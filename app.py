"""
Dashboard Resumen - Modernización RF (R4 Bogotá)
Dos formas de cargar los datos:
  1) Lista de sitios + export de RF + export de baseline (la app calcula todo), o
  2) El Excel completo de control (lee la pestaña "Resumen" y recalcula desde "Sitios").

Ejecutar:
    pip install -r requirements.txt
    streamlit run app.py
"""

import datetime
import io
import unicodedata

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ---------------------------------------------------------------------------
# Configuración general
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Dashboard Resumen · Modernización R4",
    page_icon="📡",
    layout="wide",
)

HOJA = "Resumen"

# Hitos que forman el embudo, en orden de avance. Los demás (negados, cosite)
# se muestran como indicadores aparte.
ORDEN_EMBUDO = [
    "Total Sitios",
    "TSS Cargados",
    "TSS Aprobados",
    "Obra Civil Ok",
    "Diseño TX",
    "RFIC-OK",
    "Instalados",
    "Integrados",
    "InSrv",
]
HITOS_ALERTA = ["CARGA NEGADOS", "COSITE"]

# Paleta categórica (orden fijo, validada para daltonismo) y rampa secuencial azul
COLORES_TERRITORIO = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100",
                      "#e87ba4", "#008300", "#6250d6", "#e34948"]
AZUL = "#2a78d6"
ESCALA_SECUENCIAL = [[0.0, "#cde2fb"], [0.5, "#5598e7"], [1.0, "#0d366b"]]
GRIS = "#8a8984"

# ---------------------------------------------------------------------------
# Estilos
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
      .block-container {padding-top: 1.6rem; padding-bottom: 2rem;}
      .kpi {
        border: 1px solid rgba(128,128,128,.25);
        border-radius: 12px;
        padding: 14px 18px;
        background: rgba(128,128,128,.04);
        height: 100%;
      }
      .kpi .lbl {font-size: .82rem; opacity: .7; margin-bottom: 2px;}
      .kpi .val {font-size: 1.9rem; font-weight: 700; line-height: 1.15;}
      .kpi .sub {font-size: .8rem; opacity: .65; margin-top: 2px;}
      .k2 {border: 1px solid rgba(128,128,128,.25); border-radius: 12px; padding: 12px 16px;
           background: rgba(128,128,128,.04); height: 118px; box-sizing: border-box;
           border-left: 5px solid var(--c);}
      .k2 .lbl {font-size: .8rem; opacity: .7; white-space: nowrap; overflow: hidden;
                text-overflow: ellipsis;}
      .k2 .val {font-size: 1.85rem; font-weight: 700; line-height: 1.2;}
      .k2 .sub {font-size: .76rem; opacity: .65; white-space: nowrap;}
      .k2 .bar {height: 5px; border-radius: 3px; background: rgba(128,128,128,.2); margin-top: 7px;}
      .k2 .bar > div {height: 5px; border-radius: 3px; background: var(--c);}
      .k2 .chips {display: flex; gap: 6px; margin-top: 4px;}
      .k2 .chip {font-size: .78rem; padding: 2px 8px; border-radius: 10px; color: #fff; font-weight: 600;}
      .kpi.alerta {border-left: 4px solid #e34948;}
      .kpi.ok {border-left: 4px solid #1baf7a;}
      .kpi.info {border-left: 4px solid #2a78d6;}
      .panel-t {font-size: 1.05rem; font-weight: 600; margin: 0 0 .15rem 0;}
      .panel-s {font-size: .82rem; opacity: .65; margin: 0 0 .6rem 0;}
      .hero {font-size: 3rem; font-weight: 700; line-height: 1; color: #1baf7a;}
      .hero-sub {font-size: .9rem; opacity: .75; margin-top: 6px;}
      .pbar {position: relative; height: 12px; border-radius: 6px;
             background: rgba(128,128,128,.18); margin: 14px 0 6px 0; overflow: hidden;}
      .pbar > div {position: absolute; top: 0; left: 0; height: 100%; border-radius: 6px;}
      .pleg {display: flex; gap: 16px; font-size: .78rem; opacity: .8; flex-wrap: wrap;}
      .dot {display: inline-block; width: 9px; height: 9px; border-radius: 50%;
            margin-right: 5px; vertical-align: middle;}
      .bn-row {margin: 0 0 12px 0;}
      .bn-top {display: flex; justify-content: space-between; font-size: .86rem;}
      .bn-top b {font-variant-numeric: tabular-nums;}
      .bn-bar {height: 6px; border-radius: 3px; background: #eb6834; margin-top: 4px;}
      .chips {display: flex; flex-wrap: wrap; gap: 8px 18px; margin-top: 4px;}
      .chip {font-size: .84rem;}
      .chip b {font-variant-numeric: tabular-nums;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Lectura y normalización de la hoja Resumen
# ---------------------------------------------------------------------------
def _norm(txt) -> str:
    return str(txt).strip().upper()


def parsear_resumen(raw: pd.DataFrame):
    """Convierte la hoja Resumen (sin encabezado) en un DataFrame limpio.

    Estructura esperada:
        fila 0  -> [región, territorio 1..n, columna total, (extras)]
        filas   -> [hito, valor por territorio..., total, (extras)]
    Devuelve (df, territorios, total_col, info) donde df tiene índice = hito.
    """
    raw = raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    encabezado = raw.iloc[0].tolist()
    region = str(encabezado[0]).strip() if pd.notna(encabezado[0]) else "Región"

    # Columnas con encabezado de texto (territorios + total). Se ignoran
    # columnas auxiliares cuyo encabezado es numérico o vacío.
    cols_txt = [i for i, v in enumerate(encabezado[1:], start=1)
                if isinstance(v, str) and v.strip()]
    if len(cols_txt) < 2:
        raise ValueError("No se encontraron columnas de territorios en la hoja Resumen.")

    # La columna total es la que no parece un territorio (p. ej. MODERNIZACION_2026);
    # por defecto, la última columna de texto.
    idx_total = next((i for i in cols_txt if "MODERN" in _norm(encabezado[i])
                      or "TOTAL" in _norm(encabezado[i])), cols_txt[-1])
    idx_terr = [i for i in cols_txt if i != idx_total]

    territorios = [str(encabezado[i]).strip().title() for i in idx_terr]
    total_col = str(encabezado[idx_total]).strip()

    cuerpo = raw.iloc[1:]
    cuerpo = cuerpo[cuerpo[0].notna()]
    df = pd.DataFrame(
        {t: pd.to_numeric(cuerpo[i], errors="coerce").fillna(0).astype(int).values
         for t, i in zip(territorios, idx_terr)},
        index=[str(h).strip() for h in cuerpo[0]],
    )
    df["Total archivo"] = pd.to_numeric(cuerpo[idx_total], errors="coerce").fillna(0).astype(int).values
    df["Total"] = df[territorios].sum(axis=1)
    df.index.name = "Hito"

    # Valor numérico suelto en el encabezado (p. ej. 208) como referencia
    extra_num = next((v for v in encabezado[idx_total + 1:]
                      if pd.api.types.is_number(v) and pd.notna(v) and v), None)

    info = {"region": region, "total_col": total_col, "ref_encabezado": extra_num}
    return df, territorios, total_col, info


# ---------------------------------------------------------------------------
# Recálculo desde la pestaña Sitios (mismas reglas que las fórmulas del
# Resumen, sin los errores de celda: #REF!, territorio equivocado, etc.)
# ---------------------------------------------------------------------------
HOJA_SITIOS = "Sitios"
SIN_TERRITORIO = "Sin territorio"

# Columna de Sitios que usa cada hito en el Resumen original (A=0, B=1, ...)
COL_SITIOS = {"region": 1, "proyecto": 4, "carga": 9}
HITO_COL = {
    "TOTAL SITIOS": None,
    "TSS CARGADOS": 5,    # F  FECHA TSS
    "TSS APROBADOS": 10,  # K  TSS OK
    "OBRA CIVIL OK": 12,  # M  OC
    "DISEÑO TX": 14,      # O  DISEÑO TX
    "RFIC-OK": 13,        # N  RFI-C
    "INSTALADOS": 15,     # P  Fecha Instalación
    "INTEGRADOS": 16,     # Q  Fecha Integración
    "INSRV": 17,          # R  Fecha Inservice
}
HITO_TEXTO = {"CARGA NEGADOS": "NEGADO", "COSITE": "COSITE"}  # columna J
# Estados de la columna "carga" (para la pestaña 📦 Carga)
ORDEN_CARGA = ["Viable", "Pendiente", "Cosite", "Negado", "Con fecha", "Sin dato"]
COLOR_CARGA = {"Viable": "#1baf7a", "Pendiente": "#eda100", "Cosite": "#5598e7",
               "Negado": "#d6453d", "Con fecha": "#8e6bd6", "Sin dato": "#c3c2b7"}
PRIORIDAD_CARGA = ["Negado", "Cosite", "Pendiente", "Viable"]  # si un sitio tiene varios
PREFIJO_CARGA = "Carga · "
PREFIJO_PRES = "Presentación · "
ROJO_CLARO = "#ff4b3e"


def estado_carga(x) -> str:
    """Normaliza el texto de carga: 'NEGADO' -> 'Negado', vacío -> 'Sin dato'."""
    if x is None or x is pd.NaT or (isinstance(x, float) and pd.isna(x)):
        return "Sin dato"
    if hasattr(x, "year"):  # hay sitios con una fecha escrita en la columna carga
        return "Con fecha"
    if hasattr(x, "hour") or (pd.api.types.is_number(x) and x == 0):
        return "Sin dato"  # 0 / 00:00:00 de una fórmula: equivale a vacío
    t = str(x).strip()
    if not t or t.lower() in ("nan", "none", "nat"):
        return "Sin dato"
    return t.capitalize()


def estado_sitio(valores) -> str:
    """Estado de carga de un sitio con varios trabajos: gana el más crítico."""
    v = set(valores)
    for e in PRIORIDAD_CARGA:
        if e in v:
            return e
    otros = sorted(v - {"Sin dato"})
    return otros[0] if otros else "Sin dato"
# Criterio para contar RFIC-OK
RFIC_COMBINADO = "TSS ok + CW Lista + Orden Tx con fecha"
RFIC_COLUMNA = "Fecha en la columna RFIC"


def es_positivo(v) -> bool:
    """Equivale al criterio ">0" de CONTAR.SI.CONJUNTO en Excel."""
    if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT:
        return False
    if isinstance(v, bool):
        return False
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        return True
    if isinstance(v, datetime.time):
        return v != datetime.time(0, 0)
    if isinstance(v, (int, float)):
        return v > 0
    return False  # texto, errores (#N/A), etc.


def recalcular_desde_sitios(raw_sitios: pd.DataFrame, df_excel: pd.DataFrame,
                            territorios: list, proyecto: str,
                            criterio: str = RFIC_COLUMNA) -> pd.DataFrame:
    s = raw_sitios.iloc[1:]  # fila 0 = encabezados
    s = s[s[COL_SITIOS["proyecto"]].astype(str).str.strip().str.upper() == _norm(proyecto)]

    mapa = {_norm(t): t for t in territorios}
    terr = s[COL_SITIOS["region"]].map(
        lambda x: mapa.get(_norm(x), SIN_TERRITORIO) if pd.notna(x) else SIN_TERRITORIO)

    columnas = list(territorios)
    if (terr == SIN_TERRITORIO).any():
        columnas.append(SIN_TERRITORIO)

    filas = {}
    for hito in df_excel.index:
        clave = _norm(hito)
        if clave in HITO_COL:
            col = HITO_COL[clave]
            mask = (pd.Series(True, index=s.index) if col is None
                    else s[col].map(es_positivo))
            if clave == "RFIC-OK" and criterio == RFIC_COMBINADO:
                mask = (s[HITO_COL["TSS APROBADOS"]].map(es_positivo)
                        & s[HITO_COL["OBRA CIVIL OK"]].map(es_positivo)
                        & s[HITO_COL["DISEÑO TX"]].map(es_positivo))
        elif clave in HITO_TEXTO:
            mask = s[COL_SITIOS["carga"]].astype(str).str.strip().str.upper() == HITO_TEXTO[clave]
        else:
            filas[hito] = {c: int(df_excel.loc[hito, c]) if c in df_excel.columns else 0
                           for c in columnas}  # hito desconocido: se deja el del Excel
            continue
        filas[hito] = {c: int((mask & (terr == c)).sum()) for c in columnas}

    df = pd.DataFrame.from_dict(filas, orient="index")[columnas]
    df["Total"] = df[columnas].sum(axis=1)
    df.index.name = "Hito"
    sin_terr = s.loc[terr == SIN_TERRITORIO, [0, COL_SITIOS["region"]]].copy()
    sin_terr.columns = ["Sitio", "Región en Sitios"]
    return df, sin_terr.reset_index(drop=True)


# Columnas de fecha que se muestran en el detalle (nombre visible -> columna de Sitios)
COLS_DETALLE = {
    "Fecha TSS": 5, "TSS OK": 10, "Obra civil OK": 12, "Diseño TX": 14, "RFIC": 13,
    "Instalación": 15, "Integración": 16, "InSrv": 17,
}
# Qué columna de fecha corresponde a cada hito del Resumen
HITO_FECHA = {"TSS CARGADOS": "Fecha TSS", "TSS APROBADOS": "TSS OK",
              "OBRA CIVIL OK": "Obra civil OK", "DISEÑO TX": "Diseño TX", "RFIC-OK": "RFIC",
              "INSTALADOS": "Instalación", "INTEGRADOS": "Integración", "INSRV": "InSrv"}
ETAPAS = [  # (nombre, color) de la etapa más temprana a la más avanzada
    ("Sin TSS cargado", "#c3c2b7"), ("TSS por aprobar", "#86b6ef"),
    ("Aprobado, sin instalar", "#5598e7"), ("Instalado, sin integrar", "#256abf"),
    ("Integrado, sin InSrv", "#104281"), ("En servicio", "#1baf7a"),
]
PREFIJO_ETAPA = "Etapa · "


def a_fecha(v):
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        return pd.Timestamp(v)
    return pd.NaT


def asignar_etapa(det: pd.DataFrame) -> pd.Series:
    """Último hito que cumple cada sitio (cada sitio queda en una sola etapa)."""
    def etapa(r):
        if r["ok_INSRV"]:
            return ETAPAS[5][0]
        if r["ok_INTEGRADOS"]:
            return ETAPAS[4][0]
        if r["ok_INSTALADOS"]:
            return ETAPAS[3][0]
        if r["ok_TSS APROBADOS"]:
            return ETAPAS[2][0]
        if r["ok_TSS CARGADOS"]:
            return ETAPAS[1][0]
        return ETAPAS[0][0]
    return det.apply(etapa, axis=1)


def construir_detalle(raw_sitios: pd.DataFrame, territorios: list, proyecto: str,
                      criterio: str = RFIC_COLUMNA) -> pd.DataFrame:
    """Una fila por sitio del proyecto, con su territorio, fechas y si cumple cada hito."""
    s = raw_sitios.iloc[1:]
    s = s[s[COL_SITIOS["proyecto"]].astype(str).str.strip().str.upper() == _norm(proyecto)]
    mapa = {_norm(t): t for t in territorios}
    det = pd.DataFrame({
        "Sitio": s[0].astype(str).str.strip(),
        "Territorio": s[COL_SITIOS["region"]].map(
            lambda x: mapa.get(_norm(x), SIN_TERRITORIO) if pd.notna(x) else SIN_TERRITORIO),
    })
    for nombre, col in COLS_DETALLE.items():
        det[nombre] = s[col].map(a_fecha)
    for clave, col in HITO_COL.items():
        det["ok_" + clave] = True if col is None else s[col].map(es_positivo)
    if criterio == RFIC_COMBINADO:
        det["ok_RFIC-OK"] = (det["ok_TSS APROBADOS"] & det["ok_OBRA CIVIL OK"]
                             & det["ok_DISEÑO TX"])
        det["RFIC"] = det[["TSS OK", "Obra civil OK", "Diseño TX"]].max(axis=1).where(
            det["ok_RFIC-OK"])
    carga = s[COL_SITIOS["carga"]].astype(str).str.strip().str.upper()
    for clave, txt in HITO_TEXTO.items():
        det["ok_" + clave] = carga == txt
    det["Carga"] = s[COL_SITIOS["carga"]].map(lambda x: x if isinstance(x, str) else "")
    det["Estado carga"] = s[COL_SITIOS["carga"]].map(estado_carga)
    det["Autorización"] = s[3].map(lambda x: x if isinstance(x, str) else "")
    det["Scope"] = s[21].map(lambda x: x if isinstance(x, str) else "") if 21 in s.columns else ""

    det["Etapa"] = asignar_etapa(det)
    return det.reset_index(drop=True)


@st.cache_data(show_spinner="Leyendo archivo…")
def cargar(contenido: bytes, criterio: str = RFIC_COLUMNA):
    tipo = tipo_contenido(contenido)
    if tipo == "cifrado":
        raise ValueError(MENSAJE_CIFRADO)
    if tipo in ("csv", "html"):
        raise ValueError("Este archivo no es un libro de Excel. Para exports de texto o HTML "
                         "usa el modo **Lista de sitios + RF + baseline**.")
    try:
        xls = pd.ExcelFile(io.BytesIO(contenido))  # openpyxl (.xlsx) o xlrd (.xls antiguo)
    except ImportError as e:
        raise ValueError("Este es un Excel antiguo (.xls) y falta la librería 'xlrd' "
                         "(agrégala a requirements.txt) o guárdalo como .xlsx.") from e
    buscar_hoja = lambda n: next((h for h in xls.sheet_names  # noqa: E731
                                  if h.strip().lower() == n.lower()), None)
    hoja = buscar_hoja(HOJA)
    if hoja is None:
        raise ValueError(f"El archivo no tiene una pestaña llamada '{HOJA}'. "
                         f"Pestañas encontradas: {', '.join(xls.sheet_names)}")
    raw = pd.read_excel(xls, sheet_name=hoja, header=None)
    df_excel, territorios, total_col, info = parsear_resumen(raw)

    df_recalc, sin_terr, det, error_recalc = None, None, None, None
    hoja_s = buscar_hoja(HOJA_SITIOS)
    if hoja_s is None:
        error_recalc = f"No existe la pestaña '{HOJA_SITIOS}'; solo se muestra el Resumen del Excel."
    else:
        try:
            raw_s = pd.read_excel(xls, sheet_name=hoja_s, header=None)
            df_recalc, sin_terr = recalcular_desde_sitios(
                raw_s, df_excel, territorios, total_col, criterio)
            det = construir_detalle(raw_s, territorios, total_col, criterio)
        except Exception as e:  # noqa: BLE001
            error_recalc = f"No se pudo recalcular desde '{HOJA_SITIOS}': {e}"
    return df_excel, df_recalc, sin_terr, det, territorios, total_col, info, error_recalc


def buscar(df: pd.DataFrame, nombre: str):
    """Busca un hito sin importar mayúsculas/espacios."""
    for h in df.index:
        if _norm(h) == _norm(nombre):
            return h
    return None


def pct(a, b) -> float:
    return (a / b * 100) if b else 0.0


def kpi(col, etiqueta, valor, sub="", tipo="info"):
    col.markdown(
        f"""<div class="kpi {tipo}">
              <div class="lbl">{etiqueta}</div>
              <div class="val">{valor}</div>
              <div class="sub">{sub}</div>
            </div>""",
        unsafe_allow_html=True,
    )


def kpi2(col, etiqueta, valor, sub="", color="#2a78d6", barra=None, extra=""):
    """Tarjeta de indicador: borde de color, número grande y barra de avance opcional."""
    bar = (f'<div class="bar"><div style="width:{max(0, min(barra, 100)):.0f}%"></div></div>'
           if barra is not None else "")
    col.markdown(
        f'<div class="k2" style="--c:{color}"><div class="lbl">{etiqueta}</div>'
        f'<div class="val">{valor}</div><div class="sub">{sub}</div>{bar}{extra}</div>',
        unsafe_allow_html=True)


def layout_base(fig, alto=380):
    fig.update_layout(
        height=alto,
        margin=dict(l=10, r=20, t=30, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(size=13),
        hoverlabel=dict(font_size=13),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="rgba(128,128,128,.18)", zeroline=False)
    return fig



# ---------------------------------------------------------------------------
# Modo "lista de sitios + RF + baseline"
# ---------------------------------------------------------------------------
RF_REQ = ("Sitio", "Id", "MasterPlan")
BL_REQ = ("Sitio", "Territorio Comercial")
# Hito -> nombres posibles (normalizados) de la columna de RF con su fecha
HITO_RF = {
    "TSS CARGADOS": ("tss",),
    "TSS APROBADOS": ("tss ok",),
    "OBRA CIVIL OK": ("cw lista",),
    "DISEÑO TX": ("orden tx",),
    "RFIC-OK": ("rfic",),
    "INSTALADOS": ("fecha instalacion",),
    "INTEGRADOS": ("fecha integracion",),
    "INSRV": ("fecha insrv", "fecha inservice"),
}
# Hito -> columna de fecha que se muestra en el detalle
HITO_A_DETALLE = {"TSS CARGADOS": "Fecha TSS", "TSS APROBADOS": "TSS OK",
                  "OBRA CIVIL OK": "Obra civil OK", "DISEÑO TX": "Diseño TX",
                  "RFIC-OK": "RFIC", "INSTALADOS": "Instalación",
                  "INTEGRADOS": "Integración", "INSRV": "InSrv"}
ORDEN_HITOS = ["Total Sitios", "TSS Cargados", "TSS Aprobados", "Obra Civil Ok", "Diseño TX",
               "RFIC-OK", "Instalados", "Integrados", "InSrv", "CARGA NEGADOS", "COSITE"]
ORDEN_TERR = ["Bogota Oriente", "Bogota Norte", "Bogota Sur", "Bogota Occidente"]


def _plano(t) -> str:
    """Minúsculas, sin tildes y sin espacios repetidos (para comparar nombres)."""
    s = unicodedata.normalize("NFKD", str(t))
    return " ".join("".join(ch for ch in s if not unicodedata.combining(ch)).lower().split())


def mapa_cols(df: pd.DataFrame) -> dict:
    m = {}
    for c in df.columns:
        m.setdefault(_plano(c), c)
    return m


def decodificar(contenido: bytes) -> str:
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return contenido.decode(enc)
        except UnicodeDecodeError:
            continue
    return contenido.decode("utf-8", errors="ignore")


def limpiar_sitios(valores) -> list:
    """Quita vacíos, encabezado y repetidos (sin importar tildes/mayúsculas)."""
    vistos, salida = set(), []
    for v in valores:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            continue
        n = str(v).strip()
        k = _plano(n)
        if not n or k in ("nan", "sitio", "sitios", "nombre", "none") or k in vistos:
            continue
        vistos.add(k)
        salida.append(n)
    return salida


@st.cache_data(show_spinner=False)
def leer_lista(contenido: bytes, nombre: str) -> tuple:
    n = nombre.lower()
    if n.endswith((".xlsx", ".xlsm")):
        x = pd.ExcelFile(io.BytesIO(contenido), engine="openpyxl")
        t = pd.read_excel(x, x.sheet_names[0], header=None)
        col = 0
        fila0 = [_plano(v) for v in t.iloc[0].tolist()] if len(t) else []
        if "sitio" in fila0:
            col = fila0.index("sitio")
        return tuple(limpiar_sitios(t.iloc[:, col].tolist()))
    lineas = decodificar(contenido).splitlines()
    if n.endswith(".csv"):
        import re
        filas = [re.split(r"[;\t,]", ln) for ln in lineas]
        col = 0
        if filas and "sitio" in [_plano(c) for c in filas[0]]:
            col = [_plano(c) for c in filas[0]].index("sitio")
        return tuple(limpiar_sitios([f[col] if len(f) > col else "" for f in filas]))
    return tuple(limpiar_sitios(lineas))


def _convertir_fechas_texto(df: pd.DataFrame) -> pd.DataFrame:
    """Los exports de texto (csv/html) traen las fechas como texto: se convierten a fecha."""
    conocidas = {n for nombres in HITO_RF.values() for n in nombres}
    for c in df.columns:
        if _plano(c) in conocidas:
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True, format="mixed")
    return df


MENSAJE_CIFRADO = (
    "Este archivo está **protegido** (cifrado con contraseña o con una etiqueta de "
    "confidencialidad / gestión de derechos) y solo Excel con tu cuenta puede abrirlo; "
    "la app no puede leerlo así. Para usarlo: en Excel ve a *Archivo → Información → "
    "Proteger libro* (o a *Confidencialidad* en la cinta) y quita la protección, o copia las "
    "hojas que necesitas a un libro nuevo y guárdalo. También puedes usar el modo "
    "**Lista de sitios + RF + baseline**, que no necesita este archivo.")


def tipo_contenido(contenido: bytes) -> str:
    """Detecta el formato por el contenido (no por la extensión): muchos '.xls' de los
    sistemas internos son en realidad tablas HTML."""
    if contenido[:4] == b"PK\x03\x04":
        return "xlsx"
    if contenido[:4] == b"\xd0\xcf\x11\xe0":
        # Contenedor OLE: puede ser un .xls real o un libro cifrado (contraseña / IRM)
        if (b"E\x00n\x00c\x00r\x00y\x00p\x00t\x00e\x00d\x00P\x00a\x00c\x00k\x00a\x00g\x00e\x00" in contenido
                or b"D\x00R\x00M\x00E\x00n\x00c\x00r\x00y\x00p\x00t\x00e\x00d\x00" in contenido):
            return "cifrado"
        return "xls"
    muestra = decodificar(contenido[:200000]).lower()
    return "html" if ("<table" in muestra or "<html" in muestra) else "csv"


@st.cache_data(show_spinner="Leyendo archivo…")
def leer_tabla(contenido: bytes, nombre: str, hojas_pref: tuple, requeridas: tuple):
    """Lee un export (xlsx, xls, csv o tabla HTML). Busca la fila de encabezados y la hoja."""
    req = [_plano(r) for r in requeridas]

    def fila_encabezado(crudo):
        for i in range(min(15, len(crudo))):
            celdas = {_plano(c) for c in crudo.iloc[i].tolist() if pd.notna(c)}
            if all(r in celdas for r in req):
                return i
        return None

    tipo = tipo_contenido(contenido)
    if tipo == "cifrado":
        raise ValueError(MENSAJE_CIFRADO)
    if tipo in ("csv", "html"):
        if tipo == "csv":
            crudos = [pd.read_csv(io.StringIO(decodificar(contenido)), sep=None,
                                  engine="python", header=None, dtype=object)]
        else:
            try:
                tablas = pd.read_html(io.StringIO(decodificar(contenido)), header=0)
            except ImportError as e:
                raise ValueError("Falta la librería 'lxml' para leer este tipo de archivo "
                                 "(agrégala a requirements.txt).") from e
            except ValueError as e:
                raise ValueError(f"'{nombre}' no contiene ninguna tabla legible.") from e
            crudos = []
            for t in tablas:
                cab = pd.DataFrame([[str(c) for c in t.columns]])
                t = t.copy()
                t.columns = range(t.shape[1])
                crudos.append(pd.concat([cab, t.astype(object)], ignore_index=True))
        for crudo in crudos:
            i = fila_encabezado(crudo)
            if i is not None:
                df = crudo.iloc[i + 1:].copy()
                df.columns = [str(c) for c in crudo.iloc[i].tolist()]
                df = df.loc[:, ~pd.Index(df.columns).duplicated()].reset_index(drop=True)
                return _convertir_fechas_texto(df)
    else:
        try:
            x = pd.ExcelFile(io.BytesIO(contenido))
        except ImportError as e:
            raise ValueError("Este es un Excel antiguo (.xls binario) y falta la librería "
                             "'xlrd'. Guárdalo como .xlsx o agrega xlrd a requirements.txt.") from e
        pref = [_plano(h) for h in hojas_pref]
        hojas = ([h for h in x.sheet_names if _plano(h) in pref]
                 + [h for h in x.sheet_names if _plano(h) not in pref])
        for h in hojas:
            i = fila_encabezado(pd.read_excel(x, sheet_name=h, header=None, nrows=15))
            if i is not None:
                return pd.read_excel(x, sheet_name=h, header=i)
    raise ValueError(f"No encontré las columnas {', '.join(requeridas)} en '{nombre}'. "
                     "Revisa que sea el export correcto.")


@st.cache_data(show_spinner=False)
def planes_disponibles(rf_bytes: bytes, rf_nombre: str, sitios: tuple) -> list:
    rf = leer_tabla(rf_bytes, rf_nombre, ("RF",), RF_REQ)
    m = mapa_cols(rf)
    claves = {_plano(x) for x in sitios}
    mask = rf[m["sitio"]].map(_plano).isin(claves)
    vc = rf.loc[mask, m["masterplan"]].dropna().astype(str).str.strip().value_counts()
    return list(vc.index)


@st.cache_data(show_spinner="Calculando…")
def procesar_archivos(rf_bytes, rf_nombre, bl_bytes, bl_nombre, sitios: tuple, planes: tuple,
                      criterio: str = RFIC_COLUMNA):
    """Arma el resumen por territorio y el detalle por sitio desde RF + baseline."""
    rf = leer_tabla(rf_bytes, rf_nombre, ("RF",), RF_REQ)
    bl = leer_tabla(bl_bytes, bl_nombre, ("baseline",), BL_REQ)
    mrf, mbl = mapa_cols(rf), mapa_cols(bl)
    claves = [_plano(x) for x in sitios]

    # Territorio de cada sitio (baseline)
    t = pd.DataFrame({"k": bl[mbl["sitio"]].map(_plano),
                      "t": bl[mbl["territorio comercial"]].map(
                          lambda x: str(x).strip() if pd.notna(x) else "")})
    t = t[t["t"] != ""].drop_duplicates("k").set_index("k")["t"]
    en_baseline = {_plano(x) for x in bl[mbl["sitio"]].dropna()}
    terr_sitio = [t.get(k, SIN_TERRITORIO) for k in claves]

    # Trabajos de RF de esos sitios y de los MasterPlan elegidos
    r = rf.copy()
    r["_k"] = r[mrf["sitio"]].map(_plano)
    r["_p"] = r[mrf["masterplan"]].astype(str).str.strip()
    r = r[r["_k"].isin(set(claves)) & r["_p"].isin(set(planes))].copy()
    n_jobs = r.groupby("_k").size()

    det = pd.DataFrame({"Sitio": list(sitios), "Territorio": terr_sitio})
    faltantes = []
    for clave_h, candidatos in HITO_RF.items():
        col = next((mrf[c] for c in candidatos if c in mrf), None)
        if col is None:
            faltantes.append(f"{HITO_A_DETALLE[clave_h]} (columna '{candidatos[0]}')")
            det["ok_" + clave_h] = False
            det[HITO_A_DETALLE[clave_h]] = pd.NaT
            continue
        f = pd.DataFrame({"_k": r["_k"], "f": r[col].map(a_fecha)})
        ag = f.groupby("_k")["f"].agg(["count", "max"])
        ok = (ag["count"] == n_jobs.reindex(ag.index)) & (ag["count"] > 0)
        det["ok_" + clave_h] = [bool(ok.get(k, False)) for k in claves]
        det[HITO_A_DETALLE[clave_h]] = [ag["max"].get(k, pd.NaT) if ok.get(k, False) else pd.NaT
                                        for k in claves]
    det["ok_TOTAL SITIOS"] = True

    # RFIC-OK combinado: cada trabajo debe tener fecha en TSS ok, CW Lista y Orden Tx
    if criterio == RFIC_COMBINADO:
        cols3 = [next((mrf[c] for c in HITO_RF[k] if c in mrf), None)
                 for k in ("TSS APROBADOS", "OBRA CIVIL OK", "DISEÑO TX")]
        if any(c is None for c in cols3):
            faltantes.append("RFIC combinado (falta alguna de: TSS ok, CW Lista, Orden Tx)")
            det["ok_RFIC-OK"] = False
            det["RFIC"] = pd.NaT
        else:
            f3 = pd.DataFrame({i: r[c].map(a_fecha) for i, c in enumerate(cols3)})
            job_ok = f3.notna().all(axis=1)
            tmp = pd.DataFrame({"_k": r["_k"], "ok": job_ok, "f": f3.max(axis=1).where(job_ok)})
            ag = tmp.groupby("_k").agg(n_ok=("ok", "sum"), mx=("f", "max"))
            ok = (ag["n_ok"] == n_jobs.reindex(ag.index)) & (ag["n_ok"] > 0)
            det["ok_RFIC-OK"] = [bool(ok.get(k, False)) for k in claves]
            det["RFIC"] = [ag["mx"].get(k, pd.NaT) if ok.get(k, False) else pd.NaT
                           for k in claves]

    # Carga: negado / cosite
    ccol = mrf.get("carga")
    if ccol:
        r["_c"] = r[ccol].astype(str).str.strip().str.upper()
        neg = r.assign(x=r["_c"] == "NEGADO").groupby("_k")["x"].any()
        cos = r.assign(x=r["_c"] == "COSITE").groupby("_k")["x"].any()
        primera = r.groupby("_k")[ccol].first()
        est_carga = r[ccol].map(estado_carga).groupby(r["_k"]).agg(estado_sitio)
    else:
        est_carga = pd.Series(dtype=object)
        neg = cos = pd.Series(dtype=bool)
        primera = pd.Series(dtype=object)
        faltantes.append("Carga (columna 'carga')")
    det["ok_CARGA NEGADOS"] = [bool(neg.get(k, False)) for k in claves]
    det["ok_COSITE"] = [bool(cos.get(k, False)) and not bool(neg.get(k, False)) for k in claves]
    det["Carga"] = ["Negado" if neg.get(k, False) else "Cosite" if cos.get(k, False)
                    else (str(primera.get(k, "")) if pd.notna(primera.get(k, "")) else "")
                    for k in claves]
    det["Estado carga"] = [est_carga.get(k, "Sin dato") for k in claves]
    acol = mrf.get("alcance")
    prim_a = r.groupby("_k")[acol].first() if acol else pd.Series(dtype=object)
    det["Autorización"] = [str(prim_a.get(k, "")) if pd.notna(prim_a.get(k, "")) else ""
                           for k in claves]
    det["Scope"] = ""
    det["Trabajos"] = [int(n_jobs.get(k, 0)) for k in claves]
    det["Etapa"] = asignar_etapa(det)

    # Resumen por territorio
    terrs = [x for x in ORDEN_TERR if x in set(terr_sitio)]
    terrs += sorted(x for x in set(terr_sitio) if x not in ORDEN_TERR and x != SIN_TERRITORIO)
    if SIN_TERRITORIO in set(terr_sitio):
        terrs.append(SIN_TERRITORIO)
    filas = {h: {tt: int((det["ok_" + _norm(h)] & (det["Territorio"] == tt)).sum())
                 for tt in terrs} for h in ORDEN_HITOS}
    df = pd.DataFrame.from_dict(filas, orient="index")[terrs]
    df["Total"] = df[terrs].sum(axis=1)
    df.index.name = "Hito"

    sin_terr = pd.DataFrame({
        "Sitio": [x for x, tt in zip(sitios, terr_sitio) if tt == SIN_TERRITORIO],
        "Motivo": ["No aparece en baseline" if _plano(x) not in en_baseline
                   else "En baseline, pero sin 'Territorio Comercial'"
                   for x, tt in zip(sitios, terr_sitio) if tt == SIN_TERRITORIO]})
    calidad = {
        "n_sitios": len(sitios), "n_trabajos": int(len(r)), "planes": list(planes),
        "sin_trabajos": det.loc[det["Trabajos"] == 0, ["Sitio", "Territorio"]].reset_index(drop=True),
        "faltantes": faltantes,
    }
    return df, det, sin_terr, calidad


# ---------------------------------------------------------------------------
# Detalle de sitios al hacer clic
# ---------------------------------------------------------------------------
def leer_click(evento, key):
    """Devuelve el punto clicado solo si la selección de ese gráfico cambió."""
    try:
        puntos = list(evento["selection"]["points"])
    except (KeyError, TypeError, AttributeError):
        try:
            puntos = list(evento.selection.points)
        except AttributeError:
            puntos = []
    firma = repr(puntos)
    if st.session_state.get("_firma_" + key) == firma:
        return None
    st.session_state["_firma_" + key] = firma
    return puntos[0] if puntos else None


def ir_a_detalle(prefijo, ver, territorio="Todos", modo="Cumplen"):
    """Deja el panel de abajo apuntando a lo clicado y abre una ventana emergente con la lista."""
    st.session_state[f"{prefijo}_ver"] = ver
    st.session_state[f"{prefijo}_terr"] = territorio
    st.session_state[f"{prefijo}_modo"] = modo
    try:
        ventana_sitios(ver, territorio, modo)
    except Exception as e:  # noqa: BLE001
        st.warning(f"No pude abrir la ventana emergente ({e}). La lista está en la pestaña "
                   "📋 Tabla › 🔍 Explorar sitios.")


def grafico_clic(fig, key):
    """Muestra un gráfico de Plotly que responde a clics (Streamlit >= 1.35)."""
    return st.plotly_chart(fig, use_container_width=True, on_select="rerun",
                           selection_mode="points", key=key,
                           config={"displayModeBar": False})


def leer_celda(evento, key):
    """Celda clicada en una tabla (fila, columna) solo si la selección cambió."""
    try:
        celdas = list(evento["selection"]["cells"])
    except (KeyError, TypeError, AttributeError):
        try:
            celdas = list(evento.selection.cells)
        except AttributeError:
            celdas = []
    firma = repr(celdas)
    st.session_state["_dbg_" + key] = celdas
    if st.session_state.get("_firma_" + key) == firma:
        return None
    st.session_state["_firma_" + key] = firma
    return celdas[0] if celdas else None


def tabla_clicable(objeto, key, **kw):
    """st.dataframe con selección de casillas (Streamlit >= 1.49). Si la versión instalada
    no lo soporta, se muestra la tabla normal."""
    kw_nuevo = dict(kw)
    if kw_nuevo.pop("use_container_width", False):
        kw_nuevo["width"] = "stretch"
    try:
        ev = st.dataframe(objeto, on_select="rerun", selection_mode="single-cell",
                          key=key, **kw_nuevo)
        st.session_state.pop("_sin_clic_" + key, None)
        return ev
    except Exception as e:  # noqa: BLE001
        st.session_state["_sin_clic_" + key] = str(e)
        st.dataframe(objeto, **kw)
        return None


def clic_en_tabla(evento, key, tabla, prefijo, modo):
    """Si el usuario hizo clic en una casilla, deja el detalle apuntando a ese hito/territorio."""
    celda = leer_celda(evento, key)
    if celda is None:
        return
    try:
        fila, col = celda
    except (TypeError, ValueError):
        return
    if not isinstance(fila, int) or not (0 <= fila < len(tabla)):
        return
    if col in ("Total", "% del total"):
        terr = "Todos"
    elif col in sel_terr:
        terr = col
    else:
        return
    ir_a_detalle(prefijo, tabla.index[fila], terr, modo)


def obtener_lista(ver, terr_sel, modo):
    """Sitios que cumplen (o no) un hito en un territorio. Devuelve (tabla, título, col_fecha)."""
    es_etapa = ver.startswith(PREFIJO_ETAPA)
    es_total = _norm(ver) == "TOTAL SITIOS"
    d = det[det["Territorio"].isin(sel_terr)]
    if terr_sel != "Todos":
        d = d[d["Territorio"] == terr_sel]
    col_fecha = None
    if es_etapa:
        d = d[d["Etapa"] == ver[len(PREFIJO_ETAPA):]]
        titulo_lista = ver
    elif ver.startswith(PREFIJO_CARGA):
        d = d[d["Estado carga"] == ver[len(PREFIJO_CARGA):]]
        titulo_lista = ver
    elif ver.startswith(PREFIJO_PRES):
        clave_p = ver[len(PREFIJO_PRES):]
        inc = st.session_state.get("_pres_incluir")
        if inc and "Caja presentación" in d.columns:
            d = d[d["Caja presentación"].isin(inc)]
        if clave_p in ("RFIC OK", "RFIC NOK"):
            d = d[d["Bloque RFIC"] == clave_p]
        elif clave_p != "Meta":
            d = d[d["Caja presentación"] == clave_p]
        titulo_lista = ver
    else:
        clave = _norm(ver)
        if "ok_" + clave in d.columns:
            mask = d["ok_" + clave]
            if not es_total:
                d = d[mask] if modo == "Cumplen" else d[~mask]
        col_fecha = HITO_FECHA.get(clave)
        titulo_lista = ver if es_total else f"{ver} · {modo.lower()}"

    columnas = ["Sitio", "Territorio", "Etapa"]
    if "Trabajos" in d.columns:
        columnas.append("Trabajos")
    if col_fecha:
        columnas.append(col_fecha)
    columnas += [c for c in COLS_DETALLE if c != col_fecha]
    columnas += [c for c in ("Carga", "Autorización", "Scope")
                 if (d[c].astype(str).str.strip() != "").any()]
    d = d[columnas].sort_values(["Territorio", "Sitio"])
    return d, titulo_lista, col_fecha


def config_columnas(col_fecha):
    cfg = {c: st.column_config.DateColumn(c, format="DD/MM/YYYY") for c in COLS_DETALLE}
    if col_fecha:
        cfg[col_fecha] = st.column_config.DateColumn(f"⭐ {col_fecha}", format="DD/MM/YYYY")
    return cfg


def excel_bytes(tabla, hoja="Datos", titulo=None, indice=False, destacar=None):
    """Excel ordenado: título, encabezado azul, filtros, columnas ajustadas y fechas dd/mm/aaaa."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    t = tabla.reset_index() if indice else tabla.copy()
    for c in t.columns:  # fechas sin hora y sin NaT
        if pd.api.types.is_datetime64_any_dtype(t[c]):
            t[c] = t[c].dt.date.where(t[c].notna(), None)
    fila0 = 3 if titulo else 1
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        t.to_excel(w, sheet_name=hoja[:31], index=False, startrow=fila0 - 1)
        ws = w.sheets[hoja[:31]]
        linea = Side(style="thin", color="D9D9D9")
        borde = Border(left=linea, right=linea, top=linea, bottom=linea)
        if titulo:
            ws.cell(1, 1, titulo).font = Font(bold=True, size=14, color="1F3864")
            ws.cell(2, 1, f"Generado el {datetime.date.today():%d/%m/%Y}").font = Font(
                italic=True, size=9, color="7F7F7F")
        for j, col in enumerate(t.columns, start=1):
            h = ws.cell(fila0, j)
            h.font = Font(bold=True, color="FFFFFF")
            h.fill = PatternFill("solid", fgColor="2A78D6")
            h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            h.border = borde
            ancho = max([len(str(col))] + [len(str(x)) for x in t[col].head(500)
                                            if x is not None and not pd.isna(x)])
            ws.column_dimensions[get_column_letter(j)].width = min(max(ancho + 3, 10), 45)
            for i in range(len(t)):
                c = ws.cell(fila0 + 1 + i, j)
                c.border = borde
                if isinstance(c.value, (datetime.date, datetime.datetime)):
                    c.number_format = "DD/MM/YYYY"
                    c.alignment = Alignment(horizontal="center")
                elif isinstance(c.value, (int, float)):
                    c.alignment = Alignment(horizontal="center")
                    if isinstance(c.value, float):
                        c.number_format = "0.0"
                if i % 2 == 1:
                    c.fill = PatternFill("solid", fgColor="F3F7FC")
        if destacar:  # {(fila_0, columna_nombre)} en amarillo
            pos = {c: j for j, c in enumerate(t.columns, start=1)}
            for i, cn in destacar:
                if cn in pos:
                    ws.cell(fila0 + 1 + i, pos[cn]).fill = PatternFill("solid", fgColor="EDA100")
        ws.row_dimensions[fila0].height = 30
        ws.freeze_panes = ws.cell(fila0 + 1, 2 if indice else 1)
        if len(t):
            ws.auto_filter.ref = f"A{fila0}:{get_column_letter(len(t.columns))}{fila0 + len(t)}"
        ws.sheet_view.showGridLines = False
    return buf.getvalue()


XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def boton_excel(d, titulo_lista, terr_sel, key):
    nombre = "".join(ch if ch.isalnum() else "_" for ch in titulo_lista + "_" + terr_sel)
    sub = titulo_lista + ("" if terr_sel == "Todos" else f" · {terr_sel}")
    args = ("⬇️ Descargar lista (Excel)",
            excel_bytes(d, "Sitios", f"Sitios · {sub} ({len(d)})"))
    kw = dict(file_name=f"sitios_{nombre}.xlsx", mime=XLSX_MIME, key=key)
    try:
        st.download_button(*args, on_click="ignore", **kw)  # sin recargar la página
    except TypeError:
        st.download_button(*args, **kw)


@st.dialog("Sitios", width="large")
def ventana_sitios(ver, terr_sel, modo):
    """Ventana emergente con los sitios detrás de la casilla o barra clicada."""
    if det is None:
        return
    d, titulo_lista, col_fecha = obtener_lista(ver, terr_sel, modo)
    st.markdown(f"**{len(d):,} sitio(s)** · {titulo_lista}"
                + ("" if terr_sel == "Todos" else f" · {terr_sel}"))
    if d.empty:
        st.info("No hay sitios con esa condición.")
        return
    st.dataframe(d, hide_index=True, width="stretch", column_config=config_columnas(col_fecha),
                 height=min(38 + 35 * len(d), 420))
    boton_excel(d, titulo_lista, terr_sel, "dlg_dl")


def panel_detalle(prefijo):
    if det is None:
        return
    with st.container(border=True):
        st.markdown('<p class="panel-t">🔍 Sitios detrás de cada número</p>'
                    '<p class="panel-s">Al hacer clic en una barra o casilla se abre una ventana '
                    'con la lista. Aquí abajo también puedes elegir qué ver.</p>',
                    unsafe_allow_html=True)
        opciones_ver = list(df.index) + [PREFIJO_ETAPA + e for e, _ in ETAPAS] + [PREFIJO_CARGA + e for e in ORDEN_CARGA]
        opciones_terr = ["Todos"] + [t for t in sel_terr]
        for k, ops in ((f"{prefijo}_ver", opciones_ver), (f"{prefijo}_terr", opciones_terr)):
            if st.session_state.get(k) not in ops:
                st.session_state[k] = ops[0]
        if st.session_state.get(f"{prefijo}_modo") not in ("Cumplen", "Pendientes"):
            st.session_state[f"{prefijo}_modo"] = "Cumplen"

        c1, c2, c3 = st.columns([3, 2, 2])
        ver = c1.selectbox("Ver", opciones_ver, key=f"{prefijo}_ver")
        terr_sel = c2.selectbox("Territorio", opciones_terr, key=f"{prefijo}_terr")
        es_etapa = ver.startswith((PREFIJO_ETAPA, PREFIJO_CARGA, PREFIJO_PRES))
        es_total = _norm(ver) == "TOTAL SITIOS"
        modo = c3.radio("Mostrar", ["Cumplen", "Pendientes"], key=f"{prefijo}_modo",
                        horizontal=True, disabled=es_etapa or es_total,
                        help="Cumplen: sitios que ya tienen el hito. Pendientes: los que no.")

        d, titulo_lista, col_fecha = obtener_lista(ver, terr_sel, modo)
        st.markdown(f"**{len(d):,} sitio(s)** · {titulo_lista}"
                    + ("" if terr_sel == "Todos" else f" · {terr_sel}"))
        if fuente == FUENTE_EXCEL and not modo_arch:
            st.caption("La lista sale de la pestaña Sitios con el cálculo corregido; puede "
                       "diferir del Resumen del Excel en las celdas con error.")
        if d.empty:
            st.info("No hay sitios con esa condición.")
            return
        st.dataframe(d, hide_index=True, use_container_width=True,
                     column_config=config_columnas(col_fecha),
                     height=min(38 + 35 * len(d), 460))
        boton_excel(d, titulo_lista, terr_sel, f"{prefijo}_dl")


# ---------------------------------------------------------------------------
# Interfaz
# ---------------------------------------------------------------------------
MODO_ARCHIVOS = "Lista de sitios + RF + baseline"
MODO_EXCEL = "Excel completo de control"
FUENTE_RECALC = "Recalculado (corregido)"
FUENTE_EXCEL = "Resumen del Excel"

with st.sidebar:
    st.header("📂 Datos")
    modo = st.radio("¿Qué vas a subir?", [MODO_ARCHIVOS, MODO_EXCEL], key="modo_carga")
    modo_arch = modo == MODO_ARCHIVOS
    if modo_arch:
        f_sitios = st.file_uploader("1 · Lista de sitios (.csv, .txt, .xlsx)",
                                    type=["csv", "txt", "xlsx", "xlsm"], key="up_sitios")
        txt_sitios = st.text_area("…o pega los nombres (uno por línea)", height=90,
                                  key="txt_sitios")
        f_rf = st.file_uploader("2 · Export de RF (.xls, .xlsx, .csv)", type=["xlsx", "xlsm", "xls", "csv"],
                                key="up_rf")
        f_bl = st.file_uploader("3 · Export de baseline (.xls, .xlsx, .csv)",
                                type=["xlsx", "xlsm", "xls", "csv"], key="up_bl")
        proyecto = st.text_input("Nombre del proyecto", "MODERNIZACION_2026", key="proy")
        region_txt = st.text_input("Región", "R4", key="reg")
        st.caption("Los archivos no se guardan en ningún lado; solo se procesan en esta sesión.")
    else:
        archivo = st.file_uploader("Sube el control (.xlsx)", type=["xlsx", "xlsm"])
        st.caption("Se lee la pestaña **Resumen** y se recalcula desde **Sitios**. El archivo "
                   "no se guarda en ningún lado.")

with st.sidebar:
    st.header("⚙️ Criterio de RFIC-OK")
    criterio_rfic = st.radio(
        "RFIC-OK se cuenta cuando…", [RFIC_COMBINADO, RFIC_COLUMNA], key="criterio_rfic",
        help="• TSS ok + CW Lista + Orden Tx con fecha: el sitio cuenta cuando todos sus "
             "trabajos tienen fecha en esas tres columnas de RF.\n"
             "• Fecha en la columna RFIC: el criterio anterior.\n"
             "No aplica a la vista «Resumen del Excel», que muestra los valores tal cual.")

_t1, _t2 = st.columns([4, 1])
_t1.title("📡 Dashboard Resumen · Modernización")
presentacion = _t2.toggle("🎞️ Modo presentación", key="presentacion",
                          help="Oculta la barra lateral, las ayudas y los botones para "
                               "proyectar o capturar el tablero.")
if presentacion:
    st.markdown("<style>[data-testid='stSidebar'],[data-testid='stSidebarCollapsedControl'],"
                "[data-testid='collapsedControl'],[data-testid='stHeader'],header"
                "{display:none !important}.block-container{padding-top:1rem;"
                "max-width:1400px}</style>", unsafe_allow_html=True)


def pista(texto):
    """Ayuda breve que no se muestra en modo presentación."""
    if not presentacion:
        st.caption(texto)



df_excel = sin_terr = det = calidad = error_recalc = None
diferencias = []

if modo_arch:
    sitios = list(leer_lista(f_sitios.getvalue(), f_sitios.name)) if f_sitios else []
    sitios = limpiar_sitios(sitios + txt_sitios.splitlines())
    if not (f_rf and f_bl and sitios):
        faltan = [n for n, ok in (("la lista de sitios", bool(sitios)), ("el export de RF", bool(f_rf)),
                                  ("el export de baseline", bool(f_bl))) if not ok]
        st.info("Para empezar sube " + ", ".join(faltan) + " en la barra lateral.")
        st.stop()
    with st.sidebar:
        st.caption(f"✅ {len(sitios)} sitios en la lista")
    try:
        planes_disp = planes_disponibles(f_rf.getvalue(), f_rf.name, tuple(sitios))
    except Exception as e:  # noqa: BLE001
        st.error(f"No se pudo leer RF: {e}")
        st.stop()
    if not planes_disp:
        st.error("Ninguno de los sitios de la lista aparece en el export de RF. "
                 "Revisa que los nombres coincidan con la columna Sitio de RF.")
        st.stop()
    with st.sidebar:
        st.header("🗂️ Trabajos a contar")
        planes_sel = st.multiselect(
            "MasterPlan incluidos", planes_disp,
            default=[p for p in planes_disp if "modern" in _plano(p)] or planes_disp,
            key="planes", help="Solo los trabajos de estos MasterPlan cuentan. Un sitio cumple "
                               "un hito cuando todos sus trabajos tienen fecha.")
    if not planes_sel:
        st.warning("Elige al menos un MasterPlan.")
        st.stop()
    try:
        df_recalc, det, sin_terr, calidad = procesar_archivos(
            f_rf.getvalue(), f_rf.name, f_bl.getvalue(), f_bl.name,
            tuple(sitios), tuple(planes_sel), criterio_rfic)
    except Exception as e:  # noqa: BLE001
        st.error(f"No se pudo procesar: {e}")
        st.stop()
    total_col = proyecto.strip() or "Proyecto"
    info = {"region": region_txt.strip() or "—", "total_col": total_col, "ref_encabezado": None}
    terr_excel = [c for c in df_recalc.columns if c != "Total"]
    fuente = FUENTE_RECALC
    origen = "lista de sitios + RF + baseline"
else:
    if archivo is None:
        st.info("Sube el archivo de control en la barra lateral para ver el resumen.")
        st.stop()
    try:
        df_excel, df_recalc, sin_terr, det, terr_excel, total_col, info, error_recalc = cargar(archivo.getvalue(), criterio_rfic)
    except Exception as e:  # noqa: BLE001
        st.error(f"No se pudo leer el archivo: {e}")
        st.stop()

    with st.sidebar:
        st.header("🧮 Fuente de datos")
        if df_recalc is not None:
            fuente = st.radio("Mostrar", [FUENTE_RECALC, FUENTE_EXCEL], index=0,
                              help="Recalculado: cuenta los sitios directamente en la pestaña "
                                   "Sitios con las mismas reglas del Resumen, sin los errores "
                                   "de fórmula. Resumen del Excel: los valores tal cual.")
        else:
            fuente = FUENTE_EXCEL
            st.warning(error_recalc)
    origen = f"`{archivo.name}`"

if fuente == FUENTE_RECALC:
    df = df_recalc
    territorios = [c for c in df.columns if c != "Total"]
else:
    df = df_excel
    territorios = terr_excel

# ---- Comparación Excel vs recalculado ---------------------------------------
if df_excel is not None and df_recalc is not None:
    for h in df_excel.index:
        for t in terr_excel:
            a, b = int(df_excel.loc[h, t]), int(df_recalc.loc[h, t])
            if a != b:
                diferencias.append({"Hito": h, "Columna": t, "Excel": a, "Recalculado": b})
        if SIN_TERRITORIO in df_recalc.columns and df_recalc.loc[h, SIN_TERRITORIO]:
            diferencias.append({"Hito": h, "Columna": SIN_TERRITORIO, "Excel": 0,
                                "Recalculado": int(df_recalc.loc[h, SIN_TERRITORIO])})
        a, b = int(df_excel.loc[h, "Total archivo"]), int(df_recalc.loc[h, "Total"])
        if a != b:
            diferencias.append({"Hito": h, "Columna": "Total", "Excel": a, "Recalculado": b})

st.caption(f"Región **{info['region']}** · Proyecto **{total_col}** · "
           + (f"Fuente: **{origen}**" if modo_arch else f"Fuente: **{fuente}** · Archivo: {origen}")
           + (" · RFIC-OK = TSS ok + CW Lista + Orden Tx" if criterio_rfic == RFIC_COMBINADO
              and fuente == FUENTE_RECALC else ""))

# Filtro de territorios
with st.sidebar:
    st.header("🔎 Filtros")
    sel_terr = st.multiselect("Territorios", territorios, default=territorios,
                              key=f"terr_{fuente}")
if not sel_terr:
    st.warning("Selecciona al menos un territorio.")
    st.stop()

color_terr = {t: COLORES_TERRITORIO[i % len(COLORES_TERRITORIO)]
              for i, t in enumerate(terr_excel)}  # color fijo por territorio
color_terr[SIN_TERRITORIO] = GRIS
dfv = df[sel_terr].copy()
dfv["Total"] = dfv.sum(axis=1)

h_total = buscar(df, "Total Sitios")
h_insrv = buscar(df, "InSrv")
h_inst = buscar(df, "Instalados")
h_tss = buscar(df, "TSS Aprobados")
h_neg = buscar(df, "CARGA NEGADOS")
h_cos = buscar(df, "COSITE")

total = int(dfv.loc[h_total, "Total"]) if h_total else 0
v = lambda h: int(dfv.loc[h, "Total"]) if h else 0  # noqa: E731

# ---- Alertas de calidad de datos ----------------------------------------
alertas = []
if fuente == FUENTE_EXCEL and set(sel_terr) == set(territorios):
    dif = df[df["Total archivo"] != df["Total"]]
    for h, r in dif.iterrows():
        alertas.append(f"**{h}**: el total del archivo ({r['Total archivo']}) no coincide "
                       f"con la suma de territorios ({r['Total']}).")
if fuente == FUENTE_EXCEL and diferencias:
    alertas.append(f"Hay **{len(diferencias)}** celda(s) del Resumen del Excel con errores de "
                   f"fórmula. Cambia a *{FUENTE_RECALC}* o mira la pestaña 🔧 Correcciones.")
if fuente == FUENTE_RECALC and SIN_TERRITORIO in df.columns:
    alertas.append(("Hay sitios sin territorio en baseline" if modo_arch
                    else "Hay sitios del proyecto sin región en la pestaña Sitios")
                   + f"; se muestran en la columna **{SIN_TERRITORIO}** "
                   "(detalle en la última pestaña).")
if modo_arch and calidad is not None:
    if len(calidad["sin_trabajos"]):
        alertas.append(f"**{len(calidad['sin_trabajos'])}** sitio(s) de la lista no tienen "
                       "trabajos de los MasterPlan elegidos en RF: cuentan en el total, pero "
                       "en ningún hito.")
    if calidad["faltantes"]:
        alertas.append("En RF no encontré: " + ", ".join(calidad["faltantes"]) + ".")
if (fuente == FUENTE_EXCEL and info["ref_encabezado"] and h_total
        and int(info["ref_encabezado"]) != int(df.loc[h_total, "Total"])):
    alertas.append(f"El encabezado indica **{int(info['ref_encabezado'])}** sitios, "
                   f"pero el total de sitios es **{int(df.loc[h_total, 'Total'])}**.")
ceros = [h for h in df.index if h != h_total and h not in [h_neg, h_cos]
         and df.loc[h, "Total"] == 0]
if ceros:
    alertas.append(f"Hitos en cero en todos los territorios: **{', '.join(ceros)}** "
                   f"(revisar fórmula u origen del dato).")
if alertas and not presentacion:
    with st.expander(f"⚠️ {len(alertas)} observación(es) sobre los datos", expanded=False):
        for a in alertas:
            st.markdown(f"- {a}")

# ---- KPIs ------------------------------------------------------------------
c = st.columns(6)
pend = total - v(h_insrv)
kpi2(c[0], "Total sitios", f"{total:,}", f"{len(sel_terr)} territorio(s)", GRIS)
kpi2(c[1], "TSS aprobados", f"{v(h_tss):,}", f"{pct(v(h_tss), total):.0f}% del total",
     AZUL, pct(v(h_tss), total))
kpi2(c[2], "Instalados", f"{v(h_inst):,}", f"{pct(v(h_inst), total):.0f}% del total",
     AZUL, pct(v(h_inst), total))
kpi2(c[3], "En servicio", f"{v(h_insrv):,}", f"{pct(v(h_insrv), total):.0f}% del total",
     "#1baf7a", pct(v(h_insrv), total))
kpi2(c[4], "Pendientes InSrv", f"{pend:,}", f"{pct(pend, total):.0f}% del total",
     "#eda100", pct(pend, total))
kpi2(c[5], "Carga con alerta", f"{v(h_neg) + v(h_cos):,}", "", "#d6453d",
     extra=(f'<div class="chips"><span class="chip" style="background:#d6453d">'
            f'{v(h_neg)} Negado</span><span class="chip" style="background:#5598e7">'
            f'{v(h_cos)} Cosite</span></div>'))

st.write("")

_n5 = (int(len(calidad["sin_trabajos"]) + len(sin_terr)) if modo_arch else len(diferencias))
_nombres = ["📊 Resumen", "🎞️ Presentación", "🗺️ Territorios", "📦 Carga", "📋 Tabla"]
if not presentacion:
    _nombres.append(f"🔎 Calidad de datos ({_n5})" if modo_arch else f"🔧 Correcciones ({_n5})")
_tabs = st.tabs(_nombres)
tab1, tab_p, tab2, tab_c, tab4 = _tabs[:5]
tab5 = _tabs[5] if len(_tabs) > 5 else None

hitos_embudo = [buscar(df, h) for h in ORDEN_EMBUDO]
hitos_embudo = [h for h in hitos_embudo if h is not None]
# Agrega hitos no previstos (que no sean alertas) al final del embudo
hitos_embudo += [h for h in df.index if h not in hitos_embudo and h not in [h_neg, h_cos]]

# ---- Tab 1: Embudo ---------------------------------------------------------
VERDE = "#1baf7a"


def titulo(t, sub=""):
    st.markdown(f'<p class="panel-t">{t}</p>' + (f'<p class="panel-s">{sub}</p>' if sub else ""),
                unsafe_allow_html=True)


def html(x):
    """Inserta HTML sin sangrías (Markdown interpreta 4 espacios como bloque de código)."""
    st.markdown(" ".join(l.strip() for l in x.splitlines()), unsafe_allow_html=True)


with tab1:
    col_a, col_b = st.columns([3, 2], gap="large")

    # -- Avance por hito -----------------------------------------------------
    with col_a, st.container(border=True):
        titulo("Avance por hito", f"Sitios que ya cumplen cada hito, sobre {total:,} sitios")
        valores = [int(dfv.loc[h, "Total"]) for h in hitos_embudo]
        colores = [GRIS if val == 0 else (VERDE if h == h_insrv else AZUL)
                   for h, val in zip(hitos_embudo, valores)]
        fig = go.Figure(go.Bar(
            x=valores, y=hitos_embudo, orientation="h",
            marker=dict(color=colores, cornerradius=4),
            text=[f"<b>{val:,}</b>   {pct(val, total):.0f}%" for val in valores],
            textposition="outside", cliponaxis=False,
            textfont=dict(size=13),
            customdata=[pct(val, total) for val in valores],
            hovertemplate="<b>%{y}</b><br>%{x:,} sitios · %{customdata:.1f}% del total<extra></extra>",
        ))
        layout_base(fig, 44 * len(hitos_embudo) + 40)
        fig.update_layout(bargap=0.38, margin=dict(l=10, r=10, t=10, b=10))
        fig.update_yaxes(autorange="reversed", showgrid=False, ticklabelstandoff=10,
                         tickfont=dict(size=13))
        fig.update_xaxes(visible=False, range=[0, max(valores + [1]) * 1.22])
        fig.update_layout(clickmode="event+select")
        pto = leer_click(grafico_clic(fig, "clk_avance"), "clk_avance")
        if pto and pto.get("y") in list(df.index):
            ir_a_detalle("t4", pto["y"])
        pista("👆 Haz clic en una barra para ver sus sitios abajo.")

    with col_b:
        # -- En servicio -----------------------------------------------------
        with st.container(border=True):
            titulo("En servicio")
            p_ins, p_srv = pct(v(h_inst), total), pct(v(h_insrv), total)
            html(
                f"""<div class="hero">{p_srv:.1f}%</div>
                <div class="hero-sub"><b>{v(h_insrv):,}</b> de {total:,} sitios en servicio ·
                faltan <b>{total - v(h_insrv):,}</b></div>
                <div class="pbar">
                  <div style="width:{p_ins:.1f}%; background:#86b6ef;"></div>
                  <div style="width:{p_srv:.1f}%; background:{VERDE};"></div>
                </div>
                <div class="pleg">
                  <span><span class="dot" style="background:{VERDE}"></span>En servicio {p_srv:.0f}%</span>
                  <span><span class="dot" style="background:#86b6ef"></span>Instalados {p_ins:.0f}%</span>
                </div>"""
            )

        # -- Cuellos de botella ---------------------------------------------
        with st.container(border=True):
            titulo("Dónde se quedan los sitios", "Sitios que pasan un hito pero no el siguiente")
            caidas = []
            for x, y in zip(hitos_embudo[:-1], hitos_embudo[1:]):
                vx, vy = int(dfv.loc[x, "Total"]), int(dfv.loc[y, "Total"])
                if vx > 0 and vy > 0 and vx > vy:
                    caidas.append((x, y, vx - vy))
            caidas = sorted(caidas, key=lambda c: -c[2])[:4]
            if caidas:
                mx = caidas[0][2]
                filas_html = "".join(
                    f"""<div class="bn-row"><div class="bn-top">
                          <span>{x} → {y}</span><b>{n:,} sitios</b></div>
                          <div class="bn-bar" style="width:{n / mx * 100:.0f}%"></div></div>"""
                    for x, y, n in caidas)
                html(filas_html)
            else:
                st.caption("Sin caídas entre hitos.")

    # -- Etapa actual de cada sitio ---------------------------------------------
    cadena = [("Sin TSS cargado", h_total, buscar(df, "TSS Cargados"), "#c3c2b7"),
              ("TSS por aprobar", buscar(df, "TSS Cargados"), h_tss, "#86b6ef"),
              ("Aprobado, sin instalar", h_tss, h_inst, "#5598e7"),
              ("Instalado, sin integrar", h_inst, buscar(df, "Integrados"), "#256abf"),
              ("Integrado, sin InSrv", buscar(df, "Integrados"), h_insrv, "#104281"),
              ("En servicio", h_insrv, None, VERDE)]
    if det is not None:
        d_et = det[det["Territorio"].isin(sel_terr)]
        etapas = [(nm, int((d_et["Etapa"] == nm).sum()), col) for nm, col in ETAPAS]
    elif all(c[1] for c in cadena):
        etapas = []
        for nombre, desde, hasta, color in cadena:
            n = v(desde) - (v(hasta) if hasta else 0)
            etapas.append((nombre, max(n, 0), color))
    else:
        etapas = []
    if etapas:
        with st.container(border=True):
            titulo("¿En qué etapa está cada sitio hoy?",
                   "Cada sitio se cuenta una sola vez, en el último hito que cumple")
            fig_e = go.Figure()
            trazas_etapa = []
            for nombre, n, color in etapas:
                if n == 0:
                    continue
                trazas_etapa.append(nombre)
                fig_e.add_bar(
                    x=[n], y=[""], orientation="h", name=nombre,
                    marker=dict(color=color, line=dict(width=0)),
                    text=[f"{n}" if pct(n, total) >= 4 else ""], textposition="inside",
                    insidetextanchor="middle", textfont=dict(size=13),
                    hovertemplate=f"<b>{nombre}</b><br>{n:,} sitios · {pct(n, total):.1f}%<extra></extra>",
                )
            fig_e.update_layout(barmode="stack", showlegend=False, height=70,
                                margin=dict(l=0, r=0, t=0, b=0),
                                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            fig_e.update_xaxes(visible=False, range=[0, max(total, 1)])
            fig_e.update_yaxes(visible=False)
            fig_e.update_layout(clickmode="event+select")
            pto = leer_click(grafico_clic(fig_e, "clk_etapa"), "clk_etapa")
            if pto is not None and det is not None:
                i = pto.get("curve_number")
                if isinstance(i, int) and 0 <= i < len(trazas_etapa):
                    ir_a_detalle("t4", PREFIJO_ETAPA + trazas_etapa[i])
            chips = "".join(
                f'<span class="chip"><span class="dot" style="background:{c}"></span>'
                f'{nm} <b>{n:,}</b> <span style="opacity:.6">({pct(n, total):.0f}%)</span></span>'
                for nm, n, c in etapas)
            st.markdown(f'<div class="chips">{chips}</div>', unsafe_allow_html=True)


# ---- Tab Presentación (árbol de la meta) -----------------------------------
def clasificar_presentacion(d):
    """Devuelve (Bloque RFIC, Caja) por sitio para el árbol de la pestaña Presentación."""
    def g(c):
        return d[c].astype(bool) if c in d.columns else pd.Series(False, index=d.index)
    rfic, inst, integ = g("ok_RFIC-OK"), g("ok_INSTALADOS"), g("ok_INTEGRADOS")
    tss_c, tss_a, cw, tx = (g("ok_TSS CARGADOS"), g("ok_TSS APROBADOS"),
                            g("ok_OBRA CIVIL OK"), g("ok_DISEÑO TX"))
    est = d["Estado carga"] if "Estado carga" in d.columns else pd.Series("", index=d.index)
    caja = pd.Series("Otros", index=d.index)
    caja[rfic & ~inst & ~integ] = "Pendiente de instalar"
    caja[rfic & inst & ~integ] = "En instalación"
    caja[rfic & integ] = "Integrados"
    nok = ~rfic
    caja[nok & tss_a & cw & ~tx] = "Transmisión"
    caja[nok & tss_a & ~cw] = "Obra civil"
    caja[nok & tss_c & ~tss_a] = "En validación TSS"
    caja[nok & ~tss_c] = "Construcción TSS"
    caja[nok & (est == "Negado")] = "Carga negada"
    caja[nok & (est == "Cosite")] = "Aval Cosite"
    return pd.Series(np.where(rfic, "RFIC OK", "RFIC NOK"), index=d.index), caja


COLOR_CAJA = {
    "Meta": "#e2231a", "RFIC OK": "#1baf7a", "RFIC NOK": "#ff5a4d",
    "Integrados": "#1baf7a", "En instalación": "#4f9cf0", "Pendiente de instalar": "#9b7bea",
    "Construcción TSS": "#f2b01e", "En validación TSS": "#f58a3c", "Aval Cosite": "#2fc3e8",
    "Transmisión": "#e86ad0", "Obra civil": "#c9a46a", "Carga negada": "#d6453d",
    "Otros": "#8a8984",
}


def _rgba(hex_, a):
    h = hex_.lstrip("#")
    return f"rgba({int(h[0:2], 16)},{int(h[2:4], 16)},{int(h[4:6], 16)},{a})"


def arbol_meta(titulo_txt, etiqueta_meta, n_meta, n_ok, n_nok, hijos_ok, hijos_nok,
               regiones=(), region_sel="Todas"):
    """Árbol Meta → RFIC OK / RFIC NOK → cajas, con tarjetas clicables. Devuelve (fig, nodos)."""
    R = len(hijos_ok) + len(hijos_nok) + 1
    ys_ok = [R - i for i in range(len(hijos_ok))]
    ys_nok = [R - len(hijos_ok) - 1 - j for j in range(len(hijos_nok))]
    y_ok, y_nok = sum(ys_ok) / len(ys_ok), sum(ys_nok) / len(ys_nok)
    y_meta = (y_ok + y_nok) / 2
    con_chips = len(regiones) > 1
    y_top = R + (1.75 if con_chips else 0.85)
    alto = max(580, 78 * (R + 1) + 110 + (60 if con_chips else 0))
    px_y = alto / (y_top - 0.15)               # píxeles por unidad vertical (aprox.)
    rx, ry = 10 * 100 / 1100, 10 / px_y   # esquinas de ~10 px
    h = 0.84                              # alto de cada tarjeta (unidades y)

    # (clave, etiqueta, n, x0, x1, y, tamaño del número)
    META = (2, 21)
    PAD = (31, 50)
    HIJO = (67, 98)
    nodos = [("Meta", etiqueta_meta, n_meta, *META, y_meta, 34),
             ("RFIC OK", "RFIC OK", n_ok, *PAD, y_ok, 30),
             ("RFIC NOK", "RFIC NOK", n_nok, *PAD, y_nok, 30)]
    nodos += [(nm, nm, n, *HIJO, y, 26) for (nm, n), y in zip(hijos_ok, ys_ok)]
    nodos += [(nm, nm, n, *HIJO, y, 26) for (nm, n), y in zip(hijos_nok, ys_nok)]

    formas, anot, flechas = [], [], []

    # Resplandor de fondo
    for cx, cy, rad, col in ((0, 1.0, 38, .16), (100, 0.0, 34, .12), (60, 0.55, 30, .05)):
        formas.append(dict(type="circle", xref="x", yref="y", layer="below", line_width=0,
                           x0=cx - rad, x1=cx + rad, y0=(R + 0.7) * cy - rad * R / 70,
                           y1=(R + 0.7) * cy + rad * R / 70, fillcolor=f"rgba(226,35,26,{col})"))

    def codo(x_ini, y_ini, x_bus, x_fin, y_fin, color):
        """Línea con esquinas redondeadas: sale del padre, sube/baja por el bus y entra a la caja."""
        d = f"M {x_ini},{y_ini} "
        if abs(y_fin - y_ini) < 1e-6:
            d += f"L {x_fin},{y_fin}"
        else:
            sg = 1 if y_fin > y_ini else -1
            d += (f"L {x_bus - rx},{y_ini} Q {x_bus},{y_ini} {x_bus},{y_ini + sg * ry} "
                  f"L {x_bus},{y_fin - sg * ry} Q {x_bus},{y_fin} {x_bus + rx},{y_fin} "
                  f"L {x_fin},{y_fin}")
        formas.append(dict(type="path", path=d, layer="below",
                           line=dict(color=_rgba(color, .85), width=2.6), fillcolor="rgba(0,0,0,0)"))
        flechas.append((x_fin, y_fin, color))

    # Conectores
    for clave, yp in (("RFIC OK", y_ok), ("RFIC NOK", y_nok)):
        codo(META[1], y_meta, 26, PAD[0] - 0.4, yp, COLOR_CAJA[clave])
    for clave, yp, ys in (("RFIC OK", y_ok, ys_ok), ("RFIC NOK", y_nok, ys_nok)):
        for yh in ys:
            codo(PAD[1], yp, 58.5, HIJO[0] - 0.4, yh, COLOR_CAJA[clave])

    # Tarjetas
    xs_clic, ys_clic, claves_clic, textos = [], [], [], []
    for clave, etq, n, x0, x1, y, tam in nodos:
        col = COLOR_CAJA.get(clave, GRIS)
        ya, yb = y - h / 2, y + h / 2
        formas.append(dict(
            type="path", layer="below",
            path=(f"M {x0 + rx},{ya} L {x1 - rx},{ya} Q {x1},{ya} {x1},{ya + ry} "
                  f"L {x1},{yb - ry} Q {x1},{yb} {x1 - rx},{yb} L {x0 + rx},{yb} "
                  f"Q {x0},{yb} {x0},{yb - ry} L {x0},{ya + ry} Q {x0},{ya} {x0 + rx},{ya} Z"),
            fillcolor=_rgba(col, .16), line=dict(color=col, width=2)))
        pc = 100.0 if clave == "Meta" else pct(n, n_meta)
        # barra de avance (sobre la meta)
        anc = (x1 - x0) - 2 * rx - 0.4
        formas.append(dict(type="rect", layer="below", line_width=0, fillcolor="rgba(255,255,255,.14)",
                           x0=x0 + rx + 0.2, x1=x0 + rx + 0.2 + anc, y0=ya + 0.07, y1=ya + 0.07 + 0.05))
        formas.append(dict(type="rect", layer="below", line_width=0, fillcolor=col,
                           x0=x0 + rx + 0.2, x1=x0 + rx + 0.2 + anc * min(pc, 100) / 100,
                           y0=ya + 0.07, y1=ya + 0.07 + 0.05))
        anot.append(dict(x=x0 + 1.4, y=y + 0.07, text=f"<b>{n:,}</b>", showarrow=False,
                         font=dict(size=tam, color="#ffffff"), xanchor="left", yanchor="middle"))
        anot.append(dict(x=x0 + 1.4, y=y - 0.22, text=f"<b>{etq}</b>", showarrow=False,
                         font=dict(size=13, color=col), xanchor="left", yanchor="middle"))
        if clave != "Meta":
            anot.append(dict(x=x1 - 1.2, y=y + 0.07, text=f"<b>{pc:.0f}%</b>", showarrow=False,
                             font=dict(size=17, color="#ffffff"), xanchor="right", yanchor="middle"))
        for fx in (x0 + 0.2 * (x1 - x0), (x0 + x1) / 2, x0 + 0.8 * (x1 - x0)):
            xs_clic.append(fx); ys_clic.append(y); claves_clic.append(clave)
            textos.append(f"<b>{etq}</b><br>{n:,} sitios · {pc:.0f}% de la meta")

    # Selector de región dentro del gráfico (casillas clicables)
    if con_chips:
        regs = ["Todas"] + list(regiones)
        ancho, sep, hch = 13.5, 1.2, 0.5
        x_ini = 50 - (len(regs) * ancho + (len(regs) - 1) * sep) / 2 + 4
        yc = R + 1.2
        ya, yb = yc - hch / 2, yc + hch / 2
        rxc, ryc = (hch * px_y / 2) * 100 / 1100, hch / 2
        anot.append(dict(x=x_ini - 1.5, y=yc, text="<b>Región</b>", showarrow=False,
                         font=dict(size=13, color="#c9b8ba"), xanchor="right", yanchor="middle"))
        for k, rg in enumerate(regs):
            x0 = x_ini + k * (ancho + sep)
            x1 = x0 + ancho
            activo = rg == region_sel
            formas.append(dict(
                type="path", layer="below",
                path=(f"M {x0 + rxc},{ya} L {x1 - rxc},{ya} Q {x1},{ya} {x1},{yc} "
                      f"Q {x1},{yb} {x1 - rxc},{yb} L {x0 + rxc},{yb} Q {x0},{yb} {x0},{yc} "
                      f"Q {x0},{ya} {x0 + rxc},{ya} Z"),
                fillcolor="#e2231a" if activo else "rgba(255,255,255,.07)",
                line=dict(color="#ffffff" if activo else "rgba(255,255,255,.4)", width=2 if activo else 1.4)))
            etq_r = "Todas" if rg == "Todas" else rg.replace("Bogota ", "").replace("Bogotá ", "")
            anot.append(dict(x=(x0 + x1) / 2, y=yc, text=f"<b>{etq_r}</b>", showarrow=False,
                             font=dict(size=14, color="#ffffff" if activo else "#d9c9cb"),
                             xanchor="center", yanchor="middle"))
            for fx in (x0 + 0.25 * ancho, (x0 + x1) / 2, x0 + 0.75 * ancho):
                xs_clic.append(fx); ys_clic.append(yc); claves_clic.append("reg::" + rg)
                textos.append(f"<b>{etq_r}</b>")

    # Puntas de flecha
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=[f[0] for f in flechas], y=[f[1] for f in flechas], mode="markers",
        hoverinfo="skip", showlegend=False,
        marker=dict(symbol="triangle-right", size=10, color=[f[2] for f in flechas])))
    fig.add_trace(go.Scatter(
        x=xs_clic, y=ys_clic, mode="markers", showlegend=False, customdata=claves_clic,
        marker=dict(symbol="square", size=58, color="rgba(255,255,255,0.01)"),
        unselected=dict(marker=dict(opacity=0.01)), selected=dict(marker=dict(opacity=0.01)),
        text=textos, hovertemplate="%{text}<extra></extra>"))

    anot.append(dict(xref="paper", yref="paper", x=0.5, y=1.0, text=f"<b>{titulo_txt}</b>",
                     showarrow=False, font=dict(size=26, color="#ffffff"), bgcolor="#e2231a",
                     bordercolor="#ffffff", borderwidth=2, borderpad=10, yanchor="bottom"))
    fig.update_layout(
        shapes=formas, annotations=anot, height=alto, clickmode="event+select", dragmode=False,
        paper_bgcolor="#140a0c", plot_bgcolor="#140a0c", margin=dict(l=10, r=10, t=78, b=10),
        showlegend=False, hoverlabel=dict(bgcolor="#2a1a1d", font_size=13, font_color="#ffffff"),
        xaxis=dict(visible=False, range=[0, 100], fixedrange=True),
        yaxis=dict(visible=False, range=[0.15, y_top], fixedrange=True))
    mapa = {(round(x, 3), round(y, 3)): k for x, y, k in zip(xs_clic, ys_clic, claves_clic)}
    return fig, mapa


with tab_p:
    if det is None:
        st.info("Esta pestaña necesita el cálculo desde la pestaña Sitios o el modo lista + RF + baseline.")
    else:
        dall = det[det["Territorio"].isin(sel_terr)].copy()
        dall["Bloque RFIC"], dall["Caja presentación"] = clasificar_presentacion(dall)
        det.loc[dall.index, "Bloque RFIC"] = dall["Bloque RFIC"]
        det.loc[dall.index, "Caja presentación"] = dall["Caja presentación"]
        cajas_ok = ["Integrados", "En instalación", "Pendiente de instalar"]
        cajas_nok = ["Construcción TSS", "En validación TSS", "Aval Cosite", "Transmisión"]
        extras = ["Obra civil", "Carga negada", "Otros"]
        cuenta_all = dall["Caja presentación"].value_counts()

        # Región: se elige con las casillas dentro del gráfico
        reg_opciones = ["Todas"] + list(sel_terr)
        region_p = st.session_state.get("_pres_region", "Todas")
        if region_p not in reg_opciones:
            region_p = st.session_state["_pres_region"] = "Todas"

        if not presentacion:
            ct1, ct2 = st.columns(2)
            ct1.text_input("Título", "MODERNIZACIÓN R4 - 4Q", key="pres_titulo")
            ct2.text_input("Etiqueta de la meta", "Meta 4Q", key="pres_meta")
            st.multiselect(
                "Cajas adicionales dentro de la meta", extras,
                default=[e for e in extras if cuenta_all.get(e, 0) > 0], key="pres_extras",
                format_func=lambda e: f"{e} ({int(cuenta_all.get(e, 0))})",
                help="La meta es la suma de RFIC OK + RFIC NOK. Quita aquí las cajas que no "
                     "deban contarse (por ejemplo Carga negada u Otros).")
            # Se copian a claves propias: los widgets ocultos pierden su valor en modo presentación
            st.session_state["_pres_titulo_v"] = st.session_state.get("pres_titulo", "MODERNIZACIÓN R4 - 4Q")
            st.session_state["_pres_meta_v"] = st.session_state.get("pres_meta", "Meta 4Q")
            st.session_state["_pres_extras_v"] = list(st.session_state.get(
                "pres_extras", [e for e in extras if cuenta_all.get(e, 0) > 0]))
        titulo_p = st.session_state.get("_pres_titulo_v", "MODERNIZACIÓN R4 - 4Q")
        meta_p = st.session_state.get("_pres_meta_v", "Meta 4Q")
        extras_sel = st.session_state.get(
            "_pres_extras_v", [e for e in extras if cuenta_all.get(e, 0) > 0])
        incluidas = cajas_ok + cajas_nok + [e for e in extras if e in extras_sel]
        st.session_state["_pres_incluir"] = incluidas

        dreg = dall if region_p == "Todas" else dall[dall["Territorio"] == region_p]
        dp = dreg[dreg["Caja presentación"].isin(incluidas)]
        n_meta = len(dp)
        n_ok = int((dp["Bloque RFIC"] == "RFIC OK").sum())
        cuenta = dp["Caja presentación"].value_counts()
        extras_nok = [c for c in extras if c in extras_sel and cuenta_all.get(c, 0) > 0]
        hijos_ok = [(c, int(cuenta.get(c, 0))) for c in cajas_ok]
        hijos_nok = [(c, int(cuenta.get(c, 0))) for c in cajas_nok + extras_nok]
        titulo_arbol = titulo_p + ("" if region_p == "Todas"
                                   else f" · {region_p.replace('Bogota ', '').upper()}")

        if n_meta == 0:
            st.info("No hay sitios para esta selección.")
            if st.button("Ver todas las regiones", key="pres_reset_region"):
                st.session_state["_pres_region"] = "Todas"
                st.rerun()
        else:
            fig_t, mapa_nodos = arbol_meta(titulo_arbol, meta_p, n_meta, n_ok, n_meta - n_ok,
                                           hijos_ok, hijos_nok, list(sel_terr), region_p)
            pto = leer_click(grafico_clic(fig_t, "clk_arbol"), "clk_arbol")
            if pto:
                clave_n = mapa_nodos.get((round(float(pto.get("x", -1)), 3),
                                          round(float(pto.get("y", -1)), 3)))
                if clave_n and clave_n.startswith("reg::"):
                    st.session_state["_pres_region"] = clave_n[5:]
                    st.rerun()
                elif clave_n:
                    ir_a_detalle("t4", PREFIJO_PRES + clave_n, region_p if region_p != "Todas" else "Todos")
            fuera = len(dreg) - n_meta
            if fuera > 0:
                pista(f"La lista tiene {len(dreg):,} sitios; {fuera:,} quedan fuera de la meta "
                      "por las cajas adicionales desmarcadas.")
        pista("👆 Haz clic en una región para filtrar, o en cualquier número para ver sus sitios.")
        if not presentacion:
            with st.expander("ℹ️ ¿Cómo se clasifica cada caja?"):
                st.markdown(
                    f"- **Meta**: suma de RFIC OK + RFIC NOK de la región elegida ({n_meta:,} sitios).\n"
                    f"- **RFIC OK**: cumple el criterio elegido en la barra lateral "
                    f"(*{criterio_rfic}*). Se divide en:\n"
                    "  - **Integrados**: tiene fecha de integración.\n"
                    "  - **En instalación**: tiene fecha de instalación pero no de integración.\n"
                    "  - **Pendiente de instalar**: RFIC OK sin fecha de instalación. Aquí irán "
                    "*Hardware* y *Acceso / MOS* cuando haya datos de HW completo y MOS.\n"
                    "- **RFIC NOK** (el resto), con este orden de prioridad:\n"
                    "  - **Aval Cosite**: la carga está en Cosite.\n"
                    "  - **Construcción TSS**: sin TSS cargado.\n"
                    "  - **En validación TSS**: TSS cargado pero sin aprobar.\n"
                    "  - **Transmisión**: TSS ok y CW Lista, sin Orden Tx.\n"
                    "  - Cajas extra que solo aparecen si tienen sitios: **Obra civil** (TSS ok sin "
                    "CW Lista), **Carga negada** y **Otros**.")

# ---- Tab 2: Por territorio -------------------------------------------------
with tab2:
    vista_t = st.radio("Ver", ["📈 Avance (%)", "⏳ Pendientes"], horizontal=True,
                       key="vista_terr", label_visibility="collapsed")
    if vista_t.startswith("📈"):
        st.subheader("% de avance por territorio (sobre el total de sitios de cada uno)")
        base = dfv.loc[h_total, sel_terr].replace(0, pd.NA)
        porc = (dfv.loc[hitos_embudo, sel_terr].div(base, axis=1) * 100).astype(float).round(1)

        heat = go.Figure(go.Heatmap(
            z=porc.values, x=sel_terr, y=porc.index,
            colorscale=ESCALA_SECUENCIAL, zmin=0, zmax=100,
            text=[[f"{val:.0f}%<br>({int(dfv.loc[h, t])})" for t, val in zip(sel_terr, fila)]
                  for h, fila in zip(porc.index, porc.values)],
            texttemplate="%{text}", xgap=2, ygap=2,
            hovertemplate="<b>%{x}</b> · %{y}<br>%{z:.1f}% del total<extra></extra>",
            colorbar=dict(title="%", ticksuffix="%"),
        ))
        layout_base(heat, 460)
        heat.update_yaxes(autorange="reversed", showgrid=False)
        st.plotly_chart(heat, use_container_width=True)

        st.subheader("Comparar un hito entre territorios")
        opciones = list(df.index)
        hito_sel = st.selectbox("Hito", opciones,
                                index=opciones.index(h_insrv) if h_insrv in opciones else 0)
        vals = dfv.loc[hito_sel, sel_terr]
        bar = go.Figure(go.Bar(
            x=sel_terr, y=vals.values,
            marker=dict(color=[color_terr[t] for t in sel_terr], cornerradius=4),
            text=[f"{int(x)} · {pct(x, dfv.loc[h_total, t]):.0f}%" for t, x in zip(sel_terr, vals.values)],
            textposition="outside", cliponaxis=False,
            hovertemplate="<b>%{x}</b><br>" + hito_sel + ": %{y}<extra></extra>",
        ))
        layout_base(bar, 360)
        bar.update_yaxes(range=[0, max(vals.max(), 1) * 1.2])
        bar.update_layout(clickmode="event+select")
        pto = leer_click(grafico_clic(bar, "clk_terr"), "clk_terr")
        if pto and pto.get("x") in sel_terr:
            ir_a_detalle("t4", hito_sel, pto["x"])
        pista("👆 Haz clic en un territorio para ver sus sitios abajo.")

    else:
        st.subheader("Sitios pendientes por hito y territorio")
        st.caption("Pendiente = Total sitios del territorio − sitios que ya cumplen el hito.")
        hitos_pend = [h for h in hitos_embudo if h != h_total]
        pend = pd.DataFrame(
            {t: [int(dfv.loc[h_total, t] - dfv.loc[h, t]) for h in hitos_pend] for t in sel_terr},
            index=hitos_pend,
        )
        pend["Total"] = pend.sum(axis=1)
        pend.index.name = "Hito"

        fig_p = go.Figure()
        for t in sel_terr:
            fig_p.add_bar(name=t, y=hitos_pend, x=pend[t], orientation="h",
                          marker=dict(color=color_terr[t], line=dict(width=0)),
                          hovertemplate="<b>" + t + "</b><br>%{y}: %{x} pendientes<extra></extra>")
        fig_p.update_layout(barmode="stack", bargap=.35)
        layout_base(fig_p, 440)
        fig_p.update_yaxes(autorange="reversed", showgrid=False)
        fig_p.update_xaxes(showgrid=True, gridcolor="rgba(128,128,128,.18)")
        fig_p.update_layout(clickmode="event+select")
        pto = leer_click(grafico_clic(fig_p, "clk_pend"), "clk_pend")
        if pto and pto.get("y") in hitos_pend:
            i = pto.get("curve_number")
            terr_clic = sel_terr[i] if isinstance(i, int) and 0 <= i < len(sel_terr) else "Todos"
            ir_a_detalle("t4", pto["y"], terr_clic, "Pendientes")
        pista("👆 Haz clic en un tramo de color para ver los sitios pendientes de ese "
                   "territorio y hito.")

        ev_pend = tabla_clicable(pend, "clk_pend_tabla", use_container_width=True)
        clic_en_tabla(ev_pend, "clk_pend_tabla", pend, "t4", "Pendientes")


# ---- Tab Carga -------------------------------------------------------------
with tab_c:
    dc = det[det["Territorio"].isin(sel_terr)]
    n_c = len(dc)
    cuentas = dc["Estado carga"].value_counts()
    extra = [e for e in cuentas.index if e not in ORDEN_CARGA]  # estados no previstos
    estados = [e for e in ORDEN_CARGA + extra if cuentas.get(e, 0) > 0 or e in ORDEN_CARGA[:4]]
    st.caption("Estado de la carga de cada sitio (columna *carga*). Si un sitio tiene varios "
               "trabajos con estados distintos, se muestra el más crítico: "
               "Negado › Cosite › Pendiente › Viable.")

    cols_e = st.columns(len(estados))
    for ce, e in zip(cols_e, estados):
        n = int(cuentas.get(e, 0))
        color = COLOR_CARGA.get(e, GRIS)
        with ce, st.container(border=True):
            html(f"""<div style="border-left:5px solid {color};padding-left:10px">
                       <div style="font-size:.85rem;opacity:.75">{e}</div>
                       <div style="font-size:1.9rem;font-weight:700;line-height:1.2">{n:,}</div>
                       <div style="font-size:.8rem;opacity:.65">{pct(n, n_c):.0f}% de {n_c:,} sitios</div>
                     </div>""")
            if (not presentacion) and st.button("Ver sitios", key=f"btn_carga_{e}", disabled=n == 0,
                         use_container_width=True):
                ir_a_detalle("t4", PREFIJO_CARGA + e, "Todos")

    with st.container(border=True):
        titulo("Carga por territorio", "Porcentaje de sitios de cada territorio por estado. "
               "Haz clic en un tramo para ver sus sitios.")
        tc = pd.crosstab(dc["Territorio"], dc["Estado carga"]).reindex(
            index=sel_terr, columns=estados, fill_value=0)
        fig_c = go.Figure()
        for e in estados:
            n_e = tc[e].values
            tot_t = tc.sum(axis=1).replace(0, 1).values
            p_e = n_e / tot_t * 100
            fig_c.add_bar(
                name=e, y=list(tc.index), x=p_e, orientation="h",
                marker=dict(color=COLOR_CARGA.get(e, GRIS), line=dict(width=0)),
                text=[f"{p:.0f}%" if p >= 7 else "" for p in p_e], textposition="inside",
                insidetextanchor="middle", textfont=dict(size=13, color="#ffffff"),
                customdata=n_e,
                hovertemplate="<b>%{y}</b> · " + e + "<br>%{customdata} sitios · %{x:.1f}%<extra></extra>")
        fig_c.update_layout(barmode="stack", bargap=.35, clickmode="event+select")
        layout_base(fig_c, 120 + 52 * len(tc))
        fig_c.update_xaxes(range=[0, 100], ticksuffix="%", showgrid=False)
        fig_c.update_yaxes(autorange="reversed", showgrid=False)
        pto = leer_click(grafico_clic(fig_c, "clk_carga"), "clk_carga")
        if pto:
            i = pto.get("curve_number")
            if isinstance(i, int) and 0 <= i < len(estados) and pto.get("y") in sel_terr:
                ir_a_detalle("t4", PREFIJO_CARGA + estados[i], pto["y"])

        with st.expander("Ver números por territorio"):
            tabla_c = tc.copy()
            tabla_c["Total"] = tabla_c.sum(axis=1)
            st.dataframe(tabla_c, use_container_width=True)


# ---- Tab 4: Tabla ----------------------------------------------------------
with tab4:
    st.subheader(f"Hoja Resumen · {fuente}")
    mostrar = dfv.copy()
    if h_total:
        mostrar["% del total"] = (mostrar["Total"] / max(total, 1) * 100).round(1)

    # Celdas que difieren del Excel (solo tiene sentido en la vista recalculada)
    celdas_dif = set()
    if fuente == FUENTE_RECALC:
        todos = set(sel_terr) == set(territorios)
        for d in diferencias:
            if d["Columna"] in sel_terr or (d["Columna"] == "Total" and todos):
                celdas_dif.add((d["Hito"], d["Columna"]))

    def resaltar(data):
        estilos = pd.DataFrame("", index=data.index, columns=data.columns)
        for h, c in celdas_dif:
            if h in estilos.index and c in estilos.columns:
                estilos.loc[h, c] = ("background-color: #eda100; color: #1a1a19; "
                                     "font-weight: 700;")
        return estilos

    vista = st.radio("Vista", ["🗂️ Tarjetas", "📋 Tabla"], horizontal=True,
                     key="vista_resumen", label_visibility="collapsed")
    corto = lambda t: t.replace("Bogota ", "").replace("Bogotá ", "")  # noqa: E731

    if vista.startswith("🗂️"):
        hitos_t = list(df.index)
        n_col = 4
        for i0 in range(0, len(hitos_t), n_col):
            fila_c = st.columns(n_col)
            for cc, h in zip(fila_c, hitos_t[i0:i0 + n_col]):
                n = int(dfv.loc[h, "Total"])
                es_total_h = h == h_total
                color = (GRIS if es_total_h else VERDE if h == h_insrv
                         else "#d6453d" if h == h_neg else "#5598e7" if h == h_cos else AZUL)
                avance = 100 if es_total_h else pct(n, total)
                dif_h = sorted({c for (hh, c) in celdas_dif if hh == h})
                aviso = (f'<div style="font-size:.75rem;color:#b07a00;margin-top:2px">'
                         f'⚠ corregido vs. Excel ({", ".join(corto(x) for x in dif_h)})</div>'
                         if dif_h else "")
                por_terr = " · ".join(
                    f'<span style="color:{color_terr.get(t, GRIS)}">●</span> {corto(t)} '
                    f'<b>{int(dfv.loc[h, t])}</b>' for t in sel_terr
                    if t != SIN_TERRITORIO or int(dfv.loc[h, t]) > 0)
                with cc, st.container(border=True):
                    html(f"""<div style="border-left:5px solid {color};padding-left:10px">
                               <div style="font-size:.85rem;opacity:.75">{h}</div>
                               <div style="font-size:1.9rem;font-weight:700;line-height:1.2">{n:,}</div>
                               <div style="font-size:.8rem;opacity:.65">
                                 {"sitios del proyecto" if es_total_h else f"{avance:.0f}% de {total:,} sitios"}</div>
                               <div style="height:5px;border-radius:3px;background:rgba(128,128,128,.2);margin:6px 0">
                                 <div style="height:5px;border-radius:3px;width:{min(avance, 100):.0f}%;background:{color}"></div></div>
                               <div style="font-size:.72rem;opacity:.8;line-height:1.5">{por_terr}</div>
                               {aviso}
                             </div>""")
                    if det is not None and (not presentacion) and st.button("Ver sitios", key=f"btn_res_{h}",
                                                     use_container_width=True):
                        ir_a_detalle("t4", h, "Todos")
        st.caption("Cada tarjeta muestra el total, el avance sobre el total de sitios y el "
                   "desglose por territorio. Para ver un territorio específico, cambia a la "
                   "vista 📋 Tabla y haz clic en la casilla.")
    else:
        ev_tabla = tabla_clicable(
            mostrar.style
            .background_gradient(cmap="Blues", subset=sel_terr + ["Total"], axis=None)
            .apply(resaltar, axis=None)
            .format({"% del total": "{:.1f}%"} if h_total else {}),
            "clk_tabla", use_container_width=True, height=470,
        )
        clic_en_tabla(ev_tabla, "clk_tabla", mostrar, "t4", "Cumplen")
        if det is not None:
            pista("👆 Haz clic en una casilla: se abre una ventana con los sitios de ese hito "
                       "y territorio. La columna *Total* muestra todos los territorios.")
            if st.session_state.get("_sin_clic_clk_tabla"):
                st.warning("Esta versión de Streamlit no permite hacer clic en casillas "
                           f"(tengo la {getattr(st, '__version__', '?')}; hace falta 1.49 o más). "
                           "Actualiza `requirements.txt` y reinicia la app (Manage app → Reboot). "
                           "Mientras tanto usa el panel 🔍 de abajo. "
                           f"Detalle: {st.session_state['_sin_clic_clk_tabla']}")
            if not presentacion:
                with st.expander("🛠 Diagnóstico del clic"):
                    st.caption(f"Streamlit {getattr(st, '__version__', '?')} · casilla recibida: "
                               f"{st.session_state.get('_dbg_clk_tabla', [])}")
    if celdas_dif:
        st.caption(f"🟨 {len(celdas_dif)} celda(s) resaltada(s): el valor corregido es distinto "
                   f"al de la pestaña Resumen del Excel. Detalle en 🔧 Correcciones.")
    _dest = {(list(mostrar.index).index(h), c) for (h, c) in celdas_dif if h in mostrar.index}
    _tab = mostrar.rename_axis("Hito")
    if not presentacion:
        st.download_button(
            "⬇️ Descargar Excel",
            excel_bytes(_tab, "Resumen", f"Resumen · {fuente}", indice=True, destacar=_dest),
            file_name=f"resumen_{'recalculado' if fuente == FUENTE_RECALC else 'excel'}.xlsx",
            mime=XLSX_MIME,
        )
    if not presentacion:
        with st.expander("🔍 Explorar sitios por hito, territorio o etapa"):
            panel_detalle("t4")

# ---- Tab 5: Correcciones ---------------------------------------------------
def vista_correcciones_excel():
    st.subheader("Diferencias entre el Resumen del Excel y el recálculo")
    if df_recalc is None:
        st.info(error_recalc)
    elif not diferencias:
        st.success("El Resumen del Excel coincide con el recálculo desde la pestaña Sitios.")
    else:
        st.caption("El recálculo cuenta los sitios de la pestaña **Sitios** del proyecto "
                   f"**{total_col}** que tienen fecha en cada hito, igual que las fórmulas "
                   "del Resumen, pero sin sus errores de referencia. El Excel no se modifica.")

        def motivo(d):
            if d["Columna"] == SIN_TERRITORIO:
                return "Sitio sin región en Sitios: el Excel no lo cuenta en ningún territorio"
            if d["Columna"] == "Total":
                return "El total del Excel no coincide con la suma de territorios"
            return "Fórmula del Resumen con referencia errada (#REF!, otro territorio u otra columna)"

        tabla = pd.DataFrame(diferencias)
        tabla["Diferencia"] = tabla["Recalculado"] - tabla["Excel"]
        tabla["Motivo"] = tabla.apply(motivo, axis=1)
        st.dataframe(tabla, hide_index=True, use_container_width=True)

    if sin_terr is not None and len(sin_terr):
        st.markdown(f"**Sitios del proyecto sin región ({len(sin_terr)})** — corregir en la "
                    "pestaña `moder` o en la columna B de `Sitios`:")
        st.dataframe(sin_terr, hide_index=True, use_container_width=True)

    h_rfic = buscar(df, "RFIC-OK")
    if (h_rfic and df_recalc is not None and criterio_rfic == RFIC_COMBINADO
            and int(df_recalc.loc[h_rfic, "Total"]) == 0):
        st.info("**RFIC-OK está en 0:** ningún sitio tiene fecha en TSS OK, obra civil y "
                "diseño TX a la vez.")
    elif h_rfic and df_recalc is not None and int(df_recalc.loc[h_rfic, "Total"]) == 0:
        st.info("**RFIC-OK sigue en 0 aunque ya está bien calculado:** en Sitios un sitio "
                "solo cuenta si **todos** sus trabajos tienen fecha de RFIC, y casi ningún "
                "trabajo la tiene en la pestaña RF. Es un tema de los datos de origen, "
                "no de la fórmula.")


def vista_calidad():
    st.subheader("Calidad de los datos cargados")
    c1, c2, c3 = st.columns(3)
    c1.metric("Sitios en la lista", f"{calidad['n_sitios']:,}")
    c2.metric("Trabajos de RF considerados", f"{calidad['n_trabajos']:,}")
    c3.metric("Sitios sin trabajos en RF", f"{len(calidad['sin_trabajos']):,}")
    st.caption("MasterPlan incluidos: " + ", ".join(calidad["planes"]) + ". "
               "Un sitio cumple un hito cuando **todos** sus trabajos de esos MasterPlan tienen "
               "fecha en la columna correspondiente de RF.")

    if len(sin_terr):
        st.markdown(f"**Sitios sin territorio ({len(sin_terr)})** — revisa el nombre en "
                    "baseline o la columna *Territorio Comercial*:")
        st.dataframe(sin_terr, hide_index=True, use_container_width=True)
    if len(calidad["sin_trabajos"]):
        st.markdown(f"**Sitios sin trabajos en RF ({len(calidad['sin_trabajos'])})** — no "
                    "aparecen en RF con los MasterPlan elegidos (pueden estar bajo otro plan "
                    "o faltar en el export):")
        st.dataframe(calidad["sin_trabajos"], hide_index=True, use_container_width=True)
    if calidad["faltantes"]:
        st.warning("Columnas que no encontré en RF: " + ", ".join(calidad["faltantes"]))
    if not len(sin_terr) and not len(calidad["sin_trabajos"]) and not calidad["faltantes"]:
        st.success("Todos los sitios tienen territorio y trabajos en RF.")

    h_rfic = buscar(df, "RFIC-OK")
    if h_rfic and int(df.loc[h_rfic, "Total"]) == 0 and criterio_rfic == RFIC_COMBINADO:
        st.info("**RFIC-OK está en 0:** ningún sitio tiene TSS ok, CW Lista y Orden Tx con "
                "fecha en todos sus trabajos.")
    elif h_rfic and int(df.loc[h_rfic, "Total"]) == 0:
        st.info("**RFIC-OK está en 0:** casi ningún trabajo tiene fecha de RFIC en RF, y un "
                "sitio solo cuenta cuando todos sus trabajos la tienen. Es un tema de los "
                "datos de origen.")


if tab5 is not None:
    with tab5:
        if modo_arch:
            vista_calidad()
        else:
            vista_correcciones_excel()

"""
Dashboard Resumen - Modernización RF (R4 Bogotá)
Dos formas de cargar los datos:
  1) Lista de sitios + export de RF + export de baseline (la app calcula todo), o
  2) El Excel completo de control (lee la pestaña "Resumen" y recalcula desde "Sitios").

Ejecutar:
    pip install -r requirements.txt
    streamlit run app.py
"""

import base64
import datetime
import io
import json
import zipfile
from pathlib import Path
import unicodedata
from zoneinfo import ZoneInfo

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
      .block-container {padding-top: 4.2rem; padding-bottom: 2rem;}
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


def es_nokia(v) -> bool:
    """True si la celda es texto y menciona a Nokia (p. ej. «Nokia», «NOKIA»)."""
    return isinstance(v, str) and "nokia" in v.lower()


def _convertir_fechas_texto(df: pd.DataFrame) -> pd.DataFrame:
    """Los exports de texto (csv/html) traen las fechas como texto: se convierten a fecha."""
    conocidas = {n for nombres in HITO_RF.values() for n in nombres}
    for c in df.columns:
        if _plano(c) == "cw lista":
            # «Nokia» en CW Lista cuenta como obra civil lista aunque no tenga fecha
            df["__cw_nokia"] = df[c].map(es_nokia)
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


def cw_nokia(r: pd.DataFrame, col) -> pd.Series:
    """Trabajos cuya CW Lista dice «Nokia» (en csv/html se guardó aparte al leer)."""
    if "__cw_nokia" in r.columns:
        return r["__cw_nokia"].fillna(False).astype(bool)
    return r[col].map(es_nokia)


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
        f["v"] = f["f"].notna()
        if clave_h == "OBRA CIVIL OK":
            f["v"] = f["v"] | cw_nokia(r, col).values
        ag = f.groupby("_k").agg(count=("v", "sum"), max=("f", "max"))
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
            v3 = f3.notna()
            v3[1] = v3[1] | cw_nokia(r, cols3[1])  # CW Lista: «Nokia» cuenta como OK
            job_ok = v3.all(axis=1)
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
# Datos publicados: quien administra sube los archivos una vez y los demás solo ven
# ---------------------------------------------------------------------------
DIR_PUB = Path(__file__).parent / "datos_publicados"
RUTA_GH = "datos_publicados"


@st.cache_resource
def _memoria_pub() -> dict:
    return {}


class ArchivoPub:
    """Archivo publicado, con la misma interfaz que un archivo subido (name, getvalue)."""

    def __init__(self, nombre, datos):
        self.name, self._d = nombre, datos

    def getvalue(self):
        return self._d


def _ahora() -> str:
    return datetime.datetime.now(ZoneInfo("America/Bogota")).strftime("%d/%m/%Y %H:%M")


# --- GitHub como almacén (rama aparte: no reinicia la app al guardar) -----------------
def gh_cfg():
    """Configuración de GitHub desde los Secrets; None si no está configurado."""
    try:
        tok, repo = st.secrets.get("GITHUB_TOKEN"), st.secrets.get("GITHUB_REPO")
        if not tok or not repo:
            return None
        return {"token": tok, "repo": repo, "rama": st.secrets.get("GITHUB_RAMA_DATOS", "datos"),
                "api": st.secrets.get("GITHUB_API", "https://api.github.com")}
    except Exception:  # noqa: BLE001
        return None


def _gh(cfg, metodo, ruta, crudo=False, **kw):
    import requests
    h = {"Authorization": f"Bearer {cfg['token']}", "X-GitHub-Api-Version": "2022-11-28",
         "Accept": "application/vnd.github.raw+json" if crudo else "application/vnd.github+json"}
    return requests.request(metodo, f"{cfg['api']}/repos/{cfg['repo']}/{ruta}", headers=h,
                            timeout=120, **kw)


def gh_leer(cfg, ruta):
    r = _gh(cfg, "GET", f"contents/{ruta}", crudo=True, params={"ref": cfg["rama"]})
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.content


def gh_asegurar_rama(cfg):
    r = _gh(cfg, "GET", f"git/ref/heads/{cfg['rama']}")
    if r.status_code == 200:
        return
    base = _gh(cfg, "GET", "").json().get("default_branch", "main")
    sha = _gh(cfg, "GET", f"git/ref/heads/{base}").json()["object"]["sha"]
    _gh(cfg, "POST", "git/refs", json={"ref": f"refs/heads/{cfg['rama']}", "sha": sha}
        ).raise_for_status()


def gh_escribir(cfg, ruta, datos: bytes, mensaje: str):
    r = _gh(cfg, "GET", f"contents/{ruta}", params={"ref": cfg["rama"]})
    cuerpo = {"message": mensaje, "content": base64.b64encode(datos).decode(),
              "branch": cfg["rama"]}
    if r.status_code == 200:
        cuerpo["sha"] = r.json().get("sha")
    _gh(cfg, "PUT", f"contents/{ruta}", json=cuerpo).raise_for_status()


def _decodificar_pub(b: bytes) -> str:
    try:
        return b.decode("utf-8-sig")
    except UnicodeDecodeError:
        return b.decode("latin-1")


def pub_meta() -> dict:
    import time
    mem = _memoria_pub()
    cfg = gh_cfg()
    if cfg:
        if time.time() - mem.get("_t", 0) > 20:  # se refresca desde GitHub cada 20 s
            try:
                crudo = gh_leer(cfg, f"{RUTA_GH}/meta.json")
                mem["meta"] = json.loads(crudo) if crudo else {}
            except Exception:  # noqa: BLE001
                mem.setdefault("meta", {})
            mem["_t"] = time.time()
        return mem.get("meta", {})
    if "meta" in mem:
        return mem["meta"]
    try:
        meta = json.loads((DIR_PUB / "meta.json").read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        meta = {}
    return meta


def pub_archivo(clave):
    meta = pub_meta()
    nombre = meta.get("archivos", {}).get(clave)
    if not nombre:
        return None
    mem = _memoria_pub()
    version = f"{meta.get('publicado', '')}|{meta.get('versiones', {}).get(clave, '')}|{nombre}"
    guardado = mem.get("f_" + clave)
    if guardado and guardado[0] == version:
        return ArchivoPub(nombre, guardado[1])
    datos = None
    cfg = gh_cfg()
    try:
        if cfg:
            datos = gh_leer(cfg, f"{RUTA_GH}/{nombre}")
        else:
            datos = mem.get(clave) or (DIR_PUB / nombre).read_bytes()
    except Exception:  # noqa: BLE001
        datos = None
    if datos is None:
        return None
    mem["f_" + clave] = (version, datos)
    return ArchivoPub(nombre, datos)


def pub_guardar(entradas: dict, extra: dict):
    """entradas: {clave: (nombre_original, bytes)}. Con GitHub configurado se guarda allí
    (queda permanente); si no, en memoria y en disco (se pierde al reiniciar la app)."""
    import hashlib
    mem = _memoria_pub()
    meta = json.loads(json.dumps(pub_meta())) if pub_meta() else {}
    meta.setdefault("archivos", {})
    meta.setdefault("versiones", {})
    cfg = gh_cfg()
    aviso = None
    if cfg:
        try:
            gh_asegurar_rama(cfg)
        except Exception as e:  # noqa: BLE001
            return f"No pude preparar la rama de datos en GitHub: {e}"
    else:
        try:
            DIR_PUB.mkdir(exist_ok=True)
        except Exception as e:  # noqa: BLE001
            aviso = str(e)
    for clave, (nombre, datos) in entradas.items():
        guardado = f"{clave}{Path(nombre).suffix.lower() or '.dat'}"
        meta["archivos"][clave] = guardado
        meta["versiones"][clave] = hashlib.sha1(datos).hexdigest()[:10]
        try:
            if cfg:
                gh_escribir(cfg, f"{RUTA_GH}/{guardado}", datos, f"Datos publicados: {guardado}")
            else:
                (DIR_PUB / guardado).write_bytes(datos)
        except Exception as e:  # noqa: BLE001
            aviso = f"{guardado}: {e}"
    meta.update(extra)
    meta["publicado"] = _ahora()
    for clave in entradas:
        mem.pop("f_" + clave, None)
        mem[clave] = entradas[clave][1]
    mem["meta"] = meta
    mem["_t"] = __import__("time").time()
    texto = json.dumps(meta, ensure_ascii=False, indent=1).encode("utf-8")
    try:
        if cfg:
            gh_escribir(cfg, f"{RUTA_GH}/meta.json", texto, "Datos publicados")
        else:
            (DIR_PUB / "meta.json").write_bytes(texto)
    except Exception as e:  # noqa: BLE001
        aviso = str(e)
    return aviso


def pub_borrar():
    mem = _memoria_pub()
    cfg = gh_cfg()
    if cfg:
        try:
            gh_escribir(cfg, f"{RUTA_GH}/meta.json", b"{}", "Datos publicados quitados")
        except Exception:  # noqa: BLE001
            pass
    else:
        for nombre in list(pub_meta().get("archivos", {}).values()) + ["meta.json"]:
            try:
                (DIR_PUB / nombre).unlink()
            except Exception:  # noqa: BLE001
                pass
    mem.clear()
    mem["meta"] = {}
    mem["_t"] = __import__("time").time()


def solicitar_actualizacion():
    """Pide al computador (agente_pc.py) que suba los archivos más recientes y espera la
    respuesta. Solo con GitHub configurado."""
    import time
    import uuid
    cfg = gh_cfg()
    sid = uuid.uuid4().hex[:10]
    try:
        gh_asegurar_rama(cfg)
        gh_escribir(cfg, f"{RUTA_GH}/solicitud.json",
                    json.dumps({"id": sid, "hora": _ahora()}).encode("utf-8"),
                    "Solicitud de actualización")
    except Exception as e:  # noqa: BLE001
        st.error(f"No pude enviar la solicitud: {e}")
        return
    espera = 180
    barra = st.progress(0.0, text="Pidiendo los datos al computador…")
    t0 = time.time()
    while time.time() - t0 < espera:
        time.sleep(5)
        try:
            crudo = gh_leer(cfg, f"{RUTA_GH}/meta.json")
            meta = json.loads(crudo) if crudo else {}
        except Exception:  # noqa: BLE001
            meta = {}
        if meta.get("solicitud") == sid:
            mem = _memoria_pub()
            for k in [k for k in mem if str(k).startswith("f_")]:
                mem.pop(k)
            mem["meta"], mem["_t"] = meta, time.time()
            barra.progress(1.0, text="Datos actualizados.")
            st.rerun()
        barra.progress(min((time.time() - t0) / espera, 0.99),
                       text="Esperando que el computador envíe los archivos…")
    barra.empty()
    st.warning("El computador no respondió en 3 minutos. Revisa que esté encendido y con el "
               "agente abierto. Mientras tanto se muestran los últimos datos publicados.")


def pub_zip() -> bytes:
    """Paquete con la carpeta datos_publicados para subirla a GitHub y que no se pierda."""
    buf = io.BytesIO()
    meta = pub_meta()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for clave, nombre in meta.get("archivos", {}).items():
            f = pub_archivo(clave)
            if f is not None:
                z.writestr(f"datos_publicados/{nombre}", f.getvalue())
        z.writestr("datos_publicados/meta.json", json.dumps(meta, ensure_ascii=False, indent=1))
    return buf.getvalue()


def es_administrador() -> bool:
    """Sin ADMIN_PASSWORD en los Secrets, todos pueden publicar. Con clave, solo quien la sepa."""
    try:
        clave = st.secrets.get("ADMIN_PASSWORD")
    except Exception:  # noqa: BLE001
        clave = None
    if not clave:
        return True
    if st.session_state.get("_admin_ok"):
        return True
    with st.sidebar.expander("🔐 Administrador"):
        intento = st.text_input("Clave", type="password", key="clave_admin")
        if intento and intento == clave:
            st.session_state["_admin_ok"] = True
            st.rerun()
        elif intento:
            st.error("Clave incorrecta.")
    return False


def bloque_publicar(entradas: dict, extra: dict):
    """Botones de publicación (solo administradores)."""
    with st.sidebar:
        st.header("📢 Publicar")
        meta = pub_meta()
        if meta.get("publicado"):
            st.caption(f"Publicado el {meta['publicado']}. Quien abra el enlace verá estos datos "
                       "sin subir nada.")
        if entradas and st.button("📢 Publicar para todos", key="btn_publicar", width="stretch",
                     help="Guarda los archivos actuales en el servidor para que cualquiera que "
                          "abra el enlace vea la información sin subir archivos."):
            aviso = pub_guardar(entradas, extra)
            st.success("Publicado. Ya lo ven todos los que abran el enlace.")
            if aviso:
                st.warning("Quedó en la memoria de la app, pero no pude escribir en disco "
                           f"({aviso}).")
        if meta.get("archivos"):
            st.download_button(
                "⬇️ Paquete para GitHub (.zip)", pub_zip(), file_name="datos_publicados.zip",
                mime="application/zip", key="dl_pub_zip",
                help="Descomprime y sube la carpeta datos_publicados a tu repositorio de GitHub: "
                     "así la información se conserva aunque la app se reinicie.")
            if st.button("🗑️ Quitar datos publicados", key="btn_quitar_pub", width="stretch"):
                pub_borrar()
                st.rerun()


# ---------------------------------------------------------------------------
# Pantalla aparte: Concepto TSS
# ---------------------------------------------------------------------------
CT_SIN_TERR = "Sin territorio"
CT_ORDEN_AREA = ["RF", "Obra Civil", "Implementacion"]
CT_COLS_SALIDA = ["Sitio", "Días hábiles", "Área sin concepto", "Territorio"]
CT_AZUL_XL = "0070C0"     # encabezado, igual al del Excel de referencia
CT_ZEBRA_XL = "DDEBF7"


# ---------------------------------------------------------------------------
# Lectura de archivos (xlsx, xls real, csv o tabla HTML disfrazada de .xls)
# ---------------------------------------------------------------------------
def _ct_decodificar(b: bytes) -> str:
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            continue
    return b.decode("utf-8", errors="ignore")


@st.cache_data(show_spinner="Leyendo archivo…")
def ct_leer(contenido: bytes, nombre: str, requeridas: tuple) -> pd.DataFrame:
    req = [_plano(r) for r in requeridas]

    def cumple(cols):
        c = {_plano(x) for x in cols}
        return all(r in c for r in req)

    if contenido[:4] == b"\xd0\xcf\x11\xe0" and (
            b"E\x00n\x00c\x00r\x00y\x00p\x00t\x00e\x00d\x00P\x00a\x00c\x00k\x00a\x00g\x00e\x00"
            in contenido):
        raise ValueError("El archivo está protegido con contraseña o permisos. Guárdalo sin "
                         "protección o súbelo como export directo del sistema.")
    if contenido[:4] == b"PK\x03\x04" or contenido[:4] == b"\xd0\xcf\x11\xe0":
        x = pd.ExcelFile(io.BytesIO(contenido))
        for h in x.sheet_names:
            crudo = pd.read_excel(x, sheet_name=h, header=None, nrows=15)
            for i in range(len(crudo)):
                if cumple([c for c in crudo.iloc[i].tolist() if pd.notna(c)]):
                    return pd.read_excel(x, sheet_name=h, header=i)
    else:
        texto = _ct_decodificar(contenido)
        if "<table" in texto[:200000].lower():
            for t in pd.read_html(io.StringIO(texto), header=0):
                if cumple(t.columns):
                    return t
        else:
            t = pd.read_csv(io.StringIO(texto), sep=None, engine="python")
            if cumple(t.columns):
                return t
    raise ValueError(f"No encontré las columnas {', '.join(requeridas)} en '{nombre}'. "
                     "Revisa que sea el reporte correcto.")


def ct_col(df, nombre):
    m = {_plano(c): c for c in df.columns}
    return m.get(_plano(nombre))


# ---------------------------------------------------------------------------
# Procesamiento
# ---------------------------------------------------------------------------
def ct_dias_habiles(desde: pd.Series, hasta: datetime.date) -> pd.Series:
    """Igual que =DIAS.LAB(fecha; AHORA()) - 1 del Excel: lunes a viernes, ambos días incluidos."""
    d = desde.dt.normalize()
    ok = d.notna()
    salida = pd.Series(np.nan, index=desde.index)
    fin = np.datetime64(hasta) + np.timedelta64(1, "D")
    salida[ok] = np.busday_count(d[ok].values.astype("datetime64[D]"), fin) - 1
    return salida.clip(lower=0)


def ct_procesar(rep: pd.DataFrame, base: pd.DataFrame | None, region: str, corte: datetime.date,
             solo_fuera: bool):
    c_reg, c_sit = ct_col(rep, "Region"), ct_col(rep, "Sitio")
    c_dia, c_area = ct_col(rep, "Dia TSS"), ct_col(rep, "Area sin concepto")
    d = rep[[c_reg, c_sit, c_dia, c_area]].copy()
    d.columns = ["Region", "Sitio", "Dia TSS", "Área sin concepto"]
    d = d.dropna(subset=["Sitio"])
    d["Sitio"] = d["Sitio"].astype(str).str.strip()
    d["Área sin concepto"] = d["Área sin concepto"].astype(str).str.strip()
    if region != "Todas":
        d = d[d["Region"].map(_plano) == _plano(region)]
    d["_f"] = pd.to_datetime(d["Dia TSS"], dayfirst=True, format="mixed", errors="coerce")
    d["Días hábiles"] = ct_dias_habiles(d["_f"], corte)
    n_filas = len(d)
    # Un solo registro por sitio y área: se queda el más antiguo (más días)
    d["_k"] = d["Sitio"].map(_plano)  # "Rb" y "RB" son el mismo sitio
    d = (d.sort_values("Días hábiles", ascending=False, na_position="last")
          .drop_duplicates(["_k", "Área sin concepto"]))
    if solo_fuera:
        d = d[d["Días hábiles"] > 5]
    if base is not None:
        cs, ct = ct_col(base, "Sitio"), ct_col(base, "Territorio Comercial")
        mp = (base[[cs, ct]].dropna(subset=[cs]).assign(k=lambda x: x[cs].map(_plano))
              .drop_duplicates("k").set_index("k")[ct])
        d["Territorio"] = d["_k"].map(mp).fillna(CT_SIN_TERR)
        nombre = (base[[cs]].dropna().assign(k=lambda x: x[cs].map(_plano))
                  .drop_duplicates("k").set_index("k")[cs])
        d["Sitio"] = d["_k"].map(nombre).fillna(d["Sitio"])  # escritura oficial del baseline
    else:
        d["Territorio"] = ""
    orden = {a: i for i, a in enumerate(CT_ORDEN_AREA)}
    d["_o"] = d["Área sin concepto"].map(lambda a: orden.get(a, 99))
    d = d.sort_values(["_o", "Días hábiles", "Sitio"], ascending=[True, False, True])
    d["Días hábiles"] = d["Días hábiles"].fillna(0).astype(int)
    return d[CT_COLS_SALIDA].reset_index(drop=True), n_filas


# ---------------------------------------------------------------------------
# Salidas: Excel e imagen
# ---------------------------------------------------------------------------
def ct_excel_bytes(tabla: pd.DataFrame, titulo: str) -> bytes:
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        tabla.to_excel(w, sheet_name="Concepto TSS", index=False, startrow=2)
        ws = w.sheets["Concepto TSS"]
        ws.cell(1, 1, titulo).font = Font(bold=True, size=14, color="1F3864")
        linea = Side(style="thin", color="BFBFBF")
        borde = Border(left=linea, right=linea, top=linea, bottom=linea)
        for j, c in enumerate(tabla.columns, start=1):
            h = ws.cell(3, j)
            h.font = Font(bold=True, color="FFFFFF")
            h.fill = PatternFill("solid", fgColor=CT_AZUL_XL)
            h.alignment = Alignment(horizontal="center", vertical="center")
            h.border = borde
            ancho = max([len(str(c))] + [len(str(x)) for x in tabla[c]])
            ws.column_dimensions[get_column_letter(j)].width = min(max(ancho + 4, 12), 45)
            for i in range(len(tabla)):
                cel = ws.cell(4 + i, j)
                cel.border = borde
                cel.alignment = Alignment(
                    horizontal="right" if c == "Días hábiles" else
                    "center" if c == "Territorio" else "left", vertical="center")
                if i % 2 == 0:
                    cel.fill = PatternFill("solid", fgColor=CT_ZEBRA_XL)
        ws.row_dimensions[3].height = 22
        ws.freeze_panes = "A4"
        if len(tabla):
            ws.auto_filter.ref = f"A3:{get_column_letter(len(tabla.columns))}{3 + len(tabla)}"
        ws.sheet_view.showGridLines = False
    return buf.getvalue()


def _ct_fuente(tam: int, negrita: bool = False):
    import os

    import matplotlib
    from PIL import ImageFont
    base = os.path.join(matplotlib.get_data_path(), "fonts", "ttf")
    candidatos = (
        ["/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf",
         "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
         os.path.join(base, "DejaVuSans-Bold.ttf")] if negrita else
        ["/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf",
         "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
         os.path.join(base, "DejaVuSans.ttf")])
    for ruta in candidatos:
        if os.path.exists(ruta):
            return ImageFont.truetype(ruta, tam)
    return ImageFont.load_default()


CT_ANCHO_COL = {"Sitio": 430, "Días hábiles": 250, "Área sin concepto": 300, "Territorio": 310}


def ct_imagen_tabla(tabla: pd.DataFrame, titulo: str | None = None) -> bytes:
    """Tabla estilo Excel (encabezado azul, filas alternadas) como PNG para WhatsApp."""
    from PIL import Image, ImageDraw
    cols = list(tabla.columns)
    anchos = [CT_ANCHO_COL.get(c, 300) for c in cols]
    alto_fila, alto_enc = 58, 62
    alto_tit = 76 if titulo else 0
    ancho = sum(anchos)
    img = Image.new("RGB", (ancho, alto_tit + alto_enc + alto_fila * len(tabla) + 2), "white")
    d = ImageDraw.Draw(img)
    f_enc, f_txt, f_tit = _ct_fuente(28, True), _ct_fuente(27), _ct_fuente(38, True)
    if titulo:
        d.text((16, alto_tit // 2), titulo, font=f_tit, fill=(31, 56, 100), anchor="lm")
    y0 = alto_tit
    d.rectangle((0, y0, ancho, y0 + alto_enc), fill=(0, 112, 192))
    x = 0
    for nombre, w in zip(cols, anchos):
        d.text((x + w // 2 - 14, y0 + alto_enc // 2), nombre, font=f_enc, fill="white", anchor="mm")
        bx = x + w - 44  # botón de filtro, como en Excel
        d.rectangle((bx, y0 + 16, bx + 30, y0 + 46), fill=(240, 240, 240), outline=(190, 190, 190))
        d.polygon([(bx + 8, y0 + 27), (bx + 22, y0 + 27), (bx + 15, y0 + 37)], fill=(60, 60, 60))
        x += w
    for i, fila in enumerate(tabla.itertuples(index=False)):
        y = y0 + alto_enc + i * alto_fila
        d.rectangle((0, y, ancho, y + alto_fila), fill=(221, 235, 247) if i % 2 == 0 else "white")
        d.line((0, y + alto_fila, ancho, y + alto_fila), fill=(200, 215, 230), width=1)
        x = 0
        for nombre, valor, w in zip(cols, fila, anchos):
            txt = str(valor)
            cy = y + alto_fila // 2
            if nombre == "Días hábiles":
                d.text((x + w - 96, cy), txt, font=f_txt, fill="black", anchor="rm")
            elif nombre == "Sitio":
                d.text((x + 12, cy), txt, font=f_txt, fill="black", anchor="lm")
            else:
                d.text((x + w // 2, cy), txt, font=f_txt, fill="black", anchor="mm")
            x += w
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def ct_boton_copiar(png: bytes):
    """Botón que copia la imagen al portapapeles para pegarla directo en WhatsApp."""
    try:
        import streamlit.components.v1 as components
        b64 = base64.b64encode(png).decode()
        components.html(
            """<style>button{font:14px sans-serif;padding:7px 14px;border:1px solid #bbb;
            border-radius:8px;background:#fff;cursor:pointer}button:hover{background:#f3f3f3}
            span{font:13px sans-serif;margin-left:10px;color:#444}</style>
            <button id="b">📋 Copiar imagen</button><span id="m"></span>
            <script>
            var b64="%s", m=document.getElementById('m');
            document.getElementById('b').onclick=async function(){
              try{
                var bin=atob(b64), u=new Uint8Array(bin.length);
                for(var i=0;i<bin.length;i++)u[i]=bin.charCodeAt(i);
                await navigator.clipboard.write([new ClipboardItem({'image/png':
                  new Blob([u],{type:'image/png'})})]);
                m.textContent='Copiada. Pégala en WhatsApp con Ctrl+V';
              }catch(e){
                m.textContent='No se pudo copiar aquí: clic derecho en la imagen → Copiar imagen';
              }
            };
            </script>""" % b64, height=46)
    except Exception:  # noqa: BLE001
        st.caption("Clic derecho sobre la imagen → Copiar imagen.")


def ct_tabla_html(tabla: pd.DataFrame) -> str:
    """Tabla HTML (texto seleccionable) para copiar y pegar en Excel."""
    import html as _h
    th = "".join(f"<th>{_h.escape(str(c))}</th>" for c in tabla.columns)
    filas = "".join("<tr>" + "".join(f"<td>{_h.escape(str(v))}</td>" for v in f) + "</tr>"
                    for f in tabla.itertuples(index=False))
    return (f"<table class='ctx'><thead><tr>{th}</tr></thead><tbody>{filas}</tbody></table>")


def ct_texto_whatsapp(tabla: pd.DataFrame, area: str) -> str:
    """Versión en texto con formato de WhatsApp (negrita y monoespaciado)."""
    cols = [c for c in tabla.columns if c != "Área sin concepto"]
    filas = [[str(v) for v in f] for f in tabla[cols].itertuples(index=False)]
    w = [max(len(c), *(len(f[i]) for f in filas)) for i, c in enumerate(cols)]
    lin = lambda f: "  ".join(x.ljust(w[i]) for i, x in enumerate(f)).rstrip()  # noqa: E731
    return f"*{area}*\n```\n" + "\n".join([lin(cols)] + [lin(f) for f in filas]) + "\n```"


def pantalla_concepto_tss():
    """Pantalla aparte: Concepto TSS pendiente por sitio y área."""
    st.title("📨 Concepto TSS · sitios pendientes por área")
    st.caption("Sube el reporte de Concepto TSS: se filtra la región, queda un solo registro por "
               "sitio y área, y se generan el Excel y la imagen para WhatsApp.")

    meta_p = pub_meta()
    hay_pub = "concepto" in meta_p.get("archivos", {})
    f_rep = f_bl = None
    if ES_ADMIN or not hay_pub:
        c1, c2 = st.columns(2)
        f_rep = c1.file_uploader("1 · Reporte Concepto TSS (.xls, .xlsx, .csv, .htm, .html)",
                                 type=["xls", "xlsx", "csv", "htm", "html"], key="ct_rep")
        f_bl = c2.file_uploader("2 · Baseline (para el territorio, opcional; .xls, .htm, .html…)",
                                type=["xls", "xlsx", "csv", "htm", "html"], key="ct_bl",
                                help="Se usa solo para poner el territorio comercial de cada sitio.")
    subido_rep, subido_bl = f_rep, f_bl
    f_rep = f_rep or pub_archivo("concepto")
    f_bl = f_bl or pub_archivo("baseline")
    if hay_pub and f_rep is not None and subido_rep is None:
        st.caption(f"📢 Reporte publicado el {meta_p.get('publicado', '—')}"
                   + (" · puedes subir uno nuevo para verlo." if ES_ADMIN else "."))
    if not f_rep:
        st.info("Sube el reporte de Concepto TSS para empezar." if ES_ADMIN or not hay_pub
                else "Todavía no hay un reporte publicado.")
        st.stop()

    try:
        rep = ct_leer(f_rep.getvalue(), f_rep.name, ("Region", "Sitio", "Dia TSS", "Area sin concepto"))
    except Exception as e:  # noqa: BLE001
        st.error(str(e))
        st.stop()
    base = None
    if f_bl:
        try:
            base = ct_leer(f_bl.getvalue(), f_bl.name, ("Sitio", "Territorio Comercial"))
        except Exception as e:  # noqa: BLE001
            st.warning(f"No pude leer el baseline: {e}. Sigo sin territorio.")

    regiones = sorted(rep[ct_col(rep, "Region")].dropna().astype(str).str.strip().unique())
    hoy = datetime.datetime.now(ZoneInfo("America/Bogota")).date()
    o1, o2, o3 = st.columns([2, 2, 3])
    region = o1.selectbox("Región", ["Todas"] + regiones,
                          index=(regiones.index("Centro") + 1) if "Centro" in regiones else 0)
    corte = o2.date_input("Fecha de corte", hoy, help="Los días hábiles se cuentan hasta esta fecha.")
    solo_fuera = o3.checkbox("Solo fuera de ventana (más de 5 días hábiles)", value=False)
    c_tit = st.checkbox("Agregar la fecha de corte al título", value=False)
    c_tit_img = st.checkbox("Escribir también el título dentro de la imagen", value=False)
    por_img = st.slider("Filas por imagen", 10, 60, 30, help="Si hay más filas, se parte en varias "
                                                             "imágenes para que se lean bien en WhatsApp.")

    tabla, n_filas = ct_procesar(rep, base, region, corte, solo_fuera)
    if tabla.empty:
        st.warning("No hay sitios con esa selección.")
        st.stop()
    if base is None:
        tabla = tabla.drop(columns=["Territorio"])

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Sitios únicos", f"{tabla['Sitio'].nunique():,}")
    k2.metric("Registros (sitio + área)", f"{len(tabla):,}", help=f"Había {n_filas:,} filas en el reporte.")
    k3.metric("Fuera de ventana (>5 días)", f"{int((tabla['Días hábiles'] > 5).sum()):,}")
    k4.metric("Máx. días hábiles", f"{int(tabla['Días hábiles'].max())}")

    por_area = tabla["Área sin concepto"].value_counts()
    st.caption("Por área: " + " · ".join(f"**{a}** {int(n)}" for a, n in
                                          por_area.reindex([a for a in CT_ORDEN_AREA if a in por_area.index]
                                                           + [a for a in por_area.index if a not in CT_ORDEN_AREA]).items()))

    if ES_ADMIN and (subido_rep is not None or subido_bl is not None):
        ent = {}
        if subido_rep is not None:
            ent["concepto"] = (subido_rep.name, subido_rep.getvalue())
        if subido_bl is not None:
            ent["baseline"] = (subido_bl.name, subido_bl.getvalue())
        bloque_publicar(ent, {})
    elif ES_ADMIN and hay_pub:
        bloque_publicar({}, {})

    titulo = f"Concepto TSS · {region} · {corte:%d/%m/%Y}"
    areas = ([a for a in CT_ORDEN_AREA if a in por_area.index]
             + [a for a in por_area.index if a not in CT_ORDEN_AREA])

    st.subheader("Para pegar en WhatsApp")
    st.caption("Cada área tiene su título en texto y su imagen. En WhatsApp pega primero la imagen "
               "(Ctrl+V) y en el cuadro de comentario pega el título. Así el título sale en negrita.")
    for area in areas:
        sub = tabla[tabla["Área sin concepto"] == area].drop(columns=["Área sin concepto"])
        partes = [sub.iloc[i:i + por_img] for i in range(0, len(sub), por_img)]
        for n, parte in enumerate(partes, start=1):
            t_txt = f"*{area}*" + (f" · {corte:%d/%m/%Y}" if c_tit else "")
            if len(partes) > 1:
                t_txt += f" ({n}/{len(partes)})"
            st.code(t_txt, language=None)
            png = ct_imagen_tabla(parte, t_txt.replace("*", "") if c_tit_img else None)
            st.image(png, width=560)
            b1, b2 = st.columns([3, 2])
            with b1:
                ct_boton_copiar(png)
            b2.download_button("⬇️ Descargar PNG", png, file_name=f"concepto_tss_{_plano(area).replace(' ', '_')}_{n}.png",
                               mime="image/png", key=f"png_{_plano(area)}_{n}")

    with st.expander("📋 Ver como tabla (para copiar y pegar en Excel)"):
        st.markdown("<style>table.ctx{border-collapse:collapse;width:100%;font-size:14px}"
                    "table.ctx th{background:#0070c0;color:#fff;padding:6px 10px;text-align:center}"
                    "table.ctx td{padding:5px 10px;text-align:center;color:#111}"
                    "table.ctx td:first-child,table.ctx th:first-child{text-align:left}"
                    "table.ctx tr:nth-child(even) td{background:#ddebf7}"
                    "table.ctx tr:nth-child(odd) td{background:#fff}</style>", unsafe_allow_html=True)
        for area in areas:
            sub = tabla[tabla["Área sin concepto"] == area].drop(columns=["Área sin concepto"])
            st.markdown(f"**{area}** · {len(sub)} sitios")
            st.markdown(ct_tabla_html(sub), unsafe_allow_html=True)

    with st.expander("💬 Versión en texto para WhatsApp"):
        st.caption("Con el ícono de copiar de la esquina. WhatsApp muestra el área en negrita y la "
                   "tabla en letra monoespaciada.")
        for area in areas:
            st.code(ct_texto_whatsapp(tabla[tabla["Área sin concepto"] == area], area), language=None)

    st.subheader("Descargas")
    st.download_button("⬇️ Descargar Excel", ct_excel_bytes(tabla, titulo),
                       file_name=f"concepto_tss_{_plano(region).replace(' ', '_')}_{corte:%Y%m%d}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------------------
# Interfaz
# ---------------------------------------------------------------------------
PANTALLA_DASH = "📊 Dashboard Modernización"
PANTALLA_TSS = "📨 Concepto TSS"
with st.sidebar:
    pantalla = st.radio("Pantalla", [PANTALLA_DASH, PANTALLA_TSS], key="pantalla")
ES_ADMIN = es_administrador()
try:
    _AGENTE = bool(gh_cfg() and st.secrets.get("AGENTE_PC", False))
except Exception:  # noqa: BLE001
    _AGENTE = False
with st.sidebar:
    _clic_actualizar = False
    if _AGENTE:
        _clic_actualizar = st.button("🔄 Actualizar datos desde el computador",
                                     key="btn_actualizar", width="stretch", type="primary",
                                     help="Le pide al computador de quien administra la app que "
                                          "envíe los archivos más recientes. Tarda 1 o 2 minutos.")
        _ult = pub_meta().get("publicado")
        if _ult:
            st.caption(f"Última actualización: {_ult}")
        _org = pub_meta().get("origen", {})
        if _org:
            st.caption("Archivos de origen: " + " · ".join(
                f"{k} ({o.get('modificado', '—')})" for k, o in _org.items()))
if _clic_actualizar:
    solicitar_actualizacion()
if pantalla == PANTALLA_TSS:
    pantalla_concepto_tss()
    st.stop()

MODO_ARCHIVOS = "Lista de sitios + RF + baseline"
MODO_EXCEL = "Excel completo de control"
SIT_PC = "📁 Archivo del PC (Excel/CSV)"
SIT_PEGAR = "✍️ Pegar los nombres aquí"
FUENTE_RECALC = "Recalculado (corregido)"
FUENTE_EXCEL = "Resumen del Excel"

_pub = pub_meta()
_arch_pub = _pub.get("archivos", {})
_hay_pub = bool(_arch_pub)
_mostrar_carg = ES_ADMIN or not _hay_pub
with st.sidebar:
    st.header("📂 Datos")
    if _mostrar_carg:
        _modos = [MODO_ARCHIVOS, MODO_EXCEL]
        modo = st.radio("¿Qué vas a subir?", _modos, key="modo_carga",
                        index=1 if _pub.get("modo") == "excel" else 0)
    else:
        modo = MODO_EXCEL if _pub.get("modo") == "excel" else MODO_ARCHIVOS
    modo_arch = modo == MODO_ARCHIVOS
    f_sitios = f_rf = f_bl = archivo = None
    txt_sitios = ""
    proyecto, region_txt = _pub.get("proyecto", "MODERNIZACION_2026"), _pub.get("region", "R4")
    if _hay_pub:
        st.caption(f"📢 Datos publicados el {_pub.get('publicado', '—')}"
                   + (". Si subes archivos nuevos, se usan en vez de los publicados."
                      if ES_ADMIN else "."))
    fuente_sitios = SIT_PC
    if modo_arch:
        _sit_pub = "sitios" in _arch_pub
        fuente_sitios = st.radio(
            "1 · ¿De dónde salen los sitios?", [SIT_PC, SIT_PEGAR], key="fuente_sitios",
            index=0 if (_sit_pub or not _hay_pub) else 1,
            help="• Archivo del PC: el Excel/CSV con la lista de sitios (el publicado desde el "
                 "computador, o uno que subas si eres administrador).\n"
                 "• Pegar: escribe o pega los nombres, uno por línea. Solo vale para esta sesión.")
        if fuente_sitios == SIT_PEGAR:
            txt_sitios = st.text_area("Pega los nombres (uno por línea)", height=120,
                                      key="txt_sitios")
        elif _mostrar_carg:
            f_sitios = st.file_uploader("Lista de sitios (.csv, .txt, .xlsx)",
                                        type=["csv", "txt", "xlsx", "xlsm"], key="up_sitios")
            if _sit_pub and f_sitios is None:
                st.caption("Se usa la lista de sitios publicada desde el PC.")
        elif _sit_pub:
            st.caption("Se usa la lista de sitios publicada desde el PC.")
        else:
            st.warning("No hay una lista de sitios publicada. Pega los nombres.")
    if modo_arch and _mostrar_carg:
        f_rf = st.file_uploader("2 · Export de RF (.xls, .xlsx, .csv, .htm, .html)",
                                type=["xlsx", "xlsm", "xls", "csv", "htm", "html"], key="up_rf")
        f_bl = st.file_uploader("3 · Export de baseline (.xls, .xlsx, .csv, .htm, .html)",
                                type=["xlsx", "xlsm", "xls", "csv", "htm", "html"], key="up_bl")
        proyecto = st.text_input("Nombre del proyecto", proyecto, key="proy")
        region_txt = st.text_input("Región", region_txt, key="reg")
        if not _hay_pub:
            st.caption("Los archivos no se guardan en ningún lado; solo se procesan en esta "
                       "sesión, salvo que los publiques.")
    elif not modo_arch and _mostrar_carg:
        archivo = st.file_uploader("Sube el control (.xlsx)", type=["xlsx", "xlsm"])
        st.caption("Se lee la pestaña **Resumen** y se recalcula desde **Sitios**.")

with st.sidebar:
    st.header("⚙️ Criterio de RFIC-OK")
    criterio_rfic = st.radio(
        "RFIC-OK se cuenta cuando…", [RFIC_COMBINADO, RFIC_COLUMNA], key="criterio_rfic",
        help="• TSS ok + CW Lista + Orden Tx con fecha: el sitio cuenta cuando todos sus "
             "trabajos tienen fecha en esas tres columnas de RF.\n"
             "• Fecha en la columna RFIC: el criterio anterior.\n"
             "No aplica a la vista «Resumen del Excel», que muestra los valores tal cual.")

_t1, _t2 = st.columns([4, 1], vertical_alignment="center")
_t1.title("📡 Dashboard Resumen · Modernización")
presentacion = _t2.toggle("🎞️ Modo presentación", key="presentacion",
                          help="Oculta la barra lateral, las ayudas y los botones para "
                               "proyectar o capturar el tablero.")
if presentacion:
    st.markdown("<style>[data-testid='stSidebar'],[data-testid='stSidebarCollapsedControl'],"
                "[data-testid='collapsedControl'],[data-testid='stHeader'],header"
                "{display:none !important}.block-container{padding-top:1rem;"
                "max-width:100%;padding-left:2rem;padding-right:2rem}</style>", unsafe_allow_html=True)


def pista(texto):
    """Ayuda breve que no se muestra en modo presentación."""
    if not presentacion:
        st.caption(texto)



df_excel = sin_terr = det = calidad = error_recalc = None
diferencias = []

if modo_arch:
    _subidos = {"rf": f_rf, "baseline": f_bl}
    if fuente_sitios == SIT_PEGAR:
        sitios = limpiar_sitios(txt_sitios.splitlines())
    else:
        sitios = limpiar_sitios(list(leer_lista(f_sitios.getvalue(), f_sitios.name))) if f_sitios else []
        if not sitios and pub_archivo("sitios") is not None:
            _sp = pub_archivo("sitios")
            sitios = limpiar_sitios(list(leer_lista(_sp.getvalue(), _sp.name)))
    f_rf = f_rf or pub_archivo("rf")
    f_bl = f_bl or pub_archivo("baseline")
    if not (f_rf and f_bl and sitios):
        faltan = [n for n, ok in (("la lista de sitios" if fuente_sitios == SIT_PC
                                   else "los nombres de los sitios (pégalos)", bool(sitios)),
                                  ("el export de RF", bool(f_rf)),
                                  ("el export de baseline", bool(f_bl))) if not ok]
        st.info("Para empezar sube " + ", ".join(faltan) + " en la barra lateral.")
        st.stop()
    with st.sidebar:
        st.caption(f"✅ {len(sitios)} sitios en la lista"
                   + (" (pegados)" if fuente_sitios == SIT_PEGAR else " (archivo)"))
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
    if ES_ADMIN:
        _ent = {k: (v.name, v.getvalue()) for k, v in _subidos.items() if v is not None}
        if f_sitios is not None or (fuente_sitios == SIT_PEGAR and txt_sitios.strip()) or not _hay_pub:
            _ent["sitios"] = ("sitios.txt", "\n".join(sitios).encode("utf-8"))
        bloque_publicar(_ent, {"modo": "archivos", "proyecto": proyecto, "region": region_txt})
    total_col = proyecto.strip() or "Proyecto"
    info = {"region": region_txt.strip() or "—", "total_col": total_col, "ref_encabezado": None}
    terr_excel = [c for c in df_recalc.columns if c != "Total"]
    fuente = FUENTE_RECALC
    origen = "lista de sitios + RF + baseline"
else:
    _archivo_subido = archivo
    archivo = archivo or pub_archivo("control")
    if archivo is None:
        st.info("Sube el archivo de control en la barra lateral para ver el resumen.")
        st.stop()
    if ES_ADMIN:
        bloque_publicar({"control": (archivo.name, archivo.getvalue())} if _archivo_subido is not None
                        else {}, {"modo": "excel"})
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


@st.cache_data(show_spinner=False)
def fondo_oscuro(ancho: int, alto: int) -> str:
    """Fondo negro con resplandor rojo en las esquinas y un patrón sutil de hexágonos."""
    import math

    from PIL import Image, ImageDraw
    ancho, alto = int(ancho), int(alto)
    yy, xx = np.mgrid[0:alto, 0:ancho].astype("float32")
    m = float(max(ancho, alto))

    def brillo(cx, cy, radio, p=2.2):
        r = np.sqrt(((xx - cx * ancho) ** 2 + (yy - cy * alto) ** 2)) / (radio * m)
        return np.clip(1 - r, 0, 1) ** p

    base = np.zeros((alto, ancho, 3), dtype="float32") + np.array([10, 5, 7], dtype="float32")
    for cx, cy, rad, col in ((0.0, 0.0, 0.62, (120, 14, 18)), (1.0, 1.0, 0.50, (95, 12, 16)),
                             (0.0, 1.0, 0.40, (45, 8, 10)), (1.0, 0.0, 0.40, (55, 9, 12)),
                             (0.5, 0.5, 0.75, (28, 7, 9))):
        base += brillo(cx, cy, rad)[..., None] * np.array(col, dtype="float32")
    img = Image.fromarray(np.clip(base, 0, 255).astype("uint8"), "RGB").convert("RGBA")

    # Hexágonos: más visibles hacia la derecha, abajo y a la izquierda
    W = np.maximum(np.maximum(brillo(1.0, 1.0, 0.5, 1.2), brillo(0.0, 0.62, 0.30, 1.4) * 0.7),
                   brillo(1.0, 0.0, 0.28, 1.4) * 0.6)
    capa = Image.new("RGBA", img.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(capa)
    lado = max(34, int(m / 26))
    dx, dy = lado * math.sqrt(3), lado * 1.5
    fila = 0
    y = -lado
    while y < alto + lado:
        x = -dx + (dx / 2 if fila % 2 else 0)
        while x < ancho + dx:
            peso = float(W[min(max(int(y), 0), alto - 1), min(max(int(x), 0), ancho - 1)])
            if peso > 0.12:
                pts = [(x + lado * math.cos(math.radians(60 * k + 30)),
                        y + lado * math.sin(math.radians(60 * k + 30))) for k in range(6)]
                d.polygon(pts, outline=(170, 30, 34, int(210 * peso)))
            x += dx
        y += dy
        fila += 1
    img = Image.alpha_composite(img, capa).convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=86)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()


# Logo TEAM R4 + Claro (PNG con fondo transparente, incluido en el código)
LOGO_TEAM_R4_B64 = (
    "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"
)


def preparar_logo(datos: bytes):
    """Deja el logo listo para ponerlo sobre el fondo oscuro: quita el fondo liso o oscuro,
    recorta los márgenes vacíos y lo devuelve (uri, proporción ancho/alto)."""
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(datos)).convert("RGBA")
        arr = np.array(im).astype("float32")
        alfa = arr[..., 3]
        if (alfa < 200).mean() < 0.02:  # sin transparencia real
            maxc = arr[..., :3].max(axis=2)
            if np.median(maxc) < 110:
                # Fondo oscuro (p. ej. un recorte de la diapositiva): lo oscuro se vuelve transparente
                t = np.clip((maxc - 95) / 70, 0, 1)
                alfa = alfa * (t * t * (3 - 2 * t))
            else:
                # Fondo claro liso: se quita el color de las esquinas
                h, w = alfa.shape
                esq = [arr[0, 0, :3], arr[0, w - 1, :3], arr[h - 1, 0, :3], arr[h - 1, w - 1, :3]]
                bg = np.mean(esq, axis=0)
                dist = np.abs(arr[..., :3] - bg).sum(axis=2)
                alfa = alfa * np.clip((dist - 60) / 60, 0, 1)
            arr[..., 3] = alfa
        im = Image.fromarray(arr.astype("uint8"), "RGBA")
        caja = im.getchannel("A").point(lambda v: 255 if v > 40 else 0).getbbox()
        if caja:
            x0, y0, x1, y1 = caja
            px_, py_ = int((x1 - x0) * .06) + 2, int((y1 - y0) * .10) + 2
            im = im.crop((max(x0 - px_, 0), max(y0 - py_, 0),
                          min(x1 + px_, im.size[0]), min(y1 + py_, im.size[1])))
        if max(im.size) > 900:
            im.thumbnail((900, 900))
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        return ("data:image/png;base64," + base64.b64encode(buf.getvalue()).decode(),
                im.size[0] / max(im.size[1], 1))
    except Exception:  # noqa: BLE001
        mime = "image/png" if datos[:4] == b"\x89PNG" else "image/jpeg"
        return f"data:{mime};base64," + base64.b64encode(datos).decode(), 2.5


def aclarar(hex_: str, a: float) -> str:
    h = hex_.lstrip("#")
    return "#%02x%02x%02x" % tuple(int(round(int(h[i:i + 2], 16) + (255 - int(h[i:i + 2], 16)) * a))
                                   for i in (0, 2, 4))


def degradado(hex_: str, n: int) -> list:
    """n tonos del mismo color, de más claro (arriba) a más oscuro (abajo)."""
    h = hex_.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    if n <= 1:
        return [hex_]
    salida = []
    for i in range(n):
        a = 0.55 - 0.95 * i / (n - 1)          # >0 aclara hacia blanco, <0 oscurece hacia negro
        mezcla = (lambda c: c + (255 - c) * a) if a >= 0 else (lambda c: c * (1 + a))
        salida.append("#%02x%02x%02x" % tuple(int(round(mezcla(c))) for c in (r, g, b)))
    return salida


def arbol_meta(titulo_txt, etiqueta_meta, n_meta, n_ok, n_nok, hijos_ok, hijos_nok,
               regiones=(), region_sel="Todas", escala=1.0, logo_uri=None):
    """Árbol Meta → RFIC OK / RFIC NOK → cajas, con tarjetas clicables. Devuelve (fig, nodos)."""
    R = len(hijos_ok) + len(hijos_nok) + 1
    ys_ok = [R - i for i in range(len(hijos_ok))]
    ys_nok = [R - len(hijos_ok) - 1 - j for j in range(len(hijos_nok))]
    y_ok, y_nok = sum(ys_ok) / len(ys_ok), sum(ys_nok) / len(ys_nok)
    y_meta = (y_ok + y_nok) / 2
    con_chips = len(regiones) > 1
    y_top = R + (1.75 if con_chips else 0.85)
    extra_t = 0
    alto = int(max(600, 80 * (R + 1) + 120 + (60 if con_chips else 0)) * escala) + int(extra_t * escala)
    px_y = (alto - (90 + extra_t) * escala) / (y_top - 0.15)               # píxeles por unidad vertical (aprox.)
    rx, ry = 10 * escala * 100 / 1400, 10 * escala / px_y   # esquinas de ~10 px
    h = 0.84                              # alto de cada tarjeta (unidades y)

    # (clave, etiqueta, n, x0, x1, y, tamaño del número)
    META = (2, 21)
    PAD = (31, 50)
    HIJO = (67, 98)
    nodos = [("Meta", etiqueta_meta, n_meta, *META, y_meta, 42 * escala),
             ("RFIC OK", "RFIC OK", n_ok, *PAD, y_ok, 36 * escala),
             ("RFIC NOK", "RFIC NOK", n_nok, *PAD, y_nok, 36 * escala)]
    nodos += [(nm, nm, n, *HIJO, y, 30 * escala) for (nm, n), y in zip(hijos_ok, ys_ok)]
    nodos += [(nm, nm, n, *HIJO, y, 30 * escala) for (nm, n), y in zip(hijos_nok, ys_nok)]

    # Las cajas de cada rama van en degradado del color de su padre
    COL = dict(COLOR_CAJA)
    TXT = {}
    for padre, hijos in (("RFIC OK", hijos_ok), ("RFIC NOK", hijos_nok)):
        for (nm, _n), tono in zip(hijos, degradado(COLOR_CAJA[padre], len(hijos))):
            COL[nm] = tono
            TXT[nm] = aclarar(tono, 0.30)   # el texto de las cajas oscuras se mantiene legible

    formas, anot, flechas = [], [], []

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
    for clave, yp, ys, hijos in (("RFIC OK", y_ok, ys_ok, hijos_ok), ("RFIC NOK", y_nok, ys_nok, hijos_nok)):
        for yh, (nm, _n) in zip(ys, hijos):
            codo(PAD[1], yp, 58.5, HIJO[0] - 0.4, yh, COL[nm])

    # Tarjetas
    xs_clic, ys_clic, claves_clic, textos = [], [], [], []
    for clave, etq, n, x0, x1, y, tam in nodos:
        col = COL.get(clave, GRIS)
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
        h_px = h * px_y
        anot.append(dict(x=x0 + 1.4, y=y, yshift=0.14 * h_px, text=f"<b>{n:,}</b>", showarrow=False,
                         font=dict(size=min(tam, 0.36 * h_px), color="#ffffff"),
                         xanchor="left", yanchor="middle"))
        anot.append(dict(x=x0 + 1.4, y=y, yshift=-0.2 * h_px, text=f"<b>{etq}</b>", showarrow=False,
                         font=dict(size=min(15 * escala, 0.19 * h_px), color=TXT.get(clave, col)),
                         xanchor="left", yanchor="middle"))
        if clave != "Meta":
            anot.append(dict(x=x1 - 1.2, y=y, yshift=0.14 * h_px, text=f"<b>{pc:.0f}%</b>",
                             showarrow=False, font=dict(size=min(21 * escala, 0.3 * h_px),
                                                        color="#ffffff"),
                             xanchor="right", yanchor="middle"))
        for fx in (x0 + 0.2 * (x1 - x0), (x0 + x1) / 2, x0 + 0.8 * (x1 - x0)):
            xs_clic.append(fx); ys_clic.append(y); claves_clic.append(clave)
            textos.append(f"<b>{etq}</b><br>{n:,} sitios · {pc:.0f}% de la meta")

    # Nombre de la región que se está viendo, sobre la caja Meta
    if con_chips:
        nombre_reg = "Todas las regiones" if region_sel == "Todas" else region_sel
        hb = 1.5                                   # alto de la etiqueta (unidades y)
        yb0 = y_meta + h / 2 + 0.55                # base, con un pequeño espacio sobre Meta
        ya_, yb_ = yb0, yb0 + hb
        x0_, x1_ = META[0], META[1]
        formas.append(dict(
            type="path", layer="below",
            path=(f"M {x0_ + rx},{ya_} L {x1_ - rx},{ya_} Q {x1_},{ya_} {x1_},{ya_ + ry} "
                  f"L {x1_},{yb_ - ry} Q {x1_},{yb_} {x1_ - rx},{yb_} L {x0_ + rx},{yb_} "
                  f"Q {x0_},{yb_} {x0_},{yb_ - ry} L {x0_},{ya_ + ry} Q {x0_},{ya_} {x0_ + rx},{ya_} Z"),
            fillcolor="#e2231a", line=dict(color="#ffffff", width=2)))
        yc_ = (ya_ + yb_) / 2
        hb_px = hb * px_y
        anot.append(dict(x=(x0_ + x1_) / 2, y=yc_, yshift=0.24 * hb_px, text="<b>REGIÓN</b>",
                         showarrow=False, font=dict(size=min(13 * escala, 0.2 * hb_px), color="#ffd2cf"),
                         xanchor="center", yanchor="middle"))
        tam_reg = min(24, 190 / (0.58 * max(len(nombre_reg), 1))) * escala
        anot.append(dict(x=(x0_ + x1_) / 2, y=yc_, yshift=-0.14 * hb_px, text=f"<b>{nombre_reg}</b>",
                         showarrow=False, font=dict(size=min(tam_reg, 0.34 * hb_px), color="#ffffff"),
                         xanchor="center", yanchor="middle"))

    # Selector de región dentro del gráfico (casillas clicables)
    if con_chips:
        regs = ["Todas"] + list(regiones)
        ancho, sep, hch = 13.5, 1.2, 0.5
        x_ini = 50 - (len(regs) * ancho + (len(regs) - 1) * sep) / 2 + 4
        yc = R + 1.2
        ya, yb = yc - hch / 2, yc + hch / 2
        rxc, ryc = (hch * px_y / 2) * 100 / 1400, hch / 2
        anot.append(dict(x=x_ini - 1.5, y=yc, text="<b>Región</b>", showarrow=False,
                         font=dict(size=15 * escala, color="#c9b8ba"), xanchor="right", yanchor="middle"))
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
                             font=dict(size=16 * escala, color="#ffffff" if activo else "#d9c9cb"),
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

    texto_titulo = f"<b>{titulo_txt}</b>"
    anot.append(dict(xref="paper", yref="paper", x=0.5, y=1.0, text=texto_titulo,
                     showarrow=False, font=dict(size=min(30 * escala, 36), color="#ffffff"), bgcolor="#e2231a",
                     bordercolor="#ffffff", borderwidth=2, borderpad=10, yanchor="bottom"))
    fig.update_layout(
        shapes=formas, annotations=anot, height=alto, clickmode="event+select", dragmode=False,
        paper_bgcolor="#0a0507", plot_bgcolor="rgba(0,0,0,0)", margin=dict(l=10, r=10, t=int((84 + extra_t) * escala), b=10),
        showlegend=False, hoverlabel=dict(bgcolor="#2a1a1d", font_size=14, font_color="#ffffff"),
        xaxis=dict(visible=False, range=[0, 100], fixedrange=True),
        yaxis=dict(visible=False, range=[0.15, y_top], fixedrange=True))
    if logo_uri:
        # El logo va en el hueco entre las cajas RFIC OK y RFIC NOK (columna del centro)
        uri, aspecto = logo_uri
        y_sup, y_inf = y_ok - h / 2, y_nok + h / 2
        fig.add_layout_image(dict(
            source=uri, xref="x", yref="y", x=43.5, y=(y_sup + y_inf) / 2 + 0.1 * (y_sup - y_inf),
            xanchor="center", yanchor="middle", sizex=27, sizey=0.62 * (y_sup - y_inf),
            sizing="contain", layer="above", opacity=1))
    fig.add_layout_image(dict(
        source=fondo_oscuro(1400, alto), xref="paper", yref="paper", x=0, y=1, sizex=1, sizey=1,
        xanchor="left", yanchor="top", sizing="stretch", layer="below"))
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
            ct1, ct2, ct3 = st.columns([3, 2, 2])
            ct1.text_input("Título", "MODERNIZACIÓN R4 - 4Q", key="pres_titulo")
            ct2.text_input("Etiqueta de la meta", "Meta 4Q", key="pres_meta")
            ct3.select_slider("Tamaño del gráfico", ["Compacto", "Normal", "Grande"],
                              value="Normal", key="pres_tam2")
            logo_up = st.file_uploader("Cambiar el logo (opcional)", type=["png", "jpg", "jpeg"],
                                       key="pres_logo",
                                       help="La app ya trae el logo TEAM R4 + Claro. Sube otro "
                                            "archivo solo si quieres reemplazarlo en esta sesión.")
            if logo_up is not None:
                st.session_state["_pres_logo_v"] = logo_up.getvalue()
            else:
                st.session_state.pop("_pres_logo_v", None)
            st.session_state["_pres_tam_v"] = st.session_state.get("pres_tam2", "Normal")
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
        escala_p = {"Compacto": 0.85, "Normal": 1.0, "Grande": 1.2}.get(
            st.session_state.get("_pres_tam_v", "Normal"), 1.0)
        logo_bytes = st.session_state.get("_pres_logo_v")
        if logo_bytes is None:
            logo_bytes = base64.b64decode(LOGO_TEAM_R4_B64)
        logo_uri = None
        if logo_bytes:
            logo_uri = preparar_logo(logo_bytes)
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
        titulo_arbol = titulo_p

        if n_meta == 0:
            st.info("No hay sitios para esta selección.")
            if st.button("Ver todas las regiones", key="pres_reset_region"):
                st.session_state["_pres_region"] = "Todas"
                st.rerun()
        else:
            fig_t, mapa_nodos = arbol_meta(titulo_arbol, meta_p, n_meta, n_ok, n_meta - n_ok,
                                           hijos_ok, hijos_nok, list(sel_terr), region_p,
                                           escala_p, logo_uri)
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

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


TIPO_CW_OK = ("nokia", "conjunta")          # «Tipo CW» que da por lista la obra civil sin fecha
TIPO_CW_NA = {"n/a", "na", "n.a", "n.a.", "n a"}


def tipo_cw_ok(v) -> bool:
    """Tipo CW Nokia, Conjunta o N/A: la obra civil no depende de una fecha de CW Lista."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    t = _plano(v)
    return any(x in t for x in TIPO_CW_OK) or t in TIPO_CW_NA


def cw_nokia(r: pd.DataFrame, col, mrf: dict | None = None) -> pd.Series:
    """Trabajos que cuentan como CW Lista OK aunque no tengan fecha:
    • CW Lista dice «Nokia» (en csv/html se guardó aparte al leer), o
    • «Tipo CW» es Nokia, Conjunta o N/A."""
    if "__cw_nokia" in r.columns:
        ok = r["__cw_nokia"].fillna(False).astype(bool)
    else:
        ok = r[col].map(es_nokia)
    c_tipo = (mrf or {}).get("tipo cw")
    if c_tipo is not None and c_tipo in r.columns:
        ok = ok | r[c_tipo].map(tipo_cw_ok)
    return ok


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
            f["v"] = f["v"] | cw_nokia(r, col, mrf).values
        ag = f.groupby("_k").agg(count=("v", "sum"), max=("f", "max"))
        ok = (ag["count"] == n_jobs.reindex(ag.index)) & (ag["count"] > 0)
        det["ok_" + clave_h] = [bool(ok.get(k, False)) for k in claves]
        det[HITO_A_DETALLE[clave_h]] = [ag["max"].get(k, pd.NaT) if ok.get(k, False) else pd.NaT
                                        for k in claves]
    det["ok_TOTAL SITIOS"] = True

    # Por qué una obra civil cuenta como lista sin fecha (CW Lista «Nokia» o Tipo CW Nokia/Conjunta/N/A)
    det["CW sin fecha por"] = ""
    c_cw = next((mrf[c] for c in HITO_RF["OBRA CIVIL OK"] if c in mrf), None)
    if c_cw is not None and len(r):
        sin_fecha = r[c_cw].map(a_fecha).isna()
        txt_nok = (r["__cw_nokia"].fillna(False).astype(bool) if "__cw_nokia" in r.columns
                   else r[c_cw].map(es_nokia))
        c_tipo = mrf.get("tipo cw")
        tipo_ok = r[c_tipo].map(tipo_cw_ok) if c_tipo in r.columns else pd.Series(False, index=r.index)
        motivo = pd.Series("", index=r.index)
        motivo[sin_fecha & tipo_ok] = "Tipo CW " + r.loc[sin_fecha & tipo_ok, c_tipo].astype(str).str.strip()
        motivo[sin_fecha & txt_nok] = "CW Lista Nokia"
        por_sitio = motivo.groupby(r["_k"]).agg(lambda x: ", ".join(sorted({v for v in x if v})))
        det["CW sin fecha por"] = [por_sitio.get(k, "") if det_ok else ""
                                   for k, det_ok in zip(claves, det["ok_OBRA CIVIL OK"])]

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
            v3[1] = v3[1] | cw_nokia(r, cols3[1], mrf)  # CW Lista sin fecha: Nokia / Conjunta / N/A cuenta como OK
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
    columnas += [c for c in ("CW sin fecha por", "Carga", "Autorización", "Scope")
                 if c in d.columns and (d[c].astype(str).str.strip() != "").any()]
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
    f_spi = None
    if _mostrar_carg:
        f_spi = st.file_uploader("SPI semanal · cuadro de la presentación (.xlsx, opcional)",
                                 type=["xlsx", "xlsm"], key="up_spi",
                                 help="Excel del SPI (hoja «MOD 2026»). En la pestaña Presentación se "
                                      "muestra solo la regional Centro y el mes actual.")

_spi_subido = f_spi
f_spi = f_spi or pub_archivo("spi")
_ent_spi = {"spi": (_spi_subido.name, _spi_subido.getvalue())} if _spi_subido is not None else {}

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
        _ent.update(_ent_spi)
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
        _ent_x = {"control": (archivo.name, archivo.getvalue())} if _archivo_subido is not None else {}
        _ent_x.update(_ent_spi)
        bloque_publicar(_ent_x, {"modo": "excel"})
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


# ---------------------------------------------------------------------------
# Cuadro SPI (Excel del SPI semanal) para la pestaña Presentación
# ---------------------------------------------------------------------------
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
            "septiembre", "octubre", "noviembre", "diciembre"]


@st.cache_data(show_spinner=False)
def spi_leer(contenido: bytes, mes: int, region: str = "centro", hoja_pref: str = "MOD 2026"):
    """Lee el bloque del mes de la hoja del SPI (solo la regional pedida).
    Devuelve {'mes', 'regional', 'proyecto', 'semanas', 'filas': {'Semanal': [...], 'Acumulado': [...]}}."""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(contenido), data_only=True, read_only=True)
    hojas = ([h for h in wb.sheetnames if _plano(h) == _plano(hoja_pref)]
             + [h for h in wb.sheetnames if _plano(h) != _plano(hoja_pref)])
    for h in hojas:
        filas = [list(r) for r in wb[h].iter_rows(min_row=1, max_row=80, values_only=True)]
        ancho = max((len(r) for r in filas), default=0)
        filas = [r + [None] * (ancho - len(r)) for r in filas]
        enc = None
        for i, r in enumerate(filas):
            pl = [_plano(v) if v is not None else "" for v in r]
            if "regional" in pl and "proyecto" in pl and "tipo" in pl and i >= 2:
                enc = (i, pl.index("regional"))
                break
        if enc is None:
            continue
        i_h, c0 = enc
        r_mes, r_pas, r_sem = filas[i_h - 2], filas[i_h - 1], filas[i_h]
        if "plan" not in [_plano(v) for v in r_pas if v is not None]:
            continue
        # bloques de mes: cada celda con texto en la fila de meses abre un bloque
        inicios = [(c, str(r_mes[c]).strip()) for c in range(c0 + 3, ancho)
                   if r_mes[c] is not None and str(r_mes[c]).strip()]
        if not inicios:
            continue
        bloques = []
        for k, (c, nombre) in enumerate(inicios):
            fin = (inicios[k + 1][0] if k + 1 < len(inicios) else ancho) - 1
            bloques.append((nombre, c, fin))
        elegido = next((b for b in bloques if _plano(b[0]) == MESES_ES[mes - 1]), None) or bloques[-1]
        nombre_mes, ini, fin = elegido
        semanas = []
        for c in range(ini, fin + 1):
            if _plano(r_pas[c]) == "plan":
                semanas.append((str(r_sem[c]).strip(), c))
        reg = {"Semanal": None, "Acumulado": None}
        regional = proyecto = ""
        for r in filas[i_h + 1:]:
            if r[c0] is None or _plano(region) not in _plano(r[c0]):
                continue
            tipo = str(r[c0 + 2]).strip().capitalize()
            if tipo not in reg or reg[tipo] is not None:
                continue
            regional, proyecto = str(r[c0]).strip(), str(r[c0 + 1] or "").strip()
            reg[tipo] = [(r[c], r[c + 1], r[c + 2]) for _, c in semanas]
        if not semanas or all(v is None for v in reg.values()):
            raise ValueError(f"No encontré la regional «{region}» en la hoja «{h}».")
        return {"mes": nombre_mes, "regional": regional, "proyecto": proyecto,
                "semanas": [w for w, _ in semanas], "filas": {k: v for k, v in reg.items() if v is not None}}
    raise ValueError("No encontré la tabla del SPI (columnas Regional, Proyecto y Tipo con "
                     "Plan, Actual y SPI).")


@st.cache_data(show_spinner=False)
def spi_imagen(datos: dict, semana_actual: str = ""):
    """Panel del SPI con el estilo oscuro del gráfico: cada semana tiene sus columnas
    Plan, Actual y SPI (el SPI en una etiqueta de color). Los anchos se miden con la tipografía
    disponible, para que nada se encime. Devuelve (uri PNG, proporción ancho/alto)."""
    from PIL import Image, ImageDraw
    S = 2                                           # sobremuestreo para que se vea nítido
    f_tit, f_sub = _ct_fuente(21 * S, True), _ct_fuente(14 * S)
    f_cab, f_men = _ct_fuente(17 * S, True), _ct_fuente(13 * S, True)
    f_num, f_pl = _ct_fuente(21 * S, True), _ct_fuente(19 * S)
    f_pil, f_lab = _ct_fuente(15 * S, True), _ct_fuente(16 * S, True)
    medidor = ImageDraw.Draw(Image.new("RGBA", (4, 4)))

    def ancho(txt, fuente):
        return medidor.textlength(str(txt), font=fuente) / S

    def fmt(v):
        return "" if v is None or v == "" else (
            f"{int(round(v)):,}" if isinstance(v, (int, float)) else str(v))

    n = len(datos["semanas"])
    filas = list(datos["filas"].items())
    nombre = datos["regional"].replace("_", " ")
    # anchos de columna según el contenido
    todos = [v for _, vals in filas for fila in vals for v in fila[:2]]
    w_pl = max([ancho(fmt(v), f_pl) for v in todos] + [ancho("Plan", f_men)]) + 26
    w_ac = max([ancho(fmt(v), f_num) for v in todos] + [ancho("Actual", f_men)]) + 26
    w_sp = max(ancho("100%", f_pil) + 34, ancho("SPI", f_men) + 30)
    w_gr = w_pl + w_ac + w_sp                       # ancho de cada semana
    w_lab = max(ancho(f"SPI {datos['mes']}", f_tit), ancho(nombre, f_sub),
                *(ancho(t, f_lab) for t, _ in filas)) + 28
    pad, h_sem, h_sub, h_fila = 6, 34, 24, 46
    W = int((w_lab + w_gr * n + 2 * pad) * S)
    H = int((h_sem + h_sub + h_fila * len(filas) + 2 * pad) * S)
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    BLANCO, MUTE = (255, 255, 255, 255), (205, 188, 190, 255)
    VERDE, AMBAR, ROJO = (27, 175, 122), (242, 176, 30), (255, 90, 77)
    d.rounded_rectangle((0, 0, W - 1, H - 1), radius=14 * S, fill=(255, 255, 255, 16),
                        outline=(255, 255, 255, 70), width=S)
    x0, y0 = pad * S, pad * S
    y_datos = y0 + (h_sem + h_sub) * S
    # título del panel (columna izquierda, sobre las cabeceras)
    d.text((x0 + 14 * S, y0 + (h_sem + h_sub) * S * 0.36), f"SPI {datos['mes']}", font=f_tit,
           fill=BLANCO, anchor="lm")
    d.text((x0 + 14 * S, y0 + (h_sem + h_sub) * S * 0.70), nombre, font=f_sub, fill=MUTE, anchor="lm")
    d.line((x0 + 10 * S, y_datos, W - x0 - 10 * S, y_datos), fill=(255, 255, 255, 60), width=S)

    for k, sem in enumerate(datos["semanas"]):
        gx = x0 + (w_lab + w_gr * k) * S                          # inicio del grupo de la semana
        c_pl = gx + w_pl * S // 2
        c_ac = gx + (w_pl + w_ac // 2) * S
        c_sp = gx + int((w_pl + w_ac + w_sp / 2) * S)
        # separador vertical de la semana
        d.line((gx, y0 + 8 * S, gx, H - y0 - 8 * S), fill=(255, 255, 255, 55), width=S)
        # cabecera: semana y subtítulos
        cx, cy = gx + w_gr * S // 2, y0 + h_sem * S // 2
        if sem == semana_actual:
            d.rounded_rectangle((cx - 36 * S, cy - 14 * S, cx + 36 * S, cy + 14 * S), radius=14 * S,
                                fill=(226, 35, 26, 255), outline=(255, 255, 255, 220), width=S)
            d.text((cx, cy), sem, font=f_cab, fill=BLANCO, anchor="mm")
        else:
            d.text((cx, cy), sem, font=f_cab, fill=MUTE, anchor="mm")
        ys = y0 + (h_sem + h_sub // 2) * S
        d.text((c_pl, ys), "Plan", font=f_men, fill=MUTE, anchor="mm")
        d.text((c_ac, ys), "Actual", font=f_men, fill=(255, 170, 160, 255), anchor="mm")
        d.text((c_sp, ys), "SPI", font=f_men, fill=MUTE, anchor="mm")
        # datos
        for fi, (tipo, vals) in enumerate(filas):
            pl, ac, spi = vals[k]
            cy = y_datos + (fi * h_fila + h_fila // 2) * S
            d.text((c_pl, cy), fmt(pl), font=f_pl, fill=MUTE, anchor="mm")
            d.text((c_ac, cy), fmt(ac), font=f_num, fill=BLANCO, anchor="mm")
            if isinstance(spi, (int, float)):
                col = VERDE if spi >= 1 else AMBAR if spi >= 0.9 else ROJO
                pw, ph = int((w_sp - 16) * S), 26 * S
                d.rounded_rectangle((c_sp - pw // 2, cy - ph // 2, c_sp + pw // 2, cy + ph // 2),
                                    radius=ph // 2, fill=col + (70,), outline=col + (255,), width=S)
                claro = tuple(int(c + (255 - c) * 0.55) for c in col)
                d.text((c_sp, cy), f"{spi * 100:.0f}%", font=f_pil, fill=claro + (255,), anchor="mm")
    # etiquetas de fila y separadores horizontales
    for fi, (tipo, _) in enumerate(filas):
        cy = y_datos + (fi * h_fila + h_fila // 2) * S
        d.text((x0 + 14 * S, cy), tipo, font=f_lab, fill=BLANCO, anchor="lm")
        if fi > 0:
            yl = y_datos + fi * h_fila * S
            d.line((x0 + 10 * S, yl, W - x0 - 10 * S, yl), fill=(255, 255, 255, 38), width=S)
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return ("data:image/png;base64," + base64.b64encode(buf.getvalue()).decode(), W / H)


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
               regiones=(), region_sel="Todas", escala=1.0, logo_uri=None, spi_uri=None):
    """Árbol Meta → RFIC OK / RFIC NOK → cajas, con tarjetas clicables. Devuelve (fig, nodos)."""
    R = len(hijos_ok) + len(hijos_nok) + 1
    ys_ok = [R - i for i in range(len(hijos_ok))]
    ys_nok = [R - len(hijos_ok) - 1 - j for j in range(len(hijos_nok))]
    y_ok, y_nok = sum(ys_ok) / len(ys_ok), sum(ys_nok) / len(ys_nok)
    y_meta = (y_ok + y_nok) / 2
    con_chips = len(regiones) > 1
    y_top = R + (1.75 if con_chips else 0.85)
    extra_t = 66 if spi_uri else 0         # espacio extra arriba para el cuadro SPI
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

    if not spi_uri:
        anot.append(dict(xref="paper", yref="paper", x=0.5, y=1.0, text=f"<b>{titulo_txt}</b>",
                         showarrow=False, font=dict(size=min(30 * escala, 36), color="#ffffff"),
                         bgcolor="#e2231a", bordercolor="#ffffff", borderwidth=2, borderpad=10,
                         yanchor="bottom"))
    fig.update_layout(
        shapes=formas, annotations=anot, height=alto, clickmode="event+select", dragmode=False,
        paper_bgcolor="#0a0507", plot_bgcolor="rgba(0,0,0,0)", margin=dict(l=10, r=10, t=int((84 + extra_t) * escala), b=10),
        showlegend=False, hoverlabel=dict(bgcolor="#2a1a1d", font_size=14, font_color="#ffffff"),
        xaxis=dict(visible=False, range=[0, 100], fixedrange=True),
        yaxis=dict(visible=False, range=[0.15, y_top], fixedrange=True))
    if logo_uri:
        # El logo va en la esquina inferior izquierda, debajo de la caja Meta
        uri, aspecto = logo_uri
        y_base = 0.45
        alto_logo = max(min(2.4, (y_meta - h / 2 - 0.25) - y_base), 0.6)
        fig.add_layout_image(dict(
            source=uri, xref="x", yref="y", x=1.5, y=y_base, xanchor="left", yanchor="bottom",
            sizex=23.5, sizey=alto_logo, sizing="contain", layer="above", opacity=1))
    if spi_uri:
        uri_s, asp_s = spi_uri
        mt_px = (84 + extra_t) * escala
        plot_h = alto - mt_px - 10
        caja_px = mt_px - 22
        fig.add_layout_image(dict(
            source=uri_s, xref="paper", yref="paper", x=0.5, y=1 + 8 / plot_h,
            xanchor="center", yanchor="bottom", sizex=0.82, sizey=caja_px / plot_h,
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
        spi_uri = None
        if f_spi is not None:
            try:
                _mes_hoy = datetime.datetime.now(ZoneInfo("America/Bogota")).month
                _spi_d = spi_leer(f_spi.getvalue(), _mes_hoy, "centro")
                spi_uri = spi_imagen(_spi_d, f"W{datetime.datetime.now(ZoneInfo('America/Bogota')).isocalendar()[1]}")
                if not presentacion:
                    st.caption(f"📈 Cuadro SPI: {_spi_d['regional']} · {_spi_d['mes']} · "
                               f"archivo {f_spi.name}. Reemplaza al título del gráfico.")
            except Exception as e:  # noqa: BLE001
                _msg = ("el archivo no es un Excel .xlsx válido" if "zip" in str(e).lower() else str(e))
                st.warning(f"No pude leer el SPI: {_msg}. Se muestra el título en su lugar.")
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
                                           escala_p, logo_uri, spi_uri)
            # La clave del gráfico cambia después de cada clic atendido: así la casilla clicada
            # no queda "seleccionada" y volver a pulsarla (o pulsar la misma tras cerrar la
            # ventana) cuenta como un clic nuevo y no como "quitar la selección".
            n_clk = st.session_state.get("_arbol_n", 0)
            k_arbol = f"clk_arbol_{n_clk}"
            pto = leer_click(grafico_clic(fig_t, k_arbol), k_arbol)
            if pto:
                clave_n = mapa_nodos.get((round(float(pto.get("x", -1)), 3),
                                          round(float(pto.get("y", -1)), 3)))
                if clave_n is None and mapa_nodos:
                    # el punto llegó con coordenadas ligeramente distintas: se toma la casilla más cercana
                    try:
                        px_, py_ = float(pto.get("x", -1)), float(pto.get("y", -1))
                        clave_n = min(mapa_nodos.items(),
                                      key=lambda kv: (kv[0][0] - px_) ** 2 + (kv[0][1] - py_) ** 2)[1]
                    except (TypeError, ValueError):
                        clave_n = None
                if clave_n and clave_n.startswith("reg::"):
                    st.session_state["_pres_region"] = clave_n[5:]
                    st.session_state["_arbol_n"] = n_clk + 1
                    st.rerun()
                elif clave_n:
                    st.session_state["_abrir_t4"] = (PREFIJO_PRES + clave_n,
                                                     region_p if region_p != "Todas" else "Todos")
                    st.session_state["_arbol_n"] = n_clk + 1
                    st.rerun()
            _ab = st.session_state.pop("_abrir_t4", None)
            if _ab:
                ir_a_detalle("t4", _ab[0], _ab[1])
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

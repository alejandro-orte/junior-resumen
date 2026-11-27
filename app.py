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
    c_tit = st.checkbox("Agregar la fecha de corte al título de cada imagen", value=False)
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
    st.caption("Cada área queda en su propia imagen, con el área como título. Pulsa **Copiar imagen** "
               "y pégala en WhatsApp con Ctrl+V (o clic derecho sobre la imagen → Copiar imagen).")
    for area in areas:
        sub = tabla[tabla["Área sin concepto"] == area].drop(columns=["Área sin concepto"])
        partes = [sub.iloc[i:i + por_img] for i in range(0, len(sub), por_img)]
        for n, parte in enumerate(partes, start=1):
            t_img = area + (f" · {corte:%d/%m/%Y}" if c_tit else "")
            if len(partes) > 1:
                t_img += f" ({n}/{len(partes)})"
            png = ct_imagen_tabla(parte, t_img)
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
    alto = int(max(600, 80 * (R + 1) + 120 + (60 if con_chips else 0) + (150 if logo_uri else 0)) * escala)
    px_y = (alto - 90 * escala) / (y_top - (-1.75 if logo_uri else 0.15))               # píxeles por unidad vertical (aprox.)
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
        h_px = h * px_y
        anot.append(dict(x=x0 + 1.4, y=y, yshift=0.14 * h_px, text=f"<b>{n:,}</b>", showarrow=False,
                         font=dict(size=min(tam, 0.36 * h_px), color="#ffffff"),
                         xanchor="left", yanchor="middle"))
        anot.append(dict(x=x0 + 1.4, y=y, yshift=-0.2 * h_px, text=f"<b>{etq}</b>", showarrow=False,
                         font=dict(size=min(15 * escala, 0.19 * h_px), color=col),
                         xanchor="left", yanchor="middle"))
        if clave != "Meta":
            anot.append(dict(x=x1 - 1.2, y=y, yshift=0.14 * h_px, text=f"<b>{pc:.0f}%</b>",
                             showarrow=False, font=dict(size=min(21 * escala, 0.3 * h_px),
                                                        color="#ffffff"),
                             xanchor="right", yanchor="middle"))
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

    anot.append(dict(xref="paper", yref="paper", x=0.5, y=1.0, text=f"<b>{titulo_txt}</b>",
                     showarrow=False, font=dict(size=min(30 * escala, 36), color="#ffffff"), bgcolor="#e2231a",
                     bordercolor="#ffffff", borderwidth=2, borderpad=10, yanchor="bottom"))
    fig.update_layout(
        shapes=formas, annotations=anot, height=alto, clickmode="event+select", dragmode=False,
        paper_bgcolor="#0a0507", plot_bgcolor="rgba(0,0,0,0)", margin=dict(l=10, r=10, t=int(84 * escala), b=10),
        showlegend=False, hoverlabel=dict(bgcolor="#2a1a1d", font_size=14, font_color="#ffffff"),
        xaxis=dict(visible=False, range=[0, 100], fixedrange=True),
        yaxis=dict(visible=False, range=[(-1.75 if logo_uri else 0.15), y_top], fixedrange=True))
    if logo_uri:
        uri, aspecto = logo_uri
        alto_u = min(112 * escala / px_y, 1.2)               # unidades y (~110 px)
        ancho_u = min(alto_u * px_y * aspecto * 100 / 1400, 34)
        alto_u = ancho_u * 1400 / 100 / (px_y * aspecto)
        fig.add_layout_image(dict(
            source=uri, xref="x", yref="y", x=98, y=-1.6, xanchor="right", yanchor="bottom",
            sizex=ancho_u, sizey=alto_u, sizing="contain", layer="above", opacity=1))
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
            logo_up = st.file_uploader("Logo (PNG o JPG, opcional)", type=["png", "jpg", "jpeg"],
                                       key="pres_logo",
                                       help="También puedes dejar un archivo logo.png junto a "
                                            "app.py en GitHub y se usa automáticamente.")
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
            for nombre_logo in ("logo.png", "logo.jpg", "logo.jpeg"):
                ruta_logo = Path(__file__).parent / nombre_logo
                if ruta_logo.exists():
                    logo_bytes = ruta_logo.read_bytes()
                    break
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
        titulo_arbol = titulo_p + ("" if region_p == "Todas"
                                   else f" · {region_p.replace('Bogota ', '').upper()}")

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

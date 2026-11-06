"""
Dashboard Resumen - Modernización RF (R4 Bogotá)
Sube el archivo de control (.xlsx) y visualiza la pestaña "Resumen".

Ejecutar:
    pip install -r requirements.txt
    streamlit run app.py
"""

import datetime
import io

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
                            territorios: list, proyecto: str) -> pd.DataFrame:
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


def construir_detalle(raw_sitios: pd.DataFrame, territorios: list, proyecto: str) -> pd.DataFrame:
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
    carga = s[COL_SITIOS["carga"]].astype(str).str.strip().str.upper()
    for clave, txt in HITO_TEXTO.items():
        det["ok_" + clave] = carga == txt
    det["Carga"] = s[COL_SITIOS["carga"]].map(lambda x: x if isinstance(x, str) else "")
    det["Autorización"] = s[3].map(lambda x: x if isinstance(x, str) else "")
    det["Scope"] = s[21].map(lambda x: x if isinstance(x, str) else "") if 21 in s.columns else ""

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

    det["Etapa"] = det.apply(etapa, axis=1)
    return det.reset_index(drop=True)


@st.cache_data(show_spinner="Leyendo archivo…")
def cargar(contenido: bytes):
    xls = pd.ExcelFile(io.BytesIO(contenido), engine="openpyxl")
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
            df_recalc, sin_terr = recalcular_desde_sitios(raw_s, df_excel, territorios, total_col)
            det = construir_detalle(raw_s, territorios, total_col)
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
    st.session_state[f"{prefijo}_ver"] = ver
    st.session_state[f"{prefijo}_terr"] = territorio
    st.session_state[f"{prefijo}_modo"] = modo


def grafico_clic(fig, key):
    """Muestra un gráfico de Plotly que responde a clics (Streamlit >= 1.35)."""
    return st.plotly_chart(fig, use_container_width=True, on_select="rerun",
                           selection_mode="points", key=key,
                           config={"displayModeBar": False})


def panel_detalle(prefijo):
    if det is None:
        return
    with st.container(border=True):
        st.markdown('<p class="panel-t">🔍 Sitios detrás de cada número</p>'
                    '<p class="panel-s">Haz clic en una barra de los gráficos de arriba, '
                    'o elige aquí qué quieres ver.</p>', unsafe_allow_html=True)
        opciones_ver = list(df_excel.index) + [PREFIJO_ETAPA + e for e, _ in ETAPAS]
        terr_disp = [t for t in sel_terr]
        opciones_terr = ["Todos"] + terr_disp
        for k, ops in ((f"{prefijo}_ver", opciones_ver), (f"{prefijo}_terr", opciones_terr)):
            if st.session_state.get(k) not in ops:
                st.session_state[k] = ops[0]
        if st.session_state.get(f"{prefijo}_modo") not in ("Cumplen", "Pendientes"):
            st.session_state[f"{prefijo}_modo"] = "Cumplen"

        c1, c2, c3 = st.columns([3, 2, 2])
        ver = c1.selectbox("Ver", opciones_ver, key=f"{prefijo}_ver")
        terr_sel = c2.selectbox("Territorio", opciones_terr, key=f"{prefijo}_terr")
        es_etapa = ver.startswith(PREFIJO_ETAPA)
        es_total = _norm(ver) == "TOTAL SITIOS"
        modo = c3.radio("Mostrar", ["Cumplen", "Pendientes"], key=f"{prefijo}_modo",
                        horizontal=True, disabled=es_etapa or es_total,
                        help="Cumplen: sitios que ya tienen el hito. Pendientes: los que no.")

        d = det[det["Territorio"].isin(sel_terr)]
        if terr_sel != "Todos":
            d = d[d["Territorio"] == terr_sel]
        col_fecha = None
        if es_etapa:
            d = d[d["Etapa"] == ver[len(PREFIJO_ETAPA):]]
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
        if col_fecha:
            columnas.append(col_fecha)
        columnas += [c for c in COLS_DETALLE if c != col_fecha]
        columnas += ["Carga", "Autorización", "Scope"]
        d = d[columnas].sort_values(["Territorio", "Sitio"])

        st.markdown(f"**{len(d):,} sitio(s)** · {titulo_lista}"
                    + ("" if terr_sel == "Todos" else f" · {terr_sel}"))
        if fuente == FUENTE_EXCEL:
            st.caption("La lista sale de la pestaña Sitios con el cálculo corregido; puede "
                       "diferir del Resumen del Excel en las celdas con error.")
        if d.empty:
            st.info("No hay sitios con esa condición.")
            return
        cfg = {c: st.column_config.DateColumn(c, format="DD/MM/YYYY") for c in COLS_DETALLE}
        if col_fecha:
            cfg[col_fecha] = st.column_config.DateColumn(f"⭐ {col_fecha}", format="DD/MM/YYYY")
        st.dataframe(d, hide_index=True, use_container_width=True, column_config=cfg,
                     height=min(38 + 35 * len(d), 460))
        nombre = "".join(ch if ch.isalnum() else "_" for ch in titulo_lista + "_" + terr_sel)
        st.download_button("⬇️ Descargar lista (CSV)", d.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"sitios_{nombre}.csv", mime="text/csv",
                           key=f"{prefijo}_dl")


# ---------------------------------------------------------------------------
# Interfaz
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("📂 Archivo")
    archivo = st.file_uploader("Sube el control (.xlsx)", type=["xlsx", "xlsm"])
    st.caption("Se lee la pestaña **Resumen**. El archivo no se guarda en ningún lado; "
               "solo se procesa en esta sesión.")

st.title("📡 Dashboard Resumen · Modernización")

if archivo is None:
    st.info("Sube el archivo de control en la barra lateral para ver el resumen.")
    st.stop()

try:
    df_excel, df_recalc, sin_terr, det, terr_excel, total_col, info, error_recalc = cargar(archivo.getvalue())
except Exception as e:  # noqa: BLE001
    st.error(f"No se pudo leer el archivo: {e}")
    st.stop()

FUENTE_RECALC = "Recalculado (corregido)"
FUENTE_EXCEL = "Resumen del Excel"

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

if fuente == FUENTE_RECALC:
    df = df_recalc
    territorios = [c for c in df.columns if c != "Total"]
else:
    df = df_excel
    territorios = terr_excel

# ---- Comparación Excel vs recalculado ---------------------------------------
diferencias = []
if df_recalc is not None:
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
           f"Fuente: **{fuente}** · Archivo: `{archivo.name}`")

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
    alertas.append(f"Hay sitios del proyecto sin región en la pestaña Sitios; se muestran en "
                   f"la columna **{SIN_TERRITORIO}** (detalle en 🔧 Correcciones).")
if (fuente == FUENTE_EXCEL and info["ref_encabezado"] and h_total
        and int(info["ref_encabezado"]) != int(df.loc[h_total, "Total"])):
    alertas.append(f"El encabezado indica **{int(info['ref_encabezado'])}** sitios, "
                   f"pero el total de sitios es **{int(df.loc[h_total, 'Total'])}**.")
ceros = [h for h in df.index if h != h_total and h not in [h_neg, h_cos]
         and df.loc[h, "Total"] == 0]
if ceros:
    alertas.append(f"Hitos en cero en todos los territorios: **{', '.join(ceros)}** "
                   f"(revisar fórmula u origen del dato).")
if alertas:
    with st.expander(f"⚠️ {len(alertas)} observación(es) sobre los datos", expanded=False):
        for a in alertas:
            st.markdown(f"- {a}")

# ---- KPIs ------------------------------------------------------------------
c = st.columns(6)
kpi(c[0], "Total sitios", f"{total:,}", f"{len(sel_terr)} territorio(s)")
kpi(c[1], "TSS aprobados", f"{v(h_tss):,}", f"{pct(v(h_tss), total):.0f}% del total")
kpi(c[2], "Instalados", f"{v(h_inst):,}", f"{pct(v(h_inst), total):.0f}% del total")
kpi(c[3], "En servicio (InSrv)", f"{v(h_insrv):,}", f"{pct(v(h_insrv), total):.0f}% del total", "ok")
kpi(c[4], "Pendientes por InSrv", f"{total - v(h_insrv):,}",
    f"{pct(total - v(h_insrv), total):.0f}% del total")
kpi(c[5], "Carga negados · Cosite", f"{v(h_neg)} · {v(h_cos)}", "requieren gestión", "alerta")

st.write("")

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["📊 Embudo", "🗺️ Por territorio", "⏳ Pendientes", "📋 Tabla",
     f"🔧 Correcciones ({len(diferencias)})"])

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
        if pto and pto.get("y") in list(df_excel.index):
            ir_a_detalle("t1", pto["y"])
        st.caption("👆 Haz clic en una barra para ver sus sitios abajo.")

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
                    ir_a_detalle("t1", PREFIJO_ETAPA + trazas_etapa[i])
            chips = "".join(
                f'<span class="chip"><span class="dot" style="background:{c}"></span>'
                f'{nm} <b>{n:,}</b> <span style="opacity:.6">({pct(n, total):.0f}%)</span></span>'
                for nm, n, c in etapas)
            st.markdown(f'<div class="chips">{chips}</div>', unsafe_allow_html=True)

    panel_detalle("t1")

# ---- Tab 2: Por territorio -------------------------------------------------
with tab2:
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
        ir_a_detalle("t2", hito_sel, pto["x"])
    st.caption("👆 Haz clic en un territorio para ver sus sitios abajo.")
    panel_detalle("t2")

# ---- Tab 3: Pendientes -----------------------------------------------------
with tab3:
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
        ir_a_detalle("t3", pto["y"], terr_clic, "Pendientes")
    st.caption("👆 Haz clic en un tramo de color para ver los sitios pendientes de ese "
               "territorio y hito.")

    st.dataframe(pend, use_container_width=True)
    panel_detalle("t3")

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

    st.dataframe(
        mostrar.style
        .background_gradient(cmap="Blues", subset=sel_terr + ["Total"], axis=None)
        .apply(resaltar, axis=None)
        .format({"% del total": "{:.1f}%"} if h_total else {}),
        use_container_width=True, height=470,
    )
    if celdas_dif:
        st.caption(f"🟨 {len(celdas_dif)} celda(s) resaltada(s): el valor corregido es distinto "
                   f"al de la pestaña Resumen del Excel. Detalle en 🔧 Correcciones.")
    st.download_button(
        "⬇️ Descargar CSV", mostrar.to_csv().encode("utf-8-sig"),
        file_name=f"resumen_{'recalculado' if fuente == FUENTE_RECALC else 'excel'}.csv",
        mime="text/csv",
    )

# ---- Tab 5: Correcciones ---------------------------------------------------
with tab5:
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

    h_rfic = buscar(df_excel, "RFIC-OK")
    if h_rfic and df_recalc is not None and int(df_recalc.loc[h_rfic, "Total"]) == 0:
        st.info("**RFIC-OK sigue en 0 aunque ya está bien calculado:** en Sitios un sitio "
                "solo cuenta si **todos** sus trabajos tienen fecha de RFIC, y casi ningún "
                "trabajo la tiene en la pestaña RF. Es un tema de los datos de origen, "
                "no de la fórmula.")

"""
Dashboard Resumen - Modernización RF (R4 Bogotá)
Sube el archivo de control (.xlsx) y visualiza la pestaña "Resumen".

Ejecutar:
    pip install -r requirements.txt
    streamlit run app.py
"""

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


@st.cache_data(show_spinner="Leyendo archivo…")
def cargar(contenido: bytes):
    xls = pd.ExcelFile(io.BytesIO(contenido), engine="openpyxl")
    hoja = next((h for h in xls.sheet_names if h.strip().lower() == HOJA.lower()), None)
    if hoja is None:
        raise ValueError(f"El archivo no tiene una pestaña llamada '{HOJA}'. "
                         f"Pestañas encontradas: {', '.join(xls.sheet_names)}")
    raw = pd.read_excel(xls, sheet_name=hoja, header=None)
    return parsear_resumen(raw)


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
    df, territorios, total_col, info = cargar(archivo.getvalue())
except Exception as e:  # noqa: BLE001
    st.error(f"No se pudo leer el archivo: {e}")
    st.stop()

st.caption(f"Región **{info['region']}** · Proyecto **{total_col}** · Archivo: `{archivo.name}`")

# Filtro de territorios
with st.sidebar:
    st.header("🔎 Filtros")
    sel_terr = st.multiselect("Territorios", territorios, default=territorios)
if not sel_terr:
    st.warning("Selecciona al menos un territorio.")
    st.stop()

color_terr = {t: COLORES_TERRITORIO[i % len(COLORES_TERRITORIO)]
              for i, t in enumerate(territorios)}  # color fijo por territorio
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
if set(sel_terr) == set(territorios):
    dif = df[df["Total archivo"] != df["Total"]]
    for h, r in dif.iterrows():
        alertas.append(f"**{h}**: el total del archivo ({r['Total archivo']}) no coincide "
                       f"con la suma de territorios ({r['Total']}).")
if info["ref_encabezado"] and h_total and int(info["ref_encabezado"]) != int(df.loc[h_total, "Total"]):
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

tab1, tab2, tab3, tab4 = st.tabs(["📊 Embudo", "🗺️ Por territorio", "⏳ Pendientes", "📋 Tabla"])

hitos_embudo = [buscar(df, h) for h in ORDEN_EMBUDO]
hitos_embudo = [h for h in hitos_embudo if h is not None]
# Agrega hitos no previstos (que no sean alertas) al final del embudo
hitos_embudo += [h for h in df.index if h not in hitos_embudo and h not in [h_neg, h_cos]]

# ---- Tab 1: Embudo ---------------------------------------------------------
with tab1:
    col_a, col_b = st.columns([3, 2])
    with col_a:
        st.subheader("Avance por hito")
        valores = [int(dfv.loc[h, "Total"]) for h in hitos_embudo]
        fig = go.Figure(go.Bar(
            x=valores, y=hitos_embudo, orientation="h",
            marker=dict(color=[GRIS if val == 0 else AZUL for val in valores],
                        cornerradius=4),
            text=[f"{val:,}  ·  {pct(val, total):.0f}%" for val in valores],
            textposition="outside", cliponaxis=False,
            hovertemplate="<b>%{y}</b><br>Sitios: %{x:,}<br>%{customdata:.1f}% del total<extra></extra>",
            customdata=[pct(val, total) for val in valores],
        ))
        layout_base(fig, 420)
        fig.update_yaxes(autorange="reversed", showgrid=False)
        fig.update_xaxes(showgrid=True, gridcolor="rgba(128,128,128,.18)",
                         range=[0, max(valores + [1]) * 1.25])
        st.plotly_chart(fig, use_container_width=True)

    with col_b:
        st.subheader("% en servicio")
        p = pct(v(h_insrv), total)
        g = go.Figure(go.Indicator(
            mode="gauge+number", value=p,
            number=dict(suffix="%", valueformat=".1f"),
            gauge=dict(axis=dict(range=[0, 100]), bar=dict(color="#1baf7a", thickness=.35),
                       bgcolor="rgba(128,128,128,.12)", borderwidth=0),
        ))
        layout_base(g, 230)
        st.plotly_chart(g, use_container_width=True)

        # Conversión entre hitos consecutivos (dónde se pierde más)
        st.markdown("**Mayores caídas entre hitos consecutivos**")
        caidas = []
        for a, b in zip(hitos_embudo[:-1], hitos_embudo[1:]):
            va, vb = int(dfv.loc[a, "Total"]), int(dfv.loc[b, "Total"])
            if va > 0 and vb > 0:
                caidas.append({"De": a, "A": b, "Sitios": va - vb})
        caidas = sorted([x for x in caidas if x["Sitios"] > 0], key=lambda x: -x["Sitios"])[:4]
        if caidas:
            st.dataframe(pd.DataFrame(caidas), hide_index=True, use_container_width=True)
        else:
            st.caption("Sin caídas entre hitos.")

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
    st.plotly_chart(bar, use_container_width=True)

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
    st.plotly_chart(fig_p, use_container_width=True)

    st.dataframe(pend, use_container_width=True)

# ---- Tab 4: Tabla ----------------------------------------------------------
with tab4:
    st.subheader("Hoja Resumen")
    mostrar = dfv.copy()
    if h_total:
        mostrar["% del total"] = (mostrar["Total"] / max(total, 1) * 100).round(1)
    st.dataframe(
        mostrar.style
        .background_gradient(cmap="Blues", subset=sel_terr + ["Total"], axis=None)
        .format({"% del total": "{:.1f}%"} if h_total else {}),
        use_container_width=True, height=470,
    )
    st.download_button(
        "⬇️ Descargar CSV", mostrar.to_csv().encode("utf-8-sig"),
        file_name="resumen.csv", mime="text/csv",
    )

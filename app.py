"""
=====================================================================
 PANEL DROPSHIPPING — Meta Ads x Dropi
 App Streamlit que une el reporte de Meta Ads (CSV) con el reporte de
 órdenes de Dropi (CSV o XLSX), calcula ROAS y margen neto por fecha
 y muestra 3 gráficas dinámicas + tabla resumen, en tema oscuro.

 Cómo ejecutarla:
     pip install -r requirements.txt
     streamlit run app.py
=====================================================================
"""

import io
import re
from datetime import timedelta

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

# ---------------------------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA Y TEMA OSCURO
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Panel Dropshipping · Meta Ads x Dropi",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

DARK_CSS = """
<style>
:root{
    --bg:#0e1117; --panel:#161a23; --panel2:#1c212c; --border:#262b38;
    --text:#e8eaf0; --muted:#9aa3b2; --accent:#7c5cff; --good:#22c55e;
    --bad:#ef4444; --warn:#f59e0b;
}
.stApp{ background:var(--bg); color:var(--text); }
section[data-testid="stSidebar"]{ background:var(--panel); border-right:1px solid var(--border); }
h1,h2,h3,h4{ color:var(--text) !important; }
.block-container{ padding-top:1.6rem; }

/* tarjetas de métricas */
div[data-testid="stMetric"]{
    background:linear-gradient(180deg,var(--panel2),var(--panel));
    border:1px solid var(--border); border-radius:14px;
    padding:14px 18px; box-shadow:0 4px 14px rgba(0,0,0,.25);
}
div[data-testid="stMetricLabel"]{ color:var(--muted) !important; }

/* botones y file uploader */
.stButton>button, .stDownloadButton>button{
    background:var(--accent); color:white; border:none; border-radius:10px;
    padding:.5rem 1.1rem; font-weight:600;
}
[data-testid="stFileUploaderDropzone"]{
    background:var(--panel2); border:1.5px dashed var(--border); border-radius:12px;
}

/* tabla */
[data-testid="stDataFrame"]{ border:1px solid var(--border); border-radius:12px; overflow:hidden; }

hr{ border-color:var(--border); }
.badge{
    display:inline-block; padding:2px 10px; border-radius:999px; font-size:.78rem;
    background:rgba(124,92,255,.15); color:#c9bdff; border:1px solid rgba(124,92,255,.35);
}
.callout{
    background:rgba(245,158,11,.08); border:1px solid rgba(245,158,11,.35);
    color:#ffd98a; padding:10px 14px; border-radius:10px; font-size:.9rem;
}
</style>
"""
st.markdown(DARK_CSS, unsafe_allow_html=True)

PLOTLY_TEMPLATE = go.layout.Template(
    layout=go.Layout(
        paper_bgcolor="#161a23",
        plot_bgcolor="#161a23",
        font=dict(color="#e8eaf0"),
        xaxis=dict(gridcolor="#262b38", zerolinecolor="#262b38"),
        yaxis=dict(gridcolor="#262b38", zerolinecolor="#262b38"),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
        colorway=["#7c5cff", "#22c55e", "#f59e0b", "#38bdf8", "#ef4444"],
    )
)

# ---------------------------------------------------------------------------
# 2. FUNCIONES DE CARGA / LIMPIEZA
# ---------------------------------------------------------------------------

def _to_number(series: pd.Series) -> pd.Series:
    """Convierte texto tipo '1,234.5' o vacío -> float, tolerante a NaN."""
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0.0)


@st.cache_data(show_spinner=False)
def cargar_meta_ads(file) -> pd.DataFrame:
    """Lee el export de Meta Ads y regresa columnas normalizadas por fecha."""
    df = pd.read_csv(file)
    df.columns = [c.strip() for c in df.columns]

    col_inicio = next((c for c in df.columns if "Inicio del informe" in c), None)
    col_fin = next((c for c in df.columns if "Fin del informe" in c), None)
    col_gasto = next((c for c in df.columns if "Importe gastado" in c), None)
    col_compras = next((c for c in df.columns if c.strip() == "Compras"), None)
    col_campana = next((c for c in df.columns if "Nombre de la campaña" in c), None)

    if col_inicio is None or col_gasto is None:
        raise ValueError(
            "No encuentro las columnas 'Inicio del informe' / 'Importe gastado' "
            "en el CSV de Meta Ads. Verifica que sea un export estándar de Ads Manager."
        )

    # CORRECCIÓN: Eliminar la fila de "Totales" o cualquier fila vacía antes de procesar fechas
    df = df.dropna(subset=[col_inicio])

    df[col_inicio] = pd.to_datetime(df[col_inicio], errors="coerce")
    df[col_fin] = pd.to_datetime(df[col_fin], errors="coerce") if col_fin else df[col_inicio]
    
    # Asegurar que las fechas no se hayan convertido en "NaT" (Not a Time)
    df = df.dropna(subset=[col_inicio])

    df["gasto_mxn"] = _to_number(df[col_gasto])
    df["compras"] = _to_number(df[col_compras]) if col_compras else 0
    df["campana"] = df[col_campana] if col_campana else "N/D"

    # Al haber borrado las filas vacías, esta condición ahora será True
    rango_por_fila = (df[col_inicio] == df[col_fin]).all()
    dias_unicos = df[col_inicio].nunique()

    filas = []
    if rango_por_fila and dias_unicos > 1:
        # Ya es diario: una fecha por fila.
        for _, r in df.iterrows():
            filas.append({"fecha": r[col_inicio], "gasto_mxn": r["gasto_mxn"],
                          "compras_meta": r["compras"], "campana": r["campana"]})
        modo = "diario"
    else:
        # Reporte por rango de fechas (prorrateado)
        for _, r in df.iterrows():
            ini, fin = r[col_inicio], r[col_fin]
            if pd.isna(ini):
                continue
            fin = fin if pd.notna(fin) else ini
            n_dias = (fin - ini).days + 1
            n_dias = max(n_dias, 1)
            for i in range(n_dias):
                filas.append({
                    "fecha": ini + timedelta(days=i),
                    "gasto_mxn": r["gasto_mxn"] / n_dias,
                    "compras_meta": r["compras"] / n_dias,
                    "campana": r["campana"],
                })
        modo = "prorrateado"

    diario = pd.DataFrame(filas)
    resumen = diario.groupby("fecha", as_index=False).agg(
        gasto_mxn=("gasto_mxn", "sum"),
        compras_meta=("compras_meta", "sum"),
    )
    return resumen, diario, modo


@st.cache_data(show_spinner=False)
def cargar_dropi(file, nombre_archivo: str) -> pd.DataFrame:
    """Lee el export de Dropi (CSV o XLSX) y regresa detalle + resumen diario."""
    if nombre_archivo.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(file)
    else:
        df = pd.read_csv(file)
    df.columns = [c.strip() for c in df.columns]

    col_fecha = next((c for c in df.columns if c.strip().upper() == "FECHA"), df.columns[0])
    col_total = next((c for c in df.columns if "TOTAL DE LA ORDEN" in c.upper()), None)
    col_ganancia = next((c for c in df.columns if c.strip().upper() == "GANANCIA"), None)
    col_proveedor = next((c for c in df.columns if "PRECIO PROVEEDOR X CANTIDAD" in c.upper()), None)
    col_flete = next((c for c in df.columns if "PRECIO FLETE" in c.upper()), None)
    col_estatus = next((c for c in df.columns if "ESTATUS" in c.upper()), None)
    col_producto = next((c for c in df.columns if c.strip().upper() == "PRODUCTO"), None)
    col_cantidad = next((c for c in df.columns if c.strip().upper() == "CANTIDAD"), None)

    if col_total is None:
        raise ValueError(
            "No encuentro la columna 'TOTAL DE LA ORDEN' en el archivo de Dropi. "
            "Verifica que sea el export de órdenes/productos de Dropi."
        )

    df["fecha"] = pd.to_datetime(df[col_fecha], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["fecha"])
    df["total_gtq"] = _to_number(df[col_total])
    df["flete_gtq"] = _to_number(df[col_flete]) if col_flete else 0.0
    df["costo_proveedor_gtq"] = _to_number(df[col_proveedor]) if col_proveedor else 0.0
    df["estatus"] = df[col_estatus].astype(str).str.upper() if col_estatus else "N/D"
    df["producto"] = df[col_producto] if col_producto else "N/D"
    df["cantidad"] = _to_number(df[col_cantidad]) if col_cantidad else 1

    # Si Dropi ya trae GANANCIA calculada la usamos; si no, la estimamos:
    # ganancia = total - costo proveedor - flete
    if col_ganancia and df[col_ganancia].notna().any():
        df["ganancia_gtq"] = _to_number(df[col_ganancia])
        df.loc[df["ganancia_gtq"] == 0, "ganancia_gtq"] = (
            df["total_gtq"] - df["costo_proveedor_gtq"] - df["flete_gtq"]
        )
    else:
        df["ganancia_gtq"] = df["total_gtq"] - df["costo_proveedor_gtq"] - df["flete_gtq"]

    return df


# ---------------------------------------------------------------------------
# 3. BARRA LATERAL — CARGA DE ARCHIVOS Y PARÁMETROS
# ---------------------------------------------------------------------------
st.sidebar.markdown("### 📤 Carga tus archivos")
archivo_meta = st.sidebar.file_uploader("Reporte de Meta Ads (.csv)", type=["csv"], key="meta")
archivo_dropi = st.sidebar.file_uploader("Reporte de órdenes Dropi (.csv o .xlsx)",
                                          type=["csv", "xlsx", "xls"], key="dropi")

st.sidebar.markdown("### ⚙️ Parámetros")
tipo_cambio = st.sidebar.number_input(
    "Tipo de cambio GTQ → MXN", min_value=0.0, value=2.35, step=0.01,
    help="Dropi reporta en quetzales (GTQ); Meta Ads reporta en pesos (MXN). "
         "Ajusta este valor al tipo de cambio del día."
)
comision_retiro_mxn = st.sidebar.number_input(
    "Comisión de retiro por orden (MXN)", min_value=0.0, value=0.0, step=1.0,
    help="Opcional: costo fijo por orden al retirar tus ganancias de Dropi."
)

st.sidebar.markdown("### 🎯 Filtro de estatus Dropi")
solo_entregado = st.sidebar.checkbox(
    "Contar solo órdenes con estatus ENTREGADO", value=False,
    help="Si lo desactivas, se incluyen todas las órdenes generadas (en ruta, "
         "recolectado, entregado, etc.)."
)

# ---------------------------------------------------------------------------
# 4. ENCABEZADO
# ---------------------------------------------------------------------------
st.markdown(
    '<span class="badge">Meta Ads × Dropi</span>',
    unsafe_allow_html=True,
)
st.title("📦 Panel de Rentabilidad — Dropshipping")
st.caption("Une tus reportes, calcula ROAS y margen neto por fecha, y visualiza el desempeño de tus campañas.")

if not archivo_meta or not archivo_dropi:
    st.info("👈 Sube el CSV de Meta Ads **y** el reporte de Dropi en la barra lateral para comenzar.")
    st.stop()

# ---------------------------------------------------------------------------
# 5. PROCESAMIENTO
# ---------------------------------------------------------------------------
try:
    meta_resumen, meta_detalle, modo_meta = cargar_meta_ads(archivo_meta)
except Exception as e:
    st.error(f"Error leyendo el archivo de Meta Ads: {e}")
    st.stop()

try:
    dropi_detalle = cargar_dropi(archivo_dropi, archivo_dropi.name)
except Exception as e:
    st.error(f"Error leyendo el archivo de Dropi: {e}")
    st.stop()

if modo_meta == "prorrateado":
    st.markdown(
        '<div class="callout">⚠️ Tu export de Meta Ads no trae desglose diario '
        '(todas las filas cubren el mismo rango de fechas). El gasto se prorrateó '
        'en partes iguales entre los días del rango para poder unirlo por fecha con Dropi. '
        'Para una unión exacta, exporta desde Ads Manager con el desglose "Por día".</div>',
        unsafe_allow_html=True,
    )

dropi_f = dropi_detalle.copy()
if solo_entregado:
    dropi_f = dropi_f[dropi_f["estatus"] == "ENTREGADO"]

dropi_resumen = dropi_f.groupby("fecha", as_index=False).agg(
    ordenes=("total_gtq", "count"),
    ventas_gtq=("total_gtq", "sum"),
    ganancia_gtq=("ganancia_gtq", "sum"),
)

# --- Unión por fecha ---------------------------------------------------
panel = pd.merge(meta_resumen, dropi_resumen, on="fecha", how="outer").fillna(0)
panel = panel.sort_values("fecha")

# --- Conversión de moneda y métricas -----------------------------------
panel["ventas_mxn"] = panel["ventas_gtq"] * tipo_cambio
panel["ganancia_dropi_mxn"] = panel["ganancia_gtq"] * tipo_cambio
panel["comisiones_mxn"] = panel["ordenes"] * comision_retiro_mxn
panel["margen_neto_mxn"] = panel["ganancia_dropi_mxn"] - panel["gasto_mxn"] - panel["comisiones_mxn"]
panel["roas"] = (panel["ventas_mxn"] / panel["gasto_mxn"]).replace([float("inf")], 0).fillna(0)
panel["fecha_str"] = panel["fecha"].dt.strftime("%d-%b")

if panel.empty:
    st.warning("No hay fechas en común entre ambos archivos con los filtros actuales.")
    st.stop()

# ---------------------------------------------------------------------------
# 6. MÉTRICAS PRINCIPALES
# ---------------------------------------------------------------------------
total_gasto = panel["gasto_mxn"].sum()
total_ventas = panel["ventas_mxn"].sum()
total_ganancia = panel["ganancia_dropi_mxn"].sum()
total_margen = panel["margen_neto_mxn"].sum()
roas_global = (total_ventas / total_gasto) if total_gasto else 0
total_ordenes = int(panel["ordenes"].sum())

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("💸 Gasto en Ads (MXN)", f"${total_gasto:,.2f}")
c2.metric("🛒 Ventas Dropi (MXN)", f"${total_ventas:,.2f}", f"{total_ordenes} órdenes")
c3.metric("📈 ROAS global", f"{roas_global:,.2f}x")
c4.metric("💰 Ganancia Dropi (MXN)", f"${total_ganancia:,.2f}")
c5.metric("🧮 Margen neto (MXN)", f"${total_margen:,.2f}",
          delta_color="normal" if total_margen >= 0 else "inverse")

st.divider()

# ---------------------------------------------------------------------------
# 7. GRÁFICAS DINÁMICAS (3)
# ---------------------------------------------------------------------------
g1, g2 = st.columns(2)

# Gráfica 1: Gasto vs Ventas por fecha (barras + línea combinadas)
with g1:
    st.subheader("Gasto en Ads vs. Ventas por día")
    fig1 = go.Figure()
    fig1.add_bar(x=panel["fecha_str"], y=panel["gasto_mxn"], name="Gasto Ads (MXN)",
                 marker_color="#ef4444")
    fig1.add_bar(x=panel["fecha_str"], y=panel["ventas_mxn"], name="Ventas Dropi (MXN)",
                 marker_color="#22c55e")
    fig1.update_layout(template=PLOTLY_TEMPLATE, barmode="group", height=380,
                        legend=dict(orientation="h", yanchor="bottom", y=1.02))
    st.plotly_chart(fig1, use_container_width=True)

# Gráfica 2: ROAS por fecha (línea)
with g2:
    st.subheader("ROAS por día")
    fig2 = px.line(panel, x="fecha_str", y="roas", markers=True)
    fig2.add_hline(y=1, line_dash="dash", line_color="#f59e0b",
                    annotation_text="Punto de equilibrio (ROAS=1)")
    fig2.update_traces(line_color="#7c5cff", line_width=3, marker=dict(size=8))
    fig2.update_layout(template=PLOTLY_TEMPLATE, height=380, yaxis_title="ROAS (x)")
    st.plotly_chart(fig2, use_container_width=True)

# Gráfica 3: Margen neto acumulado por fecha (área)
st.subheader("Margen neto diario y acumulado")
panel["margen_acumulado"] = panel["margen_neto_mxn"].cumsum()
fig3 = go.Figure()
colores_barra = ["#22c55e" if v >= 0 else "#ef4444" for v in panel["margen_neto_mxn"]]
fig3.add_bar(x=panel["fecha_str"], y=panel["margen_neto_mxn"], name="Margen neto diario (MXN)",
             marker_color=colores_barra)
fig3.add_scatter(x=panel["fecha_str"], y=panel["margen_acumulado"], name="Margen acumulado (MXN)",
                  mode="lines+markers", line=dict(color="#38bdf8", width=3), yaxis="y2")
fig3.update_layout(
    template=PLOTLY_TEMPLATE, height=380,
    yaxis=dict(title="Margen diario (MXN)"),
    yaxis2=dict(title="Margen acumulado (MXN)", overlaying="y", side="right", showgrid=False),
    legend=dict(orientation="h", yanchor="bottom", y=1.02),
)
st.plotly_chart(fig3, use_container_width=True)

st.divider()

# ---------------------------------------------------------------------------
# 8. TABLA RESUMEN
# ---------------------------------------------------------------------------
st.subheader("📋 Tabla resumen por fecha")

tabla = panel[[
    "fecha_str", "ordenes", "gasto_mxn", "ventas_mxn", "ganancia_dropi_mxn",
    "comisiones_mxn", "margen_neto_mxn", "roas",
]].rename(columns={
    "fecha_str": "Fecha", "ordenes": "Órdenes", "gasto_mxn": "Gasto Ads (MXN)",
    "ventas_mxn": "Ventas (MXN)", "ganancia_dropi_mxn": "Ganancia Dropi (MXN)",
    "comisiones_mxn": "Comisiones (MXN)", "margen_neto_mxn": "Margen Neto (MXN)",
    "roas": "ROAS",
})

st.dataframe(
    tabla.style
        .format({
            "Gasto Ads (MXN)": "${:,.2f}", "Ventas (MXN)": "${:,.2f}",
            "Ganancia Dropi (MXN)": "${:,.2f}", "Comisiones (MXN)": "${:,.2f}",
            "Margen Neto (MXN)": "${:,.2f}", "ROAS": "{:,.2f}x",
        })
        .background_gradient(subset=["ROAS"], cmap="Purples")
        .map(lambda v: "color:#22c55e" if isinstance(v, (int, float)) and v >= 0 else "color:#ef4444",
                  subset=["Margen Neto (MXN)"]),
    use_container_width=True, height=380,
)

csv_export = tabla.to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Descargar tabla resumen (CSV)", csv_export,
                    file_name="resumen_dropshipping.csv", mime="text/csv")

# ---------------------------------------------------------------------------
# 9. DESGLOSE OPCIONAL POR PRODUCTO (si el archivo de Dropi trae varios)
# ---------------------------------------------------------------------------
if dropi_f["producto"].nunique() > 1:
    st.divider()
    st.subheader("🏷️ Ventas y ganancia por producto")
    por_producto = dropi_f.groupby("producto", as_index=False).agg(
        ordenes=("total_gtq", "count"),
        ventas_gtq=("total_gtq", "sum"),
        ganancia_gtq=("ganancia_gtq", "sum"),
    )
    por_producto["ventas_mxn"] = por_producto["ventas_gtq"] * tipo_cambio
    por_producto["ganancia_mxn"] = por_producto["ganancia_gtq"] * tipo_cambio
    fig4 = px.bar(por_producto, x="producto", y=["ventas_mxn", "ganancia_mxn"],
                  barmode="group", labels={"value": "MXN", "producto": "Producto"})
    fig4.update_layout(template=PLOTLY_TEMPLATE, height=380,
                        legend=dict(orientation="h", yanchor="bottom", y=1.02))
    st.plotly_chart(fig4, use_container_width=True)

st.caption(
    "Nota: Dropi se asume en quetzales (GTQ) y Meta Ads en pesos mexicanos (MXN), "
    "según tu configuración. Ajusta el tipo de cambio en la barra lateral."
)

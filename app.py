import streamlit as st
import pandas as pd
import io

# Configuración de la página para que se vea amplia en pantallas grandes
st.set_page_config(page_title="Asistente Dropshipping", layout="wide")

st.title("📊 Asistente de Rentabilidad y Dropshipping")

# --- MENÚ LATERAL: MEMORIA Y DATOS ---
st.sidebar.header("1. Carga de Datos")

st.sidebar.write("**A. Sube tu memoria (Opcional si es tu primer día)**")
archivo_historial = st.sidebar.file_uploader("Historial Maestro (Excel/CSV)", type=["csv", "xlsx"], key="historial")

st.sidebar.write("**B. Sube los reportes de HOY**")
archivo_ordenes_hoy = st.sidebar.file_uploader("Ventas Dropi de Hoy", type=["csv", "xlsx"], key="ordenes")
archivo_anuncios_hoy = st.sidebar.file_uploader("Gastos Ads de Hoy", type=["csv", "xlsx"], key="ads")

# --- CONFIGURACIÓN DE DIVISAS ---
st.sidebar.header("2. Divisas")
tipo_cambio = st.sidebar.number_input("Tipo de cambio: 1 Dólar = X Moneda Local", min_value=0.01, value=19.50)

# FUNCIÓN PARA UNIR LA MEMORIA CON LO DE HOY
def procesar_datos(historial, nuevos_datos):
    # Si hay historial y datos nuevos, los pega uno debajo del otro
    if historial is not None and nuevos_datos is not None:
        df_historial = pd.read_csv(historial) if historial.name.endswith('.csv') else pd.read_excel(historial)
        df_nuevos = pd.read_csv(nuevos_datos) if nuevos_datos.name.endswith('.csv') else pd.read_excel(nuevos_datos)
        return pd.concat([df_historial, df_nuevos], ignore_index=True)
    # Si solo hay datos nuevos (primer día)
    elif nuevos_datos is not None:
        return pd.read_csv(nuevos_datos) if nuevos_datos.name.endswith('.csv') else pd.read_excel(nuevos_datos)
    return None

# Procesamos las ventas
df_ventas = procesar_datos(archivo_historial, archivo_ordenes_hoy)

# --- PANEL PRINCIPAL ---
tab1, tab2, tab3, tab4 = st.tabs(["📈 Dashboard y Rentabilidad", "🧮 Calculadora de Precios", "🔗 Enlace de Anuncios", "🧠 Recomendaciones (IA)"])

with tab1:
    st.header("Resumen de tu Negocio")
    if df_ventas is not None:
        st.success("Datos cargados correctamente.")
        
        # Simulamos indicadores rápidos (Estos se calcularían leyendo tus columnas reales)
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Órdenes Totales", "145", "+12 hoy")
        col2.metric("Entregadas", "120", "82%")
        col3.metric("Devoluciones", "15", "10%")
        col4.metric("Cancelaciones", "10", "8%")
        
        st.subheader("Tabla de Rentabilidad Diaria")
        st.write("Semáforo: Compara tu ingreso por ventas vs el costo de anuncios y proveedores.")
        # Tabla de ejemplo de cómo se verá el semáforo
        datos_ejemplo = pd.DataFrame({
            "Fecha": ["23-Sep", "24-Sep"],
            "Ventas Totales": ["$5,000", "$6,200"],
            "Costo Proveedor": ["$1,500", "$1,800"],
            "Gasto Publicidad": ["$800", "$950"],
            "Utilidad Neta": ["$2,700", "$3,450"],
            "Rentable": ["✅ Sí", "✅ Sí"]
        })
        st.dataframe(datos_ejemplo, use_container_width=True)

        # Botón para descargar la nueva memoria unificada
        st.subheader("💾 Guarda tu nueva Memoria")
        st.write("Descarga este archivo y súbelo mañana en 'Historial Maestro' para no perder tu información.")
        
        # Convertimos el archivo a Excel para descargar
        buffer = io.BytesIO()
        df_ventas.to_excel(buffer, index=False)
        st.download_button(
            label="⬇️ Descargar Nuevo Historial Maestro",
            data=buffer.getvalue(),
            file_name="Historial_Maestro_Actualizado.xlsx",
            mime="application/vnd.ms-excel"
        )
    else:
        st.info("Sube tus archivos en el menú lateral para ver tu Dashboard.")

with tab2:
    st.header("Calculadora de Precios Inteligente")
    tipo_prod = st.radio("Tipo de Producto", ("Nuevo (Estimación manual)", "Existente (+15 días con datos reales)"))
    
    colA, colB = st.columns(2)
    with colA:
        costo_prov = st.number_input("Costo del Proveedor", value=100.0)
        costo_flete = st.number_input("Costo de Envío/Flete", value=150.0)
        cpa_esperado = st.number_input("Costo por Compra en Ads (CPA)", value=80.0)
    
    with colB:
        if tipo_prod == "Nuevo (Estimación manual)":
            tasa_dev = st.slider("Tasa de Devoluciones (%)", 0, 100, 15) / 100
            tasa_canc = st.slider("Tasa de Cancelaciones (%)", 0, 100, 5) / 100
        else:
            st.success("Usando historial de tu base de datos: Dev 12%, Canc 4%")
            tasa_dev, tasa_canc = 0.12, 0.04
            
    precio_sugerido = (costo_prov + costo_flete + cpa_esperado) / (1 - tasa_dev - tasa_canc)
    st.metric("Precio Mínimo de Venta Sugerido (Punto de Equilibrio)", f"${precio_sugerido:,.2f}")

with tab3:
    st.header("🔗 Enlazar Publicidad con Tienda")
    st.write("Asigna cada campaña publicitaria a su producto correspondiente para medir el costo exacto.")
    
    col_camp, col_prod = st.columns(2)
    with col_camp:
        st.write("**Campañas Activas (De tu archivo Ads)**")
        st.write("1. Campaña_Reloj_TikTok")
        st.write("2. Campaña_Audifonos_FB")
        
    with col_prod:
        st.write("**Selecciona el producto (De tu archivo Dropi)**")
        prod1 = st.selectbox("Producto para campaña 1", ["Reloj Inteligente", "Audífonos Pro", "Lámpara LED"], key="p1")
        prod2 = st.selectbox("Producto para campaña 2", ["Audífonos Pro", "Reloj Inteligente", "Lámpara LED"], index=1, key="p2")
        
    st.button("Guardar Enlaces")

with tab4:
    st.header("🧠 Recomendaciones Estratégicas")
    st.write("Basado en el cálculo de tu Punto de Equilibrio y tu historial de ventas:")
    
    st.warning("⚠️ **Apagar:** El anuncio 'Campaña_Audifonos_FB' está costando $120 por venta, pero tu punto de equilibrio es $90. Estás perdiendo dinero.")
    st.success("🚀 **Escalar:** El anuncio 'Campaña_Reloj_TikTok' está consiguiendo ventas a $45. Tu margen de ganancia es altísimo. ¡Aumenta el presupuesto un 20% hoy!")
    st.info("💡 **Producto:** El 'Reloj Inteligente' tiene una tasa de entrega del 95%. Se recomienda buscar más proveedores para este tipo de producto.")

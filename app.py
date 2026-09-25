import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Asistente Dropshipping", layout="wide")

# 1. CONEXIÓN A GOOGLE SHEETS
conn = st.connection("gsheets", type=GSheetsConnection)

# 2. PANTALLA DE INICIO DE SESIÓN
def login():
    st.title("🔒 Acceso al Asistente")
    try:
        df_usuarios = conn.read(worksheet="Usuarios", ttl=0)
        
        # --- ESTA LÍNEA NUEVA LIMPIA LOS ESPACIOS FANTASMAS ---
        df_usuarios.columns = df_usuarios.columns.str.strip() 
        
    except Exception as e:
        st.error("Error al conectar con Google Sheets. Revisa tus secretos en Streamlit.")
        return False

    usuario = st.text_input("Usuario")
    contrasena = st.text_input("Contraseña", type="password")

    
    if st.button("Entrar"):
        if usuario in df_usuarios['Usuario'].values:
            pass_correcta = df_usuarios.loc[df_usuarios['Usuario'] == usuario, 'Contraseña'].values[0]
            if str(contrasena) == str(pass_correcta):
                st.session_state['logueado'] = True
                st.session_state['usuario_actual'] = usuario
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
        else:
            st.error("Usuario no encontrado.")
            
    return st.session_state.get('logueado', False)

if 'logueado' not in st.session_state:
    st.session_state['logueado'] = False

if not st.session_state['logueado']:
    login()
    st.stop()

# ---------------------------------------------------------
# APLICACIÓN PRINCIPAL
# ---------------------------------------------------------
usuario_activo = st.session_state['usuario_actual']
st.sidebar.success(f"👤 Sesión activa: {usuario_activo}")
if st.sidebar.button("Cerrar Sesión"):
    st.session_state['logueado'] = False
    st.session_state['usuario_actual'] = None
    st.rerun()

st.title("📊 Asistente de Rentabilidad para Dropshipping")

# --- SECCIÓN: CONFIGURACIÓN DE DIVISAS ---
st.sidebar.header("1. Configuración de Divisas")
divisas_opciones = ["MXN", "GTQ", "USD", "COP", "EUR"]

divisa_origen = st.sidebar.selectbox("Divisa ORIGEN (Utilidad/Retiro)", divisas_opciones, index=0)
divisa_ads = st.sidebar.selectbox("Divisa ADS (Publicidad)", divisas_opciones, index=0)
divisa_dropi = st.sidebar.selectbox("Divisa DROPI (Artículos)", divisas_opciones, index=1)

st.sidebar.write(f"**Valor respecto a 1 {divisa_origen}:**")
tasa_ads = st.sidebar.number_input(f"1 {divisa_ads} equivale a:", value=1.000)
tasa_dropi = st.sidebar.number_input(f"1 {divisa_dropi} equivale a:", value=0.430)

# --- SECCIÓN: CARGA DE DATOS ---
st.sidebar.header("2. Cargar Reportes del Día")
archivo_ordenes_hoy = st.sidebar.file_uploader("Sube Ventas (Dropi)", type=["csv", "xlsx"])
archivo_anuncios_hoy = st.sidebar.file_uploader("Sube Ads (Publicidad)", type=["csv", "xlsx"])

if st.sidebar.button("💾 Guardar Hoy en Memoria"):
    if archivo_ordenes_hoy is not None and archivo_anuncios_hoy is not None:
        with st.spinner('Guardando datos para tu usuario...'):
            df_ordenes_hoy = pd.read_excel(archivo_ordenes_hoy)
            df_ads_hoy = pd.read_csv(archivo_anuncios_hoy)
            
            # Etiquetar los datos con el usuario activo
            df_ordenes_hoy['Usuario'] = usuario_activo
            df_ads_hoy['Usuario'] = usuario_activo
            
            memoria_ordenes = conn.read(worksheet="Memoria_Ordenes")
            memoria_ads = conn.read(worksheet="Memoria_Ads")
            
            nuevas_ordenes = pd.concat([memoria_ordenes, df_ordenes_hoy], ignore_index=True)
            nuevos_ads = pd.concat([memoria_ads, df_ads_hoy], ignore_index=True)
            
            conn.update(worksheet="Memoria_Ordenes", data=nuevas_ordenes)
            conn.update(worksheet="Memoria_Ads", data=nuevos_ads)
            
            st.sidebar.success("¡Datos guardados y aislados correctamente!")
    else:
        st.sidebar.warning("Sube ambos archivos primero.")

# --- LECTURA Y FILTRADO DE MEMORIA ---
df_ordenes_full = conn.read(worksheet="Memoria_Ordenes")
df_ads_full = conn.read(worksheet="Memoria_Ads")

# Filtrar para que el usuario solo vea sus propios datos
if not df_ordenes_full.empty and 'Usuario' in df_ordenes_full.columns:
    df_ordenes = df_ordenes_full[df_ordenes_full['Usuario'] == usuario_activo]
else:
    df_ordenes = pd.DataFrame()

if not df_ads_full.empty and 'Usuario' in df_ads_full.columns:
    df_ads = df_ads_full[df_ads_full['Usuario'] == usuario_activo]
else:
    df_ads = pd.DataFrame()

# --- PANEL PRINCIPAL ---
tab1, tab2, tab3 = st.tabs(["📈 Dashboard y Rentabilidad", "🧮 Calculadora de Precios", "🧠 Recomendaciones y Ads"])

with tab1:
    st.header("Resumen General")
    if not df_ordenes.empty:
        total_ordenes = len(df_ordenes)
        entregadas = len(df_ordenes[df_ordenes['ESTATUS'] == 'ENTREGADO'])
        canceladas = len(df_ordenes[df_ordenes['ESTATUS'] == 'CANCELADO'])
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Órdenes Totales", total_ordenes)
        col2.metric("Entregadas", entregadas)
        col3.metric("Cancelaciones", canceladas)
        col4.metric("% Cancelación", f"{(canceladas/total_ordenes)*100 if total_ordenes > 0 else 0:.1f}%")

        st.subheader(f"Rentabilidad en {divisa_origen}")
        
        # Conversión de divisas al vuelo para los cálculos
        ingreso_total = df_ordenes['TOTAL DE LA ORDEN'].sum() * tasa_dropi
        costo_flete = df_ordenes['PRECIO FLETE'].sum() * tasa_dropi
        costo_proveedor = df_ordenes['PRECIO PROVEEDOR X CANTIDAD'].sum() * tasa_dropi
        
        gasto_ads = df_ads['Importe gastado (MXN)'].sum() * tasa_ads if not df_ads.empty else 0
        
        utilidad = ingreso_total - costo_flete - costo_proveedor - gasto_ads
        
        st.metric("Utilidad Total (Bruta)", f"${utilidad:,.2f} {divisa_origen}")
    else:
        st.info("Tu memoria está vacía. Sube reportes para comenzar.")

with tab2:
    st.header(f"Calculadora de Precios Inteligente ({divisa_origen})")
    colA, colB = st.columns(2)
    with colA:
        costo_prov = st.number_input(f"Costo Proveedor ({divisa_dropi})", value=45.0) * tasa_dropi
        costo_flete_est = st.number_input(f"Envío Promedio ({divisa_dropi})", value=50.0) * tasa_dropi
        cpa_esperado = st.number_input(f"CPA Esperado ({divisa_ads})", value=50.0) * tasa_ads
    
    with colB:
        tasa_dev = st.slider("Tasa Devoluciones (%)", 0, 100, 15) / 100
        tasa_canc = st.slider("Tasa Cancelaciones (%)", 0, 100, 5) / 100
            
    precio_sugerido_origen = (costo_prov + costo_flete_est + cpa_esperado) / (1 - tasa_dev - tasa_canc)
    precio_sugerido_dropi = precio_sugerido_origen / tasa_dropi if tasa_dropi > 0 else 0
    
    st.success(f"Precio Mínimo de Venta: **${precio_sugerido_dropi:,.2f} {divisa_dropi}** (Equivalente a ${precio_sugerido_origen:,.2f} {divisa_origen})")

with tab3:
    st.header("Analizador de Anuncios y Recomendaciones")
    if not df_ads.empty and 'Nombre de la campaña' in df_ads.columns:
        for index, row in df_ads.iterrows():
            costo_ads_original = row['Costo por compra (MXN)']
            if pd.notna(costo_ads_original):
                costo_origen = costo_ads_original * tasa_ads
                nombre = row['Nombre de la campaña']
                
                # Ajusta tus límites de CPA según tu estrategia
                limite_alto = 150 * tasa_ads
                limite_bajo = 80 * tasa_ads
                
                if costo_origen > limite_alto:
                    st.warning(f"⚠️ **Apagar:** '{nombre}' tiene un CPA de ${costo_origen:.2f} {divisa_origen}.")
                elif costo_origen < limite_bajo:
                    st.success(f"🚀 **Escalar:** '{nombre}' va excelente con un CPA de ${costo_origen:.2f} {divisa_origen}.")
    else:
        st.write("No hay datos de anuncios registrados para tu usuario.")

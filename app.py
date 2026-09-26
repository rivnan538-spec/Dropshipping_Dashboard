"""
================================================================================
ASISTENTE DE RENTABILIDAD PARA DROPSHIPPING
Streamlit + Google Sheets — multiusuario, multidivisa, con memoria histórica.
Fuentes de datos soportadas: reportes de órdenes de Dropi (.xlsx) y reportes
de Anuncios de Meta Ads Manager a nivel "Anuncio" con desglose de Campaña,
Conjunto de anuncios y Anuncio (.csv).
================================================================================
"""

import re
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Asistente Dropshipping", layout="wide", page_icon="📊")

# ==============================================================================
# CONSTANTES
# ==============================================================================
DIVISAS = ["MXN", "GTQ", "COP", "USD", "EUR", "GBP", "CAD", "PEN", "CLP", "ARS", "BOB", "HNL", "PAB"]

ESTATUS_ENTREGADO = {"ENTREGADO"}
ESTATUS_DEVOLUCION = {"DEVOLUCION", "DEVOLUCIÓN"}
ESTATUS_CANCELADO = {"CANCELADO"}

HOJAS = {
    "usuarios": "Usuarios",
    "config": "Config_Usuario",
    "ordenes": "Memoria_Ordenes",
    "ads": "Memoria_Ads",
    "vinculacion": "Vinculacion_Producto",
    "tipo_cambio": "Tipo_Cambio_Historico",
}

COLS_CONFIG = [
    "Usuario", "Divisa_Origen", "Divisa_Ads", "Divisa_Dropi",
    "Usar_TC_Manual", "TC_Ads_Manual", "TC_Dropi_Manual",
    "Usar_Semaforo_Manual", "CPA_Bajo_Manual", "CPA_Alto_Manual",
    "Usar_Tasas_Manual", "Tasa_Devolucion_Manual", "Tasa_Cancelacion_Manual",
]
COLS_VINCULACION = ["Usuario", "Campaña", "Conjunto_Anuncios", "Anuncio", "Producto"]
COLS_TIPO_CAMBIO = ["Fecha", "Divisa_Base", "Divisa_Destino", "Tasa", "Fuente"]

DEFAULT_CONFIG = {
    "Divisa_Origen": "MXN", "Divisa_Ads": "MXN", "Divisa_Dropi": "GTQ",
    "Usar_TC_Manual": False, "TC_Ads_Manual": 1.0, "TC_Dropi_Manual": 0.43,
    "Usar_Semaforo_Manual": False, "CPA_Bajo_Manual": 80.0, "CPA_Alto_Manual": 150.0,
    "Usar_Tasas_Manual": False, "Tasa_Devolucion_Manual": 0.15, "Tasa_Cancelacion_Manual": 0.05,
}

API_TIPO_CAMBIO = "https://open.er-api.com/v6/latest/{base}"

RANGOS_FECHA = ["Hoy", "Últimos 7 días", "Último mes", "Personalizado"]

conn = st.connection("gsheets", type=GSheetsConnection)


# ==============================================================================
# UTILIDADES DE GOOGLE SHEETS
# ==============================================================================
def leer_hoja(nombre, ttl=60):
    try:
        df = conn.read(worksheet=nombre, ttl=ttl)
        df = df.dropna(how="all")
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()


def verificar_columnas(df, columnas_requeridas, nombre_hoja):
    """Muestra un error claro en vez de un KeyError crudo si faltan columnas esperadas."""
    faltantes = [c for c in columnas_requeridas if c not in df.columns]
    if faltantes:
        st.error(
            f"En tu hoja **{nombre_hoja}** faltan estas columnas: {faltantes}. "
            f"Columnas encontradas: {list(df.columns)}. "
            "Revisa el encabezado (fila 1) de esa pestaña en Google Sheets — seguramente quedó "
            "una fila vieja sin esa columna o el nombre no coincide exactamente (mayúsculas, "
            "acentos o espacios). Si tienes datos de prueba incompatibles, lo más rápido es "
            "borrar el contenido de esa pestaña (dejando solo el encabezado correcto) y volver "
            "a subir tus reportes desde el Panel de Control."
        )
        return False
    return True


def _filtrar_usuario(df, usuario):
    if df.empty or "Usuario" not in df.columns:
        return pd.DataFrame(columns=df.columns)
    return df[df["Usuario"].astype(str).str.strip() == usuario].copy()


# ==============================================================================
# LOGIN
# ==============================================================================
def login():
    st.title("🔒 Acceso al Asistente")
    try:
        df_usuarios = leer_hoja(HOJAS["usuarios"], ttl=0)
        df_usuarios.columns = df_usuarios.columns.str.strip()
    except Exception:
        st.error("Error al conectar con Google Sheets.")
        return False

    usuario = st.text_input("Usuario")
    contrasena = st.text_input("Contraseña", type="password")

    if st.button("Entrar"):
        usuarios_lista = df_usuarios["Usuario"].astype(str).str.strip().values
        usuario_limpio = usuario.strip()
        if usuario_limpio in usuarios_lista:
            pass_correcta = str(
                df_usuarios.loc[df_usuarios["Usuario"].astype(str).str.strip() == usuario_limpio, "Contraseña"].values[0]
            ).strip()
            if pass_correcta.endswith(".0"):
                pass_correcta = pass_correcta[:-2]
            if str(contrasena).strip() == pass_correcta:
                st.session_state["logueado"] = True
                st.session_state["usuario_actual"] = usuario_limpio
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
        else:
            st.error("Usuario no encontrado.")
    return st.session_state.get("logueado", False)


# ==============================================================================
# CONFIGURACIÓN POR USUARIO
# ==============================================================================
def cargar_config(usuario):
    df = leer_hoja(HOJAS["config"], ttl=30)
    fila = _filtrar_usuario(df, usuario)
    if fila.empty:
        return dict(DEFAULT_CONFIG)
    d = fila.iloc[0].to_dict()
    salida = dict(DEFAULT_CONFIG)
    for k in COLS_CONFIG:
        if k in d and pd.notna(d[k]):
            salida[k] = d[k]
    for k in ["Usar_TC_Manual", "Usar_Semaforo_Manual", "Usar_Tasas_Manual"]:
        salida[k] = str(salida[k]).strip().lower() in ("true", "1", "verdadero", "sí", "si")
    for k in ["TC_Ads_Manual", "TC_Dropi_Manual", "CPA_Bajo_Manual", "CPA_Alto_Manual",
              "Tasa_Devolucion_Manual", "Tasa_Cancelacion_Manual"]:
        try:
            salida[k] = float(salida[k])
        except (ValueError, TypeError):
            salida[k] = DEFAULT_CONFIG[k]
    return salida


def guardar_config(usuario, nueva_config: dict):
    df = leer_hoja(HOJAS["config"], ttl=0)
    if df.empty:
        df = pd.DataFrame(columns=COLS_CONFIG)
    if "Usuario" in df.columns:
        df = df[df["Usuario"].astype(str).str.strip() != usuario]
    fila = {c: nueva_config.get(c) for c in COLS_CONFIG}
    fila["Usuario"] = usuario
    df = pd.concat([df, pd.DataFrame([fila])], ignore_index=True)
    conn.update(worksheet=HOJAS["config"], data=df)
    st.cache_data.clear()


# ==============================================================================
# TIPO DE CAMBIO
# ==============================================================================
@st.cache_data(ttl=6 * 3600, show_spinner=False)
def _obtener_tasas_api(base):
    try:
        r = requests.get(API_TIPO_CAMBIO.format(base=base), timeout=8)
        r.raise_for_status()
        data = r.json()
        if data.get("result") == "success":
            return data["rates"]
    except Exception:
        return None
    return None


def registrar_tasa_hoy(base, destino):
    """Obtiene (si hace falta) y guarda en Tipo_Cambio_Historico la tasa base->destino de HOY."""
    if base == destino:
        return 1.0
    hoy = date.today().isoformat()
    df_tc = leer_hoja(HOJAS["tipo_cambio"], ttl=60)
    if not df_tc.empty:
        m = (df_tc["Fecha"].astype(str) == hoy) & (df_tc["Divisa_Base"] == base) & (df_tc["Divisa_Destino"] == destino)
        if m.any():
            return float(df_tc[m].iloc[-1]["Tasa"])
    rates = _obtener_tasas_api(base)
    if rates and destino in rates:
        tasa = float(rates[destino])
        df_tc2 = leer_hoja(HOJAS["tipo_cambio"], ttl=0)
        if df_tc2.empty:
            df_tc2 = pd.DataFrame(columns=COLS_TIPO_CAMBIO)
        nueva = {"Fecha": hoy, "Divisa_Base": base, "Divisa_Destino": destino, "Tasa": tasa, "Fuente": "API"}
        df_tc2 = pd.concat([df_tc2, pd.DataFrame([nueva])], ignore_index=True)
        conn.update(worksheet=HOJAS["tipo_cambio"], data=df_tc2)
        return tasa
    return None


def registrar_tasa_manual(base, destino, tasa):
    if base == destino:
        return
    hoy = date.today().isoformat()
    df_tc = leer_hoja(HOJAS["tipo_cambio"], ttl=0)
    if df_tc.empty:
        df_tc = pd.DataFrame(columns=COLS_TIPO_CAMBIO)
    else:
        m = (df_tc["Fecha"].astype(str) == hoy) & (df_tc["Divisa_Base"] == base) & (df_tc["Divisa_Destino"] == destino)
        df_tc = df_tc[~m]
    nueva = {"Fecha": hoy, "Divisa_Base": base, "Divisa_Destino": destino, "Tasa": float(tasa), "Fuente": "Manual"}
    df_tc = pd.concat([df_tc, pd.DataFrame([nueva])], ignore_index=True)
    conn.update(worksheet=HOJAS["tipo_cambio"], data=df_tc)


def construir_mapa_tasas(df_tc_historico, base, destino):
    mapa = {}
    if df_tc_historico is not None and not df_tc_historico.empty:
        sub = df_tc_historico[(df_tc_historico["Divisa_Base"] == base) & (df_tc_historico["Divisa_Destino"] == destino)]
        for _, row in sub.iterrows():
            try:
                mapa[str(row["Fecha"])] = float(row["Tasa"])
            except (ValueError, TypeError):
                continue
    return mapa


def convertir_columna(df, col_fecha, col_valor, base, destino, df_tc_historico, tasa_manual_fallback=None):
    """Convierte col_valor de `base` a `destino` usando la tasa histórica del día de cada fila."""
    if df.empty:
        return pd.Series(dtype=float)
    if base == destino:
        return pd.to_numeric(df[col_valor], errors="coerce").fillna(0)
    mapa = construir_mapa_tasas(df_tc_historico, base, destino)
    fechas_ordenadas = sorted(mapa.keys())

    def tasa_de(fecha):
        fecha = str(fecha)
        if fecha in mapa:
            return mapa[fecha]
        anteriores = [f for f in fechas_ordenadas if f <= fecha]
        if anteriores:
            return mapa[anteriores[-1]]
        if fechas_ordenadas:
            return mapa[fechas_ordenadas[0]]
        return tasa_manual_fallback if tasa_manual_fallback else 1.0

    tasas_serie = df[col_fecha].astype(str).map(tasa_de)
    return pd.to_numeric(df[col_valor], errors="coerce").fillna(0) * tasas_serie


# ==============================================================================
# CARGA Y LIMPIEZA DE DATOS
# ==============================================================================
def detectar_columna_divisa(df, prefijo):
    """Busca una columna tipo 'Importe gastado (MXN)' y devuelve (columna, divisa)."""
    for c in df.columns:
        m = re.match(rf"{re.escape(prefijo)}\s*\(([A-Za-z]{{3}})\)", str(c))
        if m:
            return c, m.group(1).upper()
    return None, None


def cargar_ordenes(usuario):
    df = _filtrar_usuario(leer_hoja(HOJAS["ordenes"], ttl=120), usuario)
    if df.empty:
        return df
    df["FECHA_dt"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
    df["FECHA_iso"] = df["FECHA_dt"].dt.strftime("%Y-%m-%d")
    for c in ["TOTAL DE LA ORDEN", "PRECIO FLETE", "COSTO DEVOLUCION FLETE", "COMISION",
              "PRECIO PROVEEDOR X CANTIDAD", "CANTIDAD"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df = df.dropna(subset=["FECHA_dt"])
    return df


def cargar_ads(usuario):
    df = _filtrar_usuario(leer_hoja(HOJAS["ads"], ttl=120), usuario)
    if df.empty:
        return df, None
    df["FECHA_iso"] = pd.to_datetime(df["Inicio del informe"], errors="coerce").dt.strftime("%Y-%m-%d")
    col_gasto, divisa_detectada = detectar_columna_divisa(df, "Importe gastado")
    col_cpa, _ = detectar_columna_divisa(df, "Costo por compra")
    df["_gasto"] = pd.to_numeric(df[col_gasto], errors="coerce").fillna(0) if col_gasto else 0.0
    df["_cpa_meta"] = pd.to_numeric(df[col_cpa], errors="coerce") if col_cpa else np.nan
    df["_alcance"] = pd.to_numeric(df.get("Alcance", 0), errors="coerce").fillna(0)
    df = df.dropna(subset=["FECHA_iso"])
    return df, divisa_detectada


def guardar_ordenes(usuario, df_nuevo):
    df_nuevo = df_nuevo.copy()
    df_nuevo["Usuario"] = usuario
    df_nuevo["Fecha_Carga"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    df_actual = leer_hoja(HOJAS["ordenes"], ttl=0)
    if df_actual.empty:
        df_final = df_nuevo
        agregadas = len(df_nuevo)
    else:
        if "ID" in df_actual.columns and "ID" in df_nuevo.columns:
            existentes = set(zip(df_actual["Usuario"].astype(str), df_actual["ID"].astype(str)))
            antes = len(df_nuevo)
            df_nuevo = df_nuevo[~df_nuevo.apply(lambda r: (str(r["Usuario"]), str(r["ID"])) in existentes, axis=1)]
            agregadas = len(df_nuevo)
        else:
            agregadas = len(df_nuevo)
        df_final = pd.concat([df_actual, df_nuevo], ignore_index=True)
    conn.update(worksheet=HOJAS["ordenes"], data=df_final)
    st.cache_data.clear()
    return agregadas


def guardar_ads(usuario, df_nuevo):
    df_nuevo = df_nuevo.copy()
    df_nuevo["Usuario"] = usuario
    df_nuevo["Fecha_Carga"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    claves = ["Usuario", "Nombre de la campaña", "Nombre del conjunto de anuncios", "Nombre del anuncio", "Inicio del informe"]
    df_actual = leer_hoja(HOJAS["ads"], ttl=0)
    if df_actual.empty:
        df_final = df_nuevo
        agregadas = len(df_nuevo)
    else:
        if all(c in df_actual.columns for c in claves) and all(c in df_nuevo.columns for c in claves):
            existentes = set(df_actual[claves].astype(str).apply(tuple, axis=1))
            df_nuevo = df_nuevo[~df_nuevo[claves].astype(str).apply(tuple, axis=1).isin(existentes)]
        agregadas = len(df_nuevo)
        df_final = pd.concat([df_actual, df_nuevo], ignore_index=True)
    conn.update(worksheet=HOJAS["ads"], data=df_final)
    st.cache_data.clear()
    return agregadas


def cargar_vinculacion(usuario):
    df = _filtrar_usuario(leer_hoja(HOJAS["vinculacion"], ttl=60), usuario)
    if df.empty:
        return pd.DataFrame(columns=COLS_VINCULACION)
    return df


def guardar_vinculo(usuario, campana, conjunto, anuncio, producto):
    df = leer_hoja(HOJAS["vinculacion"], ttl=0)
    if df.empty:
        df = pd.DataFrame(columns=COLS_VINCULACION)
    mask = (
        (df.get("Usuario") == usuario) & (df.get("Campaña") == campana)
        & (df.get("Conjunto_Anuncios") == conjunto) & (df.get("Anuncio") == anuncio)
    )
    df = df[~mask]
    nueva = {"Usuario": usuario, "Campaña": campana, "Conjunto_Anuncios": conjunto, "Anuncio": anuncio, "Producto": producto}
    df = pd.concat([df, pd.DataFrame([nueva])], ignore_index=True)
    conn.update(worksheet=HOJAS["vinculacion"], data=df)
    st.cache_data.clear()


# ==============================================================================
# HELPERS DE NEGOCIO
# ==============================================================================
def selector_rango_fechas(key, index_default=1):
    eleccion = st.radio("Rango de fechas", RANGOS_FECHA, index=index_default, horizontal=True, key=f"radio_{key}")
    hoy = date.today()
    if eleccion == "Hoy":
        return hoy, hoy
    if eleccion == "Últimos 7 días":
        return hoy - timedelta(days=6), hoy
    if eleccion == "Último mes":
        return hoy - timedelta(days=30), hoy
    c1, c2 = st.columns(2)
    ini = c1.date_input("Desde", hoy - timedelta(days=30), key=f"ini_{key}")
    fin = c2.date_input("Hasta", hoy, key=f"fin_{key}")
    return ini, fin


def calcular_semaforo(valor_actual, serie_historica, usar_manual=False, umbral_bajo=None, umbral_alto=None):
    """Métrica tipo CPA (menor = mejor). Devuelve (emoji, texto)."""
    if valor_actual is None or pd.isna(valor_actual):
        return "⚪", "Sin datos"
    if usar_manual and umbral_bajo is not None and umbral_alto is not None:
        if valor_actual <= umbral_bajo:
            return "🟢", "Bueno (bajo tu umbral)"
        if valor_actual <= umbral_alto:
            return "🟡", "Regular"
        return "🔴", "Malo (sobre tu umbral)"
    serie_valida = pd.to_numeric(serie_historica, errors="coerce").dropna()
    if len(serie_valida) < 3:
        return "⚪", "Historial insuficiente"
    p33, p66 = serie_valida.quantile([0.33, 0.66])
    if valor_actual <= p33:
        return "🟢", "Mejor que tu propio histórico"
    if valor_actual <= p66:
        return "🟡", "Dentro de tu promedio histórico"
    return "🔴", "Peor que tu propio histórico"


def tasas_devolucion_cancelacion(df_ordenes, config):
    if config["Usar_Tasas_Manual"]:
        return config["Tasa_Devolucion_Manual"], config["Tasa_Cancelacion_Manual"]
    total = len(df_ordenes)
    if total == 0:
        return 0.0, 0.0
    dev = df_ordenes["ESTATUS"].isin(ESTATUS_DEVOLUCION).sum() / total
    canc = df_ordenes["ESTATUS"].isin(ESTATUS_CANCELADO).sum() / total
    return dev, canc


# ==============================================================================
# SESIÓN / LOGIN
# ==============================================================================
if "logueado" not in st.session_state:
    st.session_state["logueado"] = False

if not st.session_state["logueado"]:
    login()
    st.stop()

usuario_activo = st.session_state["usuario_actual"]
config = cargar_config(usuario_activo)
DO, DA, DD = config["Divisa_Origen"], config["Divisa_Ads"], config["Divisa_Dropi"]

st.sidebar.success(f"👤 Sesión activa: {usuario_activo}")
if st.sidebar.button("Cerrar Sesión"):
    st.session_state["logueado"] = False
    st.session_state["usuario_actual"] = None
    st.rerun()
st.sidebar.caption(f"Divisa origen actual: **{DO}** — cámbiala en ⚙️ Panel de Control")

# Aseguramos que la tasa de HOY quede registrada en el histórico (automático)
if not config["Usar_TC_Manual"]:
    registrar_tasa_hoy(DD, DO)
    registrar_tasa_hoy(DA, DO)

df_tc = leer_hoja(HOJAS["tipo_cambio"], ttl=300)
df_ordenes = cargar_ordenes(usuario_activo)
df_ads, divisa_ads_detectada = cargar_ads(usuario_activo)
df_vinc = cargar_vinculacion(usuario_activo)

TC_ADS_MANUAL = config["TC_Ads_Manual"] if config["Usar_TC_Manual"] else None
TC_DROPI_MANUAL = config["TC_Dropi_Manual"] if config["Usar_TC_Manual"] else None

st.title("📊 Asistente de Rentabilidad para Dropshipping")

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["📊 Dashboard", "📣 Anuncios", "🧮 Calculadora", "⚙️ Panel de Control", "💹 Rentabilidad", "🧭 Recomendaciones"]
)

# ==============================================================================
# TAB 1 — DASHBOARD
# ==============================================================================
COLS_ORDENES_REQUERIDAS = ["ID", "ESTATUS", "TOTAL DE LA ORDEN", "PRECIO PROVEEDOR X CANTIDAD", "PRECIO FLETE", "CANTIDAD", "PRODUCTO", "FECHA_iso"]

with tab1:
    if df_ordenes.empty:
        st.info("Tu memoria está vacía todavía. Ve a ⚙️ Panel de Control para subir tus primeros reportes.")
    elif not verificar_columnas(df_ordenes, COLS_ORDENES_REQUERIDAS, HOJAS["ordenes"]):
        pass
    else:
        ini, fin = selector_rango_fechas("dash")
        df_rango = df_ordenes[(df_ordenes["FECHA_dt"].dt.date >= ini) & (df_ordenes["FECHA_dt"].dt.date <= fin)]
        df_ads_rango = df_ads[(pd.to_datetime(df_ads["FECHA_iso"]).dt.date >= ini) & (pd.to_datetime(df_ads["FECHA_iso"]).dt.date <= fin)] if not df_ads.empty else df_ads

        entregadas_df = df_rango[df_rango["ESTATUS"].isin(ESTATUS_ENTREGADO)]
        ingreso = convertir_columna(entregadas_df, "FECHA_iso", "TOTAL DE LA ORDEN", DD, DO, df_tc, TC_DROPI_MANUAL).sum()
        costo_proveedor = convertir_columna(df_rango, "FECHA_iso", "PRECIO PROVEEDOR X CANTIDAD", DD, DO, df_tc, TC_DROPI_MANUAL).sum()
        costo_flete = convertir_columna(df_rango, "FECHA_iso", "PRECIO FLETE", DD, DO, df_tc, TC_DROPI_MANUAL).sum()
        dev_df = df_rango[df_rango["ESTATUS"].isin(ESTATUS_DEVOLUCION)]
        costo_flete_dev = convertir_columna(dev_df, "FECHA_iso", "COSTO DEVOLUCION FLETE", DD, DO, df_tc, TC_DROPI_MANUAL).sum()
        gasto_ads = convertir_columna(df_ads_rango, "FECHA_iso", "_gasto", DA, DO, df_tc, TC_ADS_MANUAL).sum() if not df_ads_rango.empty else 0.0

        total_gastos = costo_proveedor + costo_flete + costo_flete_dev + gasto_ads
        utilidad_neta = ingreso - total_gastos
        articulos_vendidos = entregadas_df["CANTIDAD"].sum()

        st.subheader(f"Resumen ({ini.strftime('%d/%m/%Y')} – {fin.strftime('%d/%m/%Y')}) en {DO}")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Órdenes generadas", len(df_rango))
        c2.metric("Artículos vendidos", int(articulos_vendidos))
        c3.metric("Ganancia de ventas", f"${ingreso:,.2f}")
        c4.metric("Gastos totales", f"${total_gastos:,.2f}")
        c5.metric("Utilidad neta", f"${utilidad_neta:,.2f}", delta=f"{(utilidad_neta/ingreso*100 if ingreso else 0):.1f}% margen")

        st.divider()
        st.subheader("Detalle de órdenes y artículos")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Órdenes generadas", len(df_rango))
        c2.metric("Entregadas", len(entregadas_df))
        c3.metric("Devoluciones", len(dev_df))
        c4.metric("Cancelaciones", len(df_rango[df_rango["ESTATUS"].isin(ESTATUS_CANCELADO)]))

        st.subheader("Gráfica dinámica (clic en la leyenda para activar/desactivar series)")
        metricas_disp = st.multiselect(
            "Datos a mostrar", ["Órdenes generadas", "Entregadas", "Devoluciones", "Cancelaciones"],
            default=["Órdenes generadas", "Entregadas", "Devoluciones", "Cancelaciones"], key="dash_metricas"
        )
        agg = df_rango.groupby(df_rango["FECHA_dt"].dt.date).agg(
            **{
                "Órdenes generadas": ("ID", "count"),
                "Entregadas": ("ESTATUS", lambda s: s.isin(ESTATUS_ENTREGADO).sum()),
                "Devoluciones": ("ESTATUS", lambda s: s.isin(ESTATUS_DEVOLUCION).sum()),
                "Cancelaciones": ("ESTATUS", lambda s: s.isin(ESTATUS_CANCELADO).sum()),
            }
        ).reset_index().rename(columns={"FECHA_dt": "Fecha"})
        if metricas_disp:
            largo = agg.melt(id_vars="Fecha", value_vars=metricas_disp, var_name="Métrica", value_name="Cantidad")
            fig = px.line(largo, x="Fecha", y="Cantidad", color="Métrica", markers=True)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Selecciona al menos un dato para graficar.")


# ==============================================================================
# TAB 2 — ANUNCIOS
# ==============================================================================
COLS_ADS_REQUERIDAS = ["Nombre de la campaña", "Nombre del conjunto de anuncios", "Nombre del anuncio", "FECHA_iso"]

with tab2:
    if df_ads.empty:
        st.info("Aún no subes reportes de Meta Ads. Ve a ⚙️ Panel de Control.")
    elif not verificar_columnas(df_ads, COLS_ADS_REQUERIDAS, HOJAS["ads"]):
        pass
    else:
        ini, fin = selector_rango_fechas("ads")
        df_ads_rango = df_ads[(pd.to_datetime(df_ads["FECHA_iso"]).dt.date >= ini) & (pd.to_datetime(df_ads["FECHA_iso"]).dt.date <= fin)].copy()

        campanas = sorted(df_ads_rango["Nombre de la campaña"].dropna().unique())
        seleccion = st.multiselect("Filtrar por campaña", campanas, default=campanas[:3] if len(campanas) > 3 else campanas)
        df_vista = df_ads_rango[df_ads_rango["Nombre de la campaña"].isin(seleccion)] if seleccion else df_ads_rango

        st.subheader("Comportamiento diario (gasto, CPA, alcance)")
        diario = df_vista.groupby("FECHA_iso").agg(Gasto=("_gasto", "sum"), CPA=("_cpa_meta", "mean"), Alcance=("_alcance", "sum")).reset_index()
        fig1 = go.Figure()
        fig1.add_trace(go.Scatter(x=diario["FECHA_iso"], y=diario["Gasto"], name="Gasto", mode="lines+markers"))
        fig1.add_trace(go.Scatter(x=diario["FECHA_iso"], y=diario["CPA"], name="CPA (Meta)", mode="lines+markers", yaxis="y2"))
        fig1.update_layout(
            yaxis=dict(title="Gasto"), yaxis2=dict(title="CPA", overlaying="y", side="right"),
            legend=dict(orientation="h"),
        )
        st.plotly_chart(fig1, use_container_width=True)
        fig1b = px.bar(diario, x="FECHA_iso", y="Alcance", title="Alcance diario")
        st.plotly_chart(fig1b, use_container_width=True)

        st.subheader("Órdenes del día vs. gasto en publicidad")
        ord_dia = df_ordenes.groupby(df_ordenes["FECHA_dt"].dt.strftime("%Y-%m-%d")).size().reset_index(name="Órdenes")
        ord_dia.columns = ["FECHA_iso", "Órdenes"]
        comparativo = diario[["FECHA_iso", "Gasto"]].merge(ord_dia, on="FECHA_iso", how="outer").fillna(0).sort_values("FECHA_iso")
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(x=comparativo["FECHA_iso"], y=comparativo["Órdenes"], name="Órdenes"))
        fig2.add_trace(go.Scatter(x=comparativo["FECHA_iso"], y=comparativo["Gasto"], name="Gasto Ads", yaxis="y2", mode="lines+markers"))
        fig2.update_layout(yaxis=dict(title="Órdenes"), yaxis2=dict(title="Gasto", overlaying="y", side="right"))
        st.plotly_chart(fig2, use_container_width=True)

        st.divider()
        st.subheader("Rentabilidad real por producto (cruce Ads × Dropi)")
        st.caption(
            "El CPA que reporta Meta suele subestimar (solo cuenta conversaciones/pixel). "
            "Aquí el CPA REAL se calcula dividiendo el gasto de los anuncios vinculados entre las "
            "órdenes reales de Dropi para ese producto en el mismo periodo."
        )
        if df_vinc.empty:
            st.warning("Todavía no vinculas anuncios a productos. Hazlo en ⚙️ Panel de Control para ver esta tabla.")
        else:
            filas_prod = []
            vinc_rango = df_vinc.merge(
                df_ads_rango, left_on=["Campaña", "Conjunto_Anuncios", "Anuncio"],
                right_on=["Nombre de la campaña", "Nombre del conjunto de anuncios", "Nombre del anuncio"], how="inner"
            )
            for producto, sub in vinc_rango.groupby("Producto"):
                gasto_prod = sub["_gasto"].sum()
                ordenes_prod = df_ordenes[
                    (df_ordenes["PRODUCTO"] == producto)
                    & (pd.to_datetime(df_ordenes["FECHA_iso"]).dt.date >= ini)
                    & (pd.to_datetime(df_ordenes["FECHA_iso"]).dt.date <= fin)
                ]
                n_ordenes = len(ordenes_prod)
                cpa_real = gasto_prod / n_ordenes if n_ordenes else np.nan
                filas_prod.append({
                    "Producto": producto, "Gasto Ads": round(gasto_prod, 2), "Órdenes reales": n_ordenes,
                    "CPA Real": round(cpa_real, 2) if pd.notna(cpa_real) else None,
                    "Alcance total": int(sub["_alcance"].sum()),
                })
            tabla_prod = pd.DataFrame(filas_prod).sort_values("Gasto Ads", ascending=False)
            st.dataframe(tabla_prod, use_container_width=True, hide_index=True)

        st.subheader("Rendimiento y ranking por anuncio")
        filas = []
        for (camp, conj, anun), sub in df_ads_rango.groupby(["Nombre de la campaña", "Nombre del conjunto de anuncios", "Nombre del anuncio"]):
            gasto = sub["_gasto"].sum()
            cpa_prom = sub["_cpa_meta"].mean()
            alcance = sub["_alcance"].sum()
            hist = df_ads[df_ads["Nombre del anuncio"] == anun]["_cpa_meta"]
            semaforo, texto = calcular_semaforo(
                cpa_prom, hist, config["Usar_Semaforo_Manual"], config["CPA_Bajo_Manual"], config["CPA_Alto_Manual"]
            )
            dia_alto = sub.loc[sub["_gasto"].idxmax(), "FECHA_iso"] if not sub.empty and sub["_gasto"].notna().any() else "-"
            dia_bajo = sub.loc[sub["_gasto"].idxmin(), "FECHA_iso"] if not sub.empty and sub["_gasto"].notna().any() else "-"
            if semaforo == "🟢":
                recomendacion = "🚀 Escalar presupuesto"
            elif semaforo == "🟡":
                recomendacion = "⚖️ Mantener / ajustar ±10%"
            elif semaforo == "🔴":
                recomendacion = "🛑 Reducir o apagar"
            else:
                recomendacion = "ℹ️ Falta historial"
            filas.append({
                "Campaña": camp, "Conjunto": conj, "Anuncio": anun, "Gasto total": round(gasto, 2),
                "CPA prom. (Meta)": round(cpa_prom, 2) if pd.notna(cpa_prom) else None,
                "Alcance": int(alcance), "Día costo más alto": dia_alto, "Día costo más bajo": dia_bajo,
                "Semáforo": semaforo, "Detalle": texto, "Recomendación": recomendacion,
            })
        tabla_anuncios = pd.DataFrame(filas).sort_values("Gasto total", ascending=False)
        st.dataframe(tabla_anuncios, use_container_width=True, hide_index=True)

        st.subheader("🏆 Ranking de mejores anuncios (por CPA promedio, de menor a mayor)")
        ranking = tabla_anuncios.dropna(subset=["CPA prom. (Meta)"]).sort_values("CPA prom. (Meta)").head(10)
        st.dataframe(ranking[["Anuncio", "Campaña", "CPA prom. (Meta)", "Gasto total", "Semáforo"]], use_container_width=True, hide_index=True)


# ==============================================================================
# TAB 3 — CALCULADORA DE PRECIOS
# ==============================================================================
with tab3:
    st.header(f"Calculadora de precios ({DO})")
    st.caption("Todos los montos se muestran ya convertidos a tu divisa origen.")

    col_a, col_b = st.columns(2)
    with col_a:
        costo_prov = st.number_input(f"Costo proveedor ({DD})", min_value=0.0, value=45.0)
        costo_flete_in = st.number_input(f"Costo de flete ({DD})", min_value=0.0, value=50.0)
        flete_devolucion = st.number_input(f"Flete por devolución ({DD})", min_value=0.0, value=35.0)
        cpa_esperado = st.number_input(f"Costo por adquisición / gasto de publicidad por venta ({DA})", min_value=0.0, value=50.0)
    with col_b:
        tasa_dev_hist, tasa_canc_hist = tasas_devolucion_cancelacion(df_ordenes, config) if not df_ordenes.empty else (0.15, 0.05)
        tasa_dev = st.slider("Tasa de devoluciones (%)", 0, 90, int(tasa_dev_hist * 100)) / 100
        tasa_canc = st.slider("Tasa de cancelaciones (%)", 0, 90, int(tasa_canc_hist * 100)) / 100
        margen_deseado = st.slider("Margen de utilidad deseado sobre el precio (%)", 0, 80, 20) / 100

    tasa_conv_dropi = construir_mapa_tasas(df_tc, DD, DO)
    tasa_conv_dropi_hoy = list(tasa_conv_dropi.values())[-1] if tasa_conv_dropi else (TC_DROPI_MANUAL or 1.0)
    tasa_conv_ads = construir_mapa_tasas(df_tc, DA, DO)
    tasa_conv_ads_hoy = list(tasa_conv_ads.values())[-1] if tasa_conv_ads else (TC_ADS_MANUAL or 1.0)

    costo_prov_o = costo_prov * tasa_conv_dropi_hoy
    costo_flete_o = costo_flete_in * tasa_conv_dropi_hoy
    flete_dev_o = flete_devolucion * tasa_conv_dropi_hoy * tasa_dev
    cpa_o = cpa_esperado * tasa_conv_ads_hoy

    costo_total_fijo = costo_prov_o + costo_flete_o + flete_dev_o + cpa_o
    precio_minimo = costo_total_fijo / (1 - tasa_dev - tasa_canc) if (1 - tasa_dev - tasa_canc) > 0 else np.nan
    precio_recomendado = precio_minimo / (1 - margen_deseado) if margen_deseado < 1 else np.nan
    utilidad_generada = precio_recomendado - (costo_total_fijo / (1 - tasa_dev - tasa_canc)) if pd.notna(precio_recomendado) else np.nan

    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric(f"Costo total (por venta efectiva) — {DO}", f"${costo_total_fijo/(1-tasa_dev-tasa_canc):,.2f}" if (1-tasa_dev-tasa_canc) > 0 else "N/A")
    c2.metric(f"Precio mínimo (punto de equilibrio) — {DO}", f"${precio_minimo:,.2f}" if pd.notna(precio_minimo) else "N/A")
    c3.metric(f"Precio recomendado (con {int(margen_deseado*100)}% margen) — {DO}", f"${precio_recomendado:,.2f}" if pd.notna(precio_recomendado) else "N/A")
    st.success(f"Utilidad generada por venta al precio recomendado: **${utilidad_generada:,.2f} {DO}**" if pd.notna(utilidad_generada) else "Ajusta los parámetros.")
    if DD != DO:
        st.caption(f"Precio recomendado equivalente: ${precio_recomendado/tasa_conv_dropi_hoy:,.2f} {DD}" if pd.notna(precio_recomendado) and tasa_conv_dropi_hoy else "")


# ==============================================================================
# TAB 4 — PANEL DE CONTROL
# ==============================================================================
with tab4:
    st.header("⚙️ Panel de Control")

    st.subheader("1. Divisas")
    c1, c2, c3 = st.columns(3)
    nueva_do = c1.selectbox("Divisa ORIGEN (utilidad/retiro)", DIVISAS, index=DIVISAS.index(DO) if DO in DIVISAS else 0)
    nueva_da = c2.selectbox("Divisa en que pagas Ads", DIVISAS, index=DIVISAS.index(DA) if DA in DIVISAS else 0)
    nueva_dd = c3.selectbox("Divisa en que reporta Dropi", DIVISAS, index=DIVISAS.index(DD) if DD in DIVISAS else 0)
    if divisa_ads_detectada and divisa_ads_detectada != nueva_da:
        st.warning(f"Tu último reporte de Ads venía en **{divisa_ads_detectada}**, pero tienes configurado **{nueva_da}**. Verifica cuál es correcto.")

    st.subheader("2. Tipo de cambio")
    usar_manual_tc = st.checkbox("Sobrescribir el tipo de cambio automático con uno manual", value=config["Usar_TC_Manual"])
    c1, c2 = st.columns(2)
    tc_ads_manual = c1.number_input(f"1 {nueva_da} = X {nueva_do}", min_value=0.0, value=float(config["TC_Ads_Manual"]), format="%.4f", disabled=not usar_manual_tc)
    tc_dropi_manual = c2.number_input(f"1 {nueva_dd} = X {nueva_do}", min_value=0.0, value=float(config["TC_Dropi_Manual"]), format="%.4f", disabled=not usar_manual_tc)
    st.caption("Con 'automático' el sistema consulta el tipo de cambio del día una vez al día y lo guarda en la hoja Tipo_Cambio_Historico, para que tus cálculos pasados no cambien.")

    st.subheader("3. Semáforo de rendimiento de anuncios")
    usar_manual_sem = st.checkbox("Usar umbrales fijos en vez de compararlo con mi histórico", value=config["Usar_Semaforo_Manual"])
    c1, c2 = st.columns(2)
    cpa_bajo = c1.number_input(f"CPA bueno (verde) si es menor o igual a ({nueva_do})", min_value=0.0, value=float(config["CPA_Bajo_Manual"]), disabled=not usar_manual_sem)
    cpa_alto = c2.number_input(f"CPA malo (rojo) si es mayor a ({nueva_do})", min_value=0.0, value=float(config["CPA_Alto_Manual"]), disabled=not usar_manual_sem)

    st.subheader("4. Tasas de devolución y cancelación")
    usar_manual_tasas = st.checkbox("Usar tasas manuales en vez de calcularlas del histórico", value=config["Usar_Tasas_Manual"])
    c1, c2 = st.columns(2)
    tasa_dev_manual = c1.slider("Tasa de devoluciones manual (%)", 0, 90, int(config["Tasa_Devolucion_Manual"] * 100), disabled=not usar_manual_tasas) / 100
    tasa_canc_manual = c2.slider("Tasa de cancelaciones manual (%)", 0, 90, int(config["Tasa_Cancelacion_Manual"] * 100), disabled=not usar_manual_tasas) / 100

    if st.button("💾 Guardar configuración"):
        guardar_config(usuario_activo, {
            "Divisa_Origen": nueva_do, "Divisa_Ads": nueva_da, "Divisa_Dropi": nueva_dd,
            "Usar_TC_Manual": usar_manual_tc, "TC_Ads_Manual": tc_ads_manual, "TC_Dropi_Manual": tc_dropi_manual,
            "Usar_Semaforo_Manual": usar_manual_sem, "CPA_Bajo_Manual": cpa_bajo, "CPA_Alto_Manual": cpa_alto,
            "Usar_Tasas_Manual": usar_manual_tasas, "Tasa_Devolucion_Manual": tasa_dev_manual, "Tasa_Cancelacion_Manual": tasa_canc_manual,
        })
        if usar_manual_tc:
            registrar_tasa_manual(nueva_da, nueva_do, tc_ads_manual)
            registrar_tasa_manual(nueva_dd, nueva_do, tc_dropi_manual)
        st.success("Configuración guardada.")
        st.rerun()

    st.divider()
    st.subheader("5. Subir reportes diarios")
    c1, c2 = st.columns(2)
    archivo_ordenes = c1.file_uploader("Reporte de órdenes de Dropi (.xlsx)", type=["xlsx"])
    archivo_ads = c2.file_uploader("Reporte de Anuncios de Meta — nivel Anuncio, con columna 'Campaña' (.csv)", type=["csv"])

    if st.button("💾 Guardar en memoria"):
        if archivo_ordenes is None and archivo_ads is None:
            st.warning("Sube al menos un archivo.")
        else:
            if archivo_ordenes is not None:
                df_new_ordenes = pd.read_excel(archivo_ordenes)
                n = guardar_ordenes(usuario_activo, df_new_ordenes)
                st.success(f"Órdenes: {n} filas nuevas agregadas (se ignoraron duplicados por ID).")
            if archivo_ads is not None:
                df_new_ads = pd.read_csv(archivo_ads)
                if "Nombre de la campaña" not in df_new_ads.columns:
                    st.error("Este CSV no trae la columna 'Nombre de la campaña'. Agrégala al exportar desde Ads Manager.")
                else:
                    n = guardar_ads(usuario_activo, df_new_ads)
                    st.success(f"Anuncios: {n} filas nuevas agregadas (se ignoraron duplicados).")
            st.rerun()

    st.divider()
    st.subheader("6. Vincular anuncios a productos")
    if df_ads.empty:
        st.info("Sube primero un reporte de anuncios.")
    elif not verificar_columnas(df_ads, COLS_ADS_REQUERIDAS, HOJAS["ads"]):
        pass
    else:
        combinaciones = df_ads[["Nombre de la campaña", "Nombre del conjunto de anuncios", "Nombre del anuncio"]].drop_duplicates()
        ya_vinculadas = set(zip(df_vinc.get("Campaña", []), df_vinc.get("Conjunto_Anuncios", []), df_vinc.get("Anuncio", [])))
        combinaciones["_clave"] = list(zip(combinaciones["Nombre de la campaña"], combinaciones["Nombre del conjunto de anuncios"], combinaciones["Nombre del anuncio"]))
        pendientes = combinaciones[~combinaciones["_clave"].isin(ya_vinculadas)]
        productos_disponibles = sorted(df_ordenes["PRODUCTO"].dropna().unique()) if not df_ordenes.empty else []

        st.caption(f"{len(pendientes)} anuncios sin vincular de {len(combinaciones)} totales.")
        if not pendientes.empty and productos_disponibles:
            fila = pendientes.iloc[0]
            st.write(f"**Campaña:** {fila['Nombre de la campaña']} · **Conjunto:** {fila['Nombre del conjunto de anuncios']} · **Anuncio:** {fila['Nombre del anuncio']}")
            producto_sel = st.selectbox("¿A qué producto pertenece?", productos_disponibles, key="vinc_producto")
            if st.button("🔗 Vincular"):
                guardar_vinculo(usuario_activo, fila["Nombre de la campaña"], fila["Nombre del conjunto de anuncios"], fila["Nombre del anuncio"], producto_sel)
                st.success("Vinculado.")
                st.rerun()
        elif pendientes.empty:
            st.success("Todos tus anuncios están vinculados a un producto. 🎉")

        if not df_vinc.empty:
            with st.expander("Ver / editar vínculos existentes"):
                st.dataframe(df_vinc[["Campaña", "Conjunto_Anuncios", "Anuncio", "Producto"]], use_container_width=True, hide_index=True)


# ==============================================================================
# TAB 5 — RENTABILIDAD
# ==============================================================================
with tab5:
    if df_ordenes.empty:
        st.info("Sin datos todavía.")
    elif not verificar_columnas(df_ordenes, COLS_ORDENES_REQUERIDAS, HOJAS["ordenes"]):
        pass
    else:
        ini, fin = selector_rango_fechas("rent")
        df_rango = df_ordenes[(df_ordenes["FECHA_dt"].dt.date >= ini) & (df_ordenes["FECHA_dt"].dt.date <= fin)]
        df_ads_rango = df_ads[(pd.to_datetime(df_ads["FECHA_iso"]).dt.date >= ini) & (pd.to_datetime(df_ads["FECHA_iso"]).dt.date <= fin)] if not df_ads.empty else pd.DataFrame()

        filas = []
        for fecha, sub in df_rango.groupby(df_rango["FECHA_dt"].dt.date):
            total = len(sub)
            devoluciones = sub["ESTATUS"].isin(ESTATUS_DEVOLUCION).sum()
            cancelaciones = sub["ESTATUS"].isin(ESTATUS_CANCELADO).sum()
            costo_prov = convertir_columna(sub, "FECHA_iso", "PRECIO PROVEEDOR X CANTIDAD", DD, DO, df_tc, TC_DROPI_MANUAL).sum()
            ingreso = convertir_columna(sub[sub["ESTATUS"].isin(ESTATUS_ENTREGADO)], "FECHA_iso", "TOTAL DE LA ORDEN", DD, DO, df_tc, TC_DROPI_MANUAL).sum()
            gasto_ads_dia = 0.0
            if not df_ads_rango.empty:
                sub_ads = df_ads_rango[df_ads_rango["FECHA_iso"] == fecha.strftime("%Y-%m-%d")]
                gasto_ads_dia = convertir_columna(sub_ads, "FECHA_iso", "_gasto", DA, DO, df_tc, TC_ADS_MANUAL).sum()
            utilidad_dia = ingreso - costo_prov - gasto_ads_dia
            filas.append({
                "Fecha": fecha, "Órdenes": total, "Devoluciones": devoluciones, "Cancelaciones": cancelaciones,
                "Tasa Dev.": f"{devoluciones/total*100:.1f}%" if total else "0%",
                "Tasa Canc.": f"{cancelaciones/total*100:.1f}%" if total else "0%",
                f"Costo proveedor ({DO})": round(costo_prov, 2), f"Gasto Ads ({DO})": round(gasto_ads_dia, 2),
                f"Utilidad del día ({DO})": round(utilidad_dia, 2), "Rentable": "🟢" if utilidad_dia > 0 else "🔴",
            })
        tabla_rent = pd.DataFrame(filas).sort_values("Fecha", ascending=False)

        st.subheader(f"Tabla de rentabilidad diaria ({DO})")
        st.dataframe(tabla_rent, use_container_width=True, hide_index=True)
        dias_rentables = (tabla_rent["Rentable"] == "🟢").sum()
        st.caption(f"{dias_rentables} de {len(tabla_rent)} días fueron rentables en el rango seleccionado.")


# ==============================================================================
# TAB 6 — RECOMENDACIONES
# ==============================================================================
with tab6:
    st.header("🧭 Recomendaciones para la próxima semana")
    if df_ads.empty or df_ordenes.empty:
        st.info("Necesito al menos historial de órdenes y de anuncios para generar recomendaciones.")
    elif not verificar_columnas(df_ads, COLS_ADS_REQUERIDAS, HOJAS["ads"]) or not verificar_columnas(df_ordenes, COLS_ORDENES_REQUERIDAS, HOJAS["ordenes"]):
        pass
    else:
        dias_analisis = st.slider("Días de historial a analizar", 7, 60, 14)
        desde = date.today() - timedelta(days=dias_analisis)
        df_ads_periodo = df_ads[pd.to_datetime(df_ads["FECHA_iso"]).dt.date >= desde]

        st.subheader("Presupuestos recomendados por anuncio")
        filas = []
        for (camp, conj, anun), sub in df_ads_periodo.groupby(["Nombre de la campaña", "Nombre del conjunto de anuncios", "Nombre del anuncio"]):
            gasto_diario_prom = sub.groupby("FECHA_iso")["_gasto"].sum().mean()
            cpa_prom = sub["_cpa_meta"].mean()
            hist = df_ads[df_ads["Nombre del anuncio"] == anun]["_cpa_meta"]
            semaforo, _ = calcular_semaforo(cpa_prom, hist, config["Usar_Semaforo_Manual"], config["CPA_Bajo_Manual"], config["CPA_Alto_Manual"])
            if semaforo == "🟢":
                accion, factor = "🚀 Subir presupuesto", 1.3
            elif semaforo == "🟡":
                accion, factor = "⚖️ Ajustar ligeramente", 1.0
            elif semaforo == "🔴":
                accion, factor = "🛑 Bajar presupuesto / apagar", 0.5
            else:
                accion, factor = "🆕 Dejar correr (poco historial)", 1.0
            filas.append({
                "Campaña": camp, "Conjunto": conj, "Anuncio": anun,
                f"Gasto diario actual ({DA})": round(gasto_diario_prom, 2) if pd.notna(gasto_diario_prom) else 0,
                "Semáforo": semaforo, "Acción sugerida": accion,
                f"Presupuesto diario sugerido ({DA})": round((gasto_diario_prom or 0) * factor, 2),
            })
        tabla_presup = pd.DataFrame(filas).sort_values(f"Gasto diario actual ({DA})", ascending=False)
        st.dataframe(tabla_presup, use_container_width=True, hide_index=True)
        st.caption("Sugerencia basada en heurística de reglas (no garantiza resultados): escalar +30% en verde, mantener en amarillo, reducir 50% en rojo.")

        st.divider()
        st.subheader("Recomendaciones por producto")
        if df_vinc.empty:
            st.info("Vincula tus anuncios a productos en ⚙️ Panel de Control para ver esta tabla.")
        else:
            filas_p = []
            vinc_periodo = df_vinc.merge(
                df_ads_periodo, left_on=["Campaña", "Conjunto_Anuncios", "Anuncio"],
                right_on=["Nombre de la campaña", "Nombre del conjunto de anuncios", "Nombre del anuncio"], how="inner"
            )
            df_ord_periodo = df_ordenes[pd.to_datetime(df_ordenes["FECHA_iso"]).dt.date >= desde]
            for producto, sub in vinc_periodo.groupby("Producto"):
                gasto_prod = convertir_columna(sub, "FECHA_iso", "_gasto", DA, DO, df_tc, TC_ADS_MANUAL).sum()
                ord_prod = df_ord_periodo[df_ord_periodo["PRODUCTO"] == producto]
                entregadas_prod = ord_prod[ord_prod["ESTATUS"].isin(ESTATUS_ENTREGADO)]
                ingreso_prod = convertir_columna(entregadas_prod, "FECHA_iso", "TOTAL DE LA ORDEN", DD, DO, df_tc, TC_DROPI_MANUAL).sum()
                costo_prod = convertir_columna(ord_prod, "FECHA_iso", "PRECIO PROVEEDOR X CANTIDAD", DD, DO, df_tc, TC_DROPI_MANUAL).sum()
                utilidad_prod = ingreso_prod - costo_prod - gasto_prod
                margen = utilidad_prod / ingreso_prod if ingreso_prod else 0
                if margen > 0.15:
                    rec = "📈 Escalar"
                elif margen > 0:
                    rec = "⚖️ Ajustar presupuesto/precio"
                else:
                    rec = "🛑 Dejar de vender o subir precio"
                filas_p.append({
                    "Producto": producto, f"Ingreso ({DO})": round(ingreso_prod, 2), f"Costo ({DO})": round(costo_prod + gasto_prod, 2),
                    f"Utilidad ({DO})": round(utilidad_prod, 2), "Margen": f"{margen*100:.1f}%", "Recomendación": rec,
                })
            tabla_prod_rec = pd.DataFrame(filas_p).sort_values(f"Utilidad ({DO})", ascending=False)
            st.dataframe(tabla_prod_rec, use_container_width=True, hide_index=True)

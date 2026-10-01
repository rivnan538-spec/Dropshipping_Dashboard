"""
================================================================================
ASISTENTE DE RENTABILIDAD PARA DROPSHIPPING  (versión corregida)
Streamlit + Google Sheets — multiusuario, multidivisa, con memoria histórica.
Fuentes: reporte de órdenes de Dropi (.xlsx) y reporte de Meta Ads (.csv) con
desglose por día, a nivel Campaña o a nivel Anuncio.

CAMBIOS PRINCIPALES vs. la versión anterior
 1. Ya no se pierde la memoria si Google Sheets falla al leer (antes se
    sobrescribía la hoja completa con solo las filas nuevas).
 2. Las órdenes existentes se ACTUALIZAN (estatus nuevo) en vez de ignorarse.
 3. Costos calculados solo sobre lo que realmente se cobra/paga:
    proveedor + flete solo en ENTREGADAS; en DEVOLUCIÓN se pierde el flete.
    Los importes a nivel orden (total, flete) se cuentan una vez por ID.
 4. Acepta el CSV de Meta a nivel Campaña (antes exigía nivel Anuncio).
 5. Métrica de Meta = costo por resultado (conversaciones) calculado
    ponderado; ya no depende de "Costo por compra" (venía vacío).
 6. Fechas con zona horaria local (antes usaba la hora del servidor, UTC).
 7. Tipo de cambio manual con validación de dirección (GTQ→MXN ≈ 2.3, no 0.43).
 8. Tasas de devolución/cancelación sobre la base correcta; calculadora
    corregida; tabla diaria incluye días con gasto y sin órdenes.
 9. Vinculación anuncio→producto con uno o varios productos.
10. Contraseñas: soporta hash sha256 y comparación en tiempo constante.
================================================================================
"""

import bisect
import hashlib
import hmac
import re
import unicodedata
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

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
ZONA_HORARIA = "America/Mexico_City"  # Guatemala y CDMX comparten UTC-6 todo el año

DIVISAS = ["MXN", "GTQ", "COP", "USD", "EUR", "GBP", "CAD", "PEN", "CLP", "ARS", "BOB", "HNL", "PAB"]

# Los estatus se normalizan (mayúsculas, sin acentos) al cargar las órdenes.
ESTATUS_ENTREGADO = {"ENTREGADO"}
ESTATUS_DEVOLUCION = {"DEVOLUCION"}
ESTATUS_CANCELADO = {"CANCELADO"}
ESTATUS_FINALES = ESTATUS_ENTREGADO | ESTATUS_DEVOLUCION | ESTATUS_CANCELADO
# Todo lo demás (PENDIENTE, GUIA_GENERADA, EN RUTA, INCIDENCIA VALIDADA...) = en tránsito.

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
    "Devolucion_Cobra_Flete",
]
COLS_VINCULACION = ["Usuario", "Campaña", "Conjunto_Anuncios", "Anuncio", "Producto"]
COLS_TIPO_CAMBIO = ["Fecha", "Divisa_Base", "Divisa_Destino", "Tasa", "Fuente"]
COLS_CLAVE_ADS = ["Nombre de la campaña", "Nombre del conjunto de anuncios", "Nombre del anuncio"]
SEP_PRODUCTOS = " || "  # separador cuando un anuncio se vincula a varios productos

DEFAULT_CONFIG = {
    "Divisa_Origen": "MXN", "Divisa_Ads": "MXN", "Divisa_Dropi": "GTQ",
    # 1 GTQ ≈ 2.3 MXN  (¡no 0.43! eso es MXN→GTQ)
    "Usar_TC_Manual": False, "TC_Ads_Manual": 1.0, "TC_Dropi_Manual": 2.3,
    "Usar_Semaforo_Manual": False, "CPA_Bajo_Manual": 80.0, "CPA_Alto_Manual": 150.0,
    "Usar_Tasas_Manual": False, "Tasa_Devolucion_Manual": 0.15, "Tasa_Cancelacion_Manual": 0.05,
    "Devolucion_Cobra_Flete": True,
}

API_TIPO_CAMBIO = "https://open.er-api.com/v6/latest/{base}"
RANGOS_FECHA = ["Hoy", "Últimos 7 días", "Últimos 30 días", "Personalizado"]
MIN_RESULTADOS_RANKING = 10

conn = st.connection("gsheets", type=GSheetsConnection)


# ==============================================================================
# FECHA / HORA LOCAL
# ==============================================================================
def ahora():
    return datetime.now(ZoneInfo(ZONA_HORARIA))


def hoy_local():
    return ahora().date()


# ==============================================================================
# UTILIDADES DE GOOGLE SHEETS
# ==============================================================================
def leer_hoja(nombre, ttl=60, estricto=False):
    """Lee una pestaña. Con estricto=True, un fallo de lectura LANZA error (obligatorio
    antes de cualquier escritura: si no, una lectura fallida parece 'hoja vacía' y
    se sobrescribe todo el historial)."""
    try:
        df = conn.read(worksheet=nombre, ttl=ttl)
    except Exception as e:
        if estricto:
            raise RuntimeError(f"No pude leer la pestaña '{nombre}' de Google Sheets ({e}). No se modificó nada.") from e
        return pd.DataFrame()
    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def escribir_hoja(nombre, df, filas_previas=None):
    """Escribe la hoja completa. Si se pasa filas_previas, se niega a escribir una tabla
    mucho más chica (protección contra pérdida accidental de historial)."""
    if filas_previas and len(df) < 0.8 * filas_previas:
        raise RuntimeError(
            f"Cancelé la escritura en '{nombre}': la tabla nueva ({len(df)} filas) es mucho más chica "
            f"que la actual ({filas_previas}). Revisa antes de continuar."
        )
    conn.update(worksheet=nombre, data=df)
    st.cache_data.clear()


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


def en_rango(df, col_fecha, ini, fin):
    """Filtra df por fecha (col_fecha en ISO o datetime) entre ini y fin, ambos incluidos."""
    if df is None or df.empty:
        return df
    f = _a_datetime(df[col_fecha])
    return df[(f >= pd.Timestamp(ini)) & (f < pd.Timestamp(fin) + pd.Timedelta(days=1))]


# ==============================================================================
# LOGIN
# ==============================================================================
def _contrasena_valida(ingresada, guardada):
    """Acepta contraseña en texto plano o 'sha256$<hash hexadecimal>' en la hoja Usuarios."""
    ingresada = str(ingresada).strip()
    guardada = str(guardada).strip()
    if guardada.lower().startswith("sha256$"):
        h = hashlib.sha256(ingresada.encode("utf-8")).hexdigest()
        return hmac.compare_digest(h, guardada[7:].strip().lower())
    if guardada.endswith(".0"):  # Sheets convierte 1234 en 1234.0
        guardada = guardada[:-2]
    return hmac.compare_digest(ingresada.encode("utf-8"), guardada.encode("utf-8"))


def login():
    st.title("🔒 Acceso al Asistente")
    try:
        df_usuarios = leer_hoja(HOJAS["usuarios"], ttl=0, estricto=True)
    except Exception as e:
        st.error(f"Error al conectar con Google Sheets: {e}")
        return False
    if df_usuarios.empty or not {"Usuario", "Contraseña"} <= set(df_usuarios.columns):
        st.error("La pestaña 'Usuarios' debe tener las columnas 'Usuario' y 'Contraseña'.")
        return False

    usuario = st.text_input("Usuario")
    contrasena = st.text_input("Contraseña", type="password")

    if st.button("Entrar"):
        usuario_limpio = usuario.strip()
        coincide = df_usuarios[df_usuarios["Usuario"].astype(str).str.strip() == usuario_limpio]
        if coincide.empty:
            st.error("Usuario o contraseña incorrectos.")
        elif _contrasena_valida(contrasena, coincide["Contraseña"].values[0]):
            st.session_state["logueado"] = True
            st.session_state["usuario_actual"] = usuario_limpio
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos.")
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
    for k in ["Usar_TC_Manual", "Usar_Semaforo_Manual", "Usar_Tasas_Manual", "Devolucion_Cobra_Flete"]:
        salida[k] = str(salida[k]).strip().lower() in ("true", "1", "verdadero", "sí", "si")
    for k in ["TC_Ads_Manual", "TC_Dropi_Manual", "CPA_Bajo_Manual", "CPA_Alto_Manual",
              "Tasa_Devolucion_Manual", "Tasa_Cancelacion_Manual"]:
        try:
            salida[k] = float(salida[k])
        except (ValueError, TypeError):
            salida[k] = DEFAULT_CONFIG[k]
    return salida


def guardar_config(usuario, nueva_config: dict):
    df = leer_hoja(HOJAS["config"], ttl=0, estricto=True)
    filas_previas = len(df)
    if df.empty:
        df = pd.DataFrame(columns=COLS_CONFIG)
    elif "Usuario" in df.columns:
        df = df[df["Usuario"].astype(str).str.strip() != usuario]
    fila = {c: nueva_config.get(c) for c in COLS_CONFIG}
    fila["Usuario"] = usuario
    df = pd.concat([df, pd.DataFrame([fila])], ignore_index=True)
    escribir_hoja(HOJAS["config"], df, filas_previas=filas_previas)


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


def _a_datetime(serie, **kw):
    """to_datetime tolerante a formatos mezclados ('2026-09-05' y '2026-09-05 00:00:00' en la misma columna)."""
    try:
        return pd.to_datetime(serie, format="mixed", errors="coerce", **kw)
    except (TypeError, ValueError):  # pandas antiguo sin format="mixed"
        return pd.to_datetime(serie, errors="coerce", **kw)


def _fecha_iso(serie):
    return _a_datetime(serie).dt.strftime("%Y-%m-%d")


def registrar_tasa_hoy(base, destino):
    """Obtiene (si hace falta) y guarda en Tipo_Cambio_Historico la tasa base->destino de HOY."""
    if base == destino:
        return 1.0
    hoy = hoy_local().isoformat()
    df_tc = leer_hoja(HOJAS["tipo_cambio"], ttl=60)
    if not df_tc.empty and {"Fecha", "Divisa_Base", "Divisa_Destino", "Tasa"} <= set(df_tc.columns):
        m = (_fecha_iso(df_tc["Fecha"]) == hoy) & (df_tc["Divisa_Base"] == base) & (df_tc["Divisa_Destino"] == destino)
        if m.any():
            return float(df_tc[m].iloc[-1]["Tasa"])
    rates = _obtener_tasas_api(base)
    if rates and destino in rates:
        tasa = float(rates[destino])
        df_tc2 = leer_hoja(HOJAS["tipo_cambio"], ttl=0, estricto=True)
        filas_previas = len(df_tc2)
        if df_tc2.empty:
            df_tc2 = pd.DataFrame(columns=COLS_TIPO_CAMBIO)
        nueva = {"Fecha": hoy, "Divisa_Base": base, "Divisa_Destino": destino, "Tasa": tasa, "Fuente": "API"}
        df_tc2 = pd.concat([df_tc2, pd.DataFrame([nueva])], ignore_index=True)
        escribir_hoja(HOJAS["tipo_cambio"], df_tc2, filas_previas=filas_previas)
        return tasa
    return None


def registrar_tasa_manual(base, destino, tasa):
    if base == destino:
        return
    hoy = hoy_local().isoformat()
    df_tc = leer_hoja(HOJAS["tipo_cambio"], ttl=0, estricto=True)
    if df_tc.empty:
        df_tc = pd.DataFrame(columns=COLS_TIPO_CAMBIO)
    elif {"Fecha", "Divisa_Base", "Divisa_Destino"} <= set(df_tc.columns):
        m = (_fecha_iso(df_tc["Fecha"]) == hoy) & (df_tc["Divisa_Base"] == base) & (df_tc["Divisa_Destino"] == destino)
        df_tc = df_tc[~m]
    nueva = {"Fecha": hoy, "Divisa_Base": base, "Divisa_Destino": destino, "Tasa": float(tasa), "Fuente": "Manual"}
    df_tc = pd.concat([df_tc, pd.DataFrame([nueva])], ignore_index=True)
    escribir_hoja(HOJAS["tipo_cambio"], df_tc)


def construir_mapa_tasas(df_tc_historico, base, destino):
    """{'YYYY-MM-DD': tasa} para base->destino (si hay varias en un día, gana la última)."""
    mapa = {}
    if df_tc_historico is None or df_tc_historico.empty:
        return mapa
    if not {"Fecha", "Divisa_Base", "Divisa_Destino", "Tasa"} <= set(df_tc_historico.columns):
        return mapa
    sub = df_tc_historico[(df_tc_historico["Divisa_Base"] == base) & (df_tc_historico["Divisa_Destino"] == destino)].copy()
    sub["_f"] = _fecha_iso(sub["Fecha"])
    sub["_t"] = pd.to_numeric(sub["Tasa"], errors="coerce")
    sub = sub.dropna(subset=["_f", "_t"])
    for f, t in zip(sub["_f"], sub["_t"]):
        mapa[f] = float(t)
    return mapa


def tasa_mas_reciente(mapa, fallback=None):
    """Tasa del día más reciente (por fecha, no por orden de filas en la hoja)."""
    if mapa:
        return mapa[max(mapa)]
    return fallback if fallback else 1.0


def convertir_columna(df, col_fecha, col_valor, base, destino, df_tc_historico, tasa_manual=None):
    """Convierte col_valor de `base` a `destino`.
    - Con tasa_manual: se usa esa tasa para todas las filas.
    - Si no: la tasa histórica del día de cada fila (o la más cercana anterior)."""
    if df is None or df.empty:
        return pd.Series(dtype=float)
    valores = pd.to_numeric(df[col_valor], errors="coerce").fillna(0)
    if base == destino:
        return valores
    if tasa_manual:
        return valores * float(tasa_manual)
    mapa = construir_mapa_tasas(df_tc_historico, base, destino)
    if not mapa:
        return valores * 1.0  # sin tasa disponible (la app avisa arriba)
    fechas = sorted(mapa)

    def tasa_de(f):
        if f in mapa:
            return mapa[f]
        i = bisect.bisect_right(fechas, f)
        return mapa[fechas[i - 1]] if i else mapa[fechas[0]]

    return valores * df[col_fecha].astype(str).map(tasa_de)


# ==============================================================================
# CARGA Y LIMPIEZA DE DATOS
# ==============================================================================
def normalizar_estatus(valor):
    if pd.isna(valor):
        return ""
    s = unicodedata.normalize("NFKD", str(valor)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s.upper().strip())


def parsear_fecha_dropi(serie):
    """Dropi exporta dd-mm-aaaa; Google Sheets puede devolver ISO o fechas reales. Se aceptan ambos."""
    dt = pd.to_datetime(serie, format="%d-%m-%Y", errors="coerce")
    resto = dt.isna() & serie.notna()
    if resto.any():
        dt.loc[resto] = _a_datetime(serie[resto])
    return dt


def detectar_columna_divisa(df, prefijo):
    """Busca una columna tipo 'Importe gastado (MXN)' y devuelve (columna, divisa)."""
    for c in df.columns:
        m = re.match(rf"{re.escape(prefijo)}\s*\(([A-Za-z]{{3}})\)", str(c))
        if m:
            return c, m.group(1).upper()
    return None, None


def _limpiar_id(serie):
    """Normaliza IDs: '1234567' y '1234567.0' (Sheets lo agrega a veces) son el mismo pedido."""
    return serie.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)


def cargar_ordenes(usuario):
    df = _filtrar_usuario(leer_hoja(HOJAS["ordenes"], ttl=120), usuario)
    if df.empty or "FECHA" not in df.columns:
        return df
    df = df.copy()
    df["FECHA_dt"] = parsear_fecha_dropi(df["FECHA"])
    df["FECHA_iso"] = df["FECHA_dt"].dt.strftime("%Y-%m-%d")
    if "ID" in df.columns:
        df["ID"] = _limpiar_id(df["ID"])
    if "ESTATUS" in df.columns:
        df["ESTATUS"] = df["ESTATUS"].map(normalizar_estatus)
    for c in ["TOTAL DE LA ORDEN", "PRECIO FLETE", "COSTO DEVOLUCION FLETE", "COMISION",
              "PRECIO PROVEEDOR X CANTIDAD", "CANTIDAD"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df.dropna(subset=["FECHA_dt"])


COLS_IMPORTE = ["_ingreso", "_costo_prov", "_flete", "_comision", "_perdida_dev", "_ganancia", "_potencial"]


def enriquecer_ordenes(df, cobra_flete_dev=True):
    """Calcula importes por fila (en divisa de Dropi) con la lógica real de Dropi:
      - Ingreso, proveedor, flete y comisión: SOLO órdenes ENTREGADAS.
      - Devolución: se pierde el flete (si cobra_flete_dev) + 'COSTO DEVOLUCION FLETE' si viniera.
      - Cancelada / en tránsito: sin costo ni ingreso todavía (en tránsito se estima aparte).
      - TOTAL, FLETE y COMISIÓN son de la ORDEN: se cuentan una sola vez por ID aunque la
        orden tenga varias líneas de producto. PROVEEDOR sí es por línea."""
    df = df.copy()
    idx = df.index

    def num(c):
        return pd.to_numeric(df[c], errors="coerce").fillna(0.0) if c in df.columns else pd.Series(0.0, index=idx)

    primera = ~df["ID"].duplicated()
    ent = df["ESTATUS"].isin(ESTATUS_ENTREGADO)
    dev = df["ESTATUS"].isin(ESTATUS_DEVOLUCION)
    transito = ~df["ESTATUS"].isin(ESTATUS_FINALES)
    total, flete, flete_dev = num("TOTAL DE LA ORDEN"), num("PRECIO FLETE"), num("COSTO DEVOLUCION FLETE")
    comision, prov = num("COMISION"), num("PRECIO PROVEEDOR X CANTIDAD")

    df["_ingreso"] = np.where(ent & primera, total, 0.0)
    df["_costo_prov"] = np.where(ent, prov, 0.0)
    df["_flete"] = np.where(ent & primera, flete, 0.0)
    df["_comision"] = np.where(ent & primera, comision, 0.0)
    perdida = (flete if cobra_flete_dev else 0.0) + flete_dev
    df["_perdida_dev"] = np.where(dev & primera, perdida, 0.0)
    df["_ganancia"] = df["_ingreso"] - df["_costo_prov"] - df["_flete"] - df["_comision"] - df["_perdida_dev"]
    df["_potencial"] = np.where(transito & primera, total - flete, 0.0) - np.where(transito, prov, 0.0)
    return df


def preparar_ordenes(df, cobra_flete_dev, DD, DO, df_tc, tc_manual):
    """Enriquece y agrega columnas *_o ya convertidas a la divisa origen."""
    df = enriquecer_ordenes(df, cobra_flete_dev)
    for c in COLS_IMPORTE:
        df[c + "_o"] = convertir_columna(df, "FECHA_iso", c, DD, DO, df_tc, tc_manual)
    return df


def cargar_ads(usuario):
    df = _filtrar_usuario(leer_hoja(HOJAS["ads"], ttl=120), usuario)
    if df.empty or "Inicio del informe" not in df.columns:
        return df, None
    df = df.copy()
    df["FECHA_iso"] = _fecha_iso(df["Inicio del informe"])

    def num(c):
        return pd.to_numeric(df[c], errors="coerce").fillna(0.0) if c in df.columns else pd.Series(0.0, index=df.index)

    col_gasto, divisa_detectada = detectar_columna_divisa(df, "Importe gastado")
    df["_gasto"] = pd.to_numeric(df[col_gasto], errors="coerce").fillna(0.0) if col_gasto else 0.0
    df["_alcance"] = num("Alcance")
    compras, resultados_meta = num("Compras"), num("Resultados")
    # Resultado principal = el 'Resultado' que Meta reporta para el objetivo de la campaña
    # (p. ej. conversaciones de mensajes iniciadas). Solo si viene vacío se usan las Compras.
    # (Las 'Compras' del píxel suelen ser muy pocas en campañas de mensajes y distorsionan el costo.)
    df["_resultados"] = np.where(resultados_meta > 0, resultados_meta, compras)
    df["_cpa_meta"] = np.where(df["_resultados"] > 0, df["_gasto"] / df["_resultados"].replace(0, np.nan), np.nan)
    return df.dropna(subset=["FECHA_iso"]), divisa_detectada


def normalizar_ads_csv(df):
    """Valida y normaliza el CSV de Meta. Devuelve (df, nivel) o lanza ValueError con un mensaje claro."""
    df = df.copy()
    if "Nombre de la campaña" not in df.columns:
        raise ValueError(f"El CSV no trae 'Nombre de la campaña'. Columnas encontradas: {list(df.columns)}")
    if "Inicio del informe" not in df.columns or "Fin del informe" not in df.columns:
        raise ValueError("El CSV no trae 'Inicio/Fin del informe'. Exporta con Desglose → Por tiempo → Día.")
    if detectar_columna_divisa(df, "Importe gastado")[0] is None:
        raise ValueError("El CSV no trae la columna 'Importe gastado (XXX)'.")
    df = df.dropna(subset=["Inicio del informe"])  # descarta filas de totales
    if (_fecha_iso(df["Inicio del informe"]) != _fecha_iso(df["Fin del informe"])).any():
        raise ValueError(
            "Hay filas que abarcan varios días. La app necesita una fila por día: en Ads Manager usa "
            "Desglose → Por tiempo → Día antes de exportar."
        )
    tiene_adset = "Nombre del conjunto de anuncios" in df.columns
    tiene_anuncio = "Nombre del anuncio" in df.columns
    if not tiene_adset:
        df["Nombre del conjunto de anuncios"] = "(todos los conjuntos)"
    if not tiene_anuncio:
        df["Nombre del anuncio"] = "(todos los anuncios)"
    nivel = "Anuncio" if (tiene_adset and tiene_anuncio) else "Campaña"
    df["Nivel_Reporte"] = nivel
    return df, nivel


# ------------------------------------------------------------------------------
# Escritura con "upsert" (lo más reciente gana) — nunca sobrescribe si la lectura falló
# ------------------------------------------------------------------------------
def _claves_orden(df):
    idx = df.index
    usr = df["Usuario"].astype(str).str.strip() if "Usuario" in df.columns else pd.Series("", index=idx)
    prod = df["PRODUCTO"].astype(str).str.strip() if "PRODUCTO" in df.columns else pd.Series("", index=idx)
    return usr + "||" + _limpiar_id(df["ID"]) + "||" + prod


def _claves_ads(df):
    k = df["Usuario"].astype(str).str.strip() if "Usuario" in df.columns else pd.Series("", index=df.index)
    for c in COLS_CLAVE_ADS:
        k = k + "||" + (df[c].astype(str).str.strip() if c in df.columns else "")
    return k + "||" + _fecha_iso(df["Inicio del informe"]).fillna("")


def guardar_ordenes(usuario, df_nuevo):
    """Inserta órdenes nuevas y ACTUALIZA las existentes (el estatus cambia cada día).
    Devuelve (nuevas, con_cambio_de_estatus)."""
    if "ID" not in df_nuevo.columns:
        raise ValueError(f"El archivo no trae la columna 'ID'. ¿Es el reporte de órdenes de Dropi? Columnas: {list(df_nuevo.columns)}")
    df_nuevo = df_nuevo.copy()
    df_nuevo["Usuario"] = usuario
    df_nuevo["Fecha_Carga"] = ahora().strftime("%Y-%m-%d %H:%M")
    k_new = _claves_orden(df_nuevo)
    df_nuevo = df_nuevo[~k_new.duplicated(keep="last")]
    k_new = _claves_orden(df_nuevo)

    df_actual = leer_hoja(HOJAS["ordenes"], ttl=0, estricto=True)
    if df_actual.empty:
        df_final, nuevas, cambios = df_nuevo, len(df_nuevo), 0
        previas = None
    else:
        k_act = _claves_orden(df_actual)
        ya = k_act.isin(set(k_new))
        nuevas = int((~k_new.isin(set(k_act))).sum())
        cambios = 0
        if "ESTATUS" in df_actual.columns and "ESTATUS" in df_nuevo.columns:
            viejo = pd.Series(df_actual.loc[ya, "ESTATUS"].map(normalizar_estatus).values, index=k_act[ya].values)
            viejo = viejo[~viejo.index.duplicated(keep="last")]
            nuevo = pd.Series(df_nuevo["ESTATUS"].map(normalizar_estatus).values, index=k_new.values)
            comunes = viejo.index.intersection(nuevo.index)
            cambios = int((viejo[comunes] != nuevo[comunes]).sum())
        df_final = pd.concat([df_actual.loc[~ya], df_nuevo], ignore_index=True)
        previas = len(df_actual)
    escribir_hoja(HOJAS["ordenes"], df_final, filas_previas=previas)
    return nuevas, cambios


def _validar_nivel_ads(df_actual, nivel_nuevo, usuario):
    sub = _filtrar_usuario(df_actual, usuario)
    if sub.empty:
        return
    niveles = set(sub["Nivel_Reporte"].fillna("Anuncio").astype(str)) if "Nivel_Reporte" in sub.columns else {"Anuncio"}
    if niveles != {nivel_nuevo}:
        raise ValueError(
            f"Ya tienes datos de Ads a nivel {sorted(niveles)} y este archivo es a nivel '{nivel_nuevo}'. "
            "Mezclar niveles duplicaría el gasto. Usa un solo nivel (o vacía Memoria_Ads, dejando el "
            "encabezado, y vuelve a cargar)."
        )


def guardar_ads(usuario, df_nuevo):
    """Inserta y ACTUALIZA filas de Ads (Meta corrige cifras de días recientes). Devuelve (nuevas, actualizadas)."""
    df_nuevo = df_nuevo.copy()
    nivel = str(df_nuevo["Nivel_Reporte"].iloc[0]) if "Nivel_Reporte" in df_nuevo.columns and len(df_nuevo) else "Anuncio"
    df_nuevo["Usuario"] = usuario
    df_nuevo["Fecha_Carga"] = ahora().strftime("%Y-%m-%d %H:%M")
    k_new = _claves_ads(df_nuevo)
    df_nuevo = df_nuevo[~k_new.duplicated(keep="last")]
    k_new = _claves_ads(df_nuevo)

    df_actual = leer_hoja(HOJAS["ads"], ttl=0, estricto=True)
    if df_actual.empty:
        df_final, nuevas, actualizadas, previas = df_nuevo, len(df_nuevo), 0, None
    else:
        _validar_nivel_ads(df_actual, nivel, usuario)
        k_act = _claves_ads(df_actual)
        ya = k_act.isin(set(k_new))
        nuevas = int((~k_new.isin(set(k_act))).sum())
        actualizadas = len(df_nuevo) - nuevas
        df_final = pd.concat([df_actual.loc[~ya], df_nuevo], ignore_index=True)
        previas = len(df_actual)
    escribir_hoja(HOJAS["ads"], df_final, filas_previas=previas)
    return nuevas, actualizadas


def cargar_vinculacion(usuario):
    df = _filtrar_usuario(leer_hoja(HOJAS["vinculacion"], ttl=60), usuario)
    if df.empty or not set(COLS_VINCULACION) <= set(df.columns):
        return pd.DataFrame(columns=COLS_VINCULACION)
    return df.drop_duplicates(["Campaña", "Conjunto_Anuncios", "Anuncio"], keep="last").reset_index(drop=True)


def _mascara_vinculo(df, usuario, campana, conjunto, anuncio):
    s = lambda c: df[c].astype(str).str.strip()
    return ((s("Usuario") == usuario) & (s("Campaña") == str(campana).strip())
            & (s("Conjunto_Anuncios") == str(conjunto).strip()) & (s("Anuncio") == str(anuncio).strip()))


def guardar_vinculo(usuario, campana, conjunto, anuncio, productos):
    """productos: lista de nombres de producto de Dropi (uno o varios)."""
    etiqueta = SEP_PRODUCTOS.join(sorted(set(productos)))
    df = leer_hoja(HOJAS["vinculacion"], ttl=0, estricto=True)
    previas = len(df)
    if df.empty or not set(COLS_VINCULACION) <= set(df.columns):
        df = pd.DataFrame(columns=COLS_VINCULACION)
    else:
        df = df[~_mascara_vinculo(df, usuario, campana, conjunto, anuncio)]
    nueva = {"Usuario": usuario, "Campaña": campana, "Conjunto_Anuncios": conjunto, "Anuncio": anuncio, "Producto": etiqueta}
    df = pd.concat([df, pd.DataFrame([nueva])], ignore_index=True)
    escribir_hoja(HOJAS["vinculacion"], df, filas_previas=previas)


def eliminar_vinculo(usuario, campana, conjunto, anuncio):
    df = leer_hoja(HOJAS["vinculacion"], ttl=0, estricto=True)
    if df.empty or not set(COLS_VINCULACION) <= set(df.columns):
        return
    df = df[~_mascara_vinculo(df, usuario, campana, conjunto, anuncio)]
    escribir_hoja(HOJAS["vinculacion"], df)  # reducir filas es intencional aquí


def limpiar_duplicados_ordenes(usuario):
    df = leer_hoja(HOJAS["ordenes"], ttl=0, estricto=True)
    if df.empty or "ID" not in df.columns:
        return 0
    antes = len(df)
    sin_dup = df[~_claves_orden(df).duplicated(keep="last")]  # se conserva la fila más reciente
    eliminadas = antes - len(sin_dup)
    if eliminadas > 0:
        escribir_hoja(HOJAS["ordenes"], sin_dup)
    return eliminadas


def limpiar_duplicados_ads(usuario):
    df = leer_hoja(HOJAS["ads"], ttl=0, estricto=True)
    if df.empty or "Inicio del informe" not in df.columns:
        return 0
    antes = len(df)
    sin_dup = df[~_claves_ads(df).duplicated(keep="last")]
    eliminadas = antes - len(sin_dup)
    if eliminadas > 0:
        escribir_hoja(HOJAS["ads"], sin_dup)
    return eliminadas


# ==============================================================================
# HELPERS DE NEGOCIO
# ==============================================================================
def selector_rango_fechas(key, index_default=1):
    eleccion = st.radio("Rango de fechas", RANGOS_FECHA, index=index_default, horizontal=True, key=f"radio_{key}")
    hoy = hoy_local()
    if eleccion == "Hoy":
        return hoy, hoy
    if eleccion == "Últimos 7 días":
        return hoy - timedelta(days=6), hoy
    if eleccion == "Últimos 30 días":
        return hoy - timedelta(days=29), hoy
    c1, c2 = st.columns(2)
    ini = c1.date_input("Desde", hoy - timedelta(days=30), key=f"ini_{key}")
    fin = c2.date_input("Hasta", hoy, key=f"fin_{key}")
    return ini, fin


def calcular_semaforo(valor_actual, serie_historica, usar_manual=False, umbral_bajo=None, umbral_alto=None):
    """Métrica tipo costo (menor = mejor). Devuelve (emoji, texto)."""
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
        return "🟢", "Mejor que sus propios días"
    if valor_actual <= p66:
        return "🟡", "Dentro de su promedio"
    return "🔴", "Peor que sus propios días"


def semaforo_producto(cpa_orden, cpa_equilibrio, config):
    """Semáforo del CPA REAL por orden generada de un producto."""
    if cpa_orden is None or pd.isna(cpa_orden):
        return "⚪", "Sin órdenes"
    if config["Usar_Semaforo_Manual"]:
        if cpa_orden <= config["CPA_Bajo_Manual"]:
            return "🟢", "Bajo tu umbral"
        if cpa_orden <= config["CPA_Alto_Manual"]:
            return "🟡", "Entre tus umbrales"
        return "🔴", "Sobre tu umbral"
    if cpa_equilibrio is None or pd.isna(cpa_equilibrio) or cpa_equilibrio <= 0:
        return "🔴", "Sin margen antes de publicidad"
    ratio = cpa_orden / cpa_equilibrio
    if ratio <= 0.6:
        return "🟢", "CPA ≤ 60% de tu punto de equilibrio"
    if ratio <= 1:
        return "🟡", "CPA cerca de tu punto de equilibrio"
    return "🔴", "CPA por encima del punto de equilibrio"


def recomendar_producto(margen, semaforo, n_ordenes):
    if n_ordenes < 10:
        return "🆕 Pocos datos: deja correr"
    if semaforo == "🟢" and margen > 0.15:
        return "📈 Escalar"
    if margen <= 0 or semaforo == "🔴":
        return "🛑 Reducir gasto / subir precio"
    return "⚖️ Ajustar presupuesto/precio"


def tasas_devolucion_cancelacion(df_ordenes, config):
    """Devolución = devueltas / (entregadas + devueltas)  [órdenes que ya se enviaron y cerraron].
    Cancelación = canceladas / órdenes generadas."""
    if config["Usar_Tasas_Manual"]:
        return config["Tasa_Devolucion_Manual"], config["Tasa_Cancelacion_Manual"]
    u = df_ordenes.drop_duplicates("ID")
    if u.empty:
        return 0.0, 0.0
    ent = u["ESTATUS"].isin(ESTATUS_ENTREGADO).sum()
    dev = u["ESTATUS"].isin(ESTATUS_DEVOLUCION).sum()
    canc = u["ESTATUS"].isin(ESTATUS_CANCELADO).sum()
    return (dev / (ent + dev) if (ent + dev) else 0.0), canc / len(u)


def calcular_precios(costo_prov, costo_flete, flete_dev, cpa_orden, tasa_dev, tasa_canc, margen):
    """Todo en la misma divisa. Modelo por cada orden GENERADA:
       s = (1-canc)*(1-dev)  -> se entrega (paga proveedor + flete)
       r = (1-canc)*dev      -> se devuelve (se pierde el flete de devolución)
       publicidad = cpa_orden por cada orden generada
    Precio mínimo por entrega = prov + flete + (flete_dev*r + cpa_orden) / s"""
    s = (1 - tasa_canc) * (1 - tasa_dev)
    r = (1 - tasa_canc) * tasa_dev
    if s <= 0:
        return dict(s=s, r=r, minimo=np.nan, recomendado=np.nan, utilidad=np.nan)
    minimo = costo_prov + costo_flete + (flete_dev * r + cpa_orden) / s
    recomendado = minimo / (1 - margen) if margen < 1 else np.nan
    return dict(s=s, r=r, minimo=minimo, recomendado=recomendado,
                utilidad=recomendado - minimo if pd.notna(recomendado) else np.nan)


def analisis_por_producto(df_vinc, df_ads_rango, df_ord_rango, config):
    """Cruza gasto de anuncios (vinculados) con órdenes reales de Dropi. Importes en divisa origen."""
    vacio = pd.DataFrame()
    if df_vinc.empty or df_ads_rango is None or df_ads_rango.empty or df_ord_rango is None or df_ord_rango.empty:
        return vacio
    vinc = df_vinc.merge(
        df_ads_rango[COLS_CLAVE_ADS + ["_gasto_o"]],
        left_on=["Campaña", "Conjunto_Anuncios", "Anuncio"], right_on=COLS_CLAVE_ADS, how="inner",
    )
    filas = []
    for etiqueta, sub in vinc.groupby("Producto"):
        productos = [p.strip() for p in str(etiqueta).split(SEP_PRODUCTOS.strip())]
        ords = df_ord_rango[df_ord_rango["PRODUCTO"].astype(str).str.strip().isin(productos)]
        u = ords.drop_duplicates("ID")
        n, n_ent = len(u), int(u["ESTATUS"].isin(ESTATUS_ENTREGADO).sum())
        n_dev, n_canc = int(u["ESTATUS"].isin(ESTATUS_DEVOLUCION).sum()), int(u["ESTATUS"].isin(ESTATUS_CANCELADO).sum())
        cerradas = n_ent + n_dev + n_canc
        gasto = sub["_gasto_o"].sum()
        ingreso = ords["_ingreso_o"].sum()
        costos = ords[["_costo_prov_o", "_flete_o", "_comision_o", "_perdida_dev_o"]].sum().sum()
        util_pre = ingreso - costos
        util_neta = util_pre - gasto
        cpa_orden = gasto / n if n else np.nan
        cpa_equilibrio = util_pre / cerradas if cerradas else np.nan  # solo órdenes ya cerradas
        sem, detalle = semaforo_producto(cpa_orden, cpa_equilibrio, config)
        margen = util_neta / ingreso if ingreso else 0.0
        filas.append({
            "Producto": etiqueta.replace(SEP_PRODUCTOS, " + "), "Gasto Ads": gasto, "Órdenes": n, "Entregadas": n_ent,
            "Devoluciones": n_dev, "En tránsito": n - cerradas,
            "CPA real x orden": cpa_orden, "CPA real x entrega": gasto / n_ent if n_ent else np.nan,
            "Ingreso": ingreso, "Utilidad antes de ads": util_pre, "Utilidad neta": util_neta, "Margen": margen,
            "CPA equilibrio x orden": cpa_equilibrio, "Semáforo": sem, "Detalle": detalle,
            "Recomendación": recomendar_producto(margen, sem, n),
        })
    return pd.DataFrame(filas).sort_values("Gasto Ads", ascending=False) if filas else vacio


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

# La tasa de HOY queda registrada en el histórico (automático, una vez por sesión)
if not config["Usar_TC_Manual"] and not st.session_state.get("tasas_registradas_hoy"):
    try:
        registrar_tasa_hoy(DD, DO)
        registrar_tasa_hoy(DA, DO)
    except Exception as e:
        st.warning(f"No se pudo actualizar el tipo de cambio automático de hoy (se usará el último disponible). Detalle: {e}")
    st.session_state["tasas_registradas_hoy"] = True

df_tc = leer_hoja(HOJAS["tipo_cambio"], ttl=300)
TC_ADS_MANUAL = config["TC_Ads_Manual"] if config["Usar_TC_Manual"] else None
TC_DROPI_MANUAL = config["TC_Dropi_Manual"] if config["Usar_TC_Manual"] else None

if not config["Usar_TC_Manual"]:
    for _base in sorted({DD, DA} - {DO}):
        if not construir_mapa_tasas(df_tc, _base, DO):
            st.warning(f"No hay tipo de cambio {_base}→{DO} guardado; los importes en {_base} se muestran SIN convertir. "
                       "Revisa tu conexión o captura un tipo de cambio manual en ⚙️ Panel de Control.")

df_ordenes = cargar_ordenes(usuario_activo)
COLS_ORDENES_BASE = ["ID", "ESTATUS", "FECHA_iso"]
if not df_ordenes.empty and all(c in df_ordenes.columns for c in COLS_ORDENES_BASE):
    df_ordenes = preparar_ordenes(df_ordenes, config["Devolucion_Cobra_Flete"], DD, DO, df_tc, TC_DROPI_MANUAL)

df_ads, divisa_ads_detectada = cargar_ads(usuario_activo)
if not df_ads.empty and {"FECHA_iso", "_gasto"} <= set(df_ads.columns):
    df_ads["_gasto_o"] = convertir_columna(df_ads, "FECHA_iso", "_gasto", DA, DO, df_tc, TC_ADS_MANUAL)

df_vinc = cargar_vinculacion(usuario_activo)

st.title("📊 Asistente de Rentabilidad para Dropshipping")

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["📊 Dashboard", "📣 Anuncios", "🧮 Calculadora", "⚙️ Panel de Control", "💹 Rentabilidad", "🧭 Recomendaciones"]
)

COLS_ORDENES_REQUERIDAS = ["ID", "ESTATUS", "TOTAL DE LA ORDEN", "PRECIO PROVEEDOR X CANTIDAD", "PRECIO FLETE",
                           "CANTIDAD", "PRODUCTO", "FECHA_iso", "_ingreso_o"]
COLS_ADS_REQUERIDAS = COLS_CLAVE_ADS + ["FECHA_iso", "_gasto_o"]

# ==============================================================================
# TAB 1 — DASHBOARD
# ==============================================================================
with tab1:
    if df_ordenes.empty:
        st.info("Tu memoria está vacía todavía. Ve a ⚙️ Panel de Control para subir tus primeros reportes.")
    elif not verificar_columnas(df_ordenes, COLS_ORDENES_REQUERIDAS, HOJAS["ordenes"]):
        pass
    else:
        ini, fin = selector_rango_fechas("dash")
        df_rango = en_rango(df_ordenes, "FECHA_iso", ini, fin)
        df_ads_rango = en_rango(df_ads, "FECHA_iso", ini, fin) if not df_ads.empty else df_ads

        ord_u = df_rango.drop_duplicates("ID")  # conteos por ORDEN, no por línea de producto
        ent_u = ord_u[ord_u["ESTATUS"].isin(ESTATUS_ENTREGADO)]
        dev_u = ord_u[ord_u["ESTATUS"].isin(ESTATUS_DEVOLUCION)]
        canc_u = ord_u[ord_u["ESTATUS"].isin(ESTATUS_CANCELADO)]
        tran_u = ord_u[~ord_u["ESTATUS"].isin(ESTATUS_FINALES)]

        ingreso = df_rango["_ingreso_o"].sum()
        costo_proveedor = df_rango["_costo_prov_o"].sum()
        costo_flete = df_rango["_flete_o"].sum()
        comision = df_rango["_comision_o"].sum()
        perdida_dev = df_rango["_perdida_dev_o"].sum()
        gasto_ads = df_ads_rango["_gasto_o"].sum() if not df_ads_rango.empty else 0.0

        total_gastos = costo_proveedor + costo_flete + comision + perdida_dev + gasto_ads
        utilidad_neta = ingreso - total_gastos
        articulos_vendidos = df_rango.loc[df_rango["ESTATUS"].isin(ESTATUS_ENTREGADO), "CANTIDAD"].sum()

        st.subheader(f"Resumen ({ini.strftime('%d/%m/%Y')} – {fin.strftime('%d/%m/%Y')}) en {DO}")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Órdenes generadas", len(ord_u))
        c2.metric("Artículos vendidos", int(articulos_vendidos))
        c3.metric("Ventas cobradas (entregadas)", f"${ingreso:,.2f}")
        c4.metric("Gastos totales", f"${total_gastos:,.2f}")
        c5.metric("Utilidad neta", f"${utilidad_neta:,.2f}", delta=f"{(utilidad_neta / ingreso * 100 if ingreso else 0):.1f}% margen")
        st.caption(
            f"Gastos = proveedor ${costo_proveedor:,.2f} + flete ${costo_flete:,.2f} + comisión ${comision:,.2f} "
            f"+ pérdida por devoluciones ${perdida_dev:,.2f} + publicidad ${gasto_ads:,.2f}. "
            "Proveedor y flete solo cuentan en órdenes entregadas."
        )

        potencial = df_rango["_potencial_o"].sum()
        cpa_orden = gasto_ads / len(ord_u) if len(ord_u) else 0.0
        cpa_entrega = gasto_ads / len(ent_u) if len(ent_u) else 0.0
        st.info(
            f"⏳ {len(tran_u)} órdenes siguen en tránsito: si se entregaran todas sumarían ≈ ${potencial:,.2f} {DO} "
            f"de utilidad antes de publicidad (no incluido arriba). "
            f"CPA real: ${cpa_orden:,.2f} por orden generada · ${cpa_entrega:,.2f} por orden entregada."
        )

        st.divider()
        st.subheader("Detalle de órdenes")
        d_ent_dev = len(ent_u) + len(dev_u)
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Entregadas", len(ent_u))
        c2.metric("Devoluciones", len(dev_u), delta=f"{(len(dev_u) / d_ent_dev * 100 if d_ent_dev else 0):.1f}% de las cerradas", delta_color="off")
        c3.metric("Cancelaciones", len(canc_u))
        c4.metric("En tránsito", len(tran_u))
        c5.metric("Órdenes generadas", len(ord_u))

        st.subheader("Gráfica dinámica (clic en la leyenda para activar/desactivar series)")
        metricas_disp = st.multiselect(
            "Datos a mostrar", ["Órdenes generadas", "Entregadas", "Devoluciones", "Cancelaciones"],
            default=["Órdenes generadas", "Entregadas", "Devoluciones", "Cancelaciones"], key="dash_metricas"
        )
        if ord_u.empty:
            st.info("No hay órdenes en el rango seleccionado.")
        elif metricas_disp:
            agg = ord_u.groupby(ord_u["FECHA_dt"].dt.date).agg(
                **{
                    "Órdenes generadas": ("ID", "count"),
                    "Entregadas": ("ESTATUS", lambda s: s.isin(ESTATUS_ENTREGADO).sum()),
                    "Devoluciones": ("ESTATUS", lambda s: s.isin(ESTATUS_DEVOLUCION).sum()),
                    "Cancelaciones": ("ESTATUS", lambda s: s.isin(ESTATUS_CANCELADO).sum()),
                }
            ).reset_index().rename(columns={"FECHA_dt": "Fecha"})
            largo = agg.melt(id_vars="Fecha", value_vars=metricas_disp, var_name="Métrica", value_name="Cantidad")
            fig_dash = px.bar(largo, x="Fecha", y="Cantidad", color="Métrica", barmode="group")
            st.plotly_chart(fig_dash, use_container_width=True)
        else:
            st.info("Selecciona al menos un dato para graficar.")


# ==============================================================================
# TAB 2 — ANUNCIOS
# ==============================================================================
with tab2:
    if df_ads.empty:
        st.info("Aún no subes reportes de Meta Ads. Ve a ⚙️ Panel de Control.")
    elif not verificar_columnas(df_ads, COLS_ADS_REQUERIDAS, HOJAS["ads"]):
        pass
    else:
        ini, fin = selector_rango_fechas("ads")
        df_ads_rango = en_rango(df_ads, "FECHA_iso", ini, fin).copy()

        gasto_camp = df_ads_rango.groupby("Nombre de la campaña")["_gasto"].sum().sort_values(ascending=False)
        campanas = list(gasto_camp.index)
        seleccion = st.multiselect("Filtrar por campaña", campanas, default=[c for c in campanas[:3] if gasto_camp[c] > 0])
        df_vista = df_ads_rango[df_ads_rango["Nombre de la campaña"].isin(seleccion)] if seleccion else df_ads_rango

        st.subheader(f"Comportamiento diario (gasto en {DA}, costo por resultado, alcance)")
        diario = df_vista.groupby("FECHA_iso").agg(Gasto=("_gasto", "sum"), Resultados=("_resultados", "sum"), Alcance=("_alcance", "sum")).reset_index()
        diario["Costo por resultado"] = diario["Gasto"] / diario["Resultados"].replace(0, np.nan)
        fig1 = go.Figure()
        fig1.add_trace(go.Bar(x=diario["FECHA_iso"], y=diario["Gasto"], name="Gasto"))
        fig1.add_trace(go.Bar(x=diario["FECHA_iso"], y=diario["Costo por resultado"], name="Costo por resultado (Meta)", yaxis="y2"))
        fig1.update_layout(barmode="group", yaxis=dict(title="Gasto"), yaxis2=dict(title="Costo por resultado", overlaying="y", side="right"), legend=dict(orientation="h"))
        st.plotly_chart(fig1, use_container_width=True)
        st.plotly_chart(px.bar(diario, x="FECHA_iso", y="Alcance", title="Alcance diario"), use_container_width=True)

        st.subheader(f"Órdenes del día vs. gasto en publicidad ({DO})")
        # Si hay campañas filtradas arriba, se usan sus productos vinculados (sección 6) para
        # contar solo LAS ÓRDENES DE ESOS PRODUCTOS — así el gasto y las órdenes comparan lo mismo,
        # en vez del total de la tienda contra el gasto de solo algunas campañas.
        productos_filtro = None
        if seleccion and not df_vinc.empty:
            vinc_sel = df_vinc[df_vinc["Campaña"].isin(seleccion)]
            if not vinc_sel.empty:
                productos_filtro = set()
                for etiqueta in vinc_sel["Producto"]:
                    productos_filtro.update(p.strip() for p in str(etiqueta).split(SEP_PRODUCTOS.strip()))

        ord_rango_all = en_rango(df_ordenes, "FECHA_iso", ini, fin) if not df_ordenes.empty else df_ordenes
        if ord_rango_all is not None and not ord_rango_all.empty:
            ord_base = ord_rango_all.drop_duplicates("ID")
            if productos_filtro:
                ord_base = ord_base[ord_base["PRODUCTO"].astype(str).str.strip().isin(productos_filtro)]
            ord_dia = ord_base.groupby("FECHA_iso").size().reset_index(name="Órdenes")
        else:
            ord_dia = pd.DataFrame(columns=["FECHA_iso", "Órdenes"])
        gasto_dia = df_vista.groupby("FECHA_iso")["_gasto_o"].sum().reset_index(name="Gasto")  # respeta el filtro de campaña
        comparativo = gasto_dia.merge(ord_dia, on="FECHA_iso", how="outer").fillna(0).sort_values("FECHA_iso")
        comparativo["Costo por orden"] = comparativo["Gasto"] / comparativo["Órdenes"].replace(0, np.nan)

        fig2 = go.Figure()
        fig2.add_trace(go.Bar(x=comparativo["FECHA_iso"], y=comparativo["Órdenes"], name="Órdenes"))
        fig2.add_trace(go.Bar(x=comparativo["FECHA_iso"], y=comparativo["Gasto"], name="Gasto Ads", yaxis="y2"))
        fig2.update_layout(barmode="group", yaxis=dict(title="Órdenes"), yaxis2=dict(title=f"Gasto ({DO})", overlaying="y", side="right"))
        st.plotly_chart(fig2, use_container_width=True)

        tabla_comp = comparativo.rename(columns={"FECHA_iso": "Fecha", "Gasto": f"Gasto ({DO})", "Costo por orden": f"Costo por orden ({DO})"})
        st.dataframe(tabla_comp[["Fecha", "Órdenes", f"Gasto ({DO})", f"Costo por orden ({DO})"]].round(2), use_container_width=True, hide_index=True)
        if seleccion and productos_filtro:
            st.caption(f"Órdenes filtradas a los productos vinculados a {', '.join(seleccion)}.")
        elif seleccion:
            st.warning("Las campañas filtradas arriba no tienen productos vinculados todavía (sección 6 del Panel de Control), "
                       "así que la tabla muestra el total de órdenes de la tienda, no solo las de esas campañas.")

        st.divider()
        st.subheader(f"Rentabilidad real por producto (cruce Ads × Dropi) — importes en {DO}")
        st.caption(
            "El costo por resultado de Meta cuenta conversaciones/pixel, no ventas. Aquí el CPA REAL divide el gasto "
            "de los anuncios vinculados entre las órdenes reales de Dropi de ese producto en el periodo. "
            "'CPA equilibrio' = cuánto puedes pagar por orden antes de perder (calculado con tus órdenes ya cerradas)."
        )
        if df_vinc.empty:
            st.warning("Todavía no vinculas anuncios a productos. Hazlo en ⚙️ Panel de Control para ver esta tabla.")
        else:
            tabla_prod = analisis_por_producto(df_vinc, df_ads_rango, en_rango(df_ordenes, "FECHA_iso", ini, fin), config)
            if tabla_prod.empty:
                st.info("Sin coincidencias entre anuncios vinculados y órdenes en este periodo.")
            else:
                st.dataframe(
                    tabla_prod[["Producto", "Gasto Ads", "Órdenes", "Entregadas", "Devoluciones", "En tránsito", "CPA real x orden",
                                "CPA real x entrega", "CPA equilibrio x orden", "Utilidad neta", "Semáforo", "Detalle"]].round(2),
                    use_container_width=True, hide_index=True,
                )

        st.subheader("Rendimiento y ranking por anuncio")
        st.caption("El semáforo compara el costo por resultado del periodo contra los propios días del anuncio (Meta mide conversaciones, no ventas).")
        filas = []
        for (camp, conj, anun), sub in df_ads_rango.groupby(COLS_CLAVE_ADS):
            gasto = sub["_gasto"].sum()
            if gasto <= 0:
                continue  # campañas sin gasto en el periodo no aportan
            resultados = sub["_resultados"].sum()
            cpr = gasto / resultados if resultados > 0 else np.nan
            hist = df_ads[(df_ads[COLS_CLAVE_ADS[0]] == camp) & (df_ads[COLS_CLAVE_ADS[1]] == conj) & (df_ads[COLS_CLAVE_ADS[2]] == anun)]["_cpa_meta"]
            semaforo, texto = calcular_semaforo(cpr, hist)
            con_cpa = sub.dropna(subset=["_cpa_meta"])
            dia_alto = con_cpa.loc[con_cpa["_cpa_meta"].idxmax(), "FECHA_iso"] if not con_cpa.empty else "-"
            dia_bajo = con_cpa.loc[con_cpa["_cpa_meta"].idxmin(), "FECHA_iso"] if not con_cpa.empty else "-"
            recomendacion = {"🟢": "🚀 Escalar presupuesto", "🟡": "⚖️ Mantener / ajustar ±10%", "🔴": "🛑 Reducir o apagar"}.get(semaforo, "ℹ️ Falta historial")
            filas.append({
                "Campaña": camp, "Conjunto": conj, "Anuncio": anun, f"Gasto ({DA})": round(gasto, 2), "Resultados": int(resultados),
                f"Costo por resultado ({DA})": round(cpr, 2) if pd.notna(cpr) else None, "Alcance": int(sub["_alcance"].sum()),
                "Día con costo más alto": dia_alto, "Día con costo más bajo": dia_bajo,
                "Semáforo": semaforo, "Detalle": texto, "Recomendación": recomendacion,
            })
        if filas:
            tabla_anuncios = pd.DataFrame(filas).sort_values(f"Gasto ({DA})", ascending=False)
            st.dataframe(tabla_anuncios, use_container_width=True, hide_index=True)

            st.subheader(f"🏆 Ranking por costo por resultado (mínimo {MIN_RESULTADOS_RANKING} resultados, de menor a mayor)")
            ranking = tabla_anuncios[tabla_anuncios["Resultados"] >= MIN_RESULTADOS_RANKING].sort_values(f"Costo por resultado ({DA})").head(10)
            st.dataframe(ranking[["Anuncio", "Campaña", "Resultados", f"Costo por resultado ({DA})", f"Gasto ({DA})", "Semáforo"]], use_container_width=True, hide_index=True)
        else:
            st.info("Ningún anuncio tuvo gasto en este periodo.")


# ==============================================================================
# TAB 3 — CALCULADORA DE PRECIOS
# ==============================================================================
with tab3:
    st.header(f"Calculadora de precios ({DO})")
    st.caption("Todos los montos se muestran ya convertidos a tu divisa origen.")

    col_a, col_b = st.columns(2)
    with col_a:
        costo_prov = st.number_input(f"Costo proveedor ({DD})", min_value=0.0, value=45.0)
        costo_flete_in = st.number_input(f"Costo de flete de ida ({DD}) — solo se paga si se entrega", min_value=0.0, value=35.0)
        flete_devolucion = st.number_input(f"Pérdida por orden devuelta ({DD})", min_value=0.0, value=35.0)
        cpa_esperado = st.number_input(f"Publicidad por orden GENERADA ({DA})", min_value=0.0, value=100.0)
    with col_b:
        hay_hist = not df_ordenes.empty and {"ID", "ESTATUS"} <= set(df_ordenes.columns)
        tasa_dev_hist, tasa_canc_hist = tasas_devolucion_cancelacion(df_ordenes, config) if hay_hist else (0.15, 0.05)
        tasa_dev = st.slider("Devoluciones (% de las órdenes enviadas)", 0, 90, min(90, int(round(tasa_dev_hist * 100)))) / 100
        tasa_canc = st.slider("Cancelaciones (% de las órdenes generadas)", 0, 90, min(90, int(round(tasa_canc_hist * 100)))) / 100
        margen_deseado = st.slider("Margen de utilidad deseado sobre el precio (%)", 0, 80, 20) / 100

    tc_dropi_hoy = TC_DROPI_MANUAL or tasa_mas_reciente(construir_mapa_tasas(df_tc, DD, DO))
    tc_ads_hoy = TC_ADS_MANUAL or tasa_mas_reciente(construir_mapa_tasas(df_tc, DA, DO))
    res = calcular_precios(
        costo_prov * tc_dropi_hoy, costo_flete_in * tc_dropi_hoy, flete_devolucion * tc_dropi_hoy,
        cpa_esperado * tc_ads_hoy, tasa_dev, tasa_canc, margen_deseado,
    )

    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric(f"Precio mínimo (equilibrio) — {DO}", f"${res['minimo']:,.2f}" if pd.notna(res["minimo"]) else "N/A")
    c2.metric(f"Precio recomendado (con {int(margen_deseado * 100)}% margen) — {DO}", f"${res['recomendado']:,.2f}" if pd.notna(res["recomendado"]) else "N/A")
    c3.metric(f"Utilidad por orden entregada — {DO}", f"${res['utilidad']:,.2f}" if pd.notna(res["utilidad"]) else "N/A")
    st.caption(
        f"De cada 100 órdenes generadas: {res['s'] * 100:.0f} se entregan, {res['r'] * 100:.0f} se devuelven y "
        f"{tasa_canc * 100:.0f} se cancelan. La publicidad de las 100 órdenes la pagan solo las entregadas."
    )
    if DD != DO and pd.notna(res["recomendado"]) and tc_dropi_hoy:
        st.caption(f"Precio recomendado equivalente: ${res['recomendado'] / tc_dropi_hoy:,.2f} {DD}")


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
    usar_manual_tc = st.checkbox("Sobrescribir el tipo de cambio automático con uno manual (se aplica a TODAS las fechas)", value=config["Usar_TC_Manual"])
    c1, c2 = st.columns(2)
    tc_ads_manual = c1.number_input(f"1 {nueva_da} = X {nueva_do}", min_value=0.0, value=float(config["TC_Ads_Manual"]), format="%.4f", disabled=not usar_manual_tc)
    tc_dropi_manual = c2.number_input(f"1 {nueva_dd} = X {nueva_do}", min_value=0.0, value=float(config["TC_Dropi_Manual"]), format="%.4f", disabled=not usar_manual_tc)
    if usar_manual_tc:
        for _b, _v in [(nueva_da, tc_ads_manual), (nueva_dd, tc_dropi_manual)]:
            if _b != nueva_do and _v:
                _ref = (_obtener_tasas_api(_b) or {}).get(nueva_do)
                if _ref and abs(_v / _ref - 1) > 0.3:
                    st.warning(f"Tu tasa manual 1 {_b} = {_v:g} {nueva_do} se aleja mucho de la de mercado (≈ {_ref:.4f}). ¿La capturaste al revés?")
    st.caption("Con 'automático' el sistema consulta el tipo de cambio del día una vez al día y lo guarda en Tipo_Cambio_Historico, para que tus cálculos pasados no cambien.")

    st.subheader("3. Semáforo de CPA real por producto")
    usar_manual_sem = st.checkbox("Usar umbrales fijos en vez de compararlo con mi punto de equilibrio", value=config["Usar_Semaforo_Manual"])
    c1, c2 = st.columns(2)
    cpa_bajo = c1.number_input(f"CPA bueno (verde) si es menor o igual a ({nueva_do} por orden generada)", min_value=0.0, value=float(config["CPA_Bajo_Manual"]), disabled=not usar_manual_sem)
    cpa_alto = c2.number_input(f"CPA malo (rojo) si es mayor a ({nueva_do} por orden generada)", min_value=0.0, value=float(config["CPA_Alto_Manual"]), disabled=not usar_manual_sem)

    st.subheader("4. Tasas de devolución y cancelación")
    usar_manual_tasas = st.checkbox("Usar tasas manuales en vez de calcularlas del histórico", value=config["Usar_Tasas_Manual"])
    c1, c2 = st.columns(2)
    tasa_dev_manual = c1.slider("Tasa de devoluciones manual (%)", 0, 90, int(config["Tasa_Devolucion_Manual"] * 100), disabled=not usar_manual_tasas) / 100
    tasa_canc_manual = c2.slider("Tasa de cancelaciones manual (%)", 0, 90, int(config["Tasa_Cancelacion_Manual"] * 100), disabled=not usar_manual_tasas) / 100

    st.subheader("5. Costo de una devolución")
    dev_cobra_flete = st.checkbox(
        "Cuando una orden se devuelve, pierdo el flete de ida (Dropi lo cobra)", value=config["Devolucion_Cobra_Flete"],
        help="En tu reporte 'COSTO DEVOLUCION FLETE' viene en 0 incluso en devoluciones. Confirma en tu billetera de Dropi "
             "cuánto se descuenta realmente por devolución; si no descuenta nada, desmarca esta opción.",
    )

    if st.session_state.get("_msg_config"):
        tipo, texto = st.session_state.pop("_msg_config")
        getattr(st, tipo)(texto)

    if st.button("💾 Guardar configuración"):
        try:
            guardar_config(usuario_activo, {
                "Divisa_Origen": nueva_do, "Divisa_Ads": nueva_da, "Divisa_Dropi": nueva_dd,
                "Usar_TC_Manual": usar_manual_tc, "TC_Ads_Manual": tc_ads_manual, "TC_Dropi_Manual": tc_dropi_manual,
                "Usar_Semaforo_Manual": usar_manual_sem, "CPA_Bajo_Manual": cpa_bajo, "CPA_Alto_Manual": cpa_alto,
                "Usar_Tasas_Manual": usar_manual_tasas, "Tasa_Devolucion_Manual": tasa_dev_manual, "Tasa_Cancelacion_Manual": tasa_canc_manual,
                "Devolucion_Cobra_Flete": dev_cobra_flete,
            })
            if usar_manual_tc:
                registrar_tasa_manual(nueva_da, nueva_do, tc_ads_manual)
                registrar_tasa_manual(nueva_dd, nueva_do, tc_dropi_manual)
            st.session_state["_msg_config"] = ("success", "Configuración guardada.")
        except Exception as e:
            st.session_state["_msg_config"] = ("error", f"No se guardó la configuración: {e}")
        st.rerun()

    st.divider()
    st.subheader("6. Subir reportes diarios")
    c1, c2 = st.columns(2)
    archivo_ordenes = c1.file_uploader("Reporte de órdenes de Dropi (.xlsx)", type=["xlsx"])
    archivo_ads = c2.file_uploader("Reporte de Meta Ads con desglose por día — nivel Campaña o Anuncio (.csv)", type=["csv"])
    st.caption("Si subes un archivo con órdenes/días que ya existen, se ACTUALIZAN (estatus nuevos, cifras corregidas por Meta); no se duplican.")

    if st.session_state.get("_msg_carga"):
        tipo, texto = st.session_state.pop("_msg_carga")
        getattr(st, tipo)(texto)

    if st.button("💾 Guardar en memoria"):
        mensajes = []
        if archivo_ordenes is None and archivo_ads is None:
            mensajes.append(("warning", "Sube al menos un archivo."))
        else:
            if archivo_ordenes is not None:
                try:
                    nuevas, cambios = guardar_ordenes(usuario_activo, pd.read_excel(archivo_ordenes))
                    mensajes.append(("success", f"Órdenes: {nuevas} nuevas; {cambios} existentes cambiaron de estatus y se actualizaron."))
                except Exception as e:
                    mensajes.append(("error", f"Órdenes NO guardadas: {e}"))
            if archivo_ads is not None:
                try:
                    df_ads_csv, nivel = normalizar_ads_csv(pd.read_csv(archivo_ads, encoding="utf-8-sig"))
                    nuevas, actualizadas = guardar_ads(usuario_activo, df_ads_csv)
                    mensajes.append(("success", f"Anuncios (nivel {nivel}): {nuevas} filas nuevas y {actualizadas} actualizadas."))
                except Exception as e:
                    mensajes.append(("error", f"Anuncios NO guardados: {e}"))
        tipos = {t for t, _ in mensajes}
        tipo = "error" if "error" in tipos else ("warning" if "warning" in tipos else "success")
        st.session_state["_msg_carga"] = (tipo, " · ".join(m for _, m in mensajes))
        st.rerun()

    st.divider()
    st.subheader("7. Vincular anuncios a productos")
    if df_ads.empty:
        st.info("Sube primero un reporte de anuncios.")
    elif not verificar_columnas(df_ads, COLS_ADS_REQUERIDAS, HOJAS["ads"]):
        pass
    else:
        # Solo anuncios que alguna vez gastaron (las campañas apagadas sin gasto no hace falta vincularlas)
        combinaciones = df_ads[df_ads["_gasto"] > 0][COLS_CLAVE_ADS].drop_duplicates()
        ya_vinculadas = set(zip(df_vinc["Campaña"], df_vinc["Conjunto_Anuncios"], df_vinc["Anuncio"]))
        combinaciones["_clave"] = list(zip(*[combinaciones[c] for c in COLS_CLAVE_ADS]))
        pendientes = combinaciones[~combinaciones["_clave"].isin(ya_vinculadas)]
        productos_disponibles = sorted(df_ordenes["PRODUCTO"].dropna().astype(str).str.strip().unique()) if (not df_ordenes.empty and "PRODUCTO" in df_ordenes.columns) else []

        st.caption(f"{len(pendientes)} anuncios con gasto sin vincular de {len(combinaciones)} con gasto.")
        if st.session_state.get("_msg_vinc"):
            tipo, texto = st.session_state.pop("_msg_vinc")
            getattr(st, tipo)(texto)
        if not pendientes.empty and productos_disponibles:
            fila = pendientes.iloc[0]
            st.write(f"**Campaña:** {fila[COLS_CLAVE_ADS[0]]} · **Conjunto:** {fila[COLS_CLAVE_ADS[1]]} · **Anuncio:** {fila[COLS_CLAVE_ADS[2]]}")
            productos_sel = st.multiselect("¿Qué producto(s) vende este anuncio? (puedes elegir varios)", productos_disponibles, key="vinc_producto")
            if st.button("🔗 Vincular"):
                if not productos_sel:
                    st.session_state["_msg_vinc"] = ("warning", "Elige al menos un producto.")
                else:
                    try:
                        guardar_vinculo(usuario_activo, fila[COLS_CLAVE_ADS[0]], fila[COLS_CLAVE_ADS[1]], fila[COLS_CLAVE_ADS[2]], productos_sel)
                        st.session_state["_msg_vinc"] = ("success", "Vinculado.")
                    except Exception as e:
                        st.session_state["_msg_vinc"] = ("error", f"No se guardó el vínculo: {e}")
                st.rerun()
        elif pendientes.empty:
            st.success("Todos tus anuncios con gasto están vinculados a un producto. 🎉")

        if not df_vinc.empty:
            with st.expander("Ver / quitar vínculos existentes"):
                st.dataframe(df_vinc[["Campaña", "Conjunto_Anuncios", "Anuncio", "Producto"]], use_container_width=True, hide_index=True)
                i_quitar = st.selectbox(
                    "Quitar vínculo de:", list(range(len(df_vinc))), key="vinc_quitar",
                    format_func=lambda i: f"{df_vinc.loc[i, 'Campaña']} · {df_vinc.loc[i, 'Anuncio']} → {df_vinc.loc[i, 'Producto']}",
                )
                if st.button("🗑️ Quitar vínculo"):
                    r = df_vinc.loc[i_quitar]
                    try:
                        eliminar_vinculo(usuario_activo, r["Campaña"], r["Conjunto_Anuncios"], r["Anuncio"])
                        st.session_state["_msg_vinc"] = ("success", "Vínculo eliminado.")
                    except Exception as e:
                        st.session_state["_msg_vinc"] = ("error", f"No se pudo quitar: {e}")
                    st.rerun()

    st.divider()
    st.subheader("8. Limpieza de duplicados")
    st.caption("Úsalo si tus totales no cuadran con Dropi/Meta. Se conserva la fila más reciente de cada duplicado.")
    c1, c2 = st.columns(2)
    if st.session_state.get("_msg_limpieza"):
        tipo, texto = st.session_state.pop("_msg_limpieza")
        getattr(st, tipo)(texto)
    for _col, _txt, _fn, _hoja in [
        (c1, "🧹 Detectar y eliminar duplicados en Órdenes", limpiar_duplicados_ordenes, "Memoria_Ordenes"),
        (c2, "🧹 Detectar y eliminar duplicados en Anuncios", limpiar_duplicados_ads, "Memoria_Ads"),
    ]:
        if _col.button(_txt):
            try:
                eliminadas = _fn(usuario_activo)
                st.session_state["_msg_limpieza"] = (("success", f"Se eliminaron {eliminadas} filas duplicadas de {_hoja}.") if eliminadas > 0
                                                     else ("info", f"No se encontraron duplicados en {_hoja}."))
            except Exception as e:
                st.session_state["_msg_limpieza"] = ("error", f"No se pudo limpiar {_hoja}: {e}")
            st.rerun()


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
        df_rango = en_rango(df_ordenes, "FECHA_iso", ini, fin)
        df_ads_rango = en_rango(df_ads, "FECHA_iso", ini, fin) if not df_ads.empty else pd.DataFrame()
        gasto_por_dia = df_ads_rango.groupby("FECHA_iso")["_gasto_o"].sum().to_dict() if not df_ads_rango.empty else {}
        fechas_orden = set(df_rango["FECHA_iso"].dropna())
        # Incluye días con gasto en anuncios aunque no haya habido órdenes ese día
        todas_fechas = sorted(fechas_orden | set(gasto_por_dia))

        filas = []
        for f in todas_fechas:
            sub = df_rango[df_rango["FECHA_iso"] == f]
            su = sub.drop_duplicates("ID")
            total = len(su)
            devoluciones = int(su["ESTATUS"].isin(ESTATUS_DEVOLUCION).sum())
            cancelaciones = int(su["ESTATUS"].isin(ESTATUS_CANCELADO).sum())
            transito = int((~su["ESTATUS"].isin(ESTATUS_FINALES)).sum())
            ingreso = sub["_ingreso_o"].sum()
            costos = sub[["_costo_prov_o", "_flete_o", "_comision_o", "_perdida_dev_o"]].sum().sum()
            gasto_ads_dia = float(gasto_por_dia.get(f, 0.0))
            utilidad_dia = ingreso - costos - gasto_ads_dia
            filas.append({
                "Fecha": f, "Órdenes": total, "Devoluciones": devoluciones, "Cancelaciones": cancelaciones, "En tránsito": transito,
                "Tasa Dev.": f"{devoluciones / total * 100:.1f}%" if total else "-",
                "Tasa Canc.": f"{cancelaciones / total * 100:.1f}%" if total else "-",
                f"Ingreso ({DO})": round(ingreso, 2), f"Costos Dropi ({DO})": round(costos, 2),
                f"Gasto Ads ({DO})": round(gasto_ads_dia, 2), f"Utilidad del día ({DO})": round(utilidad_dia, 2),
                "Rentable": "🟢" if utilidad_dia > 0 else "🔴",
            })
        if filas:
            tabla_rent = pd.DataFrame(filas).sort_values("Fecha", ascending=False)
            st.subheader(f"Tabla de rentabilidad diaria ({DO})")
            st.dataframe(tabla_rent, use_container_width=True, hide_index=True)
            dias_rentables = (tabla_rent["Rentable"] == "🟢").sum()
            col_util = f"Utilidad del día ({DO})"
            st.caption(
                f"{dias_rentables} de {len(tabla_rent)} días fueron rentables · Utilidad acumulada del rango: ${tabla_rent[col_util].sum():,.2f} {DO}. "
                "Cada orden se asigna al día en que se generó; las que siguen en tránsito aún no suman ingreso, "
                "así que los días recientes se ven peor de lo que terminarán."
            )
        else:
            st.info("No hay órdenes ni gasto en el rango seleccionado.")


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
        desde = hoy_local() - timedelta(days=dias_analisis)
        df_ads_periodo = en_rango(df_ads, "FECHA_iso", desde, hoy_local())

        st.subheader(f"Presupuestos recomendados por anuncio ({DA})")
        filas = []
        for (camp, conj, anun), sub in df_ads_periodo.groupby(COLS_CLAVE_ADS):
            dias_con_gasto = sub[sub["_gasto"] > 0].groupby("FECHA_iso")["_gasto"].sum()
            if dias_con_gasto.empty:
                continue  # anuncio apagado en el periodo
            gasto_diario_prom = dias_con_gasto.mean()
            resultados = sub["_resultados"].sum()
            cpr = sub["_gasto"].sum() / resultados if resultados > 0 else np.nan
            hist = df_ads[(df_ads[COLS_CLAVE_ADS[0]] == camp) & (df_ads[COLS_CLAVE_ADS[1]] == conj) & (df_ads[COLS_CLAVE_ADS[2]] == anun)]["_cpa_meta"]
            semaforo, _ = calcular_semaforo(cpr, hist)
            accion, factor = {
                "🟢": ("🚀 Subir presupuesto", 1.3), "🟡": ("⚖️ Ajustar ligeramente", 1.0), "🔴": ("🛑 Bajar presupuesto / apagar", 0.5),
            }.get(semaforo, ("🆕 Dejar correr (poco historial)", 1.0))
            filas.append({
                "Campaña": camp, "Conjunto": conj, "Anuncio": anun,
                f"Gasto diario actual ({DA})": round(gasto_diario_prom, 2), "Semáforo": semaforo, "Acción sugerida": accion,
                f"Presupuesto diario sugerido ({DA})": round(gasto_diario_prom * factor, 2),
            })
        if filas:
            tabla_presup = pd.DataFrame(filas).sort_values(f"Gasto diario actual ({DA})", ascending=False)
            st.dataframe(tabla_presup, use_container_width=True, hide_index=True)
            st.caption("Heurística de reglas (no garantiza resultados): +30% en verde, mantener en amarillo, −50% en rojo. "
                       "Meta mide conversaciones; valida siempre con la tabla por producto de abajo (CPA real).")
        else:
            st.info("Ningún anuncio tuvo gasto en el periodo.")

        st.divider()
        st.subheader(f"Recomendaciones por producto ({DO})")
        if df_vinc.empty:
            st.info("Vincula tus anuncios a productos en ⚙️ Panel de Control para ver esta tabla.")
        else:
            tabla_prod_rec = analisis_por_producto(df_vinc, df_ads_periodo, en_rango(df_ordenes, "FECHA_iso", desde, hoy_local()), config)
            if tabla_prod_rec.empty:
                st.info("Sin coincidencias entre anuncios vinculados y órdenes en este periodo.")
            else:
                tabla_prod_rec = tabla_prod_rec.sort_values("Utilidad neta", ascending=False)
                tabla_prod_rec["Margen"] = (tabla_prod_rec["Margen"] * 100).round(1).astype(str) + "%"
                st.dataframe(
                    tabla_prod_rec[["Producto", "Órdenes", "En tránsito", "Ingreso", "Gasto Ads", "Utilidad neta", "Margen",
                                    "CPA real x orden", "CPA equilibrio x orden", "Semáforo", "Recomendación"]].round(2),
                    use_container_width=True, hide_index=True,
                )
                st.caption("Las órdenes en tránsito aún no suman ingreso: con periodos cortos la utilidad se ve subestimada.")

"""
Dashboard SAC
- Fuente: OneDrive via Microsoft Graph API
- Roles: viewer (público) / admin (login requerido para editar)
- Credenciales en Streamlit Secrets
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import requests
from io import BytesIO
from datetime import datetime
import numpy as np

# ─────────────────────────────────────────────
# CONFIGURACIÓN GENERAL
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="Dashboard SAC",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────

SHEET_ABIERTOS      = "Consolidado"
SHEET_CERRADOS      = "Gestion_SAC"
BINS_SEMAFORO       = [-999, 5, 10, 999]
LABELS_SEMAFORO     = ["🟢 En tiempo", "🟡 En riesgo", "🔴 Vencido"]
COLUMNAS_REQUERIDAS = ["FechaCreacion", "Responsable", "NombreSeccionales", "NUI"]

# Paleta ampliada — usada en todos los gráficos de categorías
PALETA = [
    "#8f5cda","#7069d8","#3a81d5","#38a9d2","#4cb2ca",
    "#a78bfa","#60a5fa","#34d399","#f472b6","#fb923c",
    "#facc15","#e879f9","#2dd4bf","#818cf8","#f87171",
    "#4ade80","#38bdf8","#c084fc","#fb7185","#a3e635",
]

# Columnas visibles en las tablas del dashboard (en orden)
COLUMNAS_TABLA = [
    "NUI", "NombreSeccionales", "Id_Tickets", "Semaforo",
    "Menu", "SubMenu1", "SubMenu2", "SubMenu3", "FechaCreacion", "Dias_Habiles",
    "Creador_gestion", "Responsable", "Fecha Asignación",
    "Descripción", "Fecha Respuesta",
]

# ─────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────

st.markdown("""
<style>
.main, .block-container { background-color: #0f1116; }
section[data-testid="stSidebar"] { background-color: #161a24; }
h1, h2, h3, h4 { color: #8f5cda !important; }
p, label, .stMarkdown { color: #d1d5db; }
div[data-testid="metric-container"] {
    background: linear-gradient(135deg, #161a24, #1e2130);
    border: 1px solid #2a2f3a;
    border-left: 4px solid #8f5cda;
    padding: 14px 18px; border-radius: 12px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.3);
}
div[data-testid="stMetricValue"] { color: #ffffff !important; font-size: 2rem !important; }
div[data-testid="stMetricLabel"] { color: #9ca3af !important; }
div[data-testid="stMetricDelta"] { font-size: 0.85rem; }
.stDataFrame { border-radius: 10px; overflow: hidden; }
.stButton > button {
    background: linear-gradient(135deg, #8f5cda, #3a81d5);
    color: white; border: none; border-radius: 8px;
    padding: 8px 20px; font-weight: 600; transition: opacity .2s;
}
.stButton > button:hover { opacity: .85; }
hr { border-color: #2a2f3a; }
.badge {
    display: inline-block; background: #1e2130;
    border: 1px solid #3a81d5; color: #38a9d2;
    border-radius: 20px; padding: 2px 14px;
    font-size: 0.75rem; margin-bottom: 12px;
}
.role-admin {
    display: inline-block; background: #1a2e1a;
    border: 1px solid #4ade80; color: #4ade80;
    border-radius: 20px; padding: 2px 12px; font-size: 0.75rem;
}
.role-viewer {
    display: inline-block; background: #1e2130;
    border: 1px solid #6b7280; color: #9ca3af;
    border-radius: 20px; padding: 2px 12px; font-size: 0.75rem;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# MICROSOFT GRAPH — TOKEN
# ─────────────────────────────────────────────

def get_token() -> str:
    """
    Obtiene access token usando refresh_token (flujo delegado).
    El refresh_token se renueva automáticamente en cada llamada.
    """
    cfg  = st.secrets["microsoft"]
    resp = requests.post(
        f"https://login.microsoftonline.com/{cfg['tenant_id']}/oauth2/v2.0/token",
        data={
            "grant_type":    "refresh_token",
            "client_id":     cfg["client_id"],
            "client_secret": cfg["client_secret"],
            "refresh_token": cfg["refresh_token"],
            "scope":         "Files.ReadWrite offline_access User.Read",
        },
        timeout=15,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def headers() -> dict:
    return {"Authorization": f"Bearer {get_token()}"}

# ─────────────────────────────────────────────
# MICROSOFT GRAPH — LEER / ESCRIBIR
# ─────────────────────────────────────────────

def _drive_url() -> str:
    """Construye la URL base del archivo en el OneDrive del usuario autenticado."""
    file_path = st.secrets["microsoft"]["file_path"]
    return f"https://graph.microsoft.com/v1.0/me/drive/root:{file_path}:"


# Lista fija de responsables — usada también en la normalización al cargar
RESPONSABLES_VALIDOS_GLOBAL = [
    "JURIDICO - JAVIER PRADA",
    "GERALDIN",
    "JAVIER PRADA - GERALDIN",
    "GESTORES SOCIALES",
    "JURIDICO",
    "DIANA - JUAN SEBASTIAN",
    "DIANA - GERALDIN",
    "JAVIER PRADA",
    "VANESSA",
    "GERALDIN - SAC",
    "SUI",
    "DIANA - GERALDIN - JURIDICO",
    "GERALDIN - JURIDICO",
    "GERALDIN - SUI",
]

def _normalizar_col_responsable(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza la columna Responsable al valor válido más cercano."""
    import difflib
    validos_upper = [r.upper() for r in RESPONSABLES_VALIDOS_GLOBAL]

    def _match(valor):
        if pd.isna(valor) or str(valor).strip() == "":
            return valor
        v = str(valor).strip().upper()
        # Exacto
        if v in validos_upper:
            return RESPONSABLES_VALIDOS_GLOBAL[validos_upper.index(v)]
        # Similitud
        matches = difflib.get_close_matches(v, validos_upper, n=1, cutoff=0.4)
        if matches:
            return RESPONSABLES_VALIDOS_GLOBAL[validos_upper.index(matches[0])]
        return valor  # sin match, conservar original

    if "Responsable" in df.columns:
        df = df.copy()
        df["Responsable"] = df["Responsable"].apply(_match)
    return df


@st.cache_data(ttl=60, show_spinner="📥 Descargando datos desde OneDrive…")
def cargar_datos() -> tuple[pd.DataFrame, pd.DataFrame]:
    url  = _drive_url() + "/content"
    resp = requests.get(url, headers=headers(), timeout=30)
    resp.raise_for_status()
    buf  = BytesIO(resp.content)
    df_a = pd.read_excel(buf, sheet_name=SHEET_ABIERTOS, engine="openpyxl")
    buf.seek(0)
    df_c = pd.read_excel(buf, sheet_name=SHEET_CERRADOS, engine="openpyxl")
    # Normalizar responsables en ambas hojas
    df_a = _normalizar_col_responsable(df_a)
    df_c = _normalizar_col_responsable(df_c)
    return df_a, df_c


def guardar_datos(df_abiertos: pd.DataFrame, df_cerrados: pd.DataFrame) -> None:
    """
    Descarga el archivo original, actualiza los datos fila por fila
    preservando el formato de tabla (Table) de cada hoja, y lo sube.
    """
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    # 1 — Descargar el archivo original para preservar formato
    url_content = _drive_url() + "/content"
    resp_get    = requests.get(url_content, headers=headers(), timeout=30)
    resp_get.raise_for_status()
    buf_orig = BytesIO(resp_get.content)

    wb = load_workbook(buf_orig)

    def actualizar_hoja(wb, nombre_hoja: str, df: pd.DataFrame) -> None:
        ws = wb[nombre_hoja]

        # Detectar la tabla Excel en la hoja
        tabla = list(ws.tables.values())

        # Limpiar datos existentes (mantener fila de encabezado)
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.value = None

        # Escribir nuevos datos fila por fila
        for r_idx, (_, row_data) in enumerate(df.iterrows(), start=2):
            for c_idx, value in enumerate(row_data, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        # Actualizar el rango de la tabla para que incluya los nuevos datos
        if tabla:
            t = tabla[0]
            ref_start = t.ref.split(":")[0]  # ej: A1
            col_end   = get_column_letter(len(df.columns))
            row_end   = len(df) + 1          # +1 por el encabezado
            t.ref     = f"{ref_start}:{col_end}{row_end}"

    actualizar_hoja(wb, SHEET_ABIERTOS, df_abiertos)
    actualizar_hoja(wb, SHEET_CERRADOS, df_cerrados)

    # 2 — Guardar en buffer y subir a OneDrive
    buf_out = BytesIO()
    wb.save(buf_out)
    buf_out.seek(0)

    resp_put = requests.put(
        url_content,
        headers={
            **headers(),
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
        data=buf_out.read(),
        timeout=60,
    )
    resp_put.raise_for_status()

# ─────────────────────────────────────────────
# AUTENTICACIÓN — ROLES
# ─────────────────────────────────────────────

def verificar_credenciales(usuario: str, password: str) -> str | None:
    usuarios = st.secrets.get("usuarios", {})
    cfg      = usuarios.get(usuario)
    if cfg and cfg.get("password") == password:
        return cfg.get("rol")
    return None


def es_admin() -> bool:
    return st.session_state.get("rol") == "admin"


def mostrar_login_sidebar() -> None:
    with st.sidebar:
        st.divider()
        if es_admin():
            st.markdown(
                f'<div class="role-admin">🔐 Admin: {st.session_state["usuario"]}</div>',
                unsafe_allow_html=True,
            )
            if st.button("🔌 Cerrar sesión", use_container_width=True):
                st.session_state.pop("rol",     None)
                st.session_state.pop("usuario", None)
                st.rerun()
        else:
            # Login oculto — se despliega al hacer clic en la flecha
            with st.expander("🔑 ›", expanded=False):
                with st.form("form_login", clear_on_submit=True):
                    usr = st.text_input("Usuario")
                    pwd = st.text_input("Contraseña", type="password")
                    ok  = st.form_submit_button("Iniciar sesión", use_container_width=True)
                if ok:
                    rol = verificar_credenciales(usr, pwd)
                    if rol:
                        st.session_state["rol"]     = rol
                        st.session_state["usuario"] = usr
                        st.rerun()
                    else:
                        st.error("Usuario o contraseña incorrectos.")

# ─────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────

# Festivos colombianos fijos (Ley 51/1983 + Ley 270/1996)
# Se actualiza anualmente agregando el año correspondiente
def _festivos_colombia() -> list:
    festivos = []
    for anio in range(2020, 2031):
        # Festivos fijos
        fijos = [
            f"{anio}-01-01",  # Año Nuevo
            f"{anio}-05-01",  # Día del Trabajo
            f"{anio}-07-20",  # Independencia
            f"{anio}-08-07",  # Batalla de Boyacá
            f"{anio}-12-08",  # Inmaculada Concepción
            f"{anio}-12-25",  # Navidad
        ]
        festivos.extend(fijos)

        # Festivos "Ley Emiliani" — se mueven al lunes siguiente si no caen en lunes
        def lunes_siguiente(fecha_str):
            d = pd.Timestamp(fecha_str)
            if d.weekday() == 0:
                return d
            dias = 7 - d.weekday()
            return d + pd.Timedelta(days=dias)

        emiliani = [
            f"{anio}-01-06",  # Reyes Magos
            f"{anio}-03-19",  # San José
            f"{anio}-06-29",  # San Pedro y San Pablo
            f"{anio}-08-15",  # Asunción
            f"{anio}-10-12",  # Día de la Raza
            f"{anio}-11-01",  # Todos los Santos
            f"{anio}-11-11",  # Independencia Cartagena
        ]
        for f in emiliani:
            festivos.append(str(lunes_siguiente(f).date()))

        # Semana Santa (Jueves y Viernes Santo) — cálculo algoritmo de Gauss
        a = anio % 19
        b = anio // 100
        c = anio % 100
        d = b // 4
        e = b % 4
        f = (b + 8) // 25
        g = (b - f + 1) // 3
        h = (19*a + b - d - g + 15) % 30
        i = c // 4
        k = c % 4
        l = (32 + 2*e + 2*i - h - k) % 7
        m = (a + 11*h + 22*l) // 451
        mes  = (h + l - 7*m + 114) // 31
        dia  = ((h + l - 7*m + 114) % 31) + 1
        pascua   = pd.Timestamp(year=anio, month=mes, day=dia)
        jueves   = pascua - pd.Timedelta(days=3)
        viernes  = pascua - pd.Timedelta(days=2)
        ascension = lunes_siguiente(str((pascua + pd.Timedelta(days=39)).date()))
        corpus    = lunes_siguiente(str((pascua + pd.Timedelta(days=60)).date()))
        sagrado   = lunes_siguiente(str((pascua + pd.Timedelta(days=68)).date()))
        festivos.extend([
            str(jueves.date()), str(viernes.date()),
            str(ascension.date()), str(corpus.date()), str(sagrado.date()),
        ])

    return [np.datetime64(f, "D") for f in festivos]

FESTIVOS_CO = np.array(_festivos_colombia(), dtype="datetime64[D]")


def calcular_dias_habiles(df: pd.DataFrame) -> pd.Series:
    """Días hábiles entre FechaCreacion y hoy, excluyendo festivos colombianos."""
    hoy    = np.datetime64(pd.Timestamp.today().normalize(), "D")
    fechas = pd.to_datetime(df["FechaCreacion"], errors="coerce").values.astype("datetime64[D]")
    valido = ~pd.isnull(df["FechaCreacion"])
    dias   = np.where(
        valido,
        np.busday_count(fechas, hoy, holidays=FESTIVOS_CO),
        np.nan,
    )
    return pd.Series(dias, index=df.index)


def aplicar_semaforo(df: pd.DataFrame) -> pd.DataFrame:
    """
    Agrega Dias_Habiles y Semaforo_KPI (para métricas internas).
    La columna Semaforo original del Excel NO se toca — se muestra tal cual en tablas.
    """
    df = df.copy()
    df["FechaCreacion"] = pd.to_datetime(df["FechaCreacion"], errors="coerce")
    df["Dias_Habiles"]  = calcular_dias_habiles(df)
    df["Semaforo_KPI"]  = pd.cut(df["Dias_Habiles"], bins=BINS_SEMAFORO, labels=LABELS_SEMAFORO)
    return df


def a_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()


def validar_columnas(df: pd.DataFrame, nombre: str) -> None:
    faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in df.columns]
    if faltantes:
        st.warning(f"⚠️ **{nombre}** le faltan columnas: `{', '.join(faltantes)}`")

# ─────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────

try:
    df_abiertos, df_cerrados = cargar_datos()
except requests.HTTPError as e:
    st.error(f"❌ Error al conectar con OneDrive: `{e.response.status_code} {e.response.reason}`")
    st.info("""
**Posibles causas:**
- `401 Unauthorized` → client_id, client_secret o tenant_id incorrectos
- `403 Forbidden` → permisos no otorgados o falta "Grant admin consent"
- `404 Not Found` → file_path o user_id incorrecto
    """)
    st.stop()
except Exception as e:
    st.error(f"❌ Error inesperado: `{e}`")
    st.stop()

validar_columnas(df_abiertos, SHEET_ABIERTOS)
validar_columnas(df_cerrados, SHEET_CERRADOS)

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────

mostrar_login_sidebar()

with st.sidebar:
    st.markdown("## ⚡ Filtros")

# ─────────────────────────────────────────────
# NAVEGACIÓN
# ─────────────────────────────────────────────

TABS = ["⚡ Escalamiento", "💡 Casos Gestionados", "🔧 Gestionar"] if es_admin() \
       else ["⚡ Escalamiento", "💡 Casos Gestionados"]

dashboard = st.selectbox("Seleccionar módulo", TABS, label_visibility="collapsed")

rol_html = '<span class="role-admin">🔐 Admin</span>' if es_admin() \
           else '<span class="role-viewer">👁️ Viewer</span>'

st.markdown("""
<style>
.btn-verde > button {
    background: linear-gradient(135deg, #16a34a, #15803d) !important;
    color: white !important; border: none !important;
    border-radius: 8px !important; font-weight: 600 !important;
}
.btn-verde > button:hover { opacity: .85 !important; }
</style>
""", unsafe_allow_html=True)

col_badge, col_btn = st.columns([7, 1])
with col_badge:
    st.markdown(
        f'<div class="badge">🔌 OneDrive · {datetime.now().strftime("%d/%m/%Y %H:%M")} &nbsp;{rol_html}</div>',
        unsafe_allow_html=True,
    )
with col_btn:
    st.markdown('<div class="btn-verde">', unsafe_allow_html=True)
    if st.button("⚡ Actualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
# ═══════════════════════════════════════════════
# MÓDULO: GESTIONAR  (solo admin)
# ═══════════════════════════════════════════════

if dashboard == "🔧 Gestionar":

    if not es_admin():
        st.error("🔒 Acceso restringido. Inicia sesión como administrador.")
        st.stop()

    st.title("🔧 Gestión de Tickets")
    df_edit = df_abiertos.copy()

    with st.expander("🔎 Filtros de búsqueda", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        f_nui         = c1.text_input("NUI")
        f_ticket      = c2.text_input("Id_Tickets")
        f_seccional   = c3.selectbox("Seccional",   ["Todos"] + sorted(df_edit["NombreSeccionales"].dropna().unique().tolist()))
        f_responsable = c4.selectbox("Responsable", ["Todos"] + sorted(df_edit["Responsable"].dropna().unique().tolist()))

    if f_nui:
        df_edit = df_edit[df_edit["NUI"].astype(str).str.contains(f_nui, case=False)]
    if f_ticket and "Id_Tickets" in df_edit.columns:
        df_edit = df_edit[df_edit["Id_Tickets"].astype(str).str.contains(f_ticket, case=False)]
    if f_seccional != "Todos":
        df_edit = df_edit[df_edit["NombreSeccionales"] == f_seccional]
    if f_responsable != "Todos":
        df_edit = df_edit[df_edit["Responsable"] == f_responsable]

    # ── Tabla 1: Gestión SAC ──
    st.subheader("🗂️ Gestión SAC")
    df_cerrados_edit = df_cerrados.copy()
    if f_nui:
        df_cerrados_edit = df_cerrados_edit[df_cerrados_edit["NUI"].astype(str).str.contains(f_nui, case=False)]
    if f_seccional != "Todos" and "NombreSeccionales" in df_cerrados_edit.columns:
        df_cerrados_edit = df_cerrados_edit[df_cerrados_edit["NombreSeccionales"] == f_seccional]
    if f_responsable != "Todos" and "Responsable" in df_cerrados_edit.columns:
        df_cerrados_edit = df_cerrados_edit[df_cerrados_edit["Responsable"] == f_responsable]
    cols_cerrados = [c for c in COLUMNAS_TABLA if c in df_cerrados_edit.columns]
    st.markdown(f"**{len(df_cerrados_edit)}** registros")
    st.dataframe(df_cerrados_edit[cols_cerrados] if cols_cerrados else df_cerrados_edit,
                 use_container_width=True, height=280)

    st.divider()

    # ── Tabla 2: Consolidado ──
    st.subheader("📑 Consolidado")
    st.markdown(f"**{len(df_edit)}** registros encontrados")
    cols_visibles = [c for c in COLUMNAS_TABLA if c in df_edit.columns]
    st.dataframe(df_edit[cols_visibles] if cols_visibles else df_edit,
                 use_container_width=True, height=280)

    if df_edit.empty:
        st.info("No hay registros que coincidan.")
        st.stop()

    st.divider()
    st.subheader("⚙️ Editar registro")

    index_sel = st.selectbox(
        "Selecciona un registro", df_edit.index,
        format_func=lambda i: f"#{i} — NUI: {df_abiertos.loc[i,'NUI']} | Responsable: {df_abiertos.loc[i,'Responsable']}",
    )
    registro = df_abiertos.loc[index_sel]

    responsables_lista = RESPONSABLES_VALIDOS_GLOBAL

    # Opciones fijas de SubMenu1
    OPCIONES_SUBMENU1 = [
        "Reposición",
        "Sin SISFV Instalada",
        "Restricciones Acceso",
        "Usuario Ausente",
        "Traslados No Autorizado",
        "Suspension del servicio",
        "Reconexion del servicio",
        "SISFV Funcional",
        "Ausencia de Partes",
        "SISFV Abandonado",
        "Equipos modificados",
    ]

    with st.form("form_edicion"):
        # Campos de solo lectura (informativos)
        col1, col2 = st.columns(2)
        col1.text_input("NUI",      registro.get("NUI", ""),           disabled=True)
        col2.text_input("Semáforo", str(registro.get("Semaforo", "")), disabled=True)

        # Responsable: lista desplegable con los existentes en Consolidado
        responsable_actual = registro.get("Responsable", "")
        idx_resp = responsables_lista.index(responsable_actual) if responsable_actual in responsables_lista else 0
        nuevo_responsable = st.selectbox("Responsable", responsables_lista, index=idx_resp)

        # SubMenu1: lista fija de opciones
        submenu_actual = str(registro.get("SubMenu1", ""))
        idx_sub = OPCIONES_SUBMENU1.index(submenu_actual) if submenu_actual in OPCIONES_SUBMENU1 else 0
        nuevo_submenu1 = st.selectbox("SubMenu1", OPCIONES_SUBMENU1, index=idx_sub)

        # Descripción editable
        nueva_descripcion = st.text_area("Descripción", registro.get("Descripción", ""), height=120)

        col_a, col_b, _  = st.columns([1, 1, 3])
        guardar  = col_a.form_submit_button("✅ Guardar cambios")
        eliminar = col_b.form_submit_button("🗑️ Eliminar registro", type="secondary")

    if guardar:
        df_abiertos.loc[index_sel, "Responsable"] = nuevo_responsable
        df_abiertos.loc[index_sel, "SubMenu1"]    = nuevo_submenu1
        df_abiertos.loc[index_sel, "Descripción"] = nueva_descripcion
        try:
            with st.spinner("Guardando en OneDrive…"):
                guardar_datos(df_abiertos, df_cerrados)
            st.success("✅ Cambios guardados en OneDrive.")
        except Exception as e:
            st.error(f"❌ Error al guardar: `{e}`")
        st.cache_data.clear()
        st.rerun()

    if eliminar:
        if st.session_state.get("confirm_delete") != index_sel:
            st.session_state["confirm_delete"] = index_sel
            st.warning("⚠️ Presiona **Eliminar registro** de nuevo para confirmar.")
        else:
            df_abiertos.drop(index=index_sel, inplace=True)
            try:
                with st.spinner("Guardando cambios…"):
                    guardar_datos(df_abiertos, df_cerrados)
                st.warning("🗑️ Registro eliminado correctamente.")
            except Exception as e:
                st.error(f"❌ Error al eliminar: `{e}`")
            st.session_state.pop("confirm_delete", None)
            st.cache_data.clear()
            st.rerun()

    st.stop()

# ═══════════════════════════════════════════════
# MÓDULO: ESCALAMIENTO / CASOS CERRADOS
# ═══════════════════════════════════════════════

es_escalamiento = dashboard == "⚡ Escalamiento"
df_base = df_abiertos.copy() if es_escalamiento else df_cerrados.copy()
df_base = aplicar_semaforo(df_base)

with st.sidebar:
    responsables = st.multiselect("Responsable", sorted(df_base["Responsable"].dropna().unique()))
    seccionales  = st.multiselect("Seccional",   sorted(df_base["NombreSeccionales"].dropna().unique()))

    menus     = st.multiselect("Menú",     sorted(df_base["Menu"].dropna().unique())     if "Menu"     in df_base.columns else [])
    submenus1 = st.multiselect("SubMenu1", sorted(df_base["SubMenu1"].dropna().unique()) if "SubMenu1" in df_base.columns else [])
    submenus2 = st.multiselect("SubMenu2", sorted(df_base["SubMenu2"].dropna().unique()) if "SubMenu2" in df_base.columns else [])
    submenus3 = st.multiselect("SubMenu3", sorted(df_base["SubMenu3"].dropna().unique()) if "SubMenu3" in df_base.columns else [])

    fecha_min, fecha_max = df_base["FechaCreacion"].min(), df_base["FechaCreacion"].max()
    if pd.notna(fecha_min) and pd.notna(fecha_max) and fecha_min != fecha_max:
        rango_fechas = st.date_input(
            "Rango de fechas",
            value=(fecha_min.date(), fecha_max.date()),
            min_value=fecha_min.date(), max_value=fecha_max.date(),
        )
    else:
        rango_fechas = None

df = df_base.copy()
if responsables:
    df = df[df["Responsable"].isin(responsables)]
if seccionales:
    df = df[df["NombreSeccionales"].isin(seccionales)]
if menus and "Menu" in df.columns:
    df = df[df["Menu"].isin(menus)]
if submenus1 and "SubMenu1" in df.columns:
    df = df[df["SubMenu1"].isin(submenus1)]
if submenus2 and "SubMenu2" in df.columns:
    df = df[df["SubMenu2"].isin(submenus2)]
if submenus3 and "SubMenu3" in df.columns:
    df = df[df["SubMenu3"].isin(submenus3)]
if rango_fechas and len(rango_fechas) == 2:
    df = df[(df["FechaCreacion"].dt.date >= rango_fechas[0]) &
            (df["FechaCreacion"].dt.date <= rango_fechas[1])]

# KPIs
titulo = "⚡ Gestión Escalamiento" if es_escalamiento else "💡 Gestión Casos Gestionados"
st.title(titulo)

total     = len(df)
secciones = df["NombreSeccionales"].nunique()
nuis      = df["NUI"].nunique()
prom_dias = round(df["Dias_Habiles"].mean(), 1) if not df.empty else 0
vencidos  = (df["Semaforo_KPI"] == "🔴 Vencido").sum()
en_riesgo = (df["Semaforo_KPI"] == "🟡 En riesgo").sum()

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Total Tickets", total)
k2.metric("Seccionales",   secciones)
k3.metric("NUIs únicos",   nuis)
k4.metric("Promedio días", prom_dias)
k5.metric("🔴 Vencidos",   vencidos,  delta=f"-{vencidos}"  if vencidos  else None, delta_color="inverse")
k6.metric("🟡 En riesgo",  en_riesgo, delta=f"-{en_riesgo}" if en_riesgo else None, delta_color="inverse")

st.divider()

# ── TABLA PRIMERO ──
st.subheader("⚡ Detalle de tickets")
busqueda_top = st.text_input("🔎 Búsqueda rápida", "", key="busqueda_top")
df_tabla_top = df.copy()
if busqueda_top:
    mask = df_tabla_top.astype(str).apply(lambda col: col.str.contains(busqueda_top, case=False)).any(axis=1)
    df_tabla_top = df_tabla_top[mask]
cols_top = [c for c in COLUMNAS_TABLA if c in df_tabla_top.columns]
df_top   = df_tabla_top[cols_top].copy() if cols_top else df_tabla_top.copy()

# Formatear columnas de fecha a dd/mm/yyyy
COLS_FECHA = ["FechaCreacion", "Fecha Asignación", "Fecha Respuesta"]
for col_f in COLS_FECHA:
    if col_f in df_top.columns:
        df_top[col_f] = pd.to_datetime(df_top[col_f], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")

st.dataframe(
    df_top.style.apply(lambda row: [
        "background-color: #1a2e1a" if row.get("Semaforo_KPI") == "🟢 En tiempo"
        else "background-color: #2e2a1a" if row.get("Semaforo_KPI") == "🟡 En riesgo"
        else "background-color: #2e1a1a" if row.get("Semaforo_KPI") == "🔴 Vencido"
        else "" for _ in row], axis=1),
    use_container_width=True, height=380,
)
cols_export = [c for c in COLUMNAS_TABLA if c in df_tabla_top.columns]
st.download_button(
    label="📥 Exportar tabla a Excel",
    data=a_excel(df_tabla_top[cols_export]),
    file_name=f"SAC_export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="export_top",
)

st.divider()

LAYOUT = dict(plot_bgcolor="#161a24", paper_bgcolor="#161a24", margin=dict(l=0,r=0,t=30,b=0),
              legend=dict(bgcolor="rgba(0,0,0,0)", font_color="#d1d5db"),
              font=dict(color="#d1d5db"))

# ── FILA 1: Seccional + Semáforo ──
col_g1, col_g2 = st.columns([2, 1])
with col_g1:
    st.subheader("🏘️ Tickets por Seccional")
    df_sec = df.groupby("NombreSeccionales").size().reset_index(name="Tickets").sort_values("Tickets", ascending=True)
    fig_bar = px.bar(df_sec, x="Tickets", y="NombreSeccionales", orientation="h",
                     color="NombreSeccionales", color_discrete_sequence=PALETA, template="plotly_dark")
    fig_bar.update_layout(**LAYOUT, showlegend=False, yaxis_title=None)
    st.plotly_chart(fig_bar, use_container_width=True, key="fig_bar")

with col_g2:
    st.subheader("🚦 Estado del Servicio")
    df_sem = df["Semaforo_KPI"].value_counts().reset_index()
    df_sem.columns = ["Estado", "Cantidad"]
    fig_pie = px.pie(df_sem, names="Estado", values="Cantidad", hole=0.55, template="plotly_dark",
                     color="Estado", color_discrete_map={"🟢 En tiempo": "#4ade80", "🟡 En riesgo": "#facc15", "🔴 Vencido": "#f87171"})
    fig_pie.update_layout(**{**LAYOUT, "legend": dict(orientation="h", y=-0.15, bgcolor="rgba(0,0,0,0)", font_color="#d1d5db")})
    fig_pie.update_traces(textinfo="percent+label", textfont_size=12)
    st.plotly_chart(fig_pie, use_container_width=True, key="fig_pie")

# ── FILA 3: SubMenu1 — ancho completo ──
st.subheader("⚡ Tickets por Categoría")
if "SubMenu1" in df.columns:
    df_sub = (df.groupby("SubMenu1").size().reset_index(name="Tickets")
                .sort_values("Tickets", ascending=True))
    fig_sub = px.bar(df_sub, x="Tickets", y="SubMenu1", orientation="h",
                     color="SubMenu1", color_discrete_sequence=PALETA, template="plotly_dark")
    fig_sub.update_layout(**LAYOUT, showlegend=False, yaxis_title=None)
    st.plotly_chart(fig_sub, use_container_width=True, key="fig_sub")
else:
    st.info("Columna SubMenu1 no disponible.")

# ── FILA 4: SubMenu2 ← izq · SubMenu3 → der ──
col_g5, col_g6 = st.columns(2)
with col_g5:
    st.subheader("🔌 Tickets por SubCategoría")
    if "SubMenu2" in df.columns and df["SubMenu2"].notna().any():
        df_sub2 = (df.groupby("SubMenu2").size().reset_index(name="Tickets")
                     .sort_values("Tickets", ascending=True))
        fig_sub2 = px.bar(df_sub2, x="Tickets", y="SubMenu2", orientation="h",
                          color="SubMenu2", color_discrete_sequence=PALETA, template="plotly_dark")
        fig_sub2.update_layout(**LAYOUT, showlegend=False, yaxis_title=None)
        st.plotly_chart(fig_sub2, use_container_width=True, key="fig_sub2")
    else:
        st.info("Columna SubMenu2 no disponible o sin datos.")

with col_g6:
    st.subheader("💡 Tickets por Detalle")
    if "SubMenu3" in df.columns and df["SubMenu3"].notna().any():
        df_sub3 = (df.groupby("SubMenu3").size().reset_index(name="Tickets")
                     .sort_values("Tickets", ascending=True))
        fig_sub3 = px.bar(df_sub3, x="Tickets", y="SubMenu3", orientation="h",
                          color="SubMenu3", color_discrete_sequence=PALETA, template="plotly_dark")
        fig_sub3.update_layout(**LAYOUT, showlegend=False, yaxis_title=None)
        st.plotly_chart(fig_sub3, use_container_width=True, key="fig_sub3")
    else:
        st.info("Columna SubMenu3 no disponible o sin datos.")

# ── FILA 5: Top 10 Responsables ← izq · Tiempo promedio cierre → der ──
col_g7, col_g8 = st.columns(2)
with col_g7:
    st.subheader("👷 Top 10 Responsables con más Tickets")
    if "Responsable" in df.columns:
        df_resp = (df.groupby("Responsable").size().reset_index(name="Tickets")
                     .sort_values("Tickets", ascending=False).head(10)
                     .sort_values("Tickets", ascending=True))
        fig_resp = px.bar(df_resp, x="Tickets", y="Responsable", orientation="h",
                          color="Responsable", color_discrete_sequence=PALETA, template="plotly_dark")
        fig_resp.update_layout(**LAYOUT, showlegend=False, yaxis_title=None)
        st.plotly_chart(fig_resp, use_container_width=True, key="fig_resp")

with col_g8:
    st.subheader("⏱️ Tiempo Promedio de Atención por Responsable")
    if "Responsable" in df.columns and df["Dias_Habiles"].notna().any():
        df_cierre = (df.groupby("Responsable")["Dias_Habiles"].mean().reset_index()
                       .rename(columns={"Dias_Habiles": "Dias para Cierre"})
                       .sort_values("Dias para Cierre", ascending=True))
        fig_cierre = px.bar(df_cierre, x="Dias para Cierre", y="Responsable", orientation="h",
                            color="Responsable", color_discrete_sequence=PALETA, template="plotly_dark",
                            labels={"Dias para Cierre": "Días promedio"})
        fig_cierre.update_layout(**LAYOUT, showlegend=False, yaxis_title=None)
        st.plotly_chart(fig_cierre, use_container_width=True, key="fig_cierre")

# ── FILA 2: Tendencia + Treemap ──
col_g3, col_g4 = st.columns(2)
with col_g3:
    st.subheader("📈 Tendencia de creación")
    if df["FechaCreacion"].notna().any():
        df_trend = (df.dropna(subset=["FechaCreacion"])
                      .assign(Mes=lambda x: x["FechaCreacion"].dt.to_period("M").astype(str))
                      .groupby("Mes").size().reset_index(name="Tickets"))
        fig_line = px.line(df_trend, x="Mes", y="Tickets", markers=True,
                           template="plotly_dark", color_discrete_sequence=["#a78bfa"])
        fig_line.update_layout(**LAYOUT)
        fig_line.update_traces(line_width=2.5, marker_size=7, line_color="#a78bfa",
                               marker=dict(color="#f472b6", size=8))
        st.plotly_chart(fig_line, use_container_width=True, key="fig_line")
    else:
        st.info("Sin datos de fecha disponibles.")

with col_g4:
    st.subheader("🗺️ Cobertura Seccional × Responsable")
    df_tree = df.groupby(["NombreSeccionales", "Responsable"]).size().reset_index(name="Tickets")
    fig_tree = px.treemap(df_tree, path=["NombreSeccionales", "Responsable"], values="Tickets",
                          color="Tickets", color_continuous_scale=["#3a81d5", "#8f5cda", "#f472b6"],
                          template="plotly_dark")
    fig_tree.update_layout(**LAYOUT)
    st.plotly_chart(fig_tree, use_container_width=True, key="fig_tree")

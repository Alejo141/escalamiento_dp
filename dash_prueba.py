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


@st.cache_data(ttl=60, show_spinner="📥 Descargando datos desde OneDrive…")
def cargar_datos() -> tuple[pd.DataFrame, pd.DataFrame]:
    url  = _drive_url() + "/content"
    resp = requests.get(url, headers=headers(), timeout=30)
    resp.raise_for_status()
    buf  = BytesIO(resp.content)
    df_a = pd.read_excel(buf, sheet_name=SHEET_ABIERTOS, engine="openpyxl")
    buf.seek(0)
    df_c = pd.read_excel(buf, sheet_name=SHEET_CERRADOS, engine="openpyxl")
    return df_a, df_c


def guardar_datos(df_abiertos: pd.DataFrame, df_cerrados: pd.DataFrame) -> None:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_abiertos.to_excel(w, index=False, sheet_name=SHEET_ABIERTOS)
        df_cerrados.to_excel(w, index=False, sheet_name=SHEET_CERRADOS)
    buf.seek(0)
    url  = _drive_url() + "/content"
    resp = requests.put(
        url,
        headers={
            **headers(),
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
        data=buf.read(),
        timeout=60,
    )
    resp.raise_for_status()

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
            if st.button("🚪 Cerrar sesión", use_container_width=True):
                st.session_state.pop("rol",     None)
                st.session_state.pop("usuario", None)
                st.rerun()
        else:
            st.markdown("### 🔐 Acceso Admin")
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

def calcular_dias_habiles(df: pd.DataFrame) -> pd.Series:
    hoy    = np.datetime64(pd.Timestamp.today().normalize(), "D")
    fechas = pd.to_datetime(df["FechaCreacion"], errors="coerce").values.astype("datetime64[D]")
    valido = ~pd.isnull(df["FechaCreacion"])
    return pd.Series(np.where(valido, np.busday_count(fechas, hoy), np.nan), index=df.index)


def aplicar_semaforo(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["FechaCreacion"] = pd.to_datetime(df["FechaCreacion"], errors="coerce")
    df["Dias_Habiles"]  = calcular_dias_habiles(df)
    df["Semaforo"]      = pd.cut(df["Dias_Habiles"], bins=BINS_SEMAFORO, labels=LABELS_SEMAFORO)
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
    st.markdown("## 🔎 Filtros")

# ─────────────────────────────────────────────
# NAVEGACIÓN
# ─────────────────────────────────────────────

TABS = ["📈 Escalamiento", "✅ Casos Cerrados", "🛠️ Gestionar"] if es_admin() \
       else ["📈 Escalamiento", "✅ Casos Cerrados"]

dashboard = st.selectbox("Seleccionar módulo", TABS, label_visibility="collapsed")

rol_html = '<span class="role-admin">🔐 Admin</span>' if es_admin() \
           else '<span class="role-viewer">👁️ Viewer</span>'
st.markdown(
    f'<div class="badge">☁️ OneDrive · {datetime.now().strftime("%d/%m/%Y %H:%M")} &nbsp;{rol_html}</div>',
    unsafe_allow_html=True,
)

# ═══════════════════════════════════════════════
# MÓDULO: GESTIONAR  (solo admin)
# ═══════════════════════════════════════════════

if dashboard == "🛠️ Gestionar":

    if not es_admin():
        st.error("🔒 Acceso restringido. Inicia sesión como administrador.")
        st.stop()

    st.title("🛠️ Gestión de Tickets")
    df_edit = df_abiertos.copy()

    with st.expander("🔍 Filtros de búsqueda", expanded=True):
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

    st.markdown(f"**{len(df_edit)}** registros encontrados")
    st.dataframe(df_edit, use_container_width=True, height=300)

    if df_edit.empty:
        st.info("No hay registros que coincidan.")
        st.stop()

    st.divider()
    st.subheader("✏️ Editar registro")

    index_sel = st.selectbox(
        "Selecciona un registro", df_edit.index,
        format_func=lambda i: f"#{i} — NUI: {df_abiertos.loc[i,'NUI']} | Responsable: {df_abiertos.loc[i,'Responsable']}",
    )
    registro = df_abiertos.loc[index_sel]

    with st.form("form_edicion"):
        col1, col2 = st.columns(2)
        nuevo_responsable = col1.text_input("Responsable", registro.get("Responsable", ""))
        nuevo_semaforo    = col2.selectbox(
            "Semáforo", LABELS_SEMAFORO,
            index=LABELS_SEMAFORO.index(registro["Semaforo"])
                  if registro.get("Semaforo") in LABELS_SEMAFORO else 0,
        )
        nueva_descripcion = st.text_area("Descripción", registro.get("Descripción", ""), height=120)
        col_a, col_b, _  = st.columns([1, 1, 3])
        guardar  = col_a.form_submit_button("💾 Guardar cambios")
        eliminar = col_b.form_submit_button("🗑️ Eliminar registro", type="secondary")

    if guardar:
        df_abiertos.loc[index_sel, "Responsable"] = nuevo_responsable
        df_abiertos.loc[index_sel, "Semaforo"]    = nuevo_semaforo
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

es_escalamiento = dashboard == "📈 Escalamiento"
df_base = df_abiertos.copy() if es_escalamiento else df_cerrados.copy()
df_base = aplicar_semaforo(df_base)

with st.sidebar:
    responsables = st.multiselect("Responsable", sorted(df_base["Responsable"].dropna().unique()))
    seccionales  = st.multiselect("Seccional",   sorted(df_base["NombreSeccionales"].dropna().unique()))
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
if rango_fechas and len(rango_fechas) == 2:
    df = df[(df["FechaCreacion"].dt.date >= rango_fechas[0]) &
            (df["FechaCreacion"].dt.date <= rango_fechas[1])]

# KPIs
titulo = "📈 Gestión Escalamiento" if es_escalamiento else "✅ Gestión Casos Cerrados"
st.title(titulo)

total     = len(df)
secciones = df["NombreSeccionales"].nunique()
nuis      = df["NUI"].nunique()
prom_dias = round(df["Dias_Habiles"].mean(), 1) if not df.empty else 0
vencidos  = (df["Semaforo"] == "🔴 Vencido").sum()
en_riesgo = (df["Semaforo"] == "🟡 En riesgo").sum()

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Total Tickets", total)
k2.metric("Seccionales",   secciones)
k3.metric("NUIs únicos",   nuis)
k4.metric("Promedio días", prom_dias)
k5.metric("🔴 Vencidos",   vencidos,  delta=f"-{vencidos}"  if vencidos  else None, delta_color="inverse")
k6.metric("🟡 En riesgo",  en_riesgo, delta=f"-{en_riesgo}" if en_riesgo else None, delta_color="inverse")

st.divider()

# Gráficos fila 1
col_g1, col_g2 = st.columns([2, 1])
with col_g1:
    st.subheader("📊 Tickets por Seccional")
    df_sec = df.groupby("NombreSeccionales").size().reset_index(name="Tickets").sort_values("Tickets", ascending=True)
    fig_bar = px.bar(df_sec, x="Tickets", y="NombreSeccionales", orientation="h",
                     color="Tickets", color_continuous_scale=["#3a81d5","#8f5cda"], template="plotly_dark")
    fig_bar.update_layout(plot_bgcolor="#161a24", paper_bgcolor="#161a24",
                          coloraxis_showscale=False, margin=dict(l=0,r=0,t=10,b=0), yaxis_title=None)
    st.plotly_chart(fig_bar, use_container_width=True)

with col_g2:
    st.subheader("🚦 Semáforo")
    df_sem = df["Semaforo"].value_counts().reset_index()
    df_sem.columns = ["Estado","Cantidad"]
    fig_pie = px.pie(df_sem, names="Estado", values="Cantidad", hole=0.55, template="plotly_dark",
                     color="Estado", color_discrete_map={"🟢 En tiempo":"#4ade80","🟡 En riesgo":"#facc15","🔴 Vencido":"#f87171"})
    fig_pie.update_layout(plot_bgcolor="#161a24", paper_bgcolor="#161a24",
                          legend=dict(orientation="h", y=-0.15), margin=dict(l=0,r=0,t=10,b=0))
    fig_pie.update_traces(textinfo="percent+label", textfont_size=12)
    st.plotly_chart(fig_pie, use_container_width=True)

# Gráficos fila 2
col_g3, col_g4 = st.columns(2)
with col_g3:
    st.subheader("📅 Tendencia de creación")
    if df["FechaCreacion"].notna().any():
        df_trend = (df.dropna(subset=["FechaCreacion"])
                      .assign(Mes=lambda x: x["FechaCreacion"].dt.to_period("M").astype(str))
                      .groupby("Mes").size().reset_index(name="Tickets"))
        fig_line = px.line(df_trend, x="Mes", y="Tickets", markers=True,
                           template="plotly_dark", color_discrete_sequence=["#8f5cda"])
        fig_line.update_layout(plot_bgcolor="#161a24", paper_bgcolor="#161a24", margin=dict(l=0,r=0,t=10,b=0))
        fig_line.update_traces(line_width=2.5, marker_size=7)
        st.plotly_chart(fig_line, use_container_width=True)
    else:
        st.info("Sin datos de fecha disponibles.")

with col_g4:
    st.subheader("🗺️ Treemap Seccional × Responsable")
    df_tree = df.groupby(["NombreSeccionales","Responsable"]).size().reset_index(name="Tickets")
    fig_tree = px.treemap(df_tree, path=["NombreSeccionales","Responsable"], values="Tickets",
                          color="Tickets", color_continuous_scale=["#3a81d5","#8f5cda","#f87171"],
                          template="plotly_dark")
    fig_tree.update_layout(paper_bgcolor="#161a24", margin=dict(l=0,r=0,t=10,b=0))
    st.plotly_chart(fig_tree, use_container_width=True)

# Gráfico fila 3
st.subheader("📦 Distribución de días hábiles")
fig_hist = px.histogram(df.dropna(subset=["Dias_Habiles"]), x="Dias_Habiles", nbins=20,
                        color_discrete_sequence=["#7069d8"], template="plotly_dark",
                        labels={"Dias_Habiles":"Días hábiles"})
fig_hist.add_vline(x=5,  line_dash="dash", line_color="#4ade80", annotation_text="Límite verde (5)")
fig_hist.add_vline(x=10, line_dash="dash", line_color="#facc15", annotation_text="Límite amarillo (10)")
fig_hist.update_layout(plot_bgcolor="#161a24", paper_bgcolor="#161a24", margin=dict(l=0,r=0,t=10,b=0))
st.plotly_chart(fig_hist, use_container_width=True)

# Tabla detalle
st.divider()
st.subheader("📋 Detalle de tickets")
busqueda = st.text_input("🔍 Búsqueda rápida", "")
df_tabla = df.copy()
if busqueda:
    mask     = df_tabla.astype(str).apply(lambda col: col.str.contains(busqueda, case=False)).any(axis=1)
    df_tabla = df_tabla[mask]

st.dataframe(
    df_tabla.style.apply(lambda row: [
        "background-color: #1a2e1a" if row.get("Semaforo") == "🟢 En tiempo"
        else "background-color: #2e2a1a" if row.get("Semaforo") == "🟡 En riesgo"
        else "background-color: #2e1a1a" if row.get("Semaforo") == "🔴 Vencido"
        else "" for _ in row], axis=1),
    use_container_width=True, height=400,
)

st.download_button(
    label="⬇️ Exportar vista actual a Excel",
    data=a_excel(df_tabla),
    file_name=f"SAC_export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

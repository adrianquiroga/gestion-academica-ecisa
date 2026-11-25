"""
Portal Principal - Sistema de Gestión Académica (Docentes, Matrículas y Homologaciones).
Muestra la bienvenida, verifica la conectividad con Supabase y el puente de Google Drive,
y facilita la navegación hacia los módulos de la aplicación.
"""

import importlib
import streamlit as st
from src.db.client import probar_conexion_supabase, verificar_configuracion_supabase
from src.services.drive_svc import verificar_configuracion_drive
import src.utils.auth as _auth_mod

importlib.reload(_auth_mod)
renderizar_barra_sesion_y_login = _auth_mod.renderizar_barra_sesion_y_login

st.set_page_config(
    page_title="Sistema de Gestión Académica",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🎓 Sistema de Gestión Académica")
st.subheader("Plataforma Integral de Docentes, Matrículas y Homologaciones")

if not renderizar_barra_sesion_y_login():
    st.stop()

st.markdown(
    """
    Bienvenido al portal institucional de gestión académica. Esta plataforma permite el registro y
    actualización de perfiles docentes con carga directa de evidencias y certificados hacia **Google Drive Institucional**,
    así como la generación de reportes ejecutivos y fichas individuales en formato **PDF**.
    """
)

st.divider()

# =============================================================================
# PANEL DE DIAGNÓSTICO DE INFRAESTRUCTURA (SUPABASE + GOOGLE DRIVE BRIDGE)
# =============================================================================
st.markdown("### 🔌 Estado de Conexión e Infraestructura")

col_db, col_drive = st.columns(2)

with col_db:
    st.markdown("#### 🗄️ Base de Datos (Supabase / PostgreSQL)")
    ok_cfg_db, msg_cfg_db = verificar_configuracion_supabase()
    if ok_cfg_db:
        ok_conn_db, msg_conn_db = probar_conexion_supabase()
        if ok_conn_db:
            st.success(f"**Conectado:** {msg_conn_db}")
        else:
            st.warning(f"**Configurado con advertencia:** {msg_conn_db}")
    else:
        st.error(f"**Pendiente de configuración:** {msg_cfg_db}")
        st.caption(
            "Copia `.streamlit/secrets.toml.example` a `.streamlit/secrets.toml` y agrega `SUPABASE_URL` y `SUPABASE_KEY`."
        )

with col_drive:
    st.markdown("#### 📁 Puente de Soportes (Google Drive + Apps Script)")
    ok_drive, msg_drive = verificar_configuracion_drive()
    if ok_drive:
        st.success(f"**Listo para recibir soportes:** {msg_drive}")
    else:
        st.error(f"**Pendiente de configuración:** {msg_drive}")
        st.caption(
            "Publica `scripts/google_apps_script_bridge.gs` como Aplicación Web en tu Google Drive y configura `APPS_SCRIPT_URL` y `DRIVE_API_TOKEN`."
        )

st.divider()

# =============================================================================
# TARJETAS DE MÓDULOS Y ACCESO RÁPIDO
# =============================================================================
st.markdown("### 🧭 Módulos del Sistema")

c1, c2, c3, c4, c5 = st.columns(5)

with c1:
    with st.container(border=True):
        st.markdown("#### 👨‍🏫 1. Registro Docente")
        st.caption("Fase 1 (Activo)")
        st.write(
            "Búsqueda de docentes, resoluciones UNAD, actualización de nivel de inglés (MCER), "
            "títulos universitarios y aportes con soportes en Drive."
        )
        st.page_link("pages/1_Registro_Docente.py", label="Ir a Registro Docente", icon="📝")

with c2:
    with st.container(border=True):
        st.markdown("#### 📊 2. Panel y PDF")
        st.caption("Fase 1 (Activo)")
        st.write(
            "Indicadores gerenciales de docentes, filtros por centro, programa y nivel de inglés, "
            "consulta de soportes en Drive y descarga de PDF."
        )
        st.page_link("pages/2_Panel_y_PDF.py", label="Ir a Panel Docente", icon="📄")

with c3:
    with st.container(border=True):
        st.markdown("#### 📈 3. Matrículas")
        st.caption("Fase 2 (Activo - 15.473 actas)")
        st.write(
            "Analítica comparativa de Actas de Matrícula 2025 y 2026 por periodo, centro, "
            "programa, condición (Antiguos/Nuevos) y estrato."
        )
        st.page_link("pages/3_Matriculas.py", label="Ir a Matrículas 2025-2026", icon="🗂️")

with c4:
    with st.container(border=True):
        st.markdown("#### 🔄 4. Homologaciones")
        st.caption("Fase 3 (Preparado)")
        st.write(
            "Gestión y trazabilidad de homologaciones internas y externas con soporte "
            "documental de resoluciones y sábanas de notas."
        )
        st.page_link("pages/4_Homologaciones.py", label="Ver Homologaciones", icon="⚖️")

with c5:
    with st.container(border=True):
        st.markdown("#### 🎯 5. Créditos ECISA")
        st.caption("Nuevo (Proyectado vs 2026)")
        st.write(
            "Comparativo de créditos proyectados (`proy_zon`) vs realizados (`creditos`) en 2026 "
            "con gráficas por programa (`pro_descripcion`) y CEAD (`nombrecead`)."
        )
        st.page_link("pages/5_Proyeccion_Creditos.py", label="Ir a Créditos 2026", icon="🎯")

st.divider()
st.caption(
    "Arquitectura optimizada para Streamlit Community Cloud (límite 1 GB RAM), "
    "Supabase PostgreSQL y Google Drive Institucional (archivos PDF/PNG/JPG máx. 5 MB)."
)

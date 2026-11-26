"""
Módulo 3 (Fase 3): Gestión de Homologaciones Internas y Externas.
Esqueleto base preparado con la tabla `public.homologaciones` (incluyendo `soporte_url`) lista en `src/db/schema.sql`.
"""

import importlib
from pathlib import Path
import sys

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import streamlit as st
import src.utils.auth as _auth_mod

importlib.reload(_auth_mod)
renderizar_barra_sesion_y_login = _auth_mod.renderizar_barra_sesion_y_login

st.set_page_config(
    page_title="Módulo de Homologaciones - Fase 3",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded",
)

if not renderizar_barra_sesion_y_login("4_homologaciones"):
    st.stop()

st.title("🔄 Módulo de Homologaciones Internas y Externas")

st.info("🚧 **Módulo en construcción - Fase 2/3**")

with st.container(border=True):
    st.markdown("### Especificaciones Técnicas Preparadas en este Proyecto")
    st.markdown(
        """
        * **Tabla `public.homologaciones`** creada en `src/db/schema.sql` con soporte para homologaciones `Interna` y `Externa`, control de estados (`En estudio`, `Aprobada`, `Rechazada`) y columna `soporte_url`.
        * **Puente con Google Drive Institucional (`src/services/drive_svc.py`)** listo para almacenar resoluciones y sábanas de notas en la subcarpeta `Homologaciones`.
        """
    )

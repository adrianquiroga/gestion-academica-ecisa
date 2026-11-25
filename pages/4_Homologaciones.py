"""
Módulo 3 (Fase 3): Gestión de Homologaciones Internas y Externas.
Esqueleto base preparado con la tabla `public.homologaciones` (incluyendo `soporte_url`) lista en `src/db/schema.sql`.
"""

import streamlit as st

st.set_page_config(
    page_title="Módulo de Homologaciones - Fase 3",
    page_icon="🔄",
    layout="wide",
)

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

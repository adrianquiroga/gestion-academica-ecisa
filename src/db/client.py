"""
Módulo de conexión singleton a Supabase.
Utiliza @st.cache_resource y st.secrets para garantizar una única instancia segura.
"""

from typing import Optional, Tuple
import streamlit as st
from supabase import create_client, Client


def verificar_configuracion_supabase() -> Tuple[bool, str]:
    """
    Verifica si las credenciales SUPABASE_URL y SUPABASE_KEY están configuradas
    en `.streamlit/secrets.toml` y no contienen los valores de plantilla por defecto.
    """
    try:
        url = st.secrets.get("SUPABASE_URL", "")
        key = st.secrets.get("SUPABASE_KEY", "")
    except Exception:
        return False, "No se encontró el archivo `.streamlit/secrets.toml` o no se pudo leer."

    if not url or not key:
        return False, "Faltan las variables `SUPABASE_URL` o `SUPABASE_KEY` en `.streamlit/secrets.toml`."

    if "tu-proyecto-id.supabase.co" in url or key.endswith("..."):
        return False, "Las credenciales de Supabase tienen valores de ejemplo. Actualízalas en `.streamlit/secrets.toml`."

    return True, "Credenciales de Supabase configuradas correctamente."


@st.cache_resource(show_spinner=False)
def get_supabase_client() -> Optional[Client]:
    """
    Retorna la instancia singleton del cliente de Supabase.
    Si la configuración no es válida, retorna None sin lanzar excepciones fatales en la UI.
    """
    ok, _ = verificar_configuracion_supabase()
    if not ok:
        return None

    url: str = st.secrets["SUPABASE_URL"].strip()
    key: str = st.secrets["SUPABASE_KEY"].strip()
    return create_client(url, key)


def probar_conexion_supabase() -> Tuple[bool, str]:
    """
    Realiza una consulta ligera (HEAD/COUNT de 1 registro) a la tabla `docentes`
    para comprobar conectividad y existencia del esquema en Supabase.
    """
    ok_cfg, msg_cfg = verificar_configuracion_supabase()
    if not ok_cfg:
        return False, msg_cfg

    try:
        client = get_supabase_client()
        if client is None:
            return False, "No fue posible inicializar el cliente de Supabase."
        client.table("docentes").select("id", count="exact").limit(1).execute()
        return True, "Conectado a Supabase y esquema `docentes` verificado."
    except Exception as exc:
        return False, f"Error al conectar con Supabase o tabla `docentes` no creada aún: {exc}"

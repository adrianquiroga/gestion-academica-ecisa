"""
Módulo de Autenticación y Control de Acceso Basado en Roles (RBAC):
1. Rol `admin`:
   - Ingresa con usuario `admin` y clave configurada en `secrets.toml` (`ADMIN_PASS`, por defecto `admin2026`).
   - Único autorizado para consultar el **Directorio de Docentes y Resoluciones (101 resultados)** en `2_Panel_y_PDF.py`,
     así como para buscar y editar cualquier docente o ejecutar sincronizaciones masivas desde Excel.
2. Rol `docente`:
   - Ingresa inicialmente con su **número de cédula (documento)** tanto en Usuario como en Contraseña.
   - Puede cambiar su contraseña en cualquier momento (se almacena cifrada con SHA-256 en `docentes.password_hash`,
     con respaldo local automático si la columna SQL aún no ha sido agregada).
   - Solo puede ver y editar su propio perfil, estudios, nivel de inglés, aportes y descargar su propia Ficha PDF.
"""

import hashlib
import json
from pathlib import Path
import runpy
import sys
from typing import Any, Dict, Optional, Tuple
import streamlit as st
from src.db.client import get_supabase_client

ROOT_PROJECT_DIR = Path(__file__).resolve().parent.parent.parent

MAPA_MODULOS_SISTEMA = {
    "inicio": ("🎓 Inicio / Portal Principal", "🎓 Inicio", ROOT_PROJECT_DIR / "app.py"),
    "1_registro": ("👨‍🏫 1. Registro Docente", "👨‍🏫 1. Registro", ROOT_PROJECT_DIR / "pages" / "1_Registro_Docente.py"),
    "2_panel": ("📊 2. Panel y PDF", "📊 2. Panel y PDF", ROOT_PROJECT_DIR / "pages" / "2_Panel_y_PDF.py"),
    "3_matriculas": ("📈 3. Matrículas (2025-2026)", "📈 3. Matrículas", ROOT_PROJECT_DIR / "pages" / "3_Matriculas.py"),
    "4_homologaciones": ("🔄 4. Homologaciones", "🔄 4. Homolog.", ROOT_PROJECT_DIR / "pages" / "4_Homologaciones.py"),
    "5_creditos": ("🎯 5. Proyección Créditos ECISA", "🎯 5. Créditos", ROOT_PROJECT_DIR / "pages" / "5_Proyeccion_Creditos.py"),
}


def _detectar_modulo_script_principal() -> str:
    """Detecta cuál de los archivos del proyecto fue ejecutado como script raíz por Streamlit."""
    if "_modulo_script_raiz_fijo" in st.session_state and st.session_state["_modulo_script_raiz_fijo"]:
        return str(st.session_state["_modulo_script_raiz_fijo"])

    nombre_arch = ""
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        ctx = get_script_run_ctx()
        if ctx and getattr(ctx, "main_script_path", None):
            nombre_arch = Path(ctx.main_script_path).name.lower()
    except Exception:
        nombre_arch = ""

    if not nombre_arch:
        try:
            main_mod = sys.modules.get("__main__")
            main_file = getattr(main_mod, "__file__", "") or ""
            nombre_arch = Path(main_file).name.lower()
        except Exception:
            nombre_arch = ""

    mod_detectado = "inicio"
    if "1_registro" in nombre_arch:
        mod_detectado = "1_registro"
    elif "2_panel" in nombre_arch:
        mod_detectado = "2_panel"
    elif "3_matriculas" in nombre_arch:
        mod_detectado = "3_matriculas"
    elif "4_homologaciones" in nombre_arch:
        mod_detectado = "4_homologaciones"
    elif "5_proyeccion" in nombre_arch:
        mod_detectado = "5_creditos"

    st.session_state["_modulo_script_raiz_fijo"] = mod_detectado
    return mod_detectado


def navegar_a_modulo(clave_mod: str) -> None:
    """
    Navega de forma segura y consistente hacia cualquiera de los módulos del sistema.
    Funciona tanto si Streamlit se inició desde `app.py` como si en Streamlit Cloud
    el Main file path apunta a `pages/3_Matriculas.py` u otra página.
    """
    if clave_mod not in MAPA_MODULOS_SISTEMA:
        return

    st.session_state["modulo_navegacion_activo"] = clave_mod
    st.session_state["_navegacion_forzada"] = True
    st.rerun()

RUTA_CLAVES_LOCAL = Path(".streamlit/claves_docentes.json")


def hash_password(clave: str) -> str:
    """Genera el hash SHA-256 de una contraseña."""
    return hashlib.sha256(clave.strip().encode("utf-8")).hexdigest()


def _leer_claves_locales() -> Dict[str, str]:
    if not RUTA_CLAVES_LOCAL.exists():
        return {}
    try:
        return json.loads(RUTA_CLAVES_LOCAL.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _guardar_clave_local(documento: str, pwd_hash: str) -> None:
    try:
        mapa = _leer_claves_locales()
        mapa[str(documento).strip()] = pwd_hash
        RUTA_CLAVES_LOCAL.parent.mkdir(parents=True, exist_ok=True)
        RUTA_CLAVES_LOCAL.write_text(json.dumps(mapa, indent=2), encoding="utf-8")
    except Exception:
        pass


def obtener_credenciales_admin() -> Tuple[str, str]:
    """Obtiene el usuario y contraseña del Administrador desde `st.secrets` o valores por defecto."""
    try:
        user = str(st.secrets.get("ADMIN_USER", "admin")).strip()
        pwd = str(st.secrets.get("ADMIN_PASS", "admin2026")).strip()
        return (user or "admin"), (pwd or "admin2026")
    except Exception:
        return "admin", "admin2026"


def inicializar_estado_sesion() -> None:
    """Inicializa las variables de sesión de autenticación en `st.session_state`."""
    defaults = {
        "autenticado": False,
        "rol_usuario": None,          # 'admin' o 'docente'
        "usuario_login": None,        # 'admin' o número de documento del docente
        "nombre_usuario": None,
        "docente_autenticado": None,  # dict del docente cuando rol == 'docente'
        "docente_actual": None,       # dict del docente cargado en el formulario de edición
        "forzar_cambio_clave": False, # True cuando el docente sigue con su cédula como clave inicial
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def cerrar_sesion() -> None:
    """Limpia la sesión actual y redirige al login en el Portal Principal."""
    for k in [
        "autenticado",
        "rol_usuario",
        "usuario_login",
        "nombre_usuario",
        "docente_autenticado",
        "docente_actual",
        "forzar_cambio_clave",
        "criterio_ultimo",
    ]:
        if k in st.session_state:
            st.session_state[k] = False if k in {"autenticado", "forzar_cambio_clave"} else None
    st.session_state["modulo_navegacion_activo"] = "inicio"


def autenticar_usuario(usuario: str, clave: str) -> Tuple[bool, str]:
    """
    Valida las credenciales ingresadas:
    1. Primero verifica si corresponde al usuario Administrador (`admin`).
    2. Si no, busca el docente por su número de cédula (`documento`) en Supabase:
       - Si el docente aún no ha cambiado su clave (`password_hash` vacío), valida que `clave == documento`.
       - Si ya tiene `password_hash`, valida contra `hash_password(clave)`.
    """
    inicializar_estado_sesion()
    user_clean = usuario.strip()
    pass_clean = clave.strip()

    if not user_clean or not pass_clean:
        return False, "Por favor ingresa tanto el usuario como la contraseña."

    # 1. Verificar si es el Administrador
    admin_user, admin_pass = obtener_credenciales_admin()
    if user_clean.lower() == admin_user.lower():
        if pass_clean == admin_pass:
            st.session_state["autenticado"] = True
            st.session_state["rol_usuario"] = "admin"
            st.session_state["usuario_login"] = admin_user
            st.session_state["nombre_usuario"] = "Administrador del Sistema"
            st.session_state["docente_autenticado"] = None
            st.session_state["forzar_cambio_clave"] = False
            return True, "Bienvenido, Administrador."
        return False, "Contraseña de Administrador incorrecta."

    # 2. Verificar si es un Docente (por número de cédula)
    client = get_supabase_client()
    if client is None:
        return False, "No hay conexión con Supabase para validar el documento."

    try:
        res = (
            client.table("docentes")
            .select("*")
            .eq("documento", user_clean)
            .limit(1)
            .execute()
        )
        if not res.data:
            return (
                False,
                f"No se encontró ningún docente registrado con la cédula `{user_clean}`.",
            )

        docente = res.data[0]
        from src.services.docentes_svc import obtener_estado_efectivo_docente
        docente["estado"] = obtener_estado_efectivo_docente(docente)
        doc_cedula = str(docente.get("documento", "")).strip()

        # Consultar hash guardado en Supabase o en el archivo de respaldo local
        hash_db: Optional[str] = docente.get("password_hash")
        if not hash_db:
            claves_locales = _leer_claves_locales()
            hash_db = claves_locales.get(doc_cedula)

        hash_ingresado = hash_password(pass_clean)

        if hash_db:
            # El docente ya había personalizado su contraseña
            if hash_ingresado != hash_db:
                return False, "Contraseña incorrecta. Si ya cambiaste tu clave, usa tu nueva contraseña."
            es_clave_inicial = (pass_clean == doc_cedula)
        else:
            # Clave inicial: número de cédula en ambos campos
            if pass_clean != doc_cedula:
                return (
                    False,
                    "Contraseña incorrecta. Para tu primer ingreso, tu contraseña inicial es tu mismo número de cédula.",
                )
            es_clave_inicial = True

        st.session_state["autenticado"] = True
        st.session_state["rol_usuario"] = "docente"
        st.session_state["usuario_login"] = doc_cedula
        st.session_state["nombre_usuario"] = docente.get("nombre_completo", doc_cedula)
        st.session_state["docente_autenticado"] = docente
        st.session_state["docente_actual"] = docente
        st.session_state["forzar_cambio_clave"] = es_clave_inicial
        return True, f"Bienvenido(a), {docente.get('nombre_completo')}."

    except Exception as exc:
        return False, f"Error al validar credenciales en Supabase: {exc}"


def cambiar_clave_docente(
    documento: str,
    clave_actual: str,
    clave_nueva: str,
    clave_confirmacion: str,
) -> Tuple[bool, str]:
    """
    Permite al docente cambiar su contraseña (validando su clave actual y guardando el hash SHA-256).
    """
    doc_clean = documento.strip()
    act_clean = clave_actual.strip()
    nueva_clean = clave_nueva.strip()
    conf_clean = clave_confirmacion.strip()

    if not act_clean or not nueva_clean or not conf_clean:
        return False, "Todos los campos de contraseña son obligatorios."

    if len(nueva_clean) < 4:
        return False, "La nueva contraseña debe tener al menos 4 caracteres."

    if nueva_clean != conf_clean:
        return False, "La nueva contraseña y su confirmación no coinciden."

    if nueva_clean == doc_clean:
        return False, "La nueva contraseña debe ser diferente a tu número de cédula."

    client = get_supabase_client()
    if client is None:
        return False, "Cliente de Supabase no disponible."

    try:
        res = (
            client.table("docentes")
            .select("*")
            .eq("documento", doc_clean)
            .limit(1)
            .execute()
        )
        if not res.data:
            return False, "No se encontró el registro del docente."

        docente = res.data[0]
        hash_db: Optional[str] = docente.get("password_hash")
        if not hash_db:
            hash_db = _leer_claves_locales().get(doc_clean)

        if hash_db:
            if hash_password(act_clean) != hash_db:
                return False, "La contraseña actual ingresada no es correcta."
        else:
            if act_clean != doc_clean:
                return False, "Tu contraseña actual es tu número de cédula."

        nuevo_hash = hash_password(nueva_clean)

        # Intentar guardar en Supabase; si la columna password_hash aún no se ha corrido en SQL, guarda también en local
        try:
            client.table("docentes").update(
                {
                    "password_hash": nuevo_hash,
                    "debe_cambiar_clave": False,
                }
            ).eq("documento", doc_clean).execute()
        except Exception:
            pass

        _guardar_clave_local(doc_clean, nuevo_hash)
        st.session_state["forzar_cambio_clave"] = False
        return True, "✅ Tu contraseña ha sido actualizada exitosamente."
    except Exception as exc:
        return False, f"Error al actualizar la contraseña: {exc}"


def renderizar_barra_sesion_y_login(modulo_actual: Optional[str] = None) -> bool:
    """
    Renderiza tanto una barra superior horizontal fija como la barra lateral (`st.sidebar`) con las 6 pestañas
    del sistema, el estado de la sesión activa y el formulario de inicio de sesión.

    Siempre arranca por defecto en `"inicio"` (Portal Principal `app.py`), incluso si en Streamlit Cloud
    el Main file path fue configurado apuntando a `pages/3_Matriculas.py` o a cualquier otra página.

    Retorna `True` si el usuario está autenticado, `False` si aún debe iniciar sesión.
    """
    inicializar_estado_sesion()

    mod_raiz = _detectar_modulo_script_principal()
    mod_este_archivo = modulo_actual or mod_raiz
    # Si estamos en el script raíz real de Streamlit, nunca estamos en delegación
    # (previene que un st.stop() o st.rerun() dentro de una página delegada deje pegado el flag)
    if mod_este_archivo == mod_raiz:
        st.session_state["_ejecutando_delegacion_modulo"] = False
    en_delegacion = bool(st.session_state.get("_ejecutando_delegacion_modulo"))

    # Siempre iniciar por defecto en el Portal Principal ("inicio") al abrir la aplicación por primera vez
    if "modulo_navegacion_activo" not in st.session_state or not st.session_state["modulo_navegacion_activo"]:
        st.session_state["modulo_navegacion_activo"] = "inicio"

    if not en_delegacion:
        mod_activo = st.session_state.get("modulo_navegacion_activo", "inicio")

        # 1. BARRA DE NAVEGACIÓN SUPERIOR VISIBLE SIEMPRE (INCLUSO SI EL SIDEBAR ESTÁ COLAPSADO)
        with st.container(border=True):
            cols_top = st.columns([1, 1.1, 1.1, 1.1, 1.1, 1.1, 0.9])
            for idx_m, (clave_mod, (_, etiqueta_corta, _)) in enumerate(MAPA_MODULOS_SISTEMA.items()):
                es_activo = (clave_mod == mod_activo)
                with cols_top[idx_m]:
                    if st.button(
                        etiqueta_corta,
                        key=f"btn_nav_top_{clave_mod}",
                        type="primary" if es_activo else "secondary",
                        use_container_width=True,
                    ):
                        navegar_a_modulo(clave_mod)

            with cols_top[6]:
                if st.session_state.get("autenticado"):
                    if st.button("🚪 Salir", key="btn_logout_top", use_container_width=True, help="Cerrar sesión"):
                        cerrar_sesion()
                        st.rerun()
                else:
                    st.caption("🔒 Sin sesión")

        # 2. MENÚ LATERAL (SIDEBAR)
        with st.sidebar:
            st.markdown("### 🧭 Módulos del Sistema")
            for clave_mod, (etiqueta_mod, _, _) in MAPA_MODULOS_SISTEMA.items():
                es_activo = (clave_mod == mod_activo)
                if st.button(
                    etiqueta_mod,
                    key=f"btn_nav_global_{clave_mod}",
                    type="primary" if es_activo else "secondary",
                    use_container_width=True,
                ):
                    navegar_a_modulo(clave_mod)

            st.divider()
            st.markdown("### 🔐 Sesión de Usuario")
            if st.session_state.get("autenticado"):
                rol = st.session_state.get("rol_usuario")
                nombre = st.session_state.get("nombre_usuario")
                login_id = st.session_state.get("usuario_login")

                if rol == "admin":
                    st.success(f"👑 **Rol:** Administrador\n\n**Usuario:** `{login_id}`")
                else:
                    st.info(f"👨‍🏫 **Docente:** {nombre}\n\n**Cédula:** `{login_id}`")

                    with st.expander("🔑 Cambiar mi Contraseña", expanded=bool(st.session_state.get("forzar_cambio_clave"))):
                        if st.session_state.get("forzar_cambio_clave"):
                            st.caption("⚠️ Estás usando tu cédula como clave inicial. Te recomendamos personalizarla:")
                        with st.form("form_cambio_clave_sidebar", clear_on_submit=True):
                            c_act = st.text_input("Contraseña Actual", type="password")
                            c_nue = st.text_input("Nueva Contraseña", type="password")
                            c_con = st.text_input("Confirmar Nueva Contraseña", type="password")
                            btn_cambiar = st.form_submit_button("Actualizar Clave", type="primary", use_container_width=True)

                        if btn_cambiar:
                            ok_c, msg_c = cambiar_clave_docente(
                                documento=str(login_id),
                                clave_actual=c_act,
                                clave_nueva=c_nue,
                                clave_confirmacion=c_con,
                            )
                            if ok_c:
                                st.success(msg_c)
                            else:
                                st.error(msg_c)

                if st.button("🚪 Cerrar Sesión", key="btn_logout_sidebar", use_container_width=True):
                    cerrar_sesion()
                    st.rerun()
            else:
                st.warning("No has iniciado sesión.")

        # 3. SI AÚN NO HA INICIADO SESIÓN, MOSTRAR EL LOGIN ANTES DE CUALQUIER MÓDULO
        if not st.session_state.get("autenticado"):
            st.markdown("### 🎓 Sistema de Gestión Académica — UNAD (ECISA)")
            col_l, col_center, col_r = st.columns([1, 1.6, 1])
            with col_center:
                with st.container(border=True):
                    st.markdown("### 🔐 Ingreso al Sistema de Gestión Académica")
                    st.markdown(
                        """
                        * **👨‍🏫 Docentes:** Ingresa tu **número de cédula** como Usuario y como Contraseña inicial (luego podrás cambiar tu clave en el menú lateral).
                        * **👑 Administrador:** Ingresa con tus credenciales de administración (`admin`).
                        """
                    )
                    with st.form("form_login_principal", clear_on_submit=False):
                        inp_user = st.text_input(
                            "Usuario (Cédula del Docente o `admin`)",
                            placeholder="Ej: 72142321 o admin",
                        )
                        inp_pass = st.text_input(
                            "Contraseña (Inicialmente tu número de cédula)",
                            type="password",
                            placeholder="Ingresa tu contraseña",
                        )
                        btn_login = st.form_submit_button(
                            "🔓 Iniciar Sesión",
                            type="primary",
                            use_container_width=True,
                        )

                    if btn_login:
                        ok_auth, msg_auth = autenticar_usuario(inp_user, inp_pass)
                        if ok_auth:
                            st.session_state["modulo_navegacion_activo"] = "inicio"
                            st.success(msg_auth)
                            st.rerun()
                        else:
                            st.error(f"❌ {msg_auth}")
            return False

        # 4. SI EL MÓDULO SELECCIONADO ES DISTINTO AL ARCHIVO RAÍZ ACTUAL, EJECUTARLO CON RUNPY
        mod_destino = st.session_state.get("modulo_navegacion_activo", "inicio")
        if mod_destino != mod_este_archivo and mod_destino in MAPA_MODULOS_SISTEMA:
            _, _, ruta_destino = MAPA_MODULOS_SISTEMA[mod_destino]
            if ruta_destino.exists():
                orig_set_page_config = st.set_page_config
                st.set_page_config = lambda *a, **kw: None
                st.session_state["_ejecutando_delegacion_modulo"] = True
                try:
                    runpy.run_path(str(ruta_destino), run_name="__main__")
                finally:
                    # SIEMPRE limpiar la bandera incluso si la página delegada llamó a st.stop() o st.rerun()
                    st.set_page_config = orig_set_page_config
                    st.session_state["_ejecutando_delegacion_modulo"] = False
                st.stop()

    return bool(st.session_state.get("autenticado"))


def es_admin() -> bool:
    """Retorna True únicamente si el usuario autenticado tiene rol `admin`."""
    return bool(st.session_state.get("autenticado") and st.session_state.get("rol_usuario") == "admin")

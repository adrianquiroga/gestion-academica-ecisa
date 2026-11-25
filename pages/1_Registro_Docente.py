"""
Módulo 1 (Fase 1 - MVP): Registro, Actualización e Importación de Resoluciones Docentes.
Adaptado para manejar todos los campos de `Resoluciones Docentes.xlsx` (UNAD):
1. Importación con 1 clic desde `Resoluciones Docentes.xlsx` (sincroniza los 101+ docentes y desglosa sus títulos).
2. Búsqueda de docente por documento, correo institucional o nombre.
3. Edición de datos personales, contacto, nivel de inglés (con certificado en Google Drive) y datos de Resolución/Vinculación.
4. Gestión dinámica de títulos académicos (`docente_estudios`) y aportes (`docente_aportes`) con soportes en Drive.
"""

from datetime import datetime
import importlib
import streamlit as st
from src.db.client import verificar_configuracion_supabase
import src.services.docentes_svc as _doc_svc
import src.services.drive_svc as _drive_svc
from src.services.pdf_svc import generar_pdf_ficha_docente
import src.utils.auth as _auth_mod

importlib.reload(_doc_svc)
importlib.reload(_drive_svc)
importlib.reload(_auth_mod)

FOLDER_ID_SOPORTES_APORTES = _drive_svc.FOLDER_ID_SOPORTES_APORTES
FOLDER_ID_SOPORTES_INGLES = _drive_svc.FOLDER_ID_SOPORTES_INGLES
URL_CARPETA_DRIVE_APORTES = _drive_svc.URL_CARPETA_DRIVE_APORTES
subir_soporte_drive = _drive_svc.subir_soporte_drive
validar_archivo_soporte = _drive_svc.validar_archivo_soporte

CARPETA_DRIVE_SOPORTES_INGLES_URL = _doc_svc.CARPETA_DRIVE_SOPORTES_INGLES_URL
CATEGORIAS_APORTE = _doc_svc.CATEGORIAS_APORTE
MODALIDADES_VINCULACION = _doc_svc.MODALIDADES_VINCULACION
NIVELES_FORMACION = _doc_svc.NIVELES_FORMACION
NIVELES_FORMACION_MAXIMO = _doc_svc.NIVELES_FORMACION_MAXIMO
NIVELES_INGLES = _doc_svc.NIVELES_INGLES
RUTA_EXCEL_INGLES_RESPUESTAS_DEFECTO = _doc_svc.RUTA_EXCEL_INGLES_RESPUESTAS_DEFECTO
RUTA_EXCEL_RESOLUCIONES_DEFECTO = _doc_svc.RUTA_EXCEL_RESOLUCIONES_DEFECTO
agregar_aporte_docente = _doc_svc.agregar_aporte_docente
agregar_estudio_docente = _doc_svc.agregar_estudio_docente
buscar_docente = _doc_svc.buscar_docente
eliminar_aporte_docente = _doc_svc.eliminar_aporte_docente
eliminar_estudio_docente = _doc_svc.eliminar_estudio_docente
guardar_o_actualizar_docente = _doc_svc.guardar_o_actualizar_docente
importar_resoluciones_excel_a_supabase = _doc_svc.importar_resoluciones_excel_a_supabase
importar_respuestas_ingles_a_supabase = _doc_svc.importar_respuestas_ingles_a_supabase
obtener_aportes_docente = _doc_svc.obtener_aportes_docente
obtener_estudios_docente = _doc_svc.obtener_estudios_docente
previsualizar_excel_local = _doc_svc.previsualizar_excel_local

cambiar_clave_docente = _auth_mod.cambiar_clave_docente
es_admin = _auth_mod.es_admin
renderizar_barra_sesion_y_login = _auth_mod.renderizar_barra_sesion_y_login

st.set_page_config(
    page_title="Registro y Resoluciones Docentes",
    page_icon="👨‍🏫",
    layout="wide",
)

st.title("👨‍🏫 Gestión de Perfil y Resoluciones Docentes")

if not renderizar_barra_sesion_y_login():
    st.stop()

usuario_es_admin = es_admin()

if "busqueda_realizada" not in st.session_state:
    st.session_state["busqueda_realizada"] = False
if "criterio_ultimo" not in st.session_state:
    st.session_state["criterio_ultimo"] = ""

if usuario_es_admin:
    st.markdown(
        "**Modo Administrador (`admin`):** Puedes buscar y editar a cualquier docente por **documento, correo institucional o nombre**, "
        "o sincronizar la base de datos desde **`Resoluciones Docentes.xlsx`**."
    )
else:
    st.markdown(
        "Bienvenido(a) a tu portal docente. Aquí puedes actualizar tu **nivel de inglés, adjuntar tus certificados en Google Drive**, "
        "registrar tus **títulos académicos y aportes**, y **cambiar tu contraseña de acceso**."
    )

ok_db, msg_db = verificar_configuracion_supabase()
if not ok_db:
    st.error(f"⚠️ Base de datos no configurada: {msg_db}")
    st.stop()

# Si el usuario autenticado es un docente, cargamos su propio perfil automáticamente
if not usuario_es_admin:
    doc_sesion = st.session_state.get("docente_autenticado")
    if doc_sesion:
        st.session_state["docente_actual"] = doc_sesion

    # Panel directo para que el docente cambie su contraseña en su propia pantalla
    with st.expander(
        "🔑 Cambiar mi Contraseña de Acceso (Inicialmente es tu número de cédula)",
        expanded=bool(st.session_state.get("forzar_cambio_clave")),
    ):
        c_pw1, c_pw2, c_pw3, c_pw_btn = st.columns([1.2, 1.2, 1.2, 1])
        with c_pw1:
            inp_pw_act = st.text_input("Contraseña Actual", type="password", key="pw_act_main")
        with c_pw2:
            inp_pw_nue = st.text_input("Nueva Contraseña", type="password", key="pw_nue_main")
        with c_pw3:
            inp_pw_con = st.text_input("Confirmar Nueva Contraseña", type="password", key="pw_con_main")
        with c_pw_btn:
            st.write("")
            if st.button("💾 Guardar Nueva Clave", type="primary", use_container_width=True):
                ok_pw, msg_pw = cambiar_clave_docente(
                    documento=str(st.session_state.get("usuario_login")),
                    clave_actual=inp_pw_act,
                    clave_nueva=inp_pw_nue,
                    clave_confirmacion=inp_pw_con,
                )
                if ok_pw:
                    st.success(msg_pw)
                else:
                    st.error(msg_pw)

# =============================================================================
# 0. PANEL DESPLEGABLE: IMPORTACIÓN / SINCRONIZACIÓN DE `Resoluciones Docentes.xlsx` (SOLO ADMIN)
# =============================================================================
if usuario_es_admin:
    with st.expander("📥 Importar / Sincronizar Base de Datos desde `Resoluciones Docentes.xlsx` (Solo Admin)", expanded=False):
        st.markdown(
            """
            Esta herramienta procesa las **32 columnas** de `Resoluciones Docentes.xlsx`:
            * Normaliza el **Nivel de Inglés** al estándar MCER (`A1`..`C2`, `En caracterización / Prueba`, `Sin certificar`) conservando el texto original.
            * Carga los datos de **Resolución, Periodo, Modalidad, Horas, Salario, Total Vinculación, Escuela, Centro (CEAD/CCAV/UDR) y Programa**.
            * Desglosa automáticamente los títulos de las columnas **`PROFESIONAL`**, **`ESPECIALIZACIÓN`**, **`MAESTRÍA`** y **`DOCTORADO`** en la tabla `docente_estudios` sin duplicar registros.
            """
        )

        col_imp_local, col_imp_upload = st.columns(2)

        with col_imp_local:
            df_local = previsualizar_excel_local()
            if df_local is not None:
                st.info(
                    f"📄 Archivo detectado en el proyecto: **`{RUTA_EXCEL_RESOLUCIONES_DEFECTO}`** "
                    f"({len(df_local)} docentes válidos encontrados)."
                )
                if st.button(
                    f"🚀 Sincronizar `{RUTA_EXCEL_RESOLUCIONES_DEFECTO}` con Supabase ({len(df_local)} docentes)",
                    type="primary",
                    use_container_width=True,
                ):
                    with st.spinner("Sincronizando docentes y desglosando títulos académicos en Supabase..."):
                        ok_imp, n_doc, n_est, msg_imp = importar_resoluciones_excel_a_supabase(
                            fuente_excel=None,
                            sincronizar_estudios=True,
                        )
                    if ok_imp:
                        st.success(f"✅ {msg_imp}")
                    else:
                        st.error(f"❌ {msg_imp}")
            else:
                st.warning("No se encontró `Resoluciones Docentes.xlsx` en la carpeta local.")

        with col_imp_upload:
            excel_subido = st.file_uploader(
                "O sube una versión actualizada de `Resoluciones Docentes.xlsx`:",
                type=["xlsx"],
                key="uploader_excel_resoluciones",
            )
            if excel_subido is not None:
                if st.button("📥 Importar Archivo Excel Subido", use_container_width=True):
                    with st.spinner("Procesando archivo Excel subido y sincronizando con Supabase..."):
                        ok_imp, n_doc, n_est, msg_imp = importar_resoluciones_excel_a_supabase(
                            fuente_excel=excel_subido,
                            sincronizar_estudios=True,
                        )
                    if ok_imp:
                        st.success(f"✅ {msg_imp}")
                    else:
                        st.error(f"❌ {msg_imp}")

        st.divider()
        st.markdown("##### 🌐 Sincronizar Encuesta `Evaluacion de Nivel de Ingles para Docentes (respuestas).xlsx`")
        col_ing_loc, col_ing_up = st.columns(2)
        with col_ing_loc:
            if RUTA_EXCEL_INGLES_RESPUESTAS_DEFECTO.exists():
                st.info(f"📄 Archivo detectado: **`{RUTA_EXCEL_INGLES_RESPUESTAS_DEFECTO}`**")
                if st.button(
                    "🚀 Cruzar y Actualizar Niveles y Soportes de Inglés en Drive",
                    use_container_width=True,
                ):
                    with st.spinner("Actualizando niveles de inglés y enlaces a certificados en Drive..."):
                        ok_i, n_act, n_no, msg_i = importar_respuestas_ingles_a_supabase(None)
                    if ok_i:
                        st.success(f"✅ {msg_i}")
                    else:
                        st.error(f"❌ {msg_i}")
            st.link_button(
                "📂 Abrir Carpeta de Certificados de Inglés en Google Drive",
                url=CARPETA_DRIVE_SOPORTES_INGLES_URL,
                use_container_width=True,
            )
        with col_ing_up:
            excel_ing_subido = st.file_uploader(
                "O sube una versión actualizada de `Evaluacion de Nivel de Ingles (respuestas).xlsx`:",
                type=["xlsx"],
                key="uploader_excel_ingles_respuestas",
            )
            if excel_ing_subido is not None:
                if st.button("📥 Importar Respuestas de Inglés Subidas", use_container_width=True):
                    with st.spinner("Cruzando respuestas de inglés con Supabase..."):
                        ok_i, n_act, n_no, msg_i = importar_respuestas_ingles_a_supabase(excel_ing_subido)
                    if ok_i:
                        st.success(f"✅ {msg_i}")
                    else:
                        st.error(f"❌ {msg_i}")

    # =============================================================================
    # 1. BARRA DE BÚSQUEDA DE DOCENTE (SOLO ADMIN)
    # =============================================================================
    with st.container(border=True):
        st.markdown("#### 🔍 1. Buscar o Inicializar Perfil Docente (Modo Administrador)")
        col_input, col_btn_buscar, col_btn_limpiar = st.columns([3, 1, 1])

        with col_input:
            criterio = st.text_input(
                "Documento, correo institucional o nombre del docente",
                value=st.session_state.get("criterio_ultimo") or "",
                placeholder="Ej: 72142321, adrian.quiroga@unad.edu.co o ADRIAN QUIROGA",
                label_visibility="collapsed",
            )

        with col_btn_buscar:
            btn_buscar = st.button("🔍 Buscar / Cargar", type="primary", use_container_width=True)

        with col_btn_limpiar:
            btn_nuevo = st.button("➕ Nuevo / Limpiar", use_container_width=True)

        if btn_nuevo:
            st.session_state["docente_actual"] = None
            st.session_state["busqueda_realizada"] = True
            st.session_state["criterio_ultimo"] = ""
            st.rerun()

        if btn_buscar:
            st.session_state["criterio_ultimo"] = criterio.strip()
            st.session_state["busqueda_realizada"] = True
            docente_encontrado, mensaje_busqueda = buscar_docente(criterio)
            st.session_state["docente_actual"] = docente_encontrado
            if docente_encontrado:
                st.success(f"✅ {mensaje_busqueda}")
            else:
                st.info(f"ℹ️ {mensaje_busqueda}")

docente_actual = st.session_state.get("docente_actual")

# Inicializar valores del formulario con los datos actuales o por defecto
docente_id_actual = docente_actual.get("id") if docente_actual else None
val_doc = docente_actual.get("documento", "") if docente_actual else ""
val_nombre = docente_actual.get("nombre_completo", "") if docente_actual else ""
val_correo = docente_actual.get("correo_institucional", "") if docente_actual else ""
val_email_personal = (docente_actual.get("email_personal") or "") if docente_actual else ""
val_telefono = (docente_actual.get("telefono") or "") if docente_actual else ""
val_cumple = (docente_actual.get("cumpleanos_mes_dia") or "") if docente_actual else ""
val_ingles = docente_actual.get("nivel_ingles", "Sin certificar") if docente_actual else "Sin certificar"
val_detalle_ingles = (docente_actual.get("detalle_ingles_original") or "") if docente_actual else ""
val_soporte_ingles = docente_actual.get("soporte_ingles_url") if docente_actual else None

# Valores de Resolución / Vinculación
val_nivel_max = (docente_actual.get("nivel_formacion_maximo") or "MAGISTER") if docente_actual else "MAGISTER"
val_modalidad = (docente_actual.get("modalidad") or "TIEMPO COMPLETO") if docente_actual else "TIEMPO COMPLETO"
val_periodo = (docente_actual.get("periodo_academico") or "2201") if docente_actual else "2201"
val_estado_acep = (docente_actual.get("estado_aceptacion") or "ACEPTO") if docente_actual else "ACEPTO"
val_horas = int(docente_actual.get("horas_vinculacion") or 0) if docente_actual else 0
val_res_num = (docente_actual.get("resolucion_numero") or "") if docente_actual else ""
val_res_fecha = (docente_actual.get("resolucion_fecha") or "") if docente_actual else ""
val_f_inicio = (docente_actual.get("fecha_inicio") or "") if docente_actual else ""
val_f_termina = (docente_actual.get("fecha_termina") or "") if docente_actual else ""
val_escuela = (docente_actual.get("escuela_descripcion") or "ESCUELA DE CIENCIAS DE LA SALUD") if docente_actual else "ESCUELA DE CIENCIAS DE LA SALUD"
val_centro = (docente_actual.get("centro_descripcion") or "") if docente_actual else ""
val_zona = (docente_actual.get("zona_descripcion") or "") if docente_actual else ""
val_programa = (docente_actual.get("programa_descripcion") or "") if docente_actual else ""
val_salario = float(docente_actual.get("salario_mensual") or 0.0) if docente_actual else 0.0
val_total_vinc = float(docente_actual.get("total_vinculacion") or 0.0) if docente_actual else 0.0

if not docente_actual and st.session_state.get("criterio_ultimo"):
    crit = st.session_state["criterio_ultimo"]
    if "@" in crit:
        val_correo = crit
    elif crit.isdigit():
        val_doc = crit
    else:
        val_nombre = crit.upper()

# =============================================================================
# 2. FORMULARIO DE DATOS PERSONALES, INGLÉS Y RESOLUCIÓN DE VINCULACIÓN
# =============================================================================
with st.container(border=True):
    encabezado_perfil = (
        f"📝 2. Perfil Docente y Resolución — Editando a: {val_nombre} (Doc: {val_doc})"
        if docente_id_actual
        else "📝 2. Perfil Docente y Resolución (Nuevo Registro)"
    )
    st.markdown(f"#### {encabezado_perfil}")

    with st.form("form_datos_basicos_docente", clear_on_submit=False):
        st.markdown("##### 👤 Identificación y Contacto")
        c_doc, c_nom, c_cor = st.columns([1, 1.6, 1.4])
        with c_doc:
            input_documento = st.text_input("Número de Documento *", value=val_doc)
        with c_nom:
            input_nombre = st.text_input("Nombre Completo *", value=val_nombre)
        with c_cor:
            input_correo = st.text_input("Correo Institucional *", value=val_correo)

        c_tel, c_em_per, c_cum, c_niv_max = st.columns([1, 1.4, 0.8, 1.2])
        with c_tel:
            input_telefono = st.text_input("Teléfono Celular", value=val_telefono)
        with c_em_per:
            input_email_personal = st.text_input("Email Personal", value=val_email_personal)
        with c_cum:
            input_cumple = st.text_input("Cumpleaños (MM-DD)", value=val_cumple, placeholder="Ej: 07-17")
        with c_niv_max:
            opciones_niv_max = NIVELES_FORMACION_MAXIMO
            if val_nivel_max and val_nivel_max not in opciones_niv_max:
                opciones_niv_max = [val_nivel_max] + opciones_niv_max
            input_nivel_max = st.selectbox(
                "Nivel Formación en Resolución",
                options=opciones_niv_max,
                index=opciones_niv_max.index(val_nivel_max) if val_nivel_max in opciones_niv_max else 0,
            )

        st.markdown("##### 🌐 Competencia en Inglés y Soporte en Google Drive")
        c_ing, c_det_ing, c_file = st.columns([1, 1.1, 1.9])
        with c_ing:
            idx_ingles = NIVELES_INGLES.index(val_ingles) if val_ingles in NIVELES_INGLES else 0
            input_nivel_ingles = st.selectbox(
                "Nivel de Inglés (MCER) *",
                options=NIVELES_INGLES,
                index=idx_ingles,
            )
            if val_soporte_ingles:
                st.markdown(f"📎 **Certificado en Drive:** [Abrir soporte]({val_soporte_ingles})")
        with c_det_ing:
            input_detalle_ingles = st.text_input(
                "Detalle / Examen de Inglés",
                value=val_detalle_ingles,
                placeholder="Ej: B2 - ADVANCED o STEP UNAD",
            )
        with c_file:
            archivo_ingles = st.file_uploader(
                "Certificado de Inglés (Opcional - PDF, PNG o JPG máx. 5 MB)",
                type=["pdf", "png", "jpg", "jpeg"],
                key="uploader_ingles",
            )

        with st.expander("📋 Datos de Resolución y Vinculación Académica (UNAD)", expanded=bool(val_res_num)):
            r1, r2, r3, r4, r5 = st.columns(5)
            with r1:
                input_periodo = st.text_input("Periodo Académico", value=val_periodo)
            with r2:
                input_res_num = st.text_input("N° Resolución", value=val_res_num)
            with r3:
                input_res_fecha = st.text_input("Fecha Resolución (YYYY-MM-DD)", value=str(val_res_fecha)[:10])
            with r4:
                opciones_mod = MODALIDADES_VINCULACION
                if val_modalidad and val_modalidad not in opciones_mod:
                    opciones_mod = [val_modalidad] + opciones_mod
                input_modalidad = st.selectbox(
                    "Modalidad",
                    options=opciones_mod,
                    index=opciones_mod.index(val_modalidad) if val_modalidad in opciones_mod else 0,
                )
            with r5:
                input_horas = st.number_input("Horas Vinculación", min_value=0, max_value=3000, value=val_horas, step=1)

            r6, r7, r8, r9 = st.columns([1.3, 1.3, 1.4, 1])
            with r6:
                input_escuela = st.text_input("Escuela", value=val_escuela)
            with r7:
                input_centro = st.text_input("Centro (CEAD/CCAV/UDR)", value=val_centro)
            with r8:
                input_programa = st.text_input("Programa Académico", value=val_programa)
            with r9:
                input_zona = st.text_input("Zona", value=val_zona)

            r10, r11, r12, r13, r14 = st.columns(5)
            with r10:
                input_f_inicio = st.text_input("Fecha Inicio (YYYY-MM-DD)", value=str(val_f_inicio)[:10])
            with r11:
                input_f_termina = st.text_input("Fecha Termina (YYYY-MM-DD)", value=str(val_f_termina)[:10])
            with r12:
                input_estado_acep = st.text_input("Estado Aceptación", value=val_estado_acep)
            with r13:
                input_salario = st.number_input("Salario Mensual ($)", min_value=0.0, value=val_salario, step=100000.0)
            with r14:
                input_total_vinc = st.number_input("Total Vinculación ($)", min_value=0.0, value=val_total_vinc, step=500000.0)

        btn_guardar_perfil = st.form_submit_button(
            "💾 Guardar Perfil Docente, Nivel de Inglés y Datos de Resolución",
            type="primary",
            use_container_width=True,
        )

    if btn_guardar_perfil:
        if not input_documento.strip() or not input_nombre.strip() or not input_correo.strip():
            st.error("Por favor completa los campos obligatorios: Documento, Nombre Completo y Correo Institucional.")
        else:
            nueva_url_ingles = val_soporte_ingles
            error_subida = False

            if archivo_ingles is not None:
                es_val, msg_val = validar_archivo_soporte(archivo_ingles)
                if not es_val:
                    st.error(f"❌ Archivo de inglés inválido: {msg_val}")
                    error_subida = True
                else:
                    with st.spinner("Subiendo certificado de inglés a Google Drive Institucional..."):
                        ok_up, url_drive, msg_up = subir_soporte_drive(
                            uploaded_file=archivo_ingles,
                            documento=input_documento.strip(),
                            carpeta_categoria="Docentes/Ingles",
                            etiqueta_archivo=f"Ingles_{input_nivel_ingles}",
                        )
                    if ok_up and url_drive:
                        nueva_url_ingles = url_drive
                        st.toast(msg_up, icon="✅")
                    else:
                        st.error(f"❌ {msg_up}")
                        error_subida = True

            if not error_subida:
                datos_extra = {
                    "telefono": input_telefono.strip() or None,
                    "email_personal": input_email_personal.strip() or None,
                    "cumpleanos_mes_dia": input_cumple.strip() or None,
                    "nivel_formacion_maximo": input_nivel_max,
                    "detalle_ingles_original": input_detalle_ingles.strip() or input_nivel_ingles,
                    "periodo_academico": input_periodo.strip() or None,
                    "resolucion_numero": input_res_num.strip() or None,
                    "resolucion_fecha": input_res_fecha.strip() or None,
                    "modalidad": input_modalidad,
                    "horas_vinculacion": int(input_horas),
                    "escuela_descripcion": input_escuela.strip() or None,
                    "centro_descripcion": input_centro.strip() or None,
                    "programa_descripcion": input_programa.strip() or None,
                    "zona_descripcion": input_zona.strip() or None,
                    "fecha_inicio": input_f_inicio.strip() or None,
                    "fecha_termina": input_f_termina.strip() or None,
                    "estado_aceptacion": input_estado_acep.strip() or None,
                    "salario_mensual": float(input_salario) if input_salario else None,
                    "total_vinculacion": float(input_total_vinc) if input_total_vinc else None,
                }

                ok_save, doc_guardado, msg_save = guardar_o_actualizar_docente(
                    documento=input_documento,
                    nombre_completo=input_nombre,
                    correo_institucional=input_correo,
                    nivel_ingles=input_nivel_ingles,
                    soporte_ingles_url=nueva_url_ingles,
                    docente_id=docente_id_actual,
                    datos_adicionales=datos_extra,
                )
                if ok_save and doc_guardado:
                    st.session_state["docente_actual"] = doc_guardado
                    if not usuario_es_admin:
                        st.session_state["docente_autenticado"] = doc_guardado
                    st.session_state["criterio_ultimo"] = doc_guardado.get("documento", "")
                    st.success(f"✅ {msg_save}")
                    st.rerun()
                else:
                    st.error(f"❌ {msg_save}")

# =============================================================================
# 3. SECCIONES DINÁMICAS: ESTUDIOS Y APORTES
# =============================================================================
docente_actual = st.session_state.get("docente_actual")

if not docente_actual or not docente_actual.get("id"):
    st.info(
        "💡 **Paso siguiente:** Busca un docente arriba (o importa `Resoluciones Docentes.xlsx`) "
        "para gestionar sus **Títulos Académicos** y **Aportes Institucionales**."
    )
    st.stop()

docente_id = docente_actual["id"]
documento_docente = docente_actual["documento"]
estudios_actuales = obtener_estudios_docente(docente_id)
aportes_actuales = obtener_aportes_docente(docente_id)

col_resumen, col_pdf_rapido = st.columns([3.2, 1])
with col_resumen:
    st.success(
        f"**Docente:** {docente_actual['nombre_completo']} (`{documento_docente}`) | "
        f"**Centro:** `{docente_actual.get('centro_descripcion') or 'N/A'}` | "
        f"**Resolución:** `{docente_actual.get('resolucion_numero') or 'S/N'}` | "
        f"**Inglés:** `{docente_actual.get('nivel_ingles')}` | "
        f"**Títulos:** `{len(estudios_actuales)}` | **Aportes:** `{len(aportes_actuales)}`"
    )
with col_pdf_rapido:
    pdf_bytes_individual = generar_pdf_ficha_docente(
        docente=docente_actual,
        estudios=estudios_actuales,
        aportes=aportes_actuales,
    )
    st.download_button(
        label="📄 Descargar Ficha en PDF",
        data=pdf_bytes_individual,
        file_name=f"Ficha_Docente_{documento_docente}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

tab_estudios, tab_aportes = st.tabs(
    [
        f"🎓 3. Formación Académica / Títulos ({len(estudios_actuales)})",
        f"🚀 4. Aportes Académicos e Institucionales ({len(aportes_actuales)})",
    ]
)

# -----------------------------------------------------------------------------
# PESTAÑA 3: FORMACIÓN ACADÉMICA (SIN CARGA DE SOPORTE DIPLOMA/ACTA)
# -----------------------------------------------------------------------------
with tab_estudios:
    col_form_est, col_list_est = st.columns([1.15, 1.85])

    with col_form_est:
        with st.container(border=True):
            st.markdown("##### ➕ Agregar Nuevo Título Académico")
            with st.form("form_nuevo_estudio", clear_on_submit=True):
                est_nivel = st.selectbox("Nivel de Formación *", options=NIVELES_FORMACION, index=1)
                est_titulo = st.text_input("Título Obtenido *", placeholder="Ej: Magíster en Administración de Salud")
                est_universidad = st.text_input("Universidad / Institución *", placeholder="Ej: Universidad Nacional de Colombia")
                anio_actual = datetime.now().year
                est_anio = st.number_input(
                    "Año de Graduación",
                    min_value=1960,
                    max_value=anio_actual + 2,
                    value=min(2022, anio_actual),
                    step=1,
                )
                btn_add_estudio = st.form_submit_button("➕ Registrar Título", type="primary", use_container_width=True)

            if btn_add_estudio:
                if not est_titulo.strip() or not est_universidad.strip():
                    st.error("El título obtenido y la universidad son obligatorios.")
                else:
                    ok_ins, msg_ins = agregar_estudio_docente(
                        docente_id=docente_id,
                        nivel_formacion=est_nivel,
                        titulo_obtenido=est_titulo,
                        universidad=est_universidad,
                        anio_graduacion=int(est_anio),
                        soporte_url=None,
                    )
                    if ok_ins:
                        st.toast(msg_ins, icon="🎓")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg_ins}")

    with col_list_est:
        st.markdown("##### 📚 Títulos Registrados del Docente")
        if not estudios_actuales:
            st.info("Aún no hay títulos registrados para este docente.")
        else:
            for est in estudios_actuales:
                with st.container(border=True):
                    c_info, c_del = st.columns([4.5, 0.7])
                    with c_info:
                        st.markdown(f"**{est.get('nivel_formacion')}** — {est.get('titulo_obtenido')}")
                        st.caption(f"🏛️ {est.get('universidad')} | 📅 Año: {est.get('anio_graduacion') or 'N/A'}")
                    with c_del:
                        if st.button("🗑️", key=f"del_est_{est['id']}", help="Eliminar este título", use_container_width=True):
                            ok_del, msg_del = eliminar_estudio_docente(est["id"])
                            if ok_del:
                                st.toast(msg_del, icon="🗑️")
                                st.rerun()
                            else:
                                st.error(msg_del)

# -----------------------------------------------------------------------------
# PESTAÑA 4: APORTES ACADÉMICOS E INSTITUCIONALES (SUBCARPETA POR DOCENTE EN DRIVE)
# -----------------------------------------------------------------------------
with tab_aportes:
    nombre_subcarpeta_docente = f"{documento_docente} - {docente_actual.get('nombre_completo', 'Docente').strip()}"
    col_form_apo, col_list_apo = st.columns([1.15, 1.85])

    with col_form_apo:
        with st.container(border=True):
            st.markdown("##### ➕ Registrar Nuevo Aporte")
            st.caption(
                f"📁 Los anexos se guardarán automáticamente en Google Drive dentro de la subcarpeta del docente: "
                f"**`{nombre_subcarpeta_docente}`**"
            )
            with st.form("form_nuevo_aporte", clear_on_submit=True):
                apo_categoria = st.selectbox("Categoría del Aporte *", options=CATEGORIAS_APORTE)
                apo_periodo = st.text_input(
                    "Periodo Académico *",
                    value=str(docente_actual.get("periodo_academico") or "2026-2201"),
                )
                apo_titulo = st.text_input("Título del Aporte *", placeholder="Ej: Publicación artículo / Semillero / Ponencia")
                apo_descripcion = st.text_area("Descripción y Resultados", height=90)
                apo_archivo = st.file_uploader(
                    "Evidencia o Certificado (Opcional - PDF/PNG/JPG máx. 5 MB)",
                    type=["pdf", "png", "jpg", "jpeg"],
                    key="uploader_aporte",
                )
                btn_add_aporte = st.form_submit_button("➕ Registrar Aporte", type="primary", use_container_width=True)

            if btn_add_aporte:
                if not apo_titulo.strip() or not apo_periodo.strip():
                    st.error("El título del aporte y el periodo académico son obligatorios.")
                else:
                    url_soporte_apo = None
                    fallo_drive_apo = False
                    if apo_archivo is not None:
                        es_val, msg_val = validar_archivo_soporte(apo_archivo)
                        if not es_val:
                            st.error(f"❌ {msg_val}")
                            fallo_drive_apo = True
                        else:
                            with st.spinner(
                                f"Subiendo evidencia a Google Drive en la subcarpeta `{nombre_subcarpeta_docente}`..."
                            ):
                                ok_up, url_up, msg_up = subir_soporte_drive(
                                    uploaded_file=apo_archivo,
                                    documento=documento_docente,
                                    carpeta_categoria=f"{nombre_subcarpeta_docente}/{apo_categoria}",
                                    etiqueta_archivo=f"Aporte_{apo_categoria}",
                                    root_folder_id=FOLDER_ID_SOPORTES_APORTES,
                                )
                            if ok_up and url_up:
                                url_soporte_apo = url_up
                            else:
                                st.error(f"❌ {msg_up}")
                                fallo_drive_apo = True

                    if not fallo_drive_apo:
                        ok_ins, msg_ins = agregar_aporte_docente(
                            docente_id=docente_id,
                            categoria=apo_categoria,
                            titulo_aporte=apo_titulo,
                            descripcion=apo_descripcion,
                            periodo_academico=apo_periodo,
                            soporte_url=url_soporte_apo,
                        )
                        if ok_ins:
                            st.toast(msg_ins, icon="🚀")
                            st.rerun()
                        else:
                            st.error(f"❌ {msg_ins}")

    with col_list_apo:
        c_tit_apo, c_btn_repo_apo = st.columns([2, 1.3])
        with c_tit_apo:
            st.markdown("##### 🏆 Aportes Registrados del Docente")
        with c_btn_repo_apo:
            st.link_button(
                "📂 Abrir Repositorio Drive de Aportes",
                url=URL_CARPETA_DRIVE_APORTES,
                use_container_width=True,
            )
        if not aportes_actuales:
            st.info("Aún no hay aportes registrados para este docente.")
        else:
            for ap in aportes_actuales:
                with st.container(border=True):
                    c_info, c_link, c_del = st.columns([3.2, 1.3, 0.7])
                    with c_info:
                        st.markdown(f"**[{ap.get('categoria')}]** {ap.get('titulo_aporte')}")
                        st.caption(f"🗓️ Periodo: `{ap.get('periodo_academico')}`")
                        if ap.get("descripcion"):
                            st.write(ap.get("descripcion"))
                    with c_link:
                        if ap.get("soporte_url"):
                            st.link_button("📂 Ver Evidencia Drive", url=str(ap["soporte_url"]), use_container_width=True)
                        else:
                            st.caption("Sin evidencia adjunta")
                    with c_del:
                        if st.button("🗑️", key=f"del_apo_{ap['id']}", help="Eliminar este aporte", use_container_width=True):
                            ok_del, msg_del = eliminar_aporte_docente(ap["id"])
                            if ok_del:
                                st.toast(msg_del, icon="🗑️")
                                st.rerun()
                            else:
                                st.error(msg_del)

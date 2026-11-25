"""
Módulo 1 (Fase 1 - MVP): Panel Gerencial de Resoluciones Docentes, Soportes en Drive y Reportes PDF.
Adaptado para aprovechar todas las dimensiones de `Resoluciones Docentes.xlsx`:
- KPIs de Docentes, Horas Contratadas, Valor Total Vinculación, Posgrados, Inglés y Aportes.
- Gráficas interactivas por Nivel de Inglés, Centro (CEAD/CCAV/UDR), Nivel de Formación y Categoría de Aporte.
- Filtros combinados por Nivel de Inglés, Centro, Programa, Modalidad y Categoría de Aporte.
- Descarga en memoria (`BytesIO`) del Informe Consolidado Gerencial y de la Ficha Individual Docente en PDF.
"""

import importlib
from pathlib import Path
import sys

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import pandas as pd
import plotly.express as px
import streamlit as st
from src.db.client import verificar_configuracion_supabase
import src.services.docentes_svc as _doc_svc
import src.services.pdf_svc as _pdf_svc
import src.utils.auth as _auth_mod

importlib.reload(_doc_svc)
importlib.reload(_pdf_svc)
importlib.reload(_auth_mod)

CARPETA_DRIVE_SOPORTES_INGLES_URL = _doc_svc.CARPETA_DRIVE_SOPORTES_INGLES_URL
CATEGORIAS_APORTE = _doc_svc.CATEGORIAS_APORTE
NIVELES_INGLES = _doc_svc.NIVELES_INGLES
limpiar_cache_docentes = _doc_svc.limpiar_cache_docentes
listar_docentes_con_relaciones = _doc_svc.listar_docentes_con_relaciones
obtener_estadisticas_consolidadas = _doc_svc.obtener_estadisticas_consolidadas
generar_pdf_ficha_docente = _pdf_svc.generar_pdf_ficha_docente
generar_pdf_informe_consolidado = _pdf_svc.generar_pdf_informe_consolidado
es_admin = _auth_mod.es_admin
renderizar_barra_sesion_y_login = _auth_mod.renderizar_barra_sesion_y_login

st.set_page_config(
    page_title="Panel Gerencial y Reportes PDF",
    page_icon="📊",
    layout="wide",
)

c_title, c_refresh = st.columns([4, 1])
with c_title:
    st.title("📊 Panel Gerencial de Resoluciones Docentes y Reportes PDF")
    st.markdown(
        "Analiza la planta docente por **Centro, Programa, Modalidad, Formación y Nivel de Inglés**, "
        "consulta los soportes en **Google Drive Institucional** y descarga los informes oficiales en **PDF**."
    )
with c_refresh:
    st.write("")
    if st.button("🔄 Actualizar Datos", use_container_width=True):
        limpiar_cache_docentes()
        st.rerun()

if not renderizar_barra_sesion_y_login():
    st.stop()

usuario_es_admin = es_admin()

ok_db, msg_db = verificar_configuracion_supabase()
if not ok_db:
    st.error(f"⚠️ Base de datos no configurada: {msg_db}")
    st.stop()

stats = obtener_estadisticas_consolidadas()

# =============================================================================
# 1. MÉTRICAS CLAVE (KPIs) Y ACCESO A CARPETA DE DRIVE
# =============================================================================
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("👨‍🏫 Total Docentes", stats["total_docentes"])
k2.metric("⏱️ Total Horas Vinculación", f"{stats['total_horas_vinculacion']:,}".replace(",", "."))
k3.metric("🎓 Títulos de Posgrado", stats["total_posgrados"])
k4.metric("🌐 Con Soporte Inglés (Drive)", stats["docentes_con_soporte_ingles"])
k5.metric("🚀 Total Aportes", stats["total_aportes"])

if usuario_es_admin:
    c_dr1, c_dr2 = st.columns(2)
    with c_dr1:
        st.link_button(
            "📂 Abrir Carpeta Drive: Certificados de Inglés",
            url=CARPETA_DRIVE_SOPORTES_INGLES_URL,
            use_container_width=True,
        )
    with c_dr2:
        st.link_button(
            "📂 Abrir Carpeta Drive: Repositorio de Aportes Docentes",
            url="https://drive.google.com/drive/folders/1HSLyyW8LzqKmSLvlhpj6-fZHeTgkfBvP?usp=drive_link",
            use_container_width=True,
        )

st.divider()

# =============================================================================
# 2. VISUALIZACIÓN ESTADÍSTICA (PLOTLY)
# =============================================================================
col_chart_ing, col_chart_cen = st.columns(2)

with col_chart_ing:
    df_ing = pd.DataFrame(
        list(stats["distribucion_ingles"].items()),
        columns=["Nivel de Inglés", "Docentes"],
    )
    fig_ing = px.bar(
        df_ing,
        x="Nivel de Inglés",
        y="Docentes",
        text="Docentes",
        title="Distribución de Docentes por Nivel de Inglés (MCER)",
        color="Nivel de Inglés",
        color_discrete_sequence=px.colors.sequential.Blues_r,
    )
    fig_ing.update_layout(showlegend=False, height=330, margin=dict(t=45, b=20, l=20, r=20))
    st.plotly_chart(fig_ing, use_container_width=True)

with col_chart_cen:
    dist_centro = stats.get("distribucion_centro") or {}
    if dist_centro:
        df_cen = pd.DataFrame(
            sorted(dist_centro.items(), key=lambda x: x[1], reverse=True),
            columns=["Centro", "Docentes"],
        )
        fig_cen = px.bar(
            df_cen,
            x="Docentes",
            y="Centro",
            orientation="h",
            text="Docentes",
            title="Docentes Vinculados por Centro (CEAD / CCAV / UDR)",
            color="Centro",
            color_discrete_sequence=px.colors.qualitative.Safe,
        )
        fig_cen.update_layout(showlegend=False, height=330, margin=dict(t=45, b=20, l=20, r=20))
        st.plotly_chart(fig_cen, use_container_width=True)
    else:
        df_apo = pd.DataFrame(
            list(stats["distribucion_aportes"].items()),
            columns=["Categoría", "Aportes"],
        )
        fig_apo = px.bar(
            df_apo,
            x="Categoría",
            y="Aportes",
            text="Aportes",
            title="Total de Aportes por Categoría Institucional",
            color="Categoría",
        )
        fig_apo.update_layout(showlegend=False, height=330, margin=dict(t=45, b=20, l=20, r=20))
        st.plotly_chart(fig_apo, use_container_width=True)

st.divider()

# =============================================================================
# 3. SECCIÓN EXCLUSIVA PARA ADMINISTRADOR: DIRECTORIO DE DOCENTES Y RESOLUCIONES (101 RESULTADOS)
# =============================================================================
if not usuario_es_admin:
    st.info(
        "🔒 **Acceso Restringido:** La consulta del **Directorio de Docentes y Resoluciones (101 resultados)** "
        "está reservada exclusivamente para el usuario **Administrador (`admin`)**. "
        "Como docente, puedes consultar y descargar tu propia Ficha Individual en la pestaña **1. Registro Docente**."
    )
    st.stop()

st.markdown("### 🔎 Filtros de Consulta y Exportación Consolidada (Exclusivo Administrador)")

f1, f2, f3, f4, f5 = st.columns(5)

with f1:
    filtro_ingles = st.selectbox("Nivel de Inglés", options=["Todos"] + NIVELES_INGLES)

with f2:
    filtro_centro = st.selectbox(
        "Centro (CEAD/CCAV/UDR)",
        options=["Todos"] + stats.get("centros_disponibles", []),
    )

with f3:
    filtro_programa = st.selectbox(
        "Programa Académico",
        options=["Todos"] + stats.get("programas_disponibles", []),
    )

with f4:
    filtro_modalidad = st.selectbox(
        "Modalidad",
        options=["Todas"] + stats.get("modalidades_disponibles", []),
    )

with f5:
    filtro_categoria = st.selectbox(
        "Categoría de Aporte",
        options=["Todas"] + CATEGORIAS_APORTE,
    )

c_busq, c_btn_pdf = st.columns([3, 1.3])
with c_busq:
    busqueda_txt = st.text_input(
        "Buscar por Nombre, Documento, Correo o N° de Resolución",
        placeholder="Ej: QUIROGA, 72142321, 164...",
    )

docentes_filtrados = listar_docentes_con_relaciones(
    filtro_nivel_ingles=filtro_ingles,
    filtro_categoria_aporte=filtro_categoria,
    filtro_modalidad=filtro_modalidad,
    filtro_centro=filtro_centro,
    filtro_programa=filtro_programa,
    busqueda_texto=busqueda_txt,
    limite=150,
)

with c_btn_pdf:
    st.write("")
    pdf_consolidado_bytes = generar_pdf_informe_consolidado(
        stats=stats,
        docentes_resumen=docentes_filtrados,
    )
    st.download_button(
        label="📥 Descargar Informe Consolidado (PDF)",
        data=pdf_consolidado_bytes,
        file_name="Informe_Consolidado_Resoluciones_Docentes.pdf",
        mime="application/pdf",
        type="primary",
        use_container_width=True,
    )

# =============================================================================
# 4. TABLA DE DOCENTES Y RESOLUCIONES CON ENLACES A DRIVE
# =============================================================================
st.markdown(f"#### 📋 Directorio de Docentes y Resoluciones ({len(docentes_filtrados)} resultados)")

if not docentes_filtrados:
    st.info(
        "No se encontraron docentes con los criterios seleccionados. "
        "Si aún no has cargado la base inicial, ve a **1. Registro Docente** y haz clic en **Sincronizar `Resoluciones Docentes.xlsx`**."
    )
else:
    filas_tabla = []
    for d in docentes_filtrados:
        estudios_d = d.get("docente_estudios") or []
        aportes_d = d.get("docente_aportes") or []
        filas_tabla.append(
            {
                "Documento": d.get("documento", ""),
                "Nombre Completo": d.get("nombre_completo", ""),
                "Formación": d.get("nivel_formacion_maximo") or "-",
                "Nivel Inglés": d.get("nivel_ingles", "Sin certificar"),
                "Certificado Inglés": d.get("soporte_ingles_url") or "",
                "Centro": d.get("centro_descripcion") or "-",
                "Programa": d.get("programa_descripcion") or "-",
                "Modalidad": d.get("modalidad") or "-",
                "Resolución": d.get("resolucion_numero") or "-",
                "Horas": d.get("horas_vinculacion") or 0,
                "Títulos": len(estudios_d),
                "Aportes": len(aportes_d),
            }
        )

    df_tabla = pd.DataFrame(filas_tabla)
    st.dataframe(
        df_tabla,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Certificado Inglés": st.column_config.LinkColumn(
                "Certificado Inglés",
                help="Haz clic para abrir el soporte de inglés en Google Drive",
                display_text="📂 Abrir Drive",
            ),
        },
    )

    st.divider()

    # =========================================================================
    # 5. INSPECCIÓN DETALLADA Y DESCARGA DE FICHA INDIVIDUAL PDF
    # =========================================================================
    st.markdown("### 📄 Ficha Individual Docente, Detalle de Resolución y Soportes en Google Drive")

    opciones_docentes = {
        f"{d.get('nombre_completo')} — Doc: {d.get('documento')} | {d.get('centro_descripcion') or 'Sin centro'} | Inglés: {d.get('nivel_ingles')}": d
        for d in docentes_filtrados
    }

    col_sel, col_btn_ficha = st.columns([3, 1])
    with col_sel:
        etiqueta_sel = st.selectbox(
            "Selecciona un docente para ver su resolución completa, abrir sus soportes en Drive o descargar su Ficha PDF:",
            options=list(opciones_docentes.keys()),
        )

    docente_sel = opciones_docentes[etiqueta_sel]
    estudios_sel = docente_sel.get("docente_estudios") or []
    aportes_sel = docente_sel.get("docente_aportes") or []

    with col_btn_ficha:
        st.write("")
        pdf_individual_bytes = generar_pdf_ficha_docente(
            docente=docente_sel,
            estudios=estudios_sel,
            aportes=aportes_sel,
        )
        st.download_button(
            label="📄 Descargar Ficha Individual (PDF)",
            data=pdf_individual_bytes,
            file_name=f"Ficha_Docente_{docente_sel.get('documento')}.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True,
        )

    with st.container(border=True):
        c_d1, c_d2, c_d3 = st.columns([1.4, 1.5, 1.1])
        with c_d1:
            st.markdown(f"**👤 Docente:** {docente_sel.get('nombre_completo')}")
            st.markdown(f"**🪪 Documento:** `{docente_sel.get('documento')}`")
            st.markdown(f"**✉️ Correo Inst.:** {docente_sel.get('correo_institucional')}")
            st.markdown(f"**📞 Teléfono / Email Pers.:** {docente_sel.get('telefono') or '-'} | {docente_sel.get('email_personal') or '-'}")
        with c_d2:
            st.markdown(f"**🏛️ Centro / Zona:** {docente_sel.get('centro_descripcion') or '-'} ({docente_sel.get('zona_descripcion') or '-'})")
            st.markdown(f"**📚 Programa:** {docente_sel.get('programa_descripcion') or '-'}")
            st.markdown(
                f"**📜 Resolución:** N° `{docente_sel.get('resolucion_numero') or '-'}` del `{docente_sel.get('resolucion_fecha') or '-'}` "
                f"| **Modalidad:** `{docente_sel.get('modalidad') or '-'}` (`{docente_sel.get('horas_vinculacion') or 0}` hrs)"
            )
            st.markdown(
                f"**🗓️ Vigencia:** `{docente_sel.get('fecha_inicio') or '-'}` al `{docente_sel.get('fecha_termina') or '-'}` "
                f"| **Periodo:** `{docente_sel.get('periodo_academico') or '-'}`"
            )
        with c_d3:
            st.markdown(f"**🎓 Nivel Formación:** `{docente_sel.get('nivel_formacion_maximo') or '-'}`")
            st.markdown(
                f"**🌐 Inglés (MCER):** `{docente_sel.get('nivel_ingles')}` "
                f"({docente_sel.get('detalle_ingles_original') or '-'})"
            )
            if docente_sel.get("soporte_ingles_url"):
                st.link_button(
                    "📂 Abrir Certificado de Inglés en Drive",
                    url=str(docente_sel["soporte_ingles_url"]),
                    use_container_width=True,
                )
            else:
                st.caption("Sin soporte de inglés adjunto en Drive.")

        st.markdown("---")
        col_est_det, col_apo_det = st.columns(2)

        with col_est_det:
            st.markdown(f"##### 🎓 Títulos Académicos ({len(estudios_sel)})")
            if not estudios_sel:
                st.caption("Sin títulos desglosados.")
            else:
                for e in estudios_sel:
                    with st.container(border=True):
                        st.markdown(f"**{e.get('nivel_formacion')}**: {e.get('titulo_obtenido')}")
                        st.caption(f"🏛️ {e.get('universidad')} | Año: {e.get('anio_graduacion') or '-'}")

        with col_apo_det:
            st.markdown(f"##### 🚀 Aportes Académicos ({len(aportes_sel)})")
            if not aportes_sel:
                st.caption("Sin aportes registrados.")
            else:
                for a in aportes_sel:
                    with st.container(border=True):
                        st.markdown(
                            f"**[{a.get('categoria')}]** {a.get('titulo_aporte')} (`{a.get('periodo_academico')}`)"
                        )
                        if a.get("descripcion"):
                            st.caption(a.get("descripcion"))
                        if a.get("soporte_url"):
                            st.markdown(f"🔗 [Abrir Evidencia en Google Drive]({a['soporte_url']})")
                        else:
                            st.caption("Sin evidencia adjunta")

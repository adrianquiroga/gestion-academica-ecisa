"""
Módulo 2 (Fase 2 Activa): Gestión y Analítica de Matrículas por Periodo (ACTAS 2025 y ACTAS 2026).
Diseñado siguiendo estrictamente las reglas de rendimiento para +15.000 registros y límite de 1 GB RAM:
- Lectura en streaming con `openpyxl(..., read_only=True)` saltando las 13 filas de título del reporte UNAD.
- Inserción por lotes (`Batching` de 750 filas por petición) con barra de progreso en vivo.
- Consultas paginadas (máximo 100 filas en pantalla) y agregaciones en caché (`@st.cache_data(ttl=300)`).
- Comparativo 2025 vs 2026, distribución por Centro, Programa, Periodo, Condición (Nuevos vs Antiguos) y Estrato.
- Exportación de Informe Ejecutivo de Matrículas en PDF (`fpdf2` en memoria).
"""

import importlib
import math
import pandas as pd
import plotly.express as px
import streamlit as st
from src.db.client import verificar_configuracion_supabase
from src.services.matriculas_svc import (
    consultar_matriculas_paginadas,
    detectar_archivos_actas_locales,
    importar_archivo_actas_a_supabase,
    limpiar_cache_matriculas,
    obtener_estadisticas_matriculas,
)
from src.services.pdf_svc import generar_pdf_informe_matriculas
import src.utils.auth as _auth_mod

importlib.reload(_auth_mod)
es_admin = _auth_mod.es_admin
renderizar_barra_sesion_y_login = _auth_mod.renderizar_barra_sesion_y_login

st.set_page_config(
    page_title="Módulo de Matrículas 2025-2026",
    page_icon="📈",
    layout="wide",
)

c_title, c_refresh = st.columns([4, 1])
with c_title:
    st.title("📈 Gestión y Analítica de Matrículas (Actas 2025 y 2026)")
    st.markdown(
        "Módulo de alto rendimiento optimizado para **+15.000 registros de matrícula** (ECISA - Zona Centro Bogotá Cundinamarca), "
        "con carga masiva por lotes (*Batching*), paginación SQL y exportación gerencial en **PDF**."
    )
with c_refresh:
    st.write("")
    if st.button("🔄 Actualizar Datos", use_container_width=True):
        limpiar_cache_matriculas()
        st.rerun()

if not renderizar_barra_sesion_y_login():
    st.stop()

usuario_es_admin = es_admin()

ok_db, msg_db = verificar_configuracion_supabase()
if not ok_db:
    st.error(f"⚠️ Base de datos no configurada: {msg_db}")
    st.stop()

# =============================================================================
# 0. PANEL DESPLEGABLE: CARGA MASIVA POR LOTES (BATCHING) — SOLO ADMIN
# =============================================================================
if usuario_es_admin:
    with st.expander("📥 Carga Masiva por Lotes (Batching) — `ACTAS 2025.xlsx` y `ACTAS 2026.xlsx` (Solo Admin)", expanded=False):
        st.markdown(
            """
            Este importador lee automáticamente los reportes oficiales de **Actas de Matrícula UNAD** (detectando los encabezados en la fila 14)
            e inserta los registros en Supabase en **bloques de 750 filas** para proteger el límite de 1 GB de RAM de Streamlit Cloud.
            """
        )

        col_local, col_upload = st.columns(2)

        with col_local:
            archivos_locales = detectar_archivos_actas_locales()
            if archivos_locales:
                nombres_str = ", ".join([f"`{n}` ({a})" for n, a in archivos_locales])
                st.info(f"📂 Archivos detectados en la carpeta del proyecto: {nombres_str}")
                if st.button(
                    "🚀 Sincronizar `ACTAS 2025.xlsx` y `ACTAS 2026.xlsx` por Lotes",
                    type="primary",
                    use_container_width=True,
                ):
                    barra = st.progress(0.0, text="Iniciando lectura optimizada...")
                    resumen_mensajes = []
                    for idx_arch, (nombre_arch, anio_arch) in enumerate(archivos_locales, start=1):
                        def _actualizar_progreso(actual: int, total: int, na: str = nombre_arch) -> None:
                            frac = min(actual / max(total, 1), 1.0)
                            barra.progress(
                                frac,
                                text=f"Sincronizando `{na}`: {actual:,} de {total:,} registros...".replace(",", "."),
                            )

                        ok_imp, n_ins, msg_imp = importar_archivo_actas_a_supabase(
                            fuente_archivo=nombre_arch,
                            anio_forzado=anio_arch,
                            progreso_callback=_actualizar_progreso,
                        )
                        if ok_imp:
                            resumen_mensajes.append(f"✅ **`{nombre_arch}`**: {msg_imp}")
                        else:
                            st.error(f"❌ **`{nombre_arch}`**: {msg_imp}")
                    barra.empty()
                    for m in resumen_mensajes:
                        st.success(m)
            else:
                st.warning("No se encontraron `ACTAS 2025.xlsx` o `ACTAS 2026.xlsx` en la carpeta local.")

        with col_upload:
            archivo_subido = st.file_uploader(
                "O sube un nuevo archivo Excel de Actas de Matrícula (.xlsx):",
                type=["xlsx"],
                key="uploader_actas_matriculas",
            )
            anio_manual = st.selectbox("Año del Acta para el archivo subido:", options=[2026, 2025, 2027], index=0)
            if archivo_subido is not None:
                if st.button("📥 Importar Archivo Subido por Lotes", use_container_width=True):
                    barra_up = st.progress(0.0, text="Procesando archivo subido...")
                    ok_up, n_up, msg_up = importar_archivo_actas_a_supabase(
                        fuente_archivo=archivo_subido,
                        anio_forzado=int(anio_manual),
                        progreso_callback=lambda act, tot: barra_up.progress(
                            min(act / max(tot, 1), 1.0),
                            text=f"Insertando lote: {act:,} / {tot:,} registros...".replace(",", "."),
                        ),
                    )
                    barra_up.empty()
                    if ok_up:
                        st.success(f"✅ {msg_up}")
                    else:
                        st.error(f"❌ {msg_up}")

# =============================================================================
# 1. SELECTOR GLOBAL DE AÑO Y MÉTRICAS CLAVE (KPIs)
# =============================================================================
col_filtro_global, col_pdf_global = st.columns([3, 1.3])
with col_filtro_global:
    filtro_anio_kpi = st.radio(
        "Seleccionar Vista de Año de Actas:",
        options=["Todos", "2025", "2026"],
        horizontal=True,
    )

stats = obtener_estadisticas_matriculas(filtro_anio=filtro_anio_kpi)

with col_pdf_global:
    st.write("")
    pdf_matriculas_bytes = generar_pdf_informe_matriculas(
        stats=stats,
        filtro_anio=filtro_anio_kpi,
    )
    st.download_button(
        label="📥 Descargar Informe Matrículas (PDF)",
        data=pdf_matriculas_bytes,
        file_name=f"Informe_Matriculas_{filtro_anio_kpi}.pdf",
        mime="application/pdf",
        type="primary",
        use_container_width=True,
    )

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("🗂️ Total Actas de Matrícula", f"{stats['total_matriculas']:,}".replace(",", "."))
k2.metric("👨‍🎓 Estudiantes Únicos", f"{stats['estudiantes_unicos']:,}".replace(",", "."))
k3.metric("🔄 Estudiantes Antiguos", f"{stats['total_antiguos']:,}".replace(",", "."))
k4.metric("✨ Estudiantes Nuevos", f"{stats['total_nuevos']:,}".replace(",", "."))
k5.metric("📌 Continuidad / Aplazados", f"{stats['total_continuidad_otros']:,}".replace(",", "."))

if stats["total_matriculas"] == 0:
    st.info(
        "💡 Aún no hay registros en la tabla `matriculas` para el filtro seleccionado. "
        "Abre el panel superior **'📥 Carga Masiva por Lotes'** para sincronizar `ACTAS 2025.xlsx` y `ACTAS 2026.xlsx`."
    )
    st.stop()

st.divider()

# =============================================================================
# 2. GRÁFICAS EJECUTIVAS (PLOTLY)
# =============================================================================
g1, g2 = st.columns(2)

with g1:
    df_centro = pd.DataFrame(
        sorted(stats["por_centro"].items(), key=lambda x: x[1], reverse=True),
        columns=["Centro", "Matrículas"],
    )
    fig_centro = px.bar(
        df_centro,
        x="Matrículas",
        y="Centro",
        orientation="h",
        text="Matrículas",
        title="Matrículas por Centro Regional (CEAD / CCAV / UDR)",
        color="Centro",
        color_discrete_sequence=px.colors.qualitative.Safe,
    )
    fig_centro.update_layout(showlegend=False, height=350, margin=dict(t=45, b=20, l=20, r=20))
    st.plotly_chart(fig_centro, use_container_width=True)

with g2:
    df_prog = pd.DataFrame(
        sorted(stats["por_programa"].items(), key=lambda x: x[1], reverse=True),
        columns=["Programa Académico", "Matrículas"],
    )
    fig_prog = px.bar(
        df_prog,
        x="Matrículas",
        y="Programa Académico",
        orientation="h",
        text="Matrículas",
        title="Matrículas por Programa Académico (ECISA)",
        color="Programa Académico",
        color_discrete_sequence=px.colors.sequential.Blues_r,
    )
    fig_prog.update_layout(showlegend=False, height=350, margin=dict(t=45, b=20, l=20, r=20))
    st.plotly_chart(fig_prog, use_container_width=True)

g3, g4, g5 = st.columns(3)

with g3:
    df_per = pd.DataFrame(
        sorted(stats["por_periodo"].items(), key=lambda x: x[1], reverse=True)[:10],
        columns=["Periodo", "Matrículas"],
    )
    fig_per = px.bar(
        df_per,
        x="Periodo",
        y="Matrículas",
        text="Matrículas",
        title="Top Periodos Académicos",
        color="Periodo",
    )
    fig_per.update_layout(showlegend=False, height=300, margin=dict(t=45, b=20, l=20, r=20))
    st.plotly_chart(fig_per, use_container_width=True)

with g4:
    df_cond = pd.DataFrame(
        list(stats["por_condicion"].items()),
        columns=["Condición", "Matrículas"],
    )
    fig_cond = px.pie(
        df_cond,
        names="Condición",
        values="Matrículas",
        hole=0.45,
        title="Distribución por Condición (Antiguo / Nuevo)",
    )
    fig_cond.update_layout(height=300, margin=dict(t=45, b=20, l=20, r=20))
    st.plotly_chart(fig_cond, use_container_width=True)

with g5:
    df_est = pd.DataFrame(
        sorted(stats["por_estrato"].items()),
        columns=["Estrato", "Matrículas"],
    )
    fig_est = px.bar(
        df_est,
        x="Estrato",
        y="Matrículas",
        text="Matrículas",
        title="Distribución por Estrato Socioeconómico",
        color="Estrato",
    )
    fig_est.update_layout(showlegend=False, height=300, margin=dict(t=45, b=20, l=20, r=20))
    st.plotly_chart(fig_est, use_container_width=True)

st.divider()

# =============================================================================
# 3. EXPLORADOR PAGINADO DE ESTUDIANTES — EXCLUSIVO ADMINISTRADOR
# =============================================================================
if not usuario_es_admin:
    st.info(
        "🔒 **Acceso Restringido:** El **Explorador Paginado de Actas de Matrícula (Consultas SQL Filtradas)** "
        "con datos individuales de estudiantes está habilitado exclusivamente para el usuario **Administrador (`admin`)**."
    )
    st.stop()

st.markdown("### 🔎 Explorador Paginado de Actas de Matrícula (Consultas SQL Filtradas — Exclusivo Admin)")

f1, f2, f3, f4, f5 = st.columns([1, 1.2, 1.4, 1.1, 1.5])

with f1:
    sel_periodo = st.selectbox("Periodo", options=["Todos"] + stats["periodos_disponibles"])
with f2:
    sel_centro = st.selectbox("Centro", options=["Todos"] + stats["centros_disponibles"])
with f3:
    sel_programa = st.selectbox("Programa", options=["Todos"] + stats["programas_disponibles"])
with f4:
    sel_condicion = st.selectbox("Condición", options=["Todas"] + stats["condiciones_disponibles"])
with f5:
    busq_est = st.text_input(
        "Buscar Estudiante (Documento, Nombre o Correo)",
        placeholder="Ej: 1023026660 o RAMIREZ...",
    )

col_pag_info, col_pag_num = st.columns([3.5, 1])
with col_pag_num:
    pagina_actual = st.number_input("Página (100 filas/pág.)", min_value=1, value=1, step=1)

filas_pagina, total_filtrado = consultar_matriculas_paginadas(
    filtro_anio=filtro_anio_kpi,
    filtro_periodo=sel_periodo,
    filtro_centro=sel_centro,
    filtro_programa=sel_programa,
    filtro_condicion=sel_condicion,
    busqueda_estudiante=busq_est,
    pagina=int(pagina_actual),
    filas_por_pagina=100,
)

total_paginas = max(math.ceil(total_filtrado / 100), 1)
with col_pag_info:
    st.markdown(
        f"**Resultados encontrados:** `{total_filtrado:,}` actas | "
        f"Mostrando página **`{pagina_actual}`** de **`{total_paginas}`** (máximo 100 registros en memoria)".replace(",", ".")
    )

if not filas_pagina:
    st.warning("No se encontraron registros para los filtros o número de página seleccionados.")
else:
    df_detalle = pd.DataFrame(filas_pagina)
    df_detalle = df_detalle.rename(
        columns={
            "anio_acta": "Año",
            "periodo": "Periodo",
            "tipo_documento": "Tipo Doc",
            "documento_estudiante": "Documento",
            "nombre_estudiante": "Nombre Estudiante",
            "correo_institucional": "Correo Institucional",
            "programa_academico": "Programa Académico",
            "centro_zona": "Centro",
            "tipo_matricula": "Condición",
            "cursos_matriculados": "Cursos Matr.",
            "cursos_aprobados": "Cursos Aprob.",
            "promedio": "Promedio",
            "resultado": "Resultado",
            "estrato": "Estrato",
            "sexo": "Sexo",
            "ciudad_residencia": "Ciudad",
        }
    )
    st.dataframe(df_detalle, use_container_width=True, hide_index=True)

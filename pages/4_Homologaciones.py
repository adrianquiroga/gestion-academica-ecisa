"""
Módulo 4 (Fase 3 Activa): Seguimiento de Homologaciones y Tiempos de Respuesta 2026 (ECISA - ZCBC).
Alimentado por el archivo oficial `Reporte_Tiempos_20261123_093339` (01/01/2026 al 23/11/2026):
- KPIs de Estudios de Homologación (112 registros en ZCBC), Tipos (`int` Interna, `egr` Egresado/Convenio, `ext` Externa),
  y Tiempos Promedio de Respuesta (Tiempo Total, Días Calendario, Tiempo Actores, Tiempo Administración y Tiempo Envío).
- Analítica de tiempos por cada actor del flujo:
  1. Líder Nacional de Homologación
  2. Tutor Encargado
  3. Líder Nacional de Programa
- Gráficas interactivas por Programa Académico, Centro Regional (JAG, Zipaquirá, Soacha, Facatativá, Gachetá, Girardot, Fusagasugá, Chipaque),
  Tipo de Homologación, Tutor Encargado y Evolución Mensual.
- Descarga de Informe Ejecutivo de Seguimiento y Tiempos de Homologación en PDF (`fpdf2` en memoria).
"""

import importlib
from pathlib import Path
import sys

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from src.db.client import verificar_configuracion_supabase
import src.services.homologaciones_svc as _hom_svc
import src.services.pdf_svc as _pdf_svc
import src.utils.auth as _auth_mod

importlib.reload(_hom_svc)
importlib.reload(_pdf_svc)
importlib.reload(_auth_mod)

calcular_estadisticas_homologaciones = _hom_svc.calcular_estadisticas_homologaciones
detectar_archivo_reporte_tiempos_local = _hom_svc.detectar_archivo_reporte_tiempos_local
limpiar_cache_homologaciones = _hom_svc.limpiar_cache_homologaciones
obtener_datos_homologaciones_2026 = _hom_svc.obtener_datos_homologaciones_2026
sincronizar_homologaciones_a_supabase = _hom_svc.sincronizar_homologaciones_a_supabase
generar_pdf_informe_homologaciones = _pdf_svc.generar_pdf_informe_homologaciones
es_admin = _auth_mod.es_admin
renderizar_barra_sesion_y_login = _auth_mod.renderizar_barra_sesion_y_login

st.set_page_config(
    page_title="Seguimiento y Tiempos de Homologación 2026 - ECISA",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded",
)

if not renderizar_barra_sesion_y_login("4_homologaciones"):
    st.stop()

df_base, meta_rep = obtener_datos_homologaciones_2026()
f_ini_corte = meta_rep.get("fecha_inicial_reporte", "01/01/2026")
f_fin_corte = meta_rep.get("fecha_final_reporte", "23/11/2026")

c_title, c_refresh = st.columns([4, 1])
with c_title:
    st.title("🔄 Seguimiento y Tiempos de Homologación 2026 (ECISA - ZCBC)")
    st.markdown(
        f"**Informe Oficial de Seguimiento de Tiempos de Homologación** (Corte: **`{f_ini_corte}`** al **`{f_fin_corte}`** — "
        f"Zona Centro Bogotá y Cundinamarca): análisis por **Actor (Líder Homologación, Tutor Encargado, Líder Programa), "
        f"Programa Académico, Centro Regional y Tipo (`int`, `egr`, `ext`)**."
    )
with c_refresh:
    st.write("")
    if st.button("🔄 Actualizar Datos", use_container_width=True):
        limpiar_cache_homologaciones()
        st.rerun()

usuario_es_admin = es_admin()

ok_db, msg_db = verificar_configuracion_supabase()
if not ok_db:
    st.warning(f"⚠️ Aviso de conexión a Supabase: {msg_db}. Visualizando desde caché sincronizada.")

# =============================================================================
# 0. PANEL DESPLEGABLE: SINCRONIZACIÓN DE `Reporte_Tiempos_20261123_093339` (SOLO ADMIN)
# =============================================================================
if usuario_es_admin:
    with st.expander(
        "📥 Cargar / Sincronizar `Reporte_Tiempos_20261123_093339` con Supabase (Solo Admin)",
        expanded=df_base.empty,
    ):
        st.markdown(
            """
            Este importador procesa directamente el reporte oficial del sistema de homologaciones (**`Reporte_Tiempos_20261123_093339`** en formato `.csv` o `.xlsx`),
            corrigiendo tildes/UTF-8 y desglosando los **4 actores del proceso** (*Estudiante, Líder Nacional de Homologación, Tutor Encargado y Líder Nacional de Programa*)
            junto con los indicadores de **`Tiempo actores`**, **`Tiempo admon`**, **`Tiempo Envio`**, **`Tiempo Total`** y **`dias`**.
            """
        )
        col_loc, col_up = st.columns(2)

        with col_loc:
            ruta_local = detectar_archivo_reporte_tiempos_local()
            if ruta_local is not None:
                st.info(f"📂 Archivo detectado en `datos_locales`: **`{ruta_local}`**")
                if st.button(
                    f"🚀 Sincronizar `{ruta_local.name}` con Supabase",
                    type="primary",
                    use_container_width=True,
                ):
                    with st.spinner("Sincronizando Reporte de Tiempos de Homologación con Supabase..."):
                        ok_s, n_s, msg_s = sincronizar_homologaciones_a_supabase(None, reemplazar_existentes=True)
                    if ok_s:
                        st.success(f"✅ {msg_s}")
                        st.rerun()
                    else:
                        st.warning(f"⚠️ {msg_s}")
                        st.rerun()
            else:
                st.info("Sube un archivo `Reporte_Tiempos` en el panel de la derecha 👉.")

        with col_up:
            archivo_subido = st.file_uploader(
                "O sube un nuevo archivo `Reporte_Tiempos` (.csv o .xlsx):",
                type=["csv", "xlsx", "xls"],
                key="uploader_reporte_tiempos_hom",
            )
            if archivo_subido is not None:
                if st.button("📥 Procesar y Guardar Reporte Subido", type="primary", use_container_width=True):
                    with st.spinner("Procesando archivo subido y actualizando Supabase..."):
                        ok_u, n_u, msg_u = sincronizar_homologaciones_a_supabase(
                            archivo_subido,
                            reemplazar_existentes=True,
                        )
                    if ok_u:
                        st.success(f"✅ {msg_u}")
                        st.rerun()
                    else:
                        st.warning(f"⚠️ {msg_u}")
                        st.rerun()

if df_base.empty:
    st.info(
        "💡 Aún no hay registros cargados. Abre el panel superior para sincronizar `Reporte_Tiempos_20261123_093339`."
    )
    st.stop()

# =============================================================================
# 1. FILTROS COMBINADOS Y DESCARGA DE INFORME EJECUTIVO PDF
# =============================================================================
st.markdown("### 🔎 Filtros de Consulta por Tipo, Programa, Centro, Tutor y Líder de Programa")

f1, f2, f3, f4, f5 = st.columns([1.05, 1.45, 1.25, 1.25, 1.25])

tipos_disp = ["Todos"] + sorted(df_base["tipo_homologacion"].dropna().unique().tolist())
progs_disp = ["Todos"] + sorted(df_base["programa_destino"].dropna().unique().tolist())
centros_disp = ["Todos"] + sorted(df_base["centro"].dropna().unique().tolist())
tutores_disp = ["Todos"] + sorted(df_base["tutor_nombre"].dropna().unique().tolist())
lideres_disp = ["Todos"] + sorted(df_base["lider_programa_nombre"].dropna().unique().tolist())

with f1:
    sel_tipo = st.selectbox("Tipo (`int` / `egr` / `ext`)", options=tipos_disp)
with f2:
    sel_prog = st.selectbox("Programa Académico (ECISA)", options=progs_disp)
with f3:
    sel_centro = st.selectbox("Centro Regional (ZCBC)", options=centros_disp)
with f4:
    sel_tutor = st.selectbox("Tutor Encargado", options=tutores_disp)
with f5:
    sel_lider_prog = st.selectbox("Líder Nacional de Programa", options=lideres_disp)

c_busq, c_pdf = st.columns([3.2, 1.3])
with c_busq:
    busq_txt = st.text_input(
        "Buscar por ID de Homologación (`Idhomologacion`), Documento o Nombre del Estudiante:",
        placeholder="Ej: 39232, 1082746042, DELGADO TELLO...",
    )

df_filtrado = df_base.copy()
if sel_tipo != "Todos":
    df_filtrado = df_filtrado[df_filtrado["tipo_homologacion"] == sel_tipo]
if sel_prog != "Todos":
    df_filtrado = df_filtrado[df_filtrado["programa_destino"] == sel_prog]
if sel_centro != "Todos":
    df_filtrado = df_filtrado[df_filtrado["centro"] == sel_centro]
if sel_tutor != "Todos":
    df_filtrado = df_filtrado[df_filtrado["tutor_nombre"] == sel_tutor]
if sel_lider_prog != "Todos":
    df_filtrado = df_filtrado[df_filtrado["lider_programa_nombre"] == sel_lider_prog]

if busq_txt and busq_txt.strip():
    term = busq_txt.strip().lower()
    mask = (
        df_filtrado["id_homologacion"].astype(str).str.lower().str.contains(term, na=False)
        | df_filtrado["documento_estudiante"].astype(str).str.lower().str.contains(term, na=False)
        | df_filtrado["nombre_estudiante"].astype(str).str.lower().str.contains(term, na=False)
        | df_filtrado["tutor_nombre"].astype(str).str.lower().str.contains(term, na=False)
    )
    df_filtrado = df_filtrado[mask]

stats = calcular_estadisticas_homologaciones(df_filtrado)

with c_pdf:
    st.write("")
    pdf_bytes = generar_pdf_informe_homologaciones(
        stats=stats,
        homologaciones_muestra=df_filtrado.to_dict(orient="records"),
    )
    st.download_button(
        label="📥 Descargar Informe PDF",
        data=pdf_bytes,
        file_name="Informe_Tiempos_Homologaciones_2026_ECISA.pdf",
        mime="application/pdf",
        type="primary",
        use_container_width=True,
    )

# =============================================================================
# 2. INDICADORES CLAVE (KPIs DE VOLUMEN Y TIEMPOS DE RESPUESTA)
# =============================================================================
k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric(
    "🗂️ Total Homologaciones",
    f"{stats['total_homologaciones']}",
    delta=f"{stats['estudiantes_unicos']} estudiantes únicos",
    delta_color="off",
)
k2.metric(
    "🏛️ Internas (`int`)",
    f"{stats['total_int']}",
    delta=f"{(stats['total_int'] / max(stats['total_homologaciones'], 1) * 100):.1f}% del total",
    delta_color="off",
)
k3.metric(
    "🎓 Egresados / Conv. (`egr`)",
    f"{stats['total_egr']}",
    delta=f"{(stats['total_egr'] / max(stats['total_homologaciones'], 1) * 100):.1f}% del total",
    delta_color="off",
)
k4.metric(
    "🌐 Externas (`ext`)",
    f"{stats['total_ext']}",
    delta=f"{(stats['total_ext'] / max(stats['total_homologaciones'], 1) * 100):.1f}% del total",
    delta_color="off",
)
k5.metric(
    "⏱️ Tiempo Total Promedio",
    f"{stats['promedio_tiempo_total']} días",
    delta=f"Calendario: {stats['promedio_dias_calendario']} días",
    delta_color="off",
)
k6.metric(
    "⚙️ Actores vs Admon",
    f"{stats['promedio_tiempo_actores']}d / {stats['promedio_tiempo_admon']}d",
    delta=f"Líder: {stats['promedio_lider_hom']}d | Tutor: {stats['promedio_tutor']}d | Prog: {stats['promedio_lider_prog']}d",
    delta_color="off",
)

st.divider()

# =============================================================================
# 3. GRÁFICAS EJECUTIVAS DE VOLUMEN Y TIEMPOS (PLOTLY)
# =============================================================================
g1, g2 = st.columns([1.45, 1])

with g1:
    df_prog_t = stats["df_programas_tiempos"]
    if not df_prog_t.empty:
        fig_prog = px.bar(
            df_prog_t,
            x="Homologaciones",
            y="programa_destino",
            orientation="h",
            text=df_prog_t.apply(
                lambda r: f"{int(r['Homologaciones'])} est. (Prom: {r['Prom_Total']} días)",
                axis=1,
            ),
            title="Homologaciones y Tiempo Total Promedio por Programa Académico (ECISA)",
            color="Prom_Total",
            color_continuous_scale="Blues",
        )
        fig_prog.update_layout(
            yaxis=dict(autorange="reversed", title=""),
            xaxis_title="Cantidad de Homologaciones",
            coloraxis_colorbar=dict(title="Días Prom."),
            height=360,
            margin=dict(t=45, b=20, l=20, r=20),
        )
        st.plotly_chart(fig_prog, use_container_width=True)

with g2:
    if stats["por_tipo"]:
        df_g_tipo = pd.DataFrame(
            list(stats["por_tipo"].items()),
            columns=["Tipo de Homologación", "Cantidad"],
        )
        fig_tipo = px.pie(
            df_g_tipo,
            names="Tipo de Homologación",
            values="Cantidad",
            hole=0.45,
            title="Distribución por Tipo de Homologación (int / egr / ext)",
            color_discrete_sequence=["#004B87", "#0284C7", "#38BDF8"],
        )
        fig_tipo.update_traces(textposition="inside", textinfo="percent+value")
        fig_tipo.update_layout(
            legend=dict(orientation="h", yanchor="bottom", y=-0.18, xanchor="center", x=0.5),
            height=360,
            margin=dict(t=45, b=40, l=20, r=20),
        )
        st.plotly_chart(fig_tipo, use_container_width=True)

g3, g4 = st.columns(2)

with g3:
    df_cen_t = stats["df_centros_tiempos"]
    if not df_cen_t.empty:
        fig_cen = px.bar(
            df_cen_t,
            x="Homologaciones",
            y="centro",
            orientation="h",
            text=df_cen_t.apply(
                lambda r: f"{int(r['Homologaciones'])} ({r['Prom_Total']} d)",
                axis=1,
            ),
            title="Homologaciones y Tiempo Promedio por Centro Regional (ZCBC)",
            color="centro",
            color_discrete_sequence=px.colors.qualitative.Safe,
        )
        fig_cen.update_layout(
            showlegend=False,
            yaxis=dict(autorange="reversed", title=""),
            xaxis_title="Homologaciones",
            height=350,
            margin=dict(t=45, b=20, l=20, r=20),
        )
        st.plotly_chart(fig_cen, use_container_width=True)

with g4:
    # Desglose comparativo de tiempos promedio por etapa y actor
    etapas_df = pd.DataFrame(
        [
            {"Etapa / Actor": "1. Líder Nac. Homologación", "Días Promedio": stats["promedio_lider_hom"]},
            {"Etapa / Actor": "2. Tutor Encargado", "Días Promedio": stats["promedio_tutor"]},
            {"Etapa / Actor": "3. Líder Nac. Programa", "Días Promedio": stats["promedio_lider_prog"]},
            {"Etapa / Actor": "4. Tiempo Administración", "Días Promedio": stats["promedio_tiempo_admon"]},
            {"Etapa / Actor": "5. Tiempo Envío", "Días Promedio": stats["promedio_tiempo_envio"]},
            {"Etapa / Actor": "TOTAL TRÁMITE (Tiempo Total)", "Días Promedio": stats["promedio_tiempo_total"]},
        ]
    )
    fig_tiempos = px.bar(
        etapas_df,
        x="Días Promedio",
        y="Etapa / Actor",
        orientation="h",
        text="Días Promedio",
        title="Desglose del Tiempo Promedio de Respuesta por Actor y Etapa (Días)",
        color="Etapa / Actor",
        color_discrete_sequence=["#38BDF8", "#0284C7", "#0369A1", "#64748B", "#94A3B8", "#004B87"],
    )
    fig_tiempos.update_layout(
        showlegend=False,
        yaxis=dict(autorange="reversed", title=""),
        height=350,
        margin=dict(t=45, b=20, l=20, r=20),
    )
    st.plotly_chart(fig_tiempos, use_container_width=True)

g5, g6 = st.columns(2)

with g5:
    df_tut = stats["df_tutores"]
    if not df_tut.empty:
        fig_tut = go.Figure()
        fig_tut.add_trace(
            go.Bar(
                name="Homologaciones Asignadas",
                y=df_tut["tutor_nombre"],
                x=df_tut["Homologaciones"],
                orientation="h",
                text=df_tut.apply(
                    lambda r: f"{int(r['Homologaciones'])} est. (Tutor: {r['Promedio_Dias_Tutor']}d)",
                    axis=1,
                ),
                textposition="auto",
                marker_color="#004B87",
            )
        )
        fig_tut.update_layout(
            title="Carga y Tiempo Promedio de Respuesta por Tutor Encargado",
            yaxis=dict(autorange="reversed", title=""),
            xaxis_title="Estudios de Homologación",
            height=380,
            margin=dict(t=45, b=20, l=20, r=20),
        )
        st.plotly_chart(fig_tut, use_container_width=True)

with g6:
    if stats["por_mes"]:
        df_mes = pd.DataFrame(
            list(stats["por_mes"].items()),
            columns=["Mes de Aprobación (2026)", "Homologaciones"],
        )
        fig_mes = px.bar(
            df_mes,
            x="Mes de Aprobación (2026)",
            y="Homologaciones",
            text="Homologaciones",
            title="Evolución Mensual de Homologaciones Aprobadas (Ene - Nov 2026)",
            color="Homologaciones",
            color_continuous_scale="Blues",
        )
        fig_mes.update_layout(
            coloraxis_showscale=False,
            height=380,
            margin=dict(t=45, b=20, l=20, r=20),
        )
        st.plotly_chart(fig_mes, use_container_width=True)

st.divider()

# =============================================================================
# 4. TABLAS DETALLADAS DE TRAZABILIDAD Y TIEMPOS POR ACTOR
# =============================================================================
tab_det, tab_actores, tab_prog_cen = st.tabs(
    [
        f"📋 1. Sábana Completa de Seguimiento y Tiempos ({len(df_filtrado)} estudios)",
        "👨‍🏫 2. Desempeño por Tutor Encargado y Líder Nacional de Programa",
        "🏛️ 3. Consolidado por Programa Académico y Centro Regional",
    ]
)

with tab_det:
    if not usuario_es_admin:
        st.info(
            "🔒 **Nota de Privacidad:** La consulta detallada con números de identificación de estudiantes "
            "está reservada para el usuario **Administrador (`admin`)**."
        )
    else:
        if df_filtrado.empty:
            st.warning("No se encontraron estudios de homologación con los filtros seleccionados.")
        else:
            df_show = df_filtrado[
                [
                    "id_homologacion",
                    "documento_estudiante",
                    "nombre_estudiante",
                    "programa_destino",
                    "centro",
                    "tipo_homologacion",
                    "lider_homologacion_nombre",
                    "tiempo_lider_hom",
                    "tutor_nombre",
                    "tiempo_tutor",
                    "lider_programa_nombre",
                    "tiempo_lider_prog",
                    "fecha_aprobacion",
                    "fecha_envio",
                    "tiempo_envio",
                    "tiempo_actores",
                    "tiempo_admon",
                    "tiempo_total",
                    "dias_calendario",
                ]
            ].rename(
                columns={
                    "id_homologacion": "ID Homolog.",
                    "documento_estudiante": "Identificación",
                    "nombre_estudiante": "Estudiante",
                    "programa_destino": "Programa Académico",
                    "centro": "Centro Regional",
                    "tipo_homologacion": "Tipo",
                    "lider_homologacion_nombre": "Líder Nac. Homologación",
                    "tiempo_lider_hom": "Días Líder Hom.",
                    "tutor_nombre": "Tutor Encargado",
                    "tiempo_tutor": "Días Tutor",
                    "lider_programa_nombre": "Líder Nac. Programa",
                    "tiempo_lider_prog": "Días Líder Prog.",
                    "fecha_aprobacion": "Fecha Aprobación",
                    "fecha_envio": "Fecha Envío",
                    "tiempo_envio": "T. Envío",
                    "tiempo_actores": "T. Actores",
                    "tiempo_admon": "T. Admon",
                    "tiempo_total": "Tiempo Total (d)",
                    "dias_calendario": "Días Cal.",
                }
            )
            st.dataframe(df_show, use_container_width=True, hide_index=True, height=450)

with tab_actores:
    c_t1, c_t2 = st.columns(2)
    with c_t1:
        st.markdown("##### 👨‍🏫 Tiempos por Tutor Encargado")
        df_tut_show = stats["df_tutores"].rename(
            columns={
                "tutor_nombre": "Tutor Encargado",
                "Homologaciones": "Estudios Asignados",
                "Promedio_Dias_Tutor": "Prom. Días Tutor",
                "Max_Dias_Tutor": "Máx. Días Tutor",
                "Promedio_Tiempo_Total": "Prom. Tiempo Total (d)",
            }
        )
        st.dataframe(df_tut_show, use_container_width=True, hide_index=True)

    with c_t2:
        st.markdown("##### 🎓 Tiempos por Líder Nacional de Programa")
        df_lp_show = stats["df_lideres_programa"].rename(
            columns={
                "lider_programa_nombre": "Líder Nacional de Programa",
                "Homologaciones": "Estudios Avalados",
                "Promedio_Dias_Lider_Prog": "Prom. Días Líder Prog.",
                "Max_Dias_Lider_Prog": "Máx. Días Líder Prog.",
                "Promedio_Tiempo_Total": "Prom. Tiempo Total (d)",
            }
        )
        st.dataframe(df_lp_show, use_container_width=True, hide_index=True)

with tab_prog_cen:
    c_p1, c_p2 = st.columns(2)
    with c_p1:
        st.markdown("##### 📚 Resumen por Programa Académico")
        df_pr_show = stats["df_programas_tiempos"].rename(
            columns={
                "programa_destino": "Programa Académico",
                "Homologaciones": "Total",
                "Internas_int": "int",
                "Egresados_egr": "egr",
                "Externas_ext": "ext",
                "Prom_Actores": "Prom. Actores (d)",
                "Prom_Admon": "Prom. Admon (d)",
                "Prom_Total": "Prom. Total (d)",
                "Prom_Dias_Cal": "Prom. Días Cal.",
            }
        )
        st.dataframe(df_pr_show, use_container_width=True, hide_index=True)

    with c_p2:
        st.markdown("##### 🏛️ Resumen por Centro Regional (ZCBC)")
        df_ce_show = stats["df_centros_tiempos"].rename(
            columns={
                "centro": "Centro Regional",
                "Homologaciones": "Total",
                "Internas_int": "int",
                "Egresados_egr": "egr",
                "Externas_ext": "ext",
                "Prom_Actores": "Prom. Actores (d)",
                "Prom_Admon": "Prom. Admon (d)",
                "Prom_Total": "Prom. Total (d)",
            }
        )
        st.dataframe(df_ce_show, use_container_width=True, hide_index=True)

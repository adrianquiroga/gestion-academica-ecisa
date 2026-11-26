"""
Módulo 4 (Fase 3 Activa): Gestión, Analítica y Trazabilidad de Homologaciones 2026 (ECISA - UNAD).
Procesa y visualiza los acuerdos de Consejo de Escuela de `Homologaciones 2026.xlsx` (corte noviembre de 2026):
- KPIs de Homologaciones Estudiantiles (630 acuerdos), Estudiantes Únicos, Internas (UNAD), Convenio SENA y Otras IES, más 21 Acuerdos Normativos de Consejo de Escuela.
- Gráficas interactivas por Programa a Cursar (Destino), Tipo/Institución de Origen (UNAD vs SENA vs Otras IES), Sesión de Consejo de Escuela, Centro Regional, Nivel de Formación de Origen y Programa de Origen.
- Filtros combinados por Tipo (Interna/Externa), Programa a Cursar, Consejo de Escuela, Institución de Origen, Centro, Nivel y buscador por Estudiante / Documento / N° de Acuerdo.
- Descarga de Informe Ejecutivo de Homologaciones 2026 en PDF (`fpdf2` en memoria).
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
import src.services.homologaciones_svc as _hom_svc
import src.services.pdf_svc as _pdf_svc
import src.utils.auth as _auth_mod

importlib.reload(_hom_svc)
importlib.reload(_pdf_svc)
importlib.reload(_auth_mod)

RUTA_EXCEL_HOMOLOGACIONES_DEFECTO = _hom_svc.RUTA_EXCEL_HOMOLOGACIONES_DEFECTO
calcular_estadisticas_homologaciones = _hom_svc.calcular_estadisticas_homologaciones
limpiar_cache_homologaciones = _hom_svc.limpiar_cache_homologaciones
obtener_datos_homologaciones_2026 = _hom_svc.obtener_datos_homologaciones_2026
sincronizar_homologaciones_a_supabase = _hom_svc.sincronizar_homologaciones_a_supabase
generar_pdf_informe_homologaciones = _pdf_svc.generar_pdf_informe_homologaciones
es_admin = _auth_mod.es_admin
renderizar_barra_sesion_y_login = _auth_mod.renderizar_barra_sesion_y_login

st.set_page_config(
    page_title="Homologaciones 2026 - ECISA UNAD",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded",
)

if not renderizar_barra_sesion_y_login("4_homologaciones"):
    st.stop()

c_title, c_refresh = st.columns([4, 1])
with c_title:
    st.title("🔄 Gestión y Analítica de Homologaciones 2026 (ECISA)")
    st.markdown(
        "Trazabilidad de **Acuerdos de Homologación por Consejo de Escuela 2026** (corte al 23 de noviembre de 2026): "
        "análisis por **Programa a Cursar, Tipo (Interna UNAD / Convenio SENA / Otras IES), Centro Regional, Programa de Origen y Acuerdos Normativos**."
    )
with c_refresh:
    st.write("")
    if st.button("🔄 Actualizar Datos", use_container_width=True):
        limpiar_cache_homologaciones()
        st.rerun()

usuario_es_admin = es_admin()

ok_db, msg_db = verificar_configuracion_supabase()
if not ok_db:
    st.warning(f"⚠️ Aviso de conexión a Supabase: {msg_db}. Se utilizará la base local sincronizada.")

df_hom_base, df_ac_base = obtener_datos_homologaciones_2026()

# =============================================================================
# 0. PANEL DESPLEGABLE DE CARGA / SINCRONIZACIÓN DE `Homologaciones 2026.xlsx` (SOLO ADMIN)
# =============================================================================
if usuario_es_admin:
    with st.expander(
        "📥 Cargar / Sincronizar `Homologaciones 2026.xlsx` con Supabase (Solo Admin)",
        expanded=df_hom_base.empty,
    ):
        st.markdown(
            """
            Este importador procesa automáticamente la matriz oficial **`Homologaciones 2026.xlsx`**:
            * **Limpieza inteligente:** Une automáticamente las filas donde el nombre del programa de origen se partió en dos líneas y corrige las filas donde `Nombre` e `Identidad` venían invertidos.
            * **Clasificación automática:** Identifica homologaciones **Internas (UNAD)**, **Externas por Convenio SENA** y **Externas de otras Universidades/IES**, además de separar los **Acuerdos Normativos/Generales de Consejo de Escuela**.
            """
        )
        col_loc, col_up = st.columns(2)

        with col_loc:
            if RUTA_EXCEL_HOMOLOGACIONES_DEFECTO.exists():
                st.info(f"📂 Archivo detectado en el proyecto: **`{RUTA_EXCEL_HOMOLOGACIONES_DEFECTO}`**")
                if st.button(
                    "🚀 Sincronizar `Homologaciones 2026.xlsx` con Supabase",
                    type="primary",
                    use_container_width=True,
                ):
                    with st.spinner("Procesando acuerdos de Consejo de Escuela y sincronizando con Supabase..."):
                        ok_s, n_h, n_a, msg_s = sincronizar_homologaciones_a_supabase(None, reemplazar_existentes=True)
                    if ok_s:
                        st.success(f"✅ {msg_s}")
                        st.rerun()
                    else:
                        st.warning(f"⚠️ {msg_s}")
                        st.rerun()
            else:
                st.info("Puedes subir una nueva versión de `Homologaciones 2026.xlsx` en el panel de la derecha 👉.")

        with col_up:
            excel_subido = st.file_uploader(
                "O sube una versión actualizada de `Homologaciones 2026.xlsx`:",
                type=["xlsx", "xls"],
                key="uploader_homologaciones_2026",
            )
            if excel_subido is not None:
                if st.button("📥 Procesar y Sincronizar Archivo Subido", type="primary", use_container_width=True):
                    with st.spinner("Procesando archivo subido y actualizando Supabase..."):
                        ok_u, n_hu, n_au, msg_u = sincronizar_homologaciones_a_supabase(
                            excel_subido,
                            reemplazar_existentes=True,
                        )
                    if ok_u:
                        st.success(f"✅ {msg_u}")
                        st.rerun()
                    else:
                        st.warning(f"⚠️ {msg_u}")
                        st.rerun()

if df_hom_base.empty:
    st.info(
        "💡 Aún no hay registros de homologaciones cargados. "
        "Abre el panel superior **'📥 Cargar / Sincronizar Homologaciones 2026.xlsx'** para importar la base."
    )
    st.stop()

# =============================================================================
# 1. FILTROS INTERACTIVOS Y DESCARGA DE INFORME PDF
# =============================================================================
st.markdown("### 🔎 Filtros de Consulta y Reporte Ejecutivo PDF")

f1, f2, f3, f4, f5 = st.columns([1.05, 1.45, 1.35, 1.35, 1.1])

tipos_disp = ["Todos", "Interna (UNAD)", "Externa (SENA + Otras IES)", "Solo Convenio SENA", "Solo Otras Universidades / IES"]
progs_disp = ["Todos"] + sorted(df_hom_base["programa_destino"].dropna().unique().tolist())
consejos_disp = ["Todos"] + sorted(df_hom_base["consejo_escuela"].dropna().unique().tolist())
centros_disp = ["Todos"] + sorted(df_hom_base["centro"].dropna().unique().tolist())
niveles_disp = ["Todos"] + sorted(df_hom_base["nivel_origen"].dropna().unique().tolist())

with f1:
    sel_tipo = st.selectbox("Tipo / Origen", options=tipos_disp)
with f2:
    sel_prog_dest = st.selectbox("Programa a Cursar (ECISA)", options=progs_disp)
with f3:
    sel_consejo = st.selectbox("Consejo de Escuela", options=consejos_disp)
with f4:
    sel_centro = st.selectbox("Centro Regional", options=centros_disp)
with f5:
    sel_nivel = st.selectbox("Nivel de Origen", options=niveles_disp)

c_busq, c_pdf = st.columns([3.2, 1.3])
with c_busq:
    busq_txt = st.text_input(
        "Buscar por Nombre del Estudiante, Documento (Identidad), Número de Acuerdo, Programa o Universidad de Origen:",
        placeholder="Ej: 1083910506, Rosero, ACUERDO NÚMERO 639, SENA, Regencia...",
    )

# Aplicar filtros al DataFrame
df_filtrado = df_hom_base.copy()

if sel_tipo == "Interna (UNAD)":
    df_filtrado = df_filtrado[df_filtrado["tipo"] == "Interna"]
elif sel_tipo == "Externa (SENA + Otras IES)":
    df_filtrado = df_filtrado[df_filtrado["tipo"] == "Externa"]
elif sel_tipo == "Solo Convenio SENA":
    df_filtrado = df_filtrado[df_filtrado["institucion_origen"].str.contains("SENA", case=False, na=False)]
elif sel_tipo == "Solo Otras Universidades / IES":
    df_filtrado = df_filtrado[
        (df_filtrado["tipo"] == "Externa")
        & (~df_filtrado["institucion_origen"].str.contains("SENA", case=False, na=False))
    ]

if sel_prog_dest != "Todos":
    df_filtrado = df_filtrado[df_filtrado["programa_destino"] == sel_prog_dest]

if sel_consejo != "Todos":
    df_filtrado = df_filtrado[df_filtrado["consejo_escuela"] == sel_consejo]

if sel_centro != "Todos":
    df_filtrado = df_filtrado[df_filtrado["centro"] == sel_centro]

if sel_nivel != "Todos":
    df_filtrado = df_filtrado[df_filtrado["nivel_origen"] == sel_nivel]

if busq_txt and busq_txt.strip():
    term = busq_txt.strip().lower()
    mask = (
        df_filtrado["nombre_estudiante"].astype(str).str.lower().str.contains(term, na=False)
        | df_filtrado["documento_estudiante"].astype(str).str.lower().str.contains(term, na=False)
        | df_filtrado["acuerdo"].astype(str).str.lower().str.contains(term, na=False)
        | df_filtrado["programa_origen"].astype(str).str.lower().str.contains(term, na=False)
        | df_filtrado["institucion_origen"].astype(str).str.lower().str.contains(term, na=False)
        | df_filtrado["snies_origen"].astype(str).str.lower().str.contains(term, na=False)
    )
    df_filtrado = df_filtrado[mask]

stats = calcular_estadisticas_homologaciones(df_filtrado, df_ac_base)

with c_pdf:
    st.write("")
    pdf_bytes = generar_pdf_informe_homologaciones(
        stats=stats,
        homologaciones_muestra=df_filtrado.to_dict(orient="records"),
    )
    st.download_button(
        label="📥 Descargar Informe PDF",
        data=pdf_bytes,
        file_name="Informe_Homologaciones_2026_ECISA.pdf",
        mime="application/pdf",
        type="primary",
        use_container_width=True,
    )

# =============================================================================
# 2. INDICADORES CLAVE (KPIs)
# =============================================================================
total_h = stats["total_homologaciones"]
pct_int = (stats["total_internas"] / total_h * 100.0) if total_h else 0.0
pct_sena = (stats["total_sena"] / total_h * 100.0) if total_h else 0.0
pct_otras = (stats["total_otras_ies"] / total_h * 100.0) if total_h else 0.0

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("📜 Acuerdos Homologación", f"{total_h:,}".replace(",", "."), delta=f"{stats['total_consejos']} Consejos de Escuela", delta_color="off")
k2.metric("👨‍🎓 Estudiantes Únicos", f"{stats['estudiantes_unicos']:,}".replace(",", "."))
k3.metric("🏛️ Internas (UNAD)", f"{stats['total_internas']:,}".replace(",", "."), delta=f"{pct_int:.1f}% del total", delta_color="off")
k4.metric("🤝 Convenio SENA", f"{stats['total_sena']:,}".replace(",", "."), delta=f"{pct_sena:.1f}% del total", delta_color="off")
k5.metric("🎓 Otras Universidades", f"{stats['total_otras_ies']:,}".replace(",", "."), delta=f"{pct_otras:.1f}% del total", delta_color="off")
k6.metric("⚖️ Acuerdos Normativos", f"{stats['total_acuerdos_generales']}", delta="Aplazamientos / Rutas", delta_color="off")

st.divider()

# =============================================================================
# 3. GRÁFICAS EJECUTIVAS INTERACTIVAS (PLOTLY)
# =============================================================================
g1, g2 = st.columns([1.45, 1])

with g1:
    if stats["por_programa_destino"]:
        df_g_prog = pd.DataFrame(
            sorted(stats["por_programa_destino"].items(), key=lambda x: x[1], reverse=True),
            columns=["Programa a Cursar (ECISA)", "Homologaciones"],
        )
        fig_prog = px.bar(
            df_g_prog,
            x="Homologaciones",
            y="Programa a Cursar (ECISA)",
            orientation="h",
            text="Homologaciones",
            title="Homologaciones Aprobadas por Programa a Cursar en ECISA (Destino)",
            color="Programa a Cursar (ECISA)",
            color_discrete_sequence=px.colors.sequential.Blues_r,
        )
        fig_prog.update_layout(
            showlegend=False,
             yaxis=dict(autorange="reversed"),
            height=360,
            margin=dict(t=45, b=20, l=20, r=20),
        )
        st.plotly_chart(fig_prog, use_container_width=True)

with g2:
    if stats["por_tipo"]:
        df_g_tipo = pd.DataFrame(
            [(k, v) for k, v in stats["por_tipo"].items() if v > 0],
            columns=["Origen de Homologación", "Cantidad"],
        )
        fig_tipo = px.pie(
            df_g_tipo,
            names="Origen de Homologación",
            values="Cantidad",
            hole=0.45,
            title="Distribución por Tipo e Institución de Origen",
            color_discrete_sequence=["#004B87", "#0284C7", "#38BDF8"],
        )
        fig_tipo.update_traces(textposition="inside", textinfo="percent+value")
        fig_tipo.update_layout(
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
            height=360,
            margin=dict(t=45, b=40, l=20, r=20),
        )
        st.plotly_chart(fig_tipo, use_container_width=True)

g3, g4 = st.columns(2)

with g3:
    if stats["por_consejo"]:
        df_g_cons = pd.DataFrame(
            sorted(stats["por_consejo"].items(), key=lambda x: x[0]),
            columns=["Consejo de Escuela", "Acuerdos"],
        )
        fig_cons = px.bar(
            df_g_cons,
            x="Acuerdos",
            y="Consejo de Escuela",
            orientation="h",
            text="Acuerdos",
            title="Evolución de Acuerdos por Sesión de Consejo de Escuela (2026)",
            color="Acuerdos",
            color_continuous_scale="Blues",
        )
        fig_cons.update_layout(
            coloraxis_showscale=False,
            height=360,
            margin=dict(t=45, b=20, l=20, r=20),
        )
        st.plotly_chart(fig_cons, use_container_width=True)

with g4:
    if stats["por_centro"]:
        df_g_cen = pd.DataFrame(
            sorted(stats["por_centro"].items(), key=lambda x: x[1], reverse=True)[:15],
            columns=["Centro Regional", "Homologaciones"],
        )
        fig_cen = px.bar(
            df_g_cen,
            x="Homologaciones",
            y="Centro Regional",
            orientation="h",
            text="Homologaciones",
            title="Top 15 Centros Regionales con Mayor Número de Homologaciones",
            color="Centro Regional",
            color_discrete_sequence=px.colors.qualitative.Safe,
        )
        fig_cen.update_layout(
            showlegend=False,
            yaxis=dict(autorange="reversed"),
            height=360,
            margin=dict(t=45, b=20, l=20, r=20),
        )
        st.plotly_chart(fig_cen, use_container_width=True)

g5, g6 = st.columns([1.35, 1])

with g5:
    if stats["por_programa_origen"]:
        df_g_porig = pd.DataFrame(
            sorted(stats["por_programa_origen"].items(), key=lambda x: x[1], reverse=True)[:12],
            columns=["Programa de Origen", "Homologaciones"],
        )
        fig_porig = px.bar(
            df_g_porig,
            x="Homologaciones",
            y="Programa de Origen",
            orientation="h",
            text="Homologaciones",
            title="Top 12 Programas Académicos de Origen más Homologados",
            color="Homologaciones",
            color_continuous_scale="Tealgrn",
        )
        fig_porig.update_layout(
            coloraxis_showscale=False,
            yaxis=dict(autorange="reversed"),
            height=350,
            margin=dict(t=45, b=20, l=20, r=20),
        )
        st.plotly_chart(fig_porig, use_container_width=True)

with g6:
    if stats["por_nivel_origen"]:
        df_g_niv = pd.DataFrame(
            sorted(stats["por_nivel_origen"].items(), key=lambda x: x[1], reverse=True),
            columns=["Nivel de Formación de Origen", "Homologaciones"],
        )
        fig_niv = px.bar(
            df_g_niv,
            x="Nivel de Formación de Origen",
            y="Homologaciones",
            text="Homologaciones",
            title="Distribución por Nivel de Formación de Origen",
            color="Nivel de Formación de Origen",
            color_discrete_sequence=px.colors.qualitative.Pastel,
        )
        fig_niv.update_layout(
            showlegend=False,
            height=350,
            margin=dict(t=45, b=20, l=20, r=20),
        )
        st.plotly_chart(fig_niv, use_container_width=True)

st.divider()

# =============================================================================
# 4. PESTAÑAS DE DETALLE: ACUERDOS ESTUDIANTILES, RESUMEN POR IES Y ACUERDOS NORMATIVOS
# =============================================================================
tab_est, tab_univ, tab_norm = st.tabs(
    [
        f"👨‍🎓 1. Directorio de Acuerdos de Homologación Estudiantil ({len(df_filtrado)})",
        "🏛️ 2. Consolidado por Universidad / Institución de Origen",
        f"⚖️ 3. Acuerdos Normativos y Generales de Consejo de Escuela ({len(df_ac_base)})",
    ]
)

with tab_est:
    if not usuario_es_admin:
        st.info(
            "🔒 **Nota de Privacidad:** La tabla detallada con números de documento de identidad de estudiantes "
            "está reservada para el rol **Administrador (`admin`)**."
        )
    else:
        if df_filtrado.empty:
            st.warning("No se encontraron acuerdos de homologación con los filtros seleccionados.")
        else:
            df_mostrar = df_filtrado[
                [
                    "numero_acuerdo",
                    "consejo_escuela",
                    "acuerdo",
                    "fecha_resolucion",
                    "documento_estudiante",
                    "nombre_estudiante",
                    "programa_destino",
                    "centro",
                    "programa_origen",
                    "snies_origen",
                    "nivel_origen",
                    "institucion_origen",
                    "tipo",
                ]
            ].rename(
                columns={
                    "numero_acuerdo": "N° Ac.",
                    "consejo_escuela": "Consejo de Escuela",
                    "acuerdo": "Acuerdo de Homologación",
                    "fecha_resolucion": "Fecha Acuerdo",
                    "documento_estudiante": "Identidad (Documento)",
                    "nombre_estudiante": "Nombre del Estudiante",
                    "programa_destino": "Programa a Cursar (ECISA)",
                    "centro": "Centro Regional",
                    "programa_origen": "Programa de Origen",
                    "snies_origen": "SNIES Origen",
                    "nivel_origen": "Nivel Origen",
                    "institucion_origen": "Universidad / Institución de Origen",
                    "tipo": "Tipo",
                }
            )
            st.dataframe(df_mostrar, use_container_width=True, hide_index=True, height=460)

with tab_univ:
    if df_filtrado.empty:
        st.info("Sin registros para agrupar.")
    else:
        df_u_grp = (
            df_filtrado.groupby(["institucion_origen", "tipo", "nivel_origen"], as_index=False)
            .agg(
                Total_Homologaciones=("documento_estudiante", "count"),
                Estudiantes_Unicos=("documento_estudiante", "nunique"),
                Programas_Origen=("programa_origen", "nunique"),
            )
            .sort_values(by="Total_Homologaciones", ascending=False)
            .reset_index(drop=True)
        )
        df_u_grp = df_u_grp.rename(
            columns={
                "institucion_origen": "Universidad / Institución de Origen",
                "tipo": "Tipo de Homologación",
                "nivel_origen": "Nivel de Formación",
                "Total_Homologaciones": "Total Acuerdos",
                "Estudiantes_Unicos": "Estudiantes Únicos",
                "Programas_Origen": "Programas de Origen Distintos",
            }
        )
        st.dataframe(df_u_grp, use_container_width=True, hide_index=True)

with tab_norm:
    st.markdown(
        "Acuerdos generales expedidos por el **Consejo de Escuela de Ciencias de la Salud (ECISA)** en 2026 "
        "(solicitudes de aplazamiento de periodo, condiciones de homologación interna hacia posgrados, diplomados, prácticas y suficiencia por competencias):"
    )
    if df_ac_base.empty:
        st.info("No hay acuerdos generales registrados.")
    else:
        df_ac_show = df_ac_base[
            ["numero_acuerdo", "consejo_escuela", "acuerdo", "fecha_acuerdo", "asunto"]
        ].rename(
            columns={
                "numero_acuerdo": "N° Acuerdo",
                "consejo_escuela": "Consejo de Escuela",
                "acuerdo": "Acuerdo",
                "fecha_acuerdo": "Fecha",
                "asunto": "Descripción / Objeto del Acuerdo",
            }
        )
        st.dataframe(df_ac_show, use_container_width=True, hide_index=True)

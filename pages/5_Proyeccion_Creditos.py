"""
Módulo 5: Comparativo de Créditos Proyectados (`proy_zon`) vs Realizados (`creditos`) en 2026 (`data ECISA.xlsx`).
Enfocado en las columnas:
- `codigo_real_snies`
- `pro_descripcion`
- `nombrecead`
- `proy_zon`
- `creditos`
Incluye gráficas comparativas por Programa Académico y por Nombre del CEAD, porcentaje de cumplimiento,
matriz cruzada y descarga de Informe Ejecutivo en PDF.
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
from src.services.creditos_svc import (
    RUTA_DATA_ECISA_DEFECTO,
    calcular_resumenes_creditos,
    limpiar_cache_creditos,
    obtener_datos_creditos_ecisa,
    sincronizar_data_ecisa_a_supabase,
)
from src.services.pdf_svc import generar_pdf_informe_creditos_ecisa
import src.utils.auth as _auth_mod

importlib.reload(_auth_mod)
renderizar_barra_sesion_y_login = _auth_mod.renderizar_barra_sesion_y_login

st.set_page_config(
    page_title="Créditos Proyectados vs Realizados 2026 - ECISA",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

if not renderizar_barra_sesion_y_login("5_creditos"):
    st.stop()

c_title, c_refresh = st.columns([4, 1])
with c_title:
    st.title("🎯 Comparativo de Créditos Proyectados vs Realizados 2026 (ECISA)")
    st.markdown(
        "Análisis de cumplimiento de metas (`proy_zon` vs `creditos`) desglosado por "
        "**Programa Académico (`pro_descripcion` / `codigo_real_snies`)** y **Centro Regional (`nombrecead`)**."
    )
with c_refresh:
    st.write("")
    if st.button("🔄 Actualizar Datos", use_container_width=True):
        limpiar_cache_creditos()
        st.rerun()

ok_db, msg_db = verificar_configuracion_supabase()
if not ok_db:
    st.error(f"⚠️ Base de datos no configurada: {msg_db}")
    st.stop()

df_base = obtener_datos_creditos_ecisa()

# =============================================================================
# 0. PANEL DE CARGA / SINCRONIZACIÓN DE `data ECISA.xlsx`
# =============================================================================
with st.expander(
    "📥 Cargar / Sincronizar Archivo `data ECISA.xlsx` (Campos: `codigo_real_snies`, `pro_descripcion`, `nombrecead`, `proy_zon`, `creditos`)",
    expanded=df_base.empty,
):
    col_loc, col_up = st.columns(2)

    with col_loc:
        if RUTA_DATA_ECISA_DEFECTO.exists():
            st.info(f"📂 Archivo detectado en la carpeta del proyecto: **`{RUTA_DATA_ECISA_DEFECTO}`**")
            if st.button(
                "🚀 Sincronizar `data ECISA.xlsx` local con Supabase",
                type="primary",
                use_container_width=True,
            ):
                with st.spinner("Sincronizando proyección y ejecución de créditos 2026 en Supabase..."):
                    ok_s, n_s, msg_s = sincronizar_data_ecisa_a_supabase(None)
                if ok_s:
                    st.success(f"✅ {msg_s}")
                    st.rerun()
                else:
                    st.error(f"❌ {msg_s}")
        else:
            st.warning(
                "No se detectó `data ECISA.xlsx` en la carpeta `C:\\PROYECTOS\\U APLICATIVO`. "
                "Puedes pegarlo en esa carpeta o seleccionarlo directamente en el cuadro de la derecha 👉."
            )

    with col_up:
        excel_subido = st.file_uploader(
            "Selecciona o arrastra aquí tu archivo `data ECISA.xlsx`:",
            type=["xlsx", "xls"],
            key="uploader_data_ecisa",
        )
        if excel_subido is not None:
            if st.button("📥 Procesar y Guardar `data ECISA.xlsx` en Supabase", type="primary", use_container_width=True):
                with st.spinner("Procesando archivo y guardando en Supabase..."):
                    ok_u, n_u, msg_u = sincronizar_data_ecisa_a_supabase(excel_subido)
                if ok_u:
                    st.success(f"✅ {msg_u}")
                    st.rerun()
                else:
                    st.error(f"❌ {msg_u}")

if df_base.empty:
    st.info(
        "👆 Para visualizar las gráficas comparativas, selecciona tu archivo **`data ECISA.xlsx`** en el panel de arriba "
        "o pégalo en `C:\\PROYECTOS\\U APLICATIVO\\data ECISA.xlsx` y haz clic en **Actualizar Datos**."
    )
    st.stop()

# =============================================================================
# 1. FILTROS INTERACTIVOS Y DESCARGA DE PDF
# =============================================================================
st.markdown("### 🔎 Filtros por Centro (`nombrecead`) y Programa (`pro_descripcion`)")

f_cead, f_prog, f_snies, f_pdf = st.columns([1.4, 1.8, 1.0, 1.2])

ceads_unicos = sorted(df_base["nombrecead"].dropna().unique().tolist())
progs_unicos = sorted(df_base["pro_descripcion"].dropna().unique().tolist())
snies_unicos = sorted(df_base["codigo_real_snies"].dropna().unique().tolist())

with f_cead:
    sel_ceads = st.multiselect("Filtrar por CEAD (`nombrecead`)", options=ceads_unicos, placeholder="Todos los CEAD")
with f_prog:
    sel_progs = st.multiselect("Filtrar por Programa (`pro_descripcion`)", options=progs_unicos, placeholder="Todos los programas")
with f_snies:
    sel_snies = st.multiselect("Código SNIES", options=snies_unicos, placeholder="Todos")

resumen = calcular_resumenes_creditos(
    df=df_base,
    filtro_cead=sel_ceads if sel_ceads else None,
    filtro_programa=sel_progs if sel_progs else None,
    filtro_snies=sel_snies if sel_snies else None,
)

with f_pdf:
    st.write("")
    pdf_bytes = generar_pdf_informe_creditos_ecisa(resumen)
    st.download_button(
        label="📥 Descargar Informe PDF",
        data=pdf_bytes,
        file_name="Informe_Creditos_Proyectados_vs_Realizados_2026_ECISA.pdf",
        mime="application/pdf",
        type="primary",
        use_container_width=True,
    )

# =============================================================================
# 2. INDICADORES CLAVE (KPIs)
# =============================================================================
k1, k2, k3, k4 = st.columns(4)
k1.metric("🎯 Créditos Proyectados (`proy_zon`)", f"{resumen['total_proyectado']:,.0f}".replace(",", "."))
k2.metric("✅ Créditos Realizados (`creditos`)", f"{resumen['total_realizado']:,.0f}".replace(",", "."))
k3.metric(
    "📊 Diferencia Neta (Realizados - Proyectados)",
    f"{resumen['diferencia']:+,.0f}".replace(",", "."),
    delta=f"{resumen['diferencia']:+,.0f} créditos".replace(",", "."),
)
k4.metric(
    "🏆 % Cumplimiento de Meta 2026",
    f"{resumen['cumplimiento_pct']:.1f}%",
    delta=f"{resumen['cumplimiento_pct'] - 100.0:+.1f}% vs meta (100%)",
)

st.divider()

# =============================================================================
# 3. GRÁFICAS COMPARATIVAS POR NOMBRE DEL CEAD (`nombrecead`)
# =============================================================================
st.markdown("### 🏛️ 1. Comparativo por Nombre del CEAD (`nombrecead`)")

df_cead = resumen["df_por_cead"]
if not df_cead.empty:
    col_g_cead1, col_g_cead2 = st.columns([1.45, 1])

    with col_g_cead1:
        fig_cead_bar = go.Figure()
        fig_cead_bar.add_trace(
            go.Bar(
                name="Créditos Proyectados (proy_zon)",
                x=df_cead["nombrecead"],
                y=df_cead["proy_zon"],
                text=df_cead["proy_zon"].apply(lambda v: f"{v:,.0f}".replace(",", ".")),
                textposition="auto",
                marker_color="#94A3B8",
            )
        )
        fig_cead_bar.add_trace(
            go.Bar(
                name="Créditos Realizados (creditos)",
                x=df_cead["nombrecead"],
                y=df_cead["creditos"],
                text=df_cead["creditos"].apply(lambda v: f"{v:,.0f}".replace(",", ".")),
                textposition="auto",
                marker_color="#004B87",
            )
        )
        fig_cead_bar.update_layout(
            barmode="group",
            title="Créditos Proyectados vs Realizados 2026 por CEAD (nombrecead)",
            xaxis_title="Centro (nombrecead)",
            yaxis_title="Cantidad de Créditos",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            height=400,
            margin=dict(t=60, b=40, l=20, r=20),
        )
        st.plotly_chart(fig_cead_bar, use_container_width=True)

    with col_g_cead2:
        df_cead_sorted_cump = df_cead.sort_values(by="cumplimiento_pct", ascending=True)
        fig_cead_cump = px.bar(
            df_cead_sorted_cump,
            x="cumplimiento_pct",
            y="nombrecead",
            orientation="h",
            text=df_cead_sorted_cump["cumplimiento_pct"].apply(lambda v: f"{v:.1f}%"),
            title="% de Cumplimiento de Meta por CEAD",
            color="cumplimiento_pct",
            color_continuous_scale="Tealgrn",
        )
        fig_cead_cump.add_vline(x=100, line_dash="dash", line_color="#DC2626", annotation_text="Meta 100%")
        fig_cead_cump.update_layout(
            xaxis_title="% Cumplimiento",
            yaxis_title="",
            coloraxis_showscale=False,
            height=400,
            margin=dict(t=60, b=40, l=20, r=20),
        )
        st.plotly_chart(fig_cead_cump, use_container_width=True)

st.divider()

# =============================================================================
# 4. GRÁFICAS COMPARATIVAS POR PROGRAMA DESCRIPCIÓN (`pro_descripcion`)
# =============================================================================
st.markdown("### 📚 2. Comparativo por Programa Académico (`pro_descripcion` y `codigo_real_snies`)")

df_prog = resumen["df_por_programa"]
if not df_prog.empty:
    df_prog["etiqueta_programa"] = df_prog.apply(
        lambda r: f"{r['pro_descripcion']} (SNIES {r['codigo_real_snies']})"
        if str(r["codigo_real_snies"]) not in {"S/N", "0"}
        else str(r["pro_descripcion"]),
        axis=1,
    )

    col_g_prog1, col_g_prog2 = st.columns([1.45, 1])

    with col_g_prog1:
        df_prog_plot = df_prog.sort_values(by="creditos", ascending=True)
        fig_prog_bar = go.Figure()
        fig_prog_bar.add_trace(
            go.Bar(
                name="Créditos Proyectados (proy_zon)",
                y=df_prog_plot["pro_descripcion"],
                x=df_prog_plot["proy_zon"],
                orientation="h",
                text=df_prog_plot["proy_zon"].apply(lambda v: f"{v:,.0f}".replace(",", ".")),
                textposition="auto",
                marker_color="#94A3B8",
            )
        )
        fig_prog_bar.add_trace(
            go.Bar(
                name="Créditos Realizados (creditos)",
                y=df_prog_plot["pro_descripcion"],
                x=df_prog_plot["creditos"],
                orientation="h",
                text=df_prog_plot["creditos"].apply(lambda v: f"{v:,.0f}".replace(",", ".")),
                textposition="auto",
                marker_color="#0284C7",
            )
        )
        fig_prog_bar.update_layout(
            barmode="group",
            title="Créditos Proyectados vs Realizados 2026 por Programa (pro_descripcion)",
            xaxis_title="Cantidad de Créditos",
            yaxis_title="",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            height=max(420, len(df_prog_plot) * 42),
            margin=dict(t=60, b=40, l=20, r=20),
        )
        st.plotly_chart(fig_prog_bar, use_container_width=True)

    with col_g_prog2:
        df_prog_cump = df_prog.sort_values(by="cumplimiento_pct", ascending=True)
        fig_prog_cump = px.bar(
            df_prog_cump,
            x="cumplimiento_pct",
            y="pro_descripcion",
            orientation="h",
            text=df_prog_cump["cumplimiento_pct"].apply(lambda v: f"{v:.1f}%"),
            title="% de Cumplimiento por Programa",
            color="cumplimiento_pct",
            color_continuous_scale="Blues",
        )
        fig_prog_cump.add_vline(x=100, line_dash="dash", line_color="#DC2626", annotation_text="Meta 100%")
        fig_prog_cump.update_layout(
            xaxis_title="% Cumplimiento",
            yaxis_title="",
            coloraxis_showscale=False,
            height=max(420, len(df_prog_cump) * 42),
            margin=dict(t=60, b=40, l=20, r=20),
        )
        st.plotly_chart(fig_prog_cump, use_container_width=True)

st.divider()

# =============================================================================
# 5. TABLAS DE RESUMEN Y DETALLE CRUZADO (PROGRAMA x CEAD)
# =============================================================================
tab_t_cead, tab_t_prog, tab_t_cruce = st.tabs(
    [
        "🏛️ Tabla Consolidada por CEAD (`nombrecead`)",
        "📚 Tabla Consolidada por Programa (`pro_descripcion`)",
        "🔍 Detalle Cruzado (SNIES + Programa + CEAD)",
    ]
)

with tab_t_cead:
    df_cead_show = df_cead.rename(
        columns={
            "nombrecead": "Nombre del CEAD (nombrecead)",
            "proy_zon": "Créditos Proyectados (proy_zon)",
            "creditos": "Créditos Realizados (creditos)",
            "diferencia": "Diferencia (Realizados - Proy.)",
            "cumplimiento_pct": "% Cumplimiento",
        }
    )
    st.dataframe(df_cead_show, use_container_width=True, hide_index=True)

with tab_t_prog:
    df_prog_show = df_prog[
        ["codigo_real_snies", "pro_descripcion", "proy_zon", "creditos", "diferencia", "cumplimiento_pct"]
    ].rename(
        columns={
            "codigo_real_snies": "Código SNIES (codigo_real_snies)",
            "pro_descripcion": "Programa Descripción (pro_descripcion)",
            "proy_zon": "Créditos Proyectados (proy_zon)",
            "creditos": "Créditos Realizados (creditos)",
            "diferencia": "Diferencia (Realizados - Proy.)",
            "cumplimiento_pct": "% Cumplimiento",
        }
    )
    st.dataframe(df_prog_show, use_container_width=True, hide_index=True)

with tab_t_cruce:
    df_det_show = resumen["df_detalle"].rename(
        columns={
            "codigo_real_snies": "Código SNIES",
            "pro_descripcion": "Programa (pro_descripcion)",
            "nombrecead": "CEAD (nombrecead)",
            "proy_zon": "Proyectado (proy_zon)",
            "creditos": "Realizado (creditos)",
            "diferencia": "Diferencia",
            "cumplimiento_pct": "% Cumplimiento",
        }
    )
    st.dataframe(df_det_show, use_container_width=True, hide_index=True)

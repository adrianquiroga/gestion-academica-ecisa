"""
Servicio para el Módulo de Comparativo de Créditos Proyectados vs Realizados 2026 (`data ECISA.xlsx`).
Enfocado en las columnas:
- `codigo_real_snies` (Código SNIES del programa)
- `pro_descripcion` (Nombre del Programa Académico)
- `nombrecead` (Nombre del Centro / CEAD / CCAV / UDR)
- `proy_zon` (Créditos Proyectados en la Zona/Centro)
- `creditos` (Créditos Realizados / Matriculados en 2026)
"""

from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd
import streamlit as st
from src.db.client import get_supabase_client

RUTA_DATA_ECISA_DEFECTO = (
    Path("datos_locales/data ECISA.xlsx")
    if Path("datos_locales/data ECISA.xlsx").exists()
    else Path("data ECISA.xlsx")
)
BATCH_SIZE_CREDITOS = 750

COLUMNAS_REQUERIDAS = {
    "codigo_real_snies",
    "pro_descripcion",
    "nombrecead",
    "proy_zon",
    "creditos",
}


def limpiar_cache_creditos() -> None:
    """Invalida la caché de consultas de proyección de créditos ECISA."""
    obtener_datos_creditos_ecisa.clear()


def _normalizar_nombre_columna(col: Any) -> str:
    return str(col).strip().lower()


def _limpiar_numero(val: Any) -> float:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    try:
        txt = str(val).strip().replace(",", ".")
        if not txt or txt.lower() in {"nan", "none", "-"}:
            return 0.0
        return round(float(txt), 2)
    except Exception:
        return 0.0


def _limpiar_texto(val: Any, defecto: str = "SIN ESPECIFICAR") -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return defecto
    txt = str(val).strip()
    if not txt or txt.lower() in {"nan", "none"}:
        return defecto
    # Si viene como número flotante tipo '15297.0', quitar '.0'
    if txt.endswith(".0") and txt[:-2].isdigit():
        return txt[:-2]
    return txt


def leer_y_procesar_data_ecisa(fuente_excel: Any = None) -> Tuple[List[Dict[str, Any]], str]:
    """
    Lee `data ECISA.xlsx` (desde la raíz del proyecto o desde un archivo subido),
    localiza automáticamente la fila donde están las columnas
    `codigo_real_snies`, `pro_descripcion`, `nombrecead`, `proy_zon`, `creditos`,
    y retorna los registros normalizados listos para Supabase.
    """
    try:
        if fuente_excel is None:
            if not RUTA_DATA_ECISA_DEFECTO.exists():
                return [], f"No se encontró el archivo `{RUTA_DATA_ECISA_DEFECTO}` en la carpeta del proyecto."
            df = pd.read_excel(RUTA_DATA_ECISA_DEFECTO, sheet_name=0)
        else:
            df = pd.read_excel(fuente_excel, sheet_name=0)
    except Exception as exc:
        return [], f"Error al abrir el archivo Excel: {exc}"

    # Mapear nombres de columnas insensibles a mayúsculas/espacios
    mapa_cols = {_normalizar_nombre_columna(c): c for c in df.columns}

    # Si las columnas están más abajo (ej. encabezados en otra fila), buscar en las primeras 20 filas
    if not COLUMNAS_REQUERIDAS.issubset(set(mapa_cols.keys())):
        encontrado = False
        for idx_fila in range(min(20, len(df))):
            valores_fila = [_normalizar_nombre_columna(v) for v in df.iloc[idx_fila].values]
            if COLUMNAS_REQUERIDAS.issubset(set(valores_fila)):
                df.columns = [str(v).strip() for v in df.iloc[idx_fila].values]
                df = df.iloc[idx_fila + 1 :].reset_index(drop=True)
                mapa_cols = {_normalizar_nombre_columna(c): c for c in df.columns}
                encontrado = True
                break

        if not encontrado:
            faltantes = COLUMNAS_REQUERIDAS - set(mapa_cols.keys())
            return (
                [],
                f"No se encontraron las columnas requeridas en el Excel: {', '.join(sorted(faltantes))}. "
                f"Columnas detectadas: {', '.join(str(c) for c in df.columns[:10])}",
            )

    col_snies = mapa_cols["codigo_real_snies"]
    col_prog = mapa_cols["pro_descripcion"]
    col_cead = mapa_cols["nombrecead"]
    col_proy = mapa_cols["proy_zon"]
    col_cred = mapa_cols["creditos"]

    registros: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        prog_val = row.get(col_prog)
        cead_val = row.get(col_cead)
        if pd.isna(prog_val) or pd.isna(cead_val):
            continue

        proy_num = _limpiar_numero(row.get(col_proy))
        cred_num = _limpiar_numero(row.get(col_cred))
        prog_str = _limpiar_texto(prog_val, "SIN PROGRAMA")
        cead_str = _limpiar_texto(cead_val, "SIN CENTRO")
        snies_str = _limpiar_texto(row.get(col_snies), "S/N")

        # Ignorar filas de subtotales/totales del Excel (ej. pro_descripcion='Total' o nombrecead='Total')
        if prog_str.upper() in {"TOTAL", "TOTAL GENERAL", "SIN PROGRAMA"}:
            continue
        if cead_str.upper() in {"TOTAL", "TOTAL GENERAL", "SIN CENTRO"}:
            continue

        registros.append(
            {
                "anio": 2026,
                "codigo_real_snies": snies_str,
                "pro_descripcion": prog_str,
                "nombrecead": cead_str,
                "proy_zon": proy_num,
                "creditos": cred_num,
            }
        )

    if not registros:
        return [], "El archivo Excel no contiene filas válidas con datos de créditos."

    return registros, f"Se procesaron {len(registros):,} registros desde el archivo Excel.".replace(",", ".")


def sincronizar_data_ecisa_a_supabase(
    fuente_excel: Any = None,
    reemplazar_existentes: bool = True,
) -> Tuple[bool, int, str]:
    """
    Importa los registros de `data ECISA.xlsx` hacia la tabla `public.proyeccion_creditos_ecisa`
    en Supabase por lotes de 750 filas.
    """
    client = get_supabase_client()
    if client is None:
        return False, 0, "Cliente de Supabase no configurado."

    registros, msg_lectura = leer_y_procesar_data_ecisa(fuente_excel)
    if not registros:
        return False, 0, msg_lectura

    try:
        if reemplazar_existentes:
            # Limpiar registros previos del año 2026 antes de insertar la nueva foto
            client.table("proyeccion_creditos_ecisa").delete().eq("anio", 2026).execute()

        total_insertados = 0
        for inicio in range(0, len(registros), BATCH_SIZE_CREDITOS):
            lote = registros[inicio : inicio + BATCH_SIZE_CREDITOS]
            client.table("proyeccion_creditos_ecisa").insert(lote).execute()
            total_insertados += len(lote)

        limpiar_cache_creditos()
        return (
            True,
            total_insertados,
            f"Sincronización exitosa: {total_insertados:,} registros de créditos proyectados vs realizados (2026) guardados en Supabase.".replace(
                ",", "."
            ),
        )
    except Exception as exc:
        return (
            False,
            0,
            f"Error al guardar en Supabase (asegúrate de ejecutar el bloque SQL de `proyeccion_creditos_ecisa`): {exc}",
        )


@st.cache_data(ttl=300, show_spinner=False)
def obtener_datos_creditos_ecisa() -> pd.DataFrame:
    """
    Consulta los registros de `public.proyeccion_creditos_ecisa` desde Supabase.
    Si la tabla en Supabase aún está vacía pero el archivo `data ECISA.xlsx` existe en la carpeta,
    retorna directamente los datos del Excel local para visualización inmediata.
    """
    client = get_supabase_client()
    filas_totales: List[Dict[str, Any]] = []

    if client is not None:
        try:
            page_size = 1000
            offset = 0
            while True:
                res = (
                    client.table("proyeccion_creditos_ecisa")
                    .select("codigo_real_snies, pro_descripcion, nombrecead, proy_zon, creditos")
                    .range(offset, offset + page_size - 1)
                    .execute()
                )
                lote = res.data or []
                if not lote:
                    break
                filas_totales.extend(lote)
                if len(lote) < page_size:
                    break
                offset += page_size
        except Exception:
            pass

    if not filas_totales and RUTA_DATA_ECISA_DEFECTO.exists():
        registros_locales, _ = leer_y_procesar_data_ecisa(None)
        filas_totales = registros_locales

    if not filas_totales:
        return pd.DataFrame(
            columns=["codigo_real_snies", "pro_descripcion", "nombrecead", "proy_zon", "creditos"]
        )

    df = pd.DataFrame(filas_totales)
    df["proy_zon"] = pd.to_numeric(df["proy_zon"], errors="coerce").fillna(0.0)
    df["creditos"] = pd.to_numeric(df["creditos"], errors="coerce").fillna(0.0)
    return df


def calcular_resumenes_creditos(
    df: pd.DataFrame,
    filtro_cead: Optional[List[str]] = None,
    filtro_programa: Optional[List[str]] = None,
    filtro_snies: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Filtra el DataFrame y calcula las tablas agregadas por:
    - `pro_descripcion` (Programa Académico + SNIES)
    - `nombrecead` (Centro Regional)
    - Cruce detallado `(codigo_real_snies, pro_descripcion, nombrecead)`
    junto con los indicadores globales de cumplimiento (%).
    """
    if df.empty:
        return {
            "total_proyectado": 0.0,
            "total_realizado": 0.0,
            "diferencia": 0.0,
            "cumplimiento_pct": 0.0,
            "df_por_programa": pd.DataFrame(),
            "df_por_cead": pd.DataFrame(),
            "df_detalle": pd.DataFrame(),
        }

    df_f = df.copy()
    if filtro_cead:
        df_f = df_f[df_f["nombrecead"].isin(filtro_cead)]
    if filtro_programa:
        df_f = df_f[df_f["pro_descripcion"].isin(filtro_programa)]
    if filtro_snies:
        df_f = df_f[df_f["codigo_real_snies"].isin(filtro_snies)]

    tot_proy = float(df_f["proy_zon"].sum())
    tot_cred = float(df_f["creditos"].sum())
    dif = tot_cred - tot_proy
    cump_global = (tot_cred / tot_proy * 100.0) if tot_proy > 0 else 0.0

    # 1. Agrupación por Programa (pro_descripcion + codigo_real_snies)
    df_prog = (
        df_f.groupby(["codigo_real_snies", "pro_descripcion"], as_index=False)[["proy_zon", "creditos"]]
        .sum()
    )
    df_prog["diferencia"] = df_prog["creditos"] - df_prog["proy_zon"]
    df_prog["cumplimiento_pct"] = df_prog.apply(
        lambda r: round((r["creditos"] / r["proy_zon"] * 100.0), 1) if r["proy_zon"] > 0 else 0.0,
        axis=1,
    )
    df_prog = df_prog.sort_values(by="creditos", ascending=False).reset_index(drop=True)

    # 2. Agrupación por Centro (nombrecead)
    df_cead = (
        df_f.groupby(["nombrecead"], as_index=False)[["proy_zon", "creditos"]]
        .sum()
    )
    df_cead["diferencia"] = df_cead["creditos"] - df_cead["proy_zon"]
    df_cead["cumplimiento_pct"] = df_cead.apply(
        lambda r: round((r["creditos"] / r["proy_zon"] * 100.0), 1) if r["proy_zon"] > 0 else 0.0,
        axis=1,
    )
    df_cead = df_cead.sort_values(by="creditos", ascending=False).reset_index(drop=True)

    # 3. Agrupación detallada (SNIES + Programa + CEAD)
    df_det = (
        df_f.groupby(["codigo_real_snies", "pro_descripcion", "nombrecead"], as_index=False)[["proy_zon", "creditos"]]
        .sum()
    )
    df_det["diferencia"] = df_det["creditos"] - df_det["proy_zon"]
    df_det["cumplimiento_pct"] = df_det.apply(
        lambda r: round((r["creditos"] / r["proy_zon"] * 100.0), 1) if r["proy_zon"] > 0 else 0.0,
        axis=1,
    )
    df_det = df_det.sort_values(by=["nombrecead", "creditos"], ascending=[True, False]).reset_index(drop=True)

    return {
        "total_proyectado": tot_proy,
        "total_realizado": tot_cred,
        "diferencia": dif,
        "cumplimiento_pct": round(cump_global, 1),
        "df_por_programa": df_prog,
        "df_por_cead": df_cead,
        "df_detalle": df_det,
    }

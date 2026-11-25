"""
Servicio para el Módulo 2: Matrículas por Periodo (Fase 2 - ACTAS 2025 y ACTAS 2026 UNAD).
Optimizado para procesar +15.000 registros sin superar el límite de 1 GB de RAM:
1. Lectura en streaming con `openpyxl(..., read_only=True)` detectando automáticamente la fila de encabezados (fila 14).
2. Inserción/Upsert por lotes (`Batching` de 750 registros por petición).
3. Consultas paginadas (máximo 100 filas en tabla de detalle) y estadísticas consolidadas en caché (`@st.cache_data(ttl=300)`).
"""

import re
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple
import openpyxl
import streamlit as st
from src.db.client import get_supabase_client

DEFAULT_BATCH_SIZE = 750
ARCHIVOS_ACTAS_LOCALES = [
    ("datos_locales/ACTAS 2025.xlsx", 2025),
    ("datos_locales/ACTAS 2026.xlsx", 2026),
    ("ACTAS 2025.xlsx", 2025),
    ("ACTAS 2026.xlsx", 2026),
]


def limpiar_cache_matriculas() -> None:
    """Invalida la caché de estadísticas y consultas de matrículas."""
    obtener_estadisticas_matriculas.clear()
    consultar_matriculas_paginadas.clear()


def _limpiar_str(val: Any) -> Optional[str]:
    if val is None:
        return None
    txt = str(val).strip()
    return txt if txt and txt.lower() not in {"nan", "none", ""} else None


def _limpiar_int(val: Any) -> Optional[int]:
    if val is None:
        return None
    try:
        return int(float(str(val).strip()))
    except Exception:
        return None


def _limpiar_float(val: Any) -> Optional[float]:
    if val is None:
        return None
    try:
        return round(float(str(val).strip().replace(",", ".")), 2)
    except Exception:
        return None


def detectar_anio_archivo(nombre_archivo: str, periodo_ejemplo: Optional[str] = None) -> int:
    """Detecta el año del acta a partir del nombre del archivo (ej. 'ACTAS 2025.xlsx') o del código de periodo."""
    match = re.search(r"(202\d)", str(nombre_archivo))
    if match:
        return int(match.group(1))
    if periodo_ejemplo:
        if str(periodo_ejemplo).startswith("203"):
            return 2025
        if str(periodo_ejemplo).startswith("220") or str(periodo_ejemplo).startswith("221"):
            return 2026
    return 2026


def leer_actas_excel_optimizado(
    fuente_archivo: Any,
    anio_forzado: Optional[int] = None,
) -> Tuple[List[Dict[str, Any]], str]:
    """
    Lee un archivo Excel de Actas de Matrícula UNAD (`ACTAS 2025.xlsx` o `ACTAS 2026.xlsx`)
    usando `openpyxl` en modo `read_only=True` (bajo consumo de memoria).
    Salta automáticamente las primeras 13 filas de encabezado institucional hasta encontrar
    la fila con `['Periodo', 'Tipo Doc', 'Documento', 'Nombres', ...]`.
    """
    try:
        wb = openpyxl.load_workbook(fuente_archivo, read_only=True, data_only=True)
        ws = wb.active
    except Exception as exc:
        return [], f"No se pudo abrir el archivo Excel: {exc}"

    nombre_fuente = getattr(fuente_archivo, "name", str(fuente_archivo))
    encabezados: List[str] = []
    mapa_registros: Dict[Tuple[str, str, str], Dict[str, Any]] = {}

    try:
        for fila in ws.iter_rows(values_only=True):
            if not fila or all(c is None for c in fila[:10]):
                continue

            # Detectar fila de encabezados (fila 14 en los reportes UNAD)
            if not encabezados:
                primeras_celdas = [str(c).strip().lower() if c is not None else "" for c in fila[:6]]
                if "periodo" in primeras_celdas and "documento" in primeras_celdas:
                    encabezados = [str(c).strip() if c is not None else f"col_{i}" for i, c in enumerate(fila)]
                continue

            # Mapear fila de datos
            row_dict = {
                encabezados[i]: fila[i]
                for i in range(min(len(encabezados), len(fila)))
            }

            periodo_val = _limpiar_int(row_dict.get("Periodo"))
            periodo_str = str(periodo_val) if periodo_val is not None else _limpiar_str(row_dict.get("Periodo"))
            doc_val = _limpiar_int(row_dict.get("Documento"))
            doc_str = str(doc_val) if doc_val is not None else _limpiar_str(row_dict.get("Documento"))
            programa_str = _limpiar_str(row_dict.get("Programa")) or "SIN PROGRAMA"

            if not periodo_str or not doc_str:
                continue

            anio_acta = anio_forzado or detectar_anio_archivo(nombre_fuente, periodo_str)

            registro: Dict[str, Any] = {
                "anio_acta": anio_acta,
                "periodo": periodo_str,
                "tipo_documento": _limpiar_str(row_dict.get("Tipo Doc")),
                "documento_estudiante": doc_str,
                "nombre_estudiante": _limpiar_str(row_dict.get("Nombres")),
                "correo_institucional": _limpiar_str(row_dict.get("Correo Institucional")),
                "escuela": _limpiar_str(row_dict.get("Escuela")),
                "programa_academico": programa_str,
                "zona": _limpiar_str(row_dict.get("Zona")),
                "centro_zona": _limpiar_str(row_dict.get("Centro")),
                "tipo_matricula": _limpiar_str(row_dict.get("Condición")),
                "num_matricula": _limpiar_str(row_dict.get("Num de Matricula")),
                "cursos_aprobados": _limpiar_int(row_dict.get("Cursos aprobados")),
                "cursos_matriculados": _limpiar_int(row_dict.get("Cursos matriculados")),
                "promedio": _limpiar_float(row_dict.get("Promedio")),
                "resultado": _limpiar_str(row_dict.get("Resultado")),
                "fecha_ult_acceso": _limpiar_str(row_dict.get("Fecha Ult Acceso")),
                "pais_residencia": _limpiar_str(row_dict.get("Pais de residencia")),
                "departamento_residencia": _limpiar_str(row_dict.get("Departamento de residencia")),
                "ciudad_residencia": _limpiar_str(row_dict.get("Ciudad de residencia")),
                "zona_residencia": _limpiar_str(row_dict.get("Zona de Residencia")),
                "estrato": _limpiar_int(row_dict.get("Estrato")),
                "sexo": _limpiar_str(row_dict.get("Sexo")),
                "necesidades_especiales": _limpiar_str(row_dict.get("Necesidades especiales")),
                "creditos_pregrado": _limpiar_int(row_dict.get("Creditos Pregrado")),
                "creditos_disciplinares": _limpiar_int(row_dict.get("Creditos Disciplinares")),
                "creditos_postgrado": _limpiar_int(row_dict.get("Creditos PostGrado")),
                "creditos_matriculados": _limpiar_int(row_dict.get("Creditos Totales")),
            }

            clave_unica = (periodo_str, doc_str, programa_str)
            mapa_registros[clave_unica] = registro
    finally:
        wb.close()

    registros = list(mapa_registros.values())
    if not registros:
        return [], "No se encontraron registros válidos de matrícula en el archivo."

    return registros, f"Se leyeron {len(registros):,} registros únicos de matrícula.".replace(",", ".")


def insertar_matriculas_por_lotes(
    registros: List[Dict[str, Any]],
    batch_size: int = DEFAULT_BATCH_SIZE,
    progreso_callback: Optional[Callable[[int, int], None]] = None,
) -> Tuple[bool, int, str]:
    """
    Inserta o actualiza (`upsert`) registros de matrículas en bloques de `batch_size` (750 filas)
    para mantener el consumo de memoria muy por debajo de 1 GB de RAM.
    """
    client = get_supabase_client()
    if client is None:
        return False, 0, "Cliente de Supabase no disponible."

    total = len(registros)
    if total == 0:
        return False, 0, "La lista de registros está vacía."

    total_procesados = 0
    try:
        for inicio in range(0, total, batch_size):
            lote = registros[inicio : inicio + batch_size]
            client.table("matriculas").upsert(
                lote,
                on_conflict="periodo,documento_estudiante,programa_academico",
            ).execute()
            total_procesados += len(lote)
            if progreso_callback:
                progreso_callback(total_procesados, total)

        limpiar_cache_matriculas()
        return (
            True,
            total_procesados,
            f"Se sincronizaron exitosamente {total_procesados:,} registros en lotes de {batch_size}.".replace(",", "."),
        )
    except Exception as exc:
        return (
            False,
            total_procesados,
            f"Error al procesar el lote en la fila {total_procesados}: {exc}",
        )


def importar_archivo_actas_a_supabase(
    fuente_archivo: Any,
    anio_forzado: Optional[int] = None,
    batch_size: int = DEFAULT_BATCH_SIZE,
    progreso_callback: Optional[Callable[[int, int], None]] = None,
) -> Tuple[bool, int, str]:
    """Lee un archivo Excel de actas y lo sincroniza por lotes en Supabase."""
    registros, msg_lectura = leer_actas_excel_optimizado(fuente_archivo, anio_forzado=anio_forzado)
    if not registros:
        return False, 0, msg_lectura

    return insertar_matriculas_por_lotes(
        registros=registros,
        batch_size=batch_size,
        progreso_callback=progreso_callback,
    )


def detectar_archivos_actas_locales() -> List[Tuple[str, int]]:
    """Retorna la lista de archivos `ACTAS 2025.xlsx` y `ACTAS 2026.xlsx` presentes en la carpeta del proyecto."""
    encontrados: List[Tuple[str, int]] = []
    for nombre, anio in ARCHIVOS_ACTAS_LOCALES:
        if Path(nombre).exists():
            encontrados.append((nombre, anio))
    return encontrados


# =============================================================================
# CONSULTAS AGREGADAS Y PAGINADAS (PROHIBIDO SELECT * COMPLETO EN UI)
# =============================================================================

@st.cache_data(ttl=300, show_spinner=False)
def obtener_estadisticas_matriculas(filtro_anio: Optional[str] = None) -> Dict[str, Any]:
    """
    Calcula las métricas consolidadas sobre los ~15.500 registros de matrículas
    recorriendo únicamente las columnas de agrupación en bloques paginados ligeros.
    """
    client = get_supabase_client()
    stats: Dict[str, Any] = {
        "total_matriculas": 0,
        "estudiantes_unicos": 0,
        "total_nuevos": 0,
        "total_antiguos": 0,
        "total_continuidad_otros": 0,
        "por_anio": {},
        "por_periodo": {},
        "por_centro": {},
        "por_programa": {},
        "por_condicion": {},
        "por_estrato": {},
        "por_sexo": {},
        "anios_disponibles": [],
        "periodos_disponibles": [],
        "centros_disponibles": [],
        "programas_disponibles": [],
        "condiciones_disponibles": [],
    }

    if client is None:
        return stats

    try:
        # Paginación ligera de 1.000 en 1.000 proyectando solo las columnas livianas de agrupación
        columnas_ligeras = (
            "anio_acta, periodo, documento_estudiante, centro_zona, "
            "programa_academico, tipo_matricula, estrato, sexo"
        )
        page_size = 1000
        offset = 0
        docs_unicos = set()
        anios_set = set()
        periodos_set = set()
        centros_set = set()
        programas_set = set()
        condiciones_set = set()

        while True:
            query = (
                client.table("matriculas")
                .select(columnas_ligeras)
                .range(offset, offset + page_size - 1)
            )
            if filtro_anio and filtro_anio != "Todos":
                query = query.eq("anio_acta", int(filtro_anio))

            res = query.execute()
            filas = res.data or []
            if not filas:
                break

            for r in filas:
                stats["total_matriculas"] += 1
                doc = r.get("documento_estudiante")
                if doc:
                    docs_unicos.add(doc)

                anio = str(r.get("anio_acta") or "2026")
                anios_set.add(anio)
                stats["por_anio"][anio] = stats["por_anio"].get(anio, 0) + 1

                per = str(r.get("periodo") or "S/P")
                periodos_set.add(per)
                stats["por_periodo"][per] = stats["por_periodo"].get(per, 0) + 1

                cen = r.get("centro_zona") or "Sin centro"
                centros_set.add(cen)
                stats["por_centro"][cen] = stats["por_centro"].get(cen, 0) + 1

                prog = r.get("programa_academico") or "Sin programa"
                programas_set.add(prog)
                stats["por_programa"][prog] = stats["por_programa"].get(prog, 0) + 1

                cond = r.get("tipo_matricula") or "Sin condición"
                condiciones_set.add(cond)
                stats["por_condicion"][cond] = stats["por_condicion"].get(cond, 0) + 1

                if cond.lower() == "nuevo":
                    stats["total_nuevos"] += 1
                elif cond.lower() == "antiguo":
                    stats["total_antiguos"] += 1
                else:
                    stats["total_continuidad_otros"] += 1

                est = f"Estrato {r.get('estrato')}" if r.get("estrato") is not None else "No reporta"
                stats["por_estrato"][est] = stats["por_estrato"].get(est, 0) + 1

                sex = r.get("sexo") or "No reporta"
                stats["por_sexo"][sex] = stats["por_sexo"].get(sex, 0) + 1

            if len(filas) < page_size:
                break
            offset += page_size

        stats["estudiantes_unicos"] = len(docs_unicos)
        stats["anios_disponibles"] = sorted(anios_set)
        stats["periodos_disponibles"] = sorted(periodos_set, reverse=True)
        stats["centros_disponibles"] = sorted(centros_set)
        stats["programas_disponibles"] = sorted(programas_set)
        stats["condiciones_disponibles"] = sorted(condiciones_set)

    except Exception:
        pass

    return stats


@st.cache_data(ttl=300, show_spinner=False)
def consultar_matriculas_paginadas(
    filtro_anio: Optional[str] = None,
    filtro_periodo: Optional[str] = None,
    filtro_centro: Optional[str] = None,
    filtro_programa: Optional[str] = None,
    filtro_condicion: Optional[str] = None,
    busqueda_estudiante: Optional[str] = None,
    pagina: int = 1,
    filas_por_pagina: int = 100,
) -> Tuple[List[Dict[str, Any]], int]:
    """
    Consulta páginas de máximo 100 filas aplicando los filtros directamente en PostgreSQL
    (cumpliendo estrictamente la regla de rendimiento de no traer tablas completas a la UI).
    Retorna `(filas_pagina, total_coincidencias)`.
    """
    client = get_supabase_client()
    if client is None:
        return [], 0

    filas_por_pagina = min(max(filas_por_pagina, 10), 100)
    inicio = max(pagina - 1, 0) * filas_por_pagina
    fin = inicio + filas_por_pagina - 1

    columnas_detalle = (
        "anio_acta, periodo, tipo_documento, documento_estudiante, nombre_estudiante, "
        "correo_institucional, programa_academico, centro_zona, tipo_matricula, "
        "cursos_matriculados, cursos_aprobados, promedio, resultado, estrato, sexo, ciudad_residencia"
    )

    try:
        query = (
            client.table("matriculas")
            .select(columnas_detalle, count="exact")
            .order("anio_acta", desc=True)
            .order("periodo", desc=True)
            .order("nombre_estudiante", desc=False)
        )

        if filtro_anio and filtro_anio != "Todos":
            query = query.eq("anio_acta", int(filtro_anio))
        if filtro_periodo and filtro_periodo != "Todos":
            query = query.eq("periodo", filtro_periodo)
        if filtro_centro and filtro_centro != "Todos":
            query = query.eq("centro_zona", filtro_centro)
        if filtro_programa and filtro_programa != "Todos":
            query = query.eq("programa_academico", filtro_programa)
        if filtro_condicion and filtro_condicion != "Todas":
            query = query.eq("tipo_matricula", filtro_condicion)

        if busqueda_estudiante and busqueda_estudiante.strip():
            term = busqueda_estudiante.strip()
            query = query.or_(
                f"documento_estudiante.ilike.%{term}%,nombre_estudiante.ilike.%{term}%,correo_institucional.ilike.%{term}%"
            )

        res = query.range(inicio, fin).execute()
        total_count = res.count if res.count is not None else len(res.data or [])
        return res.data or [], total_count
    except Exception:
        return [], 0

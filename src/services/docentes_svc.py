"""
Servicio de lógica de negocio para el Módulo 1: Gestión Docente y Resoluciones Académicas.
Optimizado para manejar toda la estructura de `Resoluciones Docentes.xlsx` (32 columnas),
incluyendo:
- Normalización automática de niveles de inglés (MCER) conservando el detalle original.
- Desglose inteligente de títulos de Pregrado, Especialización, Maestría y Doctorado desde las columnas del Excel.
- Importación y sincronización masiva por lotes (Upsert) hacia Supabase.
- CRUD completo de docentes, estudios y aportes con caché `@st.cache_data(ttl=300)`.
"""

import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd
import streamlit as st
from src.db.client import get_supabase_client

RUTA_ESTADOS_LOCAL = Path(".streamlit/estados_docentes.json")

RUTA_EXCEL_RESOLUCIONES_DEFECTO = (
    Path("datos_locales/Resoluciones Docentes.xlsx")
    if Path("datos_locales/Resoluciones Docentes.xlsx").exists()
    else Path("Resoluciones Docentes.xlsx")
)

RUTA_EXCEL_INGLES_RESPUESTAS_DEFECTO = (
    Path("datos_locales/Evaluacion de Nivel de Ingles para Docentes (respuestas).xlsx")
    if Path("datos_locales/Evaluacion de Nivel de Ingles para Docentes (respuestas).xlsx").exists()
    else Path("Evaluacion de Nivel de Ingles para Docentes (respuestas).xlsx")
)

CARPETA_DRIVE_SOPORTES_INGLES_URL = (
    "https://drive.google.com/drive/folders/"
    "13FAIUxhKdZLR2PWSr0HBKQfvSdDNREoEEHm4FAPU8J6MwZbtXPgHIj2M6z8lqCVh8LdMSUUC?usp=drive_link"
)

NIVELES_INGLES = [
    "Sin certificar",
    "En caracterización / Prueba",
    "A1",
    "A2",
    "B1",
    "B1+",
    "B2",
    "C1",
    "C2",
]

NIVELES_FORMACION = [
    "Tecnología",
    "Pregrado",
    "Especialización",
    "Maestría",
    "Doctorado",
    "Posdoctorado",
]

NIVELES_FORMACION_MAXIMO = [
    "DOCTORADO",
    "MAGISTER",
    "ESPECIALIZACION",
    "TECNOLOGIA ESPECIALIZADA",
    "PROFESIONAL",
    "TECNOLOGIA",
]

MODALIDADES_VINCULACION = [
    "TIEMPO COMPLETO",
    "MEDIO TIEMPO",
    "HORA CATEDRA",
]

CATEGORIAS_APORTE = [
    "Investigación",
    "Innovación",
    "Internacionalización",
    "Vida Académica",
    "Vida Universitaria",
]

ESTADOS_DOCENTE = {
    "A": "Activo (A)",
    "R": "Retirado (R)",
}


def normalizar_estado_docente(valor: Any) -> str:
    """
    Normaliza el estado del docente a 'A' (Activo) o 'R' (Retirado).
    Por defecto retorna 'A'.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "A"
    txt = str(valor).strip().upper()
    if txt in {"R", "RETIRADO", "RETIRADA", "INACTIVO", "INACTIVA", "RETIRADO (R)"} or txt.startswith("R"):
        return "R"
    return "A"


def etiqueta_estado_docente(codigo: Any) -> str:
    """Retorna la etiqueta legible ('Activo (A)' o 'Retirado (R)') para el código de estado."""
    cod = normalizar_estado_docente(codigo)
    return ESTADOS_DOCENTE.get(cod, "Activo (A)")


def _leer_estados_locales() -> Dict[str, str]:
    if not RUTA_ESTADOS_LOCAL.exists():
        return {}
    try:
        return json.loads(RUTA_ESTADOS_LOCAL.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _guardar_estado_local(documento: str, estado: str) -> None:
    try:
        mapa = _leer_estados_locales()
        mapa[str(documento).strip()] = normalizar_estado_docente(estado)
        RUTA_ESTADOS_LOCAL.parent.mkdir(parents=True, exist_ok=True)
        RUTA_ESTADOS_LOCAL.write_text(json.dumps(mapa, indent=2), encoding="utf-8")
    except Exception:
        pass


def obtener_estado_efectivo_docente(docente: Optional[Dict[str, Any]]) -> str:
    """
    Obtiene el estado ('A' o 'R') de un registro docente desde la columna `estado` de Supabase
    o desde el respaldo local si la columna aún no existe en la tabla remota.
    """
    if not docente:
        return "A"
    if docente.get("estado"):
        return normalizar_estado_docente(docente.get("estado"))
    doc_str = str(docente.get("documento") or "").strip()
    if doc_str:
        mapa_local = _leer_estados_locales()
        if doc_str in mapa_local:
            return normalizar_estado_docente(mapa_local[doc_str])
    return "A"


def limpiar_cache_docentes() -> None:
    """Invalida las consultas en caché cuando se producen cambios en la base de datos."""
    listar_docentes_con_relaciones.clear()
    obtener_estadisticas_consolidadas.clear()


# =============================================================================
# NORMALIZACIÓN Y PARSING DE `Resoluciones Docentes.xlsx`
# =============================================================================

def normalizar_nivel_ingles(texto_raw: Any) -> Tuple[str, Optional[str]]:
    """
    Convierte los valores heterogéneos de la columna `Nivel de Inglés` del Excel
    (ej. 'B2 - ADVANCED *', 'B1+ - UPPER INTERMEDIATE', 'CARACTERIZACION UNAD...', 'No regsitra')
    en un código MCER limpio (`A1`, `A2`, `B1`, `B1+`, `B2`, `C1`, `C2`, `En caracterización / Prueba`, `Sin certificar`)
    y retorna `(nivel_mcer, detalle_original)`.
    """
    if texto_raw is None or (isinstance(texto_raw, float) and pd.isna(texto_raw)):
        return "Sin certificar", None

    limpio = str(texto_raw).strip()
    if not limpio or limpio.lower() in {"nan", "none", "no registra", "no regsitra", "no aplica", "n/a"}:
        return "Sin certificar", limpio if limpio else None

    upper = limpio.upper()

    if "C2" in upper:
        return "C2", limpio
    if "C1" in upper:
        return "C1", limpio
    if "B2" in upper:
        return "B2", limpio
    if "B1+" in upper or "B1 +" in upper or "UPPER INTERMEDIATE" in upper:
        return "B1+", limpio
    if "B1" in upper or "INTERMEDIATE" in upper:
        return "B1", limpio
    if "A2" in upper or "PRE INTERMEDIATE" in upper:
        return "A2", limpio
    if "A1" in upper or "BREAKTHROUGH" in upper:
        return "A1", limpio

    # Casos de pruebas institucionales sin nivel MCER explícito (STEP, Caracterización, Inglés II/III, Proficiency Test)
    palabras_prueba = [
        "STEP",
        "CARACTERIZACION",
        "PRUEBA DE NIVEL",
        "INGLES II",
        "INGLES III",
        "INGLÉS",
        "PROFICIENCY TEST",
        "UNAD ENGLISH",
    ]
    if any(p in upper for p in palabras_prueba):
        return "En caracterización / Prueba", limpio

    return "Sin certificar", limpio


def _es_texto_vacio_o_no_aplica(valor: Any) -> bool:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return True
    txt = str(valor).strip().lower()
    return txt in {
        "",
        "nan",
        "none",
        "no aplica",
        "ni aplica",
        "n/a",
        "no registra",
        "no regsitra",
        "sin informacion",
        "sin información",
        "-",
    }


def extraer_estudios_de_texto(texto_celda: Any, nivel_formacion: str) -> List[Dict[str, Any]]:
    """
    Desglosa una celda del Excel como:
    `Magister en Ciencias Farmacología - Universidad Nacional de Colombia; Máster en Atención Farmacéutica Integral - Universidad de Barcelona`
    en múltiples registros estructurados `{'nivel_formacion', 'titulo_obtenido', 'universidad', 'anio_graduacion'}`.
    """
    if _es_texto_vacio_o_no_aplica(texto_celda):
        return []

    texto = str(texto_celda).strip()
    # Separar múltiples títulos delimitados por ';'
    bloques = [b.strip() for b in texto.split(";") if b.strip()]
    estudios_extraidos: List[Dict[str, Any]] = []

    for bloque in bloques:
        if _es_texto_vacio_o_no_aplica(bloque):
            continue

        # Detectar año si aparece un número de 4 dígitos (1970-2035)
        match_anio = re.search(r"\b(19[7-9]\d|20[0-3]\d)\b", bloque)
        anio = int(match_anio.group(1)) if match_anio else None

        # Separar Título y Universidad por el primer guion ' - '
        if " - " in bloque:
            partes = bloque.split(" - ", 1)
            titulo = partes[0].strip()
            universidad = partes[1].strip()
        elif "-" in bloque and not bloque.upper().endswith("UNAD-CAFAM"):
            partes = bloque.rsplit("-", 1)
            titulo = partes[0].strip()
            universidad = partes[1].strip()
        else:
            titulo = bloque
            universidad = "Por especificar (Importado de Resolución)"

        nivel_efectivo = nivel_formacion
        if nivel_formacion == "Pregrado" and "tecnólog" in titulo.lower() or "tecnolog" in titulo.lower():
            nivel_efectivo = "Tecnología"

        if titulo:
            estudios_extraidos.append(
                {
                    "nivel_formacion": nivel_efectivo,
                    "titulo_obtenido": titulo,
                    "universidad": universidad or "Por especificar",
                    "anio_graduacion": anio,
                }
            )

    return estudios_extraidos


def _limpiar_fecha_iso(valor: Any) -> Optional[str]:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    try:
        dt = pd.to_datetime(valor, errors="coerce")
        if pd.isna(dt):
            return None
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return None


def _limpiar_cumpleanos_mes_dia(valor: Any) -> Optional[str]:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    try:
        dt = pd.to_datetime(valor, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%m-%d")
        txt = str(valor).strip()
        return txt[:10] if txt and txt.lower() != "nan" else None
    except Exception:
        return None


def _limpiar_entero(valor: Any) -> Optional[int]:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    try:
        return int(float(str(valor).strip()))
    except Exception:
        return None


def _limpiar_flotante(valor: Any) -> Optional[float]:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    try:
        return round(float(str(valor).strip()), 2)
    except Exception:
        return None


def _limpiar_texto(valor: Any) -> Optional[str]:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    txt = str(valor).strip()
    return txt if txt and txt.lower() not in {"nan", "none"} else None


def procesar_dataframe_resoluciones(df: pd.DataFrame) -> Tuple[List[Dict[str, Any]], Dict[str, List[Dict[str, Any]]]]:
    """
    Transforma el DataFrame leído de `Resoluciones Docentes.xlsx` en:
    1. `lista_docentes`: diccionarios listos para `upsert` en `public.docentes`.
    2. `mapa_estudios_por_doc`: diccionario `{documento: [estudios...]}` desglosados de
       las columnas `PROFESIONAL`, `ESPECIALIZACIÓN`, `MAESTRÍA` y `DOCTORADO`.
    """
    lista_docentes: List[Dict[str, Any]] = []
    mapa_estudios: Dict[str, List[Dict[str, Any]]] = {}

    for _, row in df.iterrows():
        doc_int = _limpiar_entero(row.get("DOCUMENTO"))
        nombre = _limpiar_texto(row.get("NOMBRES"))
        correo = _limpiar_texto(row.get("Correos Institucionales"))

        if not doc_int or not nombre:
            continue

        documento_str = str(doc_int)
        if not correo:
            correo = f"docente.{documento_str}@unad.edu.co"
        else:
            correo = correo.lower()

        nivel_mcer, detalle_ing = normalizar_nivel_ingles(row.get("Nivel de Inglés"))

        prof_txt = _limpiar_texto(row.get("PROFESIONAL"))
        esp_txt = _limpiar_texto(row.get("ESPECIALIZACIÓN"))
        mae_txt = _limpiar_texto(row.get("MAESTRÍA"))
        doc_txt = _limpiar_texto(row.get("DOCTORADO"))

        periodo_val = _limpiar_entero(row.get("PERIODO ACADEMICO"))
        res_num_val = _limpiar_entero(row.get("RESOLUCION NUMERO"))
        tel_val = _limpiar_entero(row.get("Telefono"))

        registro_docente: Dict[str, Any] = {
            "documento": documento_str,
            "nombre_completo": nombre,
            "correo_institucional": correo,
            "email_personal": _limpiar_texto(row.get("EMAIL PERSONAL")),
            "telefono": str(tel_val) if tel_val else _limpiar_texto(row.get("Telefono")),
            "cumpleanos_mes_dia": _limpiar_cumpleanos_mes_dia(row.get("CUMPLEAÑOS (MES Y DIA no año)")),
            "nivel_ingles": nivel_mcer,
            "detalle_ingles_original": detalle_ing,
            "nivel_formacion_maximo": _limpiar_texto(row.get("NIVEL FORMACION")),
            "formacion_profesional_texto": None if _es_texto_vacio_o_no_aplica(prof_txt) else prof_txt,
            "formacion_especializacion_texto": None if _es_texto_vacio_o_no_aplica(esp_txt) else esp_txt,
            "formacion_maestria_texto": None if _es_texto_vacio_o_no_aplica(mae_txt) else mae_txt,
            "formacion_doctorado_texto": None if _es_texto_vacio_o_no_aplica(doc_txt) else doc_txt,
            "consecutivo_vinculacion": _limpiar_entero(row.get("#")),
            "periodo_academico": str(periodo_val) if periodo_val else _limpiar_texto(row.get("PERIODO ACADEMICO")),
            "estado_aceptacion": _limpiar_texto(row.get("ESTADO ACEPTACION")),
            "modalidad": _limpiar_texto(row.get("MODALIDAD")),
            "horas_vinculacion": _limpiar_entero(row.get("HORAS")),
            "escuela_codigo": _limpiar_entero(row.get("ESCUELA CODIGO")),
            "escuela_descripcion": _limpiar_texto(row.get("ESCUELA DESCRIPCION")),
            "centro_descripcion": _limpiar_texto(row.get("CENTRO DESCRIPCION")),
            "zona_descripcion": _limpiar_texto(row.get("ZONA DESCRIPCION")),
            "programa_codigo": _limpiar_entero(row.get("PROGRAMA CODIGO")),
            "programa_descripcion": _limpiar_texto(row.get("PROGRAMA DESCRIPCION")),
            "fecha_inicio": _limpiar_fecha_iso(row.get("FECHA INICIO")),
            "fecha_termina": _limpiar_fecha_iso(row.get("FECHA TERMINA")),
            "resolucion_numero": str(res_num_val) if res_num_val else _limpiar_texto(row.get("RESOLUCION NUMERO")),
            "resolucion_fecha": _limpiar_fecha_iso(row.get("RESOLUCION FECHA")),
            "salario_mensual": _limpiar_flotante(row.get("SALARIO MENSUAL")),
            "total_vinculacion": _limpiar_flotante(row.get("TOTAL VINCULACION")),
            "dofe": _limpiar_texto(row.get("DOFE")),
            "es_nuevo": _limpiar_texto(row.get("NUEVO")),
            "ultima_vinculacion": _limpiar_fecha_iso(row.get("ULTIMA VINCULACION")),
            "estado": normalizar_estado_docente(row.get("ESTADO") or row.get("estado") or "A"),
        }

        lista_docentes.append(registro_docente)

        estudios_doc: List[Dict[str, Any]] = []
        estudios_doc.extend(extraer_estudios_de_texto(prof_txt, "Pregrado"))
        estudios_doc.extend(extraer_estudios_de_texto(esp_txt, "Especialización"))
        estudios_doc.extend(extraer_estudios_de_texto(mae_txt, "Maestría"))
        estudios_doc.extend(extraer_estudios_de_texto(doc_txt, "Doctorado"))
        mapa_estudios[documento_str] = estudios_doc

    return lista_docentes, mapa_estudios


def importar_resoluciones_excel_a_supabase(
    fuente_excel: Any = None,
    sincronizar_estudios: bool = True,
) -> Tuple[bool, int, int, str]:
    """
    Lee `Resoluciones Docentes.xlsx` (desde el archivo local en la raíz del proyecto o desde un archivo subido),
    realiza `upsert` por lotes en `public.docentes` y opcionalmente sincroniza los títulos desglosados
    en `public.docente_estudios` evitando duplicados.

    Retorna: `(exito, docentes_procesados, estudios_creados, mensaje)`
    """
    client = get_supabase_client()
    if client is None:
        return False, 0, 0, "Cliente de Supabase no configurado."

    try:
        if fuente_excel is None:
            if not RUTA_EXCEL_RESOLUCIONES_DEFECTO.exists():
                return False, 0, 0, f"No se encontró el archivo `{RUTA_EXCEL_RESOLUCIONES_DEFECTO}` en la raíz."
            df = pd.read_excel(RUTA_EXCEL_RESOLUCIONES_DEFECTO, sheet_name=0)
        else:
            df = pd.read_excel(fuente_excel, sheet_name=0)
    except Exception as exc:
        return False, 0, 0, f"Error al leer el archivo Excel de resoluciones: {exc}"

    lista_docentes, mapa_estudios = procesar_dataframe_resoluciones(df)
    if not lista_docentes:
        return False, 0, 0, "No se encontraron filas válidas con DOCUMENTO y NOMBRES en el archivo Excel."

    try:
        # 1. Upsert masivo de docentes (respetando soporte_ingles_url existente al no sobreescribirlo)
        # Si un docente que estaba Retirado ('R') vuelve en las resoluciones del nuevo periodo, pasa a Activo ('A')
        try:
            res_upsert = (
                client.table("docentes")
                .upsert(lista_docentes, on_conflict="documento")
                .execute()
            )
        except Exception as exc_up:
            if "estado" in str(exc_up).lower():
                for d_item in lista_docentes:
                    _guardar_estado_local(d_item["documento"], d_item.get("estado", "A"))
                    d_item.pop("estado", None)
                res_upsert = (
                    client.table("docentes")
                    .upsert(lista_docentes, on_conflict="documento")
                    .execute()
                )
            else:
                raise exc_up
        for d_item in lista_docentes:
            _guardar_estado_local(d_item["documento"], d_item.get("estado", "A"))
        docentes_guardados = res_upsert.data or []
        total_docentes = len(docentes_guardados)

        total_estudios_nuevos = 0
        if sincronizar_estudios and docentes_guardados:
            # Mapear documento -> uuid
            doc_a_id = {d["documento"]: d["id"] for d in docentes_guardados if d.get("documento") and d.get("id")}
            lista_ids = list(doc_a_id.values())

            # Consultar estudios ya existentes para no duplicarlos si se importa varias veces
            res_est_existentes = (
                client.table("docente_estudios")
                .select("docente_id, nivel_formacion, titulo_obtenido")
                .in_("docente_id", lista_ids)
                .execute()
            )
            claves_existentes = {
                (
                    e["docente_id"],
                    (e.get("nivel_formacion") or "").strip().lower(),
                    (e.get("titulo_obtenido") or "").strip().lower(),
                )
                for e in (res_est_existentes.data or [])
            }

            nuevos_estudios_lote: List[Dict[str, Any]] = []
            for doc_str, estudios_Extraidos in mapa_estudios.items():
                doc_uuid = doc_a_id.get(doc_str)
                if not doc_uuid:
                    continue
                for est in estudios_Extraidos:
                    clave = (
                        doc_uuid,
                        est["nivel_formacion"].strip().lower(),
                        est["titulo_obtenido"].strip().lower(),
                    )
                    if clave not in claves_existentes:
                        claves_existentes.add(clave)
                        nuevos_estudios_lote.append(
                            {
                                "docente_id": doc_uuid,
                                "nivel_formacion": est["nivel_formacion"],
                                "titulo_obtenido": est["titulo_obtenido"],
                                "universidad": est["universidad"],
                                "anio_graduacion": est["anio_graduacion"],
                            }
                        )

            if nuevos_estudios_lote:
                client.table("docente_estudios").insert(nuevos_estudios_lote).execute()
                total_estudios_nuevos = len(nuevos_estudios_lote)

        limpiar_cache_docentes()
        return (
            True,
            total_docentes,
            total_estudios_nuevos,
            f"Importación exitosa: {total_docentes} docentes sincronizados y {total_estudios_nuevos} títulos académicos desglosados.",
        )
    except Exception as exc:
        return (
            False,
            0,
            0,
            f"Error al sincronizar con Supabase (asegúrate de haber ejecutado el nuevo `src/db/schema.sql`): {exc}",
        )


def previsualizar_excel_local() -> Optional[pd.DataFrame]:
    """Lee el archivo `Resoluciones Docentes.xlsx` local si existe para previsualización rápida."""
    if not RUTA_EXCEL_RESOLUCIONES_DEFECTO.exists():
        return None
    try:
        df = pd.read_excel(RUTA_EXCEL_RESOLUCIONES_DEFECTO, sheet_name=0)
        df = df.dropna(subset=["DOCUMENTO", "NOMBRES"])
        return df
    except Exception:
        return None


def importar_respuestas_ingles_a_supabase(fuente_excel: Any = None) -> Tuple[bool, int, int, str]:
    """
    Lee `Evaluacion de Nivel de Ingles para Docentes (respuestas).xlsx` (formulario de Google Forms),
    cruza cada docente por su `Numero de identificacion` (o por coincidencia de nombre si escribieron mal un dígito),
    y actualiza en `public.docentes`:
    - `nivel_ingles` (normalizado a MCER: A1, A2, B1, B1+, B2, C1, C2, En caracterización / Prueba, Sin certificar)
    - `detalle_ingles_original`
    - `soporte_ingles_url` (enlace directo `https://drive.google.com/open?id=...` al PDF adjunto en Drive).

    Retorna: `(exito, actualizados_en_bd, no_encontrados, mensaje)`
    """
    client = get_supabase_client()
    if client is None:
        return False, 0, 0, "Cliente de Supabase no configurado."

    try:
        if fuente_excel is None:
            if not RUTA_EXCEL_INGLES_RESPUESTAS_DEFECTO.exists():
                return (
                    False,
                    0,
                    0,
                    f"No se encontró el archivo `{RUTA_EXCEL_INGLES_RESPUESTAS_DEFECTO}`.",
                )
            df = pd.read_excel(RUTA_EXCEL_INGLES_RESPUESTAS_DEFECTO, sheet_name=0)
        else:
            df = pd.read_excel(fuente_excel, sheet_name=0)
    except Exception as exc:
        return False, 0, 0, f"Error al leer el archivo de respuestas de inglés: {exc}"

    if df.empty or len(df.columns) < 4:
        return False, 0, 0, "El archivo no tiene las columnas esperadas del formulario de evaluación de inglés."

    # Identificar columnas por posición o palabra clave
    col_nombre = df.columns[1]
    col_doc = df.columns[2]
    col_nivel = df.columns[3]
    col_url = df.columns[4] if len(df.columns) > 4 else None

    for c in df.columns:
        cl = str(c).lower()
        if "identificacion" in cl or "documento" in cl:
            col_doc = c
        elif "nombre" in cl:
            col_nombre = c
        elif "adjunte" in cl or "certificado" in cl or "soporte" in cl:
            col_url = c
        elif "seleccione su nivel" in cl or "suficiencia" in cl or "marco comun" in cl:
            col_nivel = c

    try:
        # Traer docentes actuales en BD para cruce rápido por documento o nombre
        res_bd = client.table("docentes").select("id, documento, nombre_completo").execute()
        docentes_bd = res_bd.data or []
        mapa_doc_a_id = {str(d["documento"]).strip(): d["id"] for d in docentes_bd if d.get("documento")}

        def _norm_nombre(n: str) -> str:
            reemplazos = str.maketrans("ÁÉÍÓÚÜÑáéíóúüñ", "AEIOUUNAEIOUUN")
            return re.sub(r"[^A-Z ]+", "", str(n).upper().translate(reemplazos)).strip()

        mapa_nombre_a_id = {_norm_nombre(d["nombre_completo"]): d["id"] for d in docentes_bd if d.get("nombre_completo")}

        actualizados = 0
        no_encontrados = 0

        for _, row in df.iterrows():
            raw_doc = row.get(col_doc)
            raw_nom = _limpiar_texto(row.get(col_nombre)) or ""
            raw_niv = row.get(col_nivel)
            raw_url = _limpiar_texto(row.get(col_url)) if col_url else None

            # Limpiar número de identificación (quitar puntos, espacios, comas)
            doc_str = re.sub(r"[^0-9]", "", str(raw_doc or ""))
            if not doc_str and not raw_nom:
                continue

            nivel_mcer, detalle_orig = normalizar_nivel_ingles(raw_niv)

            docente_uuid = mapa_doc_a_id.get(doc_str)
            if not docente_uuid and raw_nom:
                nom_norm = _norm_nombre(raw_nom)
                docente_uuid = mapa_nombre_a_id.get(nom_norm)

            if docente_uuid:
                payload_upd: Dict[str, Any] = {
                    "nivel_ingles": nivel_mcer,
                    "detalle_ingles_original": f"{detalle_orig} (Encuesta Inglés)" if detalle_orig else nivel_mcer,
                }
                if raw_url and raw_url.startswith("http"):
                    payload_upd["soporte_ingles_url"] = raw_url

                client.table("docentes").update(payload_upd).eq("id", docente_uuid).execute()
                actualizados += 1
            else:
                no_encontrados += 1

        limpiar_cache_docentes()
        return (
            True,
            actualizados,
            no_encontrados,
            f"Se actualizaron {actualizados} docentes con su nivel de inglés y enlace directo a su certificado PDF en Google Drive "
            f"({no_encontrados} respuestas no coincidieron con la planta actual de resoluciones).",
        )
    except Exception as exc:
        return False, 0, 0, f"Error al actualizar soportes de inglés en Supabase: {exc}"


# =============================================================================
# OPERACIONES CRUD INDIVIDUALES
# =============================================================================

def buscar_docente(criterio: str) -> Tuple[Optional[Dict[str, Any]], str]:
    """
    Busca un docente por número de documento exacto, por correo institucional o por coincidencia parcial de nombre.
    """
    client = get_supabase_client()
    if client is None:
        return None, "Cliente de Supabase no configurado."

    valor = criterio.strip()
    if not valor:
        return None, "Ingresa un número de documento, correo institucional o nombre para buscar."

    try:
        # 1. Búsqueda exacta por documento
        res_doc = (
            client.table("docentes")
            .select("*")
            .eq("documento", valor)
            .limit(1)
            .execute()
        )
        if res_doc.data:
            doc_encontrado = res_doc.data[0]
            doc_encontrado["estado"] = obtener_estado_efectivo_docente(doc_encontrado)
            return doc_encontrado, "Docente encontrado por número de documento."

        # 2. Búsqueda por correo institucional
        res_mail = (
            client.table("docentes")
            .select("*")
            .ilike("correo_institucional", valor)
            .limit(1)
            .execute()
        )
        if res_mail.data:
            doc_encontrado = res_mail.data[0]
            doc_encontrado["estado"] = obtener_estado_efectivo_docente(doc_encontrado)
            return doc_encontrado, "Docente encontrado por correo institucional."

        # 3. Búsqueda por nombre parcial
        res_nom = (
            client.table("docentes")
            .select("*")
            .ilike("nombre_completo", f"%{valor}%")
            .limit(1)
            .execute()
        )
        if res_nom.data:
            doc_encontrado = res_nom.data[0]
            doc_encontrado["estado"] = obtener_estado_efectivo_docente(doc_encontrado)
            return doc_encontrado, f"Docente encontrado por coincidencia de nombre: {doc_encontrado.get('nombre_completo')}."

        return None, "No se encontró ningún docente con ese criterio. Puedes crear su perfil o importar `Resoluciones Docentes.xlsx`."
    except Exception as exc:
        return None, f"Error al consultar docente en Supabase: {exc}"


def guardar_o_actualizar_docente(
    documento: str,
    nombre_completo: str,
    correo_institucional: str,
    nivel_ingles: str,
    soporte_ingles_url: Optional[str] = None,
    docente_id: Optional[str] = None,
    datos_adicionales: Optional[Dict[str, Any]] = None,
) -> Tuple[bool, Optional[Dict[str, Any]], str]:
    """
    Crea o actualiza el perfil de un docente incluyendo sus datos de contacto, inglés y vinculación/resolución.
    """
    client = get_supabase_client()
    if client is None:
        return False, None, "Cliente de Supabase no disponible."

    doc_limpio = documento.strip()
    nombre_limpio = nombre_completo.strip()
    correo_limpio = correo_institucional.strip().lower()

    if not doc_limpio or not nombre_limpio or not correo_limpio:
        return False, None, "Documento, nombre completo y correo institucional son obligatorios."

    if nivel_ingles not in NIVELES_INGLES:
        return False, None, f"Nivel de inglés inválido. Opciones permitidas: {', '.join(NIVELES_INGLES)}."

    payload: Dict[str, Any] = {
        "documento": doc_limpio,
        "nombre_completo": nombre_limpio,
        "correo_institucional": correo_limpio,
        "nivel_ingles": nivel_ingles,
    }
    if soporte_ingles_url is not None:
        payload["soporte_ingles_url"] = soporte_ingles_url

    if datos_adicionales:
        for k, v in datos_adicionales.items():
            payload[k] = v

    estado_solicitado: Optional[str] = None
    if "estado" in payload and payload["estado"] is not None:
        estado_solicitado = normalizar_estado_docente(payload["estado"])
        payload["estado"] = estado_solicitado
        _guardar_estado_local(doc_limpio, estado_solicitado)

    try:
        if docente_id:
            try:
                res = (
                    client.table("docentes")
                    .update(payload)
                    .eq("id", docente_id)
                    .execute()
                )
            except Exception as exc_upd:
                if "estado" in str(exc_upd).lower() and "estado" in payload:
                    payload_sin_estado = {k: v for k, v in payload.items() if k != "estado"}
                    res = (
                        client.table("docentes")
                        .update(payload_sin_estado)
                        .eq("id", docente_id)
                        .execute()
                    )
                else:
                    raise exc_upd
            limpiar_cache_docentes()
            if res.data:
                doc_res = res.data[0]
                doc_res["estado"] = estado_solicitado or obtener_estado_efectivo_docente(doc_res)
                return True, doc_res, "Perfil y datos de resolución actualizados correctamente."
            return False, None, "No se pudo actualizar el perfil del docente."
        else:
            try:
                res = (
                    client.table("docentes")
                    .upsert(payload, on_conflict="documento")
                    .execute()
                )
            except Exception as exc_ins:
                if "estado" in str(exc_ins).lower() and "estado" in payload:
                    payload_sin_estado = {k: v for k, v in payload.items() if k != "estado"}
                    res = (
                        client.table("docentes")
                        .upsert(payload_sin_estado, on_conflict="documento")
                        .execute()
                    )
                else:
                    raise exc_ins
            limpiar_cache_docentes()
            if res.data:
                doc_res = res.data[0]
                doc_res["estado"] = estado_solicitado or obtener_estado_efectivo_docente(doc_res)
                return True, doc_res, "Perfil docente registrado correctamente."
            return False, None, "No se recibió confirmación al registrar el docente."
    except Exception as exc:
        return False, None, f"Error al guardar perfil docente en Supabase: {exc}"


def obtener_estudios_docente(docente_id: str) -> List[Dict[str, Any]]:
    """Retorna la lista de títulos académicos de un docente."""
    client = get_supabase_client()
    if client is None or not docente_id:
        return []
    try:
        res = (
            client.table("docente_estudios")
            .select("*")
            .eq("docente_id", docente_id)
            .order("anio_graduacion", desc=True)
            .execute()
        )
        return res.data or []
    except Exception:
        return []


def agregar_estudio_docente(
    docente_id: str,
    nivel_formacion: str,
    titulo_obtenido: str,
    universidad: str,
    anio_graduacion: Optional[int],
    soporte_url: Optional[str] = None,
) -> Tuple[bool, str]:
    """Inserta un nuevo título académico asociado a un docente."""
    client = get_supabase_client()
    if client is None:
        return False, "Cliente de Supabase no disponible."

    if not docente_id:
        return False, "Primero debes guardar el perfil básico del docente."

    if nivel_formacion not in NIVELES_FORMACION:
        return False, "Nivel de formación inválido."

    if not titulo_obtenido.strip() or not universidad.strip():
        return False, "El título obtenido y la universidad son obligatorios."

    payload: Dict[str, Any] = {
        "docente_id": docente_id,
        "nivel_formacion": nivel_formacion,
        "titulo_obtenido": titulo_obtenido.strip(),
        "universidad": universidad.strip(),
        "anio_graduacion": int(anio_graduacion) if anio_graduacion else None,
        "soporte_url": soporte_url,
    }

    try:
        client.table("docente_estudios").insert(payload).execute()
        limpiar_cache_docentes()
        return True, "Estudio registrado exitosamente."
    except Exception as exc:
        return False, f"Error al registrar estudio: {exc}"


def eliminar_estudio_docente(estudio_id: str) -> Tuple[bool, str]:
    """Elimina un estudio específico por su UUID."""
    client = get_supabase_client()
    if client is None:
        return False, "Cliente de Supabase no disponible."
    try:
        client.table("docente_estudios").delete().eq("id", estudio_id).execute()
        limpiar_cache_docentes()
        return True, "Estudio eliminado correctamente."
    except Exception as exc:
        return False, f"Error al eliminar estudio: {exc}"


def obtener_aportes_docente(docente_id: str) -> List[Dict[str, Any]]:
    """Retorna la lista de aportes académicos de un docente."""
    client = get_supabase_client()
    if client is None or not docente_id:
        return []
    try:
        res = (
            client.table("docente_aportes")
            .select("*")
            .eq("docente_id", docente_id)
            .order("periodo_academico", desc=True)
            .execute()
        )
        return res.data or []
    except Exception:
        return []


def agregar_aporte_docente(
    docente_id: str,
    categoria: str,
    titulo_aporte: str,
    descripcion: str,
    periodo_academico: str,
    soporte_url: Optional[str] = None,
) -> Tuple[bool, str]:
    """Inserta un nuevo aporte académico asociado a un docente."""
    client = get_supabase_client()
    if client is None:
        return False, "Cliente de Supabase no disponible."

    if not docente_id:
        return False, "Primero debes guardar el perfil básico del docente."

    if categoria not in CATEGORIAS_APORTE:
        return False, "Categoría de aporte inválida."

    if not titulo_aporte.strip() or not periodo_academico.strip():
        return False, "El título del aporte y el periodo académico son obligatorios."

    payload: Dict[str, Any] = {
        "docente_id": docente_id,
        "categoria": categoria,
        "titulo_aporte": titulo_aporte.strip(),
        "descripcion": descripcion.strip() if descripcion else None,
        "periodo_academico": periodo_academico.strip(),
        "soporte_url": soporte_url,
    }

    try:
        client.table("docente_aportes").insert(payload).execute()
        limpiar_cache_docentes()
        return True, "Aporte registrado exitosamente."
    except Exception as exc:
        return False, f"Error al registrar aporte: {exc}"


def eliminar_aporte_docente(aporte_id: str) -> Tuple[bool, str]:
    """Elimina un aporte específico por su UUID."""
    client = get_supabase_client()
    if client is None:
        return False, "Cliente de Supabase no disponible."
    try:
        client.table("docente_aportes").delete().eq("id", aporte_id).execute()
        limpiar_cache_docentes()
        return True, "Aporte eliminado correctamente."
    except Exception as exc:
        return False, f"Error al eliminar aporte: {exc}"


# =============================================================================
# CONSULTAS CONSOLIDADAS Y FILTROS GERENCIALES CON CACHÉ
# =============================================================================

@st.cache_data(ttl=300, show_spinner=False)
def listar_docentes_con_relaciones(
    filtro_nivel_ingles: Optional[str] = None,
    filtro_categoria_aporte: Optional[str] = None,
    filtro_modalidad: Optional[str] = None,
    filtro_centro: Optional[str] = None,
    filtro_programa: Optional[str] = None,
    filtro_estado: Optional[str] = "A",
    busqueda_texto: Optional[str] = None,
    limite: int = 150,
) -> List[Dict[str, Any]]:
    """
    Consulta docentes junto con sus estudios, aportes y datos de resolución.
    Permite filtrar por `estado`:
    - `'A'` o `'Activo (A)'`: Solo docentes Activos (valor por defecto para reportes).
    - `'R'` o `'Retirado (R)'`: Solo docentes Retirados.
    - `'Todos'` o `None`: Todos los docentes.
    """
    client = get_supabase_client()
    if client is None:
        return []

    try:
        query = (
            client.table("docentes")
            .select("*, docente_estudios(*), docente_aportes(*)")
            .order("nombre_completo", desc=False)
            .limit(limite)
        )

        if filtro_nivel_ingles and filtro_nivel_ingles != "Todos":
            query = query.eq("nivel_ingles", filtro_nivel_ingles)

        if filtro_modalidad and filtro_modalidad != "Todas":
            query = query.eq("modalidad", filtro_modalidad)

        if filtro_centro and filtro_centro != "Todos":
            query = query.eq("centro_descripcion", filtro_centro)

        if filtro_programa and filtro_programa != "Todos":
            query = query.eq("programa_descripcion", filtro_programa)

        if busqueda_texto and busqueda_texto.strip():
            term = busqueda_texto.strip()
            query = query.or_(
                f"nombre_completo.ilike.%{term}%,documento.ilike.%{term}%,correo_institucional.ilike.%{term}%,resolucion_numero.ilike.%{term}%"
            )

        res = query.execute()
        docentes: List[Dict[str, Any]] = res.data or []

        for d in docentes:
            d["estado"] = obtener_estado_efectivo_docente(d)

        if filtro_estado and filtro_estado not in {"Todos", "Todas"}:
            cod_filtro = normalizar_estado_docente(filtro_estado)
            docentes = [d for d in docentes if d.get("estado", "A") == cod_filtro]

        if filtro_categoria_aporte and filtro_categoria_aporte != "Todas":
            docentes = [
                d
                for d in docentes
                if any(
                    a.get("categoria") == filtro_categoria_aporte
                    for a in (d.get("docente_aportes") or [])
                )
            ]

        return docentes
    except Exception:
        return []


@st.cache_data(ttl=300, show_spinner=False)
def obtener_estadisticas_consolidadas() -> Dict[str, Any]:
    """
    Calcula las métricas gerenciales consolidadas incluyendo la información de resoluciones.
    IMPORTANTE: Excluye automáticamente de todas las estadísticas a los docentes cuyo
    campo `estado` sea Retirado ('R'), contabilizando únicamente a los docentes Activos ('A').
    """
    client = get_supabase_client()
    stats: Dict[str, Any] = {
        "total_docentes": 0,
        "total_docentes_retirados": 0,
        "total_docentes_registrados": 0,
        "total_estudios": 0,
        "total_posgrados": 0,
        "total_aportes": 0,
        "total_horas_vinculacion": 0,
        "suma_total_vinculacion": 0.0,
        "docentes_con_soporte_ingles": 0,
        "distribucion_ingles": {n: 0 for n in NIVELES_INGLES},
        "distribucion_formacion": {nf: 0 for nf in NIVELES_FORMACION},
        "distribucion_formacion_maxima": {},
        "distribucion_modalidad": {},
        "distribucion_centro": {},
        "distribucion_programa": {},
        "distribucion_aportes": {cat: 0 for cat in CATEGORIAS_APORTE},
        "centros_disponibles": [],
        "programas_disponibles": [],
        "modalidades_disponibles": [],
    }

    if client is None:
        return stats

    try:
        res_doc = (
            client.table("docentes")
            .select("*")
            .execute()
        )
        todos_docentes = res_doc.data or []
        stats["total_docentes_registrados"] = len(todos_docentes)

        docentes_activos: List[Dict[str, Any]] = []
        ids_activos = set()

        for d in todos_docentes:
            est_doc = obtener_estado_efectivo_docente(d)
            d["estado"] = est_doc
            if est_doc == "R":
                stats["total_docentes_retirados"] += 1
            else:
                docentes_activos.append(d)
                if d.get("id"):
                    ids_activos.add(d["id"])

        stats["total_docentes"] = len(docentes_activos)

        centros_set = set()
        programas_set = set()
        modalidades_set = set()

        for d in docentes_activos:
            niv = d.get("nivel_ingles") or "Sin certificar"
            stats["distribucion_ingles"][niv] = stats["distribucion_ingles"].get(niv, 0) + 1
            if d.get("soporte_ingles_url"):
                stats["docentes_con_soporte_ingles"] += 1

            nf_max = d.get("nivel_formacion_maximo") or "Sin especificar"
            stats["distribucion_formacion_maxima"][nf_max] = (
                stats["distribucion_formacion_maxima"].get(nf_max, 0) + 1
            )

            mod = d.get("modalidad") or "Sin especificar"
            stats["distribucion_modalidad"][mod] = stats["distribucion_modalidad"].get(mod, 0) + 1
            if d.get("modalidad"):
                modalidades_set.add(d["modalidad"])

            centro = d.get("centro_descripcion") or "Sin especificar"
            stats["distribucion_centro"][centro] = stats["distribucion_centro"].get(centro, 0) + 1
            if d.get("centro_descripcion"):
                centros_set.add(d["centro_descripcion"])

            prog = d.get("programa_descripcion") or "Sin especificar"
            stats["distribucion_programa"][prog] = stats["distribucion_programa"].get(prog, 0) + 1
            if d.get("programa_descripcion"):
                programas_set.add(d["programa_descripcion"])

            if d.get("horas_vinculacion"):
                stats["total_horas_vinculacion"] += int(d["horas_vinculacion"])
            if d.get("total_vinculacion"):
                stats["suma_total_vinculacion"] += float(d["total_vinculacion"])

        # Incluir todos los centros, programas y modalidades conocidos para los filtros del directorio
        for d in todos_docentes:
            if d.get("centro_descripcion"):
                centros_set.add(d["centro_descripcion"])
            if d.get("programa_descripcion"):
                programas_set.add(d["programa_descripcion"])
            if d.get("modalidad"):
                modalidades_set.add(d["modalidad"])

        stats["centros_disponibles"] = sorted(centros_set)
        stats["programas_disponibles"] = sorted(programas_set)
        stats["modalidades_disponibles"] = sorted(modalidades_set)

        res_est = (
            client.table("docente_estudios")
            .select("id, docente_id, nivel_formacion")
            .execute()
        )
        estudios = [
            e for e in (res_est.data or [])
            if e.get("docente_id") in ids_activos
        ]
        stats["total_estudios"] = len(estudios)
        for e in estudios:
            nf = e.get("nivel_formacion") or "Pregrado"
            stats["distribucion_formacion"][nf] = stats["distribucion_formacion"].get(nf, 0) + 1
            if nf in {"Especialización", "Maestría", "Doctorado", "Posdoctorado"}:
                stats["total_posgrados"] += 1

        res_apo = (
            client.table("docente_aportes")
            .select("id, docente_id, categoria")
            .execute()
        )
        aportes = [
            a for a in (res_apo.data or [])
            if a.get("docente_id") in ids_activos
        ]
        stats["total_aportes"] = len(aportes)
        for a in aportes:
            cat = a.get("categoria") or "Investigación"
            stats["distribucion_aportes"][cat] = stats["distribucion_aportes"].get(cat, 0) + 1

    except Exception:
        pass

    return stats

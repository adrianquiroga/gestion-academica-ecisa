"""
Servicio de lógica de negocio para el Módulo 4: Gestión y Analítica de Homologaciones 2026 (ECISA - UNAD).
Procesa el archivo `Homologaciones 2026.xlsx` (Consejos de Escuela N° 2 al N° 19 de 2026):
- Une automáticamente las filas de continuación donde `Programa de origen` se partió en dos líneas en el Excel.
- Corrige automáticamente la inversión de las columnas `Nombre` e `Identidad` a partir del Acuerdo N° 370 (Consejo N° 13 y N° 19).
- Normaliza nombres de Programas a Cursar (ECISA), Universidades de Origen (UNAD, SENA, Externas), Centros, Niveles de Formación y Fechas de Acuerdo.
- Clasifica cada registro en Homologación `Interna` (UNAD) o `Externa` (SENA y otras IES), separando también los Acuerdos Normativos/Generales de Consejo de Escuela.
- Sincroniza los registros con la tabla `public.homologaciones` en Supabase y provee respaldo local en JSON para todos los metadatos de Consejo y Acuerdo.
"""

from datetime import date
import json
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd
import streamlit as st
from src.db.client import get_supabase_client

RUTA_EXCEL_HOMOLOGACIONES_DEFECTO = (
    Path("datos_locales/Homologaciones 2026.xlsx")
    if Path("datos_locales/Homologaciones 2026.xlsx").exists()
    else Path("Homologaciones 2026.xlsx")
)

RUTA_JSON_HOMOLOGACIONES_FALLBACK = Path("src/data/homologaciones_2026_cache.json")
BATCH_SIZE_HOMOLOGACIONES = 500

MESES_ES = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "abrild": 4,
    "abrilde": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}

NOMBRES_MESES_ORDEN = {
    1: "01. Enero",
    2: "02. Febrero",
    3: "03. Marzo",
    4: "04. Abril",
    5: "05. Mayo",
    6: "06. Junio",
    7: "07. Julio",
    8: "08. Agosto",
    9: "09. Septiembre",
    10: "10. Octubre",
    11: "11. Noviembre",
    12: "12. Diciembre",
}


def limpiar_cache_homologaciones() -> None:
    """Invalida la caché de consultas de homologaciones."""
    obtener_datos_homologaciones_2026.clear()


def _es_vacio(val: Any) -> bool:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return True
    txt = str(val).strip()
    return not txt or txt.lower() in {"nan", "none", "-"}


def _limpiar_texto(val: Any, defecto: str = "") -> str:
    if _es_vacio(val):
        return defecto
    txt = re.sub(r"\s+", " ", str(val)).strip()
    if txt.endswith(".0") and txt[:-2].isdigit():
        return txt[:-2]
    return txt


def _es_documento_identidad(texto: str) -> bool:
    limpio = re.sub(r"[\.\,\s\-]", "", texto.strip())
    return limpio.isdigit() and 5 <= len(limpio) <= 15


def _extraer_fecha_y_numero_acuerdo(acuerdo_txt: str, consejo_txt: str) -> Tuple[Optional[str], str, int]:
    """
    Extrae `(fecha_iso, mes_etiqueta, numero_acuerdo)` desde textos como:
    `ACUERDO NÚMERO 639 DEL 29 DE SEPTIEMBRE DE 2026` o `ACUERDO NÚMERO 131 DEL 17 DE ABRILDE 2026`.
    """
    texto_busq = f"{acuerdo_txt} {consejo_txt}".lower()

    # Extraer número de acuerdo
    num_acuerdo = 0
    m_num = re.search(r"acuerdo\s+n[úu]mero\s*(\d+)", acuerdo_txt.lower())
    if m_num:
        num_acuerdo = int(m_num.group(1))

    # Extraer día, mes y año
    m_fecha = re.search(
        r"(\d{1,2})\s*(?:de)?\s*(enero|febrero|marzo|abrilde|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)\s*(?:de|del)?\s*(202\d)",
        texto_busq,
    )
    if m_fecha:
        dia = int(m_fecha.group(1))
        mes_nombre = m_fecha.group(2)
        anio = int(m_fecha.group(3))
        mes_num = MESES_ES.get(mes_nombre, 1)
        try:
            f_iso = date(anio, mes_num, min(max(dia, 1), 28 if mes_num == 2 else 30 if mes_num in {4, 6, 9, 11} else 31)).isoformat()
            return f_iso, NOMBRES_MESES_ORDEN.get(mes_num, "Sin mes"), num_acuerdo
        except Exception:
            pass

    return None, "Sin fecha", num_acuerdo


def _normalizar_consejo(consejo_raw: str) -> str:
    txt = re.sub(r"\s+", " ", consejo_raw).strip()
    m = re.search(r"n[°ºo]?\s*(\d+)", txt.lower())
    if not m:
        return txt or "Sin Consejo"
    num = int(m.group(1))
    # Extraer fecha legible del consejo
    m_f = re.search(
        r"(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre).*?(202\d)",
        txt.lower(),
    )
    rest = txt[m.end():].strip(" -delDE")
    return f"Consejo N° {num:02d} ({rest})" if rest else f"Consejo N° {num:02d}"


def _normalizar_programa_cursar(prog_raw: str) -> Tuple[str, str]:
    """
    Retorna `(programa_base_homologado, programa_detalle_original)` para agrupar limpiamente
    variantes de mayúsculas/acentos y resoluciones en ECISA.
    """
    txt = _limpiar_texto(prog_raw, "SIN ESPECIFICAR")
    up = (
        txt.upper()
        .replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
    )

    if "PROFESIONAL EN SEGURIDAD Y SALUD EN EL TRABAJO" in up:
        return "Profesional en Seguridad y Salud en el Trabajo", txt
    if "TECNOLOGIA EN SEGURIDAD Y SALUD EN EL TRABAJO" in up:
        return "Tecnología en Seguridad y Salud en el Trabajo", txt
    if "ADMINISTRACION EN SALUD" in up:
        if "24218" in up:
            return "Administración en Salud (Res. 24218)", txt
        return "Administración en Salud", txt
    if "REGENCIA DE FARMACIA" in up:
        if "1309" in up:
            return "Tecnología en Regencia de Farmacia (Res. 1309)", txt
        if "08200" in up or "8200" in up:
            return "Tecnología en Regencia de Farmacia (Res. 08200)", txt
        return "Tecnología en Regencia de Farmacia", txt
    if "RADIOLOGIA E IMAGENES DIAGNOSTICAS" in up:
        return "Tecnología en Radiología e Imágenes Diagnósticas", txt
    if "GERENCIA Y AUDITORIA DE LA CALIDAD EN SALUD" in up:
        return "Esp. en Gerencia y Auditoría de la Calidad en Salud", txt
    if "BIOETICA" in up:
        return "Maestría en Bioética", txt
    if "SALUD PUBLICA" in up:
        return "Maestría en Salud Pública", txt

    return txt.title(), txt


def _normalizar_universidad_y_tipo(univ_raw: str) -> Tuple[str, str]:
    """
    Normaliza el nombre de la institución de origen y clasifica la homologación en:
    - `'Interna'`: Universidad Nacional Abierta y a Distancia (UNAD)
    - `'Externa'`: SENA u otras universidades/instituciones externas.
    """
    txt = _limpiar_texto(univ_raw, "Por especificar")
    up = (
        txt.upper()
        .replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
    )

    if "UNAD" in up or "ABIERTA Y A DISTANCIA" in up:
        return "UNAD - Universidad Nacional Abierta y a Distancia", "Interna"
    if "SENA" in up or "SERVICIO NACIONAL DE APRENDIZAJE" in up:
        return "SENA - Servicio Nacional de Aprendizaje", "Externa"
    if "UNIMINUTO" in up or "MINUTO DE DIOS" in up:
        return "Corporación Universitaria Minuto de Dios (UNIMINUTO)", "Externa"
    if "ANTIOQUIA" in up:
        return "Universidad de Antioquia", "Externa"
    if "NACIONAL DE COLOMBIA" in up:
        return "Universidad Nacional de Colombia", "Externa"
    if "POLITECNICO GRANCOLOMBIANO" in up:
        return "Politécnico Grancolombiano", "Externa"
    if "UNINPAHU" in up:
        return "Fundación Universitaria UNINPAHU", "Externa"
    if "UNINAVARRA" in up:
        return "Fundación Universitaria Navarra (UNINAVARRA)", "Externa"

    return txt.title(), "Externa"


def _normalizar_nivel_origen(nivel_raw: str) -> str:
    txt = _limpiar_texto(nivel_raw, "Sin especificar")
    up = (
        txt.upper()
        .replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
    )
    if "TECNOLOG" in up:
        return "Tecnológico"
    if "UNIVERSITAR" in up or "PREGADO" in up or "PREGRADO" in up or "PROFESIONAL" in up:
        if "ESPECIALIZACION" in up:
            return "Especialización"
        if "TECNICA" in up or "TECNICO" in up:
            return "Técnico Profesional"
        return "Universitario / Profesional"
    if "TECNIC" in up:
        return "Técnico Profesional"
    if "MAESTRIA" in up:
        return "Maestría"
    if "ESPECIALIZACION" in up:
        return "Especialización"
    return txt.title()


def _normalizar_centro(centro_raw: str) -> str:
    txt = _limpiar_texto(centro_raw, "Sin Centro")
    up = (
        txt.upper()
        .replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
    )
    if "JOSE ACEVEDO" in up or "CRA 30" in up or "JAG" in up:
        return "JAG - José Acevedo y Gómez (Bogotá)"
    if "SOACHA" in up:
        return "UDR Soacha"
    if "ZIPAQUIRA" in up:
        return "CEAD Zipaquirá"
    if "FACATATIVA" in up:
        return "CEAD Facatativá"
    if "GIRARDOT" in up:
        return "CEAD Girardot"
    if "FUSAGASUGA" in up:
        return "CEAD Fusagasugá"
    if "GACHETA" in up:
        return "CEAD Gachetá"
    if "CHIPAQUE" in up:
        return "UDR Chipaque"
    return txt.title()


def leer_y_procesar_excel_homologaciones(
    fuente_excel: Any = None,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], str]:
    """
    Lee `Homologaciones 2026.xlsx`, une las filas de continuación, corrige las columnas invertidas
    (`Nombre` <-> `Identidad`) y retorna:
    1. `homologaciones_estudiantes`: Lista de diccionarios con las homologaciones individuales de estudiantes.
    2. `acuerdos_generales`: Lista de acuerdos normativos/generales de Consejo de Escuela (aplazamientos, tablas de homologación, etc.).
    3. `mensaje`: Resumen del procesamiento.
    """
    try:
        if fuente_excel is None:
            if not RUTA_EXCEL_HOMOLOGACIONES_DEFECTO.exists():
                return [], [], f"No se encontró el archivo `{RUTA_EXCEL_HOMOLOGACIONES_DEFECTO}`."
            df_raw = pd.read_excel(RUTA_EXCEL_HOMOLOGACIONES_DEFECTO, sheet_name=0, header=None)
        else:
            df_raw = pd.read_excel(fuente_excel, sheet_name=0, header=None)
    except Exception as exc:
        return [], [], f"Error al leer el archivo Excel de Homologaciones: {exc}"

    if df_raw.empty:
        return [], [], "El archivo Excel está vacío."

    # Localizar la fila de encabezados ('Consejo de escuela', 'Acuerdo', 'Nombre', 'Identidad')
    idx_header = 0
    for idx in range(min(15, len(df_raw))):
        vals_lower = [str(v).strip().lower() for v in df_raw.iloc[idx].values if not _es_vacio(v)]
        if any("consejo" in v for v in vals_lower) and any("acuerdo" in v for v in vals_lower):
            idx_header = idx
            break

    headers = [re.sub(r"\s+", " ", str(v)).strip() for v in df_raw.iloc[idx_header].values]
    df = df_raw.iloc[idx_header + 1 :].copy()
    df.columns = headers[: len(df.columns)]
    df = df.reset_index(drop=True)

    # Identificar nombres exactos de columnas en el DataFrame
    mapa_c = {c.lower(): c for c in df.columns}
    col_no = mapa_c.get("no.", "No.")
    col_consejo = next((c for k, c in mapa_c.items() if "consejo" in k), "Consejo de escuela")
    col_acuerdo = next((c for k, c in mapa_c.items() if "acuerdo" in k), "Acuerdo")
    col_nombre = next((c for k, c in mapa_c.items() if "nombre" in k), "Nombre")
    col_identidad = next((c for k, c in mapa_c.items() if "identidad" in k or "documento" in k), "Identidad")
    col_prog_cursar = next((c for k, c in mapa_c.items() if "programa a cursar" in k or "cursar" in k), "Programa a cursar")
    col_centro = next((c for k, c in mapa_c.items() if "centro" in k), "Centro")
    col_prog_origen = next((c for k, c in mapa_c.items() if "programa de origen" in k or "origen" in k), "Programa de origen")
    col_snies = next((c for k, c in mapa_c.items() if "snies" in k), "Snies")
    col_nivel = next((c for k, c in mapa_c.items() if "nivel" in k), "Nivel")
    col_univ = next((c for k, c in mapa_c.items() if "universidad" in k or "institucion" in k), "Universidad de Origen")

    # Paso 1: Unir filas de continuación (cuando Acuerdo y Consejo están vacíos pero continúa el texto de Programa de origen o Nombre)
    filas_consolidadas: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        acuerdo_val = _limpiar_texto(row.get(col_acuerdo))
        consejo_val = _limpiar_texto(row.get(col_consejo))
        no_val = _limpiar_texto(row.get(col_no))
        nombre_val = _limpiar_texto(row.get(col_nombre))
        ident_val = _limpiar_texto(row.get(col_identidad))
        prog_cur_val = _limpiar_texto(row.get(col_prog_cursar))
        centro_val = _limpiar_texto(row.get(col_centro))
        prog_ori_val = _limpiar_texto(row.get(col_prog_origen))
        snies_val = _limpiar_texto(row.get(col_snies))
        nivel_val = _limpiar_texto(row.get(col_nivel))
        univ_val = _limpiar_texto(row.get(col_univ))

        # Si toda la fila está vacía, ignorar
        if not any([acuerdo_val, consejo_val, no_val, nombre_val, ident_val, prog_cur_val, prog_ori_val, univ_val]):
            continue

        # Caso fila de continuación (sin Acuerdo ni Consejo ni Identidad)
        if not acuerdo_val and not consejo_val and not ident_val and filas_consolidadas:
            ult = filas_consolidadas[-1]
            if prog_ori_val:
                ult["programa_origen"] = f"{ult['programa_origen']} {prog_ori_val}".strip()
            if nombre_val and not _es_documento_identidad(nombre_val):
                ult["nombre_raw"] = f"{ult['nombre_raw']} {nombre_val}".strip()
            continue

        filas_consolidadas.append(
            {
                "no_raw": no_val,
                "consejo_raw": consejo_val,
                "acuerdo_raw": acuerdo_val,
                "nombre_raw": nombre_val,
                "identidad_raw": ident_val,
                "programa_cursar_raw": prog_cur_val,
                "centro_raw": centro_val,
                "programa_origen": prog_ori_val,
                "snies_origen": snies_val,
                "nivel_raw": nivel_val,
                "universidad_raw": univ_val,
            }
        )

    # Paso 2: Clasificar y normalizar cada registro consolidado
    homologaciones_estudiantes: List[Dict[str, Any]] = []
    acuerdos_generales: List[Dict[str, Any]] = []
    ultimo_consejo = "Consejo de Escuela N° 2 febrero 10 de 2026"

    for idx_f, item in enumerate(filas_consolidadas, start=1):
        consejo_txt = item["consejo_raw"] or ultimo_consejo
        if item["consejo_raw"]:
            ultimo_consejo = item["consejo_raw"]

        acuerdo_txt = item["acuerdo_raw"] or f"REGISTRO SIN ACUERDO #{idx_f}"
        nom_r = item["nombre_raw"]
        id_r = item["identidad_raw"]

        fecha_iso, mes_etiqueta, num_acuerdo = _extraer_fecha_y_numero_acuerdo(acuerdo_txt, consejo_txt)
        consejo_norm = _normalizar_consejo(consejo_txt)

        # Detectar si `Nombre` e `Identidad` vienen invertidos (ej. Nombre='1051478139', Identidad='Astrid Carolina Diaz Rojas')
        documento_est = ""
        nombre_est = ""

        if _es_documento_identidad(id_r) and not _es_documento_identidad(nom_r):
            documento_est = re.sub(r"[\.\,\s\-]", "", id_r)
            nombre_est = nom_r.title()
        elif _es_documento_identidad(nom_r) and not _es_documento_identidad(id_r):
            documento_est = re.sub(r"[\.\,\s\-]", "", nom_r)
            nombre_est = id_r.title()
        elif _es_documento_identidad(id_r):
            documento_est = re.sub(r"[\.\,\s\-]", "", id_r)
            nombre_est = nom_r.title()

        # Si no hay documento de identidad ni programa a cursar, es un Acuerdo General/Normativo de Consejo de Escuela
        if not documento_est and not item["programa_cursar_raw"]:
            acuerdos_generales.append(
                {
                    "consecutivo": idx_f,
                    "numero_acuerdo": num_acuerdo,
                    "consejo_escuela": consejo_norm,
                    "acuerdo": acuerdo_txt,
                    "fecha_acuerdo": fecha_iso or "2026-01-01",
                    "mes_acuerdo": mes_etiqueta,
                    "asunto": nom_r or "Acuerdo General de Consejo de Escuela",
                }
            )
            continue

        prog_base, prog_detalle = _normalizar_programa_cursar(item["programa_cursar_raw"])
        univ_norm, tipo_homolog = _normalizar_universidad_y_tipo(item["universidad_raw"])
        nivel_norm = _normalizar_nivel_origen(item["nivel_raw"])
        centro_norm = _normalizar_centro(item["centro_raw"])
        prog_origen_limpio = _limpiar_texto(item["programa_origen"], "Sin especificar").title()
        snies_limpio = _limpiar_texto(item["snies_origen"], "S/N")

        homologaciones_estudiantes.append(
            {
                "consecutivo": len(homologaciones_estudiantes) + 1,
                "numero_acuerdo": num_acuerdo,
                "periodo": "2026",
                "consejo_escuela": consejo_norm,
                "consejo_original": consejo_txt,
                "acuerdo": acuerdo_txt,
                "fecha_resolucion": fecha_iso,
                "mes_acuerdo": mes_etiqueta,
                "documento_estudiante": documento_est or "S/D",
                "nombre_estudiante": nombre_est or "Sin Nombre",
                "programa_destino": prog_base,
                "programa_destino_original": prog_detalle,
                "centro": centro_norm,
                "programa_origen": prog_origen_limpio,
                "snies_origen": snies_limpio,
                "nivel_origen": nivel_norm,
                "institucion_origen": univ_norm,
                "tipo": tipo_homolog,  # 'Interna' o 'Externa'
                "estado": "Aprobada",
            }
        )

    msg = (
        f"Se procesaron {len(homologaciones_estudiantes):,} homologaciones estudiantiles aprobadas "
        f"y {len(acuerdos_generales)} acuerdos generales de Consejo de Escuela.".replace(",", ".")
    )
    return homologaciones_estudiantes, acuerdos_generales, msg


def guardar_cache_json_homologaciones(
    homologaciones_estudiantes: List[Dict[str, Any]],
    acuerdos_generales: List[Dict[str, Any]],
) -> None:
    """Guarda un respaldo JSON en el proyecto para que Streamlit Cloud disponga siempre de todos los campos."""
    try:
        RUTA_JSON_HOMOLOGACIONES_FALLBACK.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "actualizado": date.today().isoformat(),
            "total_homologaciones": len(homologaciones_estudiantes),
            "total_acuerdos_generales": len(acuerdos_generales),
            "homologaciones": homologaciones_estudiantes,
            "acuerdos_generales": acuerdos_generales,
        }
        RUTA_JSON_HOMOLOGACIONES_FALLBACK.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass


def sincronizar_homologaciones_a_supabase(
    fuente_excel: Any = None,
    reemplazar_existentes: bool = True,
) -> Tuple[bool, int, int, str]:
    """
    Lee `Homologaciones 2026.xlsx`, actualiza el respaldo JSON enriquecido y sincroniza
    los registros con la tabla `public.homologaciones` en Supabase.
    Compatible tanto con el esquema base de `public.homologaciones` como con las columnas extendidas.
    """
    homologaciones_est, acuerdos_gen, msg_proc = leer_y_procesar_excel_homologaciones(fuente_excel)
    if not homologaciones_est:
        return False, 0, 0, msg_proc

    guardar_cache_json_homologaciones(homologaciones_est, acuerdos_gen)

    client = get_supabase_client()
    if client is None:
        return True, len(homologaciones_est), len(acuerdos_gen), f"{msg_proc} (Guardado en caché local)."

    try:
        if reemplazar_existentes:
            client.table("homologaciones").delete().eq("periodo", "2026").execute()

        # Empaquetar metadatos extra en soporte_url como JSON compacto si la tabla remota solo tiene las columnas base,
        # o insertar directamente las columnas base de `public.homologaciones` definidas en schema.sql
        registros_sql: List[Dict[str, Any]] = []
        for h in homologaciones_est:
            meta_extra = json.dumps(
                {
                    "no": h["consecutivo"],
                    "num_ac": h["numero_acuerdo"],
                    "consejo": h["consejo_escuela"],
                    "acuerdo": h["acuerdo"],
                    "mes": h["mes_acuerdo"],
                    "nombre": h["nombre_estudiante"],
                    "centro": h["centro"],
                    "prog_orig": h["programa_origen"],
                    "snies": h["snies_origen"],
                    "nivel": h["nivel_origen"],
                    "prog_det": h["programa_destino_original"],
                },
                ensure_ascii=False,
            )
            registros_sql.append(
                {
                    "periodo": "2026",
                    "tipo": h["tipo"],
                    "documento_estudiante": h["documento_estudiante"],
                    "programa_destino": h["programa_destino"],
                    "institucion_origen": h["institucion_origen"],
                    "creditos_homologados": None,
                    "estado": "Aprobada",
                    "fecha_resolucion": h["fecha_resolucion"],
                    "soporte_url": meta_extra,
                }
            )

        total_ins = 0
        for inicio in range(0, len(registros_sql), BATCH_SIZE_HOMOLOGACIONES):
            lote = registros_sql[inicio : inicio + BATCH_SIZE_HOMOLOGACIONES]
            client.table("homologaciones").insert(lote).execute()
            total_ins += len(lote)

        limpiar_cache_homologaciones()
        return (
            True,
            total_ins,
            len(acuerdos_gen),
            f"Sincronización exitosa en Supabase: {total_ins:,} acuerdos de homologación estudiantil "
            f"y {len(acuerdos_gen)} acuerdos normativos de Consejo de Escuela procesados.".replace(",", "."),
        )
    except Exception as exc:
        limpiar_cache_homologaciones()
        return (
            False,
            len(homologaciones_est),
            len(acuerdos_gen),
            f"Se actualizaron los datos locales ({len(homologaciones_est)} homologaciones), pero ocurrió un aviso en Supabase: {exc}",
        )


@st.cache_data(ttl=300, show_spinner=False)
def obtener_datos_homologaciones_2026() -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Obtiene `(df_homologaciones, df_acuerdos_generales)` desde:
    1. El archivo Excel local `datos_locales/Homologaciones 2026.xlsx` o caché JSON enriquecida.
    2. La tabla `public.homologaciones` de Supabase (desempaquetando los metadatos de Consejo, Acuerdo, Estudiante y Centro).
    """
    homologaciones: List[Dict[str, Any]] = []
    acuerdos_gen: List[Dict[str, Any]] = []

    # 1. Intentar leer desde caché JSON enriquecida si existe
    if RUTA_JSON_HOMOLOGACIONES_FALLBACK.exists():
        try:
            data_json = json.loads(RUTA_JSON_HOMOLOGACIONES_FALLBACK.read_text(encoding="utf-8"))
            homologaciones = data_json.get("homologaciones") or []
            acuerdos_gen = data_json.get("acuerdos_generales") or []
        except Exception:
            pass

    # 2. Si no hay JSON pero está el Excel local, procesarlo y generar el JSON
    if not homologaciones and RUTA_EXCEL_HOMOLOGACIONES_DEFECTO.exists():
        homologaciones, acuerdos_gen, _ = leer_y_procesar_excel_homologaciones(None)
        if homologaciones:
            guardar_cache_json_homologaciones(homologaciones, acuerdos_gen)

    # 3. Si aún no hay datos locales, reconstruir desde Supabase (`public.homologaciones`)
    if not homologaciones:
        client = get_supabase_client()
        if client is not None:
            try:
                res = (
                    client.table("homologaciones")
                    .select("*")
                    .eq("periodo", "2026")
                    .order("id", desc=False)
                    .limit(2000)
                    .execute()
                )
                for idx, r in enumerate(res.data or [], start=1):
                    meta = {}
                    raw_sop = r.get("soporte_url") or ""
                    if str(raw_sop).startswith("{"):
                        try:
                            meta = json.loads(raw_sop)
                        except Exception:
                            meta = {}
                    homologaciones.append(
                        {
                            "consecutivo": meta.get("no", idx),
                            "numero_acuerdo": meta.get("num_ac", idx),
                            "periodo": r.get("periodo", "2026"),
                            "consejo_escuela": meta.get("consejo", "Consejo de Escuela 2026"),
                            "acuerdo": meta.get("acuerdo", f"Acuerdo #{idx}"),
                            "fecha_resolucion": r.get("fecha_resolucion"),
                            "mes_acuerdo": meta.get("mes", "2026"),
                            "documento_estudiante": r.get("documento_estudiante", ""),
                            "nombre_estudiante": meta.get("nombre", "-"),
                            "programa_destino": r.get("programa_destino", ""),
                            "programa_destino_original": meta.get("prog_det", r.get("programa_destino", "")),
                            "centro": meta.get("centro", "-"),
                            "programa_origen": meta.get("prog_orig", "-"),
                            "snies_origen": meta.get("snies", "S/N"),
                            "nivel_origen": meta.get("nivel", "-"),
                            "institucion_origen": r.get("institucion_origen", ""),
                            "tipo": r.get("tipo", "Externa"),
                            "estado": r.get("estado", "Aprobada"),
                        }
                    )
            except Exception:
                pass

    df_hom = pd.DataFrame(homologaciones) if homologaciones else pd.DataFrame()
    df_ac = pd.DataFrame(acuerdos_gen) if acuerdos_gen else pd.DataFrame()
    return df_hom, df_ac


def calcular_estadisticas_homologaciones(
    df_hom: pd.DataFrame,
    df_ac: pd.DataFrame,
) -> Dict[str, Any]:
    """
    Calcula todos los indicadores gerenciales y distribuciones sobre el DataFrame filtrado de Homologaciones 2026.
    """
    if df_hom.empty:
        return {
            "total_homologaciones": 0,
            "estudiantes_unicos": 0,
            "total_internas": 0,
            "total_externas": 0,
            "total_sena": 0,
            "total_otras_ies": 0,
            "total_consejos": 0,
            "total_acuerdos_generales": len(df_ac),
            "por_programa_destino": {},
            "por_institucion_origen": {},
            "por_centro": {},
            "por_consejo": {},
            "por_mes": {},
            "por_nivel_origen": {},
            "por_tipo": {},
            "por_programa_origen": {},
        }

    total_hom = len(df_hom)
    est_unicos = int(df_hom["documento_estudiante"].nunique())
    internas = int((df_hom["tipo"] == "Interna").sum())
    externas = int((df_hom["tipo"] == "Externa").sum())
    sena_cnt = int(df_hom["institucion_origen"].str.contains("SENA", case=False, na=False).sum())
    otras_ies = max(externas - sena_cnt, 0)
    total_consejos = int(df_hom["consejo_escuela"].nunique())

    return {
        "total_homologaciones": total_hom,
        "estudiantes_unicos": est_unicos,
        "total_internas": internas,
        "total_externas": externas,
        "total_sena": sena_cnt,
        "total_otras_ies": otras_ies,
        "total_consejos": total_consejos,
        "total_acuerdos_generales": len(df_ac),
        "por_programa_destino": df_hom["programa_destino"].value_counts().to_dict(),
        "por_institucion_origen": df_hom["institucion_origen"].value_counts().to_dict(),
        "por_centro": df_hom["centro"].value_counts().to_dict(),
        "por_consejo": df_hom["consejo_escuela"].value_counts().to_dict(),
        "por_mes": df_hom["mes_acuerdo"].value_counts().sort_index().to_dict(),
        "por_nivel_origen": df_hom["nivel_origen"].value_counts().to_dict(),
        "por_tipo": {
            "Interna (UNAD)": internas,
            "Externa - SENA (Convenio)": sena_cnt,
            "Externa - Otras Universidades / IES": otras_ies,
        },
        "por_programa_origen": df_hom["programa_origen"].value_counts().to_dict(),
    }

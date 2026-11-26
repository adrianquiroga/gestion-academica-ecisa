"""
Servicio de lógica de negocio para el Módulo 4: Seguimiento de Homologaciones y Tiempos de Respuesta 2026 (ECISA - ZCBC).
Basado en el archivo oficial `Reporte_Tiempos_20261123_093339.csv` / `.xlsx` (Informe seguimiento tiempo homologacion
01/01/2026 al 23/11/2026):
- Soporta tanto el formato `.csv` (separado por `;`) como `.xlsx` con las 34 columnas del sistema de homologaciones UNAD.
- Repara automáticamente problemas de doble codificación UTF-8 / Latin-1 (ej. `GÃ“MEZ` -> `GÓMEZ`, `CARREÃ‘O` -> `CARREÑO`).
- Desglosa los 4 actores del flujo y sus tiempos individuales de respuesta (en días):
  1. Estudiante (`Identificacion`, `Nombre`)
  2. Líder Nacional de Homologación (`Identificacion.1`, `Nombre.1`, `Fecha Inicial`, `Fecha Final`, `Tiempo`)
  3. Tutor Encargado (`Identificacion.2`, `Nombre.2`, `Fecha Inicial.1`, `Fecha Final.1`, `Tiempo.1`)
  4. Líder Nacional de Programa (`Identificacion.3`, `Nombre.3`, `Fecha Inicial.2`, `Fecha Final.2`, `Tiempo.2`)
- Procesa los tiempos consolidados del trámite (`Fecha Aprobacion`, `Fecha de Envio`, `Tiempo Envio`, `Tiempo actores`,
  `Tiempo admon`, `Tiempo Total`, `dias`), además de `Idhomologacion`, `Escuela`, `Zona`, `Programa`, `Centro` y `Tipo`
  (`int` = Interna, `egr` = Egresado / Convenio, `ext` = Externa).
- Sincroniza los registros con la tabla `public.homologaciones` en Supabase y guarda respaldo en `src/data/homologaciones_2026_cache.json`.
"""

from datetime import date, datetime
import io
import json
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd
import streamlit as st
from src.db.client import get_supabase_client

RUTAS_CANDIDATAS_REPORTE_TIEMPOS = [
    Path("datos_locales/Reporte_Tiempos_20261123_093339.csv"),
    Path("datos_locales/Reporte_Tiempos_20261123_093339.xlsx"),
    Path("Reporte_Tiempos_20261123_093339.csv"),
    Path("Reporte_Tiempos_20261123_093339.xlsx"),
]

RUTA_JSON_HOMOLOGACIONES_FALLBACK = Path("src/data/homologaciones_2026_cache.json")
BATCH_SIZE_HOMOLOGACIONES = 500

MAPA_TIPOS_HOMOLOGACION = {
    "int": ("Interna (int)", "Interna"),
    "egr": ("Egresado / Convenio (egr)", "Interna"),
    "ext": ("Externa (ext)", "Externa"),
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


def detectar_archivo_reporte_tiempos_local() -> Optional[Path]:
    """Localiza el archivo `Reporte_Tiempos_20261123_093339` (`.csv` o `.xlsx`) en `datos_locales/` o en la raíz."""
    for p in RUTAS_CANDIDATAS_REPORTE_TIEMPOS:
        if p.exists():
            return p
    # Buscar cualquier archivo que empiece por Reporte_Tiempos en datos_locales/
    carpeta_datos = Path("datos_locales")
    if carpeta_datos.exists():
        encontrados = sorted(carpeta_datos.glob("Reporte_Tiempos*"), reverse=True)
        if encontrados:
            return encontrados[0]
    return None


def limpiar_cache_homologaciones() -> None:
    """Invalida la caché de consultas de homologaciones."""
    obtener_datos_homologaciones_2026.clear()


def _reparar_mojibake(texto: str) -> str:
    """Repara textos con doble codificación UTF-8 leído como Latin-1/CP1252 (ej. 'GÃ“MEZ' -> 'GÓMEZ', 'MUÃ‘OZ' -> 'MUÑOZ')."""
    if "Ã" not in texto and "Â" not in texto:
        return texto
    try:
        return texto.encode("cp1252").decode("utf-8")
    except Exception:
        try:
            return texto.encode("latin-1").decode("utf-8")
        except Exception:
            reemplazos = {
                "Ã“": "Ó",
                "Ã‘": "Ñ",
                "Ã\x81": "Á",
                "Ã‰": "É",
                "Ã\x8d": "Í",
                "Ãš": "Ú",
                "Ã¡": "á",
                "Ã©": "é",
                "Ã­": "í",
                "Ã³": "ó",
                "Ãº": "ú",
                "Ã±": "ñ",
            }
            for k, v in reemplazos.items():
                texto = texto.replace(k, v)
            return texto


def _limpiar_texto(val: Any, defecto: str = "") -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return defecto
    txt = _reparar_mojibake(str(val)).strip()
    txt = re.sub(r"\s+", " ", txt)
    if not txt or txt.lower() in {"nan", "none"}:
        return defecto
    if txt.endswith(".0") and txt[:-2].isdigit():
        return txt[:-2]
    return txt


def _limpiar_entero(val: Any, defecto: int = 0) -> int:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return defecto
    try:
        txt = str(val).strip().replace(",", ".")
        if not txt or txt.lower() in {"nan", "none", "-"}:
            return defecto
        return int(round(float(txt)))
    except Exception:
        return defecto


def _parsear_fecha_ddmmyyyy(val: Any) -> Tuple[Optional[str], str]:
    """
    Convierte fechas del reporte como `'10/02/2026 0:00'` o `'19/01/2026 0:00'`
    a formato ISO `'2026-02-10'` y etiqueta de mes `'02. Febrero'`.
    """
    txt = _limpiar_texto(val)
    if not txt:
        return None, "Sin fecha"
    solo_fecha = txt.split(" ")[0].strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            dt = datetime.strptime(solo_fecha, fmt).date()
            return dt.isoformat(), NOMBRES_MESES_ORDEN.get(dt.month, "Sin fecha")
        except Exception:
            continue
    try:
        dt_pd = pd.to_datetime(solo_fecha, dayfirst=True, errors="coerce")
        if not pd.isna(dt_pd):
            d = dt_pd.date()
            return d.isoformat(), NOMBRES_MESES_ORDEN.get(d.month, "Sin fecha")
    except Exception:
        pass
    return None, "Sin fecha"


def _normalizar_centro(centro_raw: str) -> str:
    txt = _limpiar_texto(centro_raw, "SIN CENTRO").upper()
    if "JOSE ACEVEDO" in txt or "CRA 30" in txt:
        return "(BOGOTÁ CRA 30) JOSÉ ACEVEDO Y GÓMEZ"
    if "ZIPAQUIRA" in txt or "ZIPAQUIRÁ" in txt:
        return "ZIPAQUIRÁ"
    if "FACATATIVA" in txt or "FACATATIVÁ" in txt:
        return "FACATATIVÁ"
    if "FUSAGASUGA" in txt or "FUSAGASUGÁ" in txt:
        return "FUSAGASUGÁ"
    if "GACHETA" in txt or "GACHETÁ" in txt:
        return "GACHETÁ"
    if "SOACHA" in txt:
        return "UDR SOACHA"
    if "GIRARDOT" in txt:
        return "GIRARDOT"
    if "CHIPAQUE" in txt:
        return "CHIPAQUE"
    return txt


def _normalizar_programa(prog_raw: str) -> str:
    txt = _limpiar_texto(prog_raw, "SIN PROGRAMA").upper()
    txt = txt.replace("RESOLUCIÃ“N", "RESOLUCIÓN").replace("RESOLUCION", "RESOLUCIÓN")
    return txt


def _cargar_dataframe_crudo(fuente_archivo: Any = None) -> Tuple[Optional[pd.DataFrame], str]:
    """
    Carga el DataFrame crudo (sin encabezados) desde `Reporte_Tiempos_20261123_093339.csv` o `.xlsx`
    (ya sea desde el disco local o desde un `UploadedFile` de Streamlit).
    """
    try:
        if fuente_archivo is None:
            ruta = detectar_archivo_reporte_tiempos_local()
            if ruta is None:
                return None, "No se encontró el archivo `Reporte_Tiempos_20261123_093339` en `datos_locales/`."
            if ruta.suffix.lower() == ".csv":
                contenido = ruta.read_bytes()
                try:
                    texto = contenido.decode("utf-8-sig")
                except UnicodeDecodeError:
                    texto = contenido.decode("latin-1")
                df_raw = pd.read_csv(io.StringIO(texto), sep=";", header=None)
            else:
                df_raw = pd.read_excel(ruta, sheet_name=0, header=None)
            return df_raw, ruta.name
        else:
            nombre_arch = getattr(fuente_archivo, "name", "Reporte_Tiempos.csv")
            if str(nombre_arch).lower().endswith(".csv"):
                contenido = fuente_archivo.getvalue()
                try:
                    texto = contenido.decode("utf-8-sig")
                except UnicodeDecodeError:
                    texto = contenido.decode("latin-1")
                df_raw = pd.read_csv(io.StringIO(texto), sep=";", header=None)
            else:
                try:
                    df_raw = pd.read_excel(fuente_archivo, sheet_name=0, header=None)
                except Exception:
                    contenido = fuente_archivo.getvalue()
                    texto = contenido.decode("utf-8-sig", errors="replace")
                    df_raw = pd.read_csv(io.StringIO(texto), sep=";", header=None)
            return df_raw, str(nombre_arch)
    except Exception as exc:
        return None, f"Error al abrir el archivo de Reporte de Tiempos: {exc}"


def leer_y_procesar_reporte_tiempos(
    fuente_archivo: Any = None,
) -> Tuple[List[Dict[str, Any]], Dict[str, str], str]:
    """
    Procesa `Reporte_Tiempos_20261123_093339` (CSV o Excel) y extrae:
    1. `registros`: Lista con los 112 trámites de homologación y todos sus tiempos por actor y etapa.
    2. `metadatos_reporte`: Diccionario con `fecha_inicial_reporte` (`1/01/2026`), `fecha_final_reporte` (`23/11/2026`) y `archivo_origen`.
    3. `mensaje`: Resumen de la lectura.
    """
    df_raw, info_origen = _cargar_dataframe_crudo(fuente_archivo)
    if df_raw is None or df_raw.empty:
        return [], {}, info_origen

    fecha_ini_rep = "01/01/2026"
    fecha_fin_rep = "23/11/2026"

    # Detectar fechas del encabezado (fila 1: Fecha inicial; 1/01/2026; Fecha final; 23/11/2026)
    for idx in range(min(3, len(df_raw))):
        fila_vals = [_limpiar_texto(v) for v in df_raw.iloc[idx].values]
        if "rol" in [v.lower() for v in fila_vals]:
            continue
        for j, val in enumerate(fila_vals):
            if val.lower() == "fecha inicial" and j + 1 < len(fila_vals) and fila_vals[j + 1]:
                fecha_ini_rep = fila_vals[j + 1]
            elif val.lower() == "fecha final" and j + 1 < len(fila_vals) and fila_vals[j + 1]:
                fecha_fin_rep = fila_vals[j + 1]

    # Localizar fila de encabezados de columnas (donde aparecen 'Rol', 'Identificacion', 'Idhomologacion', 'Programa')
    idx_header = 3
    for idx in range(min(10, len(df_raw))):
        vals_low = [_limpiar_texto(v).lower() for v in df_raw.iloc[idx].values]
        if "idhomologacion" in vals_low or ("rol" in vals_low and "programa" in vals_low):
            idx_header = idx
            break

    df_datos = df_raw.iloc[idx_header + 1 :].reset_index(drop=True)
    registros: List[Dict[str, Any]] = []

    for idx_row, row in df_datos.iterrows():
        vals = list(row.values)
        if len(vals) < 34:
            vals.extend([None] * (34 - len(vals)))

        doc_est = _limpiar_texto(vals[1])
        nom_est = _limpiar_texto(vals[2]).upper()
        id_hom = _limpiar_texto(vals[29])

        if not doc_est and not nom_est and not id_hom:
            continue

        # Actor 2: Líder Nacional de Homologación (cols 3..8)
        lider_hom_doc = _limpiar_texto(vals[4])
        lider_hom_nom = _limpiar_texto(vals[5]).upper()
        f_ini_lhom, _ = _parsear_fecha_ddmmyyyy(vals[6])
        f_fin_lhom, _ = _parsear_fecha_ddmmyyyy(vals[7])
        t_lider_hom = _limpiar_entero(vals[8], 0)

        # Actor 3: Tutor Encargado (cols 9..14)
        tutor_doc = _limpiar_texto(vals[10])
        tutor_nom = _limpiar_texto(vals[11]).upper()
        f_ini_tut, _ = _parsear_fecha_ddmmyyyy(vals[12])
        f_fin_tut, _ = _parsear_fecha_ddmmyyyy(vals[13])
        t_tutor = _limpiar_entero(vals[14], 0)

        # Actor 4: Líder Nacional de Programa (cols 15..20)
        lider_prog_doc = _limpiar_texto(vals[16])
        lider_prog_nom = _limpiar_texto(vals[17]).upper()
        f_ini_lprog, _ = _parsear_fecha_ddmmyyyy(vals[18])
        f_fin_lprog, _ = _parsear_fecha_ddmmyyyy(vals[19])
        t_lider_prog = _limpiar_entero(vals[20], 0)

        # Fechas y tiempos de aprobación / envío / consolidados (cols 21..26)
        f_aprobacion, mes_aprobacion = _parsear_fecha_ddmmyyyy(vals[21])
        f_envio, _ = _parsear_fecha_ddmmyyyy(vals[22])
        t_envio = _limpiar_entero(vals[23], 0)
        t_actores = _limpiar_entero(vals[24], 0)
        t_admon = _limpiar_entero(vals[25], 0)
        t_total = _limpiar_entero(vals[26], 0)

        # Metadatos del estudio de homologación (cols 27..33)
        escuela = _limpiar_texto(vals[27], "ESCUELA DE CIENCIAS DE LA SALUD")
        zona = _limpiar_texto(vals[28], "Centro Bogotá y Cundinamarca")
        programa = _normalizar_programa(vals[30])
        centro = _normalizar_centro(vals[31])
        tipo_cod = _limpiar_texto(vals[32], "int").lower()
        dias_calendario = _limpiar_entero(vals[33], t_total)

        tipo_etiqueta, tipo_sql = MAPA_TIPOS_HOMOLOGACION.get(
            tipo_cod,
            (f"Otro ({tipo_cod})", "Externa"),
        )

        registros.append(
            {
                "consecutivo": len(registros) + 1,
                "id_homologacion": id_hom or str(len(registros) + 1),
                "periodo": "2026",
                "documento_estudiante": doc_est,
                "nombre_estudiante": nom_est,
                "programa_destino": programa,
                "centro": centro,
                "zona": zona,
                "escuela": escuela,
                "tipo_codigo": tipo_cod,
                "tipo_homologacion": tipo_etiqueta,
                "tipo_sql": tipo_sql,
                "lider_homologacion_doc": lider_hom_doc,
                "lider_homologacion_nombre": lider_hom_nom,
                "fecha_ini_lider_hom": f_ini_lhom,
                "fecha_fin_lider_hom": f_fin_lhom,
                "tiempo_lider_hom": t_lider_hom,
                "tutor_doc": tutor_doc,
                "tutor_nombre": tutor_nom,
                "fecha_ini_tutor": f_ini_tut,
                "fecha_fin_tutor": f_fin_tut,
                "tiempo_tutor": t_tutor,
                "lider_programa_doc": lider_prog_doc,
                "lider_programa_nombre": lider_prog_nom,
                "fecha_ini_lider_prog": f_ini_lprog,
                "fecha_fin_lider_prog": f_fin_lprog,
                "tiempo_lider_prog": t_lider_prog,
                "fecha_aprobacion": f_aprobacion,
                "mes_aprobacion": mes_aprobacion,
                "fecha_envio": f_envio,
                "tiempo_envio": t_envio,
                "tiempo_actores": t_actores,
                "tiempo_admon": t_admon,
                "tiempo_total": t_total,
                "dias_calendario": dias_calendario,
                "estado": "Aprobada",
            }
        )

    metadatos = {
        "fecha_inicial_reporte": fecha_ini_rep,
        "fecha_final_reporte": fecha_fin_rep,
        "archivo_origen": info_origen,
        "fecha_sincronizacion": date.today().isoformat(),
    }
    msg = (
        f"Se procesaron {len(registros):,} estudios de homologación del reporte de tiempos "
        f"(corte {fecha_ini_rep} al {fecha_fin_rep}).".replace(",", ".")
    )
    return registros, metadatos, msg


def guardar_cache_json_homologaciones(
    registros: List[Dict[str, Any]],
    metadatos: Dict[str, str],
) -> None:
    """Guarda el respaldo JSON en `src/data/homologaciones_2026_cache.json` para disponibilidad inmediata en Streamlit Cloud."""
    try:
        RUTA_JSON_HOMOLOGACIONES_FALLBACK.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "metadatos": metadatos,
            "total_registros": len(registros),
            "registros": registros,
        }
        RUTA_JSON_HOMOLOGACIONES_FALLBACK.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass


def sincronizar_homologaciones_a_supabase(
    fuente_archivo: Any = None,
    reemplazar_existentes: bool = True,
) -> Tuple[bool, int, str]:
    """
    Lee `Reporte_Tiempos_20261123_093339` (`.csv` o `.xlsx`), actualiza la caché JSON del proyecto
    y sincroniza los registros con la tabla `public.homologaciones` en Supabase.
    """
    registros, metadatos, msg_proc = leer_y_procesar_reporte_tiempos(fuente_archivo)
    if not registros:
        return False, 0, msg_proc

    guardar_cache_json_homologaciones(registros, metadatos)

    client = get_supabase_client()
    if client is None:
        return True, len(registros), f"{msg_proc} (Guardado en caché local)."

    try:
        if reemplazar_existentes:
            client.table("homologaciones").delete().eq("periodo", "2026").execute()

        registros_sql: List[Dict[str, Any]] = []
        for r in registros:
            meta_json = json.dumps(
                {
                    "id_hom": r["id_homologacion"],
                    "nombre": r["nombre_estudiante"],
                    "centro": r["centro"],
                    "zona": r["zona"],
                    "tipo_cod": r["tipo_codigo"],
                    "tipo_hom": r["tipo_homologacion"],
                    "lider_hom": r["lider_homologacion_nombre"],
                    "f_ini_lh": r["fecha_ini_lider_hom"],
                    "f_fin_lh": r["fecha_fin_lider_hom"],
                    "t_lh": r["tiempo_lider_hom"],
                    "tutor": r["tutor_nombre"],
                    "tutor_doc": r["tutor_doc"],
                    "f_ini_tut": r["fecha_ini_tutor"],
                    "f_fin_tut": r["fecha_fin_tutor"],
                    "t_tut": r["tiempo_tutor"],
                    "lider_prog": r["lider_programa_nombre"],
                    "f_ini_lp": r["fecha_ini_lider_prog"],
                    "f_fin_lp": r["fecha_fin_lider_prog"],
                    "t_lp": r["tiempo_lider_prog"],
                    "f_envio": r["fecha_envio"],
                    "mes": r["mes_aprobacion"],
                    "t_env": r["tiempo_envio"],
                    "t_act": r["tiempo_actores"],
                    "t_adm": r["tiempo_admon"],
                    "t_tot": r["tiempo_total"],
                    "dias": r["dias_calendario"],
                },
                ensure_ascii=False,
            )
            registros_sql.append(
                {
                    "periodo": "2026",
                    "tipo": r["tipo_sql"],
                    "documento_estudiante": r["documento_estudiante"],
                    "programa_destino": r["programa_destino"],
                    "institucion_origen": f"{r['tipo_homologacion']} | Tutor: {r['tutor_nombre']}",
                    "creditos_homologados": r["tiempo_total"],
                    "estado": "Aprobada",
                    "fecha_resolucion": r["fecha_aprobacion"],
                    "soporte_url": meta_json,
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
            f"Sincronización exitosa en Supabase: {total_ins:,} estudios de homologación del Reporte de Tiempos "
            f"(01/01/2026 al 23/11/2026) guardados.".replace(",", "."),
        )
    except Exception as exc:
        limpiar_cache_homologaciones()
        return (
            False,
            len(registros),
            f"Datos locales actualizados ({len(registros)} registros), aviso en Supabase: {exc}",
        )


@st.cache_data(ttl=300, show_spinner=False)
def obtener_datos_homologaciones_2026() -> Tuple[pd.DataFrame, Dict[str, str]]:
    """
    Obtiene `(df_homologaciones, metadatos)` desde:
    1. El archivo local `Reporte_Tiempos_20261123_093339` o el respaldo `src/data/homologaciones_2026_cache.json`.
    2. La tabla `public.homologaciones` de Supabase.
    """
    registros: List[Dict[str, Any]] = []
    metadatos: Dict[str, str] = {
        "fecha_inicial_reporte": "01/01/2026",
        "fecha_final_reporte": "23/11/2026",
        "archivo_origen": "Reporte_Tiempos_20261123_093339.csv",
    }

    if RUTA_JSON_HOMOLOGACIONES_FALLBACK.exists():
        try:
            data_json = json.loads(RUTA_JSON_HOMOLOGACIONES_FALLBACK.read_text(encoding="utf-8"))
            candidatos = data_json.get("registros") or []
            if candidatos and "id_homologacion" in candidatos[0]:
                registros = candidatos
                metadatos = data_json.get("metadatos") or metadatos
        except Exception:
            pass

    if not registros and detectar_archivo_reporte_tiempos_local() is not None:
        registros, metadatos, _ = leer_y_procesar_reporte_tiempos(None)
        if registros:
            guardar_cache_json_homologaciones(registros, metadatos)

    if not registros:
        client = get_supabase_client()
        if client is not None:
            try:
                res = (
                    client.table("homologaciones")
                    .select("*")
                    .eq("periodo", "2026")
                    .order("id", desc=False)
                    .limit(1000)
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
                    registros.append(
                        {
                            "consecutivo": idx,
                            "id_homologacion": meta.get("id_hom", str(r.get("id", idx))),
                            "periodo": "2026",
                            "documento_estudiante": r.get("documento_estudiante", ""),
                            "nombre_estudiante": meta.get("nombre", "-"),
                            "programa_destino": r.get("programa_destino", ""),
                            "centro": meta.get("centro", "-"),
                            "zona": meta.get("zona", "Centro Bogotá y Cundinamarca"),
                            "escuela": "ESCUELA DE CIENCIAS DE LA SALUD",
                            "tipo_codigo": meta.get("tipo_cod", "int"),
                            "tipo_homologacion": meta.get("tipo_hom", r.get("tipo", "Interna")),
                            "tipo_sql": r.get("tipo", "Interna"),
                            "lider_homologacion_nombre": meta.get("lider_hom", "ADRIAN QUIROGA RODRIGUEZ"),
                            "fecha_ini_lider_hom": meta.get("f_ini_lh"),
                            "fecha_fin_lider_hom": meta.get("f_fin_lh"),
                            "tiempo_lider_hom": int(meta.get("t_lh", 0)),
                            "tutor_doc": meta.get("tutor_doc", ""),
                            "tutor_nombre": meta.get("tutor", "-"),
                            "fecha_ini_tutor": meta.get("f_ini_tut"),
                            "fecha_fin_tutor": meta.get("f_fin_tut"),
                            "tiempo_tutor": int(meta.get("t_tut", 0)),
                            "lider_programa_nombre": meta.get("lider_prog", "-"),
                            "fecha_ini_lider_prog": meta.get("f_ini_lp"),
                            "fecha_fin_lider_prog": meta.get("f_fin_lp"),
                            "tiempo_lider_prog": int(meta.get("t_lp", 0)),
                            "fecha_aprobacion": r.get("fecha_resolucion"),
                            "mes_aprobacion": meta.get("mes", "2026"),
                            "fecha_envio": meta.get("f_envio"),
                            "tiempo_envio": int(meta.get("t_env", 0)),
                            "tiempo_actores": int(meta.get("t_act", 0)),
                            "tiempo_admon": int(meta.get("t_adm", 0)),
                            "tiempo_total": int(meta.get("t_tot", r.get("creditos_homologados") or 0)),
                            "dias_calendario": int(meta.get("dias", 0)),
                            "estado": "Aprobada",
                        }
                    )
            except Exception:
                pass

    df = pd.DataFrame(registros) if registros else pd.DataFrame()
    return df, metadatos


def calcular_estadisticas_homologaciones(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Calcula los KPIs de volumen y tiempos promedio de respuesta (en días) por actor, tipo,
    programa, centro, tutor encargado, líder de programa y mes de aprobación.
    """
    if df.empty:
        return {
            "total_homologaciones": 0,
            "estudiantes_unicos": 0,
            "total_int": 0,
            "total_egr": 0,
            "total_ext": 0,
            "promedio_tiempo_total": 0.0,
            "promedio_dias_calendario": 0.0,
            "promedio_tiempo_actores": 0.0,
            "promedio_tiempo_admon": 0.0,
            "promedio_tiempo_envio": 0.0,
            "promedio_lider_hom": 0.0,
            "promedio_tutor": 0.0,
            "promedio_lider_prog": 0.0,
            "por_programa": {},
            "por_centro": {},
            "por_tipo": {},
            "por_mes": {},
            "df_tutores": pd.DataFrame(),
            "df_lideres_programa": pd.DataFrame(),
            "df_programas_tiempos": pd.DataFrame(),
            "df_centros_tiempos": pd.DataFrame(),
        }

    total_hom = len(df)
    est_unicos = int(df["documento_estudiante"].nunique())
    tot_int = int((df["tipo_codigo"] == "int").sum())
    tot_egr = int((df["tipo_codigo"] == "egr").sum())
    tot_ext = int((df["tipo_codigo"] == "ext").sum())

    # Desempeño por Tutor Encargado
    df_tut = (
        df.groupby("tutor_nombre", as_index=False)
        .agg(
            Homologaciones=("id_homologacion", "count"),
            Promedio_Dias_Tutor=("tiempo_tutor", "mean"),
            Max_Dias_Tutor=("tiempo_tutor", "max"),
            Promedio_Tiempo_Total=("tiempo_total", "mean"),
        )
        .sort_values(by="Homologaciones", ascending=False)
        .reset_index(drop=True)
    )
    df_tut["Promedio_Dias_Tutor"] = df_tut["Promedio_Dias_Tutor"].round(1)
    df_tut["Promedio_Tiempo_Total"] = df_tut["Promedio_Tiempo_Total"].round(1)

    # Desempeño por Líder Nacional de Programa
    df_lprog = (
        df.groupby("lider_programa_nombre", as_index=False)
        .agg(
            Homologaciones=("id_homologacion", "count"),
            Promedio_Dias_Lider_Prog=("tiempo_lider_prog", "mean"),
            Max_Dias_Lider_Prog=("tiempo_lider_prog", "max"),
            Promedio_Tiempo_Total=("tiempo_total", "mean"),
        )
        .sort_values(by="Homologaciones", ascending=False)
        .reset_index(drop=True)
    )
    df_lprog["Promedio_Dias_Lider_Prog"] = df_lprog["Promedio_Dias_Lider_Prog"].round(1)
    df_lprog["Promedio_Tiempo_Total"] = df_lprog["Promedio_Tiempo_Total"].round(1)

    # Resumen por Programa Académico con tiempos
    df_prog_t = (
        df.groupby("programa_destino", as_index=False)
        .agg(
            Homologaciones=("id_homologacion", "count"),
            Internas_int=("tipo_codigo", lambda s: int((s == "int").sum())),
            Egresados_egr=("tipo_codigo", lambda s: int((s == "egr").sum())),
            Externas_ext=("tipo_codigo", lambda s: int((s == "ext").sum())),
            Prom_Actores=("tiempo_actores", "mean"),
            Prom_Admon=("tiempo_admon", "mean"),
            Prom_Total=("tiempo_total", "mean"),
            Prom_Dias_Cal=("dias_calendario", "mean"),
        )
        .sort_values(by="Homologaciones", ascending=False)
        .reset_index(drop=True)
    )
    for c_num in ["Prom_Actores", "Prom_Admon", "Prom_Total", "Prom_Dias_Cal"]:
        df_prog_t[c_num] = df_prog_t[c_num].round(1)

    # Resumen por Centro con tiempos
    df_cen_t = (
        df.groupby("centro", as_index=False)
        .agg(
            Homologaciones=("id_homologacion", "count"),
            Internas_int=("tipo_codigo", lambda s: int((s == "int").sum())),
            Egresados_egr=("tipo_codigo", lambda s: int((s == "egr").sum())),
            Externas_ext=("tipo_codigo", lambda s: int((s == "ext").sum())),
            Prom_Actores=("tiempo_actores", "mean"),
            Prom_Admon=("tiempo_admon", "mean"),
            Prom_Total=("tiempo_total", "mean"),
        )
        .sort_values(by="Homologaciones", ascending=False)
        .reset_index(drop=True)
    )
    for c_num in ["Prom_Actores", "Prom_Admon", "Prom_Total"]:
        df_cen_t[c_num] = df_cen_t[c_num].round(1)

    return {
        "total_homologaciones": total_hom,
        "estudiantes_unicos": est_unicos,
        "total_int": tot_int,
        "total_egr": tot_egr,
        "total_ext": tot_ext,
        "promedio_tiempo_total": round(float(df["tiempo_total"].mean()), 1),
        "promedio_dias_calendario": round(float(df["dias_calendario"].mean()), 1),
        "promedio_tiempo_actores": round(float(df["tiempo_actores"].mean()), 1),
        "promedio_tiempo_admon": round(float(df["tiempo_admon"].mean()), 1),
        "promedio_tiempo_envio": round(float(df["tiempo_envio"].mean()), 1),
        "promedio_lider_hom": round(float(df["tiempo_lider_hom"].mean()), 1),
        "promedio_tutor": round(float(df["tiempo_tutor"].mean()), 1),
        "promedio_lider_prog": round(float(df["tiempo_lider_prog"].mean()), 1),
        "por_programa": df["programa_destino"].value_counts().to_dict(),
        "por_centro": df["centro"].value_counts().to_dict(),
        "por_tipo": df["tipo_homologacion"].value_counts().to_dict(),
        "por_mes": df["mes_aprobacion"].value_counts().sort_index().to_dict(),
        "df_tutores": df_tut,
        "df_lideres_programa": df_lprog,
        "df_programas_tiempos": df_prog_t,
        "df_centros_tiempos": df_cen_t,
    }

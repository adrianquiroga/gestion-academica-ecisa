"""
Servicio de integración con Google Drive Institucional a través del puente en Google Apps Script.
Valida formato (pdf, png, jpg) y peso máximo (5 MB), renombra estandarizadamente y envía en Base64.
"""

import base64
import re
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple

import requests
import streamlit as st

MAX_FILE_SIZE_MB = 5
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024
ALLOWED_EXTENSIONS = {"pdf", "png", "jpg", "jpeg"}

# IDs de las carpetas raíz en Google Drive
FOLDER_ID_SOPORTES_INGLES = "13FAIUxhKdZLR2PWSr0HBKQfvSdDNREoEEHm4FAPU8J6MwZbtXPgHIj2M6z8lqCVh8LdMSUUC"
FOLDER_ID_SOPORTES_APORTES = "1HSLyyW8LzqKmSLvlhpj6-fZHeTgkfBvP"
URL_CARPETA_DRIVE_APORTES = "https://drive.google.com/drive/folders/1HSLyyW8LzqKmSLvlhpj6-fZHeTgkfBvP?usp=drive_link"

MIME_TYPES = {
    "pdf": "application/pdf",
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
}


def verificar_configuracion_drive() -> Tuple[bool, str]:
    """
    Verifica si las variables `APPS_SCRIPT_URL` y `DRIVE_API_TOKEN` están definidas
    en `.streamlit/secrets.toml` y no contienen valores de ejemplo.
    """
    try:
        url = st.secrets.get("APPS_SCRIPT_URL", "")
        token = st.secrets.get("DRIVE_API_TOKEN", "")
    except Exception:
        return False, "No se pudo leer `.streamlit/secrets.toml`."

    if not url or not token:
        return False, "Faltan `APPS_SCRIPT_URL` o `DRIVE_API_TOKEN` en `.streamlit/secrets.toml`."

    if "AKfycb..." in url or not url.startswith("https://script.google.com/"):
        return False, "La URL `APPS_SCRIPT_URL` contiene el valor de ejemplo o no es un enlace válido de Google Apps Script."

    return True, "Puente de Google Drive configurado correctamente."


def _limpiar_identificador(texto: str) -> str:
    """Limpia un texto para que sea seguro en nombres de archivo."""
    limpio = re.sub(r"[^A-Za-z0-9_-]+", "_", texto.strip())
    return limpio.strip("_") or "doc"


def construir_nombre_archivo(documento: str, categoria: str, nombre_original: str) -> str:
    """
    Genera el nombre estandarizado: `documento_categoria_YYYYMMDD_HHMMSS.ext`
    """
    ext = Path(nombre_original).suffix.lower().lstrip(".")
    if ext == "jpeg":
        ext = "jpg"
    doc_limpio = _limpiar_identificador(documento)
    cat_limpia = _limpiar_identificador(categoria)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{doc_limpio}_{cat_limpia}_{timestamp}.{ext}"


def validar_archivo_soporte(uploaded_file) -> Tuple[bool, str]:
    """
    Valida que el archivo cargado en `st.file_uploader` cumpla con:
    - Extensión permitida: pdf, png, jpg/jpeg.
    - Peso máximo: 5 MB.
    """
    if uploaded_file is None:
        return False, "No se ha seleccionado ningún archivo."

    nombre = getattr(uploaded_file, "name", "")
    ext = Path(nombre).suffix.lower().lstrip(".")
    if ext not in ALLOWED_EXTENSIONS:
        return (
            False,
            f"Formato `.{ext}` no permitido. Solo se admiten archivos PDF, PNG o JPG.",
        )

    size_bytes = getattr(uploaded_file, "size", None)
    if size_bytes is None:
        raw_bytes = uploaded_file.getvalue()
        size_bytes = len(raw_bytes)

    if size_bytes == 0:
        return False, "El archivo seleccionado está vacío (0 bytes)."

    if size_bytes > MAX_FILE_SIZE_BYTES:
        peso_mb = size_bytes / (1024 * 1024)
        return (
            False,
            f"El archivo pesa {peso_mb:.2f} MB y supera el límite máximo permitido de {MAX_FILE_SIZE_MB} MB.",
        )

    return True, "Archivo válido."


def subir_soporte_drive(
    uploaded_file,
    documento: str,
    carpeta_categoria: str,
    etiqueta_archivo: Optional[str] = None,
    root_folder_id: Optional[str] = None,
) -> Tuple[bool, Optional[str], str]:
    """
    Valida, renombra y sube un archivo de soporte hacia Google Drive mediante el puente de Apps Script.

    Args:
        uploaded_file: Objeto retornado por `st.file_uploader`.
        documento: Número de documento del docente o estudiante.
        carpeta_categoria: Ruta de subcarpeta dentro de Drive (ej. '72142321 - ADRIAN QUIROGA/Investigación').
        etiqueta_archivo: Etiqueta corta opcional para el nombre del archivo.
        root_folder_id: ID opcional de la carpeta raíz en Google Drive donde se creará `carpeta_categoria`.

    Returns:
        (exito: bool, url_drive: Optional[str], mensaje: str)
    """
    es_valido, msg_val = validar_archivo_soporte(uploaded_file)
    if not es_valido:
        return False, None, msg_val

    ok_cfg, msg_cfg = verificar_configuracion_drive()
    if not ok_cfg:
        return (
            False,
            None,
            f"No se pudo subir el soporte a Drive: {msg_cfg}",
        )

    apps_script_url: str = st.secrets["APPS_SCRIPT_URL"].strip()
    drive_token: str = st.secrets["DRIVE_API_TOKEN"].strip()

    ext = Path(uploaded_file.name).suffix.lower().lstrip(".")
    mime_type = MIME_TYPES.get(ext, "application/octet-stream")
    categoria_corta = etiqueta_archivo or carpeta_categoria.split("/")[-1]
    nombre_final = construir_nombre_archivo(documento, categoria_corta, uploaded_file.name)

    try:
        file_bytes = uploaded_file.getvalue()
        contenido_b64 = base64.b64encode(file_bytes).decode("utf-8")

        payload = {
            "token": drive_token,
            "root_folder_id": root_folder_id or FOLDER_ID_SOPORTES_INGLES,
            "carpeta_categoria": carpeta_categoria,
            "nombre_archivo": nombre_final,
            "mime_type": mime_type,
            "contenido_base64": contenido_b64,
        }

        response = requests.post(
            apps_script_url,
            json=payload,
            timeout=45,
        )
        response.raise_for_status()

        try:
            data = response.json()
        except ValueError:
            return (
                False,
                None,
                "La respuesta del puente de Google Apps Script no es un JSON válido. "
                "Verifica que hayas publicado el script con acceso para 'Cualquier persona'.",
            )

        if data.get("ok") and data.get("url"):
            return True, str(data["url"]), f"Soporte `{nombre_final}` subido exitosamente a Google Drive."

        error_remoto = data.get("error", "Error desconocido en Google Apps Script.")
        return False, None, f"El puente de Drive rechazó la carga: {error_remoto}"

    except requests.exceptions.Timeout:
        return (
            False,
            None,
            "Tiempo de espera agotado al conectar con Google Apps Script. Intenta nuevamente con una conexión estable.",
        )
    except requests.exceptions.RequestException as req_err:
        return (
            False,
            None,
            f"Error de red al comunicarse con el puente de Google Drive: {req_err}",
        )
    except Exception as exc:
        return (
            False,
            None,
            f"Error inesperado al procesar el archivo de soporte: {exc}",
        )

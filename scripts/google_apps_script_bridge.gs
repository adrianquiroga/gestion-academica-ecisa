/**
 * ============================================================================
 * PUENTE DE CARGA DE SOPORTES A GOOGLE DRIVE INSTITUCIONAL (APPS SCRIPT)
 * ============================================================================
 *
 * INSTRUCCIONES RÁPIDAS DE CONFIGURACIÓN:
 * 1. Crea una carpeta raíz en tu Google Drive institucional (ej. "Soportes_Gestion_Academica").
 * 2. Abre esa carpeta en el navegador y copia el ID que aparece al final de la URL:
 *    https://drive.google.com/drive/folders/ESTE_ES_EL_ROOT_FOLDER_ID
 * 3. Reemplaza las constantes ROOT_FOLDER_ID y DRIVE_API_TOKEN abajo.
 * 4. Publica este script como "Aplicación web":
 *    - Ejecutar como: "Yo" (tu cuenta institucional).
 *    - Quién tiene acceso: "Cualquier persona" (la seguridad la valida DRIVE_API_TOKEN).
 */

const ROOT_FOLDER_ID = "13FAIUxhKdZLR2PWSr0HBKQfvSdDNREoEEHm4FAPU8J6MwZbtXPgHIj2M6z8lqCVh8LdMSUUC";
const APORTES_FOLDER_ID = "1HSLyyW8LzqKmSLvlhpj6-fZHeTgkfBvP";
const DRIVE_API_TOKEN = "MiTokenInstitucional_2026_Seguro";

/**
 * Endpoint GET de diagnóstico para verificar que el Web App está activo.
 */
function doGet(e) {
  return responderJSON({
    ok: true,
    mensaje: "Puente de Google Drive para Gestión Académica activo. Usa peticiones POST para subir soportes."
  });
}

/**
 * Endpoint POST principal que recibe archivos en Base64 desde Streamlit.
 * Payload JSON esperado:
 * {
 *   "token": "...",
 *   "carpeta_categoria": "Docentes/Ingles",
 *   "nombre_archivo": "10203040_Ingles_20261124_213000.pdf",
 *   "mime_type": "application/pdf",
 *   "contenido_base64": "JVBERi0xLjQK..."
 * }
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return responderJSON({ ok: false, error: "No se recibió cuerpo JSON en la petición POST." });
    }

    const payload = JSON.parse(e.postData.contents);

    // 1. Validar token de seguridad
    if (!payload.token || payload.token !== DRIVE_API_TOKEN) {
      return responderJSON({ ok: false, error: "Token de autorización inválido (DRIVE_API_TOKEN no coincide)." });
    }

    // 2. Validar parámetros obligatorios
    const carpetaCategoria = (payload.carpeta_categoria || "General").trim();
    const nombreArchivo = (payload.nombre_archivo || ("soporte_" + new Date().getTime() + ".pdf")).trim();
    const mimeType = (payload.mime_type || "application/octet-stream").trim();
    const contenidoBase64 = payload.contenido_base64;

    if (!contenidoBase64) {
      return responderJSON({ ok: false, error: "El campo 'contenido_base64' está vacío." });
    }

    // Permitir especificar un root_folder_id por petición (ej. Carpeta de Aportes vs Carpeta de Inglés)
    const folderIdBase = (payload.root_folder_id || ROOT_FOLDER_ID || "").trim();

    if (!folderIdBase || folderIdBase === "PEGA_AQUI_EL_ID_DE_TU_CARPETA_RAIZ_EN_DRIVE") {
      return responderJSON({ ok: false, error: "ROOT_FOLDER_ID no ha sido configurado en el script de Google Apps Script." });
    }

    // 3. Obtener carpeta raíz indicada y crear/navegar jerarquía de subcarpetas (ej. "72142321 - ADRIAN QUIROGA")
    const carpetaDestino = obtenerOCrearJerarquiaCarpetas(folderIdBase, carpetaCategoria);

    // 4. Decodificar Base64 y crear el Blob del archivo
    const bytes = Utilities.base64Decode(contenidoBase64);
    const blob = Utilities.newBlob(bytes, mimeType, nombreArchivo);

    // 5. Guardar el archivo en la subcarpeta correspondiente
    const archivo = carpetaDestino.createFile(blob);

    // Configurar permiso de lectura mediante enlace para facilitar la consulta en los reportes
    try {
      archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (permErr) {
      // En algunos dominios institucionales ANYONE_WITH_LINK está restringido al dominio:
      try {
        archivo.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (domErr) {
        // Si la política del dominio no permite cambiar permisos vía script, se conserva el permiso por defecto de la carpeta
      }
    }

    const webViewLink = archivo.getUrl();

    return responderJSON({
      ok: true,
      url: webViewLink,
      file_id: archivo.getId(),
      nombre_archivo: archivo.getName(),
      carpeta: carpetaCategoria
    });

  } catch (err) {
    return responderJSON({
      ok: false,
      error: "Excepción en Google Apps Script: " + err.toString()
    });
  }
}

/**
 * Navega o crea recursivamente las subcarpetas indicadas en rutaRelativa (separadas por '/')
 * dentro de la carpeta raíz con ID rootFolderId.
 */
function obtenerOCrearJerarquiaCarpetas(rootFolderId, rutaRelativa) {
  let carpetaActual = DriveApp.getFolderById(rootFolderId);

  if (!rutaRelativa) {
    return carpetaActual;
  }

  const segmentos = rutaRelativa
    .split("/")
    .map(function(s) { return s.trim(); })
    .filter(function(s) { return s.length > 0; });

  for (let i = 0; i < segmentos.length; i++) {
    const nombreSubcarpeta = segmentos[i];
    const iterador = carpetaActual.getFoldersByName(nombreSubcarpeta);
    if (iterador.hasNext()) {
      carpetaActual = iterador.next();
    } else {
      carpetaActual = carpetaActual.createFolder(nombreSubcarpeta);
    }
  }

  return carpetaActual;
}

/**
 * Construye una respuesta HTTP con Content-Type application/json.
 */
function responderJSON(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

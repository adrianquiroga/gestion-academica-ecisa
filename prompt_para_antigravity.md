Actúa como un desarrollador Senior en Python, Streamlit, PostgreSQL (Supabase) y Google Apps Script.

Lee detenidamente el archivo `ARQUITECTURA.md` que se encuentra en la raíz de este proyecto y sigue todas sus especificaciones técnicas, modelo de datos y reglas de rendimiento.

Tu objetivo es construir la **Fase 1 (MVP)** de la aplicación paso a paso, incluyendo la carga de soportes (PDF, PNG, JPG de máximo 5 MB) hacia Google Drive institucional mediante un puente en Google Apps Script. Ejecuta las siguientes tareas en orden:

1. **Entorno y Estructura:**
   * Crea la estructura de carpetas exacta definida en `ARQUITECTURA.md`.
   * Crea el archivo `requirements.txt` con las librerías necesarias (`streamlit`, `supabase`, `pandas`, `openpyxl`, `fpdf2`, `plotly`, `requests`).
   * Crea un `.gitignore` que proteja `.streamlit/secrets.toml`, `.venv/` y archivos temporales.
   * Crea un archivo de plantilla `.streamlit/secrets.toml.example` indicando las variables:
     * `SUPABASE_URL`
     * `SUPABASE_KEY`
     * `APPS_SCRIPT_URL`
     * `DRIVE_API_TOKEN`

2. **Base de Datos (`src/db/schema.sql` y `src/db/client.py`):**
   * Genera el script SQL completo en `src/db/schema.sql` con todas las tablas (Docentes, Estudios, Aportes, Matrículas y Homologaciones), incluyendo las columnas `soporte_ingles_url` y `soporte_url`, restricciones `CHECK`, llaves foráneas e índices optimizados para consultas rápidas sobre 13.000+ registros.
   * Implementa la conexión segura a Supabase en `src/db/client.py` usando `@st.cache_resource` y `st.secrets`.

3. **Puente de Soportes en Google Drive (`scripts/google_apps_script_bridge.gs` y `src/services/drive_svc.py`):**
   * Genera el archivo `scripts/google_apps_script_bridge.gs` listo para copiar y pegar en Google Apps Script. Debe recibir peticiones `doPost(e)` en JSON, validar `DRIVE_API_TOKEN`, decodificar el archivo en Base64, crear automáticamente subcarpetas dentro de `ROOT_FOLDER_ID` según `carpeta_categoria`, guardar el archivo y devolver `{ "ok": true, "url": "<webViewLink>" }`.
   * Implementa `src/services/drive_svc.py` en Python para validar formato (`pdf`, `png`, `jpg`) y tamaño máximo (5 MB), renombrar ordenadamente el archivo (`documento_categoria_YYYYMMDD_HHMMSS.ext`), enviarlo vía `requests.post` a `APPS_SCRIPT_URL` y retornar la URL resultante (manejando errores de red o configuración faltante de forma amigable).

4. **Servicios de Negocio y PDF (`src/services/`):**
   * Implementa `docentes_svc.py` con funciones limpias para buscar, crear y actualizar docentes (incluyendo `soporte_ingles_url`), así como agregar/eliminar sus estudios y aportes con su respectiva `soporte_url`.
   * Implementa `pdf_svc.py` usando `fpdf2` para generar dos tipos de PDF descargables en memoria (usando `BytesIO`, sin guardar archivos en disco):
     a) Ficha individual de un docente.
     b) Informe consolidado gerencial con tablas de resumen (niveles de inglés y aportes por categoría).

5. **Interfaz Multipágina (`app.py` y `pages/`):**
   * Construye `app.py` como portal de bienvenida con estado de conexión a Supabase y verificación de configuración del puente de Drive.
   * Construye `pages/1_Registro_Docente.py` con una interfaz intuitiva donde el docente pueda buscarse por correo o documento, actualizar su nivel de inglés (A1, A2, B1, B1+, B2, etc.), adjuntar soportes (`st.file_uploader` para PDF/PNG/JPG máx. 5 MB) y agregar múltiples estudios y aportes sin perder el estado de la pantalla.
   * Construye `pages/2_Panel_y_PDF.py` con métricas clave, filtros por nivel de inglés y categoría de aporte, enlaces directos para abrir cada soporte en Google Drive y botones de descarga de los informes en PDF.
   * Deja creados los esqueletos base de `pages/3_Matriculas.py` y `pages/4_Homologaciones.py` con un mensaje de "Módulo en construcción - Fase 2/3".

**Reglas de ejecución:**
* Crea el entorno virtual local e instala las dependencias de `requirements.txt`.
* Verifica que no haya errores de sintaxis en ningún archivo Python.
* Al terminar, indícame paso a paso:
  1. Cómo ejecutar `schema.sql` en Supabase.
  2. Cómo pegar y publicar `scripts/google_apps_script_bridge.gs` como Aplicación Web en mi Google Drive institucional (ejecutándose como "Yo").
  3. Qué valores pegar en `.streamlit/secrets.toml` y cómo lanzar la app localmente con `streamlit run app.py`.
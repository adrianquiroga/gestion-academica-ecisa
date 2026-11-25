# Arquitectura del Sistema de Gestión Académica (Docentes, Matrículas y Homologaciones)

## 1. Visión General y Stack Tecnológico
Aplicación web modular diseñada para operar en infraestructura gratuita, con capacidad de escalar a decenas de miles de registros académicos sin degradar el rendimiento y con almacenamiento de soportes en Google Drive institucional.

* **Frontend y Lógica de Aplicación:** Python 3.11+ con **Streamlit** (despliegue gratuito en Streamlit Community Cloud).
* **Base de Datos Relacional:** **Supabase** (PostgreSQL, capa gratuita de 500 MB) para datos estructurados y enlaces a soportes.
* **Almacenamiento de Soportes (PDF, PNG, JPG):** **Google Drive Institucional** integrado mediante un puente ligero en **Google Apps Script (Web App)** usando `requests` en Python (evita bloqueos de cuentas de servicio, cuotas de 0 GB y restricciones OAuth de Google Workspace).
* **Procesamiento de Datos Masivos:** `pandas` y `openpyxl` (para carga masiva de archivos Excel/CSV de matrículas y homologaciones).
* **Generación de Informes PDF:** `fpdf2` (ligero, puro en Python, sin dependencias de sistema operativo que fallen en Streamlit Cloud).
* **Visualización:** `plotly` o gráficas nativas de Streamlit.

---

## 2. Reglas de Rendimiento y Límites de Infraestructura
Streamlit Community Cloud asigna **1 GB de RAM**. Para manejar ~13.000 registros de matrícula por periodo y carga de archivos sin colapsar la memoria:
1. **Nunca cargar tablas completas en memoria** (`SELECT *` prohibido en `matriculas` y `homologaciones`).
2. **Filtros y agregaciones en SQL:** Usar consultas filtradas o funciones RPC/Vistas en PostgreSQL para traer solo los totales o páginas de máximo 100 filas.
3. **Carga masiva por lotes (Batching):** Al importar archivos Excel de 13.000 filas, insertar en Supabase en bloques de 500 a 1.000 registros por petición.
4. **Soportes fuera de SQL:** Nunca guardar archivos binarios en PostgreSQL. Validar tamaño máximo de **5 MB** por archivo en Streamlit (`pdf`, `png`, `jpg`), enviar en Base64 al puente de Apps Script y guardar únicamente el enlace (`webViewLink`) en Supabase.
5. **Caché:** Utilizar `@st.cache_data(ttl=300)` para catálogos y consultas de resumen.

---

## 3. Modelo de Datos (PostgreSQL / Supabase)

### Módulo 1: Gestión Docente (Fase 1 - Prioridad Inmediata)
* **`docentes`**:
  * `id` (UUID, PK, default `gen_random_uuid()`)
  * `documento` (TEXT, UNIQUE, NOT NULL)
  * `nombre_completo` (TEXT, NOT NULL)
  * `correo_institucional` (TEXT, UNIQUE, NOT NULL)
  * `nivel_ingles` (TEXT, CHECK IN `('A1', 'A2', 'B1', 'B1+', 'B2', 'C1', 'C2', 'Sin certificar')`)
  * `soporte_ingles_url` (TEXT, opcional - enlace de Drive al certificado de inglés)
  * `fecha_actualizacion` (TIMESTAMPTZ, default `now()`)

* **`docente_estudios`** (Relación 1 a N con `docentes`):
  * `id` (UUID, PK)
  * `docente_id` (UUID, FK -> `docentes.id` ON DELETE CASCADE)
  * `nivel_formacion` (TEXT, CHECK IN `('Pregrado', 'Especialización', 'Maestría', 'Doctorado', 'Posdoctorado')`)
  * `titulo_obtenido` (TEXT, NOT NULL)
  * `universidad` (TEXT, NOT NULL)
  * `anio_graduacion` (INT)
  * `soporte_url` (TEXT, opcional - enlace de Drive al diploma/acta)

* **`docente_aportes`** (Relación 1 a N con `docentes`):
  * `id` (UUID, PK)
  * `docente_id` (UUID, FK -> `docentes.id` ON DELETE CASCADE)
  * `categoria` (TEXT, CHECK IN `('Investigación', 'Innovación', 'Internacionalización', 'Vida Académica', 'Vida Universitaria')`)
  * `titulo_aporte` (TEXT, NOT NULL)
  * `descripcion` (TEXT)
  * `periodo_academico` (TEXT, ej. `'2026-1604'`)
  * `soporte_url` (TEXT, opcional - enlace de Drive a la evidencia)

### Módulo 2: Matrículas por Periodo (Fase 2 - ~13.000 registros/periodo)
* **`matriculas`**:
  * `id` (BIGSERIAL, PK)
  * `periodo` (TEXT, NOT NULL, indexado)
  * `documento_estudiante` (TEXT, NOT NULL, indexado)
  * `nombre_estudiante` (TEXT)
  * `programa_academico` (TEXT, indexado)
  * `centro_zona` (TEXT, indexado)
  * `tipo_matricula` (TEXT: Nuevo, Antiguo, Reingreso)
  * `creditos_matriculados` (INT)
  * *Índice compuesto recomendado:* `(periodo, programa_academico)` y `(periodo, documento_estudiante)`.

### Módulo 3: Homologaciones (Fase 3)
* **`homologaciones`**:
  * `id` (BIGSERIAL, PK)
  * `periodo` (TEXT, NOT NULL, indexado)
  * `tipo` (TEXT, CHECK IN `('Interna', 'Externa')`)
  * `documento_estudiante` (TEXT, NOT NULL, indexado)
  * `programa_destino` (TEXT, NOT NULL)
  * `institucion_origen` (TEXT)
  * `creditos_homologados` (INT)
  * `estado` (TEXT: En estudio, Aprobada, Rechazada)
  * `fecha_resolucion` (DATE)
  * `soporte_url` (TEXT, opcional - enlace de Drive a la resolución o sábana de notas)

---

## 4. Arquitectura del Puente de Soportes (Google Drive + Apps Script)
1. En Streamlit, el usuario adjunta un archivo (`pdf`, `png`, `jpg`, máx. 5 MB) mediante `st.file_uploader`.
2. `drive_svc.py` valida peso y extensión, codifica los bytes en **Base64** y envía un `POST` JSON a `APPS_SCRIPT_URL` incluyendo:
   * `token`: clave secreta (`DRIVE_API_TOKEN` en `secrets.toml`).
   * `carpeta_categoria`: ej. `Docentes/Ingles`, `Docentes/Estudios`, `Docentes/Aportes`, `Homologaciones`.
   * `nombre_archivo`: formato ordenado (ej. `documento_categoria_YYYYMMDD_HHMMSS.ext`).
   * `mime_type` y `contenido_base64`.
3. El script `google_apps_script_bridge.gs` (desplegado como Web App en el Drive institucional ejecutándose como *"Yo"*) valida el token, crea la subcarpeta si no existe dentro de la carpeta raíz (`ROOT_FOLDER_ID`), guarda el archivo y retorna un JSON con `{ "ok": true, "url": "<webViewLink>" }`.
4. Streamlit almacena esa URL en la columna correspondiente de Supabase.

---

## 5. Estructura de Carpetas del Proyecto

```text
gestion_academica_app/
├── .streamlit/
│   ├── config.toml                    # Tema visual (colores institucionales)
│   ├── secrets.toml                   # Credenciales Supabase y Drive (NO subir a Git)
│   └── secrets.toml.example           # Plantilla de variables de entorno
├── scripts/
│   └── google_apps_script_bridge.gs   # Código listo para pegar en Google Apps Script
├── src/
│   ├── db/
│   │   ├── client.py                  # Conexión singleton a Supabase
│   │   └── schema.sql                 # Script SQL listo para ejecutar en Supabase
│   ├── services/
│   │   ├── docentes_svc.py            # CRUD de docentes, estudios y aportes
│   │   ├── drive_svc.py               # Envío de soportes en Base64 al puente de Drive
│   │   ├── matriculas_svc.py          # Carga masiva por lotes y consultas agregadas
│   │   └── pdf_svc.py                 # Generación de informes PDF con fpdf2
│   └── utils/
│       └── auth.py                    # Control básico de acceso (Docente vs Admin)
├── pages/
│   ├── 1_Registro_Docente.py          # Formulario dinámico + carga de soportes
│   ├── 2_Panel_y_PDF.py               # Filtros, enlaces a soportes y descarga PDF
│   ├── 3_Matriculas.py                # Módulo preparado para carga de 13.000 registros
│   └── 4_Homologaciones.py            # Módulo preparado para homologaciones
├── app.py                             # Página principal (Inicio y navegación)
├── requirements.txt                   # Dependencias exactas
├── .gitignore                         # Excluir .venv, __pycache__, secrets.toml
└── ARQUITECTURA.md                    # Este documento
```

---

## 6. Requisitos Funcionales de la Fase 1 (MVP)
1. **Registro y Actualización Docente con Soportes:**
   * El docente busca su perfil por documento o correo institucional. Si no existe, lo crea; si existe, carga sus datos para actualizarlos.
   * Selección de nivel de inglés (A1, A2, B1, B1+, B2, C1, C2, Sin certificar) con opción de adjuntar el certificado (PDF/PNG/JPG).
   * Sección dinámica para agregar y eliminar múltiples títulos (Pregrado y Posgrados) indicando universidad y adjuntando soporte opcional.
   * Sección dinámica para registrar aportes clasificados en: Investigación, Innovación, Internacionalización, Vida Académica y Vida Universitaria, con soporte opcional.
2. **Consulta de Soportes y Generación de Informes PDF:**
   * En las tablas y fichas de consulta, mostrar enlaces directos para abrir cada soporte en Google Drive.
   * **Hoja de Vida / Ficha Individual:** PDF limpio con los datos, títulos, nivel de inglés y aportes de un docente específico.
   * **Informe Consolidado:** PDF con resumen estadístico (distribución de docentes por nivel de inglés, cantidad de posgrados y total de aportes por categoría) con encabezado y fecha de generación.
3. **Estructura preparada para Fases 2 y 3:**
   * Dejar creadas las tablas en `schema.sql` (incluyendo `soporte_url` en homologaciones) y las páginas base (`3_Matriculas.py` y `4_Homologaciones.py`) listas para implementar la carga masiva de Excel en la siguiente iteración.

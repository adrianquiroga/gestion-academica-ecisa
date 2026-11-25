-- =============================================================================
-- ESQUEMA DE BASE DE DATOS: SISTEMA DE GESTIÓN ACADÉMICA (SUPABASE / POSTGRESQL)
-- Adaptado y optimizado para la estructura real de "Resoluciones Docentes.xlsx" (UNAD)
-- Incluye: Módulo 1 (Docentes + Resoluciones, Estudios, Aportes), Módulo 2 (Matrículas), Módulo 3 (Homologaciones)
-- =============================================================================

CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- =============================================================================
-- MÓDULO 1: GESTIÓN DOCENTE Y RESOLUCIONES DE VINCULACIÓN (FASE 1 - MVP)
-- =============================================================================

-- 1. Tabla principal de Docentes (Perfil Maestro + Datos de Resolución / Vinculación)
CREATE TABLE IF NOT EXISTS public.docentes (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    -- Identificación y Contacto
    documento TEXT UNIQUE NOT NULL,
    nombre_completo TEXT NOT NULL,
    correo_institucional TEXT UNIQUE NOT NULL,
    email_personal TEXT,
    telefono TEXT,
    cumpleanos_mes_dia TEXT, -- Formato MM-DD (ej. '07-17') sin año

    -- Competencia en Inglés
    nivel_ingles TEXT NOT NULL DEFAULT 'Sin certificar'
        CHECK (nivel_ingles IN ('A1', 'A2', 'B1', 'B1+', 'B2', 'C1', 'C2', 'En caracterización / Prueba', 'Sin certificar')),
    detalle_ingles_original TEXT, -- Conserva el texto original del Excel (ej. 'B2 - ADVANCED', 'CARACTERIZACION UNAD...')
    soporte_ingles_url TEXT,

    -- Resumen de Formación reportado en Resolución
    nivel_formacion_maximo TEXT, -- Ej: DOCTORADO, MAGISTER, ESPECIALIZACION, PROFESIONAL, TECNOLOGIA
    formacion_profesional_texto TEXT,
    formacion_especializacion_texto TEXT,
    formacion_maestria_texto TEXT,
    formacion_doctorado_texto TEXT,

    -- Datos de Vinculación y Resolución Académica
    consecutivo_vinculacion INT,
    periodo_academico TEXT,
    estado_aceptacion TEXT,
    modalidad TEXT, -- Ej: TIEMPO COMPLETO, MEDIO TIEMPO, HORA CATEDRA
    horas_vinculacion INT,
    escuela_codigo INT,
    escuela_descripcion TEXT,
    centro_descripcion TEXT,
    zona_descripcion TEXT,
    programa_codigo INT,
    programa_descripcion TEXT,
    fecha_inicio DATE,
    fecha_termina DATE,
    resolucion_numero TEXT,
    resolucion_fecha DATE,
    salario_mensual NUMERIC(15, 2),
    total_vinculacion NUMERIC(15, 2),
    dofe TEXT,
    es_nuevo TEXT,
    ultima_vinculacion DATE,
    estado TEXT NOT NULL DEFAULT 'A'
        CHECK (estado IN ('A', 'R')), -- 'A' = Activo, 'R' = Retirado

    fecha_actualizacion TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Migración idempotente (por si la tabla ya había sido creada con la versión anterior)
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS email_personal TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS telefono TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS cumpleanos_mes_dia TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS detalle_ingles_original TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS nivel_formacion_maximo TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS formacion_profesional_texto TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS formacion_especializacion_texto TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS formacion_maestria_texto TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS formacion_doctorado_texto TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS consecutivo_vinculacion INT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS periodo_academico TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS estado_aceptacion TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS modalidad TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS horas_vinculacion INT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS escuela_codigo INT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS escuela_descripcion TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS centro_descripcion TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS zona_descripcion TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS programa_codigo INT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS programa_descripcion TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS fecha_inicio DATE;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS fecha_termina DATE;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS resolucion_numero TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS resolucion_fecha DATE;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS salario_mensual NUMERIC(15, 2);
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS total_vinculacion NUMERIC(15, 2);
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS dofe TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS es_nuevo TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS ultima_vinculacion DATE;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS password_hash TEXT;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS debe_cambiar_clave BOOLEAN NOT NULL DEFAULT true;
ALTER TABLE public.docentes ADD COLUMN IF NOT EXISTS estado TEXT NOT NULL DEFAULT 'A';
ALTER TABLE public.docentes DROP CONSTRAINT IF EXISTS docentes_estado_check;
ALTER TABLE public.docentes ADD CONSTRAINT docentes_estado_check CHECK (estado IN ('A', 'R'));
CREATE INDEX IF NOT EXISTS idx_docentes_estado ON public.docentes (estado);

-- Actualizar restricción CHECK de nivel_ingles para incluir 'En caracterización / Prueba'
ALTER TABLE public.docentes DROP CONSTRAINT IF EXISTS docentes_nivel_ingles_check;
ALTER TABLE public.docentes ADD CONSTRAINT docentes_nivel_ingles_check
    CHECK (nivel_ingles IN ('A1', 'A2', 'B1', 'B1+', 'B2', 'C1', 'C2', 'En caracterización / Prueba', 'Sin certificar'));

-- Índices para búsqueda y filtrado gerencial rápido
CREATE INDEX IF NOT EXISTS idx_docentes_documento ON public.docentes (documento);
CREATE INDEX IF NOT EXISTS idx_docentes_correo ON public.docentes (correo_institucional);
CREATE INDEX IF NOT EXISTS idx_docentes_nivel_ingles ON public.docentes (nivel_ingles);
CREATE INDEX IF NOT EXISTS idx_docentes_modalidad ON public.docentes (modalidad);
CREATE INDEX IF NOT EXISTS idx_docentes_centro ON public.docentes (centro_descripcion);
CREATE INDEX IF NOT EXISTS idx_docentes_programa ON public.docentes (programa_descripcion);
CREATE INDEX IF NOT EXISTS idx_docentes_nivel_formacion ON public.docentes (nivel_formacion_maximo);
CREATE INDEX IF NOT EXISTS idx_docentes_periodo ON public.docentes (periodo_academico);

-- Trigger para actualizar automáticamente fecha_actualizacion en docentes
CREATE OR REPLACE FUNCTION public.fn_actualizar_timestamp_docente()
RETURNS TRIGGER AS $$
BEGIN
    NEW.fecha_actualizacion = now();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_docentes_fecha_actualizacion ON public.docentes;
CREATE TRIGGER trg_docentes_fecha_actualizacion
BEFORE UPDATE ON public.docentes
FOR EACH ROW
EXECUTE FUNCTION public.fn_actualizar_timestamp_docente();


-- 2. Tabla de Estudios de Docentes (Relación 1 a N con docentes)
CREATE TABLE IF NOT EXISTS public.docente_estudios (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    docente_id UUID NOT NULL REFERENCES public.docentes(id) ON DELETE CASCADE,
    nivel_formacion TEXT NOT NULL
        CHECK (nivel_formacion IN ('Tecnología', 'Pregrado', 'Especialización', 'Maestría', 'Doctorado', 'Posdoctorado')),
    titulo_obtenido TEXT NOT NULL,
    universidad TEXT NOT NULL,
    anio_graduacion INT CHECK (anio_graduacion IS NULL OR (anio_graduacion >= 1950 AND anio_graduacion <= 2100)),
    soporte_url TEXT,
    fecha_registro TIMESTAMPTZ NOT NULL DEFAULT now()
);

ALTER TABLE public.docente_estudios DROP CONSTRAINT IF EXISTS docente_estudios_nivel_formacion_check;
ALTER TABLE public.docente_estudios ADD CONSTRAINT docente_estudios_nivel_formacion_check
    CHECK (nivel_formacion IN ('Tecnología', 'Pregrado', 'Especialización', 'Maestría', 'Doctorado', 'Posdoctorado'));

CREATE INDEX IF NOT EXISTS idx_docente_estudios_docente_id ON public.docente_estudios (docente_id);
CREATE INDEX IF NOT EXISTS idx_docente_estudios_nivel ON public.docente_estudios (nivel_formacion);


-- 3. Tabla de Aportes Académicos de Docentes (Relación 1 a N con docentes)
CREATE TABLE IF NOT EXISTS public.docente_aportes (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    docente_id UUID NOT NULL REFERENCES public.docentes(id) ON DELETE CASCADE,
    categoria TEXT NOT NULL
        CHECK (categoria IN ('Investigación', 'Innovación', 'Internacionalización', 'Vida Académica', 'Vida Universitaria')),
    titulo_aporte TEXT NOT NULL,
    descripcion TEXT,
    periodo_academico TEXT NOT NULL,
    soporte_url TEXT,
    fecha_registro TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_docente_aportes_docente_id ON public.docente_aportes (docente_id);
CREATE INDEX IF NOT EXISTS idx_docente_aportes_categoria ON public.docente_aportes (categoria);
CREATE INDEX IF NOT EXISTS idx_docente_aportes_periodo ON public.docente_aportes (periodo_academico);


-- =============================================================================
-- MÓDULO 2: MATRÍCULAS POR PERIODO (FASE 2 - ACTAS 2025 Y 2026 UNAD ~15.500 REGISTROS)
-- =============================================================================

CREATE TABLE IF NOT EXISTS public.matriculas (
    id BIGSERIAL PRIMARY KEY,
    anio_acta INT NOT NULL DEFAULT 2026,
    periodo TEXT NOT NULL,
    tipo_documento TEXT,
    documento_estudiante TEXT NOT NULL,
    nombre_estudiante TEXT,
    correo_institucional TEXT,
    escuela TEXT,
    programa_academico TEXT NOT NULL,
    zona TEXT,
    centro_zona TEXT,
    tipo_matricula TEXT, -- Condición UNAD: 'Antiguo', 'Nuevo', 'Continuidad', 'Aplazada', 'Reingreso'
    num_matricula TEXT,
    cursos_aprobados INT,
    cursos_matriculados INT,
    promedio NUMERIC(5, 2),
    resultado TEXT,
    fecha_ult_acceso TEXT,
    pais_residencia TEXT,
    departamento_residencia TEXT,
    ciudad_residencia TEXT,
    zona_residencia TEXT,
    estrato INT,
    sexo TEXT,
    necesidades_especiales TEXT,
    creditos_pregrado INT,
    creditos_disciplinares INT,
    creditos_postgrado INT,
    creditos_matriculados INT, -- Creditos Totales
    fecha_carga TIMESTAMPTZ NOT NULL DEFAULT now(),
    CONSTRAINT uq_matricula_periodo_doc_prog UNIQUE (periodo, documento_estudiante, programa_academico)
);

-- Migración idempotente por si la tabla matriculas ya fue creada con la versión inicial
ALTER TABLE public.matriculas DROP CONSTRAINT IF EXISTS matriculas_tipo_matricula_check;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS anio_acta INT NOT NULL DEFAULT 2026;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS tipo_documento TEXT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS correo_institucional TEXT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS escuela TEXT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS zona TEXT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS num_matricula TEXT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS cursos_aprobados INT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS cursos_matriculados INT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS promedio NUMERIC(5, 2);
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS resultado TEXT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS fecha_ult_acceso TEXT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS pais_residencia TEXT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS departamento_residencia TEXT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS ciudad_residencia TEXT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS zona_residencia TEXT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS estrato INT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS sexo TEXT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS necesidades_especiales TEXT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS creditos_pregrado INT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS creditos_disciplinares INT;
ALTER TABLE public.matriculas ADD COLUMN IF NOT EXISTS creditos_postgrado INT;

DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM pg_constraint WHERE conname = 'uq_matricula_periodo_doc_prog'
    ) THEN
        ALTER TABLE public.matriculas
            ADD CONSTRAINT uq_matricula_periodo_doc_prog UNIQUE (periodo, documento_estudiante, programa_academico);
    END IF;
END $$;

CREATE INDEX IF NOT EXISTS idx_matriculas_anio ON public.matriculas (anio_acta);
CREATE INDEX IF NOT EXISTS idx_matriculas_periodo ON public.matriculas (periodo);
CREATE INDEX IF NOT EXISTS idx_matriculas_documento ON public.matriculas (documento_estudiante);
CREATE INDEX IF NOT EXISTS idx_matriculas_programa ON public.matriculas (programa_academico);
CREATE INDEX IF NOT EXISTS idx_matriculas_centro_zona ON public.matriculas (centro_zona);
CREATE INDEX IF NOT EXISTS idx_matriculas_condicion ON public.matriculas (tipo_matricula);
CREATE INDEX IF NOT EXISTS idx_matriculas_periodo_programa ON public.matriculas (periodo, programa_academico);
CREATE INDEX IF NOT EXISTS idx_matriculas_periodo_documento ON public.matriculas (periodo, documento_estudiante);
CREATE INDEX IF NOT EXISTS idx_matriculas_anio_centro ON public.matriculas (anio_acta, centro_zona);


-- =============================================================================
-- MÓDULO 3: HOMOLOGACIONES (FASE 3)
-- =============================================================================

CREATE TABLE IF NOT EXISTS public.homologaciones (
    id BIGSERIAL PRIMARY KEY,
    periodo TEXT NOT NULL,
    tipo TEXT NOT NULL CHECK (tipo IN ('Interna', 'Externa')),
    documento_estudiante TEXT NOT NULL,
    programa_destino TEXT NOT NULL,
    institucion_origen TEXT,
    creditos_homologados INT CHECK (creditos_homologados IS NULL OR creditos_homologados >= 0),
    estado TEXT NOT NULL DEFAULT 'En estudio'
        CHECK (estado IN ('En estudio', 'Aprobada', 'Rechazada')),
    fecha_resolucion DATE,
    soporte_url TEXT,
    fecha_registro TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_homologaciones_periodo ON public.homologaciones (periodo);
CREATE INDEX IF NOT EXISTS idx_homologaciones_documento ON public.homologaciones (documento_estudiante);
CREATE INDEX IF NOT EXISTS idx_homologaciones_estado ON public.homologaciones (estado);
CREATE INDEX IF NOT EXISTS idx_homologaciones_periodo_programa ON public.homologaciones (periodo, programa_destino);


-- =============================================================================
-- MÓDULO 4: PROYECCIÓN VS EJECUCIÓN DE CRÉDITOS ECISA 2026 (data ECISA.xlsx)
-- Campos principales: codigo_real_snies, pro_descripcion, nombrecead, proy_zon, creditos
-- =============================================================================

CREATE TABLE IF NOT EXISTS public.proyeccion_creditos_ecisa (
    id BIGSERIAL PRIMARY KEY,
    anio INT NOT NULL DEFAULT 2026,
    codigo_real_snies TEXT NOT NULL,
    pro_descripcion TEXT NOT NULL,
    nombrecead TEXT NOT NULL,
    proy_zon NUMERIC(15, 2) NOT NULL DEFAULT 0,
    creditos NUMERIC(15, 2) NOT NULL DEFAULT 0,
    diferencia_creditos NUMERIC(15, 2) GENERATED ALWAYS AS (creditos - proy_zon) STORED,
    porcentaje_cumplimiento NUMERIC(8, 2) GENERATED ALWAYS AS (
        CASE WHEN proy_zon > 0 THEN ROUND((creditos / proy_zon) * 100, 2) ELSE 0 END
    ) STORED,
    fecha_actualizacion TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_creditos_ecisa_snies ON public.proyeccion_creditos_ecisa (codigo_real_snies);
CREATE INDEX IF NOT EXISTS idx_creditos_ecisa_programa ON public.proyeccion_creditos_ecisa (pro_descripcion);
CREATE INDEX IF NOT EXISTS idx_creditos_ecisa_cead ON public.proyeccion_creditos_ecisa (nombrecead);
CREATE INDEX IF NOT EXISTS idx_creditos_ecisa_prog_cead ON public.proyeccion_creditos_ecisa (pro_descripcion, nombrecead);

-- =============================================================================
-- PERMISOS PARA ROLES DE API DE SUPABASE (anon, authenticated, service_role)
-- =============================================================================
GRANT USAGE ON SCHEMA public TO anon, authenticated, service_role;
GRANT ALL ON ALL TABLES IN SCHEMA public TO anon, authenticated, service_role;
GRANT ALL ON ALL SEQUENCES IN SCHEMA public TO anon, authenticated, service_role;
ALTER DEFAULT PRIVILEGES IN SCHEMA public GRANT ALL ON TABLES TO anon, authenticated, service_role;
ALTER DEFAULT PRIVILEGES IN SCHEMA public GRANT ALL ON SEQUENCES TO anon, authenticated, service_role;


-- ============================================================
-- SCHEMA para Dashboard de Urgencias SSMC
-- Ejecuta este SQL en: Supabase → SQL Editor → New query
-- ============================================================

-- Tabla principal de registros diarios de urgencias
CREATE TABLE IF NOT EXISTS registros (
  id                      UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  created_at              TIMESTAMPTZ DEFAULT NOW(),
  fecha                   DATE NOT NULL,
  semana_epi              TEXT,
  establecimiento         TEXT NOT NULL,
  demanda_total           INTEGER,
  pacientes_atendidos     INTEGER,
  atenciones_respiratorias INTEGER,
  abandonos               INTEGER,
  derivaciones_hec        INTEGER,
  derivaciones_hcsba      INTEGER,
  derivaciones_huap       INTEGER,
  tiempo_espera           NUMERIC,
  tiene_refuerzo          BOOLEAN DEFAULT FALSE,
  tipo_refuerzo           TEXT,
  horas_refuerzo          NUMERIC,
  observaciones           TEXT
);

-- Tabla de retenciones de ambulancias
CREATE TABLE IF NOT EXISTS retenciones_ambulancias (
  id                UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  created_at        TIMESTAMPTZ DEFAULT NOW(),
  fecha             DATE NOT NULL,
  establecimiento   TEXT NOT NULL,
  hora_traslado     TIME,
  tiempo_retencion  NUMERIC
);

-- ── Índices para mejorar rendimiento ──────────────────────────
CREATE INDEX IF NOT EXISTS idx_registros_fecha ON registros(fecha DESC);
CREATE INDEX IF NOT EXISTS idx_registros_establecimiento ON registros(establecimiento);
CREATE INDEX IF NOT EXISTS idx_ambulancias_fecha ON retenciones_ambulancias(fecha DESC);

-- ── Row Level Security (RLS) ──────────────────────────────────
-- Por ahora se permite acceso público con anon key (ajusta según necesites auth)
ALTER TABLE registros ENABLE ROW LEVEL SECURITY;
ALTER TABLE retenciones_ambulancias ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Acceso público a registros"
  ON registros FOR ALL
  USING (true)
  WITH CHECK (true);

CREATE POLICY "Acceso público a ambulancias"
  ON retenciones_ambulancias FOR ALL
  USING (true)
  WITH CHECK (true);

-- ── Habilitar Realtime ────────────────────────────────────────
-- Esto permite que la app reciba cambios en tiempo real
ALTER PUBLICATION supabase_realtime ADD TABLE registros;
ALTER PUBLICATION supabase_realtime ADD TABLE retenciones_ambulancias;

-- ============================================================
-- ¡Listo! Tu base de datos está configurada.
-- ============================================================

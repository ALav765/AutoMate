-- Run once: psql supply_plan -f schema.sql

CREATE TABLE IF NOT EXISTS checks (
  id          SERIAL PRIMARY KEY,
  check_num   INTEGER NOT NULL UNIQUE,   -- matches results[N] / checklist_engine.py numbering
  label       TEXT NOT NULL,
  description TEXT NOT NULL,
  threshold   TEXT NOT NULL,
  fail_type   TEXT NOT NULL DEFAULT 'business_alert', -- 'business_alert' | 'calc_error'
  created_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS jobs (
  job_id        TEXT PRIMARY KEY,        -- matches the short uuid[:8] used in main.py
  kind          TEXT NOT NULL,           -- 'run' | 'add_check'
  status        TEXT NOT NULL DEFAULT 'queued', -- 'queued' | 'running' | 'done' | 'error'
  month_label   TEXT,
  output_path   TEXT,
  summary       JSONB,                   -- { passed, total, calc_errors, business_alerts, loops }
  events        JSONB NOT NULL DEFAULT '[]'::jsonb,  -- full event log, for SSE replay/history
  created_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Seed the 15 existing checks referenced in checklist_engine.py.
-- Edit labels/thresholds to match your actual checklist before running.
-- (check_num 1-7 = band checks, 8-12 = threshold checks, 13-15 = summary checks — adjust as needed)

-- On-demand audit runs and their results. Re-runnable.
--
-- Unrelated to prism365_audit, the access ledger. A run is one execution of one audit (inactive
-- accounts, MFA registration, ...): who started it, when, how it ended, and the rows it found,
-- kept so a result can be reopened, compared and exported later. The API keeps the most recent
-- runs per audit and deletes older ones.

CREATE TABLE IF NOT EXISTS prism365_audit_run (
  id uuid PRIMARY KEY,
  tenant_id uuid NOT NULL,
  audit_id text NOT NULL CHECK (length(audit_id) <= 80),
  audit_version integer NOT NULL CHECK (audit_version > 0),
  status text NOT NULL CHECK (status IN ('running','succeeded','failed')),
  complete boolean NOT NULL DEFAULT false,
  started_at timestamptz NOT NULL,
  finished_at timestamptz,
  started_by uuid NOT NULL,
  started_by_name text CHECK (length(started_by_name) <= 500),
  error_code text CHECK (length(error_code) <= 60),
  error_message text CHECK (length(error_message) <= 500),
  summary jsonb NOT NULL DEFAULT '[]'::jsonb,
  notes jsonb NOT NULL DEFAULT '[]'::jsonb,
  result_rows jsonb NOT NULL DEFAULT '[]'::jsonb,
  row_count integer NOT NULL DEFAULT 0 CHECK (row_count >= 0)
);
CREATE INDEX IF NOT EXISTS prism365_audit_run_recent
  ON prism365_audit_run (tenant_id, audit_id, started_at DESC);
SELECT prism365_tenant_table('prism365_audit_run');

-- The scope a run used (the account domains it was limited to; null for every domain) and the
-- domains it saw before the scope was applied, so the scope chooser always lists every domain
-- and a result always says what it covered. Added in place: core migrations re-run on every
-- start, so existing installs gain the columns.
ALTER TABLE prism365_audit_run ADD COLUMN IF NOT EXISTS run_scope jsonb;
ALTER TABLE prism365_audit_run ADD COLUMN IF NOT EXISTS domain_counts jsonb NOT NULL DEFAULT '[]'::jsonb;

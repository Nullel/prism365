-- Module runtime: the tenant-isolation helper, the shared tables every module uses, and the
-- migration ledger. Re-runnable. Requires the prism365_app role (created by the migrator).

-- Applies Prism365's one tenant policy to a table. Every tenant-scoped table, shared or
-- module-owned, goes through this function; the runtime verifier rejects any table in the
-- prism365_ or mod_ namespace that has not.
CREATE OR REPLACE FUNCTION prism365_tenant_table(target regclass) RETURNS void
LANGUAGE plpgsql AS $fn$
DECLARE
  column_type text;
  seq regclass;
BEGIN
  SELECT format_type(a.atttypid, a.atttypmod) INTO column_type
    FROM pg_attribute a
   WHERE a.attrelid = target AND a.attname = 'tenant_id' AND a.attnotnull AND NOT a.attisdropped;
  IF column_type IS DISTINCT FROM 'uuid' THEN
    RAISE EXCEPTION 'Table % needs a NOT NULL uuid tenant_id column.', target;
  END IF;
  EXECUTE format('ALTER TABLE %s ENABLE ROW LEVEL SECURITY', target);
  EXECUTE format('ALTER TABLE %s FORCE ROW LEVEL SECURITY', target);
  EXECUTE format('DROP POLICY IF EXISTS prism365_tenant_isolation ON %s', target);
  EXECUTE format(
    'CREATE POLICY prism365_tenant_isolation ON %s '
    'USING (tenant_id::text = current_setting(''prism365.tenant_id'', true)) '
    'WITH CHECK (tenant_id::text = current_setting(''prism365.tenant_id'', true))',
    target);
  EXECUTE format('GRANT SELECT, INSERT, UPDATE, DELETE ON %s TO prism365_app', target);
  FOR seq IN
    SELECT d.objid::regclass
      FROM pg_depend d JOIN pg_class s ON s.oid = d.objid
     WHERE d.refobjid = target AND d.classid = 'pg_class'::regclass AND s.relkind = 'S'
  LOOP
    EXECUTE format('GRANT USAGE, SELECT ON SEQUENCE %s TO prism365_app', seq);
  END LOOP;
END;
$fn$;
REVOKE ALL ON FUNCTION prism365_tenant_table(regclass) FROM PUBLIC;

-- Findings are the only module output the digest, the UI and the AI analyst read.
CREATE TABLE IF NOT EXISTS prism365_finding (
  id uuid PRIMARY KEY,
  tenant_id uuid NOT NULL,
  module text NOT NULL,
  rule_id text NOT NULL,
  rule_version integer NOT NULL CHECK (rule_version > 0),
  dedupe_key text NOT NULL,
  purpose text NOT NULL CHECK (purpose IN ('security','operations','compliance','adoption')),
  severity text NOT NULL CHECK (severity IN ('info','low','medium','high','critical')),
  score integer NOT NULL CHECK (score BETWEEN 0 AND 1000),
  subject_kind text NOT NULL CHECK (subject_kind IN ('tenant','service','resource','person')),
  subject_id text NOT NULL,
  subject_label text,
  title text NOT NULL,
  summary text NOT NULL,
  evidence jsonb NOT NULL,
  limitations jsonb NOT NULL,
  window_start timestamptz NOT NULL,
  window_end timestamptz NOT NULL,
  state text NOT NULL CHECK (state IN ('open','acknowledged','resolved')),
  occurrences integer NOT NULL CHECK (occurrences > 0),
  first_seen timestamptz NOT NULL,
  last_seen timestamptz NOT NULL,
  state_changed_at timestamptz,
  state_changed_by uuid,
  UNIQUE (tenant_id, module, rule_id, dedupe_key)
);
CREATE INDEX IF NOT EXISTS prism365_finding_recent
  ON prism365_finding (tenant_id, state, last_seen DESC);
SELECT prism365_tenant_table('prism365_finding');

CREATE TABLE IF NOT EXISTS prism365_cursor (
  tenant_id uuid NOT NULL,
  module text NOT NULL,
  collector text NOT NULL,
  value text NOT NULL CHECK (length(value) <= 2048),
  updated_at timestamptz NOT NULL,
  PRIMARY KEY (tenant_id, module, collector)
);
SELECT prism365_tenant_table('prism365_cursor');

CREATE TABLE IF NOT EXISTS prism365_module_run (
  id uuid PRIMARY KEY,
  tenant_id uuid NOT NULL,
  module text NOT NULL,
  collector text NOT NULL,
  credential_kind text NOT NULL CHECK (credential_kind IN ('delegated','application')),
  started_at timestamptz NOT NULL,
  finished_at timestamptz,
  status text NOT NULL CHECK (status IN ('running','succeeded','partial','failed','abandoned')),
  fetched integer NOT NULL DEFAULT 0 CHECK (fetched >= 0),
  inserted integer NOT NULL DEFAULT 0 CHECK (inserted >= 0),
  findings integer NOT NULL DEFAULT 0 CHECK (findings >= 0),
  error_category text
);
-- Records Graph returned that could not be read and were left out; added after first release.
ALTER TABLE prism365_module_run ADD COLUMN IF NOT EXISTS skipped integer NOT NULL DEFAULT 0 CHECK (skipped >= 0);
CREATE INDEX IF NOT EXISTS prism365_module_run_latest
  ON prism365_module_run (tenant_id, module, collector, started_at DESC);
SELECT prism365_tenant_table('prism365_module_run');

-- Schema metadata, not tenant data. The runtime role has no privilege on it.
CREATE TABLE IF NOT EXISTS prism365_schema_migration (
  scope text NOT NULL,
  migration_id text NOT NULL,
  checksum text NOT NULL,
  applied_at timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (scope, migration_id)
);
REVOKE ALL ON prism365_schema_migration FROM PUBLIC;

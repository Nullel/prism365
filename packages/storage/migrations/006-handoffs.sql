-- Handoffs to an action tool, and the receipts it sends back (docs/handoff.md). Re-runnable.
--
-- Prism365 never acts on a tenant. A handoff is a list it exported from an audit for a separate
-- tool to act on; a receipt is that tool's report of what it did. Receipts are kept as reported,
-- as claims of the tool, never as state Prism365 checked. Nothing here is ever updated. Rows
-- older than the retention period (HANDOFF_RETENTION_DAYS, 3 years by default) are deleted
-- whenever a new handoff or receipt is stored.

-- One row per export: what was exported, by whom and when, and whether the run behind it was
-- limited to some account domains or partial. The items themselves live only in the file the
-- administrator downloaded.
CREATE TABLE IF NOT EXISTS prism365_handoff (
  tenant_id uuid NOT NULL,
  id uuid NOT NULL,
  audit_id text NOT NULL CHECK (length(audit_id) <= 80),
  audit_title text NOT NULL CHECK (length(audit_title) BETWEEN 1 AND 500),
  run_id uuid NOT NULL,
  view_id text CHECK (length(view_id) <= 40),
  view_label text CHECK (length(view_label) <= 500),
  filtered boolean NOT NULL DEFAULT false,
  run_scope jsonb,
  complete boolean NOT NULL,
  items integer NOT NULL CHECK (items BETWEEN 0 AND 20000),
  generated_at timestamptz NOT NULL,
  generated_by uuid NOT NULL,
  generated_by_name text CHECK (length(generated_by_name) <= 500),
  PRIMARY KEY (tenant_id, id)
);
CREATE INDEX IF NOT EXISTS prism365_handoff_recent
  ON prism365_handoff (tenant_id, generated_at DESC);
SELECT prism365_tenant_table('prism365_handoff');

-- One row per imported receipt. The id is the tool's own, so the key includes the tenant: one
-- tenant's receipt id can never collide with, or reveal, another's. The digest is the SHA-256 of
-- the receipt as received, in canonical JSON, which makes a repeated import recognisable; its
-- version says how it was taken, so a later method never misreads an older digest. The counts
-- by status are kept here so the list never has to read every item.
CREATE TABLE IF NOT EXISTS prism365_receipt (
  tenant_id uuid NOT NULL,
  id uuid NOT NULL,
  handoff_id uuid,
  tool_name text NOT NULL CHECK (length(tool_name) BETWEEN 1 AND 80),
  tool_version text NOT NULL CHECK (length(tool_version) BETWEEN 1 AND 40),
  performed_at timestamptz NOT NULL,
  performed_by text NOT NULL CHECK (length(performed_by) BETWEEN 1 AND 200),
  items integer NOT NULL CHECK (items BETWEEN 1 AND 20000),
  done_count integer NOT NULL CHECK (done_count >= 0),
  failed_count integer NOT NULL CHECK (failed_count >= 0),
  skipped_count integer NOT NULL CHECK (skipped_count >= 0),
  imported_at timestamptz NOT NULL,
  imported_by uuid NOT NULL,
  digest text NOT NULL CHECK (digest ~ '^[0-9a-f]{64}$'),
  digest_version integer NOT NULL CHECK (digest_version > 0),
  CHECK (done_count + failed_count + skipped_count = items),
  PRIMARY KEY (tenant_id, id),
  FOREIGN KEY (tenant_id, handoff_id) REFERENCES prism365_handoff (tenant_id, id)
);
CREATE INDEX IF NOT EXISTS prism365_receipt_recent
  ON prism365_receipt (tenant_id, imported_at DESC);
SELECT prism365_tenant_table('prism365_receipt');

-- Each action the receipt reports, in the receipt's order. An item names the account by its id,
-- its principal name, or both; the person page matches on either, so both are indexed. The
-- principal name is kept as reported, and matched by a key the API normalises (NFKC, lower
-- case), never by the database's lower(), whose result depends on its locale.
CREATE TABLE IF NOT EXISTS prism365_receipt_item (
  tenant_id uuid NOT NULL,
  receipt_id uuid NOT NULL,
  position integer NOT NULL CHECK (position BETWEEN 0 AND 19999),
  user_id text CHECK (length(user_id) BETWEEN 1 AND 200),
  user_principal_name text CHECK (length(user_principal_name) BETWEEN 1 AND 500),
  user_principal_key text CHECK (length(user_principal_key) BETWEEN 1 AND 2000),
  action text NOT NULL CHECK (action ~ '^[a-z][a-z0-9.-]{1,79}$'),
  status text NOT NULL CHECK (status IN ('done','failed','skipped')),
  detail text NOT NULL CHECK (length(detail) <= 500),
  CHECK (user_id IS NOT NULL OR user_principal_name IS NOT NULL),
  CHECK ((user_principal_name IS NULL) = (user_principal_key IS NULL)),
  PRIMARY KEY (tenant_id, receipt_id, position),
  FOREIGN KEY (tenant_id, receipt_id) REFERENCES prism365_receipt (tenant_id, id)
);
CREATE INDEX IF NOT EXISTS prism365_receipt_item_user
  ON prism365_receipt_item (tenant_id, user_id);
CREATE INDEX IF NOT EXISTS prism365_receipt_item_principal
  ON prism365_receipt_item (tenant_id, user_principal_key);
SELECT prism365_tenant_table('prism365_receipt_item');

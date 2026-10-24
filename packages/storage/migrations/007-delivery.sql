-- Alert destinations and the delivery log (docs/alerts.md). Re-runnable.
--
-- Prism365 never sends through Microsoft. A destination is the operator's own mail relay
-- (configured by environment, never here) with a list of recipients, or a webhook. A webhook's
-- address carries its own token, so it is a secret: it is kept only inside `secret`, encrypted
-- with the state store's cipher together with the signing secret, and bound to its tenant and
-- id. `host` is the only part of it the page ever shows again. `key_id` names the key that
-- encrypted it, so a destination written under another key is reported as unreadable instead of
-- failing on every send.

CREATE TABLE IF NOT EXISTS prism365_destination (
  tenant_id uuid NOT NULL,
  id uuid NOT NULL,
  kind text NOT NULL CHECK (kind IN ('email','teams','slack','generic')),
  label text NOT NULL CHECK (length(label) BETWEEN 1 AND 80),
  host text CHECK (length(host) BETWEEN 1 AND 253),
  recipients jsonb,
  secret bytea CHECK (octet_length(secret) <= 8192),
  key_id text CHECK (key_id ~ '^[0-9a-f]{16}$'),
  include_details boolean NOT NULL DEFAULT false,
  created_at timestamptz NOT NULL,
  created_by uuid NOT NULL,
  -- A webhook has an encrypted secret and a host and no recipients; email has recipients only.
  CHECK (
    (kind = 'email' AND recipients IS NOT NULL AND secret IS NULL AND host IS NULL AND key_id IS NULL)
    OR (kind <> 'email' AND recipients IS NULL AND secret IS NOT NULL AND host IS NOT NULL
        AND key_id IS NOT NULL)
  ),
  PRIMARY KEY (tenant_id, id)
);
SELECT prism365_tenant_table('prism365_destination');

-- One row per delivery to one destination: what sent it (a rule, a test, a scheduled export or
-- the weekly summary), how many items it covered, and how it went. Never the message itself.
-- `subjects` holds short digests of the subjects an alert covered, so a rule's quiet period can
-- be checked without keeping names or addresses. The destination id is kept after the
-- destination is deleted, so the log still says which deliveries went where. Rows are deleted
-- after 90 days.
CREATE TABLE IF NOT EXISTS prism365_delivery (
  tenant_id uuid NOT NULL,
  id uuid NOT NULL,
  destination_id uuid NOT NULL,
  source text NOT NULL CHECK (source IN ('rule','test','export','summary')),
  rule_id uuid,
  subjects jsonb NOT NULL DEFAULT '[]'::jsonb,
  items integer NOT NULL CHECK (items BETWEEN 0 AND 100000),
  status text NOT NULL CHECK (status IN ('pending','sent','failed')),
  attempts integer NOT NULL CHECK (attempts BETWEEN 0 AND 3),
  last_error text CHECK (last_error ~ '^[a-z][a-z_]{1,39}$'),
  created_at timestamptz NOT NULL,
  updated_at timestamptz NOT NULL,
  CHECK ((source = 'rule') = (rule_id IS NOT NULL)),
  PRIMARY KEY (tenant_id, id)
);
CREATE INDEX IF NOT EXISTS prism365_delivery_recent
  ON prism365_delivery (tenant_id, created_at DESC);
CREATE INDEX IF NOT EXISTS prism365_delivery_rule
  ON prism365_delivery (tenant_id, rule_id, created_at DESC);
SELECT prism365_tenant_table('prism365_delivery');

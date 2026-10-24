-- Notification center ("attention") and operator preferences. Re-runnable.
--
-- An attention item is one thing an administrator may need to act on, reported by a connector:
-- a Microsoft service incident, a Message Center post with a deadline. Provider facts (status,
-- labels, deadline, active) and the administrator's own state (unread, review state, muted) are
-- separate columns, so Microsoft resolving an incident never marks it reviewed and reviewing it
-- never changes what Microsoft said.

CREATE TABLE IF NOT EXISTS prism365_attention (
  id uuid PRIMARY KEY,
  tenant_id uuid NOT NULL,
  module text NOT NULL,
  kind text NOT NULL,
  ref text NOT NULL CHECK (length(ref) <= 200),
  class text NOT NULL,
  title text NOT NULL,
  summary text NOT NULL,
  services jsonb NOT NULL,
  provider_status text,
  provider_labels jsonb NOT NULL,
  provider_deadline timestamptz,
  provider_updated_at timestamptz NOT NULL,
  provider_impact text,
  provider_workaround text,
  active boolean NOT NULL,
  local_labels jsonb NOT NULL,
  severity text NOT NULL CHECK (severity IN ('info','low','medium','high','critical')),
  link text,
  -- The provider fields that count as a meaningful change, as last seen.
  meaningful jsonb NOT NULL,
  revision integer NOT NULL CHECK (revision > 0),
  change_note text,
  first_seen timestamptz NOT NULL,
  last_seen timestamptz NOT NULL,
  changed_at timestamptz NOT NULL,
  unread boolean NOT NULL,
  review_state text NOT NULL CHECK (review_state IN ('new','reviewed','done')),
  review_changed_at timestamptz,
  review_changed_by uuid,
  muted boolean NOT NULL DEFAULT false,
  UNIQUE (tenant_id, module, kind, ref)
);
CREATE INDEX IF NOT EXISTS prism365_attention_recent
  ON prism365_attention (tenant_id, active, changed_at DESC);
SELECT prism365_tenant_table('prism365_attention');

CREATE TABLE IF NOT EXISTS prism365_preference (
  tenant_id uuid NOT NULL,
  key text NOT NULL CHECK (length(key) <= 100),
  value jsonb NOT NULL,
  updated_at timestamptz NOT NULL,
  updated_by uuid,
  PRIMARY KEY (tenant_id, key)
);
SELECT prism365_tenant_table('prism365_preference');

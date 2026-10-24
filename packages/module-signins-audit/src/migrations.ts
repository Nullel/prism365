import type { ModuleMigration } from "@prism365/modules";

// Normalized sign-in events and per-account baselines. Raw Graph payloads are never stored.
// `processed` makes detection restartable: an event inserted by a run that died before its
// rules ran is picked up by the next run instead of being skipped as "already seen".
// A baseline value is `trusted` once it counts as normal for the account. Values that took part
// in a flagged sign-in stay untrusted, linked to that finding, until the finding is resolved.

export const signinsAuditMigrations: readonly ModuleMigration[] = [
  {
    id: "001-events-and-baselines",
    sql: `CREATE TABLE IF NOT EXISTS mod_signins_audit_event (
  tenant_id uuid NOT NULL,
  id text NOT NULL,
  created_at timestamptz NOT NULL,
  user_id text NOT NULL,
  upn text,
  display_name text,
  app_id text,
  app_name text,
  resource_name text,
  client_app text,
  ip text,
  country text,
  state text,
  city text,
  lat double precision,
  lon double precision,
  device_id text,
  device_os text,
  device_browser text,
  is_compliant boolean,
  is_managed boolean,
  trust_type text,
  error_code integer NOT NULL,
  failure_reason text,
  ca_status text,
  risk_level text,
  risk_level_aggregated text,
  risk_state text,
  risk_event_types jsonb NOT NULL,
  is_interactive boolean,
  backfilled boolean NOT NULL,
  processed boolean NOT NULL,
  collected_at timestamptz NOT NULL,
  PRIMARY KEY (tenant_id, id)
);
CREATE INDEX IF NOT EXISTS mod_signins_audit_event_time ON mod_signins_audit_event (tenant_id, created_at);
CREATE INDEX IF NOT EXISTS mod_signins_audit_event_user ON mod_signins_audit_event (tenant_id, user_id, created_at);
CREATE INDEX IF NOT EXISTS mod_signins_audit_event_ip ON mod_signins_audit_event (tenant_id, ip, created_at);
CREATE INDEX IF NOT EXISTS mod_signins_audit_event_pending ON mod_signins_audit_event (tenant_id, processed, created_at);
SELECT prism365_tenant_table('mod_signins_audit_event');

CREATE TABLE IF NOT EXISTS mod_signins_audit_seen (
  tenant_id uuid NOT NULL,
  user_id text NOT NULL,
  kind text NOT NULL CHECK (kind IN ('ip','country','state','device')),
  value text NOT NULL,
  first_seen timestamptz NOT NULL,
  last_seen timestamptz NOT NULL,
  trusted boolean NOT NULL,
  pending_finding text,
  PRIMARY KEY (tenant_id, user_id, kind, value)
);
SELECT prism365_tenant_table('mod_signins_audit_seen');`,
  },
];

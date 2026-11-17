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
  // IPv6 locations from the operator's IP location provider (docs/modules/signins-audit.md §13).
  // `location_source` is "microsoft" or the provider's id, and null on rows stored before this
  // migration (Microsoft's) and, briefly, on an IPv6 row waiting for its lookup. `ms_*` keep what
  // Microsoft reported when the provider's answer replaced it. `ip_location` caches answers per
  // network so the same network is not sent again; it holds networks, never full addresses.
  {
    id: "002-ip-location",
    sql: `ALTER TABLE mod_signins_audit_event
  ADD COLUMN IF NOT EXISTS location_source text,
  ADD COLUMN IF NOT EXISTS location_lookup text,
  ADD COLUMN IF NOT EXISTS ms_country text,
  ADD COLUMN IF NOT EXISTS ms_state text,
  ADD COLUMN IF NOT EXISTS ms_city text;

CREATE TABLE IF NOT EXISTS mod_signins_audit_ip_location (
  tenant_id uuid NOT NULL,
  provider text NOT NULL,
  network text NOT NULL,
  found boolean NOT NULL,
  country text,
  state text,
  city text,
  lat double precision,
  lon double precision,
  accuracy_km double precision,
  looked_up_at timestamptz NOT NULL,
  PRIMARY KEY (tenant_id, provider, network)
);
SELECT prism365_tenant_table('mod_signins_audit_ip_location');`,
  },
  // Microsoft's coordinates, kept next to its place names when the provider's answer replaced
  // them, so a wrong answer can still be investigated. Its own migration: one applied is never
  // edited (the module database refuses a changed checksum).
  {
    id: "003-ms-coordinates",
    sql: `ALTER TABLE mod_signins_audit_event
  ADD COLUMN IF NOT EXISTS ms_lat double precision,
  ADD COLUMN IF NOT EXISTS ms_lon double precision;`,
  },
];

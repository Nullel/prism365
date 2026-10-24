import type { ModuleMigration } from "@prism365/modules";

// What Microsoft reported, as last read. Incidents and posts also become notification-center
// items through the attention store; these tables keep the provider's record for the panel and
// for later reports.

export const serviceHealthMigrations: readonly ModuleMigration[] = [
  {
    id: "001-services-and-issues",
    sql: `CREATE TABLE IF NOT EXISTS mod_service_health_service (
  tenant_id uuid NOT NULL,
  id text NOT NULL,
  name text NOT NULL,
  status text NOT NULL,
  collected_at timestamptz NOT NULL,
  PRIMARY KEY (tenant_id, id)
);
SELECT prism365_tenant_table('mod_service_health_service');

CREATE TABLE IF NOT EXISTS mod_service_health_issue (
  tenant_id uuid NOT NULL,
  id text NOT NULL,
  service text NOT NULL,
  feature text,
  feature_group text,
  title text NOT NULL,
  impact_description text,
  classification text NOT NULL,
  status text NOT NULL,
  origin text,
  is_resolved boolean NOT NULL,
  start_at timestamptz,
  end_at timestamptz,
  last_modified_at timestamptz NOT NULL,
  latest_post_at timestamptz,
  latest_post text,
  scope_of_impact text,
  workaround text,
  collected_at timestamptz NOT NULL,
  PRIMARY KEY (tenant_id, id)
);
CREATE INDEX IF NOT EXISTS mod_service_health_issue_open ON mod_service_health_issue (tenant_id, is_resolved, last_modified_at);
SELECT prism365_tenant_table('mod_service_health_issue');`,
  },
];

export const messageCenterMigrations: readonly ModuleMigration[] = [
  {
    id: "001-messages",
    sql: `CREATE TABLE IF NOT EXISTS mod_service_messages_post (
  tenant_id uuid NOT NULL,
  id text NOT NULL,
  title text NOT NULL,
  category text,
  severity text,
  tags jsonb NOT NULL,
  services jsonb NOT NULL,
  is_major_change boolean NOT NULL,
  action_required_by timestamptz,
  start_at timestamptz,
  end_at timestamptz,
  last_modified_at timestamptz NOT NULL,
  body_text text,
  collected_at timestamptz NOT NULL,
  PRIMARY KEY (tenant_id, id)
);
CREATE INDEX IF NOT EXISTS mod_service_messages_post_recent ON mod_service_messages_post (tenant_id, last_modified_at);
SELECT prism365_tenant_table('mod_service_messages_post');`,
  },
];

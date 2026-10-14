CREATE TABLE IF NOT EXISTS prism365_state (
  tenant_id uuid NOT NULL,
  kind text NOT NULL CHECK (kind IN ('oauth','session','aiGrant','investigation','automationPackage')),
  key text NOT NULL,
  encrypted bytea NOT NULL,
  expires_at timestamptz NOT NULL,
  PRIMARY KEY (tenant_id,kind,key)
);
CREATE INDEX IF NOT EXISTS prism365_state_expiry ON prism365_state(expires_at);
ALTER TABLE prism365_state ENABLE ROW LEVEL SECURITY;
ALTER TABLE prism365_state FORCE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS prism365_state_tenant ON prism365_state;
CREATE POLICY prism365_state_tenant ON prism365_state
  USING (tenant_id::text = current_setting('prism365.tenant_id',true))
  WITH CHECK (tenant_id::text = current_setting('prism365.tenant_id',true));

CREATE TABLE IF NOT EXISTS prism365_audit (
  id uuid PRIMARY KEY,
  tenant_id uuid NOT NULL,
  actor_id uuid NOT NULL,
  event jsonb NOT NULL,
  occurred_at timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS prism365_audit_tenant_time ON prism365_audit(tenant_id,occurred_at);
ALTER TABLE prism365_audit ENABLE ROW LEVEL SECURITY;
ALTER TABLE prism365_audit FORCE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS prism365_audit_tenant ON prism365_audit;
CREATE POLICY prism365_audit_tenant ON prism365_audit
  USING (tenant_id::text = current_setting('prism365.tenant_id',true))
  WITH CHECK (tenant_id::text = current_setting('prism365.tenant_id',true));
GRANT USAGE ON SCHEMA public TO prism365_app;
GRANT SELECT,INSERT,UPDATE,DELETE ON prism365_state TO prism365_app;
GRANT SELECT,INSERT ON prism365_audit TO prism365_app;

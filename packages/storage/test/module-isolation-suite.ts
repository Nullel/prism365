import { randomUUID } from "node:crypto";
import type { Receipt } from "@prism365/contracts";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
// The import route's store, run here so the race between imports of one id is tried on real
// PostgreSQL, where transactions truly overlap, as well as on PGlite.
import { insertReceipt } from "../../../apps/api/src/handoff-store.ts";
import { type MigrationScope, type ModuleDatabase, SqlGuardError } from "../src/index.ts";

// Shared by the PGlite run (module-db.test.ts) and the PostgreSQL run (postgres.test.ts), so
// both backends prove the same isolation properties with the same assertions.
export const tenantA = randomUUID();
export const tenantB = randomUUID();

export const probeScope: MigrationScope = {
  scope: "module.test.probe",
  migrations: [
    {
      id: "001-items",
      sql: `CREATE TABLE IF NOT EXISTS mod_test_probe_item (
              tenant_id uuid NOT NULL,
              id text NOT NULL,
              value text NOT NULL,
              created_at timestamptz NOT NULL,
              PRIMARY KEY (tenant_id, id)
            );
            SELECT prism365_tenant_table('mod_test_probe_item');`,
    },
  ],
};

export interface Fixture {
  database: ModuleDatabase;
  owner(sql: string): Promise<void>;
  close(): Promise<void>;
}

export function isolationSuite(name: string, open: () => Promise<Fixture>, enabled = true) {
  describe.skipIf(!enabled)(`${name}: module tables and tenant isolation`, () => {
    let fixture: Fixture;
    const db = () => fixture.database;

    beforeAll(async () => {
      fixture = await open();
    }, 60_000);
    afterAll(async () => {
      await fixture?.close();
    }, 30_000);

    it("verifies that every tenant table enforces row security", async () => {
      await expect(db().verify()).resolves.toBeUndefined();
    });

    it("keeps alert destinations and deliveries to their tenant, in their shape", async () => {
      const destination = randomUUID();
      const addWebhook = (tenant: string, owner: string, id = destination) =>
        db().tenant(tenant, (sql) =>
          sql.query(
            `INSERT INTO prism365_destination (tenant_id, id, kind, label, host, recipients, secret,
               key_id, include_details, created_at, created_by)
             VALUES ($1,$2,'generic','Synthetic hook','hooks.example',NULL,$3,'0123456789abcdef',
                     false,now(),$4)`,
            [owner, id, Buffer.from("synthetic-ciphertext"), randomUUID()],
          ),
        );
      await expect(addWebhook(tenantA, tenantB)).rejects.toMatchObject({ code: "42501" });
      await addWebhook(tenantA, tenantA);
      await addWebhook(tenantB, tenantB);
      const seen = await db().tenant(tenantA, (sql) =>
        sql.query<{ tenant_id: string }>("SELECT tenant_id FROM prism365_destination"),
      );
      expect(seen.rows.map((row) => row.tenant_id)).toEqual([tenantA]);
      // A webhook needs its encrypted secret; email has recipients and no secret.
      await expect(
        db().tenant(tenantA, (sql) =>
          sql.query(
            `INSERT INTO prism365_destination (tenant_id, id, kind, label, host, recipients, secret,
               key_id, include_details, created_at, created_by)
             VALUES ($1,$2,'teams','No secret','hooks.example',NULL,NULL,NULL,false,now(),$3)`,
            [tenantA, randomUUID(), randomUUID()],
          ),
        ),
      ).rejects.toMatchObject({ code: "23514" });
      await expect(
        db().tenant(tenantA, (sql) =>
          sql.query(
            `INSERT INTO prism365_destination (tenant_id, id, kind, label, host, recipients, secret,
               key_id, include_details, created_at, created_by)
             VALUES ($1,$2,'email','Mail',NULL,'["ops@example.invalid"]'::jsonb,$3,NULL,false,
                     now(),$4)`,
            [tenantA, randomUUID(), Buffer.from("x"), randomUUID()],
          ),
        ),
      ).rejects.toMatchObject({ code: "23514" });
      const addDelivery = (tenant: string, owner: string, fields: { error?: string } = {}) =>
        db().tenant(tenant, (sql) =>
          sql.query(
            `INSERT INTO prism365_delivery (tenant_id, id, destination_id, source, rule_id, subjects,
               items, status, attempts, last_error, created_at, updated_at)
             VALUES ($1,$2,$3,'test',NULL,'[]'::jsonb,0,'failed',1,$4,now(),now())`,
            [owner, randomUUID(), destination, fields.error ?? "timeout"],
          ),
        );
      await expect(addDelivery(tenantA, tenantB)).rejects.toMatchObject({ code: "42501" });
      await addDelivery(tenantA, tenantA);
      await addDelivery(tenantB, tenantB);
      const deliveries = await db().tenant(tenantB, (sql) =>
        sql.query<{ tenant_id: string }>("SELECT tenant_id FROM prism365_delivery"),
      );
      expect(deliveries.rows.map((row) => row.tenant_id)).toEqual([tenantB]);
      // Only a fixed category is ever stored as the error, never a message.
      await expect(
        addDelivery(tenantA, tenantA, { error: "connect ECONNREFUSED 10.0.0.1:443" }),
      ).rejects.toMatchObject({ code: "23514" });
    });

    it("hides other tenants' rows even when a query has no tenant filter", async () => {
      const now = new Date().toISOString();
      await db().tenant(tenantA, (sql) =>
        sql.query(
          "INSERT INTO mod_test_probe_item (tenant_id,id,value,created_at) VALUES ($1,'x','synthetic-a',$2)",
          [tenantA, now],
        ),
      );
      await db().tenant(tenantB, (sql) =>
        sql.query(
          "INSERT INTO mod_test_probe_item (tenant_id,id,value,created_at) VALUES ($1,'x','synthetic-b',$2)",
          [tenantB, now],
        ),
      );
      const seenByA = await db().tenant(tenantA, (sql) =>
        sql.query<{ value: string }>("SELECT value FROM mod_test_probe_item"),
      );
      expect(seenByA.rows.map((row) => row.value)).toEqual(["synthetic-a"]);
      const moved = await db().tenant(tenantA, (sql) =>
        sql.query("UPDATE mod_test_probe_item SET value='moved' WHERE tenant_id=$1", [tenantB]),
      );
      expect(moved.rowCount).toBe(0);
      const deleted = await db().tenant(tenantA, (sql) =>
        sql.query("DELETE FROM mod_test_probe_item WHERE tenant_id=$1", [tenantB]),
      );
      expect(deleted.rowCount).toBe(0);
      const seenByB = await db().tenant(tenantB, (sql) =>
        sql.query<{ value: string }>("SELECT value FROM mod_test_probe_item"),
      );
      expect(seenByB.rows.map((row) => row.value)).toEqual(["synthetic-b"]);
    });

    it("rejects writing a row into another tenant or moving a row across tenants", async () => {
      await expect(
        db().tenant(tenantA, (sql) =>
          sql.query(
            "INSERT INTO mod_test_probe_item (tenant_id,id,value,created_at) VALUES ($1,'y','cross',now())",
            [tenantB],
          ),
        ),
      ).rejects.toMatchObject({ code: "42501" });
      await expect(
        db().tenant(tenantA, (sql) =>
          sql.query("UPDATE mod_test_probe_item SET tenant_id=$1", [tenantB]),
        ),
      ).rejects.toMatchObject({ code: "42501" });
    });

    it("applies the same isolation to shared finding, cursor, and run tables", async () => {
      await db().tenant(tenantA, (sql) =>
        sql.query(
          "INSERT INTO prism365_cursor (tenant_id,module,collector,value,updated_at) VALUES ($1,'test.probe','poll','synthetic-cursor',now())",
          [tenantA],
        ),
      );
      const fromB = await db().tenant(tenantB, (sql) =>
        sql.query("SELECT value FROM prism365_cursor"),
      );
      expect(fromB.rows).toEqual([]);
      await expect(
        db().tenant(tenantB, (sql) =>
          sql.query(
            "INSERT INTO prism365_module_run (id,tenant_id,module,collector,credential_kind,started_at,status) VALUES ($1,$2,'m','c','delegated',now(),'running')",
            [randomUUID(), tenantA],
          ),
        ),
      ).rejects.toMatchObject({ code: "42501" });
    });

    it("isolates handoffs and receipts, whose ids come from another tool", async () => {
      // One receipt id in both tenants: the key includes the tenant, so neither blocks the other.
      const receipt = randomUUID();
      const addReceipt = (tenant: string) =>
        db().tenant(tenant, (sql) =>
          sql.query(
            `INSERT INTO prism365_receipt (tenant_id, id, handoff_id, tool_name, tool_version,
               performed_at, performed_by, items, done_count, failed_count, skipped_count,
               imported_at, imported_by, digest, digest_version)
             VALUES ($1,$2,$3,'Example tool','1',now(),'Operator',1,1,0,0,now(),$1,$4,1)`,
            [tenant, receipt, null, "a".repeat(64)],
          ),
        );
      for (const [tenant, detail] of [
        [tenantA, "synthetic-a"],
        [tenantB, "synthetic-b"],
      ] as const) {
        await addReceipt(tenant);
        await db().tenant(tenant, (sql) =>
          sql.query(
            `INSERT INTO prism365_receipt_item (tenant_id, receipt_id, position, user_id,
               user_principal_name, user_principal_key, action, status, detail)
             VALUES ($1,$2,0,NULL,'Shared@Example.invalid','shared@example.invalid',
                     'account.disabled','done',$3)`,
            [tenant, receipt, detail],
          ),
        );
      }
      const seen = await db().tenant(tenantA, (sql) =>
        sql.query<{ detail: string }>(
          "SELECT detail FROM prism365_receipt_item WHERE user_principal_key = $1",
          ["shared@example.invalid"],
        ),
      );
      expect(seen.rows.map((row) => row.detail)).toEqual(["synthetic-a"]);
      // A handoff of tenant B can be neither written from A nor answered by A's receipts.
      const handoff = randomUUID();
      const addHandoff = (tenant: string, owner: string) =>
        db().tenant(tenant, (sql) =>
          sql.query(
            `INSERT INTO prism365_handoff (tenant_id, id, audit_id, audit_title, run_id, complete,
               items, generated_at, generated_by)
             VALUES ($1,$2,'accounts.inactive','Inactive accounts',$3,true,0,now(),$3)`,
            [owner, handoff, randomUUID()],
          ),
        );
      await expect(addHandoff(tenantA, tenantB)).rejects.toMatchObject({ code: "42501" });
      await addHandoff(tenantB, tenantB);
      await expect(
        db().tenant(tenantA, (sql) =>
          sql.query(
            `INSERT INTO prism365_receipt (tenant_id, id, handoff_id, tool_name, tool_version,
               performed_at, performed_by, items, done_count, failed_count, skipped_count,
               imported_at, imported_by, digest, digest_version)
             VALUES ($1,$2,$3,'Example tool','1',now(),'Operator',1,1,0,0,now(),$1,$4,1)`,
            [tenantA, randomUUID(), handoff, "a".repeat(64)],
          ),
        ),
      ).rejects.toMatchObject({ code: "23503" });
      // An item must name an account one way or the other.
      await expect(
        db().tenant(tenantB, (sql) =>
          sql.query(
            `INSERT INTO prism365_receipt_item (tenant_id, receipt_id, position, user_id,
               user_principal_name, action, status, detail)
             VALUES ($1,$2,1,NULL,NULL,'account.disabled','done','')`,
            [tenantB, receipt],
          ),
        ),
      ).rejects.toMatchObject({ code: "23514" });
    });

    it("keeps one receipt when imports of one id overlap, and tells each which one", async () => {
      const id = randomUUID();
      const receipt = (items: number): Receipt => ({
        format: "prism365.receipt",
        version: 1,
        id,
        handoffId: null,
        tool: { name: "Example action tool", version: "1" },
        performedAt: new Date().toISOString(),
        performedBy: "Operator",
        items: Array.from({ length: items }, (_, index) => ({
          userId: `user-${index}`,
          userPrincipalName: null,
          action: "account.disabled",
          status: "done",
          detail: "",
        })),
      });
      // Eight imports at once: four of a one-item receipt, four of a two-item one.
      const outcomes = await Promise.all(
        Array.from({ length: 8 }, (_, index) =>
          db().tenant(tenantA, (sql) =>
            insertReceipt(sql, tenantA, receipt(index % 2 ? 1 : 2), {
              digest: (index % 2 ? "b" : "c").repeat(64),
              digestVersion: 1,
              importedBy: tenantA,
            }),
          ),
        ),
      );
      expect(outcomes.filter((outcome) => outcome.inserted)).toHaveLength(1);
      const winner = outcomes.find((outcome) => outcome.inserted);
      // Every other import sees the winner, whichever it lost to.
      for (const outcome of outcomes) {
        expect(outcome.digest).toBe(winner?.digest);
        expect(outcome.summary.items).toBe(winner?.summary.items);
      }
      const stored = await db().tenant(tenantA, (sql) =>
        sql.query<{ n: number }>(
          "SELECT count(*)::int AS n FROM prism365_receipt_item WHERE receipt_id=$1",
          [id],
        ),
      );
      expect(stored.rows[0]?.n).toBe(winner?.summary.items);
    });

    it("refuses identity, setting, schema, and multi-statement SQL at runtime", async () => {
      for (const statement of [
        "RESET ROLE",
        "SET LOCAL ROLE postgres",
        "SELECT set_config('prism365.tenant_id', 'x', true)",
        "SELECT current_setting('prism365.tenant_id')",
        "SELECT 1; RESET ROLE",
        "CREATE TABLE mod_test_probe_extra (tenant_id uuid NOT NULL)",
        "DROP TABLE mod_test_probe_item",
        "SELECT * FROM prism365_schema_migration",
        "SELECT pg_read_file('/etc/hosts')",
      ])
        await expect(db().tenant(tenantA, (sql) => sql.query(statement))).rejects.toBeInstanceOf(
          SqlGuardError,
        );
    });

    it("refuses to start when a module table lacks the tenant policy", async () => {
      await fixture.owner(
        "CREATE TABLE mod_test_probe_unsafe (tenant_id uuid NOT NULL, value text)",
      );
      try {
        await expect(db().verify()).rejects.toThrow("mod_test_probe_unsafe");
      } finally {
        await fixture.owner("DROP TABLE mod_test_probe_unsafe");
      }
      await expect(db().verify()).resolves.toBeUndefined();
    });

    it("refuses to start with an extra permissive policy, a trigger, or a definer function", async () => {
      for (const [setup, teardown] of [
        [
          "CREATE POLICY leak ON prism365_finding FOR SELECT USING (true)",
          "DROP POLICY leak ON prism365_finding",
        ],
        [
          "CREATE FUNCTION mod_test_probe_f() RETURNS trigger LANGUAGE plpgsql AS $$ BEGIN RETURN NEW; END $$; CREATE TRIGGER mod_test_probe_t BEFORE INSERT ON mod_test_probe_item FOR EACH ROW EXECUTE FUNCTION mod_test_probe_f()",
          "DROP TRIGGER mod_test_probe_t ON mod_test_probe_item; DROP FUNCTION mod_test_probe_f()",
        ],
        [
          "CREATE FUNCTION mod_test_probe_peek() RETURNS SETOF text LANGUAGE sql SECURITY DEFINER AS 'SELECT value FROM prism365_cursor'",
          "DROP FUNCTION mod_test_probe_peek()",
        ],
      ] as const) {
        await fixture.owner(setup);
        try {
          await expect(db().verify()).rejects.toThrow();
        } finally {
          await fixture.owner(teardown);
        }
      }
      await expect(db().verify()).resolves.toBeUndefined();
    });

    it("rejects a quoted set_config that would re-point the tenant", async () => {
      await expect(
        db().tenant(tenantA, (sql) =>
          sql.query(`SELECT "set_config"('prism365.tenant_id', $1, true)`, [tenantB]),
        ),
      ).rejects.toBeInstanceOf(SqlGuardError);
    });

    it("rejects a tenant id that is not a UUID", async () => {
      await expect(db().tenant("not-a-tenant", async () => null)).rejects.toThrow("UUID");
    });
  });
}

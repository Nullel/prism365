import { randomUUID } from "node:crypto";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
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

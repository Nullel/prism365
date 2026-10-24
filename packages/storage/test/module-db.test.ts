import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import {
  assertRuntimeStatement,
  MigrationError,
  type MigrationScope,
  PGliteModuleDatabase,
  SqlGuardError,
} from "../src/index.ts";
import { isolationSuite, probeScope, tenantA } from "./module-isolation-suite.ts";

isolationSuite("PGlite", async () => {
  const database = await PGliteModuleDatabase.open({ scopes: [probeScope] });
  return {
    database,
    owner: (sql) => database.ownerExecForTests(sql),
    close: () => database.close(),
  };
});

describe("migration ledger", () => {
  it("applies a module migration once and refuses one that changed after it was applied", async () => {
    const dataDir = await mkdtemp(join(tmpdir(), "prism365-ledger-"));
    try {
      const first = await PGliteModuleDatabase.open({ dataDir, scopes: [probeScope] });
      expect(first.durable).toBe(true);
      await first.tenant(tenantA, (sql) =>
        sql.query(
          "INSERT INTO mod_test_probe_item (tenant_id,id,value,created_at) VALUES ($1,'kept','synthetic',now())",
          [tenantA],
        ),
      );
      await first.close();

      const again = await PGliteModuleDatabase.open({ dataDir, scopes: [probeScope] });
      const kept = await again.tenant(tenantA, (sql) =>
        sql.query<{ id: string }>("SELECT id FROM mod_test_probe_item"),
      );
      expect(kept.rows).toEqual([{ id: "kept" }]);
      await again.close();

      const edited: MigrationScope = {
        scope: probeScope.scope,
        migrations: [{ id: "001-items", sql: `${probeScope.migrations[0]?.sql ?? ""}\n-- edited` }],
      };
      await expect(PGliteModuleDatabase.open({ dataDir, scopes: [edited] })).rejects.toBeInstanceOf(
        MigrationError,
      );
    } finally {
      await rm(dataDir, { recursive: true, force: true });
    }
  }, 60_000);
});

describe("runtime statement guard", () => {
  const access = {
    read: new Set(["mod_x_y", "prism365_finding"]),
    write: new Set(["mod_x_y"]),
  };

  it("allows one parameterized read or write over permitted tables", () => {
    for (const statement of [
      "SELECT 1 FROM mod_x_y WHERE id=$1;",
      "WITH t AS (SELECT 1) INSERT INTO mod_x_y SELECT * FROM t",
      'UPDATE "mod_x_y" AS s SET v = 1 FROM prism365_finding AS f WHERE s.id = f.id',
      "SELECT count(*) FILTER (WHERE a IN (1,2))::int, EXTRACT(HOUR FROM b AT TIME ZONE $1) FROM mod_x_y",
      "INSERT INTO mod_x_y (a) VALUES ($1) ON CONFLICT (a) DO UPDATE SET b = GREATEST(mod_x_y.b, EXCLUDED.b)",
    ])
      expect(() => assertRuntimeStatement(statement, access), statement).not.toThrow();
  });

  it("refuses the bypasses an adversarial review reproduced", () => {
    for (const statement of [
      `SELECT "set_config"('prism365.tenant_id', $1, true)`,
      "SELECT set_config/**/('role', 'postgres', true)",
      `SELECT "current_setting"('prism365.tenant_id')`,
      "SELECT query_to_xml('select set_config(''role'',''postgres'',true)', true, true, '')",
      "SELECT ts_stat('select 1')",
      "SELECT pg_catalog.pg_read_file('/etc/hosts')",
      "WITH prism365_state AS (SELECT * FROM prism365_state) SELECT * FROM prism365_state",
      "SELECT * FROM prism365_state",
      "DELETE FROM prism365_state WHERE kind = 'session'",
      "INSERT INTO prism365_audit (id) VALUES ($1)",
      "UPDATE prism365_finding SET state = 'open'",
      "SELECT * FROM pg_catalog.pg_roles",
      "SELECT * INTO mod_x_copy FROM mod_x_y",
      "SELECT * FROM other_schema.mod_x_y",
      "SELECT 1; RESET ROLE",
      "VACUUM",
      "COPY mod_x_y TO STDOUT",
      "DO $$ BEGIN END $$",
    ])
      expect(() => assertRuntimeStatement(statement, access), statement).toThrow(SqlGuardError);
  });
});

describe("transaction lifetime", () => {
  it("disables a client kept after its transaction and rolls back a stuck transaction", async () => {
    const database = await PGliteModuleDatabase.open({ scopes: [probeScope] });
    try {
      const kept: { client: { query(text: string): Promise<unknown> } | null } = { client: null };
      await database.tenant(tenantA, async (sql) => {
        kept.client = sql;
      });
      await expect(
        kept.client?.query("SELECT 1 FROM mod_test_probe_item") ?? Promise.resolve(),
      ).rejects.toBeInstanceOf(SqlGuardError);

      const controller = new AbortController();
      const stuck = database.tenant(
        tenantA,
        async (sql) => {
          await sql.query(
            "INSERT INTO mod_test_probe_item (tenant_id,id,value,created_at) VALUES ($1,'stuck','x',now())",
            [tenantA],
          );
          await new Promise(() => {});
        },
        { signal: controller.signal },
      );
      setTimeout(() => controller.abort(), 20);
      await expect(stuck).rejects.toMatchObject({ name: "TransactionAbortedError" });
      const after = await database.tenant(tenantA, (sql) =>
        sql.query<{ id: string }>("SELECT id FROM mod_test_probe_item WHERE id='stuck'"),
      );
      expect(after.rows).toEqual([]);
    } finally {
      await database.close();
    }
  });
});

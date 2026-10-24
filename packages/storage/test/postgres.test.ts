import { randomBytes, randomUUID } from "node:crypto";
import { readFile } from "node:fs/promises";
import { Pool, type PoolClient } from "pg";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import {
  PostgresModuleDatabase,
  PostgresStateStore,
  runPostgresMigrations,
  type SqlClient,
} from "../src/index.ts";
import { migrationsBeforeTrends, trendMigrationSuite } from "./audit-trend-migration-suite.ts";
import { isolationSuite, probeScope } from "./module-isolation-suite.ts";

// This public password belongs only to the disposable CI database; it is not a secret.
const syntheticPassword = "Prism365_Synthetic_Database_Test_Only_12345";
const adminUrl = process.env.TEST_DATABASE_ADMIN_URL;
const tenantA = randomUUID();
const tenantB = randomUUID();
const actor = randomUUID();
const suffix = randomBytes(8).toString("hex");
const databaseName = `prism365_test_${suffix}`;
const ownerRole = `prism365_test_owner_${suffix}`;

function identifier(value: string) {
  return `"${value.replaceAll('"', '""')}"`;
}

describe.skipIf(!adminUrl)("PostgreSQL encrypted state and enforced tenant isolation", () => {
  let bootstrap: Pool | undefined;
  let administrator: Pool | undefined;
  let restricted: Pool | undefined;
  let store: PostgresStateStore | undefined;
  let fixtureAdminUrl: string;
  let fixtureOwnerUrl: string;
  let createdDatabase = false;
  let createdAppRole = false;
  let createdOwnerRole = false;
  const encryptionKey = randomBytes(32).toString("hex");

  function stateStore() {
    if (!store) throw new Error("The PostgreSQL test fixture was not initialized.");
    return store;
  }

  function adminPool() {
    if (!administrator) throw new Error("The test administrator connection is unavailable.");
    return administrator;
  }

  async function withTenant<T>(tenantId: string, operation: (client: PoolClient) => Promise<T>) {
    if (!restricted) throw new Error("The restricted test connection is unavailable.");
    const client = await restricted.connect();
    try {
      await client.query("BEGIN");
      await client.query("SELECT set_config('prism365.tenant_id',$1,true)", [tenantId]);
      const result = await operation(client);
      await client.query("COMMIT");
      return result;
    } catch (error) {
      await client.query("ROLLBACK");
      throw error;
    } finally {
      client.release();
    }
  }

  beforeAll(async () => {
    const parsed = new URL(adminUrl ?? "");
    if (
      !["postgres:", "postgresql:"].includes(parsed.protocol) ||
      !["localhost", "127.0.0.1"].includes(parsed.hostname) ||
      parsed.pathname !== "/prism365_ci"
    ) {
      throw new Error(
        "TEST_DATABASE_ADMIN_URL must point to a disposable local prism365_ci database. The suite creates its own isolated test database.",
      );
    }
    bootstrap = new Pool({ connectionString: parsed.toString(), connectionTimeoutMillis: 5_000 });
    const role = await bootstrap.query<{ rolsuper: boolean; rolbypassrls: boolean }>(
      "SELECT rolsuper,rolbypassrls FROM pg_roles WHERE rolname='prism365_app'",
    );
    if (!role.rowCount) {
      await bootstrap.query(
        `CREATE ROLE prism365_app LOGIN NOSUPERUSER NOBYPASSRLS PASSWORD '${syntheticPassword}'`,
      );
      createdAppRole = true;
    } else if (role.rows[0]?.rolsuper || role.rows[0]?.rolbypassrls) {
      throw new Error("The existing test prism365_app role must be restricted.");
    }
    await bootstrap.query(
      `CREATE ROLE ${identifier(ownerRole)} LOGIN NOSUPERUSER NOBYPASSRLS PASSWORD '${syntheticPassword}'`,
    );
    createdOwnerRole = true;
    await bootstrap.query(`CREATE DATABASE ${identifier(databaseName)}`);
    createdDatabase = true;
    parsed.pathname = `/${databaseName}`;
    fixtureAdminUrl = parsed.toString();
    administrator = new Pool({ connectionString: fixtureAdminUrl, connectionTimeoutMillis: 5_000 });
    const migration = await readFile(
      new URL("../migrations/001-state.sql", import.meta.url),
      "utf8",
    );
    await administrator.query(migration);
    parsed.username = "prism365_app";
    parsed.password = syntheticPassword;
    const appUrl = parsed.toString();
    restricted = new Pool({ connectionString: appUrl, max: 3, connectionTimeoutMillis: 5_000 });
    store = new PostgresStateStore({ connectionString: appUrl, encryptionKey });
    await store.verifyRole();
    parsed.username = ownerRole;
    fixtureOwnerUrl = parsed.toString();
  }, 30_000);

  afterAll(async () => {
    await store?.close();
    await restricted?.end();
    await administrator?.end();
    if (bootstrap) {
      try {
        // These targets were created by this suite, with unguessable fixture-only names.
        if (createdDatabase)
          await bootstrap.query(`DROP DATABASE ${identifier(databaseName)} WITH (FORCE)`);
        if (createdOwnerRole) await bootstrap.query(`DROP ROLE ${identifier(ownerRole)}`);
        if (createdAppRole) await bootstrap.query("DROP ROLE prism365_app");
      } finally {
        await bootstrap.end();
      }
    }
  }, 30_000);

  it("enforces RLS when a query omits the application tenant filter", async () => {
    const key = "rls-filter-fixture";
    await stateStore().set(
      tenantA,
      "session",
      key,
      { marker: "synthetic-a" },
      new Date(Date.now() + 60_000),
    );
    await stateStore().set(
      tenantB,
      "session",
      key,
      { marker: "synthetic-b" },
      new Date(Date.now() + 60_000),
    );
    const rows = await withTenant(tenantA, (client) =>
      client.query<{ tenant_id: string }>("SELECT tenant_id FROM prism365_state"),
    );
    expect(rows.rows.length).toBeGreaterThan(0);
    expect(rows.rows.every((row) => row.tenant_id === tenantA)).toBe(true);
    const hidden = await restricted?.query("SELECT tenant_id FROM prism365_state");
    expect(hidden?.rows).toEqual([]);
    const update = await withTenant(tenantA, (client) =>
      client.query("UPDATE prism365_state SET expires_at=now() WHERE tenant_id=$1", [tenantB]),
    );
    expect(update.rowCount).toBe(0);
    const deletion = await withTenant(tenantA, (client) =>
      client.query("DELETE FROM prism365_state WHERE tenant_id=$1", [tenantB]),
    );
    expect(deletion.rowCount).toBe(0);
    await expect(
      withTenant(tenantA, (client) =>
        client.query(
          "INSERT INTO prism365_state(tenant_id,kind,key,encrypted,expires_at) VALUES($1,'session',$2,$3,now()+interval '1 minute')",
          [tenantB, "cross-tenant-insert", Buffer.from("synthetic-invalid-ciphertext")],
        ),
      ),
    ).rejects.toMatchObject({ code: "42501" });
    await expect(
      withTenant(tenantA, (client) =>
        client.query("UPDATE prism365_state SET tenant_id=$1,key=key||'-moved'", [tenantB]),
      ),
    ).rejects.toMatchObject({ code: "42501" });
    expect(await stateStore().get(tenantB, "session", key)).toEqual({ marker: "synthetic-b" });
  });

  it("encrypts credentials and rejects moving encrypted state to another tenant", async () => {
    const key = "encrypted-relocation-fixture";
    await stateStore().set(
      tenantA,
      "session",
      key,
      { token: "synthetic-sensitive-token" },
      new Date(Date.now() + 60_000),
    );
    const result = await adminPool().query<{ encrypted: Buffer }>(
      "SELECT encrypted FROM prism365_state WHERE tenant_id=$1 AND kind='session' AND key=$2",
      [tenantA, key],
    );
    const bytes = result.rows[0]?.encrypted;
    expect(bytes).toBeInstanceOf(Buffer);
    expect(bytes?.includes(Buffer.from("synthetic-sensitive-token"))).toBe(false);
    await withTenant(tenantB, (client) =>
      client.query(
        "INSERT INTO prism365_state(tenant_id,kind,key,encrypted,expires_at) VALUES($1,'session',$2,$3,now()+interval '1 minute')",
        [tenantB, key, bytes],
      ),
    );
    await expect(stateStore().get(tenantB, "session", key)).rejects.toThrow();
    expect(await stateStore().get(tenantA, "session", key)).toEqual({
      token: "synthetic-sensitive-token",
    });
  });

  it("consumes OAuth state once and prevents stale refresh from recreating a deleted session", async () => {
    const oauthKey = "atomic-oauth-fixture";
    await stateStore().set(
      tenantA,
      "oauth",
      oauthKey,
      { state: "synthetic-once" },
      new Date(Date.now() + 60_000),
    );
    const attempts = await Promise.all([
      stateStore().take(tenantA, "oauth", oauthKey),
      stateStore().take(tenantA, "oauth", oauthKey),
    ]);
    expect(attempts.filter((value) => value !== null)).toEqual([{ state: "synthetic-once" }]);
    expect(attempts.filter((value) => value === null)).toHaveLength(1);
    const sessionKey = "deleted-session-fixture";
    await stateStore().set(
      tenantA,
      "session",
      sessionKey,
      { cache: "synthetic-old" },
      new Date(Date.now() + 60_000),
    );
    expect(
      await stateStore().replace(
        tenantA,
        "session",
        sessionKey,
        { cache: "synthetic-new" },
        new Date(Date.now() + 60_000),
      ),
    ).toBe(true);
    await stateStore().delete(tenantA, "session", sessionKey);
    expect(
      await stateStore().replace(
        tenantA,
        "session",
        sessionKey,
        { cache: "synthetic-stale-refresh" },
        new Date(Date.now() + 60_000),
      ),
    ).toBe(false);
    expect(await stateStore().get(tenantA, "session", sessionKey)).toBeNull();
  });

  it("physically purges expired state only in the selected tenant", async () => {
    await stateStore().set(
      tenantA,
      "session",
      "expired-a",
      { cache: "synthetic-expired" },
      new Date(0),
    );
    await stateStore().set(
      tenantB,
      "session",
      "expired-b",
      { cache: "synthetic-other-tenant" },
      new Date(0),
    );
    await stateStore().set(
      tenantA,
      "session",
      "retained-a",
      { cache: "synthetic-current" },
      new Date(Date.now() + 60_000),
    );
    await stateStore().purgeExpired(tenantA);
    const rows = await adminPool().query<{ key: string }>(
      "SELECT key FROM prism365_state WHERE key IN ('expired-a','expired-b','retained-a') ORDER BY key",
    );
    expect(rows.rows.map((row) => row.key)).toEqual(["expired-b", "retained-a"]);
  });

  it("rejects superuser and table-owner credentials at runtime", async () => {
    const privileged = new PostgresStateStore({ connectionString: fixtureAdminUrl, encryptionKey });
    try {
      await expect(privileged.verifyRole()).rejects.toThrow("superuser");
    } finally {
      await privileged.close();
    }
    const originalOwner = await adminPool().query<{ tableowner: string }>(
      "SELECT tableowner FROM pg_tables WHERE schemaname='public' AND tablename='prism365_state'",
    );
    const owner = originalOwner.rows[0]?.tableowner;
    if (!owner) throw new Error("The isolated test table owner is unavailable.");
    await adminPool().query(`ALTER TABLE prism365_state OWNER TO ${identifier(ownerRole)}`);
    const tableOwnerStore = new PostgresStateStore({
      connectionString: fixtureOwnerUrl,
      encryptionKey,
    });
    try {
      await expect(tableOwnerStore.verifyRole()).rejects.toThrow("must not own");
    } finally {
      await tableOwnerStore.close();
      await adminPool().query(`ALTER TABLE prism365_state OWNER TO ${identifier(owner)}`);
    }
  });

  it("limits audit reads by tenant and forbids runtime UPDATE and DELETE", async () => {
    await stateStore().appendAudit(tenantA, actor, { type: "synthetic.test", marker: "a" });
    await stateStore().appendAudit(tenantB, actor, { type: "synthetic.test", marker: "b" });
    const rows = await withTenant(tenantA, (client) =>
      client.query<{ tenant_id: string }>("SELECT tenant_id FROM prism365_audit"),
    );
    expect(rows.rows).toEqual([{ tenant_id: tenantA }]);
    await expect(
      withTenant(tenantA, (client) => client.query("UPDATE prism365_audit SET event='{}'::jsonb")),
    ).rejects.toMatchObject({ code: "42501" });
    await expect(
      withTenant(tenantA, (client) => client.query("DELETE FROM prism365_audit")),
    ).rejects.toMatchObject({ code: "42501" });
  });
});

// Module storage on real PostgreSQL: the same isolation suite the PGlite run uses. It lives in
// this file so it runs after the state suite above instead of racing it for the shared
// prism365_app role, and it connects as a throwaway member of that role so the role's password
// is never read or changed.
isolationSuite(
  "PostgreSQL",
  async () => {
    const parsed = new URL(adminUrl ?? "");
    if (!["localhost", "127.0.0.1"].includes(parsed.hostname) || parsed.pathname !== "/prism365_ci")
      throw new Error("TEST_DATABASE_ADMIN_URL must point to a disposable prism365_ci database.");
    const suffix = randomBytes(8).toString("hex");
    const moduleDatabase = `prism365_modules_${suffix}`;
    const memberRole = `prism365_modules_member_${suffix}`;
    const bootstrap = new Pool({
      connectionString: parsed.toString(),
      connectionTimeoutMillis: 5_000,
    });
    const existing = await bootstrap.query("SELECT 1 FROM pg_roles WHERE rolname='prism365_app'");
    const createdAppRole = existing.rowCount === 0;
    await bootstrap.query(`CREATE DATABASE ${identifier(moduleDatabase)}`);
    parsed.pathname = `/${moduleDatabase}`;
    const moduleAdminUrl = parsed.toString();
    await runPostgresMigrations({
      adminUrl: moduleAdminUrl,
      appPassword: syntheticPassword,
      scopes: [probeScope],
    });
    await bootstrap.query(
      `CREATE ROLE ${identifier(memberRole)} LOGIN NOSUPERUSER NOBYPASSRLS PASSWORD '${syntheticPassword}' IN ROLE prism365_app`,
    );
    const owner = new Pool({ connectionString: moduleAdminUrl, connectionTimeoutMillis: 5_000 });
    parsed.username = memberRole;
    parsed.password = syntheticPassword;
    const database = new PostgresModuleDatabase({ connectionString: parsed.toString() });
    // verify() also grants the default table access; without it every query is refused.
    await database.verify();
    return {
      database,
      owner: async (sql: string) => {
        await owner.query(sql);
      },
      close: async () => {
        await database.close();
        await owner.end();
        try {
          await bootstrap.query(`DROP DATABASE ${identifier(moduleDatabase)} WITH (FORCE)`);
          await bootstrap.query(`DROP ROLE ${identifier(memberRole)}`);
          if (createdAppRole) await bootstrap.query("DROP ROLE prism365_app");
        } finally {
          await bootstrap.end();
        }
      },
    };
  },
  Boolean(adminUrl),
);

describe.skipIf(!adminUrl)("PostgreSQL module runtime role", () => {
  it("rejects a superuser connection for module data", async () => {
    const privileged = new PostgresModuleDatabase({ connectionString: adminUrl ?? "" });
    try {
      await expect(privileged.verify()).rejects.toThrow("superuser");
    } finally {
      await privileged.close();
    }
  });
});

// Migration 005 (audit trends) on real PostgreSQL, starting from a database migrated up to 004
// that already holds audit runs. Last in this file, so it never races the suites above for the
// shared prism365_app role.
trendMigrationSuite(
  "PostgreSQL",
  async () => {
    const parsed = new URL(adminUrl ?? "");
    if (!["localhost", "127.0.0.1"].includes(parsed.hostname) || parsed.pathname !== "/prism365_ci")
      throw new Error("TEST_DATABASE_ADMIN_URL must point to a disposable prism365_ci database.");
    const trendDatabase = `prism365_trends_${randomBytes(8).toString("hex")}`;
    const bootstrap = new Pool({
      connectionString: parsed.toString(),
      connectionTimeoutMillis: 5_000,
    });
    const existing = await bootstrap.query("SELECT 1 FROM pg_roles WHERE rolname='prism365_app'");
    const createdAppRole = existing.rowCount === 0;
    if (createdAppRole)
      await bootstrap.query(
        `CREATE ROLE prism365_app LOGIN NOSUPERUSER NOBYPASSRLS PASSWORD '${syntheticPassword}'`,
      );
    await bootstrap.query(`CREATE DATABASE ${identifier(trendDatabase)}`);
    parsed.pathname = `/${trendDatabase}`;
    const trendAdminUrl = parsed.toString();
    const owner = new Pool({ connectionString: trendAdminUrl, connectionTimeoutMillis: 5_000 });
    for (const migration of await migrationsBeforeTrends()) await owner.query(migration.sql);
    return {
      owner: async (sql: string, params?: unknown[]) => {
        // A migration file is several statements, which pg answers with one result each.
        const result = (await owner.query(sql, params)) as unknown;
        const last = (Array.isArray(result) ? result.at(-1) : result) as
          | { rows?: Record<string, unknown>[] }
          | undefined;
        return last?.rows ?? [];
      },
      migrateAll: () =>
        runPostgresMigrations({
          adminUrl: trendAdminUrl,
          appPassword: syntheticPassword,
          scopes: [],
        }),
      close: async () => {
        await owner.end();
        try {
          await bootstrap.query(`DROP DATABASE ${identifier(trendDatabase)} WITH (FORCE)`);
          if (createdAppRole) await bootstrap.query("DROP ROLE prism365_app");
        } finally {
          await bootstrap.end();
        }
      },
    };
  },
  Boolean(adminUrl),
);

// Two saves of the audit schedule at the same moment, on real PostgreSQL, where transactions do
// run side by side (PGlite runs one at a time). Each save locks the setting's row, so the second
// waits for the first and adds to it instead of overwriting it. The schedule code lives in
// @prism365/audits, which depends on this package, so it is loaded by path at run time and typed
// here by the little it needs.
interface ScheduleCode {
  builtInAudits: readonly unknown[];
  writeAuditSchedule(
    sql: SqlClient,
    tenantId: string,
    schedule: { audits: Record<string, string> },
    actorId: string,
    audits: readonly unknown[],
  ): Promise<unknown>;
  readStoredAuditSchedule(
    sql: SqlClient,
    tenantId: string,
  ): Promise<{ audits: Record<string, string> }>;
}
const scheduleCode = "../../audits/src/index.ts";

describe.skipIf(!adminUrl)("PostgreSQL: audit schedule saves at the same moment", () => {
  const suffix = randomBytes(8).toString("hex");
  const scheduleDatabase = `prism365_schedule_${suffix}`;
  const memberRole = `prism365_schedule_member_${suffix}`;
  const tenant = randomUUID();
  let bootstrap: Pool;
  let database: PostgresModuleDatabase;
  let createdAppRole = false;
  let schedule: ScheduleCode;

  beforeAll(async () => {
    schedule = (await import(/* @vite-ignore */ scheduleCode)) as ScheduleCode;
    const parsed = new URL(adminUrl ?? "");
    if (!["localhost", "127.0.0.1"].includes(parsed.hostname) || parsed.pathname !== "/prism365_ci")
      throw new Error("TEST_DATABASE_ADMIN_URL must point to a disposable prism365_ci database.");
    bootstrap = new Pool({ connectionString: parsed.toString(), connectionTimeoutMillis: 5_000 });
    const existing = await bootstrap.query("SELECT 1 FROM pg_roles WHERE rolname='prism365_app'");
    createdAppRole = existing.rowCount === 0;
    await bootstrap.query(`CREATE DATABASE ${identifier(scheduleDatabase)}`);
    parsed.pathname = `/${scheduleDatabase}`;
    await runPostgresMigrations({
      adminUrl: parsed.toString(),
      appPassword: syntheticPassword,
      scopes: [],
    });
    await bootstrap.query(
      `CREATE ROLE ${identifier(memberRole)} LOGIN NOSUPERUSER NOBYPASSRLS PASSWORD '${syntheticPassword}' IN ROLE prism365_app`,
    );
    parsed.username = memberRole;
    parsed.password = syntheticPassword;
    database = new PostgresModuleDatabase({ connectionString: parsed.toString() });
    await database.verify();
  }, 60_000);

  afterAll(async () => {
    await database?.close();
    if (!bootstrap) return;
    try {
      await bootstrap.query(`DROP DATABASE IF EXISTS ${identifier(scheduleDatabase)} WITH (FORCE)`);
      await bootstrap.query(`DROP ROLE IF EXISTS ${identifier(memberRole)}`);
      if (createdAppRole) await bootstrap.query("DROP ROLE IF EXISTS prism365_app");
    } finally {
      await bootstrap.end();
    }
  }, 30_000);

  it("keeps both of two saves of different audits", async () => {
    const { builtInAudits, readStoredAuditSchedule, writeAuditSchedule } = schedule;
    // Each save pauses after reading the setting, so without the row lock both would read the
    // same old value and the second write would erase the first.
    const slow = (sql: SqlClient): SqlClient => ({
      async query<T extends object>(text: string, params?: readonly unknown[]) {
        const result = await sql.query<T>(text, params);
        if (text.includes("FROM prism365_preference"))
          await new Promise((resolve) => setTimeout(resolve, 300));
        return result;
      },
    });
    const save = (audits: Record<string, "daily" | "weekly">) =>
      database.tenant(tenant, (sql) =>
        writeAuditSchedule(slow(sql), tenant, { audits }, actor, builtInAudits),
      );
    // First when no setting exists yet, then again once it does.
    await Promise.all([
      save({ "accounts.inactive": "daily" }),
      save({ "guests.review": "weekly" }),
    ]);
    await Promise.all([
      save({ "licenses.reclaim": "daily" }),
      save({ "mfa.registration": "weekly" }),
    ]);
    const stored = await database.tenant(tenant, (sql) => readStoredAuditSchedule(sql, tenant));
    expect(stored.audits).toMatchObject({
      "accounts.inactive": "daily",
      "guests.review": "weekly",
      "licenses.reclaim": "daily",
      "mfa.registration": "weekly",
    });
  }, 30_000);
});

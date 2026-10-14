import { randomBytes, randomUUID } from "node:crypto";
import { readFile } from "node:fs/promises";
import { Pool, type PoolClient } from "pg";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { PostgresStateStore } from "../src/index.ts";

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
        if (createdDatabase) {
          // Pool.end can resolve just before idle sockets finish their close handshake. Do not
          // force-terminate them: that races a late FATAL event against the test process.
          for (let attempt = 0; attempt < 50; attempt += 1) {
            const active = await bootstrap.query<{ count: string }>(
              "SELECT count(*) FROM pg_stat_activity WHERE datname=$1",
              [databaseName],
            );
            if (active.rows[0]?.count === "0") break;
            await new Promise((resolve) => setTimeout(resolve, 100));
          }
          await bootstrap.query(`DROP DATABASE ${identifier(databaseName)}`);
        }
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

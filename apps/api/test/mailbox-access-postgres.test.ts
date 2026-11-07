import { randomBytes } from "node:crypto";
import { GraphHttpClient } from "@prism365/graph";
import { ModuleScheduler } from "@prism365/modules";
import {
  MemoryStateStore,
  type ModuleDatabase,
  PostgresModuleDatabase,
  runPostgresMigrations,
  type SqlClient,
} from "@prism365/storage";
import type { InjectOptions } from "fastify";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import {
  disposableAdminUrl,
  holdPostgresSuiteLock,
  ownerPool,
} from "../../../packages/storage/test/postgres-suite-lock.ts";
import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import { currentImport } from "../src/mailbox-access-store.ts";
import { createModuleRegistry } from "../src/modules.ts";

// Mailbox access imports on real PostgreSQL, where transactions overlap and hold row locks, which
// PGlite (one connection) never does, and where the runtime role is really held to row security.
// Runs only with TEST_DATABASE_ADMIN_URL, against a database it creates and drops. Synthetic
// tenant, administrator, mailboxes and accounts only.

const adminUrl = process.env.TEST_DATABASE_ADMIN_URL;
// This public password belongs only to the disposable test database; it is not a secret.
const syntheticPassword = "Prism365_Synthetic_Database_Test_Only_12345";
const tenantId = "11111111-1111-4111-8111-111111111111";
const otherTenant = "99999999-9999-4999-8999-999999999999";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
const MAILBOX = "eeeeeeee-0000-4000-8000-000000000001";
const features = ["directory.users.read"] as const;
const identityConfig: IdentityConfig = {
  tenantId,
  clientId: "33333333-3333-4333-8333-333333333333",
  origin,
  redirectUri: `${origin}/api/v1/auth/callback`,
  allowedActorIds: [actorId],
  enabledFeatures: [...features],
  credential: { clientSecret: "synthetic-only" },
};
const config: AppConfig = {
  mode: "live",
  origin,
  identity: identityConfig,
  production: false,
  missingConfiguration: [],
  enabledFeatures: [...features],
};
const provider: MicrosoftIdentityProvider = {
  authorizationUrl: async () => "https://login.microsoftonline.com/synthetic",
  redeem: async () => ({
    tenantId,
    actorId,
    displayName: "Synthetic admin",
    accountId: "synthetic-account",
    cache: "synthetic-cache",
    scopes: ["User.Read", "User.Read.All"],
  }),
  token: async (session) => ({ accessToken: "synthetic", cache: "c", scopes: session.scopes }),
};

// An export with `grants` grants on one mailbox, each to its own synthetic account, numbered from
// `first`.
const exportFile = (grants: number, first = 0) => ({
  format: "prism365.mailbox-access",
  version: 1,
  tenantId,
  exportedAt: new Date(Date.now() - 86_400_000).toISOString(),
  includeUserMailboxes: false,
  mailboxes: [
    {
      id: MAILBOX,
      address: "lab.results@example.invalid",
      displayName: "Lab results",
      type: "SharedMailbox",
    },
  ],
  grants: Array.from({ length: grants }, (_, index) => ({
    mailboxId: MAILBOX,
    right: "FullAccess",
    trustee: {
      id: `cccccccc-0000-4000-8000-${String(first + index).padStart(12, "0")}`,
      address: `p${first + index}@example.invalid`,
      name: `Person ${first + index}`,
      kind: "user",
    },
  })),
  unresolved: 0,
});

// Runs once, right after the person route's `after`-th read of the import table has answered, as
// another administrator's import committing at that moment would.
let betweenReads: { after: number; run: () => Promise<unknown> } | null = null;
function watched(database: ModuleDatabase): ModuleDatabase {
  return {
    kind: database.kind,
    durable: database.durable,
    tenant: (tenant, operation, options) =>
      database.tenant(
        tenant,
        (sql) =>
          operation({
            query: async <T extends object>(text: string, params?: readonly unknown[]) => {
              const result = await sql.query<T>(text, params);
              if (betweenReads && text.includes("FROM prism365_mailbox_access_import")) {
                betweenReads.after -= 1;
                if (betweenReads.after === 0) {
                  const run = betweenReads.run;
                  betweenReads = null;
                  await run();
                }
              }
              return result;
            },
          } satisfies SqlClient),
        options,
      ),
    verify: () => database.verify(),
    close: () => database.close(),
  };
}

function identifier(value: string) {
  return `"${value.replaceAll('"', '""')}"`;
}
const sleep = (milliseconds: number) => new Promise((done) => setTimeout(done, milliseconds));
const within = <T>(milliseconds: number, pending: Promise<T>) =>
  Promise.race([pending, sleep(milliseconds).then(() => "waiting" as const)]);

describe.skipIf(!adminUrl)("mailbox access on PostgreSQL", () => {
  const suffix = randomBytes(8).toString("hex");
  const databaseName = `prism365_mailbox_${suffix}`;
  const memberRole = `prism365_mailbox_member_${suffix}`;
  let releaseSuiteLock: (() => Promise<void>) | undefined;
  let bootstrap: ReturnType<typeof ownerPool> | undefined;
  let owner: ReturnType<typeof ownerPool>;
  let database: PostgresModuleDatabase;
  let app: ReturnType<typeof createApp> | undefined;
  let identity: IdentityService;
  let admin: { cookie: string; csrf: string };
  let createdAppRole = false;

  beforeAll(async () => {
    releaseSuiteLock = await holdPostgresSuiteLock(adminUrl ?? "");
    const parsed = disposableAdminUrl(adminUrl ?? "");
    bootstrap = ownerPool(parsed.toString());
    const existing = await bootstrap.query("SELECT 1 FROM pg_roles WHERE rolname='prism365_app'");
    createdAppRole = existing.rowCount === 0;
    await bootstrap.query(`CREATE DATABASE ${identifier(databaseName)}`);
    parsed.pathname = `/${databaseName}`;
    const registry = createModuleRegistry(config);
    await runPostgresMigrations({
      adminUrl: parsed.toString(),
      appPassword: syntheticPassword,
      scopes: registry.migrationScopes(),
    });
    await bootstrap.query(
      `CREATE ROLE ${identifier(memberRole)} LOGIN NOSUPERUSER NOBYPASSRLS PASSWORD '${syntheticPassword}' IN ROLE prism365_app`,
    );
    owner = ownerPool(parsed.toString());
    parsed.username = memberRole;
    parsed.password = syntheticPassword;
    database = new PostgresModuleDatabase({ connectionString: parsed.toString() });
    await database.verify();
    const store = new MemoryStateStore();
    identity = new IdentityService(identityConfig, store, provider);
    const scheduler = new ModuleScheduler({
      enabledPacks: config.enabledFeatures,
      registry,
      database,
      graphClient: (getAccessToken) => new GraphHttpClient({ getAccessToken }),
    });
    app = createApp({
      config,
      store,
      identity,
      // The person route reads the account's groups; here Microsoft lists none.
      graphFetch: (async () => Response.json({ value: [] })) as typeof fetch,
      modules: { registry, database: watched(database), scheduler },
    });
  }, 180_000);

  afterAll(async () => {
    // Closing the app closes the module database too.
    await app?.close();
    await owner?.end();
    try {
      if (bootstrap) {
        await bootstrap.query(`DROP DATABASE IF EXISTS ${identifier(databaseName)} WITH (FORCE)`);
        await bootstrap.query(`DROP ROLE IF EXISTS ${identifier(memberRole)}`);
        if (createdAppRole) await bootstrap.query("DROP ROLE prism365_app");
      }
    } finally {
      await bootstrap?.end();
      await releaseSuiteLock?.();
    }
  }, 60_000);

  beforeEach(async () => {
    const login = await identity.startLogin("directory.users.read");
    const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
    admin = {
      cookie: `prism365_session=${loggedIn.handle}`,
      csrf: (await identity.session(loggedIn.handle))?.csrfToken ?? "",
    };
    await owner.query("DELETE FROM prism365_mailbox_access_grant");
    await owner.query("DELETE FROM prism365_mailbox_access_import");
    // Each test starts as a tenant that has never imported, with no lock row yet.
    await owner.query("DELETE FROM prism365_mailbox_access_lock");
  });

  const inject = (options: InjectOptions) => {
    if (!app) throw new Error("The app did not start.");
    return app.inject(options);
  };
  const upload = (payload: unknown) =>
    inject({
      method: "POST",
      url: "/v1/mailbox-access/import",
      headers: {
        cookie: admin.cookie,
        origin,
        "x-csrf-token": admin.csrf,
        "content-type": "application/json",
      },
      payload: JSON.stringify(payload),
    });
  const clear = () =>
    inject({
      method: "DELETE",
      url: "/v1/mailbox-access",
      headers: { cookie: admin.cookie, origin, "x-csrf-token": admin.csrf },
    });
  // Every grant belongs to the one import there is, and there are as many as it says.
  const consistent = async () => {
    const imports = await owner.query<{ id: string; grant_count: number }>(
      "SELECT id, grant_count FROM prism365_mailbox_access_import WHERE tenant_id=$1",
      [tenantId],
    );
    const grants = await owner.query<{ import_id: string; n: number }>(
      "SELECT import_id, count(*)::int AS n FROM prism365_mailbox_access_grant GROUP BY import_id",
    );
    expect(imports.rows).toHaveLength(1);
    expect(grants.rows).toEqual([
      { import_id: imports.rows[0]?.id, n: imports.rows[0]?.grant_count },
    ]);
  };

  it("keeps one whole import when several race, never a mix", async () => {
    const answers = await Promise.all([
      upload(exportFile(450)),
      upload(exportFile(30)),
      upload(exportFile(5)),
    ]);
    for (const answer of answers) expect([201, 409], answer.body).toContain(answer.statusCode);
    expect(answers.some((answer) => answer.statusCode === 201)).toBe(true);
    await consistent();
    // And again with an import already there, which the replace holds while it works.
    const again = await Promise.all([upload(exportFile(200)), upload(exportFile(7))]);
    for (const answer of again) expect([201, 409], answer.body).toContain(answer.statusCode);
    await consistent();
  }, 60_000);

  it("waits for a transaction that holds the tenant's lock, then replaces what it committed", async () => {
    expect((await upload(exportFile(3))).statusCode).toBe(201);
    const holder = await owner.connect();
    await holder.query("BEGIN");
    await holder.query("SELECT tenant_id FROM prism365_mailbox_access_lock FOR UPDATE");
    const pending = upload(exportFile(4));
    try {
      expect(await within(1_500, pending)).toBe("waiting");
    } finally {
      await holder.query("COMMIT");
      holder.release();
    }
    const answer = await pending;
    expect(answer.statusCode).toBe(201);
    expect(answer.json().replaced).toBe(true);
    await consistent();
  }, 30_000);

  it("makes a clear wait for a first import still in progress, then clears it", async () => {
    // A first import between taking the tenant's lock and committing: the lock row and the import
    // are written but not committed, so a clear that locked only the import row would find nothing
    // to lock, delete nothing and leave the import behind.
    const holder = await owner.connect();
    await holder.query("BEGIN");
    await holder.query(
      "INSERT INTO prism365_mailbox_access_lock (tenant_id) VALUES ($1) ON CONFLICT (tenant_id) DO NOTHING",
      [tenantId],
    );
    await holder.query(
      "SELECT tenant_id FROM prism365_mailbox_access_lock WHERE tenant_id=$1 FOR UPDATE",
      [tenantId],
    );
    await holder.query(
      `INSERT INTO prism365_mailbox_access_import (tenant_id, id, imported_at, exported_at,
         imported_by, include_user_mailboxes, mailbox_count, grant_count, unresolved_count)
       VALUES ($1,'44444444-4444-4444-8444-444444444444',now(),now(),$2,false,1,0,0)`,
      [tenantId, actorId],
    );
    const pending = clear();
    try {
      expect(await within(1_500, pending)).toBe("waiting");
    } finally {
      await holder.query("COMMIT");
      holder.release();
    }
    const answer = await pending;
    expect(answer.statusCode).toBe(200);
    expect(answer.json()).toEqual({ cleared: true });
    const left = await owner.query(
      "SELECT 1 FROM prism365_mailbox_access_import WHERE tenant_id=$1",
      [tenantId],
    );
    expect(left.rowCount).toBe(0);
  }, 30_000);

  it("pairs an export's date with its own grants when an import commits between the reads", async () => {
    const person = "cccccccc-0000-4000-8000-000000000000";
    const old = (await upload(exportFile(3))).json().import;
    // The first read is the route's check that an import exists, before it asks Microsoft. Right
    // after the second, the read that answers the page, a newer export without this person commits.
    let replaced = false;
    betweenReads = {
      after: 2,
      run: async () => {
        expect((await upload(exportFile(2, 1_000))).statusCode).toBe(201);
        replaced = true;
      },
    };
    const answer = (
      await inject({
        method: "GET",
        url: `/v1/people/${person}/mailbox-access`,
        headers: { cookie: admin.cookie },
      })
    ).json();
    expect(replaced).toBe(true);
    // Either export may answer, but never one's date with the other's grants.
    if (answer.import.id === old.id) expect(answer.grants).toHaveLength(1);
    else expect(answer.grants).toHaveLength(0);
    expect(answer.import.id).toBe(old.id);
  }, 30_000);

  it("keeps each tenant's rows to that tenant under row security", async () => {
    expect((await upload(exportFile(3))).statusCode).toBe(201);
    // The runtime role, in another tenant's transaction, without the tenant predicate.
    const seen = await database.tenant(otherTenant, async (sql) => ({
      grants: (
        await sql.query<{ n: number }>(
          "SELECT count(*)::int AS n FROM prism365_mailbox_access_grant",
        )
      ).rows[0]?.n,
      imports: (
        await sql.query<{ n: number }>(
          "SELECT count(*)::int AS n FROM prism365_mailbox_access_import",
        )
      ).rows[0]?.n,
      current: await currentImport(sql, otherTenant),
    }));
    expect(seen).toEqual({ grants: 0, imports: 0, current: null });
    // Nor can that tenant's transaction write a row for this one, or clear this one's rows.
    await expect(
      database.tenant(otherTenant, (sql) =>
        sql.query(
          `INSERT INTO prism365_mailbox_access_import (tenant_id, id, imported_at, exported_at,
             imported_by, include_user_mailboxes, mailbox_count, grant_count, unresolved_count)
           VALUES ($1,'44444444-4444-4444-8444-444444444444',now(),now(),$2,false,0,0,0)`,
          [tenantId, actorId],
        ),
      ),
    ).rejects.toThrow();
    await database.tenant(otherTenant, (sql) =>
      sql.query("DELETE FROM prism365_mailbox_access_grant"),
    );
    await consistent();
  }, 30_000);
});

import { randomBytes, randomUUID } from "node:crypto";
import { finishRun, inactiveAccountsAudit, insertRun } from "@prism365/audits";
import { type AuditRow, receiptSchema } from "@prism365/contracts";
import { GraphHttpClient } from "@prism365/graph";
import { ModuleScheduler } from "@prism365/modules";
import { MemoryStateStore, PostgresModuleDatabase, runPostgresMigrations } from "@prism365/storage";
import type { InjectOptions } from "fastify";
import { afterAll, beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import {
  disposableAdminUrl,
  holdPostgresSuiteLock,
  ownerPool,
} from "../../../packages/storage/test/postgres-suite-lock.ts";
import { createApp } from "../src/app.ts";
import { AuditRunner } from "../src/audits.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { insertReceipt, pruneHandoffs } from "../src/handoff-store.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import { createModuleRegistry } from "../src/modules.ts";

// Handoffs and receipts on real PostgreSQL, where transactions overlap and hold row locks, which
// PGlite (one connection) never does. Runs only with TEST_DATABASE_ADMIN_URL, against a database
// it creates and drops. Synthetic tenant, administrators, accounts and tools only.

const adminUrl = process.env.TEST_DATABASE_ADMIN_URL;
// This public password belongs only to the disposable test database; it is not a secret.
const syntheticPassword = "Prism365_Synthetic_Database_Test_Only_12345";
const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
const features = ["directory.users.read", "audit.signins.read"] as const;
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
    scopes: ["User.Read", "User.Read.All", "AuditLog.Read.All"],
  }),
  token: async (session) => ({ accessToken: "synthetic", cache: "c", scopes: session.scopes }),
};
const account: AuditRow = {
  key: "user-3",
  personId: "00000003-aaaa-4aaa-8aaa-000000000003",
  tags: ["d7", "d30"],
  cells: { person: { name: "Synthetic Person 3", principal: "person3@example.invalid" }, days: 40 },
};

function identifier(value: string) {
  return `"${value.replaceAll('"', '""')}"`;
}

const sleep = (milliseconds: number) => new Promise((done) => setTimeout(done, milliseconds));
// The request's answer, or "waiting" if it has none after `milliseconds`: a request that waits on
// another transaction's lock never answers while that transaction stays open.
const within = <T>(milliseconds: number, pending: Promise<T>) =>
  Promise.race([pending, sleep(milliseconds).then(() => "waiting" as const)]);

describe.skipIf(!adminUrl)("handoffs on PostgreSQL", () => {
  const suffix = randomBytes(8).toString("hex");
  const databaseName = `prism365_handoffs_${suffix}`;
  const memberRole = `prism365_handoffs_member_${suffix}`;
  let releaseSuiteLock: (() => Promise<void>) | undefined;
  let bootstrap: ReturnType<typeof ownerPool> | undefined;
  let owner: ReturnType<typeof ownerPool>;
  let database: PostgresModuleDatabase;
  let app: ReturnType<typeof createApp> | undefined;
  let store: MemoryStateStore;
  let identity: IdentityService;
  let admin: { cookie: string; csrf: string };
  let runId: string;
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
    store = new MemoryStateStore();
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
      audits: new AuditRunner({ database }),
      modules: { registry, database, scheduler },
    });
    runId = randomUUID();
    await database.tenant(tenantId, async (sql) => {
      await insertRun(sql, tenantId, {
        id: runId,
        auditId: inactiveAccountsAudit.definition.id,
        auditVersion: inactiveAccountsAudit.definition.version,
        startedAt: new Date(),
        startedBy: { id: actorId, name: "Synthetic admin" },
      });
      await finishRun(sql, tenantId, runId, new Date(), {
        status: "succeeded",
        complete: true,
        summary: [],
        notes: [],
        rows: [account],
      });
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
  });

  const inject = (options: InjectOptions) => {
    if (!app) throw new Error("The app did not start.");
    return app.inject(options);
  };
  const get = (url: string) => inject({ method: "GET", url, headers: { cookie: admin.cookie } });
  const importReceipt = (payload: unknown) =>
    inject({
      method: "POST",
      url: "/v1/receipts",
      headers: {
        cookie: admin.cookie,
        origin,
        "x-csrf-token": admin.csrf,
        "content-type": "application/json",
      },
      payload: JSON.stringify(payload),
    });
  const receipt = (handoffId: string | null, extra = {}) => ({
    format: "prism365.receipt",
    version: 1,
    id: randomUUID(),
    handoffId,
    tool: { name: "Example action tool", version: "1.0.0" },
    performedAt: new Date(Date.now() - 60_000).toISOString(),
    performedBy: "Operator One",
    items: [
      {
        userId: account.personId,
        userPrincipalName: "person3@example.invalid",
        action: "account.sign-in-blocked",
        status: "done",
        detail: "",
      },
    ],
    ...extra,
  });
  // Fixtures written as the owner, past row security, with the right tenant.
  const oldHandoff = async () => {
    const id = randomUUID();
    await owner.query(
      `INSERT INTO prism365_handoff (tenant_id, id, audit_id, audit_title, run_id, complete, items,
         generated_at, generated_by)
       VALUES ($1,$2,'accounts.inactive','Inactive accounts',$3,true,1,now() - interval '4 years',$4)`,
      [tenantId, id, randomUUID(), actorId],
    );
    return id;
  };
  const oldReceipt = async () => {
    const id = randomUUID();
    await owner.query(
      `INSERT INTO prism365_receipt (tenant_id, id, handoff_id, tool_name, tool_version,
         performed_at, performed_by, items, done_count, failed_count, skipped_count, imported_at,
         imported_by, digest, digest_version)
       VALUES ($1,$2,NULL,'Example action tool','1',now() - interval '4 years','Operator',1,1,0,0,
               now() - interval '4 years',$3,$4,1)`,
      [tenantId, id, actorId, "a".repeat(64)],
    );
    await owner.query(
      `INSERT INTO prism365_receipt_item (tenant_id, receipt_id, position, user_id, action,
         status, detail) VALUES ($1,$2,0,'user-old','account.disabled','done','')`,
      [tenantId, id],
    );
    return id;
  };
  const exists = async (table: string, id: string) =>
    ((await owner.query(`SELECT 1 FROM ${table} WHERE id=$1`, [id])).rowCount ?? 0) > 0;
  // An open transaction that answers `handoff` with a receipt of its own, as an import does
  // between its insert and its commit: its foreign key holds the handoff row.
  const answering = async (handoff: string) => {
    const holder = await owner.connect();
    await holder.query("BEGIN");
    await holder.query(
      `INSERT INTO prism365_receipt (tenant_id, id, handoff_id, tool_name, tool_version,
         performed_at, performed_by, items, done_count, failed_count, skipped_count, imported_at,
         imported_by, digest, digest_version)
       VALUES ($1,$2,$3,'Example action tool','1',now(),'Operator',1,1,0,0,now(),$4,$5,1)`,
      [tenantId, randomUUID(), handoff, actorId, "b".repeat(64)],
    );
    return holder;
  };

  it("never waits for, or fails on, an expired handoff another import is answering", async () => {
    const handoff = await oldHandoff();
    const expired = await oldReceipt();
    const holder = await answering(handoff);
    try {
      // A receipt that answers no handoff, and an export, while the old handoff is held.
      const unrelated = receipt(null);
      const imported = await within(5_000, importReceipt(unrelated));
      expect(imported === "waiting" ? "waiting" : imported.statusCode).toBe(201);
      const exported = await within(5_000, get(`/v1/audit-runs/${runId}/handoff.json`));
      expect(exported === "waiting" ? "waiting" : exported.statusCode).toBe(200);
      expect(await exists("prism365_receipt", unrelated.id)).toBe(true);
    } finally {
      await holder.query("COMMIT");
      holder.release();
    }
    // The clean-up deleted what nobody held, and skipped the held handoff, which a receipt now
    // answers and which therefore stays.
    expect(await exists("prism365_receipt", expired)).toBe(false);
    expect(await exists("prism365_handoff", handoff)).toBe(true);
  }, 30_000);

  it("refuses a receipt whose handoff retention deletes while the import waits for it", async () => {
    const handoff = await oldHandoff();
    // A clean-up that has taken the handoff and is about to commit its deletion.
    const cleanup = await owner.connect();
    await cleanup.query("BEGIN");
    await cleanup.query("DELETE FROM prism365_handoff WHERE id=$1", [handoff]);
    const answer = receipt(handoff);
    const ledgered = store.audit.filter((entry) => entry.event.type === "receipt.imported").length;
    const pending = importReceipt(answer);
    try {
      // Wait until the import is blocked on the handoff's row: it saw the handoff, recorded the
      // import and inserted, and its foreign key now waits for the deletion.
      for (let attempt = 0; attempt < 100; attempt += 1) {
        const waiting = await owner.query(
          "SELECT 1 FROM pg_stat_activity WHERE datname = current_database() AND wait_event_type = 'Lock'",
        );
        if (waiting.rowCount) break;
        await sleep(50);
      }
    } finally {
      await cleanup.query("COMMIT");
      cleanup.release();
    }
    const response = await pending;
    expect(response.statusCode).toBe(400);
    expect(response.json()).toMatchObject({ error: "unknown_handoff" });
    // It got past the check (so the refusal came from the foreign key), and kept nothing.
    expect(store.audit.filter((entry) => entry.event.type === "receipt.imported")).toHaveLength(
      ledgered + 1,
    );
    expect(await exists("prism365_receipt", answer.id)).toBe(false);
  }, 30_000);

  it("measures retention by the database's clock, never the server's", async () => {
    const tenant = randomUUID();
    const id = randomUUID();
    await database.tenant(tenant, async (sql) => {
      await insertReceipt(sql, tenant, receiptSchema.parse({ ...receipt(null), id }), {
        digest: "a".repeat(64),
        digestVersion: 1,
        importedBy: actorId,
      });
    });
    await owner.query(
      "UPDATE prism365_receipt SET imported_at = now() - interval '1 year' WHERE id=$1",
      [id],
    );
    // The server's clock five years ahead: the receipt is still a year old by the database's.
    vi.useFakeTimers({ toFake: ["Date"] });
    vi.setSystemTime(new Date(Date.now() + 5 * 365 * 86_400_000));
    try {
      const outcome = await database.tenant(tenant, (sql) =>
        pruneHandoffs(sql, tenant, { days: 1_095 }),
      );
      expect(outcome.receipts).toBe(0);
    } finally {
      vi.useRealTimers();
    }
    expect(await exists("prism365_receipt", id)).toBe(true);
  }, 30_000);
});

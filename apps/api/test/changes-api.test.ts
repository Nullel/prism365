import { GraphHttpClient } from "@prism365/graph";
import { createDirectoryChangesModule } from "@prism365/module-directory-changes";
import { createSigninsAuditModule } from "@prism365/module-signins-audit";
import { ModuleRegistry, ModuleScheduler } from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";

vi.setConfig({ testTimeout: 30_000 });

// Synthetic tenant, administrators and changes only.
const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const personId = "44444444-4444-4444-8444-444444444444";
const adminId = "55555555-5555-4555-8555-555555555555";
const origin = "http://localhost:5173";
const features = ["audit.directory.read"] as const;
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
    scopes: ["User.Read", "AuditLog.Read.All"],
  }),
  token: async (session) => ({ accessToken: "synthetic", cache: "c", scopes: session.scopes }),
};

let store: MemoryStateStore;
let database: PGliteModuleDatabase;
let app: ReturnType<typeof createApp>;
// The same installation without the change module, and one without module storage.
let withoutModule: ReturnType<typeof createApp>;
let withoutStorage: ReturnType<typeof createApp>;
let cookie: string;
let otherCookie: string;
let plainCookie: string;

async function openApp(registry: ModuleRegistry, sessionStore: MemoryStateStore) {
  const identity = new IdentityService(identityConfig, sessionStore, provider);
  const login = await identity.startLogin("audit.directory.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
  const db = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  const scheduler = new ModuleScheduler({
    enabledPacks: [...features],
    registry,
    database: db,
    graphClient: (getAccessToken) =>
      new GraphHttpClient({
        getAccessToken,
        fetchImplementation: async () => Response.json({ value: [] }),
      }),
  });
  return {
    db,
    cookie: `prism365_session=${loggedIn.handle}`,
    app: createApp({
      config,
      store: sessionStore,
      identity,
      modules: { registry, database: db, scheduler },
    }),
  };
}

async function insert(id: string, at: string, values: Record<string, unknown>) {
  const targets = values.targets ?? [];
  await database.tenant(tenantId, async (sql) => {
    await sql.query(
      `INSERT INTO mod_directory_changes_event (tenant_id, id, activity_at, category, activity,
         result, initiated_by_kind, initiated_by_id, initiated_by_name, targets, details,
         processed, collected_at)
       VALUES ($1,$2,$3,$4,$5,$6,'user',$7,'Synthetic Admin',$8::jsonb,'[]'::jsonb,true,$3)`,
      [
        tenantId,
        id,
        at,
        values.category,
        values.activity,
        values.result ?? "success",
        adminId,
        JSON.stringify(targets),
      ],
    );
    await sql.query(
      `INSERT INTO mod_directory_changes_subject (tenant_id, event_id, subject_id, subject_kind, subject_type, activity_at)
       VALUES ($1,$2,$3,'actor','user',$4)`,
      [tenantId, id, adminId, at],
    );
    for (const target of targets as { id: string; type: string }[])
      await sql.query(
        `INSERT INTO mod_directory_changes_subject (tenant_id, event_id, subject_id, subject_kind, subject_type, activity_at)
         VALUES ($1,$2,$3,'target',$4,$5)`,
        [tenantId, id, target.id, target.type.toLowerCase(), at],
      );
  });
}

const ago = (hours: number) => new Date(Date.now() - hours * 3_600_000).toISOString();

beforeAll(async () => {
  store = new MemoryStateStore();
  const main = await openApp(new ModuleRegistry().register(createDirectoryChangesModule()), store);
  database = main.db;
  app = main.app;
  cookie = main.cookie;
  const other = await openApp(
    new ModuleRegistry().register(createSigninsAuditModule()),
    new MemoryStateStore(),
  );
  withoutModule = other.app;
  otherCookie = other.cookie;
  const plainStore = new MemoryStateStore();
  const plainIdentity = new IdentityService(identityConfig, plainStore, provider);
  const login = await plainIdentity.startLogin("audit.directory.read");
  const loggedIn = await plainIdentity.finishLogin("synthetic-code", login.state, login.state);
  plainCookie = `prism365_session=${loggedIn.handle}`;
  withoutStorage = createApp({ config, store: plainStore, identity: plainIdentity });

  const person = {
    id: personId,
    name: "Synthetic Person",
    type: "User",
    principal: null,
    properties: [],
  };
  await insert("Directory_1", ago(72), {
    category: "RoleManagement",
    activity: "Add member to role",
    targets: [person],
  });
  await insert("Directory_2", ago(48), {
    category: "GroupManagement",
    activity: "Add group",
    result: "failure",
  });
  await insert("Directory_3", ago(24), { category: "UserManagement", activity: "Update user" });
  await insert("Directory_4", ago(2), {
    category: "UserManagement",
    activity: "Reset user password",
    targets: [person],
  });
  await insert("Directory_old", ago(24 * 40), { category: "UserManagement", activity: "Add user" });
}, 60_000);

afterAll(async () => {
  await app?.close();
  await withoutModule?.close();
  await withoutStorage?.close();
});

const get = (url: string, session = cookie, target = app) =>
  target.inject({ method: "GET", url, headers: session ? { cookie: session } : {} });

describe("change feed API", () => {
  it("requires a session", async () => {
    for (const url of ["/v1/changes", `/v1/people/${personId}/changes`])
      expect((await get(url, "")).statusCode).toBe(401);
  });

  it("answers 409 in plain words without module storage or without the module", async () => {
    for (const [target, session] of [
      [withoutStorage, plainCookie],
      [withoutModule, otherCookie],
    ] as const)
      for (const url of ["/v1/changes", `/v1/people/${personId}/changes`]) {
        const response = await get(url, session, target);
        expect(response.statusCode, url).toBe(409);
        expect(response.json()).toEqual({
          error: "modules_unavailable",
          message: "The directory change feed is not available in this installation.",
        });
      }
  });

  it("refuses invalid filters, and a bad page position before recording a read", async () => {
    const before = store.audit.length;
    for (const query of [
      "days=0",
      "days=3651",
      "result=maybe",
      `search=${"x".repeat(201)}`,
      "limit=201",
      "before=not-a-position",
    ]) {
      const response = await get(`/v1/changes?${query}`);
      expect(response.statusCode, query).toBe(400);
      expect(response.json().message).toBeTruthy();
    }
    expect((await get("/v1/people/not%20an%20id/changes")).statusCode).toBe(400);
    expect((await get(`/v1/people/${personId}/changes?limit=0`)).statusCode).toBe(400);
    expect(store.audit.length).toBe(before);
  });

  it("lists the last 7 days with counts and coverage, and writes the ledger entry", async () => {
    const response = await get("/v1/changes");
    expect(response.statusCode).toBe(200);
    const body = response.json();
    expect(body.days).toBe(7);
    expect(body.items.map((item: { id: string }) => item.id)).toEqual([
      "Directory_4",
      "Directory_3",
      "Directory_2",
      "Directory_1",
    ]);
    expect(body.categories).toEqual([
      { category: "UserManagement", count: 2 },
      { category: "GroupManagement", count: 1 },
      { category: "RoleManagement", count: 1 },
    ]);
    expect(body.coverage).toMatchObject({ phase: "none", through: null });
    expect(store.audit.at(-1)).toMatchObject({
      tenantId,
      actorId,
      event: { type: "changes.viewed", datasets: ["directoryAudit"], days: 7 },
    });
    // Counts and ids only: no names, filters' text or values in the ledger.
    expect(JSON.stringify(store.audit.at(-1))).not.toContain("Synthetic");
    expect((await get("/v1/changes?days=3650")).json().items).toHaveLength(5);
  });

  it("filters by category, result and text, and pages with `before`", async () => {
    const ids = async (query: string) =>
      (await get(`/v1/changes?${query}`)).json().items.map((item: { id: string }) => item.id);
    expect(await ids("category=usermanagement")).toEqual(["Directory_4", "Directory_3"]);
    expect(await ids("result=failure")).toEqual(["Directory_2"]);
    expect(await ids("result=success&search=password")).toEqual(["Directory_4"]);
    expect(await ids("search=synthetic%20person")).toEqual(["Directory_4", "Directory_1"]);
    const first = (await get("/v1/changes?limit=3")).json();
    expect(first.items).toHaveLength(3);
    expect(first.next).toBeTruthy();
    const second = (
      await get(`/v1/changes?limit=3&before=${encodeURIComponent(first.next)}`)
    ).json();
    expect(second.items.map((item: { id: string }) => item.id)).toEqual(["Directory_1"]);
    expect(second.next).toBeNull();
  });

  it("lists a person's changes as target and an administrator's as actor", async () => {
    const response = await get(`/v1/people/${personId}/changes`);
    expect(response.statusCode).toBe(200);
    expect(
      response
        .json()
        .items.map((item: { id: string; involvement: unknown }) => [item.id, item.involvement]),
    ).toEqual([
      ["Directory_4", { actor: false, target: true }],
      ["Directory_1", { actor: false, target: true }],
    ]);
    expect(store.audit.at(-1)?.event).toMatchObject({
      type: "person.changes.viewed",
      datasets: ["directoryAudit"],
    });
    expect(JSON.stringify(store.audit.at(-1))).not.toContain(personId);
    const actor = (await get(`/v1/people/${adminId}/changes?limit=2`)).json();
    expect(actor.items.map((item: { id: string }) => item.id)).toEqual([
      "Directory_4",
      "Directory_3",
    ]);
    expect(actor.next).toBeTruthy();
  });
});

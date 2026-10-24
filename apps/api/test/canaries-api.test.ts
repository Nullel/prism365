import { GraphHttpClient } from "@prism365/graph";
import { createServiceHealthModule } from "@prism365/module-service-health";
import { readCanaries, writeCanaries } from "@prism365/module-signins-audit";
import { ModuleRegistry, ModuleScheduler } from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import { createModuleRegistry } from "../src/modules.ts";

vi.setConfig({ testTimeout: 30_000 });

// Canary accounts: the list an administrator keeps, and the routes that change it. Synthetic
// tenant, administrator and accounts only; every address ends in .invalid.
const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
const stored = "aaaaaaaa-0000-4000-8000-000000000001"; // has stored sign-ins
const listed = "aaaaaaaa-0000-4000-8000-000000000002"; // only in the directory
const unlisted = "aaaaaaaa-0000-4000-8000-000000000003"; // nowhere
const broken = "aaaaaaaa-0000-4000-8000-000000000004"; // the directory fails for it

const identityFor = (features: readonly string[]): IdentityConfig => ({
  tenantId,
  clientId: "33333333-3333-4333-8333-333333333333",
  origin,
  redirectUri: `${origin}/api/v1/auth/callback`,
  allowedActorIds: [actorId],
  enabledFeatures: features as IdentityConfig["enabledFeatures"],
  credential: { clientSecret: "synthetic-only" },
});
const configFor = (features: readonly string[]): AppConfig => ({
  mode: "live",
  origin,
  identity: identityFor(features),
  production: false,
  missingConfiguration: [],
  enabledFeatures: features as AppConfig["enabledFeatures"],
});
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

async function harness(
  features: readonly string[],
  modules: "signins" | "other" | "none" = "signins",
) {
  const requested: string[] = [];
  const graphFetch = (async (input: RequestInfo | URL) => {
    const url = new URL(String(input));
    requested.push(`${url.pathname}${url.search}`);
    if (url.pathname === `/v1.0/users/${listed}`)
      return Response.json({ id: listed, userPrincipalName: "Listed.Decoy@example.invalid" });
    if (url.pathname === `/v1.0/users/${broken}`) return new Response("{}", { status: 500 });
    return new Response("{}", { status: 404 });
  }) as typeof fetch;
  const config = configFor(features);
  const store = new MemoryStateStore();
  const identity = new IdentityService(identityFor(features), store, provider);
  const login = await identity.startLogin(features[0] as "directory.users.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
  const who = {
    cookie: `prism365_session=${loggedIn.handle}`,
    csrf: (await identity.session(loggedIn.handle))?.csrfToken ?? "",
  };
  const registry =
    modules === "signins"
      ? createModuleRegistry(config)
      : new ModuleRegistry().register(createServiceHealthModule());
  const database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  const scheduler = new ModuleScheduler({
    enabledPacks: config.enabledFeatures,
    registry,
    database,
    graphClient: (getAccessToken) => new GraphHttpClient({ getAccessToken }),
  });
  const app = createApp({
    config,
    store,
    identity,
    graphFetch,
    ...(modules === "none" ? {} : { modules: { registry, database, scheduler } }),
  });
  const get = (url: string, session: typeof who | null = who) =>
    app.inject({ method: "GET", url, headers: session ? { cookie: session.cookie } : {} });
  const post = (url: string, body?: unknown, session: typeof who | null = who, csrf = true) =>
    app.inject({
      method: "POST",
      url,
      headers: {
        ...(session ? { cookie: session.cookie, origin } : {}),
        ...(session && csrf ? { "x-csrf-token": session.csrf } : {}),
      },
      ...(body === undefined ? {} : { payload: body as object }),
    });
  const list = () => database.tenant(tenantId, (sql) => readCanaries(sql, tenantId));
  const ledger = () => store.audit.map((entry) => entry.event as Record<string, unknown>);
  // A sign-in the module already collected, so an account has a name in stored history.
  const seed = (userId: string, upn: string, at: string) =>
    database.tenant(tenantId, (sql) =>
      sql.query(
        `INSERT INTO mod_signins_audit_event (tenant_id, id, created_at, user_id, upn, error_code,
           risk_event_types, backfilled, processed, collected_at)
         VALUES ($1,$2,$3,$4,$5,0,'[]'::jsonb,false,true,$3)`,
        [tenantId, `event-${userId}-${at}`, at, userId, upn],
      ),
    );
  return { app, store, database, who, get, post, list, ledger, seed, requested };
}

type Harness = Awaited<ReturnType<typeof harness>>;
const both = ["directory.users.read", "audit.signins.read"] as const;

describe("canary accounts API", () => {
  let h: Harness;
  beforeAll(async () => {
    h = await harness(both);
    await h.seed(stored, "stored.decoy@example.invalid", "2026-10-20T10:00:00.000Z");
    // The newest name wins when an account was renamed.
    await h.seed(stored, "Stored.Decoy.Renamed@example.invalid", "2026-10-21T10:00:00.000Z");
  }, 60_000);
  afterAll(async () => {
    await h?.app.close();
  });

  it("needs a session to read or change the list, and the CSRF header to change it", async () => {
    expect((await h.get("/v1/settings/canaries", null)).statusCode).toBe(401);
    expect((await h.post("/v1/settings/canaries", { add: [stored] }, null)).statusCode).toBe(401);
    const noCsrf = await h.post("/v1/settings/canaries", { add: [stored] }, h.who, false);
    expect(noCsrf.statusCode).toBe(403);
    expect(await h.list()).toEqual({ accounts: [] });
  });

  it("starts with an empty list and records nothing for reading it", async () => {
    const before = h.store.audit.length;
    const response = await h.get("/v1/settings/canaries");
    expect(response.statusCode).toBe(200);
    expect(response.json()).toEqual({ accounts: [] });
    expect(h.store.audit.length).toBe(before);
  });

  it("refuses a body that names no valid account, before anything is recorded", async () => {
    const bodies: unknown[] = [
      undefined,
      {},
      { add: [] },
      { add: "x" },
      { add: ["not an id"] },
      { remove: [7] },
      { add: [stored], remove: [stored] },
      { add: Array.from({ length: 101 }, (_value, index) => `user-${index}`) },
    ];
    const before = h.store.audit.length;
    for (const body of bodies) {
      const response = await h.post("/v1/settings/canaries", body);
      expect(response.statusCode, JSON.stringify(body)).toBe(400);
      expect(response.json().error).toBe("invalid_request");
      expect(response.json().message).toBeTruthy();
    }
    expect(h.store.audit.length).toBe(before);
    expect(await h.list()).toEqual({ accounts: [] });
  });

  it("marks an account known from stored sign-ins, with its newest name, and writes the ledger first", async () => {
    // What the list holds at the moment the change is recorded: it must still be empty.
    let heldWhenRecorded: number | null = null;
    const original = h.store.appendAudit.bind(h.store);
    h.store.appendAudit = async (tenant, actor, event) => {
      if (event.type === "signins.canaries.changed")
        heldWhenRecorded = (await h.list()).accounts.length;
      return original(tenant, actor, event);
    };
    const before = h.store.audit.length;
    const response = await h.post("/v1/settings/canaries", { add: [stored] });
    h.store.appendAudit = original;
    expect(response.statusCode).toBe(200);
    const body = response.json();
    expect(body.accounts).toEqual([
      {
        userId: stored,
        principal: "Stored.Decoy.Renamed@example.invalid",
        addedAt: expect.any(String),
        addedBy: actorId,
      },
    ]);
    expect(heldWhenRecorded).toBe(0);
    expect((await h.list()).accounts).toHaveLength(1);
    // Stored history was enough: nothing was asked of Microsoft.
    expect(h.requested).toEqual([]);
    const entries = h.ledger().slice(before);
    expect(entries.map((entry) => entry.type)).toEqual([
      "signins.canaries.lookup",
      "signins.canaries.changed",
    ]);
    expect(entries[0]).toMatchObject({ datasets: ["signIns"], accounts: 1 });
    expect(entries[1]).toMatchObject({ added: 1, removed: 0 });
    // Counts only: no id, no name, no address.
    expect(JSON.stringify(entries)).not.toContain(stored);
    expect(JSON.stringify(entries)).not.toContain("example.invalid");
    expect((await h.get("/v1/settings/canaries")).json()).toEqual(body);
  });

  it("looks up an account with no stored sign-in in the directory, and records that first", async () => {
    const before = h.store.audit.length;
    const response = await h.post("/v1/settings/canaries", { add: [listed] });
    expect(response.statusCode).toBe(200);
    expect(response.json().accounts.map((item: { userId: string }) => item.userId)).toEqual([
      stored,
      listed,
    ]);
    expect(response.json().accounts[1].principal).toBe("Listed.Decoy@example.invalid");
    expect(h.requested).toEqual([`/v1.0/users/${listed}?$select=id,userPrincipalName`]);
    const entries = h.ledger().slice(before);
    expect(entries.map((entry) => [entry.type, entry.datasets])).toEqual([
      ["signins.canaries.lookup", ["signIns"]],
      ["signins.canaries.lookup", ["directory"]],
      ["signins.canaries.changed", undefined],
    ]);
    expect(JSON.stringify(entries)).not.toContain(listed);
    expect(JSON.stringify(entries)).not.toContain("Listed.Decoy");
  });

  it("refuses an account that is nowhere, and one the directory cannot answer for, and changes nothing", async () => {
    const before = h.store.audit.length;
    const stored1 = await h.list();
    const unknown = await h.post("/v1/settings/canaries", { add: [unlisted] });
    expect(unknown.statusCode).toBe(404);
    expect(unknown.json()).toMatchObject({ error: "account_not_found" });
    expect(unknown.json().message).toBeTruthy();
    // One unknown account refuses the whole request, even next to a good one.
    const mixed = await h.post("/v1/settings/canaries", { add: [unlisted], remove: [stored] });
    expect(mixed.statusCode).toBe(404);
    // An id that is not a GUID is never put into a directory path.
    const asked = h.requested.length;
    const odd = await h.post("/v1/settings/canaries", { add: [".."] });
    expect(odd.statusCode).toBe(404);
    expect(h.requested.length).toBe(asked);
    const down = await h.post("/v1/settings/canaries", { add: [broken] });
    expect(down.statusCode).toBe(503);
    expect(down.json()).toMatchObject({ error: "directory_unavailable" });
    expect(await h.list()).toEqual(stored1);
    // The lookups were recorded; no change was.
    expect(
      h
        .ledger()
        .slice(before)
        .map((entry) => entry.type),
    ).not.toContain("signins.canaries.changed");
  });

  it("changes only what is named, and ignores what changes nothing", async () => {
    const first = (await h.list()).accounts.find((item) => item.userId === stored);
    const before = h.store.audit.length;
    // Marking a marked account again, and removing one that is not marked, change nothing.
    const same = await h.post("/v1/settings/canaries", { add: [stored], remove: [unlisted] });
    expect(same.statusCode).toBe(200);
    expect(same.json().accounts).toHaveLength(2);
    expect(
      h
        .ledger()
        .slice(before)
        .map((entry) => entry.type),
    ).not.toContain("signins.canaries.changed");
    // A real removal keeps the others untouched, including when they were added.
    const removed = await h.post("/v1/settings/canaries", { remove: [listed] });
    expect(removed.statusCode).toBe(200);
    expect(removed.json().accounts).toEqual([first]);
    const last = h.ledger().at(-1);
    expect(last).toMatchObject({ type: "signins.canaries.changed", added: 0, removed: 1 });
  });

  it("holds at most 100 accounts, counting a removal in the same request", async () => {
    const accounts = Array.from({ length: 100 }, (_value, index) => ({
      userId: `bulk-${index}`,
      principal: `bulk.${index}@example.invalid`,
      addedAt: "2026-10-22T12:00:00.000Z",
      addedBy: actorId,
    }));
    await h.database.tenant(tenantId, (sql) => writeCanaries(sql, tenantId, { accounts }, actorId));
    await h.seed(listed, "listed.decoy@example.invalid", "2026-10-21T10:00:00.000Z");
    const before = h.store.audit.length;
    const over = await h.post("/v1/settings/canaries", { add: [listed] });
    expect(over.statusCode).toBe(400);
    expect(over.json().message).toContain("100");
    expect((await h.list()).accounts).toHaveLength(100);
    // Refused before any lookup or ledger entry.
    expect(h.store.audit.length).toBe(before);
    const swap = await h.post("/v1/settings/canaries", { add: [listed], remove: ["bulk-0"] });
    expect(swap.statusCode).toBe(200);
    expect(swap.json().accounts).toHaveLength(100);
    expect(swap.json().accounts.some((item: { userId: string }) => item.userId === listed)).toBe(
      true,
    );
  });
});

describe("canary accounts API when two administrators change the list at once", () => {
  const entry = (userId: string) => ({
    userId,
    principal: `${userId}@example.invalid`,
    addedAt: "2026-10-22T12:00:00.000Z",
    addedBy: actorId,
  });

  it("keeps a change made between reading the list and saving it, and still holds the limit", async () => {
    const h = await harness(both);
    try {
      await h.seed(stored, "stored.decoy@example.invalid", "2026-10-21T10:00:00.000Z");
      // The other administrator's change lands while this request is being recorded.
      const original = h.store.appendAudit.bind(h.store);
      let rival: string[] = [];
      h.store.appendAudit = async (tenant, actor, event) => {
        if (event.type === "signins.canaries.changed") {
          const now = await h.list();
          await h.database.tenant(tenantId, (sql) =>
            writeCanaries(
              sql,
              tenantId,
              { accounts: [...now.accounts, ...rival.map(entry)] },
              actorId,
            ),
          );
        }
        return original(tenant, actor, event);
      };

      rival = ["rival-1"];
      const merged = await h.post("/v1/settings/canaries", { add: [stored] });
      expect(merged.statusCode).toBe(200);
      expect(
        merged
          .json()
          .accounts.map((item: { userId: string }) => item.userId)
          .sort(),
      ).toEqual([stored, "rival-1"].sort());

      // 99 held when the request is checked, 100 by the time it is saved: one too many.
      await h.database.tenant(tenantId, (sql) =>
        writeCanaries(
          sql,
          tenantId,
          { accounts: Array.from({ length: 98 }, (_value, index) => entry(`bulk-${index}`)) },
          actorId,
        ),
      );
      await h.seed(listed, "listed.decoy@example.invalid", "2026-10-21T10:00:00.000Z");
      rival = ["rival-2", "rival-3"];
      const refused = await h.post("/v1/settings/canaries", { add: [listed] });
      expect(refused.statusCode).toBe(400);
      expect(refused.json().message).toContain("100");
      const held = (await h.list()).accounts.map((item) => item.userId);
      expect(held).toHaveLength(100);
      expect(held).not.toContain(listed);
    } finally {
      await h.app.close();
    }
  });
});

describe("canary accounts API without what it needs", () => {
  it("answers 409 without module storage, and without the sign-in module", async () => {
    const plain = await harness(both, "none");
    try {
      const response = await plain.get("/v1/settings/canaries");
      expect(response.statusCode).toBe(409);
      expect(response.json()).toMatchObject({ error: "modules_unavailable" });
      expect((await plain.post("/v1/settings/canaries", { add: [stored] })).statusCode).toBe(409);
    } finally {
      await plain.app.close();
    }
    // Module storage, but a registry that has no sign-in module.
    const none = await harness(both, "other");
    try {
      expect((await none.get("/v1/settings/canaries")).statusCode).toBe(409);
      expect((await none.post("/v1/settings/canaries", { add: [stored] })).statusCode).toBe(409);
    } finally {
      await none.app.close();
    }
  });

  it("needs the directory pack to look up an account with no stored sign-in", async () => {
    const limited = await harness(["audit.signins.read"]);
    try {
      await limited.seed(stored, "stored.decoy@example.invalid", "2026-10-21T10:00:00.000Z");
      const known = await limited.post("/v1/settings/canaries", { add: [stored] });
      expect(known.statusCode).toBe(200);
      const unknown = await limited.post("/v1/settings/canaries", { add: [listed] });
      expect(unknown.statusCode).toBe(403);
      expect(unknown.json()).toMatchObject({ error: "missing_consent" });
      expect(limited.requested).toEqual([]);
      expect((await limited.list()).accounts).toHaveLength(1);
    } finally {
      await limited.app.close();
    }
  });

  it("in demo mode answers like any signed-out request", async () => {
    const demo = createApp();
    try {
      expect((await demo.inject({ method: "GET", url: "/v1/settings/canaries" })).statusCode).toBe(
        401,
      );
    } finally {
      await demo.close();
    }
  });
});

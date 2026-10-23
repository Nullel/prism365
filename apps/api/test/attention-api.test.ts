import { GraphHttpClient } from "@prism365/graph";
import { defaultAttentionPreferences, ModuleScheduler, upsertAttention } from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import { createModuleRegistry } from "../src/modules.ts";

// Synthetic tenant and invented incident identifiers only.
const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
const features = ["directory.users.read", "service.health.read", "service.messages.read"] as const;
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
// The administrator granted service health but not Message center.
const provider: MicrosoftIdentityProvider = {
  authorizationUrl: async () => "https://login.microsoftonline.com/synthetic",
  redeem: async () => ({
    tenantId,
    actorId,
    displayName: "Synthetic admin",
    accountId: "synthetic-account",
    cache: "synthetic-cache",
    scopes: ["User.Read", "User.Read.All", "ServiceHealth.Read.All"],
  }),
  token: async (session) => ({ accessToken: "synthetic", cache: "c", scopes: session.scopes }),
};

let app: ReturnType<typeof createApp>;
let store: MemoryStateStore;
let cookie: string;
let csrf: string;
let itemId: string;
let scheduler: ModuleScheduler;

beforeAll(async () => {
  store = new MemoryStateStore();
  const identity = new IdentityService(identityConfig, store, provider);
  const login = await identity.startLogin("directory.users.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
  cookie = `prism365_session=${loggedIn.handle}`;
  csrf = (await identity.session(loggedIn.handle))?.csrfToken ?? "";
  const registry = createModuleRegistry(config);
  const database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  scheduler = new ModuleScheduler({
    enabledPacks: config.enabledFeatures,
    registry,
    database,
    graphClient: (getAccessToken) =>
      new GraphHttpClient({
        getAccessToken,
        fetchImplementation: async () => Response.json({ value: [] }),
      }),
  });
  app = createApp({ config, store, identity, modules: { registry, database, scheduler } });
  itemId = (
    await database.tenant(tenantId, (sql) =>
      upsertAttention(
        sql,
        tenantId,
        "service.health",
        {
          kind: "service.incident",
          ref: "EX300001",
          class: "incident",
          title: "Some users cannot send mail",
          summary: "Users may be unable to send mail.",
          services: ["Exchange Online"],
          provider: {
            status: "Investigating",
            labels: ["Incident"],
            deadline: null,
            updatedAt: new Date(),
            active: true,
            impact: "Some users may be unable to send mail.",
            workaround: null,
          },
          localLabels: [],
          severity: "high",
          link: "https://admin.microsoft.com/#/servicehealth/:/alerts/EX300001",
          meaningful: { Status: "Investigating" },
          observedAt: new Date(),
        },
        defaultAttentionPreferences,
      ),
    )
  ).id;
}, 60_000);

afterAll(async () => {
  await app?.close();
});

const post = (url: string, payload?: unknown, withCsrf = true) =>
  app.inject({
    method: "POST",
    url,
    headers: {
      cookie,
      origin,
      ...(withCsrf ? { "x-csrf-token": csrf } : {}),
      ...(payload === undefined ? {} : { "content-type": "application/json" }),
    },
    ...(payload === undefined ? {} : { payload: JSON.stringify(payload) }),
  });

describe("notification center API", () => {
  it("requires a session", async () => {
    for (const url of [
      "/v1/attention",
      "/v1/settings/attention",
      "/v1/modules/service.health/status",
    ]) {
      const response = await app.inject({ method: "GET", url });
      expect(response.statusCode, url).toBe(401);
    }
  });

  it("lists items, counts, sources and the problems that stop checking", async () => {
    const response = await app.inject({ method: "GET", url: "/v1/attention", headers: { cookie } });
    expect(response.statusCode).toBe(200);
    const body = response.json();
    expect(body.items.map((item: { ref: string }) => item.ref)).toEqual(["EX300001"]);
    expect(body.counts).toEqual({ unread: 1, active: 1 });
    expect(body.sources).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ module: "service.health", state: "not_collecting" }),
        expect.objectContaining({
          module: "service.messages",
          state: "missing_permission",
          missing: ["ServiceMessage.Read.All"],
        }),
      ]),
    );
    const ids = body.system.map((alert: { id: string }) => alert.id);
    expect(ids).toContain("collection.stopped");
    expect(ids).toContain("permission.service.messages");
    const grant = body.system.find(
      (alert: { id: string }) => alert.id === "permission.service.messages",
    );
    expect(grant.action.href).toBe("/api/v1/auth/login?feature=service.messages.read");
    // The link has to work: the login route accepts every enabled read pack, not just the
    // first three.
    const consent = await app.inject({ method: "GET", url: grant.action.href.slice(4) });
    expect(consent.statusCode).toBe(302);
    expect(consent.headers.location).toBe("https://login.microsoftonline.com/synthetic");
  });

  it("changes review and mute state only with CSRF, and leaves Microsoft's status alone", async () => {
    expect(
      (await post(`/v1/attention/${itemId}/review`, { state: "reviewed" }, false)).statusCode,
    ).toBe(403);
    const reviewed = await post(`/v1/attention/${itemId}/review`, { state: "reviewed" });
    expect(reviewed.statusCode).toBe(200);
    expect(reviewed.json()).toMatchObject({
      review: { state: "reviewed" },
      unread: false,
      provider: { status: "Investigating", active: true },
    });
    expect(store.audit.at(-1)?.event).toMatchObject({
      type: "attention.reviewed",
      state: "reviewed",
    });
    const muted = await post(`/v1/attention/${itemId}/mute`, { muted: true });
    expect(muted.json()).toMatchObject({ muted: true });
    const open = await app.inject({
      method: "GET",
      url: "/v1/attention?filter=open",
      headers: { cookie },
    });
    expect(open.json().items).toEqual([]);
    expect((await post(`/v1/attention/${itemId}/review`, { state: "archived" })).statusCode).toBe(
      400,
    );
    expect((await post("/v1/attention/not-a-uuid/read")).statusCode).toBe(400);
    expect((await post("/v1/attention/read-all")).json()).toEqual({ updated: 0 });
  });

  it("saves notification settings and rejects nonsense", async () => {
    const defaults = await app.inject({
      method: "GET",
      url: "/v1/settings/attention",
      headers: { cookie },
    });
    expect(defaults.json().preferences).toMatchObject({ services: [], notify: { incident: true } });
    expect((await post("/v1/settings/attention", { services: "Exchange" })).statusCode).toBe(400);
    const saved = await post("/v1/settings/attention", {
      services: ["Exchange Online"],
      notify: {
        incident: true,
        advisory: false,
        "message.action": true,
        "message.retirement": true,
        "message.change": false,
      },
    });
    expect(saved.statusCode).toBe(200);
    expect(saved.json().preferences).toMatchObject({
      services: ["Exchange Online"],
      notify: { advisory: false },
    });
  });

  it("reports the service-health check as never run, and Message center as not granted", async () => {
    const response = await app.inject({
      method: "GET",
      url: "/v1/modules/service.health/status",
      headers: { cookie },
    });
    expect(response.statusCode).toBe(200);
    expect(response.json()).toMatchObject({
      state: "never",
      incidents: [expect.objectContaining({ ref: "EX300001" })],
      messages: { state: "missing_permission" },
    });
  });

  it("judges a source by the session that collects, not by whoever is looking", async () => {
    // Another administrator collects with a session that never granted ServiceHealth.Read.All.
    scheduler.bind({
      kind: "delegated",
      tenantId,
      actorId: "44444444-4444-4444-8444-444444444444",
      grantedPermissions: async () => ["User.Read"],
      accessToken: async () => "synthetic",
    });
    await scheduler.tick();
    const status = await app.inject({
      method: "GET",
      url: "/v1/modules/service.health/status",
      headers: { cookie },
    });
    expect(status.json().state).toBe("missing_permission");
    const list = await app.inject({ method: "GET", url: "/v1/attention", headers: { cookie } });
    const alert = list
      .json()
      .system.find((item: { id: string }) => item.id === "permission.service.health");
    expect(alert.detail).toContain("The session Prism365 collects with");
    scheduler.unbind("operator_stopped");
  });
});

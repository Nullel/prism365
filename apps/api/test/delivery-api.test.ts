import { GraphHttpClient } from "@prism365/graph";
import { type CollectionCredential, ModuleRegistry, ModuleScheduler } from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import { finding, probeModule } from "../../../packages/modules/test/fixtures.ts";
import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import type { WebhookConnector } from "../src/delivery-webhook.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";

vi.setConfig({ testTimeout: 30_000 });

// The alerts API, and alert delivery wired into a running app: the scheduler's sink, the access
// ledger, and what the answers never contain. Synthetic tenant, administrator, hosts and
// addresses only.

const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
const hookUrl = "https://hooks.example/workflows/synthetic?sig=secret-token-value";
const identityConfig: IdentityConfig = {
  tenantId,
  clientId: "33333333-3333-4333-8333-333333333333",
  origin,
  redirectUri: `${origin}/api/v1/auth/callback`,
  allowedActorIds: [actorId],
  enabledFeatures: ["directory.users.read"],
  credential: { clientSecret: "synthetic-only" },
};
const config: AppConfig = {
  mode: "live",
  origin,
  identity: identityConfig,
  production: false,
  missingConfiguration: [],
  enabledFeatures: ["directory.users.read"],
  delivery: {
    email: {
      smtp: {
        host: "relay.example",
        port: 465,
        implicitTls: true,
        user: "alerts",
        password: "synthetic-relay-password",
        from: "alerts@example.invalid",
      },
      domains: ["example.invalid"],
    },
    webhookPorts: [],
  },
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

let pending: ReturnType<typeof finding>[] = [];
const sent: { host: string; body: string }[] = [];
let connect: WebhookConnector = async (target, request) => {
  sent.push({ host: target.host, body: request.body });
  return { status: 200 };
};
const module = probeModule({
  run: async (context) => {
    for (const input of pending) await context.findings.upsert(input);
    pending = [];
    return { complete: true };
  },
});

let app: ReturnType<typeof createApp>;
let store: MemoryStateStore;
let scheduler: ModuleScheduler;
let admin: { cookie: string; csrf: string };

beforeAll(async () => {
  const registry = new ModuleRegistry().register(module);
  const database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  store = new MemoryStateStore();
  const identity = new IdentityService(identityConfig, store, provider);
  const fetchImplementation = async () =>
    Response.json({ value: [{ id: "synthetic-user-1", displayName: "Synthetic User" }] });
  scheduler = new ModuleScheduler({
    enabledPacks: ["directory.users.read"],
    registry,
    database,
    graphClient: (getAccessToken) =>
      new GraphHttpClient({ getAccessToken, fetchImplementation, sleep: async () => {} }),
  });
  app = createApp({
    config,
    store,
    identity,
    modules: { registry, database, scheduler },
    delivery: {
      lookup: async () => [{ address: "100.128.0.1", family: 4 }],
      connect: (target, request, limits) => connect(target, request, limits),
      mailSender: { sendMail: async () => ({}) },
      flushDelayMilliseconds: 0,
      log: () => {},
    },
  });
  const login = await identity.startLogin("directory.users.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
  admin = {
    cookie: `prism365_session=${loggedIn.handle}`,
    csrf: (await identity.session(loggedIn.handle))?.csrfToken ?? "",
  };
}, 60_000);

afterAll(async () => {
  await app?.close();
});

beforeEach(() => {
  sent.length = 0;
});

const get = (url: string, cookie = admin.cookie) =>
  app.inject({ method: "GET", url, headers: cookie ? { cookie } : {} });
const send = (method: "POST" | "DELETE", url: string, body?: unknown, csrf = admin.csrf) =>
  app.inject({
    method,
    url,
    headers: { cookie: admin.cookie, origin, "x-csrf-token": csrf },
    ...(body === undefined ? {} : { payload: body as object }),
  });
const ledger = () => store.audit.map((entry) => entry.event);

async function eventually<T>(read: () => Promise<T>, done: (value: T) => boolean): Promise<T> {
  for (let attempt = 0; attempt < 200; attempt += 1) {
    const value = await read();
    if (done(value)) return value;
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
  return read();
}

async function createHook(label = "Synthetic hook") {
  const response = await send("POST", "/v1/alerts/destinations", {
    kind: "generic",
    label,
    url: hookUrl,
  });
  expect(response.statusCode).toBe(201);
  return response.json() as { destination: { id: string }; signingSecret: string };
}

describe("alerts API", () => {
  it("needs a session for every route and the CSRF token for every change", async () => {
    expect((await get("/v1/alerts", "")).statusCode).toBe(401);
    const id = "44444444-4444-4444-8444-444444444444";
    for (const [method, url] of [
      ["POST", "/v1/alerts/destinations"],
      ["DELETE", `/v1/alerts/destinations/${id}`],
      ["POST", `/v1/alerts/destinations/${id}/test`],
      ["POST", "/v1/alerts/rules"],
      ["POST", `/v1/alerts/rules/${id}`],
      ["DELETE", `/v1/alerts/rules/${id}`],
      ["POST", "/v1/alerts/exports"],
      ["POST", "/v1/alerts/weekly"],
    ] as const) {
      const response = await send(method, url, {}, "wrong-token");
      expect(response.statusCode, `${method} ${url}`).toBe(403);
    }
  });

  it("creates a webhook, shows its signing secret once and its address never again", async () => {
    const before = store.audit.length;
    const created = await createHook();
    expect(created.signingSecret).toMatch(/^[A-Za-z0-9_-]{43}$/);
    expect(JSON.stringify(created)).not.toContain("secret-token-value");
    const overview = await get("/v1/alerts");
    expect(overview.statusCode).toBe(200);
    expect(overview.body).not.toContain("secret-token-value");
    expect(overview.body).not.toContain("/workflows/");
    expect(overview.body).not.toContain(created.signingSecret);
    expect(overview.body).not.toContain("synthetic-relay-password");
    expect(overview.body).not.toContain("relay.example");
    expect(overview.json().destinations).toContainEqual(
      expect.objectContaining({ id: created.destination.id, host: "hooks.example" }),
    );
    expect(overview.json().settings).toEqual({
      email: { configured: true, domains: ["example.invalid"] },
      webhookPorts: [],
      temporaryKey: true,
    });
    // The ledger entry was written first, with ids and counts only.
    const entries = store.audit.slice(before).map((entry) => entry.event);
    expect(entries).toEqual([
      expect.objectContaining({
        type: "alerts.destination.created",
        destinationId: created.destination.id,
        kind: "generic",
      }),
    ]);
    expect(JSON.stringify(entries)).not.toContain("hooks.example");
  });

  it("saves nothing when the ledger entry cannot be written first", async () => {
    const append = store.appendAudit.bind(store);
    store.appendAudit = async () => {
      throw new Error("synthetic ledger failure");
    };
    try {
      const destination = await send("POST", "/v1/alerts/destinations", {
        kind: "generic",
        label: "Unrecorded",
        url: hookUrl,
      });
      expect(destination.statusCode).toBe(500);
    } finally {
      store.appendAudit = append;
    }
    const overview = (await get("/v1/alerts")).json() as { destinations: { label: string }[] };
    expect(overview.destinations.map((item) => item.label)).not.toContain("Unrecorded");
  });

  it("refuses unsafe or unusable destinations with a plain message", async () => {
    for (const url of [
      "http://hooks.example/x",
      "https://user:pw@hooks.example/x",
      "https://127.0.0.1/x",
      "https://localhost/x",
    ]) {
      const response = await send("POST", "/v1/alerts/destinations", {
        kind: "teams",
        label: "Bad",
        url,
      });
      expect(response.statusCode, url).toBe(400);
      expect(response.json().error).toBe("invalid_url");
    }
    const port = await send("POST", "/v1/alerts/destinations", {
      kind: "slack",
      label: "Port",
      url: "https://hooks.example:8443/x",
    });
    expect(port.json().error).toBe("port_not_allowed");
    const outside = await send("POST", "/v1/alerts/destinations", {
      kind: "email",
      label: "Outside",
      recipients: ["someone@attacker.example"],
    });
    expect(outside.statusCode).toBe(400);
    expect(outside.json().error).toBe("recipient_not_allowed");
    const invalid = await send("POST", "/v1/alerts/destinations", {
      kind: "email",
      label: "Two in one",
      recipients: ["ops@example.invalid, someone@attacker.example"],
    });
    expect(invalid.statusCode).toBe(400);
    expect(invalid.json().error).toBe("invalid_request");
  });

  it("sends a test and records it before anything leaves", async () => {
    const created = await createHook("Test target");
    const before = store.audit.length;
    const response = await send("POST", `/v1/alerts/destinations/${created.destination.id}/test`);
    expect(response.statusCode).toBe(200);
    expect(response.json().delivery).toMatchObject({ source: "test", status: "sent", attempts: 1 });
    expect(sent).toHaveLength(1);
    expect(JSON.parse(sent[0]?.body ?? "{}")).toMatchObject({ kind: "test" });
    expect(store.audit.slice(before).map((entry) => entry.event)).toEqual([
      expect.objectContaining({
        type: "alerts.delivery",
        source: "test",
        destinationId: created.destination.id,
        items: 0,
      }),
    ]);
  });

  it("creates, turns off and deletes rules, refusing unknown destinations", async () => {
    const created = await createHook("Rules target");
    const unknown = await send("POST", "/v1/alerts/rules", {
      kind: "finding",
      name: "Unknown",
      destinations: ["44444444-4444-4444-8444-444444444444"],
      quietMinutes: 60,
      minimumSeverity: "high",
    });
    expect(unknown.statusCode).toBe(400);
    expect(unknown.json().error).toBe("unknown_destination");
    const made = await send("POST", "/v1/alerts/rules", {
      kind: "finding",
      name: "High findings",
      destinations: [created.destination.id],
      quietMinutes: 60,
      minimumSeverity: "high",
    });
    expect(made.statusCode).toBe(201);
    const rule = made.json().rules.at(-1);
    expect(rule).toMatchObject({ enabled: true, createdBy: actorId, modules: [] });
    const off = await send("POST", `/v1/alerts/rules/${rule.id}`, { enabled: false });
    expect(off.json().rules.find((item: { id: string }) => item.id === rule.id).enabled).toBe(
      false,
    );
    // A destination a rule sends to is not deleted.
    const inUse = await send("DELETE", `/v1/alerts/destinations/${created.destination.id}`);
    expect(inUse.statusCode).toBe(409);
    const removed = await send("DELETE", `/v1/alerts/rules/${rule.id}`);
    expect(removed.json().rules.some((item: { id: string }) => item.id === rule.id)).toBe(false);
    const deleted = await send("DELETE", `/v1/alerts/destinations/${created.destination.id}`);
    expect(deleted.json()).toEqual({ deleted: true });
    expect(ledger().map((event) => event.type)).toEqual(
      expect.arrayContaining([
        "alerts.rule.created",
        "alerts.rule.changed",
        "alerts.rule.deleted",
        "alerts.destination.deleted",
      ]),
    );
    const invalid = await send("POST", "/v1/alerts/rules", {
      kind: "finding",
      name: "Too quiet",
      destinations: [created.destination.id],
      quietMinutes: 1,
      minimumSeverity: "high",
    });
    expect(invalid.statusCode).toBe(400);
  });
});

describe("scheduled exports and the weekly summary", () => {
  it("records each choice in the ledger before saving it", async () => {
    const created = await createHook("Schedule target");
    const before = store.audit.length;
    const exported = await send("POST", "/v1/alerts/exports", {
      auditId: "accounts.inactive",
      destinationId: created.destination.id,
    });
    expect(exported.statusCode).toBe(200);
    expect(exported.json().exports).toContainEqual({
      auditId: "accounts.inactive",
      destinationId: created.destination.id,
    });
    const weekly = await send("POST", "/v1/alerts/weekly", {
      destinationId: created.destination.id,
    });
    expect(weekly.json().weekly).toEqual({
      destinationId: created.destination.id,
      timeZone: "UTC",
    });
    expect(store.audit.slice(before).map((entry) => entry.event.type)).toEqual([
      "alerts.export.changed",
      "alerts.weekly.changed",
    ]);
    const cleared = await send("POST", "/v1/alerts/weekly", { destinationId: null });
    expect(cleared.json().weekly.destinationId).toBeNull();
    await send("POST", "/v1/alerts/exports", { auditId: "accounts.inactive", destinationId: null });
  });

  it("refuses unknown audits and destinations", async () => {
    const unknownAudit = await send("POST", "/v1/alerts/exports", {
      auditId: "audit.unknown",
      destinationId: null,
    });
    expect(unknownAudit.json().error).toBe("unknown_audit");
    const unknownDestination = await send("POST", "/v1/alerts/weekly", {
      destinationId: "44444444-4444-4444-8444-444444444444",
    });
    expect(unknownDestination.json().error).toBe("unknown_destination");
    expect((await send("POST", "/v1/alerts/weekly", { destinationId: "x" })).statusCode).toBe(400);
  });
});

describe("delivery in the running app", () => {
  const credential = (): CollectionCredential => ({
    kind: "delegated",
    tenantId,
    actorId,
    grantedPermissions: async () => ["User.Read", "User.Read.All"],
    accessToken: async () => "synthetic-token",
  });

  it("alerts on a finding a collector reports, through the scheduler's sink", async () => {
    const created = await createHook("Sink target");
    await send("POST", "/v1/alerts/rules", {
      kind: "finding",
      name: "Medium and up",
      destinations: [created.destination.id],
      quietMinutes: 60,
      minimumSeverity: "medium",
      modules: ["test.probe"],
    });
    pending = [finding({ dedupeKey: "sink-1", severity: "high" })];
    scheduler.bind(credential());
    await scheduler.tick();
    const deliveries = await eventually(
      async () =>
        (await get("/v1/alerts")).json().deliveries as { source: string; status: string }[],
      (list) => list.some((item) => item.source === "rule" && item.status === "sent"),
    );
    expect(deliveries.find((item) => item.source === "rule")).toMatchObject({ status: "sent" });
    expect(sent.map((item) => JSON.parse(item.body).kind)).toContain("finding");
    scheduler.unbind("operator_stopped");
  });

  it("keeps collecting while a destination hangs", async () => {
    connect = () => new Promise<never>(() => {});
    pending = [finding({ dedupeKey: "stuck-1", severity: "critical" })];
    scheduler.bind(credential());
    const started = Date.now();
    await scheduler.tick();
    // The run finished and was recorded without waiting for the webhook.
    expect(Date.now() - started).toBeLessThan(5_000);
    const modules = (await get("/v1/modules")).json();
    expect(JSON.stringify(modules)).toContain("succeeded");
    scheduler.unbind("operator_stopped");
  });
});

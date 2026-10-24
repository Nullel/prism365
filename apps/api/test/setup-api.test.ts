import { generateKeyPairSync, randomUUID } from "node:crypto";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  builtInAudits,
  finishRun,
  guestAccountsAudit,
  insertRun,
  licenseReclaimAudit,
} from "@prism365/audits";
import { type SetupCheck, setupStatusSchema } from "@prism365/contracts";
import { GraphHttpClient, permissionPacks } from "@prism365/graph";
import { ModuleScheduler } from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from "vitest";
import { selfSignedCertificate } from "../../../scripts/setup/certificate.ts";
import { createApp } from "../src/app.ts";
import { AuditRunner } from "../src/audits.ts";
import {
  type AppConfig,
  type IdentityConfig,
  loadConfig,
  type SupportedReadFeature,
} from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import { createModuleRegistry } from "../src/modules.ts";
import { probePaths, setupChecks } from "../src/routes/setup.ts";

// Starting collection runs the collectors against the same in-memory database the checklist
// reads, which can take longer than the default five seconds on a busy machine.
vi.setConfig({ testTimeout: 30_000 });

// Synthetic tenant and administrator only.
const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const OTHER_ACTOR = "44444444-4444-4444-8444-444444444444";
const clientId = "33333333-3333-4333-8333-333333333333";
const origin = "http://localhost:5173";
const privateKey = "synthetic-private-key-never-shown";
const DAY = 86_400_000;
const features: SupportedReadFeature[] = [
  "directory.users.read",
  "audit.signins.read",
  "service.health.read",
  "reports.settings.read",
];

function identityConfig(enabledFeatures: SupportedReadFeature[]): IdentityConfig {
  return {
    tenantId,
    clientId,
    origin,
    redirectUri: `${origin}/api/v1/auth/callback`,
    allowedActorIds: [actorId],
    enabledFeatures,
    credential: { clientCertificate: { thumbprintSha256: "AB".repeat(32), privateKey } },
  };
}

function liveConfig(overrides: Partial<AppConfig> = {}): AppConfig {
  return {
    mode: "live",
    origin,
    identity: identityConfig(features),
    production: false,
    missingConfiguration: [],
    enabledFeatures: features,
    certificate: {
      thumbprintSha256: "AB".repeat(32),
      expiresAt: new Date(Date.now() + 200 * DAY).toISOString(),
    },
    ...overrides,
  };
}

// The session grants sign-ins, users and service health; never ReportSettings.Read.All.
const provider: MicrosoftIdentityProvider = {
  authorizationUrl: async () => "https://login.microsoftonline.com/synthetic",
  redeem: async () => ({
    tenantId,
    actorId,
    displayName: "Synthetic admin",
    accountId: "synthetic-account",
    cache: "synthetic-cache",
    scopes: ["User.Read", "User.Read.All", "AuditLog.Read.All", "ServiceHealth.Read.All"],
  }),
  token: async (session) => ({ accessToken: "synthetic", cache: "c", scopes: session.scopes }),
};

async function signIn(identity: IdentityService) {
  const login = await identity.startLogin("directory.users.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
  return {
    cookie: `prism365_session=${loggedIn.handle}`,
    origin,
    "x-csrf-token": loggedIn.session.csrfToken,
  };
}

// What happened, in order: ledger entries and Graph requests side by side.
const events: string[] = [];
let answer: (url: URL) => Response = () => Response.json({ value: [] });
const graphFetch = (async (input: RequestInfo | URL) => {
  const url = new URL(String(input));
  events.push(`graph ${url.pathname}${decodeURIComponent(url.search)}`);
  return answer(url);
}) as typeof fetch;

let app: ReturnType<typeof createApp>;
let store: MemoryStateStore;
let database: PGliteModuleDatabase;
let headers: Record<string, string>;

beforeAll(async () => {
  const config = liveConfig();
  store = new MemoryStateStore();
  const append = store.appendAudit.bind(store);
  store.appendAudit = async (tenant, actor, event) => {
    events.push(`ledger ${String(event.type)}`);
    return append(tenant, actor, event);
  };
  const identity = new IdentityService(identityConfig(features), store, provider);
  headers = await signIn(identity);
  const registry = createModuleRegistry(config);
  database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  const scheduler = new ModuleScheduler({
    enabledPacks: config.enabledFeatures,
    registry,
    database,
    graphClient: (getAccessToken) =>
      new GraphHttpClient({
        getAccessToken,
        fetchImplementation: async () => Response.json({ value: [] }),
      }),
  });
  app = createApp({
    config,
    store,
    identity,
    graphFetch,
    audits: new AuditRunner({ database, timeoutMilliseconds: 1_500 }),
    modules: { registry, database, scheduler },
  });
  events.length = 0;
}, 60_000);

afterAll(async () => {
  await app?.close();
});

afterEach(() => {
  events.length = 0;
  answer = () => Response.json({ value: [] });
});

async function status(url = "/v1/setup/status", extra: Record<string, string> = headers) {
  const response = await app.inject({ method: "GET", url, headers: extra });
  expect(response.statusCode, response.body).toBe(200);
  return { body: response.body, status: setupStatusSchema.parse(response.json()) };
}
const byId = (checks: SetupCheck[], id: string) => checks.find((check) => check.id === id);

describe("GET /v1/setup/status", () => {
  it("tells a caller without a session no more than /v1/session does", async () => {
    const session = (await app.inject({ method: "GET", url: "/v1/session" })).json();
    for (const url of ["/v1/setup/status", "/%76%31/setup/status", "/api/v1/setup/status"]) {
      // No cookie, and a cookie for a session that does not exist (signed out or expired).
      for (const cookie of [undefined, `prism365_session=${"A".repeat(43)}`]) {
        const response = await app.inject({
          method: "GET",
          url,
          ...(cookie ? { headers: { cookie } } : {}),
        });
        expect(response.statusCode, url).toBe(200);
        expect(response.json()).toEqual({
          mode: session.mode,
          configured: session.configured,
          missingConfiguration: session.missingConfiguration,
          checks: [],
          probed: false,
        });
      }
    }
    expect(events).toEqual([]);
  });

  it("tells an administrator taken off the list no more than anyone else", async () => {
    // A session signed in while the administrator was allowed, then an installation whose list
    // no longer has them. Its own store: Prism365 deletes such a session when it sees it.
    const isolated = new MemoryStateStore();
    const signedIn = await signIn(
      new IdentityService(identityConfig(features), isolated, provider),
    );
    const narrowed = { ...identityConfig(features), allowedActorIds: [OTHER_ACTOR] };
    const removed = createApp({
      config: liveConfig({ identity: narrowed }),
      store: isolated,
      identity: new IdentityService(narrowed, isolated, provider),
      graphFetch,
    });
    try {
      const status = await removed.inject({
        method: "GET",
        url: "/v1/setup/status",
        headers: signedIn,
      });
      expect(status.json()).toEqual({
        mode: "live",
        configured: true,
        missingConfiguration: [],
        checks: [],
        probed: false,
      });
      const probe = await removed.inject({
        method: "POST",
        url: "/v1/setup/probe",
        headers: signedIn,
      });
      expect(probe.statusCode).toBe(401);
      expect(events.filter((event) => event.startsWith("graph"))).toEqual([]);
    } finally {
      await removed.close();
    }
  });

  it("names the missing settings of an installation nobody can sign in to yet", async () => {
    const unconfigured = createApp({
      config: liveConfig({
        identity: null,
        missingConfiguration: ["ENTRA_CLIENT_ID", "ENTRA_ADMIN_OBJECT_IDS"],
      }),
    });
    try {
      const response = await unconfigured.inject({ method: "GET", url: "/v1/setup/status" });
      expect(response.json()).toEqual({
        mode: "live",
        configured: false,
        missingConfiguration: ["ENTRA_CLIENT_ID", "ENTRA_ADMIN_OBJECT_IDS"],
        checks: [],
        probed: false,
      });
      const demo = await createApp().inject({ method: "GET", url: "/v1/setup/status" });
      expect(demo.json()).toMatchObject({ mode: "demo", checks: [] });
    } finally {
      await unconfigured.close();
    }
  });

  it("lists every check for a signed-in administrator, without any secret", async () => {
    const { body, status: result } = await status();
    expect(result.probed).toBe(false);
    expect(result.checks.map((check) => check.id)).toEqual([
      "configuration",
      "credential",
      "certificate",
      "signin",
      ...features.map((feature) => `feature.${feature}`),
      "storage.sessions",
      "storage.data",
      "collection",
      ...builtInAudits.map((audit) => `audit.${audit.definition.id}`),
    ]);
    expect(byId(result.checks, "credential")?.state).toBe("ok");
    expect(byId(result.checks, "certificate")).toMatchObject({ state: "ok" });
    expect(byId(result.checks, "feature.audit.signins.read")).toMatchObject({
      state: "ok",
      title: "Sign-in logs (AuditLog.Read.All)",
    });
    expect(byId(result.checks, "feature.reports.settings.read")).toMatchObject({
      state: "warn",
      grantFeature: "reports.settings.read",
    });
    expect(byId(result.checks, "storage.sessions")?.state).toBe("warn");
    expect(byId(result.checks, "storage.data")?.state).toBe("warn");
    expect(byId(result.checks, "collection")?.state).toBe("fail");
    expect(byId(result.checks, "audit.guests.review")).toMatchObject({
      state: "warn",
      detail: "It has never run.",
    });
    expect(body).not.toContain(privateKey);
    expect(body).not.toContain("AB".repeat(32));
    expect(events.filter((event) => event.startsWith("graph"))).toEqual([]);
  });

  it("reports collection and audit runs as they are", async () => {
    const started = await app.inject({ method: "POST", url: "/v1/collection/start", headers });
    expect(started.statusCode, started.body).toBe(200);
    const actor = { id: actorId, name: null };
    const failed = randomUUID();
    const succeeded = randomUUID();
    await database.tenant(tenantId, async (sql) => {
      await insertRun(sql, tenantId, {
        id: failed,
        auditId: guestAccountsAudit.definition.id,
        auditVersion: 1,
        startedAt: new Date(),
        startedBy: actor,
      });
      await finishRun(sql, tenantId, failed, new Date(), {
        status: "failed",
        code: "x",
        message: "x",
      });
      await insertRun(sql, tenantId, {
        id: succeeded,
        auditId: licenseReclaimAudit.definition.id,
        auditVersion: 1,
        startedAt: new Date(),
        startedBy: actor,
      });
      await finishRun(sql, tenantId, succeeded, new Date(), {
        status: "succeeded",
        complete: true,
        summary: [],
        notes: [],
        rows: [],
      });
    });
    const { status: result } = await status();
    expect(byId(result.checks, "collection")).toMatchObject({ state: "ok" });
    expect(byId(result.checks, "collection")?.detail).toContain("with your session");
    expect(byId(result.checks, "audit.licenses.reclaim")?.state).toBe("ok");
    expect(byId(result.checks, "audit.guests.review")).toMatchObject({ state: "warn" });
    expect(byId(result.checks, "audit.guests.review")?.detail).toContain("failed");
    await app.inject({ method: "POST", url: "/v1/collection/stop", headers });
  });

  it("probes each granted pack after writing the ledger, minimally, and explains a refusal", async () => {
    answer = (url) =>
      url.pathname === "/v1.0/auditLogs/signIns"
        ? new Response("{}", { status: 403, headers: { "request-id": "synthetic-request" } })
        : url.pathname === "/v1.0/admin/serviceAnnouncement/healthOverviews"
          ? new Response("{}", { status: 500 })
          : Response.json({ value: [] });
    const response = await app.inject({ method: "POST", url: "/v1/setup/probe", headers });
    expect(response.statusCode, response.body).toBe(200);
    const result = setupStatusSchema.parse(response.json());
    expect(result.probed).toBe(true);
    // One ledger entry, before any read, naming only the packs the session granted.
    expect(events[0]).toBe("ledger setup.probed");
    expect(events.slice(1).sort()).toEqual([
      "graph /v1.0/admin/serviceAnnouncement/healthOverviews?$select=id",
      "graph /v1.0/auditLogs/signIns?$top=1",
      "graph /v1.0/users?$top=1&$select=id",
    ]);
    const entry = store.audit.findLast((item) => item.event.type === "setup.probed");
    expect(entry?.event).toEqual({
      type: "setup.probed",
      packs: ["directory.users.read", "audit.signins.read", "service.health.read"],
      occurredAt: expect.any(String),
    });
    const signins = byId(result.checks, "feature.audit.signins.read");
    const roleNotes = permissionPacks.find((pack) => pack.id === "audit.signins.read")?.roleNotes;
    expect(signins?.state).toBe("fail");
    expect(signins?.detail).toContain("Microsoft refused the test read");
    expect(signins?.detail).toContain(roleNotes);
    expect(byId(result.checks, "feature.directory.users.read")?.detail).toBe(
      "Granted, and Microsoft answered a test read.",
    );
    expect(byId(result.checks, "feature.service.health.read")).toMatchObject({
      state: "unknown",
      detail: "Granted, but the test read did not complete: Microsoft answered HTTP 500.",
    });
    expect(byId(result.checks, "feature.reports.settings.read")?.state).toBe("warn");
  });

  it("makes every probe minimal, and never reads the per-person usage report", async () => {
    for (const [feature, path] of Object.entries(probePaths)) {
      const url = new URL(path({ actorId } as never), "https://graph.microsoft.com/v1.0/");
      const top = url.searchParams.get("$top");
      const filter = url.searchParams.get("$filter");
      expect(top === "1" || filter !== null || !url.search.includes("$top"), feature).toBe(true);
      expect(url.searchParams.has("$expand"), feature).toBe(false);
    }
    expect(probePaths["reports.usage.read"]).toBeUndefined();
  });

  it("reads nothing when the ledger entry cannot be written", async () => {
    const append = store.appendAudit;
    store.appendAudit = async () => {
      throw new Error("ledger unavailable");
    };
    try {
      const response = await app.inject({ method: "POST", url: "/v1/setup/probe", headers });
      expect(response.statusCode).toBe(500);
      expect(events.filter((event) => event.startsWith("graph"))).toEqual([]);
    } finally {
      store.appendAudit = append;
    }
  });

  it("keeps the status read-only and the probe behind the CSRF check", async () => {
    // A link, an image or a top-level navigation is a GET: it can never reach Microsoft.
    for (const url of ["/v1/setup/status?probe=1", "/v1/setup/status?x=1"]) {
      const response = await app.inject({ method: "GET", url, headers });
      expect(response.statusCode, url).toBe(400);
    }
    const { cookie } = headers;
    for (const sent of [{ cookie }, { cookie, origin }, { cookie, "x-csrf-token": "wrong" }]) {
      const response = await app.inject({ method: "POST", url: "/v1/setup/probe", headers: sent });
      expect(response.statusCode).toBe(403);
    }
    expect(events).toEqual([]);
  });
});

describe("certificate expiry", () => {
  const directories: string[] = [];
  afterAll(async () => {
    await Promise.all(directories.map((dir) => rm(dir, { recursive: true, force: true })));
  });

  // A generated certificate through the real configuration loader, so the dates come from the
  // certificate itself.
  async function checklistFor(notBefore: Date, notAfter: Date) {
    const dir = await mkdtemp(join(tmpdir(), "prism365-setup-api-"));
    directories.push(dir);
    const { privateKey: key } = generateKeyPairSync("rsa", { modulusLength: 2048 });
    const certificatePath = join(dir, "entra-public.pem");
    const keyPath = join(dir, "entra-private.pem");
    await writeFile(
      certificatePath,
      selfSignedCertificate(key, { commonName: "Prism365 contoso.example", notBefore, notAfter }),
    );
    await writeFile(keyPath, key.export({ type: "pkcs8", format: "pem" }));
    const config = await loadConfig({
      DATA_MODE: "live",
      ENTRA_TENANT_ID: tenantId,
      ENTRA_CLIENT_ID: clientId,
      ENTRA_ADMIN_OBJECT_IDS: actorId,
      ENABLED_READ_FEATURES: "directory.users.read",
      ENTRA_CERTIFICATE_PATH: certificatePath,
      ENTRA_PRIVATE_KEY_PATH: keyPath,
    });
    const isolated = new MemoryStateStore();
    if (!config.identity) throw new Error("The configuration did not load.");
    const identity = new IdentityService(config.identity, isolated, provider);
    const local = createApp({ config, store: isolated, identity });
    try {
      const response = await local.inject({
        method: "GET",
        url: "/v1/setup/status",
        headers: await signIn(identity),
      });
      const result = setupStatusSchema.parse(response.json());
      expect(response.body).not.toContain("PRIVATE KEY");
      return byId(result.checks, "certificate");
    } finally {
      await local.close();
    }
  }

  it("fails when the certificate has expired, warns under 30 days, and passes otherwise", async () => {
    const now = Date.now();
    const expired = await checklistFor(new Date(now - 400 * DAY), new Date(now - DAY));
    expect(expired).toMatchObject({ state: "fail", title: "The certificate has expired" });
    expect(expired?.fix).toContain("--rotate-certificate");
    const soon = await checklistFor(
      new Date(now - 300 * DAY),
      new Date(now + 10 * DAY + 3_600_000),
    );
    expect(soon).toMatchObject({ state: "warn", title: "The certificate expires soon" });
    expect(soon?.detail).toContain("in 10 days");
    const valid = await checklistFor(new Date(now - DAY), new Date(now + 200 * DAY));
    expect(valid).toMatchObject({ state: "ok", title: "The certificate is valid" });
  }, 30_000);

  it("tells sessions and collected data apart: PRISM_DATA_DIR keeps only the data", () => {
    const checks = setupChecks({
      settings: {
        mode: "live",
        origin,
        production: false,
        configured: true,
        missingConfiguration: [],
        enabledFeatures: [],
        credentialKind: "certificate",
        certificate: null,
        durableStorage: true,
        durableSessions: false,
        handoffRetentionDays: 90,
      },
      moduleDataDurable: true,
      now: new Date(),
      granted: new Set(),
      probes: null,
      collection: null,
      audits: null,
    });
    expect(byId(checks, "storage.sessions")).toMatchObject({ state: "warn" });
    expect(byId(checks, "storage.sessions")?.detail).toContain(
      "PRISM_DATA_DIR does not change that",
    );
    expect(byId(checks, "storage.data")).toMatchObject({ state: "ok" });
  });

  it("warns about a client secret and leaves the certificate check out", () => {
    const checks = setupChecks({
      settings: {
        mode: "live",
        origin,
        production: false,
        configured: true,
        missingConfiguration: [],
        enabledFeatures: [],
        credentialKind: "secret",
        certificate: null,
        durableStorage: true,
        durableSessions: true,
        handoffRetentionDays: 90,
      },
      moduleDataDurable: null,
      now: new Date(),
      granted: new Set(),
      probes: null,
      collection: null,
      audits: null,
    });
    expect(byId(checks, "credential")).toMatchObject({ state: "warn" });
    expect(byId(checks, "certificate")).toBeUndefined();
    expect(byId(checks, "storage.sessions")?.state).toBe("ok");
    expect(byId(checks, "storage.data")).toBeUndefined();
    expect(byId(checks, "collection")?.state).toBe("unknown");
    expect(byId(checks, "audits")?.state).toBe("unknown");
  });
});

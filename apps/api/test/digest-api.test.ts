import { randomUUID } from "node:crypto";
import { digestSchema } from "@prism365/contracts";
import { GraphHttpClient } from "@prism365/graph";
import { createSigninsAuditModule } from "@prism365/module-signins-audit";
import { ModuleRegistry, ModuleScheduler } from "@prism365/modules";
import {
  MemoryStateStore,
  type ModuleDatabase,
  PGliteModuleDatabase,
  type SqlClient,
} from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig, SupportedReadFeature } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import { createModuleRegistry } from "../src/modules.ts";

vi.setConfig({ testTimeout: 30_000 });

// The weekly digest route: the session, the range, the ledger entry written before any read, and
// what is left when part of the installation is off. Synthetic tenant, people and findings only.
const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
const allPacks = [
  "audit.signins.read",
  "audit.directory.read",
  "service.health.read",
  "service.messages.read",
] as const;

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

interface Installation {
  app: ReturnType<typeof createApp>;
  store: MemoryStateStore;
  cookie: string;
  database: PGliteModuleDatabase | null;
  // What the module database was asked, in order, and what the ledger was told.
  events: string[];
}

// One installation: its own state store, and (when `registry` is given) its own database.
async function install(
  enabledFeatures: SupportedReadFeature[],
  registry: ModuleRegistry | null,
): Promise<Installation> {
  const identityConfig: IdentityConfig = {
    tenantId,
    clientId: "33333333-3333-4333-8333-333333333333",
    origin,
    redirectUri: `${origin}/api/v1/auth/callback`,
    allowedActorIds: [actorId],
    enabledFeatures,
    credential: { clientSecret: "synthetic-only" },
  };
  const config: AppConfig = {
    mode: "live",
    origin,
    identity: identityConfig,
    production: false,
    missingConfiguration: [],
    enabledFeatures,
  };
  const store = new MemoryStateStore();
  const events: string[] = [];
  const append = store.appendAudit.bind(store);
  store.appendAudit = async (tenant, actor, event) => {
    events.push(`ledger:${String(event.type)}`);
    await append(tenant, actor, event);
  };
  const identity = new IdentityService(identityConfig, store, provider);
  const login = await identity.startLogin("audit.signins.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
  const database = registry
    ? await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() })
    : null;
  // The same database, noting every transaction it is asked to open.
  const watched: ModuleDatabase | null = database && {
    kind: database.kind,
    durable: database.durable,
    tenant: (id, operation, options) => {
      events.push("read");
      return database.tenant(id, operation, options);
    },
    verify: () => database.verify(),
    close: () => database.close(),
  };
  const scheduler =
    registry && watched
      ? new ModuleScheduler({
          enabledPacks: enabledFeatures,
          registry,
          database: watched,
          graphClient: (getAccessToken) =>
            new GraphHttpClient({
              getAccessToken,
              fetchImplementation: async () => Response.json({ value: [] }),
            }),
        })
      : null;
  const app = createApp({
    config,
    store,
    identity,
    ...(registry && watched && scheduler
      ? { modules: { registry, database: watched, scheduler } }
      : {}),
  });
  return { app, store, cookie: `prism365_session=${loggedIn.handle}`, database, events };
}

const get = (target: Installation, url: string, cookie: string | null = target.cookie) =>
  target.app.inject({ method: "GET", url, headers: cookie ? { cookie } : {} });

let full: Installation;
let signinsOnly: Installation;
let notEnabled: Installation;
let noStorage: Installation;
const installations: Installation[] = [];

const ago = (milliseconds: number) => new Date(Date.now() - milliseconds).toISOString();
const HOUR = 3_600_000;
const DAY = 24 * HOUR;

async function seedFinding(
  database: PGliteModuleDatabase,
  values: { module: string; rule: string; title: string; severity: string; firstSeen: string },
) {
  const id = randomUUID();
  await database.tenant(tenantId, (sql: SqlClient) =>
    sql.query(
      `INSERT INTO prism365_finding (id, tenant_id, module, rule_id, rule_version, dedupe_key,
         purpose, severity, score, subject_kind, subject_id, subject_label, title, summary,
         evidence, limitations, window_start, window_end, state, occurrences, first_seen, last_seen)
       VALUES ($1,$2,$3,$4,1,$5,'security',$6,50,'person','person-1','Rowan Ellis',$7,'Synthetic.',
         '[]'::jsonb,'[]'::jsonb,$8,$8,'open',1,$8,$8)`,
      [
        id,
        tenantId,
        values.module,
        values.rule,
        id,
        values.severity,
        values.title,
        values.firstSeen,
      ],
    ),
  );
}

beforeAll(async () => {
  full = await install([...allPacks], createModuleRegistry());
  signinsOnly = await install(
    ["audit.signins.read"],
    new ModuleRegistry().register(createSigninsAuditModule()),
  );
  notEnabled = await install(["audit.signins.read"], createModuleRegistry());
  noStorage = await install([...allPacks], null);
  installations.push(full, signinsOnly, notEnabled, noStorage);
  if (!full.database) throw new Error("The full installation has no database.");
  await seedFinding(full.database, {
    module: "signins.audit",
    rule: "signins.unfamiliar_signin",
    title: "An unfamiliar sign-in",
    severity: "high",
    firstSeen: ago(2 * HOUR),
  });
  await seedFinding(full.database, {
    module: "signins.audit",
    rule: "signins.unfamiliar_signin",
    title: "An older sign-in finding",
    severity: "low",
    firstSeen: ago(20 * DAY),
  });
}, 120_000);

afterAll(async () => {
  for (const installation of installations) {
    // Closing the app closes the module database it was given.
    await installation.app.close();
  }
});

describe("GET /v1/digest", () => {
  it("requires a session, in live mode and in demo mode, before anything is read or recorded", async () => {
    const before = full.events.length;
    expect((await get(full, "/v1/digest", null)).statusCode).toBe(401);
    expect(
      (await get(full, "/v1/digest?days=7", "prism365_session=not-a-session")).statusCode,
    ).toBe(401);
    const demo = createApp();
    expect((await demo.inject({ method: "GET", url: "/v1/digest" })).statusCode).toBe(401);
    await demo.close();
    expect(full.events.length).toBe(before);
  });

  it("answers the last 7 days by default, and the other ranges it offers", async () => {
    const response = await get(full, "/v1/digest");
    expect(response.statusCode).toBe(200);
    const body = digestSchema.parse(response.json());
    expect(body.days).toBe(7);
    expect(Date.parse(body.generatedAt) - Date.parse(body.from)).toBe(7 * DAY);
    expect(body.signInFindings.state).toBe("ok");
    expect(body.signInFindings.opened.high).toBe(1);
    expect(body.signInFindings.openTotal).toBe(2);
    expect(body.signInFindings.open[0]).toMatchObject({
      title: "An unfamiliar sign-in",
      subjectLabel: "Rowan Ellis",
      personId: "person-1",
    });
    for (const [days, opened] of [
      [1, 1],
      [7, 1],
      [14, 1],
      [30, 2],
    ] as const) {
      const ranged = digestSchema.parse((await get(full, `/v1/digest?days=${days}`)).json());
      expect(ranged.days, `days=${days}`).toBe(days);
      expect(
        Object.values(ranged.signInFindings.opened).reduce((a, b) => a + b, 0),
        `days=${days}`,
      ).toBe(days === 30 ? 2 : opened);
    }
  });

  it("refuses any other range with a plain message, before the ledger or the database", async () => {
    const before = full.events.length;
    for (const query of [
      "days=0",
      "days=2",
      "days=31",
      "days=-7",
      "days=7.5",
      "days=",
      "days=abc",
      "days=7&days=14",
      "days=%20",
    ]) {
      const response = await get(full, `/v1/digest?${query}`);
      expect(response.statusCode, query).toBe(400);
      expect(response.json()).toEqual({
        error: "invalid_request",
        message: "Choose 1, 7, 14 or 30 days.",
      });
    }
    expect(full.events.length).toBe(before);
  });

  it("writes the ledger entry before it reads, naming the sections and never a person", async () => {
    full.events.length = 0;
    const response = await get(full, "/v1/digest?days=14");
    expect(response.statusCode).toBe(200);
    // One entry, then reads: the first thing recorded is the ledger, and no read comes first.
    expect(full.events[0]).toBe("ledger:digest.viewed");
    expect(full.events.filter((event) => event.startsWith("ledger"))).toEqual([
      "ledger:digest.viewed",
    ]);
    expect(full.events.indexOf("read")).toBeGreaterThan(0);
    const entry = full.store.audit.at(-1);
    expect(entry).toMatchObject({
      tenantId,
      actorId,
      event: {
        type: "digest.viewed",
        datasets: [
          "signInFindings",
          "directoryChanges",
          "quietAccounts",
          "auditDrift",
          "microsoft",
          "collection",
        ],
        days: 14,
      },
    });
    expect(typeof entry?.event.occurredAt).toBe("string");
    // Counts and section names only: no names, titles or ids from the data.
    const text = JSON.stringify(entry);
    for (const value of ["Rowan", "person-1", "unfamiliar"]) expect(text).not.toContain(value);
  });

  it("reads nothing when the ledger cannot record the read", async () => {
    const reads = full.events.filter((event) => event === "read").length;
    const append = full.store.appendAudit.bind(full.store);
    full.store.appendAudit = async () => {
      throw new Error("Synthetic ledger failure.");
    };
    try {
      const response = await get(full, "/v1/digest");
      expect(response.statusCode).toBe(500);
      expect(response.body).not.toContain("Rowan");
    } finally {
      full.store.appendAudit = append;
    }
    expect(full.events.filter((event) => event === "read").length).toBe(reads);
  });

  it("returns the sections it can when a module is not installed, and records only those", async () => {
    signinsOnly.events.length = 0;
    const response = await get(signinsOnly, "/v1/digest");
    expect(response.statusCode).toBe(200);
    const body = digestSchema.parse(response.json());
    expect(body.directoryChanges).toMatchObject({
      state: "unavailable",
      note: "The directory change feed is not available in this installation.",
    });
    expect(body.microsoft).toMatchObject({ state: "unavailable" });
    expect(body.microsoft.incidents.note).toBe(
      "Service health is not available in this installation.",
    );
    expect(body.microsoft.posts.note).toBe("Message center is not available in this installation.");
    // Audit results and run history are core data.
    expect(body.quietAccounts.state).toBe("unavailable");
    expect(body.auditDrift.state).toBe("unavailable");
    expect(body.collection.state).toBe("ok");
    expect(signinsOnly.store.audit.at(-1)?.event).toMatchObject({
      type: "digest.viewed",
      datasets: ["signInFindings", "quietAccounts", "auditDrift", "collection"],
    });
  });

  it("names the pack to enable when a module is installed but its read pack is not", async () => {
    const body = digestSchema.parse((await get(notEnabled, "/v1/digest")).json());
    expect(body.directoryChanges.note).toBe(
      "This installation has not enabled audit.directory.read. Add it to ENABLED_READ_FEATURES and restart.",
    );
    expect(body.microsoft.incidents.note).toContain("service.health.read");
    expect(body.microsoft.posts.note).toContain("service.messages.read");
    // The sign-in pack is on; nothing has been collected yet, which is a different reason.
    expect(body.signInFindings.note).not.toContain("ENABLED_READ_FEATURES");
    expect(notEnabled.store.audit.at(-1)?.event.datasets).toEqual([
      "signInFindings",
      "quietAccounts",
      "auditDrift",
      "collection",
    ]);
  });

  it("answers with every section unavailable, and records no read, without module storage", async () => {
    noStorage.events.length = 0;
    const response = await get(noStorage, "/v1/digest?days=30");
    expect(response.statusCode).toBe(200);
    const body = digestSchema.parse(response.json());
    expect(body.days).toBe(30);
    for (const section of [
      body.signInFindings,
      body.directoryChanges,
      body.quietAccounts,
      body.auditDrift,
      body.microsoft,
      body.collection,
    ]) {
      expect(section.state).toBe("unavailable");
      expect(section.note).toContain("Module storage is not configured");
    }
    expect(body.collection.collecting).toBe(false);
    expect(noStorage.events).toEqual([]);
  });

  it("reports whether collection is on", async () => {
    const body = digestSchema.parse((await get(full, "/v1/digest")).json());
    expect(body.collection.collecting).toBe(false);
    expect(body.collection.note).toContain("Collection is off");
  });

  it("is a GET only", async () => {
    const response = await full.app.inject({
      method: "POST",
      url: "/v1/digest",
      headers: { cookie: full.cookie, origin },
      payload: {},
    });
    expect([403, 404, 405]).toContain(response.statusCode);
  });
});

import {
  builtInAudits,
  finishRun,
  insertRun,
  readAuditScope,
  readAuditSettings,
  writeAuditScope,
} from "@prism365/audits";
import { GraphHttpClient } from "@prism365/graph";
import { ModuleScheduler } from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { createApp } from "../src/app.ts";
import { AuditRunner } from "../src/audits.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import { createModuleRegistry } from "../src/modules.ts";

// The audit scope: which account domains the account audits cover. Synthetic tenant,
// administrator and accounts only; every domain ends in .example.
const tenantId = "11111111-1111-4111-8111-111111111111";
const otherTenantId = "99999999-9999-4999-8999-999999999999";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
const features = [
  "directory.users.read",
  "audit.signins.read",
  "reports.authentication-methods.read",
] as const;
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

const daysAgo = (days: number) => new Date(Date.now() - days * 86_400_000).toISOString();
const E3 = "11111111-0000-4000-8000-0000000000e3";
const users = [
  {
    id: "u-a1",
    displayName: "Ada Alpha",
    userPrincipalName: "ada@a.example",
    userType: "Member",
    accountEnabled: true,
    createdDateTime: daysAgo(900),
    assignedLicenses: [{ skuId: E3 }],
    signInActivity: { lastSuccessfulSignInDateTime: daysAgo(200) },
  },
  {
    id: "u-a2",
    displayName: "Abe Alpha",
    userPrincipalName: "abe@a.example",
    userType: "Member",
    accountEnabled: true,
    createdDateTime: daysAgo(900),
    assignedLicenses: [],
    signInActivity: { lastSuccessfulSignInDateTime: daysAgo(3) },
  },
  {
    id: "u-b1",
    displayName: "Bo Beta",
    userPrincipalName: "bo@b.example",
    userType: "Member",
    accountEnabled: false,
    createdDateTime: daysAgo(900),
    assignedLicenses: [{ skuId: E3 }],
    signInActivity: { lastSuccessfulSignInDateTime: daysAgo(10) },
  },
  {
    id: "g-1",
    displayName: "Gus Guest",
    userPrincipalName: "gus_vendor.example#EXT#@tenant.example",
    mail: "gus@vendor.example",
    userType: "Guest",
    accountEnabled: true,
    createdDateTime: daysAgo(400),
    externalUserState: "PendingAcceptance",
    assignedLicenses: [],
    signInActivity: null,
  },
];
const registrations = [
  {
    id: "u-a1",
    userPrincipalName: "ada@a.example",
    userDisplayName: "Ada Alpha",
    isAdmin: true,
    isMfaRegistered: false,
    methodsRegistered: [],
    userType: "member",
  },
  {
    id: "u-b1",
    userPrincipalName: "bo@b.example",
    userDisplayName: "Bo Beta",
    isAdmin: false,
    isMfaRegistered: false,
    methodsRegistered: [],
    userType: "member",
  },
];

interface Session {
  cookie: string;
  csrf: string;
}

// One app on its own in-memory database, so a group of tests cannot see another's runs.
async function harness(withModules = true) {
  let gate: Promise<void> | null = null;
  let refuseSignInDates = false;
  const graphFetch = (async (input: RequestInfo | URL) => {
    const url = new URL(String(input));
    if (url.pathname === "/v1.0/users") {
      if (gate) await gate;
      if (refuseSignInDates && url.searchParams.get("$select")?.includes("signInActivity"))
        return new Response("{}", { status: 403 });
      return Response.json({ value: users });
    }
    if (url.pathname === "/v1.0/reports/authenticationMethods/userRegistrationDetails")
      return Response.json({ value: registrations });
    return new Response("{}", { status: 404 });
  }) as typeof fetch;
  const store = new MemoryStateStore();
  const identity = new IdentityService(identityConfig, store, provider);
  const login = await identity.startLogin("directory.users.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
  const who: Session = {
    cookie: `prism365_session=${loggedIn.handle}`,
    csrf: (await identity.session(loggedIn.handle))?.csrfToken ?? "",
  };
  const registry = createModuleRegistry(config);
  const database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  const scheduler = new ModuleScheduler({
    enabledPacks: config.enabledFeatures,
    registry,
    database,
    graphClient: (getAccessToken) => new GraphHttpClient({ getAccessToken }),
  });
  const runner = new AuditRunner({ database, timeoutMilliseconds: 5_000 });
  const app = createApp({
    config,
    store,
    identity,
    graphFetch,
    audits: runner,
    ...(withModules ? { modules: { registry, database, scheduler } } : {}),
  });
  const get = (url: string, session: Session | null = who) =>
    app.inject({ method: "GET", url, headers: session ? { cookie: session.cookie } : {} });
  const post = (url: string, body?: unknown, session: Session | null = who, csrf = true) =>
    app.inject({
      method: "POST",
      url,
      headers: {
        ...(session ? { cookie: session.cookie, origin } : {}),
        ...(session && csrf ? { "x-csrf-token": session.csrf } : {}),
      },
      ...(body === undefined ? {} : { payload: body as object }),
    });
  const save = (domains: unknown) => post("/v1/settings/audits", { domains });
  const start = async (auditId: string) => {
    const response = await post(`/v1/audits/${auditId}/runs`);
    expect(response.statusCode, auditId).toBe(202);
    return response.json().run.id as string;
  };
  // Runs an audit to the end and returns its stored result.
  const finish = async (auditId: string) => {
    const id = await start(auditId);
    await runner.idle();
    const detail = (await get(`/v1/audit-runs/${id}`)).json() as {
      status: string;
      notes: string[];
      rows: { key: string; cells: Record<string, unknown> }[];
      summary: { label: string; value: number }[];
    };
    return { id, ...detail };
  };
  return {
    app,
    store,
    database,
    runner,
    who,
    get,
    post,
    save,
    start,
    finish,
    hold: () => {
      let release = () => {};
      gate = new Promise<void>((resolve) => {
        release = resolve;
      });
      return () => {
        release();
        gate = null;
      };
    },
    refuse: (value: boolean) => {
      refuseSignInDates = value;
    },
  };
}

type Harness = Awaited<ReturnType<typeof harness>>;
const keys = (result: { rows: { key: string }[] }) => result.rows.map((row) => row.key).sort();
const settings = async (h: Harness) => (await h.get("/v1/settings/audits")).json();

describe("audit scope settings", () => {
  let h: Harness;
  beforeAll(async () => {
    h = await harness();
  }, 60_000);
  afterAll(async () => {
    await h?.app.close();
  });

  it("needs a session, and the CSRF header to save", async () => {
    expect((await h.get("/v1/settings/audits", null)).statusCode).toBe(401);
    expect((await h.post("/v1/settings/audits", { domains: [] }, null)).statusCode).toBe(401);
    const noCsrf = await h.post("/v1/settings/audits", { domains: ["a.example"] }, h.who, false);
    expect(noCsrf.statusCode).toBe(403);
    expect((await settings(h)).scope).toEqual({ domains: [] });
  });

  it("refuses a body that is not a list of domain names", async () => {
    const bodies: unknown[] = [
      undefined,
      {},
      { domains: "a.example" },
      { domains: [1] },
      { domains: ["not a domain"] },
      { domains: ["ada@a.example"] },
      { domains: Array.from({ length: 101 }, (_, index) => `d${index}.example`) },
    ];
    for (const body of bodies) {
      const response = await h.post("/v1/settings/audits", body);
      expect(response.statusCode, JSON.stringify(body)).toBe(400);
      expect(response.json()).toMatchObject({ error: "invalid_request" });
    }
    expect((await settings(h)).scope).toEqual({ domains: [] });
    expect(h.store.audit.some((entry) => entry.event.type === "audits.scope.changed")).toBe(false);
  });

  it("starts with every domain and nothing to choose from", async () => {
    expect(await settings(h)).toEqual({ scope: { domains: [] }, domains: [] });
  });

  it("lists each domain once with the accounts the latest runs found", async () => {
    const inactive = await h.finish("accounts.inactive");
    expect(inactive.status).toBe("succeeded");
    expect((await settings(h)).domains).toEqual([
      { domain: "a.example", accounts: 2 },
      { domain: "b.example", accounts: 1 },
      { domain: "tenant.example", accounts: 1 },
    ]);
    // The same accounts in the other scoped audits are not counted twice.
    await h.finish("licenses.reclaim");
    await h.finish("mfa.registration");
    // A guest audit holds a Home domain column, but it does not use the scope: not listed.
    const guests = await h.finish("guests.review");
    expect(guests.rows[0]?.cells.domain).toBe("vendor.example");
    expect((await settings(h)).domains).toEqual([
      { domain: "a.example", accounts: 2 },
      { domain: "b.example", accounts: 1 },
      { domain: "tenant.example", accounts: 1 },
    ]);
  });

  it("saves the scope in lower case, once each, and records a count in the ledger", async () => {
    const before = h.store.audit.length;
    const saved = await h.save([" B.Example ", "a.example", "A.EXAMPLE"]);
    expect(saved.statusCode).toBe(200);
    expect(saved.json().scope).toEqual({ domains: ["a.example", "b.example"] });
    expect(saved.json().domains).toHaveLength(3);
    expect((await settings(h)).scope).toEqual({ domains: ["a.example", "b.example"] });
    const changed = h.store.audit
      .slice(before)
      .filter((entry) => entry.event.type === "audits.scope.changed");
    expect(changed).toHaveLength(1);
    expect(changed[0]?.event).toEqual({
      type: "audits.scope.changed",
      domains: 2,
      occurredAt: expect.any(String),
    });
    expect(changed[0]?.actorId).toBe(actorId);
    // A count, never the names of the domains.
    expect(JSON.stringify(changed[0])).not.toContain("example");
  });

  it("applies the saved scope to the next run of each scoped audit", async () => {
    await h.save(["a.example"]);
    const inactive = await h.finish("accounts.inactive");
    expect(keys(inactive)).toEqual(["u-a1", "u-a2"]);
    expect(inactive.notes[0]).toBe("Limited to accounts in: a.example.");
    expect(inactive.summary.find((tile) => tile.label === "Accounts")?.value).toBe(2);
    expect(inactive.rows.every((row) => row.cells.domain === "a.example")).toBe(true);
    expect(keys(await h.finish("licenses.reclaim"))).toEqual(["u-a1"]);
    expect(keys(await h.finish("mfa.registration"))).toEqual(["u-a1"]);
    // Guests do not use the scope.
    const guests = await h.finish("guests.review");
    expect(keys(guests)).toEqual(["g-1"]);
    expect(guests.notes.join(" ")).not.toContain("Limited to");
  });

  it("keeps every domain on offer and marks a limited result, after a scoped run", async () => {
    await h.save(["a.example"]);
    const run = await h.finish("accounts.inactive");
    // The chooser still lists the domains the scope leaves out, from the run's own inventory.
    const offered = (await settings(h)).domains.map((item: { domain: string }) => item.domain);
    expect(offered).toEqual(expect.arrayContaining(["a.example", "b.example"]));
    const stored = (await h.get(`/v1/audit-runs/${run.id}`)).json();
    expect(stored.scope).toEqual(["a.example"]);
    expect(stored.notes.join(" ")).toContain("left out by the audit scope");
    const csv = await h.get(`/v1/audit-runs/${run.id}/export.csv`);
    expect(csv.headers["content-disposition"]).toContain("-limited-to-1-domain.csv");
  });

  it("does not change the scope when the change cannot be recorded", async () => {
    const before = (await settings(h)).scope;
    const append = h.store.appendAudit.bind(h.store);
    h.store.appendAudit = async (tenant, actor, event) => {
      if (event.type === "audits.scope.changed") throw new Error("ledger unavailable");
      return append(tenant, actor, event);
    };
    const failed = await h.save(["b.example"]);
    h.store.appendAudit = append;
    expect(failed.statusCode).toBe(500);
    expect((await settings(h)).scope).toEqual(before);
  });

  it("goes back to every domain when the list is cleared", async () => {
    const cleared = await h.save([]);
    expect(cleared.statusCode).toBe(200);
    expect(cleared.json().scope).toEqual({ domains: [] });
    const inactive = await h.finish("accounts.inactive");
    expect(keys(inactive)).toEqual(["g-1", "u-a1", "u-a2", "u-b1"]);
    expect(inactive.notes.join(" ")).not.toContain("Limited to");
    const changed = h.store.audit.filter((entry) => entry.event.type === "audits.scope.changed");
    expect(changed.at(-1)?.event).toMatchObject({ domains: 0 });
  });

  it("reads the scope once when a run starts", async () => {
    await h.save(["a.example"]);
    const release = h.hold();
    const id = await h.start("accounts.inactive");
    // Saved while the run waits on Microsoft: this run keeps the scope it started with.
    await h.save(["b.example"]);
    release();
    await h.runner.idle();
    const detail = (await h.get(`/v1/audit-runs/${id}`)).json();
    expect(keys(detail)).toEqual(["u-a1", "u-a2"]);
    expect(detail.notes[0]).toBe("Limited to accounts in: a.example.");
    await h.save([]);
  });
});

describe("audit scope settings without module storage", () => {
  it("answers 409 like the other audit routes", async () => {
    const h = await harness(false);
    try {
      const get = await h.get("/v1/settings/audits");
      expect(get.statusCode).toBe(409);
      expect(get.json()).toMatchObject({ error: "audits_unavailable" });
      expect((await h.save(["a.example"])).statusCode).toBe(409);
      expect((await h.get("/v1/settings/audits", null)).statusCode).toBe(401);
    } finally {
      await h.app.close();
    }
  }, 60_000);
});

describe("audit scope domains and storage", () => {
  let h: Harness;
  beforeAll(async () => {
    h = await harness();
  }, 60_000);
  afterAll(async () => {
    await h?.app.close();
  });

  it("keeps listing domains from the latest successful run when a newer run failed", async () => {
    await h.finish("accounts.inactive");
    h.refuse(true);
    const failed = await h.finish("accounts.inactive");
    h.refuse(false);
    expect(failed.status).toBe("failed");
    expect((await settings(h)).domains.map((item: { domain: string }) => item.domain)).toEqual([
      "a.example",
      "b.example",
      "tenant.example",
    ]);
  });

  it("uses only the latest successful run of an audit", async () => {
    await h.database.tenant(tenantId, async (sql) => {
      const id = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
      await insertRun(sql, tenantId, {
        id,
        auditId: "accounts.inactive",
        auditVersion: 2,
        startedAt: new Date(Date.now() + 60_000),
        startedBy: { id: actorId, name: null },
      });
      await finishRun(sql, tenantId, id, new Date(Date.now() + 61_000), {
        status: "succeeded",
        complete: true,
        summary: [],
        notes: [],
        rows: [
          {
            key: "only",
            personId: null,
            tags: [],
            cells: {
              person: { name: "Only One", principal: "one@z.example" },
              domain: "z.example",
            },
          },
          // A row from an older version of the audit has no Domain cell: not listed.
          {
            key: "old",
            personId: null,
            tags: [],
            cells: { person: { name: "Old", principal: null } },
          },
        ],
      });
    });
    expect((await settings(h)).domains).toEqual([
      // Reclaim and MFA have not run on this database; only the newest inactive run counts.
      { domain: "z.example", accounts: 1 },
    ]);
  });

  it("keeps each tenant's scope apart", async () => {
    await h.database.tenant(otherTenantId, (sql) =>
      writeAuditScope(sql, otherTenantId, { domains: ["other.example"] }, actorId),
    );
    await h.save(["a.example"]);
    const mine = await h.database.tenant(tenantId, (sql) => readAuditScope(sql, tenantId));
    const theirs = await h.database.tenant(otherTenantId, (sql) =>
      readAuditScope(sql, otherTenantId),
    );
    expect(mine).toEqual({ domains: ["a.example"] });
    expect(theirs).toEqual({ domains: ["other.example"] });
    // Another tenant's settings do not list this tenant's domains either.
    const otherSettings = await h.database.tenant(otherTenantId, (sql) =>
      readAuditSettings(sql, otherTenantId, builtInAudits),
    );
    expect(otherSettings).toEqual({ scope: { domains: ["other.example"] }, domains: [] });
    await h.save([]);
  });

  it("reads a damaged saved scope as every domain", async () => {
    await h.database.tenant(tenantId, (sql) =>
      sql.query(
        `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
         VALUES ($1,'audits.scope','{"domains":"not a list"}'::jsonb,now(),$2)
         ON CONFLICT (tenant_id, key) DO UPDATE SET value=EXCLUDED.value`,
        [tenantId, actorId],
      ),
    );
    expect((await settings(h)).scope).toEqual({ domains: [] });
    await h.save([]);
  });
});

import { finishRun, insertRun } from "@prism365/audits";
import type { AuditRow } from "@prism365/contracts";
import { GraphHttpClient } from "@prism365/graph";
import { ModuleScheduler } from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import { createModuleRegistry } from "../src/modules.ts";

vi.setConfig({ testTimeout: 30_000 });

// The person page's Access tab. Synthetic tenant, administrator and accounts only; every address
// ends in .invalid.
const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
const person = "aaaaaaaa-0000-4000-8000-000000000001";
const base = `/v1/people/${person}/access`;

const GROUP_M365 = "bbbbbbbb-0000-4000-8000-000000000001";
const GROUP_SECURITY = "bbbbbbbb-0000-4000-8000-000000000002";
const GROUP_NESTED = "bbbbbbbb-0000-4000-8000-000000000003";
const ROLE_OBJECT = "cccccccc-0000-4000-8000-0000000000a1";
const ROLE_TEMPLATE = "cccccccc-0000-4000-8000-0000000000b1";
const HIDDEN = "bbbbbbbb-0000-4000-8000-0000000000ff";

const ALL = ["directory.users.read", "identity.conditional-access.read", "audit.signins.read"];
const SCOPES = ["User.Read", "User.Read.All", "Policy.Read.All", "AuditLog.Read.All"];

const group = (id: string, name: string, over: Record<string, unknown> = {}) => ({
  "@odata.type": "#microsoft.graph.group",
  id,
  displayName: name,
  groupTypes: [],
  mailEnabled: false,
  securityEnabled: true,
  ...over,
});
const transitive = [
  group(GROUP_M365, "Quality team", { groupTypes: ["Unified"], mailEnabled: true }),
  group(GROUP_SECURITY, "Plant operators"),
  group(GROUP_NESTED, "All plant staff"),
  {
    "@odata.type": "#microsoft.graph.directoryRole",
    id: ROLE_OBJECT,
    displayName: "Helpdesk Administrator",
    roleTemplateId: ROLE_TEMPLATE,
  },
  {
    "@odata.type": "#microsoft.graph.administrativeUnit",
    id: "dddddddd-0000-4000-8000-000000000001",
    displayName: "Plant site",
  },
  // Without Directory.Read.All Microsoft returns only the type and the id.
  { "@odata.type": "#microsoft.graph.directoryObject", id: HIDDEN },
];
const direct = [{ id: GROUP_M365 }, { id: GROUP_SECURITY }, { id: ROLE_OBJECT }];

const policy = (id: string, name: string, users: object, state = "enabled", grant = ["mfa"]) => ({
  id,
  displayName: name,
  state,
  conditions: { users },
  grantControls: { operator: "OR", builtInControls: grant },
});
const policies = [
  policy("p-all", "Require MFA for everyone", { includeUsers: ["All"], excludeUsers: [person] }),
  policy(
    "p-nested",
    "Managed device for plant staff",
    { includeGroups: [GROUP_NESTED] },
    "enabledForReportingButNotEnforced",
    ["compliantDevice"],
  ),
  policy("p-role", "Block legacy for helpdesk", { includeRoles: [ROLE_TEMPLATE] }, "disabled", [
    "block",
  ]),
  policy("p-other", "Contractors only", {
    includeGroups: ["bbbbbbbb-0000-4000-8000-0000000000ee"],
  }),
  { id: "p-bad", displayName: 42, state: 7, conditions: "nonsense" },
];

type Handler = (url: URL) => Response | undefined;
async function harness(
  options: {
    features?: readonly string[];
    scopes?: readonly string[];
    handler?: Handler;
    modules?: boolean;
  } = {},
) {
  const features = options.features ?? ALL;
  const scopes = options.scopes ?? SCOPES;
  const requests: string[] = [];
  let ledgerAtFirstRequest: number | null = null;
  let store: MemoryStateStore | null = null;
  const handler: Handler =
    options.handler ??
    ((url) => {
      const path = url.pathname.replace("/v1.0/", "");
      if (path === `users/${person}`)
        return Response.json({
          id: person,
          userPrincipalName: "Pat.Pell@example.invalid",
          userType: "Member",
        });
      if (path === `users/${person}/transitiveMemberOf`)
        return Response.json({ value: transitive });
      if (path === `users/${person}/memberOf`) return Response.json({ value: direct });
      if (path === `users/${person}/ownedObjects`)
        return Response.json({
          value: [
            {
              "@odata.type": "#microsoft.graph.group",
              id: "o1",
              displayName: "Plant operators",
              resourceProvisioningOptions: [],
            },
            {
              "@odata.type": "#microsoft.graph.group",
              id: "o2",
              displayName: "Plant team",
              resourceProvisioningOptions: ["Team"],
            },
            {
              "@odata.type": "#microsoft.graph.application",
              id: "o3",
              displayName: "Plant dashboard",
            },
            { "@odata.type": "#microsoft.graph.device", id: "o4", displayName: "Not shown" },
          ],
        });
      if (path === "identity/conditionalAccess/policies") return Response.json({ value: policies });
      return undefined;
    });
  const graphFetch = (async (input: RequestInfo | URL) => {
    const url = new URL(String(input));
    requests.push(`${url.pathname}${url.search}`);
    if (ledgerAtFirstRequest === null) ledgerAtFirstRequest = store?.audit.length ?? 0;
    return handler(url) ?? new Response("{}", { status: 404 });
  }) as typeof fetch;
  const identityConfig: IdentityConfig = {
    tenantId,
    clientId: "33333333-3333-4333-8333-333333333333",
    origin,
    redirectUri: `${origin}/api/v1/auth/callback`,
    allowedActorIds: [actorId],
    enabledFeatures: features as IdentityConfig["enabledFeatures"],
    credential: { clientSecret: "synthetic-only" },
  };
  const config: AppConfig = {
    mode: "live",
    origin,
    identity: identityConfig,
    production: false,
    missingConfiguration: [],
    enabledFeatures: features as AppConfig["enabledFeatures"],
  };
  const provider: MicrosoftIdentityProvider = {
    authorizationUrl: async () => "https://login.microsoftonline.com/synthetic",
    redeem: async () => ({
      tenantId,
      actorId,
      displayName: "Synthetic admin",
      accountId: "synthetic-account",
      cache: "synthetic-cache",
      scopes: [...scopes],
    }),
    token: async (session) => ({ accessToken: "synthetic", cache: "c", scopes: session.scopes }),
  };
  store = new MemoryStateStore();
  const identity = new IdentityService(identityConfig, store, provider);
  const login = await identity.startLogin(features[0] as "directory.users.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
  const cookie = `prism365_session=${loggedIn.handle}`;
  const registry = createModuleRegistry(config);
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
    ...(options.modules === false ? {} : { modules: { registry, database, scheduler } }),
  });
  const get = (url: string, signedIn = true) =>
    app.inject({ method: "GET", url, headers: signedIn ? { cookie } : {} });
  const ledger = () => store.audit.map((entry) => entry.event as Record<string, unknown>);
  // A finished run of one audit, as the audit store keeps it.
  let runs = 0;
  const storeRun = async (
    auditId: string,
    rows: AuditRow[],
    scope: string[] | null = null,
    outcome: { notes?: string[]; complete?: boolean } = {},
  ) => {
    runs += 1;
    const id = `00000000-0000-4000-8000-${String(runs).padStart(12, "0")}`;
    await database.tenant(tenantId, async (sql) => {
      await insertRun(sql, tenantId, {
        id,
        auditId,
        auditVersion: 1,
        startedAt: new Date("2026-10-20T10:00:00Z"),
        startedBy: { id: actorId, name: null },
        scope,
      });
      await finishRun(sql, tenantId, id, new Date("2026-10-20T10:05:00Z"), {
        status: "succeeded",
        complete: outcome.complete ?? true,
        summary: [],
        notes: outcome.notes ?? [],
        rows,
      });
    });
  };
  const seedSignin = (upn: string) =>
    database.tenant(tenantId, (sql) =>
      sql.query(
        `INSERT INTO mod_signins_audit_event (tenant_id, id, created_at, user_id, upn, error_code,
           risk_event_types, backfilled, processed, collected_at)
         VALUES ($1,$2,$3,$4,$5,0,'[]'::jsonb,false,true,$3)`,
        [tenantId, `event-${upn}`, "2026-10-20T10:00:00.000Z", person, upn],
      ),
    );
  return {
    app,
    get,
    ledger,
    requests,
    storeRun,
    seedSignin,
    ledgerAtFirstRequest: () => ledgerAtFirstRequest,
  };
}

const row = (key: string, personId: string | null, cells: AuditRow["cells"]): AuditRow => ({
  key,
  personId,
  tags: [],
  cells,
});

async function body(h: Awaited<ReturnType<typeof harness>>, url = base) {
  const response = await h.get(url);
  expect(response.statusCode).toBe(200);
  return response.json();
}

describe("person access API", () => {
  let h: Awaited<ReturnType<typeof harness>>;
  beforeAll(async () => {
    h = await harness();
  }, 60_000);
  afterAll(async () => {
    await h?.app.close();
  });

  it("needs a session and a usable account id", async () => {
    expect((await h.get(base, false)).statusCode).toBe(401);
    expect((await h.get("/v1/people/../access")).statusCode).not.toBe(200);
    expect((await h.get("/v1/people/.../access")).statusCode).toBe(400);
    expect((await h.get("/v1/people/has%20space/access")).statusCode).toBe(400);
    expect(h.requests).toEqual([]);
  });

  it("writes the ledger entry before it reads anything, with counts and no names", async () => {
    const before = h.ledger().length;
    await body(h);
    expect(h.ledgerAtFirstRequest()).toBeGreaterThan(before);
    const entry = h.ledger().find((event) => event.type === "person.access.viewed");
    expect(entry).toMatchObject({
      datasets: ["directory", "conditionalAccess", "auditResults"],
      sections: 3,
    });
    expect(JSON.stringify(entry)).not.toContain(person);
    expect(JSON.stringify(entry)).not.toContain("Pat.Pell");
  });

  it("lists memberships by kind, direct or nested, and keeps what Microsoft hides", async () => {
    const { memberships } = await body(h);
    expect(memberships.status).toBe("ok");
    const byId = new Map<string, Record<string, unknown>>(
      memberships.entries.map((entry: { id: string }) => [entry.id, entry]),
    );
    expect(byId.get(GROUP_M365)).toMatchObject({
      kind: "microsoft365",
      relation: "direct",
      hasGroupMailbox: true,
    });
    expect(byId.get(GROUP_SECURITY)).toMatchObject({
      kind: "security",
      relation: "direct",
      hasGroupMailbox: false,
    });
    expect(byId.get(GROUP_NESTED)).toMatchObject({ kind: "security", relation: "nested" });
    expect(byId.get(ROLE_OBJECT)).toMatchObject({
      kind: "role",
      roleTemplateId: ROLE_TEMPLATE,
      relation: "direct",
    });
    expect(byId.get("dddddddd-0000-4000-8000-000000000001")).toMatchObject({
      kind: "administrative-unit",
      relation: "nested",
    });
    expect(byId.get(HIDDEN)).toMatchObject({ kind: "hidden", name: null });
    expect(memberships.entries).toHaveLength(6);
    const request = h.requests.find((path) => path.includes("transitiveMemberOf")) ?? "";
    expect(decodeURIComponent(request)).toContain(
      "$select=id,displayName,groupTypes,mailEnabled,securityEnabled,membershipRule,roleTemplateId,description",
    );
  });

  it("lists what the person owns, groups, teams and applications only", async () => {
    const { owned } = await body(h);
    // A device is not part of this page, and leaving it out is not a cut list.
    expect(owned.truncated).toBe(false);
    expect(owned.entries).toEqual([
      { id: "o1", kind: "group", name: "Plant operators" },
      { id: "o2", kind: "team", name: "Plant team" },
      { id: "o3", kind: "application", name: "Plant dashboard" },
    ]);
  });

  it("says which Conditional Access policies apply, and which leave the person out", async () => {
    const { conditionalAccess } = await body(h);
    expect(conditionalAccess).toMatchObject({
      status: "ok",
      unaffected: 1,
      unreadable: 1,
      membershipsKnown: true,
    });
    const byId = new Map<string, Record<string, unknown>>(
      conditionalAccess.policies.map((entry: { id: string }) => [entry.id, entry]),
    );
    // An exclusion wins over "all users".
    expect(byId.get("p-all")).toMatchObject({
      effect: "excluded",
      targeted: true,
      state: "on",
      controls: "Require MFA",
      excludedBy: [{ kind: "person", id: person, name: null }],
    });
    // Through a nested group, named from the memberships; report-only is shown.
    expect(byId.get("p-nested")).toMatchObject({
      effect: "applies",
      state: "report-only",
      includedBy: [{ kind: "group", id: GROUP_NESTED, name: "All plant staff" }],
    });
    // Through a role held now; off is shown.
    expect(byId.get("p-role")).toMatchObject({
      effect: "applies",
      state: "off",
      controls: "Block access",
      includedBy: [{ kind: "role", id: ROLE_TEMPLATE, name: "Helpdesk Administrator" }],
    });
    expect(byId.has("p-other")).toBe(false);
  });

  it("finds no stored results for an audit that never ran", async () => {
    const { audits } = await body(h);
    expect(audits.status).toBe("ok");
    expect(audits.results).toHaveLength(9);
    expect(audits.results.every((result: { status: string }) => result.status === "no_run")).toBe(
      true,
    );
  });
});

describe("person access API, stored audit results", () => {
  it("picks the person's rows by id, labels the cells and caps the rows at 25", async () => {
    const h = await harness();
    const rows = [
      row("other", "aaaaaaaa-0000-4000-8000-0000000000ff", { strength: "Phone" }),
      ...Array.from({ length: 30 }, (_value, index) =>
        row(`mine-${index}`, person, {
          strength: "Authenticator app",
          registered: true,
          methods: [],
          // Columns the page does not show are not sent.
          preferred: "push",
          admin: false,
        }),
      ),
    ];
    await h.storeRun("mfa.registration", rows);
    const { audits } = await body(h);
    const mfa = audits.results.find(
      (result: { auditId: string }) => result.auditId === "mfa.registration",
    );
    expect(mfa).toMatchObject({
      status: "ok",
      matchedBy: "account",
      finishedAt: "2026-10-20T10:05:00.000Z",
      scope: null,
      truncated: true,
    });
    expect(mfa.rows).toHaveLength(25);
    expect(mfa.rows[0].key).toBe("mine-0");
    expect(mfa.rows[0].cells).toEqual([
      { key: "registered", label: "MFA registered", kind: "boolean", value: true },
      { key: "strength", label: "Strongest method", kind: "enum", value: "Authenticator app" },
    ]);
    expect(JSON.stringify(mfa)).not.toContain('"other"');
    await h.app.close();
  }, 60_000);

  it("matches mailbox rows by sign-in name, normalised, from stored sign-ins", async () => {
    const h = await harness();
    await h.seedSignin("pat.pell@example.invalid");
    await h.storeRun("mail.mailbox-storage", [
      // Full-width letters and capitals fold to the same name.
      row("pat", null, { mailbox: "ＰＡＴ.Pell@Example.invalid", used: 3.1 }),
      row("lee", null, { mailbox: "lee.park@example.invalid", used: 9 }),
    ]);
    const { audits } = await body(h);
    const mailbox = audits.results.find(
      (result: { auditId: string }) => result.auditId === "mail.mailbox-storage",
    );
    expect(mailbox).toMatchObject({ status: "ok", matchedBy: "sign-in name" });
    expect(mailbox.rows.map((entry: { key: string }) => entry.key)).toEqual(["pat"]);
    await h.app.close();
  }, 60_000);

  it("never matches a mailbox on a hashed name", async () => {
    const hashed = "0123456789abcdef0123456789abcdef";
    const h = await harness({
      handler: (url) =>
        url.pathname.endsWith(`/users/${person}`)
          ? Response.json({ id: person, userPrincipalName: hashed, userType: "Member" })
          : undefined,
    });
    await h.storeRun("mail.mailbox-storage", [row("hash", null, { mailbox: hashed, used: 1 })]);
    const { audits } = await body(h);
    const mailbox = audits.results.find(
      (result: { auditId: string }) => result.auditId === "mail.mailbox-storage",
    );
    expect(mailbox.rows).toEqual([]);
    await h.app.close();
  }, 60_000);

  it("says plainly when the sign-in name is not known, instead of guessing", async () => {
    const h = await harness({ features: ["audit.signins.read"], scopes: ["User.Read"] });
    await h.storeRun("mail.mailbox-storage", [
      row("lee", null, { mailbox: "lee@example.invalid" }),
    ]);
    const { audits } = await body(h);
    const mailbox = audits.results.find(
      (result: { auditId: string }) => result.auditId === "mail.mailbox-storage",
    );
    expect(mailbox.status).toBe("failed");
    expect(mailbox.message).toContain("sign-in name");
    await h.app.close();
  }, 60_000);
});

describe("person access API, sections that cannot be read", () => {
  it("names the pack and permission when the session was not granted them", async () => {
    const h = await harness({ scopes: ["User.Read"] });
    const result = await body(h);
    expect(result.memberships).toEqual({
      status: "not_granted",
      pack: "directory.users.read",
      permission: "User.Read.All",
      enabled: true,
    });
    expect(result.owned.status).toBe("not_granted");
    expect(result.conditionalAccess).toEqual({
      status: "not_granted",
      pack: "identity.conditional-access.read",
      permission: "Policy.Read.All",
      enabled: true,
    });
    expect(result.audits.status).toBe("ok");
    expect(h.requests).toEqual([]);
    expect(h.ledger().find((event) => event.type === "person.access.viewed")).toMatchObject({
      datasets: ["auditResults"],
      sections: 1,
    });
    await h.app.close();
  }, 60_000);

  it("says when the installation has not enabled the pack", async () => {
    const h = await harness({ features: ["audit.signins.read"], scopes: SCOPES });
    const result = await body(h);
    expect(result.memberships).toMatchObject({ status: "not_granted", enabled: false });
    expect(result.conditionalAccess).toMatchObject({ status: "not_granted", enabled: false });
    await h.app.close();
  }, 60_000);

  it("reports a failed read in plain words and still answers the other sections", async () => {
    const h = await harness({
      handler: (url) => {
        const path = url.pathname.replace("/v1.0/", "");
        if (path.endsWith("/transitiveMemberOf")) return new Response("{}", { status: 500 });
        if (path === "identity/conditionalAccess/policies")
          return new Response("{}", { status: 403 });
        if (path.endsWith("/ownedObjects")) return Response.json({ value: [] });
        return undefined;
      },
    });
    const result = await body(h);
    expect(result.memberships).toEqual({
      status: "failed",
      message: "Microsoft could not answer this read just now. Try again.",
    });
    expect(result.conditionalAccess.status).toBe("failed");
    expect(result.conditionalAccess.message).toContain("refused");
    expect(result.owned).toMatchObject({ status: "ok", entries: [] });
    await h.app.close();
  }, 60_000);

  it("still matches Conditional Access on everyone when memberships cannot be read", async () => {
    const h = await harness({
      handler: (url) => {
        const path = url.pathname.replace("/v1.0/", "");
        if (path.endsWith("/transitiveMemberOf")) return new Response("{}", { status: 500 });
        if (path === "identity/conditionalAccess/policies")
          return Response.json({ value: policies });
        if (path === `users/${person}`) return Response.json({ id: person, userType: "Member" });
        return undefined;
      },
    });
    const { conditionalAccess } = await body(h);
    expect(conditionalAccess.membershipsKnown).toBe(false);
    expect(conditionalAccess.policies.map((entry: { id: string }) => entry.id)).toEqual(["p-all"]);
    await h.app.close();
  }, 60_000);

  it("explains that audit results need module storage", async () => {
    const h = await harness({ modules: false });
    const { audits } = await body(h);
    expect(audits.status).toBe("failed");
    await h.app.close();
  }, 60_000);
});

describe("person access API, limits", () => {
  it("reads at most 1,000 memberships and says the list was cut", async () => {
    const make = (from: number, to: number) =>
      Array.from({ length: to - from }, (_value, index) =>
        group(
          `eeeeeeee-0000-4000-8000-${String(from + index).padStart(12, "0")}`,
          `Group ${from + index}`,
        ),
      );
    const h = await harness({
      handler: (url) => {
        const path = url.pathname.replace("/v1.0/", "");
        if (path === `users/${person}/memberOf`) return Response.json({ value: [] });
        if (path !== `users/${person}/transitiveMemberOf`) return undefined;
        if (url.searchParams.get("$skiptoken") === "2")
          return Response.json({ value: make(600, 1_001) });
        const next = new URL(url);
        next.searchParams.set("$skiptoken", "2");
        return Response.json({ value: make(0, 600), "@odata.nextLink": next.href });
      },
    });
    const { memberships } = await body(h);
    expect(memberships.entries).toHaveLength(1_000);
    expect(memberships.truncated).toBe(true);
    await h.app.close();
  }, 60_000);
  it("reads at most 200 owned objects and says the list was cut", async () => {
    const h = await harness({
      handler: (url) => {
        const path = url.pathname.replace("/v1.0/", "");
        if (path !== `users/${person}/ownedObjects`) return undefined;
        return Response.json({
          value: Array.from({ length: 201 }, (_value, index) =>
            group(`ffffffff-0000-4000-8000-${String(index).padStart(12, "0")}`, `Owned ${index}`),
          ),
        });
      },
    });
    const { owned } = await body(h);
    expect(owned.entries).toHaveLength(200);
    expect(owned.truncated).toBe(true);
    await h.app.close();
  }, 60_000);

  it("calls an object hidden only when it has neither a type nor a name", async () => {
    const h = await harness({
      handler: (url) => {
        const path = url.pathname.replace("/v1.0/", "");
        if (path === `users/${person}/memberOf`) return Response.json({ value: [] });
        if (path !== `users/${person}/transitiveMemberOf`) return undefined;
        return Response.json({
          value: [
            { id: "bare-1" },
            { "@odata.type": "#microsoft.graph.group", id: "typed-1", securityEnabled: true },
            { id: "named-1", displayName: "Named only" },
          ],
        });
      },
    });
    const { memberships } = await body(h);
    const byId = new Map<string, Record<string, unknown>>(
      memberships.entries.map((entry: { id: string }) => [entry.id, entry]),
    );
    expect(byId.get("bare-1")).toMatchObject({ kind: "hidden", name: null });
    expect(byId.get("typed-1")).toMatchObject({ kind: "security", name: null });
    expect(byId.get("named-1")).toMatchObject({ kind: "hidden" });
    await h.app.close();
  }, 60_000);
});

describe("person access API, combined risks", () => {
  const combinedRow = (personId: string, tags: string[], reasons: string): AuditRow => ({
    key: personId,
    personId,
    tags,
    cells: { person: { name: "Pat Pell", principal: "pat@example.invalid" }, reasons },
  });

  it("gives the person's rules, highest severity, reasons and the run's time", async () => {
    const h = await harness();
    await h.storeRun("identity.combined-risks", [
      combinedRow(
        "aaaaaaaa-0000-4000-8000-0000000000ff",
        ["critical", "admin-unprotected"],
        "Not me.",
      ),
      combinedRow(
        person,
        ["critical", "dormant-admin", "admin-unprotected", "privileged"],
        "Holds Global Administrator (permanent) and has no MFA registered. Has not signed in for 90 days.",
      ),
    ]);
    const { combinedRisks } = await body(h);
    expect(combinedRisks).toEqual({
      status: "ok",
      finishedAt: "2026-10-20T10:05:00.000Z",
      severity: "critical",
      // In the rule list's order, whatever order the tags came in.
      rules: [
        { id: "admin-unprotected", label: "Administrator without MFA", severity: "critical" },
        { id: "dormant-admin", label: expect.any(String), severity: "high" },
      ],
      reasons:
        "Holds Global Administrator (permanent) and has no MFA registered. Has not signed in for 90 days.",
      gaps: [],
      complete: true,
    });
    await h.app.close();
  }, 60_000);

  it("says the person trips nothing when the run has no row for them", async () => {
    const h = await harness();
    await h.storeRun("identity.combined-risks", [
      combinedRow("aaaaaaaa-0000-4000-8000-0000000000ff", ["high", "dormant-admin"], "Not me."),
    ]);
    const { combinedRisks } = await body(h);
    expect(combinedRisks).toMatchObject({ status: "ok", severity: null, rules: [], reasons: "" });
    await h.app.close();
  }, 60_000);

  it("says there is no run when the audit has never finished one", async () => {
    const h = await harness();
    expect((await body(h)).combinedRisks).toEqual({ status: "no_run" });
    await h.app.close();
  }, 60_000);

  it("carries the run's notes about missing or stale inputs, and that it was incomplete", async () => {
    const h = await harness();
    await h.storeRun("identity.combined-risks", [], null, {
      complete: false,
      notes: [
        "Inputs used, with the date each audit ran: Privileged roles (20 Oct 2026).",
        "Not used: MFA registration last ran 1 Aug 2026, more than 14 days ago. Two audits are never joined across such a gap.",
        "Administrator without MFA was not checked. Run MFA registration.",
        "Conditional Access: 2 enabled policies require MFA.",
      ],
    });
    const { combinedRisks } = await body(h);
    expect(combinedRisks.complete).toBe(false);
    expect(combinedRisks.gaps).toEqual([
      "Not used: MFA registration last ran 1 Aug 2026, more than 14 days ago. Two audits are never joined across such a gap.",
      "Administrator without MFA was not checked. Run MFA registration.",
    ]);
    await h.app.close();
  }, 60_000);

  it("reads only the stored run, and leaves names and addresses out of the ledger", async () => {
    const h = await harness();
    await h.storeRun("identity.combined-risks", [
      combinedRow(person, ["medium", "paying-for-nothing"], "Pays for a license."),
    ]);
    const before = h.requests.length;
    await body(h);
    const entry = h.ledger().find((event) => event.type === "person.access.viewed");
    expect(JSON.stringify(entry)).not.toContain("Pat Pell");
    expect(JSON.stringify(entry)).not.toContain("pat@example.invalid");
    expect(h.requests.slice(before).some((path) => path.includes("combined"))).toBe(false);
    await h.app.close();
  }, 60_000);

  it("explains that audit results need module storage", async () => {
    const h = await harness({ modules: false });
    expect((await body(h)).combinedRisks.status).toBe("failed");
    await h.app.close();
  }, 60_000);

  it("puts the same answer on the person record for the header, only when a run exists", async () => {
    const h = await harness();
    const none = await body(h, `/v1/people/${person}`);
    expect(none).not.toHaveProperty("combinedRisks");
    await h.storeRun("identity.combined-risks", [
      combinedRow(person, ["critical", "admin-unprotected"], "Holds Global Administrator."),
    ]);
    const record = await body(h, `/v1/people/${person}`);
    expect(record.combinedRisks).toMatchObject({
      status: "ok",
      severity: "critical",
      rules: [{ id: "admin-unprotected" }],
    });
    expect(h.ledger().find((event) => event.type === "person.viewed")).toMatchObject({
      datasets: ["signIns", "findings", "combinedRisks"],
    });
    await h.app.close();
  }, 60_000);
});

import { finishRun, insertRun } from "@prism365/audits";
import type { AuditRow, Entity, EntitySearch } from "@prism365/contracts";
import { GraphHttpClient } from "@prism365/graph";
import { ModuleScheduler } from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase, type SqlClient } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import { createModuleRegistry } from "../src/modules.ts";

vi.setConfig({ testTimeout: 60_000 });

// The panel for a group, role, policy, device or app, and the search that finds one. Synthetic
// tenant, administrator, people and objects only; every address ends in .invalid.
const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
const guid = (n: number) => `bbbbbbbb-0000-4000-8000-${n.toString(16).padStart(12, "0")}`;
const ana = guid(1);
const ben = guid(2);
const cai = guid(3);
const PLANT = guid(10);
const LAB = guid(11);
const STAFF = guid(12);
const ADMINS = guid(13);
const GLOBAL_ADMIN = "62e90394-69f5-4237-9190-012177145e10";
const LAB_ROLE = guid(20);
const POLICY = guid(30);
const OTHER_POLICY = guid(31);
const DEVICE = guid(40);
const APP = guid(50);
const APP_ID = guid(51);
const features = ["directory.users.read", "audit.signins.read", "audit.directory.read"] as const;

const row = (
  key: string,
  personId: string | null,
  cells: AuditRow["cells"],
  tags: string[] = [],
): AuditRow => ({ key, personId, tags, cells });
const ownerItem = (id: string, name: string | null, kind: string) =>
  JSON.stringify({ id, name, kind });

async function harness(withModules = true) {
  const store = new MemoryStateStore();
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
  const requests: string[] = [];
  const graphFetch = (async (input: RequestInfo | URL) => {
    requests.push(String(input));
    return new Response("{}", { status: 404 });
  }) as typeof fetch;
  const identity = new IdentityService(identityConfig, store, provider);
  const login = await identity.startLogin("directory.users.read");
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
    ...(withModules ? { modules: { registry, database, scheduler } } : {}),
  });
  const get = (path: string, signedIn = true) =>
    app.inject({ method: "GET", url: path, headers: signedIn ? { cookie } : {} });
  let runs = 0;
  const storeRun = async (
    auditId: string,
    rows: AuditRow[],
    outcome: { complete?: boolean; finished?: string; version?: number } = {},
  ) => {
    runs += 1;
    const id = `00000000-0000-4000-8000-${String(runs).padStart(12, "0")}`;
    await database.tenant(tenantId, async (sql) => {
      await insertRun(sql, tenantId, {
        id,
        auditId,
        auditVersion: outcome.version ?? 2,
        startedAt: new Date("2026-10-20T10:00:00Z"),
        startedBy: { id: actorId, name: null },
      });
      await finishRun(sql, tenantId, id, new Date(outcome.finished ?? "2026-10-20T10:05:00Z"), {
        status: "succeeded",
        complete: outcome.complete ?? true,
        summary: [],
        notes: [],
        rows,
      });
    });
  };
  const sql = <T>(work: (client: SqlClient) => Promise<T>) => database.tenant(tenantId, work);
  return { app, get, store, requests, storeRun, sql, database };
}

type Harness = Awaited<ReturnType<typeof harness>>;
const ledger = (h: Harness) => h.store.audit.map((entry) => entry.event as Record<string, unknown>);
const url = (kind: string, id: string) => `/v1/entities/${kind}/${id}`;
async function entity(h: Harness, kind: string, id: string): Promise<Entity> {
  const response = await h.get(url(kind, id));
  expect(response.statusCode).toBe(200);
  return response.json();
}
const sectionOf = (found: Entity, id: string) => {
  const section = found.sections.find((candidate) => candidate.id === id);
  if (!section) throw new Error(`No section ${id}`);
  return section;
};
function okRows(found: Entity, id: string) {
  const section = sectionOf(found, id);
  if (section.status !== "ok") throw new Error(`Section ${id} is ${section.status}`);
  return section.rows;
}

// Every audit an entity page reads, with one row each for the things the tests look at.
async function storeEverything(h: Harness) {
  await h.storeRun("groups.teams", [
    row(
      PLANT,
      null,
      {
        group: "Plant team",
        kind: "Team",
        status: "One owner",
        visibility: "Public",
        owners: 3,
        ownerIds: [ana, APP, cai],
        ownerList: [
          ownerItem(ana, "Ana Ames", "user"),
          ownerItem(APP, "Plant robot", "servicePrincipal"),
          ownerItem(cai, null, "other"),
        ],
        roleAssignable: true,
      },
      ["team", "one-owner"],
    ),
    row(LAB, null, { group: "Lab 100%", kind: "Security group", status: "OK", owners: null }),
    row(STAFF, null, { group: "Staff", kind: "Security group", status: "OK" }),
    row(ADMINS, null, { group: "Admins_one", kind: "Security group", status: "OK" }),
  ]);
  await h.storeRun("directory.privileged-roles", [
    row(
      `permanent|${ana}|${GLOBAL_ADMIN}|/`,
      ana,
      {
        holder: { name: "Ana Ames", principal: "ana@example.invalid" },
        holderType: "User",
        role: "Global Administrator",
        privileged: true,
        held: "Permanent",
        through: "Direct",
        scope: "Tenant",
        ends: null,
        mfa: false,
        roleId: GLOBAL_ADMIN,
        holderId: ana,
        throughId: null,
      },
      ["privileged"],
    ),
    row(
      `permanent|${PLANT}|${GLOBAL_ADMIN}|/`,
      null,
      {
        holder: { name: PLANT, principal: null },
        holderType: "Group or other",
        role: "Global Administrator",
        privileged: true,
        held: "Permanent",
        through: "Direct",
        scope: "Tenant",
        roleId: GLOBAL_ADMIN,
        holderId: PLANT,
        throughId: null,
      },
      ["privileged"],
    ),
    row(
      `permanent|${APP}|${GLOBAL_ADMIN}|/`,
      null,
      {
        holder: { name: "Plant robot", principal: null },
        holderType: "Service principal",
        role: "Global Administrator",
        privileged: true,
        held: "Eligible",
        through: "Direct",
        scope: "Tenant",
        roleId: GLOBAL_ADMIN,
        holderId: APP,
        throughId: null,
      },
      ["privileged"],
    ),
    row(
      `active|${ben}|${LAB_ROLE}|/`,
      ben,
      {
        holder: { name: "Ben Birch", principal: "ben@example.invalid" },
        holderType: "User",
        role: "Lab Reports Reader",
        privileged: false,
        held: "Active",
        through: "Group",
        scope: "Tenant",
        roleId: LAB_ROLE,
        holderId: ben,
        throughId: STAFF,
      },
      [],
    ),
  ]);
  await h.storeRun("identity.conditional-access", [
    row(POLICY, null, {
      policy: "Require MFA for admins",
      state: "On",
      appliesTo: "1 group",
      excludes: ["1 group"],
      apps: "All apps",
      grant: "Require MFA",
      includeUserIds: [ana, ben, cai],
      excludeUserIds: [],
      includeGroupIds: [PLANT.toUpperCase().toLowerCase(), guid(99)],
      excludeGroupIds: [STAFF],
      includeRoleIds: [GLOBAL_ADMIN, guid(98)],
      excludeRoleIds: [LAB_ROLE],
    }),
    row(OTHER_POLICY, null, {
      policy: "Block legacy",
      state: "Report only",
      includeGroupIds: [],
      excludeGroupIds: [PLANT],
      includeRoleIds: [],
      excludeRoleIds: [],
    }),
  ]);
  await h.storeRun("devices.managed", [
    row(
      DEVICE,
      ana,
      {
        device: "Plant laptop",
        person: { name: "Ana Ames", principal: "ana@example.invalid" },
        os: "Windows",
        compliance: "Not compliant",
        ownership: "Corporate",
        serial: "SYN-0001",
      },
      ["not-compliant"],
    ),
    row(guid(41), null, {
      device: "Spare tablet",
      person: { name: "No primary user", principal: null },
      os: "iOS",
    }),
  ]);
  await h.storeRun("apps.consents", [
    row(APP, null, {
      app: "Plant robot",
      publisher: "Unverified",
      owner: "Another organisation",
      enabled: true,
      highRisk: ["Directory.ReadWrite.All"],
      appId: APP_ID,
    }),
    row(guid(52), null, { app: "Plant dashboard", publisher: "Plant Works", appId: guid(53) }),
  ]);
  await h.storeRun("identity.combined-risks", [
    row(
      "r-ana",
      ana,
      { person: { name: "Ana Ames", principal: "ana@example.invalid" }, reasons: "No MFA." },
      ["critical", "admin-unprotected"],
    ),
    row("r-ben", ben, { person: { name: "Ben Birch", principal: null } }, []),
  ]);
  await h.storeRun("accounts.activity", [
    row("a1", ana, { person: { name: "Ana Ames", principal: "ana@example.invalid" } }),
    row("a2", ben, { person: { name: "Ben Birch", principal: "ben@example.invalid" } }),
  ]);
}

describe("entity API", () => {
  let h: Harness;
  beforeAll(async () => {
    h = await harness();
  }, 60_000);
  afterAll(async () => {
    await h?.app.close();
  });

  it("needs a session, a known kind and a GUID, and says plainly when there is no module storage", async () => {
    expect((await h.get(url("group", PLANT), false)).statusCode).toBe(401);
    expect((await h.get(url("planet", PLANT))).statusCode).toBe(400);
    expect((await h.get(url("group", "not-a-guid"))).statusCode).toBe(400);
    expect((await h.get(url("group", `${PLANT}x`))).statusCode).toBe(400);
    expect((await h.get("/v1/entities/search?q=ab", false)).statusCode).toBe(401);
    for (const bad of ["", "?q=a", `?q=${"x".repeat(101)}`, "?q=%20%20%20"])
      expect((await h.get(`/v1/entities/search${bad}`)).statusCode, bad).toBe(400);
    expect(ledger(h).filter((event) => String(event.type).startsWith("entity."))).toEqual([]);
    const bare = await harness(false);
    expect((await bare.get(url("group", PLANT))).statusCode).toBe(409);
    expect((await bare.get("/v1/entities/search?q=ab")).statusCode).toBe(409);
    await bare.app.close();
  });

  it("writes the ledger entry first, with the kind and no id or name, and asks Microsoft nothing", async () => {
    let at: number | null = null;
    const original = h.database.tenant.bind(h.database);
    h.database.tenant = ((...args: Parameters<typeof original>) => {
      at ??= h.store.audit.length;
      return original(...args);
    }) as typeof original;
    const before = h.store.audit.length;
    await entity(h, "device", DEVICE);
    expect(at).toBeGreaterThan(before);
    at = null;
    await h.get("/v1/entities/search?q=plant");
    h.database.tenant = original;
    expect(at).toBeGreaterThan(before + 1);
    const events = ledger(h).filter((event) => String(event.type).startsWith("entity."));
    expect(events.map((event) => event.type)).toEqual(["entity.viewed", "entity.searched"]);
    expect(events[0]).toMatchObject({ kind: "device", datasets: ["auditResults"] });
    expect(JSON.stringify(events)).not.toContain(DEVICE);
    expect(JSON.stringify(events)).not.toContain("plant");
    expect(h.requests).toEqual([]);
  });

  it("says not run for every source before anything has run, and has no name", async () => {
    for (const [kind, id, sections] of [
      ["group", PLANT, 3],
      // Combined risks only has a section for a role that has holders.
      ["role", GLOBAL_ADMIN, 2],
      ["policy", POLICY, 1],
      ["device", DEVICE, 1],
      ["app", APP, 1],
    ] as const) {
      const found = await entity(h, kind, id);
      expect(found).toMatchObject({ kind, id, tags: [] });
      expect(found.name).toBe(kind === "role" ? "Global Administrator" : null);
      expect(found.sections).toHaveLength(sections);
      expect(found.sections.map((section) => section.status)).toEqual(
        Array(sections).fill("no_run"),
      );
    }
    // A role still has a name when it is one of the built-in administrator roles.
    expect((await entity(h, "role", GLOBAL_ADMIN)).name).toBe("Global Administrator");
    expect((await entity(h, "role", LAB_ROLE)).name).toBeNull();
    expect(await (await h.get("/v1/entities/search?q=plant")).json()).toEqual({
      results: [],
      more: false,
      partial: false,
    });
  });

  it("shows a group with its owners as people, the policies that reach it and the roles it gives", async () => {
    await storeEverything(h);
    const found = await entity(h, "group", PLANT.toUpperCase());
    expect(found).toMatchObject({
      id: PLANT,
      name: "Plant team",
      tags: ["Team", "One owner", "Public", "Role-assignable"],
    });
    expect(found.notes.join(" ")).toMatch(/Members are not read/);
    expect(found.notes.join(" ")).toMatch(/Access tab/);
    const [group] = okRows(found, "group");
    expect(group?.cells.map((cell) => cell.key)).toEqual([
      "kind",
      "status",
      "visibility",
      "owners",
      "roleAssignable",
    ]);
    expect(group?.cells.some((cell) => cell.key === "ownerIds")).toBe(false);
    expect(group?.lists).toEqual([
      {
        label: "Owners",
        items: [
          { kind: "person", id: ana, name: "Ana Ames", principal: null },
          { kind: "app", id: APP, name: "Plant robot", principal: null },
          { kind: "other", id: cai, name: null, principal: null },
        ],
        more: 0,
      },
    ]);
    const policies = okRows(found, "policies");
    expect(policies.map((item) => [item.subject?.name, item.tags])).toEqual([
      ["Require MFA for admins", ["Includes it"]],
      ["Block legacy", ["Leaves it out"]],
    ]);
    const roles = okRows(found, "roles");
    expect(roles).toHaveLength(1);
    expect(roles[0]).toMatchObject({
      subject: { kind: "role", id: GLOBAL_ADMIN, name: "Global Administrator" },
      tags: ["It holds the role"],
    });
    // A group that others reach a role through: the holders are listed.
    const staff = await entity(h, "group", STAFF);
    expect(okRows(staff, "roles")[0]).toMatchObject({
      subject: { kind: "role", id: LAB_ROLE },
      tags: ["Held through it"],
      lists: [{ label: "Holder", items: [{ kind: "person", id: ben, name: "Ben Birch" }] }],
    });
    expect(okRows(staff, "policies")[0]?.tags).toEqual(["Leaves it out"]);
  });

  it("says a section is not in the latest run, and keeps the others", async () => {
    const found = await entity(h, "group", guid(77));
    expect(found.name).toBeNull();
    expect(found.sections.map((section) => section.status)).toEqual(["not_in_run", "ok", "ok"]);
    const first = found.sections[0];
    expect(first).toMatchObject({ complete: true, auditName: "Groups and Teams" });
    expect(first?.status === "not_in_run" && first.finishedAt).toBe("2026-10-20T10:05:00.000Z");
    // "ok" with a plain line, not a table, when nothing reaches the group.
    const second = sectionOf(found, "policies");
    expect(second.status === "ok" && second.rows).toEqual([]);
    expect(second.status === "ok" && second.note).toMatch(/No policy/);
  });

  it("shows a role with its holders, their combined risks and the policies that name it", async () => {
    const found = await entity(h, "role", GLOBAL_ADMIN);
    expect(found).toMatchObject({
      name: "Global Administrator",
      tags: ["Privileged role", "3 assignments"],
    });
    const holders = okRows(found, "holders");
    expect(holders.map((item) => item.subject)).toEqual([
      { kind: "person", id: ana, name: "Ana Ames", principal: "ana@example.invalid" },
      // A holder the audit could not sort is a group only because Groups and Teams lists it, and is
      // named from there, not by its id.
      { kind: "group", id: PLANT, name: "Plant team", principal: null },
      { kind: "app", id: APP, name: "Plant robot", principal: null },
    ]);
    expect(holders[0]?.cells.map((cell) => cell.key)).toEqual([
      "holderType",
      "held",
      "through",
      "scope",
      "mfa",
    ]);
    const risks = okRows(found, "risks");
    expect(risks).toHaveLength(1);
    expect(risks[0]?.subject).toMatchObject({ kind: "person", id: ana, name: "Ana Ames" });
    expect(risks[0]?.risks).toEqual([
      expect.objectContaining({ id: "admin-unprotected", severity: expect.any(String) }),
    ]);
    expect(okRows(found, "policies")[0]).toMatchObject({
      subject: { kind: "policy", id: POLICY },
      tags: ["Includes it"],
    });
    const lab = await entity(h, "role", LAB_ROLE);
    expect(lab.name).toBe("Lab Reports Reader");
    expect(okRows(lab, "holders")[0]).toMatchObject({
      lists: [{ label: "Through group", items: [{ kind: "group", id: STAFF }] }],
    });
    expect(okRows(lab, "policies")[0]?.tags).toEqual(["Leaves it out"]);
    // Ben holds the role and trips no rule, so the run has nothing to say about him.
    const quiet = sectionOf(lab, "risks");
    expect(quiet.status === "ok" && quiet.rows).toEqual([]);
    expect(quiet.status === "ok" && quiet.note).toMatch(/trips a rule/);
  });

  it("leaves out the combined risks section for a role nobody holds", async () => {
    const found = await entity(h, "role", guid(21));
    expect(found.sections.map((section) => [section.id, section.status])).toEqual([
      ["holders", "not_in_run"],
      ["policies", "ok"],
    ]);
  });

  it("shows a policy with the people, groups and roles it names, by name where the audits know it", async () => {
    const found = await entity(h, "policy", POLICY);
    expect(found).toMatchObject({ name: "Require MFA for admins", tags: ["On"] });
    const [policy] = okRows(found, "policy");
    const lists = new Map(policy?.lists.map((list) => [list.label, list.items]));
    expect(lists.get("Includes these people")).toEqual([
      { kind: "person", id: ana, name: "Ana Ames", principal: "ana@example.invalid" },
      { kind: "person", id: ben, name: "Ben Birch", principal: "ben@example.invalid" },
      // Nobody lists this account, so it is unnamed and the page says so.
      { kind: "person", id: cai, name: null, principal: null },
    ]);
    expect(lists.get("Includes these groups")).toEqual([
      { kind: "group", id: PLANT, name: "Plant team", principal: null },
      { kind: "group", id: guid(99), name: null, principal: null },
    ]);
    expect(lists.get("Includes these roles")).toEqual([
      { kind: "role", id: GLOBAL_ADMIN, name: "Global Administrator", principal: null },
      { kind: "role", id: guid(98), name: null, principal: null },
    ]);
    expect(lists.get("Leaves out these groups")?.[0]).toMatchObject({ id: STAFF, name: "Staff" });
    expect(lists.get("Leaves out these roles")?.[0]).toMatchObject({
      id: LAB_ROLE,
      name: "Lab Reports Reader",
    });
    expect(lists.has("Leaves out these people")).toBe(false);
    expect(policy?.cells.map((cell) => cell.key)).toEqual([
      "state",
      "appliesTo",
      "excludes",
      "apps",
      "grant",
    ]);
  });

  it("bounds a long list and counts what it left out", async () => {
    const many = Array.from({ length: 45 }, (_value, index) => guid(1_000 + index));
    await h.storeRun("identity.conditional-access", [
      row(POLICY, null, { policy: "Wide", state: "On", includeUserIds: many }),
    ]);
    const found = await entity(h, "policy", POLICY);
    const [policy] = okRows(found, "policy");
    const people = policy?.lists.find((list) => list.label === "Includes these people");
    expect(people?.items).toHaveLength(20);
    expect(people?.more).toBe(25);
    await storeEverything(h);
  });

  it("shows a device with its owner as a person, and none for a device with no primary user", async () => {
    const found = await entity(h, "device", DEVICE);
    expect(found).toMatchObject({
      name: "Plant laptop",
      tags: ["Windows", "Not compliant", "Corporate"],
    });
    const [device] = okRows(found, "device");
    expect(device?.lists).toEqual([
      {
        label: "Owner",
        items: [{ kind: "person", id: ana, name: "Ana Ames", principal: "ana@example.invalid" }],
        more: 0,
      },
    ]);
    expect(device?.cells.map((cell) => cell.key)).toContain("serial");
    const spare = await entity(h, "device", guid(41));
    expect(okRows(spare, "device")[0]?.lists).toEqual([]);
  });

  it("shows an app found by its service principal id or by its application id", async () => {
    for (const id of [APP, APP_ID]) {
      const found = await entity(h, "app", id);
      expect(found).toMatchObject({
        name: "Plant robot",
        tags: ["Unverified", "Another organisation", "High-risk permissions"],
      });
      expect(okRows(found, "app")[0]?.key).toBe(APP);
    }
    expect(sectionOf(await entity(h, "app", guid(60)), "app").status).toBe("not_in_run");
  });

  it("keeps a failing source from taking the rest down, and never says why", async () => {
    const original = h.database.tenant.bind(h.database);
    h.database.tenant = ((
      tenant: string,
      work: (sql: SqlClient) => Promise<unknown>,
      options?: never,
    ) =>
      original(
        tenant,
        (sql) =>
          work({
            query: ((text: string, params?: readonly unknown[]) =>
              Array.isArray(params) && params.includes("identity.conditional-access")
                ? Promise.reject(new Error("Synthetic storage failure."))
                : sql.query(text, params)) as SqlClient["query"],
          } as SqlClient),
        options,
      )) as typeof original;
    const response = await h.get(url("group", PLANT));
    const search = await h.get("/v1/entities/search?q=plant");
    h.database.tenant = original;
    const found = response.json() as Entity;
    expect(found.sections.map((section) => section.status)).toEqual(["ok", "failed", "ok"]);
    expect(response.body).not.toContain("Synthetic storage failure");
    expect(search.json()).toMatchObject({ partial: true });
  });
});

describe("entity search", () => {
  let h: Harness;
  beforeAll(async () => {
    h = await harness();
    await storeEverything(h);
  }, 60_000);
  afterAll(async () => {
    await h?.app.close();
  });
  const search = async (q: string): Promise<EntitySearch> => {
    const response = await h.get(`/v1/entities/search?q=${encodeURIComponent(q)}`);
    expect(response.statusCode, q).toBe(200);
    return response.json();
  };

  it("finds each kind by name, whatever the case, with a word to tell them apart", async () => {
    const found = await search("PLANT");
    expect(found.results.map((item) => [item.kind, item.name, item.detail])).toEqual([
      ["group", "Plant team", "Team"],
      ["device", "Plant laptop", "Windows"],
      ["app", "Plant dashboard", "Plant Works"],
      ["app", "Plant robot", "Unverified"],
    ]);
    expect(found.more).toBe(false);
    expect((await search("global")).results).toEqual([
      { kind: "role", id: GLOBAL_ADMIN, name: "Global Administrator", detail: null },
    ]);
    expect((await search("mfa")).results.map((item) => item.kind)).toEqual(["policy"]);
  });

  it("lists a role once however many people hold it", async () => {
    const found = await search("administrator");
    expect(found.results.filter((item) => item.kind === "role")).toHaveLength(1);
  });

  it("treats % and _ as the characters they are", async () => {
    expect((await search("100%")).results.map((item) => item.name)).toEqual(["Lab 100%"]);
    expect((await search("%%")).results).toEqual([]);
    expect((await search("s_o")).results.map((item) => item.name)).toEqual(["Admins_one"]);
    // An underscore is not "any character", so this does not find "Admins".
    expect((await search("ad_ins")).results).toEqual([]);
    expect((await search("a\\b")).results).toEqual([]);
  });

  it("finds a name with % or _ in it past the rows the database reads first", async () => {
    // The database reads at most 100 matching rows before the answer is settled; if _ were a
    // wildcard, every one of the plain rows ahead of the odd one would match and fill that limit.
    await h.storeRun("groups.teams", [
      ...Array.from({ length: 120 }, (_value, index) =>
        row(guid(3_000 + index), null, { group: `Plain ${index}0`, kind: "Security group" }),
      ),
      row(guid(3_500), null, { group: "Plain_one 50%", kind: "Security group" }),
    ]);
    for (const text of ["in_", "0%", "_one 50%"])
      expect(
        (await search(text)).results.map((item) => item.name),
        text,
      ).toEqual(["Plain_one 50%"]);
    await storeEverything(h);
  });

  it("finds every role however many holders another has, and says when there is more", async () => {
    const holder = (n: number, role: string, name: string) =>
      row(`held|${n}|${role}`, null, {
        holder: { name: `Holder ${n}`, principal: null },
        holderType: "User",
        role: name,
        roleId: role,
        holderId: guid(5_000 + n),
      });
    await h.storeRun("directory.privileged-roles", [
      ...Array.from({ length: 150 }, (_value, index) => holder(index, guid(60), "Crowd admin")),
      holder(900, guid(61), "Crowd reader"),
      holder(901, guid(62), "Crowd writer"),
    ]);
    const found = await search("crowd");
    expect(found.results.map((item) => item.name)).toEqual([
      "Crowd admin",
      "Crowd reader",
      "Crowd writer",
    ]);
    expect(found.more).toBe(false);
    // More matching rows than the database reads: the answer says so.
    await h.storeRun(
      "groups.teams",
      Array.from({ length: 130 }, (_value, index) =>
        row(guid(7_000 + index), null, { group: `Crowd ${index}`, kind: "Security group" }),
      ),
    );
    expect((await search("crowd")).more).toBe(true);
    await storeEverything(h);
  });

  it("puts names that start with the words first, and gives each kind a turn", async () => {
    await h.storeRun(
      "groups.teams",
      Array.from({ length: 30 }, (_value, index) =>
        row(guid(2_000 + index), null, {
          group: index === 29 ? "Alpha first" : `Team alpha ${String(index).padStart(2, "0")}`,
          kind: "Security group",
        }),
      ),
    );
    await h.storeRun("devices.managed", [
      row(guid(41), null, { device: "alpha tablet", os: "iOS" }),
    ]);
    const found = await search("alpha");
    expect(found.results).toHaveLength(20);
    expect(found.more).toBe(true);
    expect(found.results[0]).toMatchObject({ kind: "group", name: "Alpha first" });
    expect(found.results[1]).toMatchObject({ kind: "device", name: "alpha tablet" });
  });

  it("leaves out an id that is not a GUID, which the panel could not open", async () => {
    await h.storeRun("groups.teams", [
      row("not-a-guid", null, { group: "Bad id group", kind: "Security group" }),
    ]);
    expect((await search("bad id")).results).toEqual([]);
  });
});

describe("entity panels over runs of other shapes", () => {
  let h: Harness;
  beforeAll(async () => {
    h = await harness();
  }, 60_000);
  afterAll(async () => {
    await h?.app.close();
  });

  const heldBy = (n: number, role: string, over: AuditRow["cells"] = {}, personId?: string) =>
    row(
      `held|${n}|${role}`,
      personId ?? guid(8_000 + n),
      {
        holder: { name: `Holder ${n}`, principal: null },
        holderType: "User",
        role: "Lab Reports Reader",
        roleId: role,
        holderId: personId ?? guid(8_000 + n),
        ...over,
      },
      [],
    );

  it("says a run from before the ids cannot answer, instead of saying nothing reaches the thing", async () => {
    await h.storeRun("groups.teams", [row(PLANT, null, { group: "Plant team", ownerIds: [ana] })], {
      version: 1,
    });
    await h.storeRun("directory.privileged-roles", [heldBy(1, LAB_ROLE)], { version: 1 });
    await h.storeRun("identity.conditional-access", [row(POLICY, null, { policy: "Old" })], {
      version: 1,
    });
    await h.storeRun("identity.combined-risks", [row("r", ana, {}, ["critical"])]);
    const group = await entity(h, "group", PLANT);
    expect(group.sections.map((section) => [section.id, section.status])).toEqual([
      ["group", "ok"],
      ["policies", "outdated"],
      ["roles", "outdated"],
    ]);
    const own = sectionOf(group, "group");
    expect(own.status === "ok" && own.note).toMatch(/Run Groups and Teams again/);
    expect(JSON.stringify(group)).not.toMatch(/No policy includes|No role in the latest/);
    const role = await entity(h, "role", LAB_ROLE);
    expect(role.sections.map((section) => [section.id, section.status])).toEqual([
      ["holders", "outdated"],
      ["policies", "outdated"],
    ]);
    expect(role.tags).toEqual([]);
    expect(JSON.stringify(role)).not.toMatch(/trips a rule|No policy includes/);
    const policy = await entity(h, "policy", POLICY);
    const row0 = sectionOf(policy, "policy");
    expect(row0.status === "ok" && row0.note).toMatch(/Run Conditional Access again/);
    // A role cannot be searched for in such a run, and the answer says the list may be short.
    expect((await h.get("/v1/entities/search?q=lab")).json()).toMatchObject({ partial: true });
    // The same runs made again by the current audit answer normally.
    await h.storeRun("directory.privileged-roles", [heldBy(1, LAB_ROLE)]);
    await h.storeRun("identity.conditional-access", [row(POLICY, null, { policy: "New" })]);
    expect((await entity(h, "role", LAB_ROLE)).sections.map((section) => section.status)).toEqual([
      "ok",
      "ok",
      "ok",
    ]);
  });

  it("links a holder as a group or an app only when an audit lists it, and names it from there", async () => {
    const known = guid(8_101);
    const robot = guid(8_102);
    const stranger = guid(8_103);
    await h.storeRun("groups.teams", [row(known, null, { group: "IT operations admins" })]);
    await h.storeRun("apps.consents", [row(robot, null, { app: "Sync robot" })]);
    const other = (id: string, name: string | null, type: string) =>
      row(`held|${id}`, null, {
        holder: { name, principal: null },
        holderType: type,
        role: "Lab Reports Reader",
        roleId: LAB_ROLE,
        holderId: id,
      });
    await h.storeRun("directory.privileged-roles", [
      other(known, known, "Group or other"),
      other(robot, null, "Service principal"),
      other(stranger, stranger, "Group or other"),
    ]);
    const holders = okRows(await entity(h, "role", LAB_ROLE), "holders").map((r) => r.subject);
    expect(holders).toEqual([
      { kind: "group", id: known, name: "IT operations admins", principal: null },
      { kind: "app", id: robot, name: "Sync robot", principal: null },
      // Nothing lists it, so it is not claimed to be a group: text, no link.
      { kind: "other", id: stranger, name: null, principal: null },
    ]);
  });

  it("checks every holder it reads, not only the 25 it shows, and says when it could not check them all", async () => {
    const role = guid(70);
    const people = Array.from({ length: 30 }, (_value, index) => guid(9_000 + index));
    await h.storeRun(
      "directory.privileged-roles",
      people.map((id, index) => heldBy(index, role, {}, id)),
    );
    // Only the 28th holder trips a rule: past the 25 that are shown.
    await h.storeRun("identity.combined-risks", [
      row("r", people[27] ?? "", { person: { name: "Holder 27", principal: null } }, [
        "admin-unprotected",
      ]),
    ]);
    const found = await entity(h, "role", role);
    const holders = sectionOf(found, "holders");
    expect(holders.status === "ok" && [holders.rows.length, holders.more]).toEqual([25, true]);
    const risks = okRows(found, "risks");
    expect(risks).toHaveLength(1);
    expect(risks[0]?.subject).toMatchObject({ id: people[27] });
    // With nobody tripping a rule among all 30, the negative is said over all 30.
    await h.storeRun("identity.combined-risks", [row("x", guid(1), {}, [])]);
    const quiet = sectionOf(await entity(h, "role", role), "risks");
    expect(quiet.status === "ok" && quiet.note).toMatch(/trips a rule/);
    // A thousand and one holders cannot all be checked, so no negative is stated.
    const many = Array.from({ length: 1_001 }, (_value, index) =>
      heldBy(index, role, {}, guid(20_000 + index)),
    );
    await h.storeRun("directory.privileged-roles", many);
    const cut = await entity(h, "role", role);
    const note = sectionOf(cut, "risks");
    expect(note.status === "ok" && note.note).toBe(
      "Checked the first 1,000 holders. The role has more, which were not checked.",
    );
    expect(JSON.stringify(cut)).not.toMatch(/trips a rule in this run/);
    expect(cut.tags).toContain("1000 or more assignments");
  });
});

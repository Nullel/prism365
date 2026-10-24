import {
  type Audit,
  type AuditResult,
  auditDefinition,
  passwordsAudit,
  privilegedRolesAudit,
} from "@prism365/audits";
import { GraphHttpClient } from "@prism365/graph";
import { scopeGraph } from "@prism365/modules";
import { describe, expect, it } from "vitest";
import {
  assertSynthetic,
  buildSyntheticTenant,
  createSyntheticGraph,
  DOMAIN,
  guid,
  SYNTHETIC_TENANT_ID,
} from "../tenant.ts";

// The synthetic tenant's directory roles, domains and passwords, read by the real audits through
// the real Graph client: the recorder does exactly this, so what these tests see is what the
// portfolio demo shows.
const now = new Date("2026-10-22T12:00:00Z");
const tenant = buildSyntheticTenant(now);
const graph = createSyntheticGraph({ now });
const bodies: string[] = [];
const requests: URL[] = [];
const recording: typeof fetch = async (input, init) => {
  requests.push(new URL(String(input)));
  const response = await graph(input, init);
  bodies.push(await response.clone().text());
  return response;
};

async function run(audit: Audit, granted: string[]) {
  const client = new GraphHttpClient({
    getAccessToken: async () => "synthetic",
    fetchImplementation: recording,
    maxResponseBytes: 8_388_608,
    maxRetries: 0,
  });
  return audit.run({
    graph: scopeGraph(client, [...audit.packs, ...audit.optionalPacks]),
    now,
    granted: new Set(granted),
    signal: new AbortController().signal,
    scope: null,
    tenantId: SYNTHETIC_TENANT_ID,
  });
}

const rolesGranted = [
  "RoleManagement.Read.Directory",
  "User.Read.All",
  "Application.Read.All",
  "AuditLog.Read.All",
];
const passwordsGranted = ["User.Read.All", "Domain.Read.All"];

const tile = (result: AuditResult, label: string) =>
  result.summary.find((item) => item.label === label);
const named = (name: string) => {
  const found = tenant.users.find((user) => user.displayName === name);
  if (!found) throw new Error(`No synthetic user ${name}.`);
  return found;
};

describe("the synthetic privileged roles", () => {
  it("hold the findings the audit looks for", async () => {
    const result = await run(privilegedRolesAudit, rolesGranted);
    expect(result.complete).toBe(true);
    // The demo administrator and the general manager hold it for good; the IT support technician
    // has it switched on. The manager's is the permanent assignment too many.
    expect(tile(result, "Global Administrators")).toMatchObject({ value: 3, tone: "good" });
    expect(tile(result, "Permanent privileged assignments")?.value).toBe(5);
    expect(tile(result, "Disabled accounts with a privileged role")?.value).toBe(1);
    expect(tile(result, "Guests with a privileged role")?.value).toBe(1);
    expect(tile(result, "Service principals with a privileged role")?.value).toBe(1);
    expect(tile(result, "Privileged role holders")?.value).toBe(8);
  });

  it("show every way a role is held, and every assignment once", async () => {
    const result = await run(privilegedRolesAudit, rolesGranted);
    const held = result.rows.map((row) => row.cells.held);
    for (const how of ["Permanent", "Time-bound", "Activated", "Eligible"])
      expect(held, how).toContain(how);
    // Every active assignment has its schedule record in this tenant.
    expect(held).not.toContain("Active");
    expect(new Set(result.rows.map((row) => row.key)).size).toBe(result.rows.length);
    const technician = result.rows.filter(
      (row) =>
        row.personId === named("Tobias Vance").id && row.cells.role === "Global Administrator",
    );
    expect(technician.map((row) => row.cells.held).sort()).toEqual(["Activated", "Eligible"]);
  });

  it("name an app and leave a group as Group or other", async () => {
    const result = await run(privilegedRolesAudit, rolesGranted);
    const app = result.rows.find((row) => row.cells.holderType === "Service principal");
    expect(app?.cells).toMatchObject({
      holder: { name: "Exchange Hybrid Agent" },
      role: "Exchange Administrator",
      held: "Permanent",
    });
    const group = result.rows.find((row) => row.cells.holderType === "Group or other");
    expect(group?.cells.holder).toMatchObject({ name: guid("group:it-operations-admins") });
    expect(group?.cells.privileged).toBe(false);
  });

  it("fill MFA registration for accounts only", async () => {
    const result = await run(privilegedRolesAudit, rolesGranted);
    const admin = result.rows.find((row) => row.personId === named("Demo administrator").id);
    expect(admin?.cells.mfa).toBe(true);
    for (const row of result.rows.filter((item) => item.personId === null))
      expect(row.cells.mfa).toBeNull();
  });

  it("are the same every time", async () => {
    const first = await run(privilegedRolesAudit, rolesGranted);
    const second = await run(privilegedRolesAudit, rolesGranted);
    expect(second.rows).toEqual(first.rows);
    expect(second.summary).toEqual(first.summary);
  });

  it("are read without the optional permissions too", async () => {
    const result = await run(privilegedRolesAudit, ["RoleManagement.Read.Directory"]);
    expect(result.complete).toBe(true);
    expect(tile(result, "Global Administrators")?.value).toBe(3);
  });
});

describe("the synthetic passwords", () => {
  it("hold the findings the audit looks for", async () => {
    const result = await run(passwordsAudit, passwordsGranted);
    const members = tenant.users.filter((user) => user.userType === "Member");
    expect(result.complete).toBe(true);
    expect(tile(result, "Member accounts")?.value).toBe(members.length);
    expect(tile(result, "Strong password not required")?.value).toBe(1);
    expect(tile(result, "Never expires")?.value).toBe(1);
    expect(tile(result, "Expired")?.value).toBe(1);
    expect(tile(result, "Expires soon")?.value).toBe(1);
    expect(tile(result, "Enabled, never changed")?.value).toBe(1);
    expect(tile(result, "Older than 1 year")?.value).toBeGreaterThanOrEqual(5);
  });

  it("tell the expired account from the one that expires this week", async () => {
    const result = await run(passwordsAudit, passwordsGranted);
    const of = (name: string) => result.rows.find((row) => row.personId === named(name).id);
    expect(of("Anselm Gray")?.cells).toMatchObject({ policy: "180 days", status: "Expired" });
    const soon = of("Desmond Ferrante");
    expect(soon?.cells.status).toBe("Expires soon");
    const days = (Date.parse(String(soon?.cells.expires)) - now.getTime()) / 86_400_000;
    expect(days).toBeGreaterThan(0);
    expect(days).toBeLessThan(7);
  });

  it("show a synced account as managed on-premises and the research domain as never expiring", async () => {
    const result = await run(passwordsAudit, passwordsGranted);
    const of = (name: string) => result.rows.find((row) => row.personId === named(name).id);
    expect(of("Magnus Sterling")?.cells).toMatchObject({
      source: "On-premises",
      policy: "Managed on-premises",
      expires: null,
      status: "Older than 1 year",
    });
    expect(of("Elias Wragg")?.cells).toMatchObject({
      policy: "Never",
      expires: null,
      status: "Older than 1 year",
    });
    expect(of("Mirela Cobb")?.cells).toMatchObject({ neverExpires: true, expires: null });
    expect(of("Lab Training Kiosk")?.cells).toMatchObject({
      weak: true,
      status: "Strong password not required",
    });
  });

  it("are listed for members only", async () => {
    const result = await run(passwordsAudit, passwordsGranted);
    const guests = new Set(
      tenant.users.filter((user) => user.userType === "Guest").map((user) => user.id),
    );
    expect(guests.size).toBeGreaterThan(0);
    for (const row of result.rows) expect(guests.has(row.personId ?? ""), row.key).toBe(false);
    expect(result.notes.join("\n")).toContain(`${guests.size} guest accounts are not listed`);
  });

  it("leave everyone else with a recent password", async () => {
    const result = await run(passwordsAudit, passwordsGranted);
    const ordinary = result.rows.filter((row) => row.cells.status === "OK");
    expect(ordinary.length).toBeGreaterThan(25);
    for (const row of ordinary) expect(Number(row.cells.days)).toBeLessThan(90);
  });
});

describe("what the lookups return", () => {
  it("is only the objects asked for, when a request filters by ID", async () => {
    const admin = named("Demo administrator");
    const filter = encodeURIComponent(`id in ('${admin.id}','${guid("sp:marrow")}')`);
    const fetchJson = async (path: string) =>
      (await graph(`https://graph.microsoft.com/v1.0/${path}$filter=${filter}`)).json() as Promise<{
        value: { id: string }[];
      }>;
    expect((await fetchJson("users?")).value.map((item) => item.id)).toEqual([admin.id]);
    expect((await fetchJson("servicePrincipals?")).value.map((item) => item.id)).toEqual([
      guid("sp:marrow"),
    ]);
  });

  it("is the domains' password policies, with the company domain on its own", async () => {
    const result = await (
      await graph(
        "https://graph.microsoft.com/v1.0/domains?$select=id,passwordValidityPeriodInDays,passwordNotificationWindowInDays",
      )
    ).json();
    expect(result.value).toEqual([
      { id: DOMAIN, passwordValidityPeriodInDays: 180, passwordNotificationWindowInDays: 14 },
      {
        id: "larkspur-research.example",
        passwordValidityPeriodInDays: 2_147_483_647,
        passwordNotificationWindowInDays: null,
      },
      {
        id: "larkspur.onmicrosoft.com",
        passwordValidityPeriodInDays: null,
        passwordNotificationWindowInDays: null,
      },
    ]);
  });
});

describe("everything the two audits read", () => {
  it("belongs to the invented tenant, and only ever asks", async () => {
    bodies.length = 0;
    requests.length = 0;
    await run(privilegedRolesAudit, rolesGranted);
    await run(passwordsAudit, passwordsGranted);
    expect(bodies.length).toBeGreaterThan(6);
    // The recorder refuses a recording with text from outside the tenant; so does this test.
    assertSynthetic(bodies.join("\n"), tenant.addresses);
    expect(requests.some((url) => url.pathname.endsWith("/domains"))).toBe(true);
  });

  it("are published with the permissions their packs carry", () => {
    expect(auditDefinition(privilegedRolesAudit).permissions).toEqual([
      "RoleManagement.Read.Directory",
    ]);
    expect(auditDefinition(passwordsAudit).optionalPermissions).toEqual(["Domain.Read.All"]);
  });
});

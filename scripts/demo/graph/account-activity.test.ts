import { type Audit, accountActivityAudit, auditDefinition } from "@prism365/audits";
import { queryAuditRows } from "@prism365/contracts";
import { GraphHttpClient } from "@prism365/graph";
import { scopeGraph } from "@prism365/modules";
import { describe, expect, it } from "vitest";
import {
  assertSynthetic,
  buildSyntheticTenant,
  createSyntheticGraph,
  SYNTHETIC_TENANT_ID,
} from "../tenant.ts";

// The synthetic tenant's seven usage reports and group memberships, read by the real Account
// activity audit through the real Graph client: the recorder does exactly this, so what these tests
// see is what the portfolio demo shows.
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

const granted = [
  "User.Read.All",
  "AuditLog.Read.All",
  "Reports.Read.All",
  "ReportSettings.Read.All",
  "LicenseAssignment.Read.All",
];

async function run(audit: Audit = accountActivityAudit, scope: ReadonlySet<string> | null = null) {
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
    scope,
    tenantId: SYNTHETIC_TENANT_ID,
  });
}

const nameOf = (row: { cells: Record<string, unknown> }) =>
  (row.cells.person as { name: string }).name;
const tile = (result: { summary: { label: string; value: number }[] }, label: string) =>
  result.summary.find((item) => item.label === label)?.value;

// Each verdict, with the accounts the synthetic tenant plants for it.
const planted: Record<string, string[]> = {
  "Signs in, uses nothing": ["Bram Okafor", "Idris Calloway"],
  "Quiet, others use their files": ["Delphine Okoro"],
  "Quiet everywhere": [
    "Anselm Gray",
    "Harriet Moss",
    "Isadora Quill",
    "Jory Fennimore",
    "Leopold Vass",
    "Magnus Sterling",
    "Mireille Tadesse",
    "Rufus Pemberton",
    "Tamsin Rooke",
  ],
  "Recently quiet": ["Calla Dunmore", "Elias Wragg", "Ravenna Doyle"],
  "Never signed in": ["Callum Rousseau", "Lab Training Kiosk", "Quentin Hale"],
  "Not enough data": ["Brigid Alder", "Sabine Okonkwo-Reyes"],
};

describe("the synthetic account activity", () => {
  it("is read whole, from all seven reports, each once", async () => {
    requests.length = 0;
    const result = await run();
    expect(result.complete).toBe(true);
    expect(result.rows).toHaveLength(tenant.users.length);
    const reports = requests.filter((url) => url.pathname.startsWith("/v1.0/reports/"));
    expect(reports.map((url) => url.pathname.split("/").at(-1)).sort()).toEqual([
      "getEmailActivityUserDetail(period='D90')",
      "getMailboxUsageDetail(period='D90')",
      "getOffice365ActiveUserDetail(period='D90')",
      "getOneDriveActivityUserDetail(period='D90')",
      "getOneDriveUsageAccountDetail(period='D90')",
      "getSharePointActivityUserDetail(period='D90')",
      "getTeamsUserActivityUserDetail(period='D90')",
    ]);
    expect(requests.filter((url) => url.hostname === "reports.office.com")).toHaveLength(7);
    expect(requests.filter((url) => url.pathname.endsWith("/memberOf"))).toHaveLength(
      tenant.users.length,
    );
  });

  it("plants an account for every verdict, and for each one the accounts that earn it", async () => {
    const result = await run();
    const verdicts = (verdict: string) =>
      result.rows
        .filter((row) => row.cells.verdict === verdict)
        .map(nameOf)
        .sort();
    for (const [verdict, names] of Object.entries(planted))
      expect(verdicts(verdict), verdict).toEqual(names);
    const active = verdicts("Active");
    expect(active.length).toBeGreaterThan(25);
    // Everyone who signed in this month and used something is active, including a shared
    // mailbox that never signs in and a guest who uses SharePoint.
    expect(active).toEqual(
      expect.arrayContaining(["Alder Whitcombe", "Lab Results (shared)", "Demo administrator"]),
    );
    const listed = new Set([...Object.values(planted).flat(), ...active]);
    expect(listed.size).toBe(tenant.users.length);
  });

  it("gives each tile something to show", async () => {
    const result = await run();
    expect(tile(result, "Accounts")).toBe(tenant.users.length);
    expect(tile(result, "Quiet everywhere, with directory roles")).toBe(1);
    expect(tile(result, "Quiet everywhere")).toBe(9);
    expect(tile(result, "Quiet, others use their files")).toBe(1);
    expect(tile(result, "Signs in, uses nothing")).toBe(2);
  });

  it("holds a quiet account whose directory role is the reason to look", async () => {
    const result = await run();
    const definition = auditDefinition(accountActivityAudit);
    const quietWithRole = queryAuditRows(definition, result.rows, { view: "quiet-roles" });
    expect(quietWithRole.map(nameOf)).toEqual(["Anselm Gray"]);
    expect(quietWithRole[0]?.cells).toMatchObject({
      roles: 1,
      roleNames: ["Exchange Administrator"],
      status: "Enabled",
    });
    // The administrators who are active hold roles too, and are not flagged.
    expect(queryAuditRows(definition, result.rows, { view: "roles" }).map(nameOf).sort()).toEqual([
      "Anselm Gray",
      "Demo administrator",
      "Tobias Vance",
    ]);
  });

  it("shows the quiet account whose OneDrive is still opened", async () => {
    const result = await run();
    const row = result.rows.find((item) => nameOf(item) === "Delphine Okoro");
    expect(row?.cells).toMatchObject({
      verdict: "Quiet, others use their files",
      oneDriveLast: null,
      oneDriveActiveFiles: 14,
      oneDriveFileCount: 212,
    });
    expect(row?.cells.oneDriveAccountLast).toBe("2026-10-16T00:00:00.000Z");
  });

  it("shows a large mailbox on a quiet account, and licenses assigned through a group", async () => {
    const result = await run();
    const definition = auditDefinition(accountActivityAudit);
    expect(queryAuditRows(definition, result.rows, { view: "large-mailbox" }).map(nameOf)).toEqual([
      "Rufus Pemberton",
    ]);
    const rooke = result.rows.find((item) => nameOf(item) === "Tamsin Rooke");
    expect(rooke?.cells).toMatchObject({
      licenseAssignment: "Both",
      licenseGroups: ["Licensing, Microsoft 365 Business Premium"],
    });
  });

  it("counts group memberships, and has a Laboratory team and a staff group", async () => {
    const result = await run();
    const row = (name: string) => result.rows.find((item) => nameOf(item) === name);
    expect(row("Dagny Holloway")?.cells).toMatchObject({
      groups: 2,
      groupNames: ["Larkspur staff", "Laboratory team"],
      roles: 0,
    });
    expect(row("Demo administrator")?.cells).toMatchObject({
      roleNames: ["Global Administrator"],
    });
    // A departed employee's report row matches no account and is counted, not shown.
    expect(result.notes.join("\n")).toContain("1 name in the reports matches no account");
    expect(JSON.stringify(result.rows)).not.toContain("Hollis");
  });

  it("identifies every account by its directory id for a handoff", async () => {
    const result = await run();
    const ids = new Set(tenant.users.map((user) => user.id));
    for (const row of result.rows) {
      expect(row.personId).toBe(row.key);
      expect(ids.has(row.key)).toBe(true);
    }
  });

  it("serves memberships by object ID and by principal name, and nothing for an unknown account", async () => {
    const admin = tenant.users.find((user) => user.displayName === "Demo administrator");
    const byId = await graph(
      `https://graph.microsoft.com/v1.0/users/${admin?.id}/memberOf?$select=id,displayName`,
    );
    const byName = await graph(
      `https://graph.microsoft.com/v1.0/users('${encodeURIComponent(admin?.userPrincipalName ?? "")}')/memberOf`,
    );
    const body = (await byId.json()) as { value: { "@odata.type": string; displayName: string }[] };
    expect(byId.status).toBe(200);
    expect(await byName.json()).toEqual(body);
    expect(body.value.map((item) => [item["@odata.type"], item.displayName])).toEqual([
      ["#microsoft.graph.group", "Larkspur staff"],
      ["#microsoft.graph.directoryRole", "Global Administrator"],
    ]);
    const unknown = await graph(
      "https://graph.microsoft.com/v1.0/users/00000000-0000-4000-8000-000000000000/memberOf",
    );
    expect(unknown.status).toBe(404);
  });

  it("names every account in the Teams report by its object ID, which is how the audit finds it", async () => {
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic",
      fetchImplementation: graph,
      maxRetries: 0,
    });
    const table = await client.getReport("reports/getTeamsUserActivityUserDetail(period='D90')", {
      feature: "reports.usage.read",
      permission: "Reports.Read.All",
    });
    const principalAt = table.index.get("User Principal Name") ?? -1;
    const idAt = table.index.get("User Id") ?? -1;
    const byPrincipal = new Map(tenant.users.map((user) => [user.userPrincipalName, user.id]));
    expect(table.rows.length).toBeGreaterThan(20);
    for (const row of table.rows) expect(row[idAt]).toBe(byPrincipal.get(row[principalAt] ?? ""));
  });

  it("applies a scope, and still counts quiet accounts with roles outside it", async () => {
    const result = await run(accountActivityAudit, new Set(["larkspur-research.example"]));
    expect(result.rows.length).toBeGreaterThan(0);
    expect(result.rows.length).toBeLessThan(tenant.users.length);
    expect(result.summary.map((item) => item.label)).toContain(
      "Quiet everywhere, with directory roles, outside the scope",
    );
  });

  it("is the same every time", async () => {
    const first = await run();
    const second = await run();
    expect(second.rows).toEqual(first.rows);
    expect(second.summary).toEqual(first.summary);
  });

  it("is made of nothing from outside the invented tenant", async () => {
    bodies.length = 0;
    await run();
    expect(bodies.length).toBeGreaterThan(60);
    // The recorder refuses a recording with text from outside the tenant; so does this test. The
    // report downloads are in it too.
    assertSynthetic(bodies.join("\n"), tenant.addresses);
  });
});

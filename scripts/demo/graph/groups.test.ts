import {
  type Audit,
  type AuditResult,
  auditDefinition,
  GROUP_KINDS,
  groupsTeamsAudit,
  MAILBOX_STATUSES,
  mailboxStorageAudit,
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
  RESEARCH_DOMAIN,
  SYNTHETIC_TENANT_ID,
} from "../tenant.ts";

// The synthetic tenant's groups, teams and mailbox sizes, read by the real audits through the real
// Graph client: the recorder does exactly this, so what these tests see is what the portfolio demo
// shows.
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

async function run(audit: Audit, granted: string[], scope: string[] | null = null) {
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
    scope: scope ? new Set(scope) : null,
    tenantId: SYNTHETIC_TENANT_ID,
  });
}

const groupsGranted = ["Group.Read.All", "Reports.Read.All", "ReportSettings.Read.All"];
const mailboxGranted = ["Reports.Read.All", "ReportSettings.Read.All"];

const tile = (result: AuditResult, label: string) =>
  result.summary.find((item) => item.label === label)?.value;
const rowOf = (result: AuditResult, name: string, column = "group") => {
  const found = result.rows.find((row) => row.cells[column] === name);
  if (!found) throw new Error(`No row ${name}.`);
  return found;
};

describe("the synthetic groups and teams", () => {
  it("hold the findings the audit looks for", async () => {
    const result = await run(groupsTeamsAudit, groupsGranted);
    expect(result.complete).toBe(true);
    expect(result.rows).toHaveLength(14);
    expect(tile(result, "Groups and teams with no owners")).toBe(2);
    expect(tile(result, "Groups with only disabled owners")).toBe(1);
    expect(tile(result, "Teams with one owner")).toBe(2);
    expect(tile(result, "Teams with guests")).toBe(3);
    expect(tile(result, "Inactive teams")).toBe(1);
    expect(tile(result, "Public teams")).toBe(2);
    expect(tile(result, "Groups expiring in 30 days")).toBe(1);
    expect(tile(result, "Role-assignable groups")).toBe(1);
  });

  it("have every kind, and every status except Owners not read, which only a very large tenant earns", async () => {
    const result = await run(groupsTeamsAudit, groupsGranted);
    expect(new Set(result.rows.map((row) => row.cells.kind))).toEqual(new Set(GROUP_KINDS));
    expect([...new Set(result.rows.map((row) => row.cells.status))].sort()).toEqual([
      "Expires within 30 days",
      "Inactive 90 days",
      "No owners",
      "No usage data",
      "OK",
      "One owner",
      "Only disabled owners",
    ]);
  });

  it("plant the stories the contract names", async () => {
    const result = await run(groupsTeamsAudit, groupsGranted);
    // A team whose owner left, and a team owned only by an account that is disabled.
    expect(rowOf(result, "Sales kickoff").cells).toMatchObject({
      kind: "Team",
      status: "No owners",
    });
    expect(rowOf(result, "Chromatography method transfer").cells).toMatchObject({
      status: "Only disabled owners",
      ownerNames: ["Harriet Moss"],
      enabledOwners: 0,
    });
    // One owner, and guests only that owner can answer for.
    expect(rowOf(result, "Quality system reviews").cells).toMatchObject({
      status: "One owner",
      owners: 1,
      guests: 3,
    });
    // Public, and quiet for 140 days.
    expect(rowOf(result, "Madison site move 2024").cells).toMatchObject({
      visibility: "Public",
      status: "Inactive 90 days",
      lastActivity: "2026-06-04T00:00:00.000Z",
    });
    // A Microsoft 365 group that expires in 12 days, renewed almost a year ago.
    const expiring = rowOf(result, "Spring conference planning").cells;
    expect(expiring.status).toBe("Expires within 30 days");
    expect(Date.parse(String(expiring.expires)) - now.getTime()).toBe(12 * 86_400_000);
    // A team the Teams report does not list.
    expect(rowOf(result, "Customer pilots").cells).toMatchObject({
      status: "No usage data",
      lastActivity: null,
    });
    // The group roles.ts gives Helpdesk Administrator, owned by two accounts from the roster.
    const assignable = result.rows.find((row) => row.key === guid("group:it-operations-admins"));
    expect(assignable?.cells).toMatchObject({
      roleAssignable: true,
      ownerNames: ["Demo administrator", "Tobias Vance"],
    });
    expect(rowOf(result, "Madison file share readers").cells.synced).toBe(true);
  });

  it("name owners from the roster, and only owners", async () => {
    const result = await run(groupsTeamsAudit, groupsGranted);
    const names = new Set(tenant.users.map((user) => user.displayName));
    for (const row of result.rows)
      for (const name of (row.cells.ownerNames as string[] | null) ?? [])
        expect(names.has(name), name).toBe(true);
    expect(requests.some((url) => url.pathname.includes("/members"))).toBe(false);
  });

  it("agree with the memberships a person's page shows", async () => {
    const result = await run(groupsTeamsAudit, groupsGranted);
    const admin = tenant.users.find((user) => user.displayName === "Dagny Holloway");
    const memberOf = (await (
      await graph(`https://graph.microsoft.com/v1.0/users/${admin?.id}/memberOf?$select=id`)
    ).json()) as { value: { id: string }[] };
    expect(memberOf.value.length).toBeGreaterThan(0);
    const listed = new Set(result.rows.map((row) => row.key));
    for (const group of memberOf.value) expect(listed.has(group.id), group.id).toBe(true);
  });

  it("are read without the usage reports too, and then no team has usage data", async () => {
    const result = await run(groupsTeamsAudit, ["Group.Read.All"]);
    expect(result.complete).toBe(true);
    expect(result.rows.filter((row) => row.cells.kind === "Team")).toSatisfy((teams: unknown[]) =>
      (teams as { cells: Record<string, unknown> }[]).every(
        (team) => team.cells.lastActivity === null,
      ),
    );
    expect(tile(result, "Inactive teams (usage not read)")).toBe(0);
  });

  it("are the same every time", async () => {
    const first = await run(groupsTeamsAudit, groupsGranted);
    const second = await run(groupsTeamsAudit, groupsGranted);
    expect(second.rows).toEqual(first.rows);
    expect(second.summary).toEqual(first.summary);
  });

  it("answer an owner list for an unknown group with 404, and serve nothing else under a group", async () => {
    const unknown = await graph(
      `https://graph.microsoft.com/v1.0/groups/${guid("group:nobody")}/owners`,
    );
    expect(unknown.status).toBe(404);
    const members = await graph(
      `https://graph.microsoft.com/v1.0/groups/${guid("group:staff")}/members`,
    );
    expect(members.status).toBe(404);
  });
});

describe("the synthetic mailboxes", () => {
  it("fill every status", async () => {
    const result = await run(mailboxStorageAudit, mailboxGranted);
    expect(result.complete).toBe(true);
    expect(result.rows).toHaveLength(49);
    expect(new Set(result.rows.map((row) => row.cells.status))).toEqual(new Set(MAILBOX_STATUSES));
    for (const label of [
      "Cannot send or receive",
      "Cannot send",
      "Over warning",
      "Near full",
      "Near full, no archive",
    ])
      expect(tile(result, label), label).toBe(1);
    expect(tile(result, "Total storage used (GB)")).toBe(316);
  });

  it("plant a mailbox for each status, on accounts the activity audit does not call quiet", async () => {
    const result = await run(mailboxStorageAudit, mailboxGranted);
    const status = (name: string) => rowOf(result, `${name}@${DOMAIN}`, "mailbox").cells.status;
    const research = (name: string) =>
      rowOf(result, `${name}@${RESEARCH_DOMAIN}`, "mailbox").cells.status;
    expect(research("lab.results")).toBe("Cannot send or receive");
    expect(status("mateus.brandt")).toBe("Cannot send");
    expect(research("dagny.holloway")).toBe("Over warning");
    expect(research("ingrid.solheim")).toBe("Near full");
    expect(status("kasimir.orne")).toBe("Near full, no archive");
    expect(research("wendell.aris")).toBe("Quota unknown");
    expect(rowOf(result, `wendell.aris@${RESEARCH_DOMAIN}`, "mailbox").cells).toMatchObject({
      warningQuota: null,
      sendQuota: null,
      sendReceiveQuota: null,
    });
  });

  it("name a shared mailbox and a room", async () => {
    const result = await run(mailboxStorageAudit, mailboxGranted);
    const types = result.rows.reduce<Record<string, number>>((count, row) => {
      const type = String(row.cells.recipientType);
      count[type] = (count[type] ?? 0) + 1;
      return count;
    }, {});
    expect(types).toEqual({ UserMailbox: 47, SharedMailbox: 1, RoomMailbox: 1 });
    expect(rowOf(result, "Lab Results (shared)", "displayName").cells.recipientType).toBe(
      "SharedMailbox",
    );
    expect(rowOf(result, "Lab Training Kiosk", "displayName").cells.recipientType).toBe(
      "RoomMailbox",
    );
  });

  it("are limited to the domains of the audit scope, and list both domains", async () => {
    const result = await run(mailboxStorageAudit, mailboxGranted, [RESEARCH_DOMAIN]);
    expect(result.rows).toHaveLength(18);
    expect(result.domains?.map((item) => item.domain)).toEqual([RESEARCH_DOMAIN, DOMAIN]);
    expect(result.notes.join("\n")).toContain(`Limited to accounts in: ${RESEARCH_DOMAIN}.`);
  });

  it("are read from the one report, once", async () => {
    requests.length = 0;
    await run(mailboxStorageAudit, mailboxGranted);
    const reports = requests.filter((url) => url.pathname.includes("/reports/"));
    expect(reports.map((url) => url.pathname.split("/").at(-1))).toEqual([
      "getMailboxUsageDetail(period='D90')",
    ]);
  });
});

describe("everything the two audits read", () => {
  it("belongs to the invented tenant, and only ever asks", async () => {
    bodies.length = 0;
    await run(groupsTeamsAudit, groupsGranted);
    await run(mailboxStorageAudit, mailboxGranted);
    expect(bodies.length).toBeGreaterThan(10);
    // The recorder refuses a recording with text from outside the tenant; so does this test.
    assertSynthetic(bodies.join("\n"), tenant.addresses);
  });

  it("are published with the permissions their packs carry", () => {
    expect(auditDefinition(groupsTeamsAudit)).toMatchObject({
      permissions: ["Group.Read.All"],
      optionalPermissions: ["Reports.Read.All", "ReportSettings.Read.All"],
    });
    expect(auditDefinition(mailboxStorageAudit)).toMatchObject({
      permissions: ["Reports.Read.All"],
      optionalPermissions: ["ReportSettings.Read.All"],
      scoped: true,
    });
  });
});

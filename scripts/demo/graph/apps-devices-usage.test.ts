import {
  type Audit,
  appConsentsAudit,
  m365AppUsageAudit,
  managedDevicesAudit,
} from "@prism365/audits";
import { GraphHttpClient } from "@prism365/graph";
import { scopeGraph } from "@prism365/modules";
import { describe, expect, it } from "vitest";
import {
  assertSynthetic,
  buildSyntheticTenant,
  createSyntheticGraph,
  SYNTHETIC_TENANT_ID,
} from "../tenant.ts";

// The synthetic tenant's devices, applications and usage report, read by the real audits through
// the real Graph client: the recorder does exactly this, so what these tests see is what the
// portfolio demo shows.
const now = new Date("2026-10-22T12:00:00Z");
const tenant = buildSyntheticTenant(now);
const graph = createSyntheticGraph({ now });
const bodies: string[] = [];
const recording: typeof fetch = async (input, init) => {
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
  const result = await audit.run({
    graph: scopeGraph(client, [...audit.packs, ...audit.optionalPacks]),
    now,
    granted: new Set(granted),
    signal: new AbortController().signal,
    scope: null,
    tenantId: SYNTHETIC_TENANT_ID,
  });
  return result;
}

const tile = (result: { summary: { label: string; value: number }[] }, label: string) =>
  result.summary.find((item) => item.label === label)?.value;

describe("the synthetic managed devices", () => {
  it("give most people two or three devices, with the findings the audit looks for", async () => {
    const result = await run(managedDevicesAudit, ["DeviceManagementManagedDevices.Read.All"]);
    const people = tenant.users.filter(
      (user) =>
        user.userType === "Member" && user.accountEnabled && user.assignedLicenses.length > 0,
    );
    expect(result.complete).toBe(true);
    expect(result.rows.length).toBeGreaterThanOrEqual(people.length * 2);
    expect(result.rows.length).toBeLessThanOrEqual(people.length * 3 + 2);
    expect(tile(result, "Jailbroken or rooted")).toBe(1);
    expect(tile(result, "Not compliant")).toBeGreaterThanOrEqual(3);
    expect(tile(result, "No check-in for 30 days")).toBe(2);
    expect(tile(result, "Not encrypted")).toBe(2);
    expect(tile(result, "Personal devices")).toBeGreaterThanOrEqual(2);
  });

  it("includes the personal laptop with no encryption and shared machines with no owner", async () => {
    const result = await run(managedDevicesAudit, ["DeviceManagementManagedDevices.Read.All"]);
    const laptop = result.rows.find(
      (row) => row.tags.includes("personal") && row.tags.includes("not-encrypted"),
    );
    expect(laptop?.cells).toMatchObject({ os: "Windows", compliance: "Not compliant" });
    const shared = result.rows.filter((row) => row.personId === null);
    expect(shared.map((row) => row.cells.device).sort()).toEqual([
      "LARK-KIOSK-01",
      "LARK-KIOSK-02",
    ]);
    // Every other device belongs to someone in the directory.
    const ids = new Set(tenant.users.map((user) => user.id));
    for (const row of result.rows.filter((item) => item.personId !== null))
      expect(ids.has(row.personId ?? ""), String(row.cells.device)).toBe(true);
  });

  it("are the same every time", async () => {
    const first = await run(managedDevicesAudit, ["DeviceManagementManagedDevices.Read.All"]);
    const second = await run(managedDevicesAudit, ["DeviceManagementManagedDevices.Read.All"]);
    expect(second.rows).toEqual(first.rows);
  });
});

describe("the synthetic applications", () => {
  const granted = ["Application.Read.All", "Directory.Read.All"];

  it("hold the permissions the audit is meant to find", async () => {
    const result = await run(appConsentsAudit, granted);
    expect(result.complete).toBe(true);
    const names = result.rows.map((row) => row.cells.app);
    expect(result.rows.length).toBeGreaterThanOrEqual(8);
    expect(result.rows.length).toBeLessThanOrEqual(12);
    // An unverified app with the Mail.Read application permission.
    const quillnet = result.rows.find((row) => row.cells.app === "Quillnet Mail Archiver");
    expect(quillnet).toMatchObject({
      tags: expect.arrayContaining(["unverified", "app-high-risk"]),
      cells: {
        appPermissions: ["Mail.Read"],
        publisher: "Unverified",
        owner: "Another organisation",
      },
    });
    // An app with admin consent for everyone, and some that people consented to themselves.
    expect(tile(result, "Apps with admin consent for everyone")).toBeGreaterThanOrEqual(2);
    expect(tile(result, "Apps consented by individual users")).toBeGreaterThanOrEqual(2);
    expect(tile(result, "Unverified apps with high-risk permissions")).toBeGreaterThanOrEqual(2);
    expect(tile(result, "Apps with high-risk application permissions")).toBeGreaterThanOrEqual(3);
    // An app of the tenant's own, and one of Microsoft's with a permission granted here; Microsoft's
    // others, and an app with nothing, are not listed.
    expect(result.rows.find((row) => row.cells.app === "Larkspur Lab Portal")?.cells.owner).toBe(
      "This tenant",
    );
    expect(
      result.rows.find((row) => row.cells.app === "Exchange Hybrid Agent")?.cells,
    ).toMatchObject({
      owner: "Microsoft",
      appPermissions: ["full_access_as_app"],
    });
    expect(names).not.toContain("Driftwood Notes Sync");
    expect(names).not.toContain("Office 365 Portal");
    expect(names).not.toContain("Microsoft Graph");
    expect(names).not.toContain("lab-pipeline-identity");
  });

  it("resolve every permission to a name", async () => {
    const result = await run(appConsentsAudit, granted);
    for (const row of result.rows)
      for (const name of row.cells.appPermissions as string[])
        expect(name, String(row.cells.app)).toMatch(/^[A-Za-z][A-Za-z0-9_.]+$/);
    expect(result.notes.join(" ")).not.toContain("could not be matched");
  });
});

describe("the synthetic usage report", () => {
  const granted = ["Reports.Read.All", "ReportSettings.Read.All"];

  it("shows names and has accounts for every tile", async () => {
    const result = await run(m365AppUsageAudit, granted);
    expect(result.complete).toBe(true);
    expect(result.rows.length).toBeGreaterThan(20);
    expect(tile(result, "No Microsoft 365 app activity in 90 days")).toBeGreaterThanOrEqual(5);
    expect(tile(result, "Desktop apps used")).toBeGreaterThan(10);
    expect(tile(result, "Web and mobile only")).toBeGreaterThanOrEqual(2);
    expect(result.rows.some((row) => row.tags.includes("teams-only"))).toBe(true);
    expect(result.rows.some((row) => row.tags.includes("web-only"))).toBe(true);
    expect(result.notes.join(" ")).not.toContain("concealed");
  });

  it("lists as idle the accounts the inactive accounts audit lists", async () => {
    const result = await run(m365AppUsageAudit, granted);
    const idle = new Set(
      result.rows
        .filter((row) => row.tags.includes("idle"))
        .map((row) => (row.cells.person as { principal: string }).principal),
    );
    const silent = tenant.users.filter((user) => {
      const last = user.signInActivity.lastSuccessfulSignInDateTime;
      return (
        user.userType === "Member" &&
        user.accountEnabled &&
        user.assignedLicenses.length > 0 &&
        (last === null || now.getTime() - Date.parse(last) >= 90 * 86_400_000)
      );
    });
    expect(silent.length).toBeGreaterThanOrEqual(5);
    for (const user of silent)
      expect(idle.has(user.userPrincipalName), user.displayName).toBe(true);
    // And nobody who signed in this month.
    for (const user of tenant.users) {
      const last = user.signInActivity.lastSuccessfulSignInDateTime;
      if (last && now.getTime() - Date.parse(last) < 30 * 86_400_000)
        expect(idle.has(user.userPrincipalName), user.displayName).toBe(false);
    }
  });
});

describe("everything the three audits read", () => {
  it("belongs to the invented tenant", async () => {
    bodies.length = 0;
    await run(managedDevicesAudit, ["DeviceManagementManagedDevices.Read.All"]);
    await run(appConsentsAudit, ["Application.Read.All", "Directory.Read.All"]);
    await run(m365AppUsageAudit, ["Reports.Read.All", "ReportSettings.Read.All"]);
    expect(bodies.length).toBeGreaterThan(3);
    // The recorder refuses a recording with text from outside the tenant; so does this test.
    assertSynthetic(bodies.join("\n"), tenant.addresses);
  });
});

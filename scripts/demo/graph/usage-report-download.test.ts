import { GraphHttpClient } from "@prism365/graph";
import { describe, expect, it } from "vitest";
import { assertSynthetic, buildSyntheticTenant, createSyntheticGraph } from "../tenant.ts";
import { ACTIVE_USER_COLUMNS } from "./usage-reports.ts";

// The synthetic tenant's report download host, read by the real Graph client the way a live
// tenant's would be: Graph answers 302 with a link on reports.office.com, and the client fetches
// the CSV from there without the token.
const now = new Date("2026-10-22T12:00:00Z");
const tenant = buildSyntheticTenant(now);
const usage = { feature: "reports.usage.read", permission: "Reports.Read.All" } as const;
const REPORT = "reports/getOffice365ActiveUserDetail(period='D30')";

function reader() {
  const graph = createSyntheticGraph({ now });
  const sent: { host: string; authorization: string | null }[] = [];
  const recording: typeof fetch = async (input, init) => {
    sent.push({
      host: new URL(String(input)).hostname,
      authorization: new Headers(init?.headers).get("Authorization"),
    });
    return graph(input, init);
  };
  const client = new GraphHttpClient({
    getAccessToken: async () => "synthetic",
    fetchImplementation: recording,
    maxRetries: 0,
  });
  return { client, graph, sent };
}

describe("the synthetic usage report downloads", () => {
  it("answer the active users report with a redirect the real client follows without the token", async () => {
    const { client, sent } = reader();
    const table = await client.getReport(REPORT, usage);
    expect(sent).toEqual([
      { host: "graph.microsoft.com", authorization: "Bearer synthetic" },
      { host: "reports.office.com", authorization: null },
    ]);
    expect(table.header).toEqual([...ACTIVE_USER_COLUMNS]);
    expect(table).toMatchObject({ skippedRows: 0, truncated: false });
    const licensed = tenant.users.filter(
      (user) => user.userType === "Member" && user.assignedLicenses.length > 0,
    );
    expect(table.rows).toHaveLength(licensed.length);
    const byPrincipal = new Map(licensed.map((user) => [user.userPrincipalName, user]));
    for (const row of table.rows) {
      const user = byPrincipal.get(row["User Principal Name"] ?? "");
      expect(user, row["User Principal Name"]).toBeDefined();
      expect(row["Display Name"]).toBe(user?.displayName);
      expect(row["Is Deleted"]).toBe("False");
      expect(row["Has Exchange License"]).toMatch(/^(True|False)$/);
      // The report is two days behind the tenant's clock.
      expect(row["Report Refresh Date"]).toBe("2026-10-20");
      if (row["Exchange Last Activity Date"])
        expect(row["Exchange Last Activity Date"] <= "2026-10-20").toBe(true);
    }
    expect(table.rows.some((row) => row["Assigned Products"]?.includes("+"))).toBe(true);
    expect(table.rows.some((row) => row["Teams Last Activity Date"] === "")).toBe(true);
  });

  it("serve every period, each link once, and nothing for an unknown link", async () => {
    const { graph } = reader();
    for (const period of ["D7", "D30", "D90", "D180"]) {
      const redirect = await graph(
        `https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='${period}')`,
      );
      expect(redirect.status, period).toBe(302);
      const link = redirect.headers.get("Location") ?? "";
      expect(link).toMatch(/^https:\/\/reports\.office\.com\/data\/download\/[A-Za-z0-9_-]{32}$/);
      const first = await graph(link);
      expect(first.status).toBe(200);
      expect(first.headers.get("Content-Type")).toBe("application/octet-stream");
      // text() would drop the byte-order mark, so the bytes are checked.
      const bytes = new Uint8Array(await first.arrayBuffer());
      expect([...bytes.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
      const text = new TextDecoder().decode(bytes);
      expect(text.startsWith("Report Refresh Date,User Principal Name,")).toBe(true);
      // The recording is public, so the report holds nothing from outside the invented tenant.
      assertSynthetic(text, tenant.addresses);
      expect((await graph(link)).status).toBe(404);
    }
    expect((await graph("https://reports.office.com/data/download/unknown")).status).toBe(404);
    expect((await graph("https://reports.office.com/other/path")).status).toBe(404);
  });

  it("refuse a download that carries a token, and keep the link for one without", async () => {
    const { graph } = reader();
    const redirect = await graph(`https://graph.microsoft.com/v1.0/${REPORT}`);
    const link = redirect.headers.get("Location") ?? "";
    const leaked = await graph(link, { headers: { Authorization: "Bearer synthetic" } });
    expect(leaked.status).toBe(400);
    expect((await graph(link)).status).toBe(200);
  });

  it("keep the JSON form of the Microsoft 365 apps report", async () => {
    const { graph } = reader();
    const response = await graph(
      "https://graph.microsoft.com/v1.0/reports/getM365AppUserDetail(period='D90')?$format=application/json",
    );
    expect(response.status).toBe(200);
    expect(response.headers.get("Content-Type")).toBe("application/json");
    const body = (await response.json()) as { value: unknown[] };
    expect(body.value.length).toBeGreaterThan(20);
  });
});

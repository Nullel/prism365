import {
  type Audit,
  type AuditContext,
  createConditionalAccessAudit,
  createRiskyUsersAudit,
} from "@prism365/audits";
import { GraphHttpClient } from "@prism365/graph";
import { scopeGraph } from "@prism365/modules";
import { describe, expect, it } from "vitest";
import {
  assertSynthetic,
  buildSyntheticTenant,
  createSyntheticGraph,
  RESEARCH_DOMAIN,
} from "../tenant.ts";

// The synthetic tenant's Conditional Access and risky user data, read by the real audits: the
// planted findings must keep showing, and nothing in the answers may come from outside the
// invented tenant.

const anchor = new Date("2026-10-20T12:00:00.000Z");
const tenant = buildSyntheticTenant(anchor);
const granted = [
  "Policy.Read.All",
  "User.Read.All",
  "LicenseAssignment.Read.All",
  "IdentityRiskyUser.Read.All",
];

function context(audit: Audit, scope: ReadonlySet<string> | null = null): AuditContext {
  const client = new GraphHttpClient({
    getAccessToken: async () => "synthetic",
    fetchImplementation: createSyntheticGraph({ now: anchor }),
    sleep: async () => {},
  });
  return {
    graph: scopeGraph(client, [...audit.packs, ...audit.optionalPacks]),
    now: anchor,
    granted: new Set(granted),
    signal: new AbortController().signal,
    scope,
  };
}

const noWait = async () => {};
const tile = (summary: { label: string; value: number }[], label: string) =>
  summary.find((item) => item.label === label);

describe("the synthetic Conditional Access policies", () => {
  const audit = createConditionalAccessAudit({ pause: noWait });

  it("plant a report-only legacy block, MFA with two exclusions and an inactive policy", async () => {
    const result = await audit.run(context(audit));
    expect(result.rows).toHaveLength(6);
    expect(result.complete).toBe(true);
    expect(tile(result.summary, "Policies enabled")?.value).toBe(4);
    expect(tile(result.summary, "Policies in report-only")?.value).toBe(1);
    expect(tile(result.summary, "Policies off")?.value).toBe(1);
    expect(tile(result.summary, "No enabled policy blocks legacy authentication")).toMatchObject({
      value: 1,
      tone: "critical",
    });
    expect(
      tile(result.summary, "No enabled policy requires MFA for administrators"),
    ).toBeUndefined();
    expect(tile(result.summary, "Users excluded from MFA policies")?.value).toBe(2);
  });

  it("name the two excluded accounts and the named locations", async () => {
    const result = await audit.run(context(audit));
    const cells = (name: string) => result.rows.find((row) => row.cells.policy === name)?.cells;
    expect(cells("Require MFA for all users")?.excludes).toEqual([
      "2 users: Demo administrator, Tobias Vance",
    ]);
    expect(cells("Block sign-ins from outside approved countries")?.conditions).toEqual([
      "Locations: any location, except Approved countries",
    ]);
    expect(cells("Require a compliant device for Office 365")).toMatchObject({
      state: "Off",
      conditions: [
        "Platforms: Windows, macOS",
        "Locations: any location, except Madison laboratory",
      ],
    });
    expect(cells("Require phishing-resistant MFA for admin portals")).toMatchObject({
      grant: "Require authentication strength Phishing-resistant MFA",
      session: "Sign in every 4 hours; Persistent browser: never",
    });
  });

  it("read the same on every run and stay inside the invented tenant", async () => {
    const first = await audit.run(context(audit));
    const second = await audit.run(context(audit));
    expect(second).toEqual(first);
    expect(() => assertSynthetic(JSON.stringify(first), tenant.addresses)).not.toThrow();
  });
});

describe("the synthetic risky users", () => {
  const audit = createRiskyUsersAudit({ pause: noWait });

  it("plant one high risk and one medium risk user still open and one recently remediated", async () => {
    const result = await audit.run(context(audit));
    expect(result.rows).toHaveLength(4);
    expect(result.complete).toBe(true);
    expect(tile(result.summary, "High risk, still at risk")).toMatchObject({
      value: 1,
      tone: "critical",
    });
    expect(tile(result.summary, "Medium risk, still at risk")).toMatchObject({
      value: 1,
      tone: "warning",
    });
    expect(tile(result.summary, "Remediated or dismissed in the last 30 days")?.value).toBe(1);
    expect(tile(result.summary, "Confirmed compromised")?.value).toBe(0);
  });

  it("put the medium risk user on the research domain, which a scope can leave out", async () => {
    const all = await audit.run(context(audit));
    const research = all.rows.filter((row) => row.cells.domain === RESEARCH_DOMAIN);
    expect(research).toHaveLength(1);
    const scoped = await audit.run(
      context(
        audit,
        new Set([all.domains?.find((item) => item.domain !== RESEARCH_DOMAIN)?.domain ?? ""]),
      ),
    );
    expect(scoped.rows).toHaveLength(3);
    expect(tile(scoped.summary, "Medium risk, still at risk")?.value).toBe(0);
  });

  it("read the same on every run and stay inside the invented tenant", async () => {
    const first = await audit.run(context(audit));
    expect(await audit.run(context(audit))).toEqual(first);
    expect(() => assertSynthetic(JSON.stringify(first), tenant.addresses)).not.toThrow();
  });
});

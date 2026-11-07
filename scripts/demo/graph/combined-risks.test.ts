import {
  type Audit,
  type AuditResult,
  builtInAudits,
  combinedRisksAudit,
  INPUT_IDS,
  INPUTS,
  RULES,
  type StoredResults,
} from "@prism365/audits";
import { GraphHttpClient } from "@prism365/graph";
import { scopeGraph } from "@prism365/modules";
import { beforeAll, describe, expect, it } from "vitest";
import { combinedRisksFor } from "../../../apps/api/src/person-access.ts";
import { createSyntheticGraph, SYNTHETIC_TENANT_ID } from "../tenant.ts";

// The synthetic tenant through the real input audits and then Combined risks, the way the recorder
// runs them: the demo must show that audit tripping a good spread of its rules, so a change to the
// synthetic accounts that quietly empties it fails here and not in a recording.
const now = new Date("2026-10-22T12:00:00Z");
const graph = createSyntheticGraph({ now });
const granted = new Set([
  "User.Read.All",
  "AuditLog.Read.All",
  "Reports.Read.All",
  "ReportSettings.Read.All",
  "LicenseAssignment.Read.All",
  "Policy.Read.All",
  "IdentityRiskyUser.Read.All",
  "DeviceManagementManagedDevices.Read.All",
  "RoleManagement.Read.Directory",
  "Group.Read.All",
  "Domain.Read.All",
]);

async function runAudit(audit: Audit, stored?: StoredResults): Promise<AuditResult> {
  const client = new GraphHttpClient({
    getAccessToken: async () => "synthetic",
    fetchImplementation: graph,
    maxResponseBytes: 8_388_608,
    maxRetries: 0,
    sleep: async () => {},
  });
  const signal = new AbortController().signal;
  return audit.run({
    graph: scopeGraph(client, [...audit.packs, ...audit.optionalPacks], () => {}, signal),
    now,
    granted,
    signal,
    scope: null,
    tenantId: SYNTHETIC_TENANT_ID,
    ...(stored ? { stored } : {}),
  });
}

let combined: AuditResult;
beforeAll(async () => {
  const results = new Map<string, AuditResult>();
  for (const id of INPUT_IDS) {
    const audit = builtInAudits.find((item) => item.definition.id === INPUTS[id].auditId);
    if (!audit) throw new Error(`No audit ${INPUTS[id].auditId}`);
    results.set(audit.definition.id, await runAudit(audit));
  }
  const stored: StoredResults = {
    async latestRun(auditId) {
      const result = results.get(auditId);
      return result
        ? {
            finishedAt: now.toISOString(),
            complete: result.complete,
            scope: null,
            rows: result.rows,
          }
        : null;
    },
    async openFindings() {
      return [];
    },
  };
  combined = await runAudit(combinedRisksAudit, stored);
}, 120_000);

const labels = (result: AuditResult) =>
  new Set(result.rows.flatMap((row) => (row.cells.rules as string[]) ?? []));

describe("Combined risks on the synthetic tenant", () => {
  it("trips at least six of its nine rules, with every input read", () => {
    const tripped = RULES.filter((rule) => labels(combined).has(rule.label));
    expect(tripped.map((rule) => rule.id)).toEqual(
      expect.arrayContaining([
        "admin-unprotected",
        "dormant-admin",
        "guest-with-power",
        "admin-weak-device",
        "disabled-still-holding",
      ]),
    );
    expect(tripped.length).toBeGreaterThanOrEqual(6);
    expect(combined.summary.find((item) => item.label === "Rules not checked")?.value).toBe(0);
  });

  it("counts the group a disabled account still owns, from Groups and Teams", () => {
    const moss = combined.rows.find(
      (row) => (row.cells.person as { name: string | null }).name === "Harriet Moss",
    );
    expect(moss?.cells.reasons).toContain("ownership of 1 group");
    expect(combined.notes.join("\n")).toContain(
      "Owning a group is checked against Groups and Teams",
    );
  });

  it("lists the worst accounts first, with the audits behind each", () => {
    const sources = (combined.rows[0]?.cells.sources ?? []) as string[];
    expect(combined.rows[0]?.cells.severity).toBe("Critical");
    expect(sources.length).toBeGreaterThan(0);
    expect(combined.complete).toBe(true);
  });

  it("shows Tobias Vance his critical combination on the person page, and a clean account none", () => {
    const run = {
      finishedAt: now.toISOString(),
      complete: combined.complete,
      notes: combined.notes,
      scope: null,
      rows: combined.rows,
    };
    const tobias = combined.rows.find(
      (row) => (row.cells.person as { name: string | null }).name === "Tobias Vance",
    );
    expect(tobias?.personId).toBeTruthy();
    const his = combinedRisksFor(run, tobias?.personId ?? "");
    expect(his.severity).toBe("critical");
    expect(his.rules.map((rule) => rule.id)).toContain("admin-unprotected");
    expect(his.reasons).toContain("Global Administrator");
    // An account the run does not list trips nothing.
    const clean = combinedRisksFor(run, "ffffffff-0000-4000-8000-00000000ffff");
    expect(clean).toMatchObject({ severity: null, rules: [], reasons: "" });
  });
});

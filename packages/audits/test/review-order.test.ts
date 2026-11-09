import {
  AUDIT_FAMILIES,
  auditFamilyOf,
  COMBINED_RISKS_AUDIT_ID,
  groupByFamily,
} from "@prism365/contracts";
import { describe, expect, it } from "vitest";
import {
  ACTIVITY_REPORT_COUNT,
  builtInAudits,
  MAX_REPORT_DOWNLOADS,
  REVIEW_ORDER,
  type ReviewStage,
} from "../src/index.ts";

const ids = builtInAudits.map((audit) => audit.definition.id);
const stageRank: Record<ReviewStage, number> = { directory: 0, paced: 1, reports: 2, summary: 3 };

describe("the review order", () => {
  it("lists every built-in audit once and no other", () => {
    const listed = REVIEW_ORDER.map((item) => item.auditId);
    expect([...listed].sort()).toEqual([...ids].sort());
    expect(new Set(listed).size).toBe(listed.length);
  });

  it("runs the directory audits, then the paced ones, then the reports, then the summary", () => {
    const ranks = REVIEW_ORDER.map((item) => stageRank[item.stage]);
    expect(ranks).toEqual([...ranks].sort((a, b) => a - b));
    expect(new Set(REVIEW_ORDER.map((item) => item.stage))).toEqual(
      new Set(["directory", "paced", "reports", "summary"]),
    );
  });

  it("runs the audits that read stored results in the summary stage, Combined risks last", () => {
    const reading = builtInAudits
      .filter((audit) => audit.usesStoredResults)
      .map((audit) => audit.definition.id)
      .sort();
    const summary = REVIEW_ORDER.filter((item) => item.stage === "summary")
      .map((item) => item.auditId)
      .sort();
    expect(summary).toEqual(reading);
    // Combined risks joins every other audit, Tenant settings among them, so it runs last.
    const joining = builtInAudits.filter(
      (audit) => audit.usesStoredResults && audit.packs.length === 0,
    );
    expect(joining).toHaveLength(1);
    expect(REVIEW_ORDER.at(-1)?.auditId).toBe(joining[0]?.definition.id);
  });

  it("runs Tenant settings after the Conditional Access run it reads, so both come from one review", () => {
    const ids = REVIEW_ORDER.map((item) => item.auditId);
    const settings = REVIEW_ORDER.find((item) => item.auditId === "tenant.settings");
    expect(settings).toMatchObject({ stage: "summary", reportDownloads: 0 });
    expect(ids.indexOf("tenant.settings")).toBeGreaterThan(
      ids.indexOf("identity.conditional-access"),
    );
    expect(ids.indexOf("tenant.settings")).toBeLessThan(ids.indexOf(COMBINED_RISKS_AUDIT_ID));
    expect(auditFamilyOf("tenant.settings")?.id).toBe("tenant-setup");
    expect(auditFamilyOf("licenses.capacity")?.id).toBe("tenant-setup");
  });

  it("puts an audit in the reports stage exactly when its packs include the usage reports", () => {
    for (const audit of builtInAudits) {
      const step = REVIEW_ORDER.find((item) => item.auditId === audit.definition.id);
      const reads = [...audit.packs, ...audit.optionalPacks].includes("reports.usage.read");
      expect(step?.stage === "reports", audit.definition.id).toBe(reads);
      expect((step?.reportDownloads ?? 0) > 0, audit.definition.id).toBe(reads);
    }
  });

  it("puts an audit in the paced stage exactly when it requires a throttled pack", () => {
    const throttled = ["identity.conditional-access.read", "identity.risk.read"];
    for (const audit of builtInAudits) {
      const step = REVIEW_ORDER.find((item) => item.auditId === audit.definition.id);
      expect(step?.stage === "paced", audit.definition.id).toBe(
        audit.packs.some((pack) => throttled.includes(pack)),
      );
    }
  });

  it("budgets Account activity at every report it can download, and fits one review in the limit", () => {
    expect(REVIEW_ORDER.find((item) => item.auditId === "accounts.activity")?.reportDownloads).toBe(
      ACTIVITY_REPORT_COUNT,
    );
    const total = REVIEW_ORDER.reduce((sum, item) => sum + item.reportDownloads, 0);
    expect(total).toBeLessThanOrEqual(MAX_REPORT_DOWNLOADS);
  });
});

describe("the audit families", () => {
  it("puts every built-in audit in exactly one family", () => {
    for (const id of ids) {
      const owners = AUDIT_FAMILIES.filter((family) =>
        family.audits.some((entry) => entry.auditId === id),
      );
      expect(
        owners.map((family) => family.id),
        `${id} belongs to one family`,
      ).toHaveLength(1);
    }
  });

  it("lists no audit twice in a family and no audit this build does not have", () => {
    const listed = AUDIT_FAMILIES.flatMap((family) => family.audits.map((entry) => entry.auditId));
    expect(new Set(listed).size).toBe(listed.length);
    for (const id of listed) expect(ids, id).toContain(id);
  });

  it("has one main audit at most per family, and narrower audits only after it", () => {
    for (const family of AUDIT_FAMILIES) {
      const roles = family.audits.map((entry) => entry.role);
      expect(roles.filter((role) => role === "main").length).toBeLessThanOrEqual(1);
      const narrower = roles.indexOf("narrower");
      if (narrower >= 0) expect(roles.indexOf("main")).toBeGreaterThanOrEqual(0);
      if (narrower >= 0) expect(roles.indexOf("main")).toBeLessThan(narrower);
    }
  });

  it("names the family of an audit", () => {
    expect(auditFamilyOf("accounts.inactive")?.id).toBe("account-lifecycle");
    expect(auditFamilyOf("not.an-audit")).toBeUndefined();
  });

  it("groups a catalog by family and keeps an audit no family lists", () => {
    const items = [...ids, "extra.audit"].map((id) => ({ definition: { id } }));
    const groups = groupByFamily(items);
    expect(
      groups.flatMap((group) => group.entries.map((e) => e.item.definition.id)).sort(),
    ).toEqual([...ids, "extra.audit"].sort());
    expect(groups.at(-1)?.family.id).toBe("other");
    expect(groups[0]?.family.id).toBe("summary");
  });
});

import type { StoredTrendPoint } from "@prism365/audits";
import { type AlertRule, alertRuleSchema } from "@prism365/contracts";
import type { FindingNotification } from "@prism365/modules";
import { describe, expect, it } from "vitest";
import {
  auditAlert,
  auditFires,
  auditMovement,
  collectionAlert,
  collectionStoppedTooLong,
  exportAlert,
  findingAlert,
  findingMatches,
  weeklySummaryAlert,
  weeklySummarySlot,
} from "../src/alert-rules.ts";
import { exportWithheld } from "../src/delivery.ts";

// What each rule kind fires on. Synthetic findings and trend points only.

const tenantId = "11111111-1111-4111-8111-111111111111";
const base = {
  id: "aaaaaaaa-0000-4000-8000-000000000001",
  name: "Synthetic rule",
  enabled: true,
  destinations: ["bbbbbbbb-0000-4000-8000-000000000001"],
  quietMinutes: 60,
  createdAt: "2026-10-24T00:00:00.000Z",
  createdBy: "22222222-2222-4222-8222-222222222222",
};
const rule = (fields: Record<string, unknown>): AlertRule =>
  alertRuleSchema.parse({ ...base, ...fields });

function finding(overrides: Partial<FindingNotification> = {}): FindingNotification {
  return {
    tenantId,
    module: "signins.audit",
    findingId: "cccccccc-0000-4000-8000-000000000001",
    ruleId: "signins.unfamiliar_signin",
    severity: "high",
    title: "Sign-in that does not match this account's history",
    summary: "Avery Stone signed in from somewhere new.",
    subjectLabel: "Avery Stone",
    change: "created",
    observedAt: "2026-10-24T10:00:00.000Z",
    ...overrides,
  };
}

describe("finding, directory and gap rules", () => {
  it("matches a finding at or above the severity, in the chosen modules", () => {
    const high = rule({ kind: "finding", minimumSeverity: "high", modules: [] });
    expect(findingMatches(high, finding())).toBe(true);
    expect(findingMatches(high, finding({ severity: "critical" }))).toBe(true);
    expect(findingMatches(high, finding({ severity: "medium" }))).toBe(false);
    const signins = rule({ kind: "finding", minimumSeverity: "low", modules: ["signins.audit"] });
    expect(findingMatches(signins, finding())).toBe(true);
    expect(findingMatches(signins, finding({ module: "directory.changes" }))).toBe(false);
    expect(findingMatches({ ...high, enabled: false }, finding())).toBe(false);
  });

  it("matches a directory change only by its rule id", () => {
    const roles = rule({ kind: "directory", ruleIds: ["role.privileged-assigned"] });
    const change = finding({ module: "directory.changes", ruleId: "role.privileged-assigned" });
    expect(findingMatches(roles, change)).toBe(true);
    expect(findingMatches(roles, { ...change, ruleId: "app.consent-granted" })).toBe(false);
    // The same rule id from another module is not a directory change.
    expect(findingMatches(roles, { ...change, module: "signins.audit" })).toBe(false);
  });

  it("matches a gap finding from any module only when gaps are watched", () => {
    const gaps = rule({ kind: "collection", stoppedHours: null, gaps: true });
    expect(findingMatches(gaps, finding({ ruleId: "signins.collection_gap" }))).toBe(true);
    expect(
      findingMatches(gaps, finding({ module: "directory.changes", ruleId: "collection.gap" })),
    ).toBe(true);
    expect(findingMatches(gaps, finding())).toBe(false);
    const stoppedOnly = rule({ kind: "collection", stoppedHours: 6, gaps: false });
    expect(findingMatches(stoppedOnly, finding({ ruleId: "signins.collection_gap" }))).toBe(false);
  });

  it("never matches a finding with an audit rule", () => {
    const audit = rule({
      kind: "audit",
      auditId: "accounts.inactive",
      tileId: "inactive",
      condition: "above",
      value: 1,
    });
    expect(findingMatches(audit, finding({ severity: "critical" }))).toBe(false);
  });

  it("puts the worst finding first and counts every one", () => {
    const alert = findingAlert(rule({ kind: "finding", minimumSeverity: "low", modules: [] }), [
      finding({ findingId: "1", severity: "low" }),
      finding({ findingId: "2", severity: "critical", change: "escalated" }),
      finding({ findingId: "3", severity: "medium" }),
    ]);
    expect(alert.count).toBe(3);
    expect(alert.severity).toBe("critical");
    expect(alert.items.map((item) => item.severity)).toEqual(["critical", "medium", "low"]);
    expect(alert.items[0]?.note).toBe("escalated");
    expect(alert.link).toBe("/#/signins");
    expect(alert.title).toBe("3 new or escalated findings, the worst at critical severity");
  });
});

function point(overrides: Partial<StoredTrendPoint> & { value?: number } = {}): StoredTrendPoint {
  const { value = 5, ...rest } = overrides;
  return {
    runId: `run-${Math.random()}`,
    finishedAt: "2026-10-24T00:00:00.000Z",
    version: 1,
    complete: true,
    scope: null,
    trigger: "schedule",
    tiles: [{ id: "inactive", label: "Inactive accounts", value, tone: "warning" }],
    ...rest,
  };
}

describe("audit rules", () => {
  const above = rule({
    kind: "audit",
    auditId: "accounts.inactive",
    tileId: "inactive",
    condition: "above",
    value: 10,
  }) as Extract<AlertRule, { kind: "audit" }>;

  it("compares with the newest earlier point of the same version, scope and completeness", () => {
    const points = [
      point({ value: 3 }),
      point({ value: 50, version: 2 }),
      point({ value: 60, scope: ["example.invalid"] }),
      point({ value: 70, complete: false }),
      point({ value: 12 }),
    ];
    expect(auditMovement(points, "inactive")).toEqual({ value: 12, previous: 3 });
  });

  it("does not compare a partial newest run, or a tile the newest run lacks", () => {
    expect(
      auditMovement([point({ value: 3 }), point({ value: 12, complete: false })], "inactive"),
    ).toBeNull();
    expect(auditMovement([point(), point()], "missing")).toBeNull();
  });

  it("has no previous value when nothing earlier is comparable", () => {
    expect(
      auditMovement([point({ value: 3, version: 2 }), point({ value: 12 })], "inactive"),
    ).toEqual({
      value: 12,
      previous: null,
    });
  });

  it("fires above and below only when the value crosses the threshold", () => {
    expect(auditFires(above, { value: 12, previous: 3 })).toBe(true);
    expect(auditFires(above, { value: 12, previous: 11 })).toBe(false);
    expect(auditFires(above, { value: 10, previous: 3 })).toBe(false);
    expect(auditFires(above, { value: 11, previous: 10 })).toBe(true);
    expect(auditFires(above, { value: 12, previous: null })).toBe(true);
    const below = { ...above, condition: "below" as const, value: 5 };
    expect(auditFires(below, { value: 4, previous: 9 })).toBe(true);
    expect(auditFires(below, { value: 4, previous: 4 })).toBe(false);
    expect(auditFires(below, { value: 5, previous: 9 })).toBe(false);
  });

  it("fires a change rule on a move of at least the threshold, either way", () => {
    const change = { ...above, condition: "change" as const, value: 5 };
    expect(auditFires(change, { value: 15, previous: 10 })).toBe(true);
    expect(auditFires(change, { value: 5, previous: 10 })).toBe(true);
    expect(auditFires(change, { value: 14, previous: 10 })).toBe(false);
    expect(auditFires(change, { value: 15, previous: null })).toBe(false);
    expect(auditFires({ ...change, enabled: false }, { value: 50, previous: 10 })).toBe(false);
  });

  it("says what moved and links to the audit", () => {
    const alert = auditAlert(
      above,
      { title: "Inactive accounts", tileLabel: "Inactive" },
      {
        value: 12,
        previous: 3,
      },
    );
    expect(alert.title).toBe("Inactive accounts: Inactive is 12");
    expect(alert.lines).toContain("It was 3 at the previous comparable run.");
    expect(alert.link).toBe("/#/audits/accounts.inactive");
  });
});

describe("collection rules", () => {
  const stopped = rule({ kind: "collection", stoppedHours: 6, gaps: false });
  const now = new Date("2026-10-24T12:00:00.000Z");
  const state = (overrides = {}) => ({
    intended: true,
    collecting: false,
    stoppedSince: new Date("2026-10-24T05:00:00.000Z"),
    reason: "signed_out" as const,
    ...overrides,
  });

  it("fires when collection meant to be on has been stopped long enough", () => {
    expect(collectionStoppedTooLong(stopped, state(), now)).toBe(true);
    expect(
      collectionStoppedTooLong(
        stopped,
        state({ stoppedSince: new Date("2026-10-24T06:30:00Z") }),
        now,
      ),
    ).toBe(false);
  });

  it("stays quiet when collection is on, or was stopped on purpose", () => {
    expect(collectionStoppedTooLong(stopped, state({ collecting: true }), now)).toBe(false);
    expect(collectionStoppedTooLong(stopped, state({ intended: false }), now)).toBe(false);
    expect(collectionStoppedTooLong(stopped, state({ stoppedSince: null }), now)).toBe(false);
    const gapsOnly = rule({ kind: "collection", stoppedHours: null, gaps: true });
    expect(collectionStoppedTooLong(gapsOnly, state(), now)).toBe(false);
  });

  it("says for how long and why", () => {
    const alert = collectionAlert(
      stopped as Extract<AlertRule, { kind: "collection" }>,
      state(),
      now,
    );
    expect(alert.lines[0]).toContain("stopped for 7 hours");
    expect(alert.lines[1]).toBe("The administrator it collects for signed out.");
  });
});

describe("the weekly summary", () => {
  it("is due from Monday 07:00 in the time zone", () => {
    // Wednesday 2026-10-21 in UTC: the slot is Monday 2026-10-19 07:00.
    expect(weeklySummarySlot(new Date("2026-10-21T12:00:00Z"), "UTC").toISOString()).toBe(
      "2026-10-19T07:00:00.000Z",
    );
    // Monday 06:59 is still last week's slot; 07:00 is this week's.
    expect(weeklySummarySlot(new Date("2026-10-19T06:59:00Z"), "UTC").toISOString()).toBe(
      "2026-10-12T07:00:00.000Z",
    );
    expect(weeklySummarySlot(new Date("2026-10-19T07:00:00Z"), "UTC").toISOString()).toBe(
      "2026-10-19T07:00:00.000Z",
    );
  });

  it("follows the zone's own clock across daylight saving", () => {
    // Chicago is UTC-5 in October and UTC-6 after 2026-11-01.
    expect(
      weeklySummarySlot(new Date("2026-10-26T13:00:00Z"), "America/Chicago").toISOString(),
    ).toBe("2026-10-26T12:00:00.000Z");
    expect(
      weeklySummarySlot(new Date("2026-11-02T14:00:00Z"), "America/Chicago").toISOString(),
    ).toBe("2026-11-02T13:00:00.000Z");
    // Sunday evening in Chicago is already Monday in UTC, but not yet Monday 07:00 there.
    expect(
      weeklySummarySlot(new Date("2026-10-26T03:00:00Z"), "America/Chicago").toISOString(),
    ).toBe("2026-10-19T12:00:00.000Z");
    // Tokyo is ahead of UTC.
    expect(weeklySummarySlot(new Date("2026-10-18T22:30:00Z"), "Asia/Tokyo").toISOString()).toBe(
      "2026-10-18T22:00:00.000Z",
    );
  });

  it("holds counts only", () => {
    const opened = { critical: 1, high: 2, medium: 3, low: 0, info: 0 };
    const digest = {
      signInFindings: {
        state: "ok",
        opened,
        openTotal: 4,
        resolved: 2,
        open: [{ title: "Synthetic", subjectLabel: "Avery Stone" }],
      },
      directoryChanges: { state: "ok", total: 9, flaggedTotal: 1, flagged: [] },
      quietAccounts: { state: "ok", total: 5, items: [{ name: "Avery Stone" }] },
      auditDrift: {
        compared: 2,
        audits: [{ tiles: [{ change: 3 }, { change: 0 }] }, { tiles: [{ change: -1 }] }],
      },
      microsoft: { state: "ok", posts: { total: 2 }, incidents: { opened: 1 } },
      collection: { collecting: true, failedRuns: 0, gaps: 1 },
    } as unknown as Parameters<typeof weeklySummaryAlert>[0];
    const alert = weeklySummaryAlert(digest);
    expect(alert.lines).toEqual([
      "New sign-in findings: 6, of which 1 critical and 2 high. 4 open, 2 resolved.",
      "Directory changes: 9, of which 1 flagged.",
      "Accounts that went quiet: 5.",
      "Audit numbers that moved: 2, across 2 audits compared with an earlier run.",
      "Microsoft: 2 Message center posts with a deadline, 1 service incidents opened.",
      "Collection is on. Failed collector runs: 0. Gaps: 1.",
    ]);
    expect(JSON.stringify(alert)).not.toContain("Avery Stone");
    expect(alert.items).toEqual([]);
  });
});

describe("scheduled exports", () => {
  it("attaches the CSV only to email with details on, up to 5 MB", () => {
    expect(exportWithheld({ kind: "email", includeDetails: true }, 1_000)).toBeNull();
    expect(exportWithheld({ kind: "email", includeDetails: true }, 5 * 1024 * 1024)).toBeNull();
    expect(exportWithheld({ kind: "email", includeDetails: true }, 5 * 1024 * 1024 + 1)).toBe(
      "too_large",
    );
    expect(exportWithheld({ kind: "email", includeDetails: false }, 10)).toBe("details_off");
    for (const kind of ["teams", "slack", "generic"] as const)
      expect(exportWithheld({ kind, includeDetails: true }, 10)).toBe("webhook");
  });

  it("says what the run held and where the file is", () => {
    const alert = exportAlert({
      auditId: "accounts.inactive",
      title: "Inactive accounts",
      rows: 12,
      complete: false,
      scoped: true,
      attached: false,
      withheld: "too_large",
    });
    expect(alert.title).toBe("Scheduled audit: Inactive accounts");
    expect(alert.lines).toEqual([
      "12 rows, from a partial result, limited to some account domains.",
      "The CSV is larger than 5 MB, so it is not attached. Open Prism365 for the file.",
    ]);
    expect(alert.link).toBe("/#/audits/accounts.inactive");
  });
});

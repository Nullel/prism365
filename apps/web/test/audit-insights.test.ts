import {
  type AuditCatalog,
  type AuditRun,
  type AuditTrend,
  type AuditTrendPoint,
  auditDefinitionSchema,
  auditTrendSchema,
} from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import {
  AuditCatalogList,
  AuditScheduleControl,
  AuditTrendPanel,
  AuditTrendView,
  defaultTrendTile,
  futureRunNote,
  runLabel,
  saveAuditSchedule,
  scheduleLimitNote,
  scheduleTag,
  trendGeometry,
  trendHollowNote,
  trendNote,
  trendValue,
} from "../src/Audits.tsx";

// The trend panel, the schedule control and the scheduled tags. Synthetic values only.

const point = (
  day: number,
  tiles: Record<string, number>,
  overrides: Partial<AuditTrendPoint> = {},
): AuditTrendPoint => ({
  runId: `00000000-0000-4000-8000-0000000000${String(day).padStart(2, "0")}`,
  finishedAt: new Date(Date.UTC(2026, 9, day, 10)).toISOString(),
  version: 3,
  complete: true,
  scopeLimited: false,
  trigger: "manual",
  tiles,
  notComparable: [],
  ...overrides,
});

const tiles: AuditTrend["tiles"] = [
  { id: "accounts", label: "Accounts", tone: "neutral" },
  { id: "inactive-30-days", label: "Inactive 30+ days", tone: "warning" },
];

const trend = (points: AuditTrendPoint[]): AuditTrend =>
  auditTrendSchema.parse({
    auditId: "accounts.inactive",
    days: 365,
    points,
    tiles,
    truncated: false,
  });

const view = (value: AuditTrend, options: { tileId?: string; showTable?: boolean } = {}) =>
  renderToStaticMarkup(
    createElement(AuditTrendView, {
      trend: value,
      tileId: options.tileId ?? "inactive-30-days",
      showTable: options.showTable ?? false,
      onTile: () => {},
      onToggleTable: () => {},
    }),
  );

const count = (html: string, pattern: RegExp) => html.match(pattern)?.length ?? 0;

describe("audit trend panel", () => {
  const two = trend([
    point(1, { accounts: 210, "inactive-30-days": 52 }),
    point(8, { accounts: 214, "inactive-30-days": 47 }, { trigger: "schedule" }),
  ]);

  it("draws an accessible line chart once there are two points", () => {
    const html = view(two);
    expect(html).toContain('<h3 id="');
    expect(html).toContain(">Trend</h3>");
    expect(html).toMatch(
      /<svg class="audit-trend__chart" viewBox="0 0 600 160" role="img" aria-labelledby="[^"]+ [^"]+"/,
    );
    expect(html).toContain("<title");
    expect(html).toContain("Inactive 30+ days over the last 365 days</title>");
    expect(html).toContain("<desc");
    expect(html).toContain("2 runs with this value.");
    expect(count(html, /<polyline /g)).toBe(1);
    expect(count(html, /<circle /g)).toBe(2);
    expect(html).not.toContain("audit-trend__point--hollow");
    // The latest value against the first one, in plain words.
    expect(html).toMatch(/Inactive 30\+ days: 47 on [^<]+?, against 52 on [^<]+\./);
    expect(html).toContain('aria-pressed="false"');
    expect(html).toContain("Show as table");
    expect(html).not.toContain("<table");
  });

  it("shows nothing with fewer than two points", () => {
    expect(view(trend([point(1, { accounts: 1 })]))).toBe("");
    expect(view(trend([]))).toBe("");
    expect(
      renderToStaticMarkup(
        createElement(AuditTrendPanel, { trend: trend([point(1, { accounts: 1 })]) }),
      ),
    ).toBe("");
  });

  it("draws a point that does not compare hollow, and says why", () => {
    const html = view(
      trend([
        point(1, { "inactive-30-days": 60 }, { version: 2, notComparable: ["version"] }),
        point(2, { "inactive-30-days": 58 }, { notComparable: ["scope"], scopeLimited: true }),
        point(3, { "inactive-30-days": 55 }, { complete: false, notComparable: ["incomplete"] }),
        point(4, { "inactive-30-days": 47 }),
      ]),
    );
    expect(count(html, /audit-trend__point--hollow/g)).toBe(3);
    expect(html).toContain(
      "Hollow points do not compare directly with the latest run: 1 by another version of the audit, 1 with another audit scope and 1 partial result.",
    );
    // Each point's own tooltip names its reason.
    expect(html).toContain(": 60. Another version of the audit.</title>");
    expect(html).toContain(": 55. Partial result.</title>");
  });

  it("lists the points as a table when asked, with who started each and why it differs", () => {
    const html = view(
      trend([
        point(1, { "inactive-30-days": 52 }, { complete: false, notComparable: ["incomplete"] }),
        point(2, { accounts: 3 }),
        point(8, { "inactive-30-days": 47 }, { trigger: "schedule" }),
      ]),
      { showTable: true },
    );
    expect(html).not.toContain("<svg");
    expect(html).toContain('aria-pressed="true"');
    expect(html).toContain("<caption");
    expect(count(html, /<tr>/g)).toBe(4);
    expect(html).toContain('<th scope="col" class="audit-numeric">Inactive 30+ days</th>');
    expect(html).toContain("Partial result.");
    expect(html).toContain("Schedule");
    expect(html).toContain("An administrator");
    // A run that had no such tile says so instead of showing zero.
    expect(html).toContain('<span class="audit-empty">None</span>');
  });

  it("picks the first critical or warning tile, else the first", () => {
    expect(defaultTrendTile(tiles)).toBe("inactive-30-days");
    expect(
      defaultTrendTile([
        { id: "a", label: "A", tone: "neutral" },
        { id: "b", label: "B", tone: "good" },
      ]),
    ).toBe("a");
    expect(defaultTrendTile([])).toBeNull();
    const html = renderToStaticMarkup(createElement(AuditTrendPanel, { trend: two }));
    expect(html).toContain('<option value="inactive-30-days" selected="">');
  });

  it("reads only a point's own values", () => {
    const value = point(1, { accounts: 4 });
    expect(trendValue(value, "accounts")).toBe(4);
    expect(trendValue(value, "constructor")).toBeNull();
    expect(trendValue(value, "toString")).toBeNull();
  });

  it("leaves a gap where a run had no such tile, and spaces runs of one moment evenly", () => {
    const gap = trendGeometry(
      [
        point(1, { accounts: 2 }),
        point(2, { accounts: 4 }),
        point(3, {}),
        point(4, { accounts: 1 }),
        point(5, { accounts: 3 }),
      ],
      "accounts",
    );
    expect(gap.dots).toHaveLength(4);
    expect(gap.segments).toHaveLength(2);
    expect(gap.max).toBe(4);
    const same = trendGeometry(
      [point(1, { accounts: 0 }), point(1, { accounts: 0 }), point(1, { accounts: 0 })],
      "accounts",
    );
    expect(same.dots.map((dot) => dot.x)).toEqual([10, 300, 590]);
    // All zero: the line sits on the baseline, not off the chart.
    expect(same.dots.every((dot) => dot.y === same.baseline)).toBe(true);
  });

  it("words notes in plain sentences", () => {
    expect(trendNote(point(1, {}))).toBe("");
    expect(trendNote(point(1, {}, { notComparable: ["version", "scope"] }))).toBe(
      "Another version of the audit, another audit scope.",
    );
    expect(trendHollowNote([point(1, {})])).toBeNull();
    expect(
      trendHollowNote([
        point(1, {}, { notComparable: ["incomplete"] }),
        point(2, {}, { notComparable: ["incomplete"] }),
      ]),
    ).toBe("Hollow points do not compare directly with the latest run: 2 partial results.");
  });
});

const control = (props: Partial<Parameters<typeof AuditScheduleControl>[0]> = {}) =>
  renderToStaticMarkup(
    createElement(AuditScheduleControl, {
      frequency: "off",
      collecting: true,
      saving: false,
      error: "",
      onChange: () => {},
      ...props,
    }),
  );

describe("audit schedule control", () => {
  it("offers off, daily and weekly with the current one pressed", () => {
    const html = control({ frequency: "weekly" });
    expect(html).toContain("Run automatically");
    expect(count(html, /<button /g)).toBe(3);
    expect(html).toMatch(/aria-pressed="false"[^>]*>Off</);
    expect(html).toMatch(/aria-pressed="false"[^>]*>Daily</);
    expect(html).toMatch(/class="is-active" aria-pressed="true"[^>]*>Weekly</);
    expect(html).toContain(
      "Runs while collection is on, with the collecting administrator&#x27;s sign-in.",
    );
    expect(html).not.toContain("Collection is off");
  });

  it("says nothing runs while collection is off, for a scheduled audit", () => {
    expect(control({ frequency: "daily", collecting: false })).toContain(
      "Collection is off, so this audit will not run on its own until someone starts it.",
    );
    expect(control({ frequency: "off", collecting: false })).not.toContain("Collection is off");
  });

  it("says you can schedule only what you could run yourself, and it reads no more", () => {
    expect(control()).toContain(
      "You can schedule an audit only if you could run it yourself, and it never reads more than you could.",
    );
  });

  it("says when a scheduled run reads less than a manual run would, and with whose grants", () => {
    const limit = { by: "Other admin", without: ["AuditLog.Read.All"] };
    expect(scheduleLimitNote(limit)).toBe(
      "Runs with the permissions Other admin had when scheduling, so it does not read with AuditLog.Read.All.",
    );
    expect(
      scheduleLimitNote({ by: null, without: ["AuditLog.Read.All", "LicenseAssignment.Read.All"] }),
    ).toBe(
      "Runs with its required permissions only, so it does not read with AuditLog.Read.All and LicenseAssignment.Read.All: it was scheduled before Prism365 kept the permissions of whoever scheduled it.",
    );
    expect(scheduleLimitNote(null)).toBeNull();
    expect(control({ frequency: "daily", limit })).toContain(
      "Runs with the permissions Other admin had when scheduling",
    );
    // Off, nothing runs, so nothing is said.
    expect(control({ frequency: "off", limit })).not.toContain("had when scheduling");
  });

  it("saves one audit at a time, never a copy of the whole schedule", async () => {
    const sent: { url: string; body: unknown; headers: Headers }[] = [];
    const original = globalThis.fetch;
    globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
      sent.push({
        url: String(input),
        body: JSON.parse(String(init?.body)),
        headers: new Headers(init?.headers),
      });
      return Response.json({ audits: { "accounts.inactive": "daily" }, collecting: true });
    }) as typeof fetch;
    try {
      const saved = await saveAuditSchedule(
        { csrfToken: "synthetic-csrf" } as Parameters<typeof saveAuditSchedule>[0],
        "accounts.inactive",
        "daily",
      );
      expect(saved).toEqual({
        audits: { "accounts.inactive": "daily" },
        collecting: true,
        limits: {},
      });
    } finally {
      globalThis.fetch = original;
    }
    expect(sent).toHaveLength(1);
    expect(sent[0]?.url).toBe("/api/v1/settings/audit-schedule");
    expect(sent[0]?.body).toEqual({ audits: { "accounts.inactive": "daily" } });
    expect(sent[0]?.headers.get("x-csrf-token")).toBe("synthetic-csrf");
  });

  it("says when the next run waits because the last one is dated in the future", () => {
    const now = Date.parse("2026-10-22T12:00:00.000Z");
    const future = "2026-11-01T09:00:00.000Z";
    const note = futureRunNote("daily", future, now);
    expect(note).toMatch(/^Next run waits for .+: the last run is dated in the future\.$/);
    expect(note).toContain(
      new Intl.DateTimeFormat(undefined, { dateStyle: "medium", timeStyle: "short" }).format(
        new Date("2026-11-02T09:00:00.000Z"),
      ),
    );
    expect(futureRunNote("weekly", future, now)).toContain(
      new Intl.DateTimeFormat(undefined, { dateStyle: "medium", timeStyle: "short" }).format(
        new Date("2026-11-08T09:00:00.000Z"),
      ),
    );
    // A start in the past, no start, or no schedule: nothing to say.
    expect(futureRunNote("daily", "2026-10-21T09:00:00.000Z", now)).toBeNull();
    expect(futureRunNote("daily", null, now)).toBeNull();
    expect(futureRunNote("off", future, now)).toBeNull();
    const html = control({ frequency: "daily", lastStart: "2999-01-01T00:00:00.000Z" });
    expect(html).toContain("the last run is dated in the future.");
    expect(control({ frequency: "daily", lastStart: "2026-10-21T09:00:00.000Z" })).not.toContain(
      "dated in the future",
    );
  });

  it("holds still while saving and shows a save error", () => {
    const saving = control({ saving: true });
    expect(saving).toContain("<fieldset");
    expect(saving).toMatch(/<fieldset[^>]*disabled=""/);
    expect(saving).toContain("Saving…");
    expect(control({ error: "The schedule could not be saved." })).toContain(
      'role="alert">The schedule could not be saved.',
    );
  });
});

const definition = auditDefinitionSchema.parse({
  id: "accounts.inactive",
  version: 3,
  title: "Inactive accounts",
  description: "Every account and how long it has gone without signing in.",
  category: "Accounts",
  subject: { one: "account", many: "accounts" },
  columns: [{ key: "person", label: "Account", kind: "person" }],
  views: [{ id: "all", label: "All", tags: [] }],
  defaultView: "all",
  defaultSort: { column: "person", direction: "asc" },
  permissions: ["User.Read.All"],
  optionalPermissions: [],
  requirements: [],
});

describe("scheduled audits in the catalog and history", () => {
  const catalog = {
    mode: "live",
    audits: [
      {
        definition,
        availability: { state: "available", missing: [], grantFeature: null, enable: [] },
        lastRun: null,
        runningId: null,
      },
    ],
  } as AuditCatalog;
  const render = (schedule: Record<string, "off" | "daily" | "weekly"> | null) =>
    renderToStaticMarkup(
      createElement(AuditCatalogList, {
        catalog,
        busy: null,
        onOpen: () => {},
        onRun: () => {},
        schedule,
      }),
    );

  it("labels an audit by how often it runs on its own", () => {
    expect(render({ "accounts.inactive": "daily" })).toContain(
      '<span class="person-tag">Daily</span>',
    );
    expect(render({ "accounts.inactive": "weekly" })).toContain(
      '<span class="person-tag">Weekly</span>',
    );
    expect(render({ "accounts.inactive": "off" })).toContain(
      '<span class="person-muted">Off</span>',
    );
    expect(render({ "accounts.inactive": "off" })).not.toContain("person-tag");
    // Before the schedule has loaded, nothing is claimed.
    expect(render(null)).not.toContain(">Off<");
    expect(render(null)).not.toContain("person-tag");
    expect(scheduleTag(undefined)).toBeNull();
  });

  it("marks scheduled runs in the history list", () => {
    const run = {
      id: "66666666-6666-4666-8666-666666666666",
      auditId: "accounts.inactive",
      auditVersion: 3,
      status: "succeeded",
      complete: true,
      startedAt: "2026-10-22T09:00:00.000Z",
      finishedAt: "2026-10-22T09:01:00.000Z",
      startedBy: { id: "22222222-2222-4222-8222-222222222222", name: "Synthetic admin" },
      error: null,
      summary: [],
      notes: [],
      rowCount: 0,
      scope: null,
      domains: [],
      trigger: "schedule",
    } satisfies AuditRun;
    expect(runLabel(run)).toMatch(/ \(scheduled\)$/);
    expect(runLabel({ ...run, status: "failed" })).toMatch(/ \(failed, scheduled\)$/);
    expect(runLabel({ ...run, trigger: "manual" })).not.toContain("(");
  });
});

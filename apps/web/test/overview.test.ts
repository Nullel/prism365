import { readFileSync } from "node:fs";
import {
  type AttentionItem,
  type AttentionList,
  type AuditCatalog,
  type AuditRun,
  auditDefinitionSchema,
  type Finding,
  type ModuleList,
  type PrismSession,
  type Severity,
  type SigninSummary,
} from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { normalizeKey } from "../src/demo/replay.ts";
import {
  buildTiles,
  collectionSummary,
  emptyOverviewData,
  findingCounts,
  findingsHome,
  LiveOverview,
  microsoftSummary,
  needsAttention,
  OVERVIEW_REQUESTS,
  type OverviewData,
  OverviewView,
  type Source,
  settle,
} from "../src/LiveOverview.tsx";
import { findAll, hostTree, nestedInteractive, textOf } from "./markup.ts";

// Invented data only.
const NOW = Date.parse("2026-10-24T12:00:00Z");
const ago = (minutes: number) => new Date(NOW - minutes * 60_000).toISOString();
const ahead = (days: number) => new Date(NOW + days * 86_400_000).toISOString();
const ready = <T>(value: T): Source<T> => ({ value, error: null });
const failed = <T>(error = "Request failed with HTTP 502."): Source<T> => ({ value: null, error });

let counter = 0;
const uuid = () => `00000000-0000-4000-8000-${String(++counter).padStart(12, "0")}`;

function finding(over: Partial<Finding> & { severity?: Severity } = {}): Finding {
  return {
    id: uuid(),
    module: "signins.audit",
    ruleId: "signins.unfamiliar_signin",
    ruleVersion: 1,
    purpose: "security",
    severity: "medium",
    score: 100,
    subject: { kind: "person", id: "u-1", label: "Avery Example" },
    title: "Sign-in from a new country",
    summary: "A sign-in came from a place the account has not used.",
    evidence: [],
    limitations: [],
    windowStart: ago(60),
    windowEnd: ago(1),
    state: "open",
    occurrences: 1,
    firstSeen: ago(30),
    lastSeen: ago(10),
    stateChangedAt: null,
    ...over,
  };
}

const run = (startedAt: string, status: "succeeded" | "partial" | "failed" = "succeeded") => ({
  status,
  credentialKind: "delegated" as const,
  startedAt,
  finishedAt: startedAt,
  fetched: 10,
  inserted: 4,
  findings: 0,
  skipped: 0,
  errorCategory: null,
});

function modules(
  state: "collecting" | "stopped" = "collecting",
  runs: { signins?: string | null; changes?: string | null } = {
    signins: ago(7),
    changes: ago(3),
  },
): ModuleList {
  const collector = (id: string, at: string | null | undefined) => ({
    id,
    title: id,
    cadenceSeconds: 300,
    lastRun: at ? run(at) : null,
  });
  const module = (id: string, collectors: ReturnType<typeof collector>[]) => ({
    id,
    title: id,
    description: id,
    packs: [],
    permissions: [],
    missingPermissions: [],
    collectors,
  });
  return {
    mode: "live",
    collection: {
      available: true,
      state,
      credentialKind: "delegated",
      boundAt: null,
      lastStopReason: null,
      durableStorage: true,
      reason: null,
      resumes: false,
    },
    modules: [
      module("signins.audit", [collector("signins", runs.signins)]),
      module("directory.changes", [collector("audits", runs.changes)]),
    ],
  };
}

const summary = (over: Partial<SigninSummary["totals"]> = {}, events = 1235): SigninSummary => ({
  days: 7,
  totals: {
    successes: 1204,
    failures: 31,
    blockedByPolicy: 2,
    legacySuccesses: 0,
    accounts: 50,
    ...over,
  },
  countries: [],
  failingNetworks: [],
  coverage: { events, firstEvent: ago(9000), lastEvent: ago(3), lastCollected: ago(3) },
  history: { phase: "live", from: ago(9000), through: ago(3) },
});

const definition = (id: string, title: string) =>
  auditDefinitionSchema.parse({
    id,
    version: 1,
    title,
    description: `${title}, in one sentence.`,
    category: "Accounts",
    subject: { one: "account", many: "accounts" },
    columns: [{ key: "person", label: "Account", kind: "person" }],
    views: [{ id: "all", label: "All", tags: [] }],
    defaultView: "all",
    defaultSort: { column: "person", direction: "asc" },
    permissions: [],
    optionalPermissions: [],
    requirements: [],
  });

function auditRun(
  auditId: string,
  tone: "neutral" | "good" | "warning" | "critical" | "failed" | "none",
): AuditRun | null {
  if (tone === "none") return null;
  return {
    id: uuid(),
    auditId,
    auditVersion: 1,
    status: tone === "failed" ? "failed" : "succeeded",
    complete: true,
    startedAt: ago(90),
    finishedAt: ago(89),
    startedBy: { id: uuid(), name: "Synthetic admin" },
    error: tone === "failed" ? { code: "forbidden", message: "Microsoft refused." } : null,
    summary:
      tone === "failed"
        ? []
        : [
            {
              label: "Accounts",
              value: 4,
              tone: tone === "critical" || tone === "warning" ? tone : tone,
              view: null,
            },
          ],
    notes: [],
    rowCount: 12,
    scope: null,
    domains: [],
    trigger: "manual",
  };
}

function audits(...tones: Parameters<typeof auditRun>[1][]): AuditCatalog {
  return {
    mode: "live",
    audits: tones.map((tone, index) => {
      const id = `audit.number-${index + 1}`;
      return {
        definition: definition(id, `Audit ${index + 1}`),
        availability: { state: "available", missing: [], grantFeature: null, enable: [] },
        lastRun: auditRun(id, tone),
        runningId: null,
      };
    }),
  };
}

function item(
  over: Partial<AttentionItem> & {
    deadline?: string | null;
    active?: boolean;
    status?: string | null;
    impact?: string | null;
    workaround?: string | null;
  } = {},
): AttentionItem {
  const {
    deadline = null,
    active = true,
    status = null,
    impact = null,
    workaround = null,
    ...rest
  } = over;
  return {
    id: uuid(),
    module: "service.health",
    kind: "incident",
    class: "incident",
    ref: "EX900001",
    title: "Some users cannot send mail",
    summary: "Mail is delayed.",
    services: ["Exchange Online"],
    provider: {
      status,
      labels: [],
      deadline,
      updatedAt: ago(20),
      active,
      impact,
      workaround,
    },
    localLabels: [],
    severity: "high",
    link: null,
    revision: 1,
    changeNote: null,
    firstSeen: ago(60),
    lastSeen: ago(20),
    changedAt: ago(20),
    unread: false,
    review: { state: "new", changedAt: null },
    muted: false,
    ...rest,
  };
}

const sources = (
  health: AttentionList["sources"][number]["state"] = "available",
  messages: AttentionList["sources"][number]["state"] = "available",
): AttentionList["sources"] => [
  { module: "service.health", title: "Service health", state: health, missing: [] },
  { module: "service.messages", title: "Message center", state: messages, missing: [] },
];

const attention = (items: AttentionItem[], list = sources()): AttentionList => ({
  mode: "live",
  items,
  system: [],
  sources: list,
  counts: { unread: 0, active: items.length },
});

const action = (title: string, deadline: string | null, over: Partial<AttentionItem> = {}) =>
  item({
    module: "service.messages",
    kind: "message",
    class: "message.action",
    ref: "MC900001",
    title,
    services: ["SharePoint Online"],
    deadline,
    ...over,
  });

describe("findingCounts", () => {
  it("counts the open findings of both modules by severity", () => {
    const counts = findingCounts([
      [
        finding({ severity: "critical" }),
        finding({ severity: "high" }),
        finding({ severity: "high" }),
      ],
      [
        finding({ severity: "medium" }),
        finding({ severity: "low" }),
        finding({ severity: "info" }),
      ],
    ]);
    expect(counts).toEqual({ total: 6, critical: 1, high: 2, medium: 1, low: 2, capped: false });
  });

  it("leaves out findings that are not open", () => {
    const counts = findingCounts([
      [
        finding({ severity: "critical", state: "resolved" }),
        finding({ severity: "high", state: "acknowledged" }),
      ],
      [finding({ severity: "high" })],
    ]);
    expect(counts).toMatchObject({ total: 1, critical: 0, high: 1 });
  });

  it("says when a module returned as many as were asked for, so there may be more", () => {
    const full = Array.from({ length: 100 }, () => finding());
    expect(findingCounts([full, []]).capped).toBe(true);
    expect(findingCounts([full.slice(1), []]).capped).toBe(false);
    expect(findingCounts([[finding()]], 1).capped).toBe(true);
  });

  it("counts nothing for nothing", () => {
    expect(findingCounts([[], []])).toEqual({
      total: 0,
      critical: 0,
      high: 0,
      medium: 0,
      low: 0,
      capped: false,
    });
  });
});

describe("needsAttention", () => {
  it("puts the most severe first and the newest first among equals", () => {
    const rows = needsAttention([
      [
        finding({ title: "medium old", severity: "medium", lastSeen: ago(500) }),
        finding({ title: "critical old", severity: "critical", lastSeen: ago(900) }),
      ],
      [
        finding({ title: "high new", severity: "high", lastSeen: ago(2) }),
        finding({ title: "medium new", severity: "medium", lastSeen: ago(5) }),
        finding({ title: "critical new", severity: "critical", lastSeen: ago(30) }),
      ],
    ]);
    expect(rows.map((row) => row.title)).toEqual([
      "critical new",
      "critical old",
      "high new",
      "medium new",
      "medium old",
    ]);
  });

  it("keeps eight rows, and only open findings", () => {
    const many = Array.from({ length: 12 }, (_, index) => finding({ lastSeen: ago(index + 1) }));
    expect(needsAttention([many, []])).toHaveLength(8);
    expect(
      needsAttention([[...many, finding({ state: "resolved", severity: "critical" })], []])[0]
        ?.severity,
    ).toBe("medium");
    expect(needsAttention([many, []], 3)).toHaveLength(3);
  });

  it("is stable for findings that tie on both", () => {
    const a = finding({ id: "00000000-0000-4000-8000-00000000000a", lastSeen: ago(5) });
    const b = finding({ id: "00000000-0000-4000-8000-00000000000b", lastSeen: ago(5) });
    expect(needsAttention([[b, a], []]).map((row) => row.id)).toEqual([a.id, b.id]);
  });
});

describe("findingsHome", () => {
  it("opens the module with more open findings, and the sign-in audit when level or empty", () => {
    const some = (n: number, over: Partial<Finding> = {}) =>
      Array.from({ length: n }, () => finding(over));
    expect(findingsHome(some(2), some(3))).toBe("changes");
    expect(findingsHome(some(3), some(2))).toBe("signins");
    expect(findingsHome(some(2), some(2))).toBe("signins");
    expect(findingsHome([], [])).toBe("signins");
    expect(findingsHome(some(1), some(5, { state: "resolved" }))).toBe("signins");
  });
});

describe("collectionSummary", () => {
  it("is collecting when it runs and nothing says a span could not be read", () => {
    expect(collectionSummary(modules(), [finding()])).toEqual({
      state: "collecting",
      lastRun: ago(3),
    });
  });

  it("is stopped when nobody collects, whatever the findings say", () => {
    expect(
      collectionSummary(modules("stopped"), [finding({ ruleId: "signins.collection_gap" })]).state,
    ).toBe("stopped");
  });

  it("is a gap while a collection gap finding is open, from either module", () => {
    expect(
      collectionSummary(modules(), [finding({ ruleId: "signins.collection_gap" })]).state,
    ).toBe("gap");
    expect(
      collectionSummary(modules(), [
        finding({ module: "directory.changes", ruleId: "collection.gap" }),
      ]).state,
    ).toBe("gap");
    expect(
      collectionSummary(modules(), [
        finding({ ruleId: "signins.collection_gap", state: "resolved" }),
      ]).state,
    ).toBe("collecting");
    expect(
      collectionSummary(modules(), [finding({ ruleId: "role.privileged-assigned" })]).state,
    ).toBe("collecting");
  });

  it("reports the newest run of any collector, and none when nothing has run", () => {
    expect(
      collectionSummary(modules("collecting", { signins: ago(2), changes: ago(40) }), []).lastRun,
    ).toBe(ago(2));
    expect(
      collectionSummary(modules("collecting", { signins: null, changes: ago(40) }), []).lastRun,
    ).toBe(ago(40));
    expect(
      collectionSummary(modules("collecting", { signins: null, changes: null }), []).lastRun,
    ).toBeNull();
  });
});

describe("microsoftSummary", () => {
  it("separates active incidents from Message center posts that need action", () => {
    const result = microsoftSummary(
      attention([
        item({ title: "older", provider: { ...item().provider, updatedAt: ago(300) } }),
        item({ title: "newer" }),
        item({ title: "closed", active: false }),
        action("Later", ahead(20)),
        action("Sooner", ahead(3)),
        action("No date", null),
        item({ class: "message.retirement", kind: "message", title: "Retirement" }),
        item({ class: "advisory", title: "Advisory" }),
      ]),
    );
    expect(result.incidents.map((row) => row.title)).toEqual(["newer", "older"]);
    expect(result.actions.map((row) => row.title)).toEqual(["Sooner", "Later", "No date"]);
    expect(result.checked).toBe(true);
    expect(result.problem).toBeNull();
  });

  it("does not take no items for good news when neither source is being read", () => {
    const result = microsoftSummary(attention([], sources("not_enabled", "missing_permission")));
    expect(result.checked).toBe(false);
    expect(result.problem).toBe("Service health is not set up. Message center needs more access.");
  });

  it("still counts what the working source reports and says what the other lacks", () => {
    const result = microsoftSummary(attention([item()], sources("available", "not_collecting")));
    expect(result.checked).toBe(true);
    expect(result.incidents).toHaveLength(1);
    expect(result.problem).toBe("Message center is not collecting.");
    expect([result.incidentsChecked, result.postsChecked]).toEqual([true, false]);
  });

  it("takes a source the list does not mention to be working", () => {
    expect(microsoftSummary(attention([], [])).checked).toBe(true);
  });
});

describe("settle", () => {
  it("takes a new value and clears the error", () => {
    expect(settle({ value: 1, error: "old" }, { status: "fulfilled", value: 2 })).toEqual({
      value: 2,
      error: null,
    });
  });

  it("keeps the last good value with the reason when a refresh fails", () => {
    expect(
      settle({ value: 1, error: null }, { status: "rejected", reason: new Error("Gone.") }),
    ).toEqual({
      value: 1,
      error: "Gone.",
    });
    expect(settle(ready(1), { status: "rejected", reason: "odd" }).error).toBe(
      "The request failed.",
    );
    expect(
      settle<number>(emptyOverviewData.modules as never, {
        status: "rejected",
        reason: new Error("x"),
      }).value,
    ).toBeNull();
  });
});

const everything = (over: Partial<OverviewData> = {}): OverviewData => ({
  modules: ready(modules()),
  summary: ready(summary()),
  signins: ready([
    finding({ severity: "critical" }),
    finding({ severity: "high" }),
    finding({ severity: "medium" }),
    finding({ severity: "low" }),
  ]),
  changes: ready([
    finding({
      module: "directory.changes",
      ruleId: "role.privileged-assigned",
      severity: "high",
      subject: { kind: "tenant", id: "t", label: "Directory" },
    }),
  ]),
  audits: ready(audits("warning", "critical", "good", "none", "failed")),
  attention: ready(
    attention([item({ status: "Investigating" }), action("Decide on sharing links", ahead(9))]),
  ),
  ...over,
});

const tileOf = (data: OverviewData, id: string) => {
  const found = buildTiles(data, NOW).find((tile) => tile.id === id);
  if (!found) throw new Error(`No tile ${id}`);
  return found;
};

describe("summary tiles", () => {
  it("are five, in order, each opening the page that owns the figure", () => {
    const tiles = buildTiles(everything(), NOW);
    expect(tiles.map((tile) => [tile.id, tile.view])).toEqual([
      ["findings", "signins"],
      ["audits", "audits"],
      ["signins", "signins"],
      ["collection", "signins"],
      ["microsoft", "notifications"],
    ]);
    const moreChanges = everything({
      changes: ready([
        finding({ module: "directory.changes" }),
        finding({ module: "directory.changes" }),
      ]),
      signins: ready([finding()]),
    });
    expect(tileOf(moreChanges, "findings").view).toBe("changes");
  });

  it("counts open findings across both modules", () => {
    const tile = tileOf(everything(), "findings");
    expect(tile.status).toBe("ready");
    expect(tile.value).toBe("5");
    const html = renderToStaticMarkup(createElement("div", null, tile.detail));
    expect(html).toContain(">1 critical<");
    expect(html).toContain(">2 high<");
    expect(html).toContain(">1 medium<");
    expect(html).not.toContain("low");
  });

  it("marks a count that may be more than shown, and says so when nothing is open", () => {
    const full = Array.from({ length: 100 }, () => finding());
    expect(tileOf(everything({ signins: ready(full) }), "findings").value).toBe("101+");
    const none = tileOf(everything({ signins: ready([]), changes: ready([]) }), "findings");
    expect(none.value).toBe("0");
    expect(none.detail).toBe("Nothing open");
    expect(
      tileOf(
        everything({ signins: ready([finding({ severity: "low" })]), changes: ready([]) }),
        "findings",
      ).detail,
    ).toBe("Low severity only");
  });

  it("counts critical audit results, a failed run among them, with the warnings beside them", () => {
    const tile = tileOf(everything(), "audits");
    expect(tile.label).toBe("Critical audit results");
    expect(tile.view).toBe("audits");
    // One critical and one failed, one warning, one clear; the audit that never ran is not counted.
    expect(tile.value).toBe("2");
    expect(tile.detail).toBe("1 more with warnings, of 4 run");
    // Without warnings the detail says only out of how many.
    const none = tileOf(everything({ audits: ready(audits("good", "critical")) }), "audits");
    expect([none.value, none.detail]).toEqual(["1", "of 2 run"]);
    const clear = tileOf(everything({ audits: ready(audits("good", "neutral")) }), "audits");
    expect([clear.value, clear.detail]).toEqual(["0", "of 2 run"]);
    expect(tileOf(everything({ audits: ready(audits("none", "none")) }), "audits").detail).toBe(
      "No audit has run yet",
    );
  });

  it("shows successful sign-ins with the failed ones beside them", () => {
    const tile = tileOf(everything(), "signins");
    expect(tile.value).toBe("1,204");
    const html = renderToStaticMarkup(createElement("div", null, tile.detail));
    expect(html).toContain("overview-tile__alert");
    expect(html).toContain("31 failed");
    const clean = renderToStaticMarkup(
      createElement(
        "div",
        null,
        tileOf(everything({ summary: ready(summary({ failures: 0 })) }), "signins").detail,
      ),
    );
    expect(clean).not.toContain("overview-tile__alert");
    expect(
      tileOf(everything({ summary: ready(summary({ successes: 0, failures: 0 }, 0)) }), "signins")
        .detail,
    ).toBe("Nothing collected yet");
  });

  it("names the collection state with the last run as a relative time", () => {
    const collecting = tileOf(everything(), "collection");
    expect([collecting.value, collecting.dot]).toEqual(["Collecting", "good"]);
    expect(renderToStaticMarkup(createElement("div", null, collecting.detail))).toContain(
      "3 min ago",
    );
    const stopped = tileOf(everything({ modules: ready(modules("stopped")) }), "collection");
    expect([stopped.value, stopped.dot]).toEqual(["Stopped", "warning"]);
    const gap = tileOf(
      everything({ signins: ready([finding({ ruleId: "signins.collection_gap" })]) }),
      "collection",
    );
    expect([gap.value, gap.dot]).toEqual(["Gap", "warning"]);
    expect(
      tileOf(
        everything({ modules: ready(modules("collecting", { signins: null, changes: null })) }),
        "collection",
      ).detail,
    ).toBe("No run yet");
  });

  it("adds Microsoft's incidents to the posts that need action, and says when it is not looking", () => {
    const tile = tileOf(everything(), "microsoft");
    expect(tile.value).toBe("2");
    expect(tile.detail).toBe("1 incident, 1 post needs action");
    const none = tileOf(everything({ attention: ready(attention([])) }), "microsoft");
    expect([none.value, none.detail]).toEqual(["0", "0 incidents, 0 posts need action"]);
    const blind = tileOf(
      everything({ attention: ready(attention([], sources("not_enabled", "not_enabled"))) }),
      "microsoft",
    );
    expect(blind.value).toBe("Not checked");
    expect(blind.detail).toBe("Service health is not set up. Message center is not set up.");
    // One source unread: its count of zero is not shown as a zero.
    const noHealth = tileOf(
      everything({ attention: ready(attention([], sources("missing_permission", "available"))) }),
      "microsoft",
    );
    expect(noHealth.detail).toBe("incidents not checked, 0 posts need action");
    const noPosts = tileOf(
      everything({ attention: ready(attention([], sources("available", "not_enabled"))) }),
      "microsoft",
    );
    expect(noPosts.detail).toBe("0 incidents, Message center not checked");
  });

  it("wait while a source loads, and are unavailable when one failed", () => {
    expect(buildTiles(emptyOverviewData, NOW).map((tile) => tile.status)).toEqual(
      Array(5).fill("loading"),
    );
    const broken = buildTiles(
      everything({
        audits: failed(),
        signins: failed(),
        summary: failed(),
        modules: failed(),
        attention: failed(),
      }),
      NOW,
    );
    expect(broken.map((tile) => tile.status)).toEqual([
      "unavailable",
      "unavailable",
      "unavailable",
      "unavailable",
      "unavailable",
    ]);
    // One failed findings request makes the count unavailable rather than a total that is too low.
    expect(tileOf(everything({ changes: failed() }), "findings").status).toBe("unavailable");
    // A request still on its way holds the collection tile back, so it never flashes the wrong state.
    expect(tileOf(everything({ signins: emptyOverviewData.signins }), "collection").status).toBe(
      "loading",
    );
  });
});

const session: PrismSession = {
  mode: "live",
  authenticated: true,
  displayName: "Synthetic admin",
  csrfToken: "t",
} as PrismSession;

const render = (data: OverviewData) =>
  renderToStaticMarkup(
    createElement(OverviewView, { data, now: NOW, onNavigate: () => {}, onOpenPerson: () => {} }),
  );

describe("the Overview page", () => {
  it("draws still skeleton rows while everything loads, and never says Loading", () => {
    const html = render(emptyOverviewData);
    expect(html).toContain("skeleton__row");
    expect(html).not.toContain("Loading…");
    expect(html.match(/<button/g)).toHaveLength(5);
    expect(html).toContain("Needs attention");
    expect(html).toContain("From Microsoft");
  });

  it("starts as the loading page before any data arrives", () => {
    const html = renderToStaticMarkup(
      createElement(LiveOverview, { session, onNavigate: () => {}, onOpenPerson: () => {} }),
    );
    expect(html).toContain("skeleton__row");
    expect(html).not.toContain("Loading…");
  });

  it("says plainly that nothing is open and nothing is wrong when that is so", () => {
    const html = render(
      everything({
        signins: ready([]),
        changes: ready([]),
        attention: ready(attention([])),
        audits: ready(audits("good")),
      }),
    );
    expect(html).toContain("No open findings. New ones appear here as collection runs.");
    expect(html).toContain(
      "Microsoft reports no incidents, and no Message center post needs action.",
    );
    expect(html).not.toContain("View all");
  });

  it("says what could not be loaded, tile by tile and list by list", () => {
    const html = render(
      everything({
        signins: failed(),
        changes: failed(),
        attention: failed(),
        audits: failed(),
        summary: failed(),
        modules: failed(),
      }),
    );
    expect(html.match(/>Unavailable</g)).toHaveLength(5);
    expect(html).toContain("Findings could not be loaded.");
    expect(html).toContain("Microsoft updates could not be loaded.");
    expect(html).not.toContain("skeleton");
  });

  it("shows what loaded when one findings request failed, and names the one that did not", () => {
    const html = render(everything({ changes: failed() }));
    expect(html).toContain("Sign-in from a new country");
    expect(html).toContain("Directory change findings could not be loaded.");
  });

  it("says when it is showing old figures because a refresh failed, and only then", () => {
    const stale = render(
      everything({ audits: { value: audits("warning"), error: "Request failed with HTTP 502." } }),
    );
    expect(stale).toContain("Showing the last figures that loaded. Could not refresh audits.");
    expect(render(everything())).not.toContain("Showing the last figures");
  });

  it("lists up to eight findings, most severe first, each with severity, title, subject and time", () => {
    const many = Array.from({ length: 11 }, (_, index) =>
      finding({ title: `Finding ${index + 1}`, severity: "medium", lastSeen: ago(index + 1) }),
    );
    const html = render(
      everything({
        signins: ready([
          ...many,
          finding({ title: "The critical one", severity: "critical", lastSeen: ago(600) }),
        ]),
        changes: ready([]),
      }),
    );
    expect(html.match(/class="overview-row"/g)).toHaveLength(8);
    expect(html.indexOf("The critical one")).toBeLessThan(html.indexOf("Finding 1<"));
    expect(html).not.toContain("Finding 8<");
    expect(html).toContain("sev--critical");
    expect(html).toContain("Avery Example");
    expect(html).toContain("1 min ago");
    expect(html).toContain("View all");
  });

  it("draws a finding as two lines: severity and the whole title, then subject and time", () => {
    const title = "Sign-in that does not match this account's history, seen from a new country";
    const html = render(
      everything({
        signins: ready([
          finding({ title, subject: { kind: "person", id: "u-4", label: "Dagny Holloway" } }),
        ]),
        changes: ready([]),
      }),
    );
    const row = html.match(/<button type="button" class="overview-row">(.*?)<\/button>/)?.[1] ?? "";
    // Line one holds the severity and the title, untouched by any cutting in code.
    expect(row).toMatch(
      /^<span class="overview-row__head" title="[^"]+"><span class="sev sev--medium">.*?<\/span> <span class="overview-row__title">Sign-in that does not match this account&#x27;s history, seen from a new country<\/span><\/span>/,
    );
    // Line two holds the subject, then when, and nothing else.
    expect(row).toMatch(
      /<span class="overview-row__meta"><span class="overview-row__subject">Dagny Holloway<\/span><span class="overview-row__when"><time [^>]*>10 min ago<\/time><\/span><\/span>$/,
    );
    expect(row.indexOf("overview-row__head")).toBeLessThan(row.indexOf("overview-row__meta"));
  });

  it("draws a Microsoft item as two lines: the whole title, then its service and when it is due", () => {
    const html = render(
      everything({
        attention: ready(
          attention([
            action(
              "Action required: decide how long anonymous sharing links should last",
              ahead(9),
              {
                services: ["SharePoint Online", "OneDrive for Business"],
              },
            ),
          ]),
        ),
      }),
    );
    const row =
      html.match(
        /<button type="button" class="overview-row overview-row--microsoft">(.*?)<\/button>/,
      )?.[1] ?? "";
    expect(row).toMatch(
      /^<span class="overview-row__head" title="[^"]+"><span class="overview-row__title">Action required: decide how long anonymous sharing links should last<\/span><\/span>/,
    );
    expect(row).toMatch(
      /<span class="overview-row__meta"><span class="overview-row__subject">SharePoint Online, OneDrive for Business<\/span><span class="overview-row__due">Due [^<]+<\/span><\/span>$/,
    );
  });

  it("lists incidents and the posts that need action with a due date, and nothing long", () => {
    const html = render(
      everything({
        attention: ready(
          attention([
            item({
              title: "Mail delayed",
              status: "Investigating",
              impact: "A very long account of what Microsoft says about the impact.",
              workaround: "A long workaround.",
            }),
            action("Decide on sharing links", ahead(9)),
            action("Overdue decision", ago(60 * 24 * 2)),
          ]),
        ),
      }),
    );
    expect(html).toContain("Exchange Online");
    expect(html).toContain("Mail delayed");
    expect(html).toContain("Investigating");
    expect(html).toContain("SharePoint Online");
    expect(html).toMatch(/Due [A-Z][a-z]{2} \d{1,2}, 2026/);
    expect(html).toContain("overview-row__due--late");
    // The long Microsoft text and the review buttons now live on the Microsoft updates page.
    expect(html).not.toContain("very long account");
    expect(html).not.toContain("long workaround");
    for (const gone of [
      "Mark reviewed",
      "Mark done",
      "Mute",
      "Acknowledge",
      "Is Microsoft down?",
      "Admin attention",
      "Recently signed in",
    ])
      expect(html).not.toContain(gone);
  });

  it("explains a Microsoft source that is not being read, with a way to fix it", () => {
    const html = render(
      everything({ attention: ready(attention([], sources("available", "not_enabled"))) }),
    );
    expect(html).toContain("Message center is not set up.");
    expect(html).toContain("Microsoft access");
    expect(html).toContain("Microsoft reports no incidents.");
    expect(html).not.toContain("no Message center post needs action");
    const noHealth = render(
      everything({ attention: ready(attention([], sources("missing_permission", "available"))) }),
    );
    expect(noHealth).toContain("No Message center post needs action.");
    expect(noHealth).not.toContain("Microsoft reports no incidents");
  });

  it("has no link inside a button or button inside a link", () => {
    expect(nestedInteractive(render(everything()))).toBe(0);
    expect(nestedInteractive(render(emptyOverviewData))).toBe(0);
    expect(
      nestedInteractive(
        render(
          everything({ attention: ready(attention([], sources("not_enabled", "not_enabled"))) }),
        ),
      ),
    ).toBe(0);
  });
});

describe("what a click opens", () => {
  const clicks = (data: OverviewData) => {
    const visited: string[] = [];
    const people: string[] = [];
    const tree = hostTree(
      createElement(OverviewView, {
        data,
        now: NOW,
        onNavigate: (view: string) => visited.push(view),
        onOpenPerson: (id: string) => people.push(id),
      }),
    );
    const press = (host: { props: Record<string, unknown> }) =>
      (host.props.onClick as () => void)();
    return {
      tree,
      visited,
      people,
      press,
      buttons: findAll(tree, (host) => host.tag === "button"),
    };
  };

  it("sends each tile to the page that owns it", () => {
    const { buttons, press, visited } = clicks(everything());
    const tiles = buttons.filter((button) => button.props.className === "overview-tile");
    expect(tiles).toHaveLength(5);
    for (const tile of tiles) press(tile);
    expect(visited).toEqual(["signins", "audits", "signins", "signins", "notifications"]);
  });

  it("opens a person's page from a finding about them, and the owning page for any other", () => {
    const data = everything({
      signins: ready([
        finding({
          severity: "critical",
          subject: { kind: "person", id: "u-9", label: "Blake Example" },
        }),
      ]),
      changes: ready([
        finding({
          module: "directory.changes",
          severity: "high",
          subject: { kind: "tenant", id: "t", label: "Directory" },
        }),
      ]),
    });
    const { buttons, press, visited, people } = clicks(data);
    const rows = buttons.filter((button) => button.props.className === "overview-row");
    expect(rows).toHaveLength(2);
    press(rows[0] as never);
    press(rows[1] as never);
    expect(people).toEqual(["u-9"]);
    expect(visited).toEqual(["changes"]);
  });

  it("sends View all and each Microsoft row to the right page", () => {
    const { buttons, press, visited } = clicks(everything());
    const viewAll = buttons.filter((button) => textOf(button) === "View all");
    expect(viewAll).toHaveLength(2);
    for (const button of viewAll) press(button);
    for (const row of buttons.filter((button) =>
      String(button.props.className).includes("overview-row--microsoft"),
    ))
      press(row);
    expect(visited).toEqual(["signins", "notifications", "notifications", "notifications"]);
  });
});

describe("the public demo", () => {
  // The demo has no server and answers only requests it recorded (apps/web/demo/fixtures.json), so
  // a request written in a form that was not recorded would show the Overview as unavailable.
  const recorded = Object.keys(
    (
      JSON.parse(readFileSync(new URL("../demo/fixtures.json", import.meta.url), "utf8")) as {
        responses: Record<string, unknown>;
      }
    ).responses,
  );

  it("was recorded for every request the Overview makes", () => {
    const requests = [
      ...Object.values(OVERVIEW_REQUESTS),
      "/api/v1/attention?filter=open&limit=200",
    ];
    for (const url of requests) expect(recorded, url).toContain(normalizeKey("GET", url));
  });

  // A title wrapped over two lines in the narrowest two-column card the Overview has: a 1024 px
  // window less the 236 px sidebar, the page's 64 px of padding, the 16 px between the columns and
  // the card's 32 px of padding. These are estimates, because the system font is not available to a
  // test: 7.2 px a character is wider than 13 px semibold text runs (about 6.7), so a title that
  // fits here fits there.
  const CARD_PX = (1024 - 236 - 64 - 16) / 2 - 32;
  const CHAR_PX = 7.2;
  // Room taken by the widest severity: dot, the word Critical, and the space after it.
  const SEVERITY_PX = 64;

  function linesNeeded(text: string, firstPx: number, px: number): number {
    const room = (line: number) => Math.floor((line === 0 ? firstPx : px) / CHAR_PX);
    let line = 0;
    let current = "";
    for (const word of text.split(" ")) {
      const next = current ? `${current} ${word}` : word;
      if (next.length <= room(line)) current = next;
      else {
        line += 1;
        current = word;
      }
    }
    return line + 1;
  }

  it("fits every title in the demo data on two lines at 1024 px, so none is cut short", () => {
    const responses = (
      JSON.parse(readFileSync(new URL("../demo/fixtures.json", import.meta.url), "utf8")) as {
        responses: Record<string, { body: { items: { title: string; class?: string }[] } }>;
      }
    ).responses;
    const findings = [OVERVIEW_REQUESTS.signins, OVERVIEW_REQUESTS.changes].flatMap(
      (url) => responses[normalizeKey("GET", url)]?.body.items ?? [],
    );
    const microsoft = (
      responses[normalizeKey("GET", "/api/v1/attention?filter=open&limit=200")]?.body.items ?? []
    ).filter((item) => item.class === "incident" || item.class === "message.action");
    expect(findings.length).toBeGreaterThan(4);
    expect(microsoft.length).toBeGreaterThan(1);
    for (const item of findings)
      expect(
        linesNeeded(item.title, CARD_PX - SEVERITY_PX, CARD_PX),
        item.title,
      ).toBeLessThanOrEqual(2);
    for (const item of microsoft)
      expect(linesNeeded(item.title, CARD_PX, CARD_PX), item.title).toBeLessThanOrEqual(2);
    // The estimate does bite: a forty-word title does not pass it.
    expect(linesNeeded("word ".repeat(40).trim(), CARD_PX - SEVERITY_PX, CARD_PX)).toBeGreaterThan(
      2,
    );
  });
});

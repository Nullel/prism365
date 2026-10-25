import {
  type Digest,
  type DigestFinding,
  digestSchema,
  type PrismSession,
  sessionSchema,
} from "@prism365/contracts";
import { createElement, isValidElement, type ReactElement, type ReactNode } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";
import { createDemoState, handleRequest } from "../src/demo/replay.ts";
import {
  builtFrom,
  coverageText,
  differenceText,
  digestFor,
  digestUrl,
  findingTarget,
  ThisWeek,
  WeekView,
  type WeekViewProps,
  signed,
} from "../src/Week.tsx";

import { at, expectSkeleton, relativeCell, severityDot } from "./markup.ts";

// Invented people, findings, audits and Microsoft posts only.

function session(overrides: Partial<PrismSession> = {}): PrismSession {
  return sessionSchema.parse({
    mode: "live",
    configured: true,
    authenticated: true,
    displayName: "Synthetic admin",
    csrfToken: "synthetic-csrf-token-value-000000000000",
    enabledFeatures: ["audit.signins.read"],
    missingConfiguration: [],
    ...overrides,
  });
}

const personId = "44444444-4444-4444-8444-444444444444";
const uuid = (n: number) =>
  `${String(n).repeat(8)}-${String(n).repeat(4)}-4${String(n).repeat(3)}-8${String(n).repeat(3)}-${String(n).repeat(12)}`;

function finding(n: number, overrides: Partial<DigestFinding> = {}): DigestFinding {
  return {
    id: uuid(n),
    title: `Finding ${n}`,
    severity: "high",
    subjectLabel: "Rowan Ellis",
    personId,
    firstSeen: "2026-10-21T09:00:00.000Z",
    ...overrides,
  };
}

const run = (n: number, at: string) => ({ runId: uuid(n), at });

const digest: Digest = {
  days: 7,
  from: "2026-10-15T12:00:00.000Z",
  generatedAt: "2026-10-22T12:00:00.000Z",
  signInFindings: {
    state: "ok",
    note: "3 sign-in findings opened this week. 12 still open, 2 resolved this week.",
    opened: { critical: 1, high: 2, medium: 0, low: 0, info: 0 },
    open: [
      finding(1, { title: "An unfamiliar sign-in", severity: "critical" }),
      finding(2, {
        title: "Many failed sign-ins from one address",
        subjectLabel: "198.51.100.7",
        personId: null,
      }),
    ],
    openTotal: 12,
    resolved: 2,
  },
  directoryChanges: {
    state: "ok",
    note: "14 changes this week, 3 flagged.",
    total: 14,
    categories: [
      { category: "UserManagement", count: 9 },
      { category: "RoleManagement", count: 4 },
      { category: null, count: 1 },
    ],
    flagged: [
      finding(3, { title: "Global Administrator assigned", severity: "critical" }),
      finding(4, {
        title: "Conditional Access policy changed",
        severity: "medium",
        subjectLabel: "Require MFA",
        personId: null,
      }),
    ],
    flaggedTotal: 3,
    coverage: {
      from: "2026-09-22T12:00:00.000Z",
      through: "2026-10-22T11:50:00.000Z",
      lastCollected: "2026-10-22T11:50:00.000Z",
      phase: "live",
    },
  },
  quietAccounts: {
    state: "ok",
    note: "52 accounts passed 30 days without a sign-in since the earlier result.",
    auditId: "accounts.inactive",
    view: "d30",
    total: 52,
    items: [
      { personId, name: "Rowan Ellis", principal: "rowan@example.invalid", days: 33 },
      { personId: null, name: null, principal: "shared@example.invalid", days: null },
    ],
    latest: run(5, "2026-10-22T10:00:00.000Z"),
    earlier: run(6, "2026-10-14T10:00:00.000Z"),
    differences: [],
  },
  auditDrift: {
    state: "ok",
    note: "2 audits changed this week. 1 could not be compared.",
    compared: 3,
    uncompared: 1,
    audits: [
      {
        auditId: "mfa.registration",
        title: "MFA registration",
        from: run(7, "2026-10-14T10:00:00.000Z"),
        to: run(8, "2026-10-22T10:00:00.000Z"),
        differences: [],
        tiles: [
          {
            id: "admins-without-mfa",
            label: "Admins without MFA",
            tone: "critical",
            from: 0,
            to: 2,
            change: 2,
          },
        ],
      },
      {
        auditId: "accounts.inactive",
        title: "Inactive accounts",
        from: run(9, "2026-10-14T10:00:00.000Z"),
        to: run(1, "2026-10-22T10:00:00.000Z"),
        differences: ["scope", "incomplete"],
        tiles: [
          {
            id: "inactive-30-days",
            label: "Inactive 30+ days",
            tone: "warning",
            from: 120,
            to: 115,
            change: -5,
          },
        ],
      },
    ],
  },
  microsoft: {
    state: "ok",
    note: "",
    posts: {
      state: "ok",
      note: "22 Message center posts have an act-by date in the next 30 days.",
      total: 22,
      items: [
        {
          id: uuid(2),
          title: "Retirement of a legacy sign-in method",
          actBy: "2026-10-30T00:00:00.000Z",
          services: ["Exchange Online", "Microsoft Teams"],
        },
      ],
    },
    incidents: {
      state: "ok",
      note: "2 opened and 1 resolved this week, 1 open now.",
      opened: 2,
      resolved: 1,
      openTotal: 1,
      open: [
        {
          id: uuid(3),
          title: "Some users cannot send mail",
          severity: "high",
          status: "Investigating",
          services: ["Exchange Online"],
          firstSeen: "2026-10-22T08:00:00.000Z",
        },
      ],
    },
  },
  collection: {
    state: "ok",
    note: "Collecting. 2 collector runs failed this week.",
    collecting: true,
    collectors: [
      {
        module: "signins.audit",
        moduleTitle: "Sign-in audit",
        collector: "signins",
        collectorTitle: "Sign-ins",
        lastSuccess: "2026-10-22T11:55:00.000Z",
        lastRun: { status: "succeeded", at: "2026-10-22T11:55:00.000Z" },
      },
      {
        module: "directory.changes",
        moduleTitle: "Directory changes",
        collector: "audits",
        collectorTitle: "Directory audit log",
        lastSuccess: null,
        lastRun: { status: "failed", at: "2026-10-22T11:50:00.000Z" },
      },
    ],
    failedRuns: 2,
    errors: ["throttled", "access_denied"],
    gaps: 1,
    audits: [
      {
        auditId: "accounts.inactive",
        title: "Inactive accounts",
        frequency: "weekly",
        lastRun: { status: "succeeded", at: "2026-10-22T10:00:00.000Z" },
      },
      { auditId: "mfa.registration", title: "MFA registration", frequency: "daily", lastRun: null },
    ],
  },
};

const emptyDigest: Digest = {
  ...digest,
  signInFindings: {
    state: "empty",
    note: "No new sign-in findings this week.",
    opened: { critical: 0, high: 0, medium: 0, low: 0, info: 0 },
    open: [],
    openTotal: 0,
    resolved: 0,
  },
  directoryChanges: {
    ...digest.directoryChanges,
    state: "empty",
    note: "No directory changes this week.",
    total: 0,
    categories: [],
    flagged: [],
    flaggedTotal: 0,
  },
  quietAccounts: {
    ...digest.quietAccounts,
    state: "empty",
    note: "No accounts passed 30 days without a sign-in since the earlier result.",
    total: 0,
    items: [],
  },
  auditDrift: {
    state: "empty",
    note: "No audit tile changed this week.",
    compared: 3,
    uncompared: 0,
    audits: [],
  },
  microsoft: {
    state: "empty",
    note: "",
    posts: {
      state: "empty",
      note: "No Message center posts have an act-by date in the next 30 days.",
      total: 0,
      items: [],
    },
    incidents: {
      state: "empty",
      note: "Microsoft reported no incidents this week.",
      opened: 0,
      resolved: 0,
      openTotal: 0,
      open: [],
    },
  },
  collection: {
    ...digest.collection,
    failedRuns: 0,
    errors: [],
    gaps: 0,
    audits: [],
    note: "Collecting.",
  },
};

const reason = "Not available here.";
const unavailableDigest: Digest = {
  ...emptyDigest,
  signInFindings: {
    ...emptyDigest.signInFindings,
    state: "unavailable",
    note: "The sign-in audit is not available in this installation.",
  },
  directoryChanges: {
    ...emptyDigest.directoryChanges,
    state: "unavailable",
    note: "The directory change feed is not available in this installation.",
  },
  quietAccounts: {
    ...emptyDigest.quietAccounts,
    state: "unavailable",
    note: "Inactive accounts has no kept result from before this range to compare with. Schedule it daily or weekly.",
    latest: null,
    earlier: null,
  },
  auditDrift: {
    ...emptyDigest.auditDrift,
    state: "unavailable",
    note: "No audit has a kept result yet, so there is nothing to compare.",
  },
  microsoft: {
    state: "unavailable",
    note: "",
    posts: {
      ...emptyDigest.microsoft.posts,
      state: "unavailable",
      note: "Message center is not available in this installation.",
    },
    incidents: {
      ...emptyDigest.microsoft.incidents,
      state: "unavailable",
      note: "Service health is not available in this installation.",
    },
  },
  collection: { ...emptyDigest.collection, state: "unavailable", note: reason, collectors: [] },
};

interface Call {
  kind: "navigate" | "person" | "audit" | "days";
  value: string | number;
}

function props(overrides: Partial<WeekViewProps> = {}): { props: WeekViewProps; calls: Call[] } {
  const calls: Call[] = [];
  return {
    calls,
    props: {
      digest,
      days: 7,
      error: "",
      onDays: (value) => calls.push({ kind: "days", value }),
      onNavigate: (value) => calls.push({ kind: "navigate", value }),
      onOpenPerson: (value) => calls.push({ kind: "person", value }),
      onOpenAudit: (value) => calls.push({ kind: "audit", value }),
      ...overrides,
    },
  };
}

const render = (overrides: Partial<WeekViewProps> = {}) =>
  renderToStaticMarkup(createElement(WeekView, props(overrides).props));

// A piece of text that holds elements, drawn the way the page draws it.
const markup = (node: ReactNode) => renderToStaticMarkup(createElement("p", null, node));

// What a click on each button of the page does, without a browser: the page is a tree of plain
// function components, so it is expanded here and its buttons are collected as they are.
interface Clickable {
  label: string;
  click: () => void;
}
function textOf(node: ReactNode): string {
  if (typeof node === "string" || typeof node === "number") return String(node);
  if (Array.isArray(node)) return node.map(textOf).join("");
  if (isValidElement(node))
    return textOf((node as ReactElement<{ children?: ReactNode }>).props.children);
  return "";
}
function clickables(node: ReactNode, found: Clickable[] = []): Clickable[] {
  if (Array.isArray(node)) {
    for (const child of node) clickables(child, found);
    return found;
  }
  if (!isValidElement(node)) return found;
  const element = node as ReactElement<{ children?: ReactNode; onClick?: () => void }>;
  if (typeof element.type === "function") {
    try {
      return clickables((element.type as (value: unknown) => ReactNode)(element.props), found);
    } catch {
      // A component with state (the copy button) is not part of what a link does.
      return found;
    }
  }
  if (element.type === "button" && typeof element.props.onClick === "function")
    found.push({ label: textOf(element.props.children), click: element.props.onClick });
  return clickables(element.props.children, found);
}
function press(label: string, overrides: Partial<WeekViewProps> = {}): Call[] {
  const { props: value, calls } = props(overrides);
  const matching = clickables(createElement(WeekView, value)).filter(
    (item) => item.label === label,
  );
  expect(matching.length, `buttons named "${label}"`).toBeGreaterThan(0);
  matching[0]?.click();
  return calls;
}

describe("the This week page", () => {
  it("opens with one line on where it comes from, the ranges, and the panels in order", () => {
    const html = render();
    expect(builtFrom).toBe(
      "Built from what Prism365 collected. Nothing here is read from Microsoft when you open the page.",
    );
    expect(html).toContain(builtFrom);
    for (const label of ["Last day", "Last 7 days", "Last 14 days", "Last 30 days"])
      expect(html).toContain(label);
    expect(html).toMatch(/aria-pressed="true"[^>]*>Last 7 days</);
    expect(html).toMatch(/aria-pressed="false"[^>]*>Last 30 days</);
    const headings = [
      "Sign-in findings",
      "Directory changes",
      "Accounts that went quiet",
      "Audit changes",
      "Microsoft",
      "Collection",
    ].map((title) => html.indexOf(`>${title}</h3>`));
    expect(headings.every((position) => position > 0)).toBe(true);
    expect([...headings].sort((a, b) => a - b)).toEqual(headings);
    expect(html.indexOf(builtFrom)).toBeLessThan(headings[0] ?? 0);
  });

  it("shows every section with its data", () => {
    const html = render();
    // Sign-in findings.
    expect(html).toContain("3 sign-in findings opened this week");
    expect(html).toContain("1 critical");
    expect(html).toContain("2 high");
    expect(html).not.toContain("0 medium");
    expect(html).toContain("Most severe still open");
    expect(html).toContain("An unfamiliar sign-in");
    expect(html).toContain("Many failed sign-ins from one address");
    expect(html).toContain("198.51.100.7");
    expect(html).toContain("Showing 2 of 12 open findings.");
    // Directory changes.
    expect(html).toContain("User management 9");
    expect(html).toContain("No category 1");
    expect(html).toContain("Flagged by the rules");
    expect(html).toContain("Global Administrator assigned");
    expect(html).toContain("Showing 2 of 3 flagged changes.");
    expect(html).toContain("History kept since");
    // Accounts that went quiet.
    expect(html).toContain("Rowan Ellis");
    expect(html).toContain("rowan@example.invalid");
    expect(html).toContain("33 days");
    expect(html).toContain("Days unknown");
    expect(html).toContain("Showing 2 of 52 accounts.");
    expect(html).toContain("Compared with the result from");
    // Audit changes.
    expect(html).toContain("MFA registration");
    expect(html).toContain("Admins without MFA");
    expect(html).toContain("0 to 2");
    expect(html).toContain("<strong>+2</strong>");
    expect(html).toContain("120 to 115");
    expect(html).toContain("<strong>-5</strong>");
    expect(html).toContain("Not directly comparable: another account scope, a partial result.");
    // Microsoft.
    expect(html).toContain("Retirement of a legacy sign-in method");
    expect(html).toContain("Exchange Online, Microsoft Teams");
    expect(html).toContain("Showing 1 of 22 posts.");
    expect(html).toContain("Some users cannot send mail");
    expect(html).toContain("Investigating, Exchange Online, first seen");
    // Collection.
    expect(html).toContain("Collecting");
    expect(html).toContain("Sign-in audit");
    expect(html).toContain("No success in the last 30 days of run history");
    expect(html).toContain("Last run failed");
    expect(html).toContain("2 failed runs (throttled, access_denied). 1 gap in what was read.");
    expect(html).toContain("Audits that run on their own");
    expect(html).toContain("Weekly.");
    expect(html).toContain("Has not run yet.");
  });

  it("says in one line what an empty section holds, and shows no list for it", () => {
    const html = render({ digest: emptyDigest });
    for (const line of [
      "No new sign-in findings this week.",
      "No directory changes this week.",
      "No accounts passed 30 days without a sign-in since the earlier result.",
      "No audit tile changed this week.",
      "No Message center posts have an act-by date in the next 30 days.",
      "Microsoft reported no incidents this week.",
    ])
      expect(html, line).toContain(line);
    for (const heading of [
      "Most severe still open",
      "Flagged by the rules",
      "Audits that run on their own",
    ])
      expect(html).not.toContain(heading);
    expect(html).not.toContain("Rowan Ellis");
    expect(html).not.toContain("An unfamiliar sign-in");
    expect(html).not.toContain('class="week-counts"');
  });

  it("gives the reason, in its own style, for a section that could not be read", () => {
    const html = render({ digest: unavailableDigest });
    for (const line of [
      "The sign-in audit is not available in this installation.",
      "The directory change feed is not available in this installation.",
      "Inactive accounts has no kept result from before this range to compare with. Schedule it daily or weekly.",
      "No audit has a kept result yet, so there is nothing to compare.",
      "Message center is not available in this installation.",
      "Service health is not available in this installation.",
      reason,
    ])
      expect(html, line).toContain(line);
    expect(html.match(/week-note week-note--unavailable/g)?.length).toBe(7);
    // An unavailable section is never worded as empty.
    expect(html).not.toContain("No new sign-in findings");
    expect(html).not.toContain("Microsoft reported no incidents");
    expect(html).not.toContain("Collecting<");
  });

  it("shows a loading line, and an error without hiding the ranges", () => {
    const loading = render({ digest: null });
    expect(loading).toContain("Loading this week");
    expect(loading).toContain("skeleton__row");
    expect(loading).toContain("Last 14 days");
    expect(loading).not.toContain("Sign-in findings");
    const failed = render({ digest: null, error: "Request failed with HTTP 500." });
    expect(failed).toContain('role="alert"');
    expect(failed).toContain("Request failed with HTTP 500.");
    // Nothing is loading after a failure, so the page does not say it is.
    expect(failed).not.toContain("Loading this week");
    expect(failed).toContain("Last 7 days");
  });

  it("uses plain words: no dashes between clauses, middle dots or exclamation marks", () => {
    for (const value of [digest, emptyDigest, unavailableDigest]) {
      const html = render({ digest: value });
      expect(html).not.toMatch(/[—–·!]/);
    }
  });

  it("escapes provider text instead of rendering it", () => {
    const html = render({
      digest: {
        ...digest,
        signInFindings: {
          ...digest.signInFindings,
          open: [
            finding(1, { title: "<img src=x onerror=alert(1)>", subjectLabel: "<b>Rowan</b>" }),
          ],
        },
      },
    });
    expect(html).not.toContain("<img src=x");
    expect(html).not.toContain("<b>Rowan</b>");
    expect(html).toContain("&lt;img src=x");
  });
});

describe("where each item opens", () => {
  it("opens a finding about a person on that person's page, and any other on its module's page", () => {
    expect(press("An unfamiliar sign-in")).toEqual([{ kind: "person", value: personId }]);
    expect(press("Many failed sign-ins from one address")).toEqual([
      { kind: "navigate", value: "signins" },
    ]);
    expect(press("Global Administrator assigned")).toEqual([{ kind: "person", value: personId }]);
    expect(press("Conditional Access policy changed")).toEqual([
      { kind: "navigate", value: "changes" },
    ]);
  });

  it("opens a quiet account on its page, and its audit with the audit's id", () => {
    expect(press("Rowan Ellis")).toEqual([{ kind: "person", value: personId }]);
    expect(press("Open Inactive accounts")).toEqual([
      { kind: "audit", value: "accounts.inactive" },
    ]);
  });

  it("does not make a link of an account that has no page", () => {
    const labels = clickables(createElement(WeekView, props().props)).map((item) => item.label);
    expect(labels).toContain("Rowan Ellis");
    expect(labels).not.toContain("shared@example.invalid");
  });

  it("opens each changed audit, and the pages the other panels belong to", () => {
    expect(press("MFA registration")).toEqual([{ kind: "audit", value: "mfa.registration" }]);
    expect(press("Open Audits")).toEqual([{ kind: "navigate", value: "audits" }]);
    expect(press("Open sign-in audit")).toEqual([{ kind: "navigate", value: "signins" }]);
    expect(press("Open Directory changes")).toEqual([{ kind: "navigate", value: "changes" }]);
    expect(press("Open Notifications")).toEqual([{ kind: "navigate", value: "notifications" }]);
    expect(press("Retirement of a legacy sign-in method")).toEqual([
      { kind: "navigate", value: "notifications" },
    ]);
    expect(press("Some users cannot send mail")).toEqual([
      { kind: "navigate", value: "notifications" },
    ]);
  });

  it("opens the scheduled audits listed under Collection", () => {
    const { props: value, calls } = props();
    const matching = clickables(createElement(WeekView, value)).filter(
      (item) => item.label === "Inactive accounts",
    );
    // The changed audit and the scheduled one are both Inactive accounts.
    expect(matching).toHaveLength(2);
    for (const item of matching) item.click();
    expect(calls).toEqual([
      { kind: "audit", value: "accounts.inactive" },
      { kind: "audit", value: "accounts.inactive" },
    ]);
  });

  it("offers Start collecting, to the page that does it, when collection is off", () => {
    const off = { ...digest, collection: { ...digest.collection, collecting: false } };
    expect(render({ digest: off })).toContain("Not collecting");
    expect(press("Start collecting", { digest: off })).toEqual([
      { kind: "navigate", value: "signins" },
    ]);
  });

  it("changes the range when a range button is pressed", () => {
    expect(press("Last 30 days")).toEqual([{ kind: "days", value: 30 }]);
    expect(press("Last day")).toEqual([{ kind: "days", value: 1 }]);
  });

  it("works out the target of a finding", () => {
    expect(findingTarget(finding(1), "signins")).toEqual({ type: "person", id: personId });
    expect(findingTarget(finding(1, { personId: null }), "changes")).toEqual({
      type: "view",
      view: "changes",
    });
  });
});

describe("small pieces", () => {
  it("writes a change with its sign", () => {
    expect(signed(5)).toBe("+5");
    expect(signed(-12)).toBe("-12");
    expect(signed(0)).toBe("0");
  });

  it("says why two results do not compare, and nothing when they do", () => {
    expect(differenceText([])).toBeNull();
    expect(differenceText(["version"])).toBe(
      "Not directly comparable: another version of the audit.",
    );
    expect(differenceText(["version", "scope", "incomplete"])).toBe(
      "Not directly comparable: another version of the audit, another account scope, a partial result.",
    );
  });

  it("describes how far change collection has read, and nothing before it has", () => {
    expect(
      coverageText({ from: null, through: null, lastCollected: null, phase: "none" }),
    ).toBeNull();
    expect(markup(coverageText(digest.directoryChanges.coverage))).toContain("History kept since");
    expect(coverageText({ ...digest.directoryChanges.coverage, phase: "history" })).toContain(
      "Loading the last 30 days, oldest first",
    );
  });

  it("shows an answer only under the range it was asked for", () => {
    expect(digestFor(digest, 7)).toBe(digest);
    expect(digestFor(digest, 30)).toBeNull();
    expect(digestFor(null, 7)).toBeNull();
  });

  it("asks for the range it shows", () => {
    expect(digestUrl(14)).toBe("/api/v1/digest?days=14");
  });

  it("builds fixtures that match the digest contract", () => {
    for (const value of [digest, emptyDigest, unavailableDigest])
      expect(digestSchema.parse(value)).toEqual(value);
  });
});

describe("the page for a session", () => {
  const links = { onNavigate: vi.fn(), onOpenPerson: vi.fn(), onOpenAudit: vi.fn() };
  const page = (value: PrismSession) =>
    renderToStaticMarkup(createElement(ThisWeek, { session: value, ...links }));

  it("asks demo mode and a signed-out visitor to connect, and shows no invented data", () => {
    const demo = page(session({ mode: "demo", authenticated: false, configured: false }));
    expect(demo).toContain("Demo mode shows none of it");
    expect(demo).not.toContain("Sign-in findings");
    const out = page(session({ authenticated: false, csrfToken: null }));
    expect(out).toContain("Sign in to see this week.");
  });

  it("starts a signed-in page on the last 7 days, loading", () => {
    const html = page(session());
    expect(html).toContain("Loading this week");
    expect(html).toMatch(/aria-pressed="true"[^>]*>Last 7 days</);
  });
});

describe("the page's address and place in the navigation", () => {
  it("opens This week from #/week", async () => {
    const { parseRoute, routeHash } = await import("../src/App.tsx");
    expect(parseRoute("#/week")).toEqual({ view: "week", person: null, audit: null });
    expect(routeHash({ view: "week", person: null, audit: null })).toBe("#/week");
  });

  it("lists This week right after Overview, and each page once", async () => {
    const { App } = await import("../src/App.tsx");
    const html = renderToStaticMarkup(createElement(App));
    // Only the navigation: the Search button above it opens a dialog and is not a page.
    const nav = html.match(/<nav\b[\s\S]*?<\/nav>/)?.[0] ?? "";
    const labels = [...nav.matchAll(/<span>([^<]+)<\/span><\/button>/g)].map((match) => match[1]);
    expect(labels).toEqual([
      "Overview",
      "This week",
      "People",
      "Microsoft updates",
      "Audits",
      "Sign-ins",
      "Directory changes",
      "Handoffs",
      "Alerts",
      "Microsoft access",
    ]);
  });

  it("sends the folded pages' old links to the page that took them over", async () => {
    const { parseRoute } = await import("../src/App.tsx");
    expect(parseRoute("#/users")).toEqual({ view: "people", person: null, audit: null });
    expect(parseRoute("#/lifecycle")).toEqual({ view: "handoffs", person: null, audit: null });
  });
});

describe("This week in the demo", () => {
  // Recorded on 2026-10-22; shown three days later, so every date moves forward three days.
  const fixtures = {
    recordedAt: "2026-10-22T12:00:00.000Z",
    responses: { [`GET ${digestUrl(7)}`]: { status: 200, body: digest } },
  };
  const state = createDemoState(fixtures, () => new Date("2026-10-25T15:00:00.000Z"));

  it("replays the recorded page for the range it was recorded for, with its dates moved", () => {
    const answer = handleRequest(state, { method: "GET", url: digestUrl(7) });
    expect(answer.status).toBe(200);
    const shown = digestSchema.parse(answer.body);
    expect(shown.generatedAt).toBe("2026-10-25T12:00:00.000Z");
    expect(shown.from).toBe("2026-10-18T12:00:00.000Z");
    expect(shown.signInFindings.open[0]?.firstSeen).toBe("2026-10-24T09:00:00.000Z");
    expect(shown.microsoft.posts.items[0]?.actBy).toBe("2026-11-02T00:00:00.000Z");
    expect(shown.signInFindings.note).toBe(digest.signInFindings.note);
  });

  it("says a range that was not recorded is not available, rather than inventing it", () => {
    const answer = handleRequest(state, { method: "GET", url: digestUrl(30) });
    expect(answer.status).toBe(404);
  });
});

describe("the page speaks the shared visual language", () => {
  const NOW = "2026-10-22T12:00:00.000Z";
  const html = at(NOW, () => render());

  it("shows when something happened as words, with the exact moment on hover", () => {
    // A finding first seen the day before, an incident hours ago, a collector minutes ago.
    expect(html).toMatch(relativeCell("2026-10-21T09:00:00.000Z", "yesterday"));
    expect(html).toMatch(relativeCell("2026-10-22T08:00:00.000Z", "4 h ago"));
    expect(html).toMatch(relativeCell("2026-10-22T11:55:00.000Z", "5 min ago"));
    expect(html).toMatch(relativeCell("2026-10-22T11:50:00.000Z", "10 min ago"));
    // A result from more than a week back is its date, still a time element.
    expect(html).toMatch(/<time dateTime="2026-10-14T10:00:00.000Z" title="[^"]+">[^<]+<\/time>/);
  });

  it("keeps the date a person has to act by as a date", () => {
    expect(html).toContain("Act by ");
    expect(html).not.toContain('dateTime="2026-10-30T00:00:00.000Z"');
  });

  it("marks severity with the shared dot and name, in the lists and in the counts", () => {
    expect(html).toContain(severityDot("critical", "Critical"));
    expect(html).toContain(severityDot("high", "High"));
    expect(html).toContain(severityDot("medium", "Medium"));
    expect(html).toContain(severityDot("critical", "1 critical"));
    expect(html).toContain(severityDot("high", "2 high"));
    expect(html).not.toContain("finding-severity");
  });

  it("draws still skeleton rows while the answer is on its way", () => {
    expectSkeleton(render({ digest: null }), "Loading this week");
  });
});

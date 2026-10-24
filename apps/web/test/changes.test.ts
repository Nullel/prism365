import {
  type ChangeEvent,
  type ChangeFeed,
  type Finding,
  type ModuleList,
  type PersonChangePage,
  type PrismSession,
  sessionSchema,
} from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import {
  type ChangesViewProps,
  ChangesView,
  categoryLabel,
  changesQuery,
  DirectoryChanges,
  defaultChangeFilters,
  startHistory,
} from "../src/Changes.tsx";
import { PersonChangeList, PersonChanges } from "../src/PersonChanges.tsx";

// Synthetic people, apps and changes only.

function session(overrides: Partial<PrismSession> = {}): PrismSession {
  return sessionSchema.parse({
    mode: "live",
    configured: true,
    authenticated: true,
    displayName: "Synthetic admin",
    csrfToken: "synthetic-csrf-token-value-000000000000",
    enabledFeatures: ["directory.users.read", "audit.directory.read"],
    missingConfiguration: [],
    ...overrides,
  });
}

const personId = "44444444-4444-4444-8444-444444444444";

function change(id: string, at: string, overrides: Partial<ChangeEvent> = {}): ChangeEvent {
  return {
    id,
    activityAt: at,
    activity: "Update user",
    category: "UserManagement",
    operationType: "Update",
    result: "success",
    resultReason: null,
    loggedByService: "Core Directory",
    correlationId: null,
    initiatedBy: {
      kind: "user",
      id: "55555555-5555-4555-8555-555555555555",
      name: "Avery Stone",
      principal: "avery@example.invalid",
    },
    targets: [
      {
        id: personId,
        name: "Rowan Ellis",
        type: "User",
        principal: "rowan@example.invalid",
        properties: [
          {
            name: "JobTitle",
            oldValue: "Analyst",
            newValue: "Lead analyst",
            oldCut: false,
            newCut: false,
          },
          {
            name: "Description",
            oldValue: null,
            newValue: "x".repeat(400),
            oldCut: false,
            newCut: true,
          },
        ],
      },
    ],
    details: [{ key: "UserType", value: "Member" }],
    ...overrides,
  };
}

const items = [
  change("c-3", "2026-10-22T15:30:00.000Z"),
  change("c-2", "2026-10-22T09:05:00.000Z", {
    activity: "Add member to role",
    category: "RoleManagement",
    initiatedBy: { kind: "app", id: "sp-1", name: "Synthetic Sync", principal: "app-1" },
  }),
  change("c-1", "2026-10-20T12:00:00.000Z", { result: "failure", targets: [] }),
];

const feed: ChangeFeed = {
  days: 7,
  items,
  next: "2026-10-20T12:00:00.000Z|c-1",
  categories: [
    { category: "UserManagement", count: 2 },
    { category: "RoleManagement", count: 1 },
  ],
  coverage: {
    from: "2026-09-22T12:00:00.000Z",
    through: "2026-10-22T16:00:00.000Z",
    lastCollected: "2026-10-22T16:00:00.000Z",
    phase: "live",
  },
};

const finding: Finding = {
  id: "66666666-6666-4666-8666-666666666666",
  module: "directory.changes",
  ruleId: "role.privileged-assigned",
  ruleVersion: 1,
  purpose: "security",
  severity: "critical",
  score: 90,
  subject: { kind: "person", id: personId, label: "Rowan Ellis" },
  title: "Global Administrator assigned",
  summary: "Avery Stone added Rowan Ellis to Global Administrator.",
  evidence: [
    {
      source: "Microsoft Entra audit log",
      ref: "c-2",
      observedAt: "2026-10-22T09:05:00.000Z",
      fields: { role: "Global Administrator" },
    },
  ],
  limitations: ["Synthetic."],
  windowStart: "2026-10-22T09:05:00.000Z",
  windowEnd: "2026-10-22T09:05:00.000Z",
  state: "open",
  occurrences: 1,
  firstSeen: "2026-10-22T09:06:00.000Z",
  lastSeen: "2026-10-22T09:06:00.000Z",
  stateChangedAt: null,
};

const modules = (lastRun: unknown, missingPermissions: string[] = []): ModuleList =>
  ({
    mode: "live",
    collection: {
      available: true,
      state: "collecting",
      credentialKind: "delegated",
      boundAt: "2026-10-22T08:00:00.000Z",
      lastStopReason: null,
      resumes: false,
      durableStorage: true,
      skipped: [],
    },
    modules: [
      {
        id: "directory.changes",
        title: "Directory changes",
        description: "Synthetic.",
        packs: ["audit.directory.read"],
        permissions: ["AuditLog.Read.All"],
        missingPermissions,
        collectors: [{ id: "audits", title: "Directory audit log", cadenceSeconds: 600, lastRun }],
      },
    ],
  }) as unknown as ModuleList;

function view(overrides: Partial<ChangesViewProps> = {}) {
  return renderToStaticMarkup(
    createElement(ChangesView, {
      session: session(),
      modules: null,
      feed,
      items,
      findings: [finding],
      filters: defaultChangeFilters,
      error: "",
      busy: null,
      onFilters: () => {},
      onMore: () => {},
      onFindingState: () => {},
      onOpenPerson: () => {},
      ...overrides,
    }),
  );
}

describe("directory changes page", () => {
  it("shows no invented changes in demo mode, and asks for sign-in first", () => {
    const demo = renderToStaticMarkup(
      createElement(DirectoryChanges, {
        session: session({ mode: "demo", authenticated: false, configured: false }),
        onOpenPerson: () => {},
      }),
    );
    expect(demo).toContain("does not invent");
    expect(demo).not.toContain("Flagged changes");
    const signedOut = renderToStaticMarkup(
      createElement(DirectoryChanges, {
        session: session({ authenticated: false, csrfToken: null }),
        onOpenPerson: () => {},
      }),
    );
    expect(signedOut).toContain("Sign in to see directory changes");
  });

  it("starts with the page loading", () => {
    const html = renderToStaticMarkup(
      createElement(DirectoryChanges, { session: session(), onOpenPerson: () => {} }),
    );
    expect(html).toContain("Flagged changes");
    expect(html).toContain("Loading…");
  });

  it("puts flagged changes first, then the feed by day, with expandable values", () => {
    const html = view();
    expect(html.indexOf("Flagged changes")).toBeLessThan(html.indexOf("Changes, last 7 days"));
    expect(html).toContain("Global Administrator assigned");
    // Two days, the newest first.
    const days = [...html.matchAll(/<h4 class="change-day">([^<]+)<\/h4>/g)].map((m) => m[1]);
    expect(days).toHaveLength(2);
    // A user who made a change, and one it was made to, open their person pages; an app does not.
    expect(html).toContain('<button class="link-button" type="button">Avery Stone</button>');
    expect(html).toContain('<button class="link-button" type="button">Rowan Ellis</button>');
    expect(html).toContain("<span>Synthetic Sync</span>");
    expect(html).toContain("Role management");
    expect(html).toContain("What changed");
    expect(html).toContain("Lead analyst");
    expect(html).toContain("(cut at 400 characters)");
    expect(html).toContain("No target recorded");
    expect(html).toContain("Failed");
    expect(html).toContain("Show older changes");
    expect(html).toContain("User management (2)");
    expect(html).toContain("History kept since");
    // The feed's own text has no dashes or middle dots (the shared finding card is not this page's).
    expect(html.slice(html.indexOf("change-panel"))).not.toMatch(/[\u2014\u00b7]/);
  });

  it("says plainly when nothing is stored, nothing matches, or nothing is flagged", () => {
    const none = view({
      feed: {
        ...feed,
        items: [],
        next: null,
        categories: [],
        coverage: {
          ...feed.coverage,
          phase: "none",
          from: null,
          through: null,
          lastCollected: null,
        },
      },
      items: [],
      findings: [],
    });
    expect(none).toContain(startHistory);
    expect(none).toContain("Nothing flagged");
    expect(none).not.toContain("Show older changes");
    const filtered = view({
      feed: { ...feed, items: [], next: null },
      items: [],
      filters: { ...defaultChangeFilters, search: "nobody" },
    });
    expect(filtered).toContain("No changes match these filters");
  });

  it("explains a refused read, a missing grant and a pack that is not enabled", () => {
    const denied = view({
      modules: modules({
        status: "failed",
        credentialKind: "delegated",
        startedAt: "2026-10-22T16:00:00.000Z",
        finishedAt: "2026-10-22T16:00:02.000Z",
        fetched: 0,
        inserted: 0,
        findings: 0,
        skipped: 0,
        errorCategory: "access_denied",
      }),
    });
    expect(denied).toContain("Reports Reader");
    expect(denied).toContain("Entra ID P1 or P2");
    expect(view({ modules: modules(null, ["AuditLog.Read.All"]) })).toContain(
      "Grant directory audit access",
    );
    expect(view({ session: session({ enabledFeatures: ["directory.users.read"] }) })).toContain(
      "Add audit.directory.read to ENABLED_READ_FEATURES",
    );
  });

  it("builds feed requests from the filters", () => {
    expect(changesQuery(defaultChangeFilters)).toBe("/api/v1/changes?days=7&limit=100");
    expect(
      changesQuery(
        { days: 30, category: "RoleManagement", result: "failure", search: " Rowan " },
        "2026-10-20T12:00:00.000Z|c-1",
      ),
    ).toBe(
      "/api/v1/changes?days=30&limit=100&category=RoleManagement&result=failure&search=Rowan&before=2026-10-20T12%3A00%3A00.000Z%7Cc-1",
    );
    expect(categoryLabel("ApplicationManagement")).toBe("Application management");
    expect(categoryLabel(null)).toBe("No category");
  });
});

describe("person page directory changes", () => {
  const page: PersonChangePage = {
    items: [
      { ...change("p-2", "2026-10-22T15:30:00.000Z"), involvement: { actor: false, target: true } },
      {
        ...change("p-1", "2026-10-21T10:00:00.000Z"),
        involvement: { actor: true, target: true },
      },
    ],
    next: "2026-10-21T10:00:00.000Z|p-1",
  };
  const panel = (overrides: Partial<Parameters<typeof PersonChangeList>[0]> = {}) =>
    renderToStaticMarkup(
      createElement(PersonChangeList, {
        page,
        error: "",
        busy: false,
        onMore: () => {},
        ...overrides,
      }),
    );

  it("lists the newest changes with how the person was involved", () => {
    const html = panel();
    expect(html).toContain("Directory changes");
    expect(html).toContain("Made to this person");
    expect(html).toContain("Made by and to this person");
    expect(html).toContain("Show older changes");
  });

  it("has plain loading, empty and error states, and nothing in demo mode", () => {
    expect(panel({ page: null })).toContain("Loading changes");
    expect(panel({ page: { items: [], next: null } })).toContain(startHistory);
    expect(
      panel({ error: "The directory change feed is not available in this installation." }),
    ).toContain("not available in this installation");
    expect(
      renderToStaticMarkup(
        createElement(PersonChanges, {
          session: session({ mode: "demo", authenticated: false }),
          userId: personId,
        }),
      ),
    ).toBe("");
  });
});

describe("the page's address", () => {
  it("opens Directory changes from #/changes", async () => {
    const { parseRoute, routeHash } = await import("../src/App.tsx");
    expect(parseRoute("#/changes")).toEqual({ view: "changes", person: null, audit: null });
    expect(routeHash({ view: "changes", person: null, audit: null })).toBe("#/changes");
  });
});

import {
  type AuditCatalog,
  type AuditColumn,
  type AuditRow,
  type AuditRun,
  type AuditRunDetail,
  type AuditSettings,
  auditDefinitionSchema,
} from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import {
  AuditCatalogList,
  AuditCatalogPage,
  AuditScopeEditor,
  AuditScopeSection,
  auditHref,
  auditResultState,
  auditResultCounts,
  exportHref,
  headlineTile,
  lowerFirst,
  ResultTable,
  RunScopeTags,
  rowEmails,
  runResult,
  runResultParts,
  scopeLimit,
  scopeSummary,
} from "../src/Audits.tsx";
import { findAll, hostTree, nestedInteractive, plainClick, textOf } from "./markup.ts";

const definition = auditDefinitionSchema.parse({
  id: "accounts.inactive",
  version: 1,
  title: "Inactive accounts",
  description: "Every account and how long it has gone without signing in.",
  category: "Accounts",
  subject: { one: "account", many: "accounts" },
  columns: [{ key: "person", label: "Account", kind: "person" }],
  views: [{ id: "all", label: "All", tags: [] }],
  defaultView: "all",
  defaultSort: { column: "person", direction: "asc" },
  permissions: ["User.Read.All", "AuditLog.Read.All"],
  optionalPermissions: [],
  requirements: [],
});

const run: AuditRun = {
  id: "66666666-6666-4666-8666-666666666666",
  auditId: "accounts.inactive",
  auditVersion: 1,
  status: "succeeded",
  complete: true,
  startedAt: new Date(Date.now() - 5 * 60_000).toISOString(),
  finishedAt: new Date().toISOString(),
  startedBy: { id: "22222222-2222-4222-8222-222222222222", name: "Synthetic admin" },
  error: null,
  summary: [
    { label: "Accounts", value: 214, tone: "neutral", view: "all" },
    { label: "Inactive 30+ days", value: 47, tone: "warning", view: "d30" },
  ],
  notes: [],
  rowCount: 214,
};

const NOW = Date.now();

const catalog = (
  overrides: Partial<AuditCatalog["audits"][number]> = {},
  mode: "live" | "demo" = "live",
) =>
  ({
    mode,
    audits: [
      {
        definition,
        availability: { state: "available", missing: [], grantFeature: null, enable: [] },
        lastRun: null,
        runningId: null,
        ...overrides,
      },
    ],
  }) as AuditCatalog;

const render = (value: AuditCatalog) =>
  renderToStaticMarkup(
    createElement(AuditCatalogList, {
      catalog: value,
      busy: null,
      onOpen: () => {},
      onRun: () => {},
      now: NOW,
    }),
  );

describe("audit catalog", () => {
  it("offers a run for an available audit that has never run", () => {
    const html = render(catalog());
    expect(html).toContain("Inactive accounts");
    expect(html).toContain('<span class="audit-result audit-result--none" title="Not run yet">');
    expect(html).toContain(">Not run yet<");
    expect(html).toContain('aria-label="Run Inactive accounts"');
    expect(html).not.toContain("disabled");
  });

  it("says what the last run found without repeating when it ran", () => {
    expect(runResult(run, definition)).toBe("214 accounts, 47 inactive 30+ days");
    expect(runResult({ ...run, rowCount: 1, summary: [] }, definition)).toBe("1 account");
    expect(runResult({ ...run, complete: false }, definition)).toBe(
      "214 accounts, 47 inactive 30+ days, partial result",
    );
    expect(
      runResult(
        { ...run, status: "failed", error: { code: "forbidden", message: "Microsoft refused." } },
        definition,
      ),
    ).toBe("Failed: Microsoft refused.");
    expect(runResult({ ...run, status: "running" }, definition)).toBe("Running");
    expect(runResult(null, definition)).toBe("Not run yet");
  });

  it("links to consent and to Microsoft access when the session lacks a permission", () => {
    const html = render(
      catalog({
        availability: {
          state: "missing_permission",
          missing: ["AuditLog.Read.All"],
          grantFeature: "audit.signins.read",
          enable: [],
        },
      }),
    );
    expect(html).toContain('href="/api/v1/auth/login?feature=audit.signins.read"');
    expect(html).toContain("Needs AuditLog.Read.All");
    expect(html).toContain('<a href="#/access">Microsoft access</a>');
    expect(html).toMatch(/<button class="[^"]*" type="button"[^>]* disabled="">Run<\/button>/);
    // Said in the Result cell, in place of a result there is none of.
    expect(html).not.toContain("Not run yet");
  });

  it("explains an audit that is not enabled, and still shows the result it has", () => {
    const html = render(
      catalog({
        availability: {
          state: "not_enabled",
          missing: ["Policy.Read.All"],
          grantFeature: null,
          enable: ["identity.policies.read"],
        },
        lastRun: run,
      }),
    );
    expect(html).toContain(
      "Not enabled here. Add identity.policies.read to ENABLED_READ_FEATURES.",
    );
    expect(html).toContain("214 accounts, 47 inactive 30+ days");
    expect(html).toContain('<a href="#/access">Microsoft access</a>');
    expect(html).not.toContain("Grant access");
    expect(html).toMatch(/ disabled="">Run<\/button>/);
  });

  it("cannot run anything in demo mode", () => {
    const html = render(catalog({}, "demo"));
    expect(html).toContain("Demo mode lists the audits but cannot run them");
    expect(html).toMatch(/disabled=""/);
  });

  it("shows a running audit as Running, with its Run button off", () => {
    const html = render(
      catalog({ lastRun: { ...run, status: "running", finishedAt: null }, runningId: run.id }),
    );
    expect(html).toContain('<span class="audit-result audit-result--running" title="Running">');
    expect(html).toContain(">Running<");
    expect(html).toMatch(/aria-busy="true" disabled="">Run<\/button>/);
  });

  it("shows Running as soon as the Run button is pressed, before the server answers", () => {
    const html = renderToStaticMarkup(
      createElement(AuditCatalogList, {
        catalog: catalog({ lastRun: run }),
        busy: "accounts.inactive",
        onOpen: () => {},
        onRun: () => {},
        now: NOW,
      }),
    );
    expect(html).toContain("audit-result--running");
    expect(html).toContain(">Running<");
  });
});

// A second category and audits in several states, for the table as a whole.
const second = auditDefinitionSchema.parse({
  ...definition,
  id: "licenses.reclaim",
  title: "Licenses to reclaim",
  description: "Licenses held by accounts that no longer need them.",
  category: "Licenses",
  subject: { one: "license", many: "licenses" },
});
const third = auditDefinitionSchema.parse({
  ...definition,
  id: "mfa.registration",
  title: "MFA registration",
  description: "Who has not registered a second sign-in method.",
  category: "Accounts",
});
const available = { state: "available", missing: [], grantFeature: null, enable: [] } as const;
const ago = (minutes: number) => new Date(NOW - minutes * 60_000).toISOString();

const table: AuditCatalog = {
  mode: "live",
  audits: [
    {
      definition,
      availability: available,
      lastRun: {
        ...run,
        startedAt: ago(5),
        summary: [{ label: "Admins without MFA", value: 1, tone: "critical", view: null }],
      },
      runningId: null,
    },
    { definition: third, availability: available, lastRun: null, runningId: null },
    {
      definition: second,
      availability: available,
      lastRun: {
        ...run,
        auditId: "licenses.reclaim",
        startedAt: ago(26 * 60),
        summary: [{ label: "Accounts", value: 9, tone: "neutral", view: null }],
        rowCount: 9,
      },
      runningId: null,
    },
  ],
};

const renderTable = (
  value: AuditCatalog,
  schedule: Record<string, "off" | "daily" | "weekly"> | null = null,
) =>
  renderToStaticMarkup(
    createElement(AuditCatalogList, {
      catalog: value,
      busy: null,
      onOpen: () => {},
      onRun: () => {},
      schedule,
      now: NOW,
    }),
  );

describe("audit catalog table", () => {
  const html = renderTable(table, { "accounts.inactive": "daily", "licenses.reclaim": "off" });

  it("is one table with a column for each thing worth comparing", () => {
    expect(html.match(/<table/g)).toHaveLength(1);
    expect(html).not.toContain("<ul");
    const headers = [
      ...html.matchAll(/<th scope="col">(?:<span class="visually-hidden">)?([^<]+)/g),
    ];
    expect(headers.map((match) => match[1])).toEqual([
      "Audit",
      "Last run",
      "Result",
      "Schedule",
      "Run",
    ]);
  });

  it("groups the audits by category under header rows, in the order the categories appear", () => {
    const groups = [...html.matchAll(/<th scope="colgroup" colSpan="5">([^<]+)</g)];
    expect(groups.map((match) => match[1])).toEqual(["Accounts", "Licenses"]);
    // MFA registration was listed after Licenses to reclaim but belongs under Accounts.
    expect(html.indexOf("MFA registration")).toBeLessThan(html.indexOf("Licenses to reclaim"));
    expect(html.match(/catalog-table__row/g)).toHaveLength(3);
  });

  it("makes the audit's name the link to its results and cuts the purpose to one line", () => {
    expect(auditHref("accounts.inactive")).toBe("#/audits/accounts.inactive");
    expect(html).toContain(
      '<a class="catalog-table__link" href="#/audits/accounts.inactive">Inactive accounts</a>',
    );
    // One line by style; the whole sentence stays available as the title.
    expect(html).toContain(
      '<span class="catalog-table__purpose" title="Every account and how long it has gone without signing in.">',
    );
  });

  it("shows when each audit last ran as a time element, and Never for one that has not", () => {
    expect(html).toContain("5 min ago</time>");
    expect(html).toContain("yesterday</time>");
    expect(html).toContain(
      '<td data-label="Last run" class="catalog-table__when"><span class="person-muted">never</span>',
    );
  });

  it("keeps the total in plain text and colours only the number worth a look", () => {
    expect(html).toContain('class="audit-result audit-result--critical"');
    expect(html).toContain('class="audit-result audit-result--clear"');
    expect(html).toContain('class="audit-result audit-result--none"');
    // The total sits outside the coloured part, which holds the dot, the hidden severity word and
    // the issue, and nothing else.
    expect(html).toContain(
      '<span class="audit-result__total">214 accounts</span><span class="audit-result__issue audit-result__issue--critical">' +
        '<span class="sev__dot sev__dot--critical" aria-hidden="true"></span>' +
        '<span class="visually-hidden">Critical: </span>' +
        '<span class="audit-result__issue-text">1 admins without MFA</span></span>',
    );
    // A run with nothing to flag has a total and no coloured part at all.
    expect(html).toContain('<span class="audit-result__total">9 licenses</span></span>');
    const licenses = html
      .split('<tr class="catalog-table__row">')
      .find((row) => row.includes("Licenses to reclaim"));
    expect(licenses).toBeDefined();
    expect(licenses).not.toContain("audit-result__issue");
  });

  it("colours a warning differently from a critical one, in the same cell structure", () => {
    const warning = renderTable({
      mode: "live",
      audits: [
        {
          definition,
          availability: available,
          lastRun: {
            ...run,
            summary: [{ label: "Inactive 30+ days", value: 17, tone: "warning", view: null }],
          },
          runningId: null,
        },
      ],
    });
    expect(warning).toContain('class="audit-result__issue audit-result__issue--warning"');
    expect(warning).toContain('<span class="sev__dot sev__dot--medium" aria-hidden="true"></span>');
    expect(warning).toContain('<span class="visually-hidden">Warning: </span>');
    expect(warning).toContain("17 inactive 30+ days");
  });

  it("builds every row from the same cells, so no row looks different from the first", () => {
    const many: AuditCatalog = {
      mode: "live",
      audits: [
        definition,
        third,
        second,
        auditDefinitionSchema.parse({
          ...third,
          id: "guests.review",
          title: "Guests",
          category: "Licenses",
        }),
      ].map((entry, index) => ({
        definition: entry,
        availability: available,
        lastRun: index % 2 === 0 ? run : null,
        runningId: null,
      })),
    };
    const rows = [
      ...renderTable(many, { "accounts.inactive": "daily" }).matchAll(
        /<tr class="catalog-table__row">(.*?)<\/tr>/g,
      ),
    ].map((match) =>
      [...(match[1] ?? "").matchAll(/<(th|td)([^>]*)>/g)].map(
        (cell) => `${cell[1]}${(cell[2] ?? "").replace(/ colSpan="\d+"/, "")}`,
      ),
    );
    expect(rows).toHaveLength(4);
    // Same cells in the same order with the same labels for every row, whatever it holds.
    for (const row of rows) expect(row).toEqual(rows[0]);
    expect(rows[0]).toEqual([
      'th scope="row" class="catalog-table__name"',
      'td data-label="Last run" class="catalog-table__when"',
      'td data-label="Result" class="catalog-table__result"',
      'td data-label="Schedule" class="catalog-table__schedule"',
      'td class="catalog-table__run"',
    ]);
  });

  it("labels the schedule as Off, Daily or Weekly, and does not guess before it has loaded", () => {
    expect(html).toContain('<span class="person-tag">Daily</span>');
    expect(html).toContain('<span class="person-muted">Off</span>');
    const weekly = renderTable(table, { "accounts.inactive": "weekly" });
    expect(weekly).toContain('<span class="person-tag">Weekly</span>');
    // An audit missing from the schedule is Off; no schedule at all is unknown.
    expect(weekly).toContain('<span class="person-muted">Off</span>');
    expect(renderTable(table, null)).not.toContain(">Off<");
  });

  it("has no link inside a button or button inside a link, whatever state the rows are in", () => {
    expect(nestedInteractive(html)).toBe(0);
    const stateful: AuditCatalog = {
      mode: "live",
      audits: [
        {
          definition,
          availability: {
            state: "missing_permission",
            missing: ["AuditLog.Read.All"],
            grantFeature: "audit.signins.read",
            enable: [],
          },
          lastRun: run,
          runningId: null,
        },
        { definition: third, availability: available, lastRun: null, runningId: third.id },
        { definition: second, availability: available, lastRun: null, runningId: null },
      ],
    };
    expect(nestedInteractive(renderTable(stateful, { "accounts.inactive": "daily" }))).toBe(0);
    expect(nestedInteractive("<a href='#'><button>x</button></a>")).toBe(1);
    expect(nestedInteractive("<button><a href='#'>x</a></button>")).toBe(1);
  });

  it("opens the results on a plain click and leaves a click for a new tab to the browser", () => {
    const opened: string[] = [];
    const ran: string[] = [];
    const tree = hostTree(
      createElement(AuditCatalogList, {
        catalog: table,
        busy: null,
        onOpen: (id: string) => opened.push(id),
        onRun: (id: string) => ran.push(id),
        now: NOW,
      }),
    );
    const link = findAll(tree, (host) => host.props.className === "catalog-table__link")[0];
    expect(link?.props.href).toBe("#/audits/accounts.inactive");
    const onClick = link?.props.onClick as (event: unknown) => void;
    onClick({ ...plainClick(), ctrlKey: true });
    onClick({ ...plainClick(), button: 1 });
    expect(opened).toEqual([]);
    onClick(plainClick());
    expect(opened).toEqual(["accounts.inactive"]);

    const buttons = findAll(tree, (host) => host.tag === "button");
    expect(buttons.map((button) => textOf(button))).toEqual(["Run", "Run", "Run"]);
    const last = buttons[2];
    expect(last).toBeDefined();
    (last?.props.onClick as (() => void) | undefined)?.();
    expect(ran).toEqual(["licenses.reclaim"]);
  });
});

describe("audit result states", () => {
  const tile = (tone: "neutral" | "good" | "warning" | "critical", value = 3) => ({
    label: `A ${tone} number`,
    value,
    tone,
    view: null,
  });
  const withTiles = (...summary: ReturnType<typeof tile>[]): AuditRun => ({ ...run, summary });

  it("leads with the first number that is neither neutral nor good and is more than zero", () => {
    expect(headlineTile(withTiles(tile("neutral"), tile("good"), tile("warning")))?.tone).toBe(
      "warning",
    );
    expect(headlineTile(withTiles(tile("warning", 0), tile("critical")))?.tone).toBe("critical");
    expect(headlineTile(withTiles(tile("warning"), tile("critical")))?.tone).toBe("warning");
    expect(headlineTile(withTiles(tile("neutral"), tile("good")))).toBeNull();
    expect(headlineTile(null)).toBeNull();
    expect(headlineTile({ ...run, status: "failed" })).toBeNull();
  });

  it("names the state of a run for its colour", () => {
    expect(auditResultState(null)).toBe("none");
    expect(auditResultState({ ...run, status: "running" })).toBe("running");
    expect(auditResultState({ ...run, status: "failed" })).toBe("failed");
    expect(auditResultState(withTiles(tile("critical")))).toBe("critical");
    expect(auditResultState(withTiles(tile("warning")))).toBe("warning");
    expect(auditResultState(withTiles(tile("good"), tile("neutral")))).toBe("clear");
  });

  it("counts critical results, treating a failed run as critical, and the warnings beside them", () => {
    const entry = (lastRun: AuditRun | null, runningId: string | null = null) => ({
      definition,
      availability: available,
      lastRun,
      runningId,
    });
    const value: AuditCatalog = {
      mode: "live",
      audits: [
        entry(withTiles(tile("warning"))),
        entry(withTiles(tile("warning"))),
        entry(withTiles(tile("critical"))),
        entry({ ...run, status: "failed" }),
        entry(withTiles(tile("good"), tile("neutral"))),
        entry(withTiles(tile("warning", 0))),
        entry({ ...run, status: "running", summary: [] }, run.id),
        entry(null),
      ],
    };
    // Seven have a run (one of them still going); the audit that never ran is not counted.
    expect(auditResultCounts(value)).toEqual({ critical: 2, warning: 2, ran: 7 });
    expect(auditResultCounts({ ...value, audits: [] })).toEqual({
      critical: 0,
      warning: 0,
      ran: 0,
    });
  });
});

describe("result wording", () => {
  it("lowercases the first letter of a label only when the second is lower case", () => {
    expect(lowerFirst("Inactive 30+ days")).toBe("inactive 30+ days");
    expect(lowerFirst("No Microsoft 365 app activity in 90 days")).toBe(
      "no Microsoft 365 app activity in 90 days",
    );
    expect(lowerFirst("Admins without MFA")).toBe("admins without MFA");
    // Acronyms, numbers and one-letter starts are left as they are.
    expect(lowerFirst("MFA not registered")).toBe("MFA not registered");
    expect(lowerFirst("A1 licenses")).toBe("A1 licenses");
    expect(lowerFirst("30 day accounts")).toBe("30 day accounts");
    expect(lowerFirst("X")).toBe("X");
    expect(lowerFirst("")).toBe("");
  });

  it("keeps proper nouns and acronyms in the issue it shows", () => {
    const issue = (label: string) =>
      runResultParts(
        { ...run, summary: [{ label, value: 9, tone: "warning", view: null }] },
        definition,
      ).issue?.text;
    expect(issue("No Microsoft 365 app activity in 90 days")).toBe(
      "9 no Microsoft 365 app activity in 90 days",
    );
    expect(issue("MFA not registered")).toBe("9 MFA not registered");
    expect(issue("Inactive 30+ days")).toBe("9 inactive 30+ days");
  });
});

describe("audit catalog page", () => {
  const page = (props: Partial<Parameters<typeof AuditCatalogPage>[0]>) =>
    renderToStaticMarkup(
      createElement(AuditCatalogPage, {
        catalog: null,
        error: "",
        notice: "",
        busy: null,
        onOpen: () => {},
        onRun: () => {},
        schedule: null,
        now: NOW,
        ...props,
      }),
    );

  it("draws still rows while the list loads, and no loading text", () => {
    const html = page({});
    expect(html).toContain("skeleton__row");
    expect(html).toContain("Loading audits");
    expect(html).not.toContain("Loading…");
    expect(html).not.toContain("<table");
  });

  it("says why when the list could not be read, and draws nothing in its place", () => {
    const html = page({ error: "Request failed with HTTP 502." });
    expect(html).toContain('role="alert"');
    expect(html).toContain("Request failed with HTTP 502.");
    expect(html).not.toContain("skeleton");
    expect(html).not.toContain("<table");
  });

  it("draws the table and a note about a joined run once loaded", () => {
    const html = page({ catalog: table, notice: "Joined a run already in progress." });
    expect(html).toContain("<table");
    expect(html).toContain("Joined a run already in progress.");
  });
});

const settings = (
  saved: string[],
  domains: AuditSettings["domains"] = [
    { domain: "a.example", accounts: 12 },
    { domain: "b.example", accounts: 1 },
  ],
): AuditSettings => ({ scope: { domains: saved }, domains });

const renderScope = (value: AuditSettings) =>
  renderToStaticMarkup(
    createElement(AuditScopeSection, { settings: value, onSave: async () => {} }),
  );

describe("audit scope", () => {
  it("says all domains when nothing is chosen, with a button to change it", () => {
    const html = renderScope(settings([]));
    expect(html).toContain("Audit scope");
    expect(html).toContain("All domains");
    expect(html).toMatch(/<button[^>]*>Change<\/button>/);
    expect(html).not.toContain("Save");
    expect(html).not.toContain("Run Inactive accounts once");
  });

  it("names the saved domains", () => {
    const html = renderScope(settings(["a.example", "b.example"]));
    expect(html).toContain("2 domains: a.example, b.example");
    expect(html).not.toContain("All domains");
    expect(renderScope(settings(["a.example"]))).toContain("1 domain: a.example");
  });

  it("asks for a first run when no domains are known yet", () => {
    expect(renderScope(settings([], []))).toContain(
      "Run Inactive accounts once to list your domains.",
    );
  });

  it("lists every domain with its account count, and the way back to all of them", () => {
    const html = renderToStaticMarkup(
      createElement(AuditScopeEditor, {
        domains: [
          { domain: "a.example", accounts: 12 },
          { domain: "b.example", accounts: 1 },
          { domain: "old.example", accounts: null },
        ],
        selected: new Set(["b.example"]),
        saving: false,
        error: "",
        onToggle: () => {},
        onAll: () => {},
        onSave: () => {},
        onCancel: () => {},
      }),
    );
    expect(html).toContain("All domains");
    expect(html).toContain("a.example");
    expect(html).toContain("12 accounts");
    expect(html).toContain("1 account");
    expect(html).toContain("Save");
    expect(html).toContain("Cancel");
    // Only the chosen domain is ticked, so "All domains" is not.
    expect(html.match(/checked=""/g)).toHaveLength(1);
    expect(html).toMatch(/checked=""[^>]*\/><span>b\.example/);
  });

  it("ticks All domains when none is chosen, and shows a save error", () => {
    const html = renderToStaticMarkup(
      createElement(AuditScopeEditor, {
        domains: [{ domain: "a.example", accounts: 2 }],
        selected: new Set<string>(),
        saving: true,
        error: "Check the domains.",
        onToggle: () => {},
        onAll: () => {},
        onSave: () => {},
        onCancel: () => {},
      }),
    );
    expect(html.match(/checked=""/g)).toHaveLength(1);
    expect(html).toContain("Check the domains.");
    expect(html).toContain("Saving");
  });

  it("words the summary and the results-page label for one domain or several", () => {
    expect(scopeSummary([])).toBe("All domains");
    expect(scopeSummary(["a.example"])).toBe("1 domain: a.example");
    expect(scopeSummary(["a.example", "b.example"])).toBe("2 domains: a.example, b.example");
    expect(scopeLimit(1)).toBe("Limited to 1 domain");
    expect(scopeLimit(3)).toBe("Limited to 3 domains");
  });
});

describe("audit helpers", () => {
  it("builds an export link for exactly what the table shows", () => {
    const href = exportHref(run, {
      view: "d90",
      search: " sam ",
      sort: { column: "days", direction: "desc" },
      facets: { status: "Disabled", type: "" },
    });
    expect(href).toBe(
      "/api/v1/audit-runs/66666666-6666-4666-8666-666666666666/export.csv?view=d90&search=sam&sort=days&direction=desc&f.status=Disabled",
    );
  });
});

describe("address bar routes", () => {
  it("round-trips views, accounts and audits, and ignores other hashes", async () => {
    const { parseRoute, routeHash } = await import("../src/App.tsx");
    expect(parseRoute("#/audits/accounts.inactive")).toEqual({
      view: "audits",
      person: null,
      audit: "accounts.inactive",
    });
    expect(parseRoute("#/people/u%3A1")).toEqual({ view: "people", person: "u:1", audit: null });
    expect(parseRoute("#main-content")).toBeNull();
    expect(parseRoute("#/nowhere")).toBeNull();
    expect(parseRoute("#/people/%E0%A4%A")).toEqual({ view: "people", person: null, audit: null });
    expect(routeHash({ view: "people", person: "u:1", audit: null })).toBe("#/people/u%3A1");
    expect(routeHash({ view: "overview", person: "ignored", audit: null })).toBe("#/overview");
  });
});

describe("the scope a run used", () => {
  const scoped = { ...definition, scoped: true, version: 3 };
  const tags = (runScope: string[] | null, saved: string[] | null, version = 3) =>
    renderToStaticMarkup(
      createElement(RunScopeTags, {
        definition: scoped,
        run: { ...run, auditVersion: version, scope: runScope },
        saved,
      }),
    );

  it("names the scope the run used, not the one saved now", () => {
    const html = tags(["a.example"], ["a.example", "b.example"]);
    expect(html).toContain("Limited to 1 domain");
    expect(html).toContain("The saved scope has changed since this run");
    expect(tags(["a.example"], ["a.example"])).not.toContain("has changed");
  });

  it("says nothing it cannot know", () => {
    // The saved scope did not load, or the run was made by an older version that did not record it.
    expect(tags(null, null)).toBe("");
    expect(tags(null, ["a.example"], 2)).toBe("");
  });
});

const column = (
  key: string,
  kind: AuditColumn["kind"],
  overrides: Partial<AuditColumn> = {},
): AuditColumn => ({
  key,
  label: key,
  kind,
  facet: false,
  hidden: false,
  help: null,
  empty: null,
  ...overrides,
});

const memberRow = (key: string, name: string, principal: string | null): AuditRow => ({
  key,
  personId: `user-${key}`,
  tags: [],
  cells: { person: { name, principal } },
});

describe("rowEmails", () => {
  const people = [column("person", "person")];

  it("gives a member's user principal name, in the order of the rows", () => {
    const rows = [
      memberRow("2", "Blake Reed", "blake@example.invalid"),
      memberRow("1", "Avery Stone", "avery@example.invalid"),
    ];
    expect(rowEmails(people, rows)).toEqual(["blake@example.invalid", "avery@example.invalid"]);
  });

  it("prefers a guest's Email cell to its #EXT# principal name", () => {
    const columns = [column("person", "person"), column("email", "text")];
    const guest: AuditRow = {
      key: "g1",
      personId: "user-g1",
      tags: [],
      cells: {
        person: {
          name: "Casey Lin",
          principal: "casey.lin_partner.example#EXT#@tenant.example.invalid",
        },
        email: "casey.lin@partner.example",
      },
    };
    expect(rowEmails(columns, [guest])).toEqual(["casey.lin@partner.example"]);
  });

  it("reads the Email cell even when the person has no column to show it in", () => {
    // A hidden column is still part of the definition, so the address is not lost.
    const columns = [column("person", "person"), column("email", "text", { hidden: true })];
    const row: AuditRow = {
      ...memberRow("1", "Casey Lin", "casey#EXT#@tenant.example.invalid"),
      cells: {
        person: { name: "Casey Lin", principal: "casey#EXT#@tenant.example.invalid" },
        email: "casey@partner.example",
      },
    };
    expect(rowEmails(columns, [row])).toEqual(["casey@partner.example"]);
  });

  it("falls back to the principal name when the Email cell is empty or not an address", () => {
    const columns = [column("person", "person"), column("email", "text")];
    const rows = ["", null, "not an address", "two@example.invalid three@example.invalid"].map(
      (value, index) => ({
        ...memberRow(String(index), "Someone", `user${index}@example.invalid`),
        cells: {
          person: { name: "Someone", principal: `user${index}@example.invalid` },
          email: value,
        },
      }),
    );
    expect(rowEmails(columns, rows)).toEqual([
      "user0@example.invalid",
      "user1@example.invalid",
      "user2@example.invalid",
      "user3@example.invalid",
    ]);
  });

  it("drops duplicates, whatever their letter case, and rows without an address", () => {
    const rows = [
      memberRow("1", "Avery Stone", "avery@example.invalid"),
      memberRow("2", "No address", null),
      memberRow("3", "Blank address", "  "),
      memberRow("4", "Avery again", "Avery@Example.invalid"),
      memberRow("5", "Blake Reed", "blake@example.invalid"),
      { key: "6", personId: null, tags: [], cells: {} },
    ];
    expect(rowEmails(people, rows)).toEqual(["avery@example.invalid", "blake@example.invalid"]);
  });

  it("is empty when no row has an address", () => {
    expect(rowEmails(people, [])).toEqual([]);
    expect(rowEmails(people, [memberRow("1", "No address", null)])).toEqual([]);
  });
});

describe("audit result table copy buttons", () => {
  const columns = [
    column("person", "person", { label: "Guest" }),
    column("email", "text", { label: "Email" }),
    column("department", "text", { label: "Department" }),
  ];
  const guestDefinition = auditDefinitionSchema.parse({
    ...definition,
    id: "guests.review",
    columns,
  });
  const detail = (rows: AuditRow[], extra: Partial<typeof guestDefinition> = {}) =>
    renderToStaticMarkup(
      createElement(ResultTable, {
        definition: { ...guestDefinition, ...extra },
        run: { ...run, summary: [], rows } as AuditRunDetail,
        onOpenPerson: () => {},
      }),
    );
  const guest: AuditRow = {
    key: "g1",
    personId: "user-g1",
    tags: [],
    cells: {
      person: {
        name: "Casey Lin",
        principal: "casey.lin_partner.example#EXT#@tenant.example.invalid",
      },
      email: "casey.lin@partner.example",
      department: "Platform team",
    },
  };

  it("shows the account address outside the button that opens the person", () => {
    const html = detail([guest]);
    const opener = html.match(/<button class="people-name"[^>]*>(.*?)<\/button>/)?.[1] ?? "";
    expect(opener).toContain("Casey Lin");
    expect(opener).not.toContain("#EXT#");
    expect(html).toMatch(
      /<span class="copy-line__text">casey\.lin_partner\.example#EXT#@tenant\.example\.invalid<\/span>/,
    );
    expect(html).toContain(
      'aria-label="Copy casey.lin_partner.example#EXT#@tenant.example.invalid"',
    );
  });

  it("puts a copy button on an email-valued text cell and on no other text", () => {
    const html = detail([guest]);
    expect(html).toMatch(
      /<span class="copy-line__text">casey\.lin@partner\.example<\/span><button[^>]*aria-label="Copy casey\.lin@partner\.example"/,
    );
    expect(html).toContain("Platform team");
    expect(html).not.toContain("Copy Platform team");
    // The account's address and the Email cell: two copy buttons for the one row.
    expect(html.match(/aria-label="Copy /g)).toHaveLength(2);
  });

  it("offers Copy emails, before Export CSV, once there is an address to copy", () => {
    const html = detail([guest, { ...guest, key: "g2", personId: "user-g2" }]);
    expect(html).toMatch(
      /<button class="button button--secondary" type="button"><svg[^>]*>.*?<\/svg> Copy emails<\/button>/,
    );
    expect(html.indexOf("Copy emails")).toBeLessThan(html.indexOf("Export CSV"));
    expect(html).not.toContain("Copied");
    expect(html).not.toContain("Copy failed");
  });

  it("disables Copy emails when no row has an address", () => {
    const html = detail([
      {
        key: "n1",
        personId: "user-n1",
        tags: [],
        cells: {
          person: { name: "No address", principal: null },
        },
      },
    ]);
    expect(html).toMatch(/<button[^>]*disabled=""[^>]*><svg[^>]*>.*?<\/svg> Copy emails<\/button>/);
  });

  it("has no Copy emails button, and no copy button, in a table without a person column", () => {
    const html = renderToStaticMarkup(
      createElement(ResultTable, {
        definition: auditDefinitionSchema.parse({
          ...definition,
          columns: [column("email", "text"), column("product", "text")],
          defaultSort: { column: "product", direction: "asc" },
        }),
        run: {
          ...run,
          summary: [],
          rows: [
            {
              key: "1",
              personId: null,
              tags: [],
              cells: { email: "ops@example.invalid", product: "Plan" },
            },
          ],
        } as AuditRunDetail,
        onOpenPerson: () => {},
      }),
    );
    expect(html).not.toContain("Copy emails");
    // An email-valued cell still gets its own button.
    expect(html).toContain('aria-label="Copy ops@example.invalid"');
  });
});

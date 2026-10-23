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
  AuditScopeEditor,
  AuditScopeSection,
  ago,
  exportHref,
  ResultTable,
  RunScopeTags,
  rowEmails,
  runHeadline,
  scopeLimit,
  scopeSummary,
} from "../src/Audits.tsx";

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
    }),
  );

describe("audit catalog", () => {
  it("offers a run for an available audit that has never run", () => {
    const html = render(catalog());
    expect(html).toContain("Inactive accounts");
    expect(html).toContain("Not run yet.");
    expect(html).toContain("Run audit");
    expect(html).not.toContain("disabled");
  });

  it("summarises the last run by its first notable number", () => {
    expect(runHeadline(run, definition)).toBe(
      "Ran 5 min ago · 214 accounts · 47 inactive 30+ days",
    );
    expect(runHeadline({ ...run, complete: false }, definition)).toContain("partial");
    expect(
      runHeadline(
        { ...run, status: "failed", error: { code: "forbidden", message: "Microsoft refused." } },
        definition,
      ),
    ).toContain("Microsoft refused.");
  });

  it("links to consent when the session lacks a permission, and disables the run", () => {
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
    expect(html).toMatch(/<button class="button button--primary" type="button" disabled="">/);
  });

  it("cannot run anything in demo mode", () => {
    const html = render(catalog({}, "demo"));
    expect(html).toContain("Demo mode lists the audits but cannot run them");
    expect(html).toMatch(/disabled=""/);
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

  it("says how long ago something happened", () => {
    const now = Date.parse("2026-10-22T12:00:00Z");
    expect(ago("2026-10-22T11:59:30Z", now)).toBe("just now");
    expect(ago("2026-10-22T11:15:00Z", now)).toBe("45 min ago");
    expect(ago("2026-10-22T07:00:00Z", now)).toBe("5 h ago");
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

import {
  type AuditCatalog,
  type AuditRun,
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
  RunScopeTags,
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

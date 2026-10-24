import { describe, expect, it } from "vitest";
import {
  type AuditRow,
  auditCsv,
  auditDefinitionSchema,
  auditFacetOptions,
  auditScopeSchema,
  auditSettingsSchema,
  queryAuditRows,
} from "../src/index.ts";

const definition = auditDefinitionSchema.parse({
  id: "accounts.example",
  version: 1,
  title: "Example",
  description: "Synthetic.",
  category: "Accounts",
  subject: { one: "account", many: "accounts" },
  columns: [
    { key: "person", label: "Account", kind: "person" },
    { key: "days", label: "Inactive for", kind: "days" },
    { key: "status", label: "Status", kind: "enum", facet: true },
    { key: "last", label: "Last sign-in", kind: "date" },
    { key: "licenses", label: "Licenses", kind: "list" },
  ],
  views: [
    { id: "all", label: "All", tags: [] },
    { id: "d30", label: "30+ days", tags: ["d30"] },
    { id: "d30-enabled", label: "30+ days, enabled", tags: ["d30", "enabled"] },
  ],
  defaultView: "all",
  defaultSort: { column: "days", direction: "desc" },
  permissions: ["User.Read.All"],
  optionalPermissions: [],
  requirements: [],
});

const rows: AuditRow[] = [
  {
    key: "a",
    personId: "a",
    tags: ["d30", "enabled"],
    cells: {
      person: { name: "Avery", principal: "avery@example.invalid" },
      days: 40,
      status: "Enabled",
      last: "2026-09-12T00:00:00Z",
      licenses: ["E3", "Visio"],
    },
  },
  {
    key: "b",
    personId: "b",
    tags: ["d30", "disabled"],
    cells: {
      person: { name: '=cmd|"/c calc"!A1', principal: "bo@example.invalid" },
      days: 400,
      status: "Disabled",
      last: null,
      licenses: [],
    },
  },
  {
    key: "c",
    personId: "c",
    tags: ["enabled"],
    cells: {
      person: { name: "Cai", principal: "cai@example.invalid" },
      days: null,
      status: "Enabled",
      last: "2026-10-20T00:00:00Z",
      licenses: [],
    },
  },
];

const keys = (list: AuditRow[]) => list.map((row) => row.key);

describe("audit table queries", () => {
  it("filters by every tag of the view, facets and search", () => {
    expect(keys(queryAuditRows(definition, rows, { view: "d30" }))).toEqual(["b", "a"]);
    expect(keys(queryAuditRows(definition, rows, { view: "d30-enabled" }))).toEqual(["a"]);
    expect(keys(queryAuditRows(definition, rows, { facets: { status: "Enabled" } }))).toEqual([
      "a",
      "c",
    ]);
    expect(keys(queryAuditRows(definition, rows, { search: "CAI@" }))).toEqual(["c"]);
    expect(keys(queryAuditRows(definition, rows, { search: "visio" }))).toEqual(["a"]);
    // A filter on a column that is not a facet is ignored rather than hiding everything.
    expect(queryAuditRows(definition, rows, { facets: { days: "40" } })).toHaveLength(3);
  });

  it("sorts with empty values last in both directions", () => {
    expect(keys(queryAuditRows(definition, rows))).toEqual(["b", "a", "c"]);
    expect(
      keys(queryAuditRows(definition, rows, { sort: { column: "days", direction: "asc" } })),
    ).toEqual(["a", "b", "c"]);
    expect(
      keys(queryAuditRows(definition, rows, { sort: { column: "last", direction: "desc" } })),
    ).toEqual(["c", "a", "b"]);
    expect(
      keys(queryAuditRows(definition, rows, { sort: { column: "person", direction: "asc" } })),
    ).toEqual(["b", "a", "c"]);
  });

  it("offers each filter's values", () => {
    expect(auditFacetOptions(definition, rows)).toEqual({ status: ["Disabled", "Enabled"] });
  });
});

describe("audit CSV formula guard", () => {
  it("keeps every formula-leading character as text", () => {
    const one = (name: string) =>
      auditCsv(definition, [
        { key: "x", personId: null, tags: [], cells: { person: { name, principal: null } } },
      ])
        .slice(1)
        .split("\r\n")[1];
    for (const lead of ["=", "+", "-", "@", "\t", "\r"])
      expect(one(`${lead}SUM(1)`)?.startsWith(`"'${lead}`), JSON.stringify(lead)).toBe(true);
    expect(one("Plain name")).toBe('"Plain name",,,,,');
  });
});

describe("audit CSV", () => {
  it("quotes text, splits the person, and keeps formulas as text", () => {
    const csv = auditCsv(definition, rows);
    expect(csv.startsWith("﻿")).toBe(true);
    const lines = csv.slice(1).trimEnd().split("\r\n");
    expect(lines[0]).toBe(
      '"Account","User principal name","Inactive for","Status","Last sign-in","Licenses"',
    );
    expect(lines[1]).toBe(
      '"Avery","avery@example.invalid",40,"Enabled","2026-09-12T00:00:00Z","E3; Visio"',
    );
    expect(lines[2]).toBe(`"'=cmd|""/c calc""!A1","bo@example.invalid",400,"Disabled",,""`);
    expect(lines[3]).toContain('"Cai","cai@example.invalid",,"Enabled"');
  });
});

describe("audit scope", () => {
  it("publishes whether an audit uses the scope, off unless it says so", () => {
    expect(definition.scoped).toBe(false);
    expect(auditDefinitionSchema.parse({ ...definition, scoped: true }).scoped).toBe(true);
  });

  it("keeps lower-case domain names once each, in order", () => {
    expect(auditScopeSchema.parse({ domains: [] })).toEqual({ domains: [] });
    expect(
      auditScopeSchema.parse({
        domains: [" B.Example ", "a.example", "A.EXAMPLE", "sub.a.example"],
      }),
    ).toEqual({ domains: ["a.example", "b.example", "sub.a.example"] });
  });

  it("refuses anything that is not a list of domain names", () => {
    const bad: unknown[] = [
      undefined,
      null,
      {},
      { domains: "a.example" },
      { domains: [1] },
      { domains: [""] },
      { domains: ["not a domain"] },
      { domains: ["ada@a.example"] },
      { domains: ["a..example"] },
      { domains: ["-a.example"] },
      { domains: ["a.example."] },
      { domains: [`${"a".repeat(64)}.example`] },
      { domains: Array.from({ length: 101 }, (_, index) => `d${index}.example`) },
    ];
    for (const body of bad)
      expect(auditScopeSchema.safeParse(body).success, String(body)).toBe(false);
    const hundred = Array.from({ length: 100 }, (_, index) => `d${index}.example`);
    expect(auditScopeSchema.parse({ domains: hundred }).domains).toHaveLength(100);
  });

  it("describes the settings: the saved scope and the domains an administrator can choose", () => {
    expect(
      auditSettingsSchema.parse({
        scope: { domains: ["a.example"] },
        domains: [{ domain: "a.example", accounts: 12 }],
      }),
    ).toEqual({
      scope: { domains: ["a.example"] },
      domains: [{ domain: "a.example", accounts: 12 }],
    });
    expect(
      auditSettingsSchema.safeParse({ scope: { domains: [] }, domains: [{ domain: "a.example" }] })
        .success,
    ).toBe(false);
  });
});

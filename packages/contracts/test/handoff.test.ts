import { readFile } from "node:fs/promises";
import { describe, expect, it } from "vitest";
import {
  type AuditRow,
  actionLabel,
  auditDefinitionSchema,
  buildHandoff,
  canonicalJson,
  handoffFileName,
  handoffItems,
  handoffSchema,
  handoffSelection,
  MAX_HANDOFF_ITEMS,
  MAX_RECEIPT_ITEMS,
  queryAuditRows,
  receiptProblem,
  receiptSchema,
} from "../src/index.ts";

// Synthetic accounts and tools only.
const definition = auditDefinitionSchema.parse({
  id: "accounts.example",
  version: 3,
  title: "Example accounts",
  description: "Synthetic.",
  category: "Accounts",
  subject: { one: "account", many: "accounts" },
  columns: [
    { key: "person", label: "Account", kind: "person" },
    { key: "days", label: "Inactive for", kind: "days" },
    { key: "status", label: "Status", kind: "enum", facet: true },
    { key: "licenses", label: "Licenses", kind: "list" },
    { key: "admin", label: "Admin", kind: "boolean" },
    { key: "department", label: "Department", kind: "text", hidden: true },
    { key: "jobTitle", label: "Status", kind: "text", hidden: true },
  ],
  views: [
    { id: "all", label: "All accounts", tags: [] },
    { id: "d30", label: "30+ days", tags: ["d30"] },
    { id: "d90", label: "90+ days", tags: ["d90"] },
    { id: "d90-admin", label: "Admins 90+ days", tags: ["d90", "admin"] },
    { id: "never", label: "Never signed in", tags: ["never"] },
  ],
  defaultView: "d30",
  defaultSort: { column: "days", direction: "desc" },
  permissions: ["User.Read.All"],
  optionalPermissions: [],
  requirements: [],
});

const row = (id: string | null, days: number, tags: string[], extra = {}): AuditRow => ({
  key: id ?? `row-${days}`,
  personId: id,
  tags,
  cells: {
    person: { name: `Person ${days}`, principal: `person${days}@example.invalid` },
    days,
    status: "Enabled",
    licenses: ["Example E3", "Example Visio"],
    admin: false,
    department: "Synthetic",
    jobTitle: "Tester",
    ...extra,
  },
});

const rows: AuditRow[] = [
  row("u-1", 10, ["enabled"]),
  row("u-2", 45, ["d30", "enabled"]),
  row("u-3", 120, ["d30", "d90", "admin", "disabled"], { status: "Disabled", admin: true }),
  row(null, 200, ["d30", "d90"]),
];

const receipt = (overrides: Record<string, unknown> = {}, item: Record<string, unknown> = {}) => ({
  format: "prism365.receipt",
  version: 1,
  id: "7b0c3a52-5c51-4c8e-9a43-1f2d3e4f5a6b",
  handoffId: null,
  tool: { name: "Example action tool", version: "2.1.0" },
  performedAt: "2026-10-22T14:30:00Z",
  performedBy: "Operator One",
  items: [
    {
      userId: "0f1e2d3c-4b5a-4968-8776-5a4b3c2d1e0f",
      userPrincipalName: "avery@example.invalid",
      action: "account.sign-in-blocked",
      status: "done",
      detail: "",
      ...item,
    },
  ],
  ...overrides,
});

const handoff = () =>
  buildHandoff({
    id: "9a8b7c6d-5e4f-4a3b-8c2d-1e0f9a8b7c6d",
    generatedAt: "2026-10-22T15:00:00.000Z",
    tenantId: "11111111-1111-4111-8111-111111111111",
    generatedBy: { actorId: "22222222-2222-4222-8222-222222222222", name: "Synthetic admin" },
    definition,
    run: {
      id: "66666666-6666-4666-8666-666666666666",
      auditVersion: 3,
      startedAt: "2026-10-21T09:00:00.000Z",
      finishedAt: "2026-10-21T09:01:00.000Z",
    },
    query: { view: "d30" },
    rows: queryAuditRows(definition, rows, { view: "d30" }),
    columns: null,
  });

describe("handoff v1", () => {
  it("builds a valid document from the rows of a view", () => {
    const document = handoff();
    expect(handoffSchema.safeParse(document).success).toBe(true);
    expect(document.source).toEqual({
      kind: "audit",
      auditId: "accounts.example",
      auditTitle: "Example accounts",
      auditVersion: 3,
      runId: "66666666-6666-4666-8666-666666666666",
      runFinishedAt: "2026-10-21T09:01:00.000Z",
      view: { id: "d30", label: "30+ days" },
      search: "",
      facets: {},
      sort: { column: "days", direction: "desc" },
    });
    // The row with no account is left out; the rest keep the view's order.
    expect(document.items.map((item) => item.userId)).toEqual(["u-3", "u-2"]);
    expect(document.items[0]).toEqual({
      userId: "u-3",
      userPrincipalName: "person120@example.invalid",
      displayName: "Person 120",
      // Single-tag views by label, in the audit's order; "disabled" names no view.
      reasons: ["30+ days", "90+ days"],
      evidence: {
        "Inactive for": 120,
        Status: "Disabled",
        Licenses: "Example E3; Example Visio",
        Admin: true,
      },
    });
  });

  it("takes the columns the page shows, keeping two columns with one label apart", () => {
    const [item] = handoffItems(
      definition,
      [rows[2] as AuditRow],
      ["person", "jobTitle", "status"],
    );
    expect(item?.evidence).toEqual({ Status: "Disabled", "Status (jobTitle)": "Tester" });
  });

  it("records what the export applied, as queryAuditRows applies it", () => {
    expect(
      handoffSelection(definition, {
        view: "missing",
        facets: { status: "Disabled", days: "1", department: "" },
        search: "  person ",
        sort: { column: "nope", direction: "asc" },
      }),
    ).toEqual({ view: null, search: "person", facets: { status: "Disabled" }, sort: null });
    expect(handoffSelection(definition, {}).view).toEqual({ id: "d30", label: "30+ days" });
  });

  it("cleans and bounds stored text on the way out", () => {
    const [item] = handoffItems(
      definition,
      [row("u-9", 40, ["d30"], { person: { name: "A\u0000B\uD800", principal: null } })],
      null,
    );
    expect(item?.displayName).toBe("AB�");
    expect(item?.userPrincipalName).toBeNull();
    const [long] = handoffItems(
      definition,
      [row("u-8", 40, [], { status: "x".repeat(3_000) })],
      ["status"],
    );
    expect(long?.evidence.Status).toHaveLength(2_000);
  });

  it("rejects unknown keys at every level and each bound", () => {
    const document = handoff();
    const item = document.items[0];
    expect(item).toBeDefined();
    const invalid = [
      { ...document, extra: true },
      { ...document, source: { ...document.source, extra: true } },
      { ...document, items: [{ ...item, extra: true }] },
      { ...document, format: "prism365.receipt" },
      { ...document, version: 2 },
      { ...document, items: Array.from({ length: MAX_HANDOFF_ITEMS + 1 }, () => item) },
      { ...document, items: [{ ...item, reasons: Array.from({ length: 21 }, () => "Reason") }] },
      {
        ...document,
        items: [
          {
            ...item,
            evidence: Object.fromEntries(Array.from({ length: 41 }, (_, i) => [`Column ${i}`, i])),
          },
        ],
      },
      { ...document, items: [{ ...item, evidence: { Status: "x".repeat(2_001) } }] },
      { ...document, items: [{ ...item, evidence: { Status: { nested: true } } }] },
      { ...document, items: [{ ...item, userId: "not an id" }] },
      { ...document, source: { ...document.source, view: { id: "Bad View", label: "x" } } },
      {
        ...document,
        source: {
          ...document.source,
          facets: Object.fromEntries(Array.from({ length: 31 }, (_, i) => [`c${i}`, "x"])),
        },
      },
    ];
    for (const value of invalid) expect(handoffSchema.safeParse(value).success).toBe(false);
  });

  it("names the file after the audit and the day of its run", () => {
    expect(handoffFileName("accounts.inactive", "2026-10-21T09:00:00.000Z")).toBe(
      "prism365-handoff-accounts.inactive-2026-10-21.json",
    );
  });
});

describe("receipt v1", () => {
  it("parses a valid receipt, with ids and GUID user ids in lower case", () => {
    const parsed = receiptSchema.parse(
      receipt(
        { id: "7B0C3A52-5C51-4C8E-9A43-1F2D3E4F5A6B" },
        { userId: "0F1E2D3C-4B5A-4968-8776-5A4B3C2D1E0F" },
      ),
    );
    expect(parsed.id).toBe("7b0c3a52-5c51-4c8e-9a43-1f2d3e4f5a6b");
    expect(parsed.items[0]?.userId).toBe("0f1e2d3c-4b5a-4968-8776-5a4b3c2d1e0f");
    // A non-GUID directory id is kept as it was sent.
    expect(receiptSchema.parse(receipt({}, { userId: "Legacy-ID.1" })).items[0]?.userId).toBe(
      "Legacy-ID.1",
    );
  });

  it("needs a user id or a principal name on every item", () => {
    expect(receiptSchema.safeParse(receipt({}, { userPrincipalName: null })).success).toBe(true);
    expect(receiptSchema.safeParse(receipt({}, { userId: null })).success).toBe(true);
    const neither = receipt({}, { userId: null, userPrincipalName: null });
    const result = receiptSchema.safeParse(neither);
    expect(result.success).toBe(false);
    if (!result.success)
      expect(receiptProblem(neither, result.error)).toBe(
        "items[0] needs a userId or a userPrincipalName.",
      );
    // A principal name of only spaces is no principal name.
    expect(
      receiptSchema.safeParse(receipt({}, { userId: null, userPrincipalName: "  " })).success,
    ).toBe(false);
  });

  it("rejects unknown keys at the top level and inside", () => {
    for (const value of [
      { ...receipt(), extra: 1 },
      receipt({ tool: { name: "Example", version: "1", vendor: "x" } }),
      receipt({}, { ticket: "x" }),
    ])
      expect(receiptSchema.safeParse(value).success).toBe(false);
  });

  it("rejects each bound", () => {
    const longText = (length: number) => "x".repeat(length);
    const item = receipt().items[0];
    const cases: [unknown, boolean][] = [
      [receipt({ tool: { name: longText(80), version: "1" } }), true],
      [receipt({ tool: { name: longText(81), version: "1" } }), false],
      [receipt({ tool: { name: "", version: "1" } }), false],
      [receipt({ tool: { name: "Example", version: longText(40) } }), true],
      [receipt({ tool: { name: "Example", version: longText(41) } }), false],
      [receipt({ performedBy: longText(200) }), true],
      [receipt({ performedBy: longText(201) }), false],
      [receipt({ performedBy: "" }), false],
      [receipt({}, { detail: longText(500) }), true],
      [receipt({}, { detail: longText(501) }), false],
      [receipt({}, { userPrincipalName: longText(500) }), true],
      [receipt({}, { userPrincipalName: longText(501) }), false],
      [receipt({}, { userId: longText(201) }), false],
      [receipt({}, { userId: "has space" }), false],
      [receipt({}, { status: "maybe" }), false],
      [receipt({ items: [] }), false],
      [receipt({ items: Array.from({ length: MAX_RECEIPT_ITEMS }, () => item) }), true],
      [receipt({ items: Array.from({ length: MAX_RECEIPT_ITEMS + 1 }, () => item) }), false],
      [receipt({ id: "not-a-guid" }), false],
      [receipt({ handoffId: "not-a-guid" }), false],
      [receipt({ performedAt: "2026-10-22" }), false],
      [receipt({ performedAt: "2026-10-22T14:30:00" }), false],
      // Dates PostgreSQL refuses, or stores past year 9999, never reach it.
      [receipt({ performedAt: "0000-01-01T00:00:00Z" }), false],
      [receipt({ performedAt: "9999-12-31T23:59:59-14:00" }), false],
      [receipt({ performedAt: "1999-12-31T23:59:59Z" }), false],
      [receipt({ performedAt: "2000-01-01T00:00:00Z" }), true],
      [receipt({ performedAt: "2099-12-31T23:59:59+02:00" }), true],
      [receipt({ format: "prism365.handoff" }), false],
      [receipt({ version: 2 }), false],
    ];
    for (const [value, valid] of cases)
      expect(receiptSchema.safeParse(value).success, JSON.stringify(value).slice(0, 160)).toBe(
        valid,
      );
  });

  it("allows only lower-case action slugs of 2 to 80 characters", () => {
    for (const action of ["ab", "account.sign-in-blocked", "license.removed", `a${"b".repeat(79)}`])
      expect(receiptSchema.safeParse(receipt({}, { action })).success, action).toBe(true);
    for (const action of [
      "a",
      `a${"b".repeat(80)}`,
      "Account.disabled",
      "1account",
      ".account",
      "account_disabled",
      "account disabled",
      "account/disabled",
    ])
      expect(receiptSchema.safeParse(receipt({}, { action })).success, action).toBe(false);
  });

  it("cleans every string the tool sent", () => {
    const parsed = receiptSchema.parse(
      receipt(
        {
          tool: { name: " Example\u0000 tool ", version: "1\u0000.0" },
          performedBy: "Op\uDC00erator",
        },
        { userPrincipalName: "avery\u0000@example.invalid", detail: "Blocked\u0000 at\uD800" },
      ),
    );
    expect(parsed.tool).toEqual({ name: "Example tool", version: "1.0" });
    expect(parsed.performedBy).toBe("Op�erator");
    expect(parsed.items[0]?.userPrincipalName).toBe("avery@example.invalid");
    expect(parsed.items[0]?.detail).toBe("Blocked at�");
  });

  it("says what is wrong in plain words", () => {
    const problem = (value: unknown) => {
      const result = receiptSchema.safeParse(value);
      if (result.success) throw new Error("expected a problem");
      return receiptProblem(value, result.error);
    };
    expect(problem([])).toBe(
      "The receipt is not a JSON object. Choose a receipt file written by the action tool.",
    );
    expect(problem(receipt({ format: "other" }))).toBe(
      'This is not a Prism365 receipt: format should be "prism365.receipt".',
    );
    expect(problem(receipt({ version: 2 }))).toBe(
      "This receipt is version 2. This Prism365 reads version 1.",
    );
    expect(problem({ ...receipt(), extra: 1 })).toBe(
      'The receipt has a field Prism365 does not read: "extra". Check the receipt\'s format and version.',
    );
    // What the tool sent is echoed only short.
    expect(problem({ ...receipt(), ["x".repeat(10_000)]: 1 }).length).toBeLessThan(200);
    expect(problem(receipt({ version: "x".repeat(10_000) })).length).toBeLessThan(200);
    const { performedBy: _performedBy, ...missing } = receipt();
    expect(problem(missing)).toBe("performedBy is missing.");
    expect(problem(receipt({}, { action: "Bad" }))).toMatch(
      /^items\[0\]\.action is not an action name/,
    );
    expect(problem(receipt({}, { detail: "x".repeat(501) }))).toBe(
      "items[0].detail is longer than 500 characters.",
    );
    expect(problem(receipt({}, { status: "maybe" }))).toBe(
      "items[0].status should be done, failed or skipped.",
    );
    expect(problem(receipt({ items: [] }))).toBe(
      "items is empty. A receipt reports at least one action.",
    );
    expect(problem(receipt({ performedAt: "yesterday" }))).toBe(
      "performedAt is not a date and time with a time zone, such as 2026-10-22T14:30:00Z.",
    );
    expect(problem(receipt({ handoffId: 7 }))).toBe("handoffId should be a string or null.");
    expect(problem(receipt({ performedAt: "1970-01-01T00:00:00Z" }))).toBe(
      "performedAt is not between the years 2000 and 2099.",
    );
  });

  it("digests the same receipt the same way whatever its key order", () => {
    const one = receiptSchema.parse(receipt());
    const reordered = receiptSchema.parse(
      Object.fromEntries(Object.entries(receipt()).reverse()) as unknown,
    );
    expect(canonicalJson(reordered)).toBe(canonicalJson(one));
    expect(canonicalJson(receiptSchema.parse(receipt({}, { detail: "changed" })))).not.toBe(
      canonicalJson(one),
    );
    expect(canonicalJson({ b: [2, { d: null, c: "x" }], a: true })).toBe(
      '{"a":true,"b":[2,{"c":"x","d":null}]}',
    );
  });

  it("turns an action slug into a readable label", () => {
    expect(actionLabel("account.sign-in-blocked")).toBe("Account sign in blocked");
    expect(actionLabel("license.removed")).toBe("License removed");
  });
});

describe("docs/handoff.md", () => {
  it("gives a complete example of each format that parses", async () => {
    const text = await readFile(new URL("../../../docs/handoff.md", import.meta.url), "utf8");
    const examples = [...text.matchAll(/```json\n([\s\S]*?)\n```/g)].map((match) =>
      JSON.parse(match[1] ?? ""),
    );
    expect(examples).toHaveLength(2);
    const [handoffExample, receiptExample] = examples;
    expect(handoffSchema.parse(handoffExample).items.length).toBeGreaterThan(0);
    expect(receiptSchema.parse(receiptExample).handoffId).toBe(handoffExample.id);
  });
});

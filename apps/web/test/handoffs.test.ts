import {
  type AuditRow,
  type AuditRunDetail,
  auditDefinitionSchema,
  type HandoffList,
  type PersonReceiptList,
  type PrismSession,
  type ReceiptImport,
  type ReceiptItemPage,
} from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { exportHref, handoffHref, handoffNote, ResultTable } from "../src/Audits.tsx";
import {
  answers,
  HandoffsView,
  type ImportState,
  importMessage,
  ReceiptItemsTable,
  readReceiptFile,
} from "../src/Handoffs.tsx";
import { PersonReceipts, PersonReceiptsPanel } from "../src/PersonReceipts.tsx";

// Invented accounts, tools and operators only.
const handoffId = "9a8b7c6d-5e4f-4a3b-8c2d-1e0f9a8b7c6d";
const receiptSummary = {
  id: "7b0c3a52-5c51-4c8e-9a43-1f2d3e4f5a6b",
  handoff: {
    id: handoffId,
    auditId: "accounts.inactive",
    generatedAt: "2026-10-21T10:00:00.000Z",
    items: 12,
  },
  tool: { name: "Example action tool", version: "1.4.0" },
  performedAt: "2026-10-21T15:00:00.000Z",
  performedBy: "Jordan Vale",
  importedAt: "2026-10-21T16:00:00.000Z",
  items: 6,
  done: 3,
  failed: 2,
  skipped: 1,
};
const list: HandoffList = {
  handoffs: [
    {
      id: handoffId,
      auditId: "accounts.inactive",
      auditTitle: "Inactive accounts",
      runId: "66666666-6666-4666-8666-666666666666",
      view: { id: "d90", label: "90+ days" },
      filtered: true,
      items: 12,
      generatedAt: "2026-10-21T10:00:00.000Z",
      generatedBy: { id: "22222222-2222-4222-8222-222222222222", name: "Synthetic admin" },
    },
  ],
  receipts: [receiptSummary],
};

const view = (
  value: HandoffList | null,
  importState: ImportState = { state: "idle" },
  error = "",
) =>
  renderToStaticMarkup(
    createElement(HandoffsView, {
      list: value,
      error,
      importState,
      selected: null,
      onImport: () => {},
      onSelect: () => {},
    }),
  );

describe("Handoffs page", () => {
  it("explains both formats and that receipts are the tool's reports", () => {
    const html = view(list);
    expect(html).toContain("Prism365 only reads.");
    expect(html).toContain("does not check them against the tenant");
    expect(html).toContain('href="https://github.com/Nullel/prism365/blob/main/docs/handoff.md"');
    expect(html).toContain('accept=".json,application/json"');
  });

  it("shows loading, then empty lists with what to do next", () => {
    expect(view(null)).toContain("Loading handoffs…");
    const empty = view({ handoffs: [], receipts: [] });
    expect(empty).toContain("Nothing exported yet. Open an audit");
    expect(empty).toContain("No receipts yet.");
    const failed = view(null, { state: "idle" }, "Handoffs could not be loaded.");
    expect(failed).toContain("Handoffs could not be loaded.");
    expect(failed).not.toContain("Loading handoffs…");
  });

  it("lists exports and receipts with their counts and the handoff each answers", () => {
    const html = view(list);
    expect(html).toContain("Inactive accounts");
    expect(html).toContain("90+ days, filtered");
    expect(html).toContain(">12<");
    expect(html).toContain("Synthetic admin");
    expect(html).toContain("Example action tool");
    expect(html).toContain("1.4.0");
    expect(html).toContain("by Jordan Vale");
    expect(html).toMatch(/>3<\/td><td class="num">2<\/td><td class="num">1</);
    expect(html).toContain("Inactive accounts, exported");
    expect(html).toContain("Show items");
    expect(answers({ ...receiptSummary, handoff: null }, list.handoffs)).toBe("No handoff");
  });

  it("says what an import did, or why it failed", () => {
    const result: ReceiptImport = { receipt: receiptSummary, imported: true };
    expect(view(list, { state: "reading", name: "receipt.json" })).toContain(
      "Importing receipt.json…",
    );
    expect(view(list, { state: "done", name: "receipt.json", result })).toContain(
      "Imported the receipt from Example action tool: 3 done, 2 failed, 1 skipped.",
    );
    expect(importMessage({ ...result, imported: false })).toMatch(
      /^This receipt was already imported on .+\. Nothing changed\.$/,
    );
    const failed = view(list, {
      state: "failed",
      name: "receipt.json",
      message: "items[0].status should be done, failed or skipped.",
    });
    expect(failed).toContain('role="alert"');
    expect(failed).toContain("receipt.json: items[0].status should be done, failed or skipped.");
  });

  it("shows a receipt's items as the tool's report", () => {
    const page: ReceiptItemPage = {
      receipt: receiptSummary,
      items: [
        {
          position: 0,
          userId: "user-1",
          userPrincipalName: "avery@example.invalid",
          action: "account.sign-in-blocked",
          status: "done",
          detail: "",
        },
        {
          position: 1,
          userId: null,
          userPrincipalName: "blake@example.invalid",
          action: "license.removed",
          status: "failed",
          detail: "<b>not markup</b>",
        },
      ],
      next: 1,
    };
    const html = renderToStaticMarkup(
      createElement(ReceiptItemsTable, { page, loading: false, onMore: () => {} }),
    );
    expect(html).toContain("As reported by Example action tool 1.4.0, run by Jordan Vale");
    expect(html).toContain("Prism365 has not checked these against the tenant.");
    expect(html).toContain("Account sign in blocked");
    expect(html).toContain('class="person-tag person-tag--ok">done');
    expect(html).toContain('class="person-tag person-tag--alert">failed');
    expect(html).toContain("&lt;b&gt;not markup&lt;/b&gt;");
    expect(html).toContain("Show more");
    expect(html).toContain("Showing 2 of 6");
  });

  it("reads a receipt file only when it is JSON and at most 4 MB", async () => {
    const file = (text: string, size = text.length) => ({ size, text: async () => text });
    await expect(readReceiptFile(file('{"format":"prism365.receipt"}'))).resolves.toEqual({
      format: "prism365.receipt",
    });
    await expect(readReceiptFile(file("not json"))).rejects.toThrow("This file is not JSON.");
    await expect(readReceiptFile(file("{}", 4 * 1024 * 1024 + 1))).rejects.toThrow("at most 4 MB");
  });
});

describe("Reported by other tools", () => {
  const items: PersonReceiptList["items"] = [
    {
      receiptId: receiptSummary.id,
      position: 0,
      userId: "user-1",
      userPrincipalName: null,
      action: "account.sign-in-blocked",
      status: "done",
      detail: "Blocked after review.",
      tool: { name: "Example action tool", version: "1.4.0" },
      performedAt: "2026-10-21T15:00:00.000Z",
      performedBy: "Jordan Vale",
      importedAt: "2026-10-21T16:00:00.000Z",
      matchedBy: "userId",
    },
    {
      receiptId: receiptSummary.id,
      position: 1,
      userId: null,
      userPrincipalName: "avery@example.invalid",
      action: "license.removed",
      status: "skipped",
      detail: "",
      tool: { name: "Example action tool", version: "1.4.0" },
      performedAt: "2026-10-21T15:00:00.000Z",
      performedBy: "Jordan Vale",
      importedAt: "2026-10-21T16:00:00.000Z",
      matchedBy: "userPrincipalName",
    },
  ];
  const panel = (value: PersonReceiptList | null, error = "") =>
    renderToStaticMarkup(createElement(PersonReceiptsPanel, { list: value, error }));

  it("shows nothing when no receipt names the person", () => {
    expect(panel({ items: [], truncated: false })).toBe("");
    expect(panel(null)).toBe("");
  });

  it("lists each reported action with tool, status, when and detail, as a report", () => {
    const html = panel({ items, truncated: true });
    expect(html).toContain("Reported by other tools");
    expect(html).toContain("These are reports from the tools that acted");
    expect(html).toContain("Prism365 has not checked them against the tenant.");
    expect(html).toContain("Account sign in blocked");
    expect(html).toContain("License removed");
    expect(html).toContain("Blocked after review.");
    expect(html).toContain("run by Jordan Vale");
    expect(html).toContain('class="person-tag">skipped');
    // Only the item matched by name says so.
    expect(html.match(/Matched by sign-in name/g)).toHaveLength(1);
    expect(html).toContain("Showing the newest 100.");
  });

  it("says when the reports could not be loaded", () => {
    expect(panel(null, "Reports from other tools could not be loaded.")).toContain(
      "Reports from other tools could not be loaded.",
    );
  });

  it("loads nothing without a signed-in live session", () => {
    const session = { mode: "demo", authenticated: false } as PrismSession;
    expect(renderToStaticMarkup(createElement(PersonReceipts, { session, userId: "user-1" }))).toBe(
      "",
    );
  });
});

describe("Export for action", () => {
  const definition = auditDefinitionSchema.parse({
    id: "accounts.inactive",
    version: 1,
    title: "Inactive accounts",
    description: "Synthetic.",
    category: "Accounts",
    subject: { one: "account", many: "accounts" },
    columns: [
      { key: "person", label: "Account", kind: "person" },
      { key: "days", label: "Inactive for", kind: "days" },
    ],
    views: [{ id: "all", label: "All", tags: [] }],
    defaultView: "all",
    defaultSort: { column: "days", direction: "desc" },
    permissions: [],
    optionalPermissions: [],
    requirements: [],
  });
  const row: AuditRow = {
    key: "user-1",
    personId: "user-1",
    tags: [],
    cells: { person: { name: "Avery Example", principal: "avery@example.invalid" }, days: 40 },
  };
  const run = {
    id: "66666666-6666-4666-8666-666666666666",
    auditId: "accounts.inactive",
    auditVersion: 1,
    status: "succeeded",
    complete: true,
    startedAt: "2026-10-21T09:00:00.000Z",
    finishedAt: "2026-10-21T09:01:00.000Z",
    startedBy: { id: "22222222-2222-4222-8222-222222222222", name: "Synthetic admin" },
    error: null,
    summary: [],
    notes: [],
    rowCount: 1,
    scope: null,
    domains: [],
    rows: [row],
  } as AuditRunDetail;
  const table = (value = definition) =>
    renderToStaticMarkup(
      createElement(ResultTable, { definition: value, run, onOpenPerson: () => {} }),
    );

  it("is offered next to Export CSV for audits that list accounts", () => {
    const html = table();
    expect(html).toContain("Export for action");
    expect(html.indexOf("Export CSV")).toBeLessThan(html.indexOf("Export for action"));
    // The note appears after the first export, not before.
    expect(html).not.toContain(handoffNote);
    const products = auditDefinitionSchema.parse({
      ...definition,
      columns: [{ key: "product", label: "Product", kind: "text" }],
      defaultSort: { column: "product", direction: "asc" },
    });
    expect(table(products)).not.toContain("Export for action");
  });

  it("asks for the same rows as the CSV export, with the columns shown", () => {
    const query = {
      view: "d90",
      search: " avery ",
      sort: { column: "days", direction: "asc" as const },
      facets: { status: "Disabled", type: "" },
    };
    const csv = new URL(exportHref(run, query), "http://demo.invalid");
    const handoff = new URL(handoffHref(run, query, ["person", "days"]), "http://demo.invalid");
    expect(handoff.pathname).toBe(`/api/v1/audit-runs/${run.id}/handoff.json`);
    expect(handoff.searchParams.get("columns")).toBe("person,days");
    handoff.searchParams.delete("columns");
    expect(handoff.searchParams.toString()).toBe(csv.searchParams.toString());
  });
});

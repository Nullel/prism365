import {
  type AuditRow,
  auditDefinitionSchema,
  handoffSchema,
  queryAuditRows,
} from "@prism365/contracts";
import { describe, expect, it } from "vitest";
import { createDemoState, type Fixtures, handleRequest } from "../src/demo/replay.ts";

// Invented data only: a tiny recording shaped like the real one.
const recordedAt = "2026-10-20T10:00:00.000Z";
const runId = "11111111-1111-4111-8111-111111111111";
const tenantId = "33333333-3333-4333-8333-333333333333";
const actorId = "22222222-2222-4222-8222-222222222222";
const recordedHandoffId = "44444444-4444-4444-8444-444444444444";

const definition = auditDefinitionSchema.parse({
  id: "accounts.inactive",
  version: 1,
  title: "Inactive accounts",
  description: "Every account and how long it has gone without signing in.",
  category: "Accounts",
  subject: { one: "account", many: "accounts" },
  columns: [
    { key: "person", label: "Account", kind: "person" },
    { key: "days", label: "Inactive for", kind: "days" },
    { key: "status", label: "Status", kind: "enum", facet: true },
  ],
  views: [
    { id: "all", label: "All accounts", tags: [] },
    { id: "d30", label: "30+ days", tags: ["d30"] },
  ],
  defaultView: "all",
  defaultSort: { column: "days", direction: "desc" },
  permissions: ["User.Read.All"],
  optionalPermissions: [],
  requirements: [],
});
const products = auditDefinitionSchema.parse({
  ...definition,
  id: "licenses.capacity",
  columns: [{ key: "product", label: "Product", kind: "text" }],
  defaultSort: { column: "product", direction: "asc" },
});

const person = (n: number, days: number, status: string, tags: string[]): AuditRow => ({
  key: `u-${n}`,
  personId: `u-${n}`,
  tags,
  cells: {
    person: { name: `Example ${n}`, principal: `person${n}@larkspur.example` },
    days,
    status,
  },
});
const rows = [
  person(1, 41, "Enabled", ["d30"]),
  person(2, 3, "Disabled", []),
  person(3, 90, "Disabled", ["d30"]),
];
const run = (id: string, auditId: string, body: Record<string, unknown> = {}) => ({
  id,
  auditId,
  auditVersion: 1,
  status: "succeeded",
  complete: true,
  startedAt: "2026-10-20T09:59:00.000Z",
  finishedAt: "2026-10-20T09:59:03.000Z",
  startedBy: { id: actorId, name: "Demo administrator" },
  error: null,
  summary: [],
  notes: [],
  rowCount: 3,
  scope: null,
  domains: [],
  rows,
  ...body,
});
const capacityRun = "55555555-5555-4555-8555-555555555555";
const recordedReceipt = {
  id: "66666666-6666-4666-8666-666666666666",
  handoff: {
    id: recordedHandoffId,
    auditId: "accounts.inactive",
    generatedAt: recordedAt,
    items: 2,
  },
  tool: { name: "Example action tool", version: "1.4.0" },
  performedAt: recordedAt,
  performedBy: "Jordan Vale",
  importedAt: recordedAt,
  items: 1,
  done: 1,
  failed: 0,
  skipped: 0,
};

const fixtures = (withHandoff = true): Fixtures => ({
  recordedAt,
  responses: {
    "GET /api/v1/audits": {
      status: 200,
      body: {
        mode: "live",
        audits: [definition, products].map((value) => ({
          definition: value,
          availability: { state: "available", missing: [], grantFeature: null, enable: [] },
          lastRun: null,
          runningId: null,
        })),
      },
    },
    [`GET /api/v1/audit-runs/${runId}`]: { status: 200, body: run(runId, "accounts.inactive") },
    [`GET /api/v1/audit-runs/${capacityRun}`]: {
      status: 200,
      body: run(capacityRun, "licenses.capacity", { rows: [] }),
    },
    ...(withHandoff
      ? {
          [`GET /api/v1/audit-runs/${runId}/handoff.json?direction=desc&sort=days&view=d30`]: {
            status: 200,
            body: {
              format: "prism365.handoff",
              version: 1,
              id: recordedHandoffId,
              generatedAt: recordedAt,
              tenantId,
              generatedBy: { actorId, name: "Demo administrator" },
              source: {},
              items: [],
            },
          },
        }
      : {}),
    "GET /api/v1/handoffs": {
      status: 200,
      body: {
        handoffs: [
          {
            id: recordedHandoffId,
            auditId: "accounts.inactive",
            auditTitle: "Inactive accounts",
            runId,
            view: { id: "d30", label: "30+ days" },
            filtered: false,
            items: 2,
            generatedAt: recordedAt,
            generatedBy: { id: actorId, name: "Demo administrator" },
          },
        ],
        receipts: [recordedReceipt],
      },
    },
  },
});

function start(withHandoff = true) {
  const state = createDemoState(fixtures(withHandoff), () => new Date(recordedAt));
  return (method: string, url: string, body?: unknown) =>
    handleRequest(state, { method, url, body });
}

const receipt = (id: string, extra: Record<string, unknown> = {}) => ({
  format: "prism365.receipt",
  version: 1,
  id,
  handoffId: recordedHandoffId,
  tool: { name: "Example action tool", version: "1.4.0" },
  performedAt: "2026-10-20T11:00:00Z",
  performedBy: "Jordan Vale",
  items: [
    {
      userId: "u-1",
      userPrincipalName: "person1@larkspur.example",
      action: "account.sign-in-blocked",
      status: "done",
      detail: "",
    },
    {
      userId: null,
      userPrincipalName: "person3@larkspur.example",
      action: "license.removed",
      status: "failed",
      detail: "Synthetic refusal.",
    },
  ],
  ...extra,
});

describe("demo handoffs", () => {
  it("builds the handoff the server would for the page's view, and lists it", () => {
    const request = start();
    const response = request(
      "GET",
      `/api/v1/audit-runs/${runId}/handoff.json?view=d30&sort=days&direction=asc&columns=person,status`,
    );
    expect(response.status).toBe(200);
    expect(response.headers?.["Content-Disposition"]).toBe(
      'attachment; filename="prism365-handoff-accounts.inactive-2026-10-20.json"',
    );
    const document = handoffSchema.parse(JSON.parse(response.body as string));
    expect(document.tenantId).toBe(tenantId);
    expect(document.items.map((item) => item.userId)).toEqual(
      queryAuditRows(definition, rows, {
        view: "d30",
        sort: { column: "days", direction: "asc" },
      }).map((row) => row.personId),
    );
    expect(document.items[0]).toMatchObject({
      reasons: ["30+ days"],
      evidence: { Status: "Enabled" },
    });
    const list = request("GET", "/api/v1/handoffs").body as { handoffs: { id: string }[] };
    expect(list.handoffs.map((item) => item.id)).toEqual([document.id, recordedHandoffId]);
  });

  it("refuses audits without accounts, unknown runs, and a recording with no handoff", () => {
    const request = start();
    expect(request("GET", `/api/v1/audit-runs/${capacityRun}/handoff.json`).status).toBe(404);
    expect(
      request("GET", "/api/v1/audit-runs/ffffffff-ffff-4fff-8fff-ffffffffffff/handoff.json").status,
    ).toBe(404);
    expect(start(false)("GET", `/api/v1/audit-runs/${runId}/handoff.json`).status).toBe(404);
  });

  it("imports a receipt for the page: once, as reported, and listed first", () => {
    const request = start();
    const id = "77777777-7777-4777-8777-777777777777";
    const first = request("POST", "/api/v1/receipts", receipt(id));
    expect(first.status).toBe(201);
    expect(first.body).toMatchObject({
      imported: true,
      receipt: { id, done: 1, failed: 1, skipped: 0, handoff: { id: recordedHandoffId } },
    });
    const again = request("POST", "/api/v1/receipts", receipt(id));
    expect(again.status).toBe(200);
    expect(again.body).toMatchObject({ imported: false });
    expect(
      request("POST", "/api/v1/receipts", receipt(id, { performedBy: "Someone else" })).status,
    ).toBe(409);
    const list = request("GET", "/api/v1/handoffs").body as { receipts: { id: string }[] };
    expect(list.receipts.map((item) => item.id)).toEqual([id, recordedReceipt.id]);
    const page = request("GET", `/api/v1/receipts/${id}?limit=1`).body as {
      items: { position: number }[];
      next: number | null;
    };
    expect(page.items.map((item) => item.position)).toEqual([0]);
    expect(page.next).toBe(0);
    // A receipt the recording already holds is answered as kept.
    expect(request("POST", "/api/v1/receipts", receipt(recordedReceipt.id)).body).toMatchObject({
      imported: false,
    });
  });

  it("refuses what the server refuses, in the same words", () => {
    const request = start();
    const invalid = request(
      "POST",
      "/api/v1/receipts",
      receipt("88888888-8888-4888-8888-888888888888", { version: 2 }),
    );
    expect(invalid.status).toBe(400);
    expect(invalid.body).toMatchObject({
      message: "This receipt is version 2. This Prism365 reads version 1.",
    });
    const unknown = request(
      "POST",
      "/api/v1/receipts",
      receipt("88888888-8888-4888-8888-888888888888", {
        handoffId: "99999999-9999-4999-8999-999999999999",
      }),
    );
    expect(unknown.status).toBe(400);
    expect(unknown.body).toMatchObject({ error: "unknown_handoff" });
  });
});

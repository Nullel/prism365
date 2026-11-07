import { randomUUID } from "node:crypto";
import { type AuditRow, auditRowSchema } from "@prism365/contracts";
import { type FindingInput, setFindingState, upsertFinding } from "@prism365/modules";
import { PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { createStoredResults, finishRun, insertRun } from "../src/index.ts";

// The stored-results accessor on a real (PGlite) database. Synthetic tenants and accounts only.

const tenantA = "11111111-1111-4111-8111-111111111111";
const tenantB = "99999999-9999-4999-8999-999999999999";
const actor = { id: "22222222-2222-4222-8222-222222222222", name: "Synthetic admin" };
const signal = new AbortController().signal;

let database: PGliteModuleDatabase;
beforeAll(async () => {
  database = await PGliteModuleDatabase.open();
}, 60_000);
afterAll(async () => {
  await database?.close();
});

const rowOf = (key: string): AuditRow =>
  auditRowSchema.parse({
    key,
    personId: key,
    tags: [],
    cells: { person: { name: key, principal: null } },
  });

async function storeRun(
  tenantId: string,
  auditId: string,
  day: number,
  outcome: { status: "succeeded"; rows: AuditRow[]; complete?: boolean } | { status: "failed" },
  scope: string[] | null = null,
) {
  const id = randomUUID();
  const start = new Date(Date.UTC(2026, 9, day, 9));
  await database.tenant(tenantId, (sql) =>
    insertRun(sql, tenantId, {
      id,
      auditId,
      auditVersion: 1,
      startedAt: start,
      startedBy: actor,
      scope,
    }),
  );
  await database.tenant(tenantId, (sql) =>
    finishRun(
      sql,
      tenantId,
      id,
      new Date(Date.UTC(2026, 9, day, 10)),
      outcome.status === "succeeded"
        ? {
            status: "succeeded",
            complete: outcome.complete ?? true,
            summary: [],
            notes: [],
            rows: outcome.rows,
          }
        : { status: "failed", code: "failed", message: "The audit failed." },
    ),
  );
  return id;
}

const input = (over: Partial<FindingInput> = {}): FindingInput => {
  const observedAt = new Date("2026-10-22T12:00:00Z");
  return {
    ruleId: "signins.synthetic",
    ruleVersion: 1,
    dedupeKey: randomUUID(),
    purpose: "security",
    severity: "high",
    score: 50,
    subject: { kind: "person", id: "u-ada", label: "Ada Admin" },
    title: "Sign-in from a new country",
    summary: "A synthetic account signed in from a country not seen before for it.",
    evidence: [
      {
        source: "Synthetic fixture",
        ref: "event-1",
        observedAt: observedAt.toISOString(),
        fields: { country: "XA" },
      },
    ],
    limitations: ["Synthetic data."],
    window: { start: new Date(observedAt.getTime() - 3_600_000), end: observedAt },
    observedAt,
    ...over,
  };
};
const addFinding = (tenantId: string, module: string, over: Partial<FindingInput> = {}) =>
  database.tenant(tenantId, (sql) => upsertFinding(sql, tenantId, module, input(over)));
const stored = (tenantId: string) => createStoredResults(database, tenantId, signal);

describe("latestRun", () => {
  it("returns the newest successful run with its rows, scope and completeness", async () => {
    await storeRun(tenantA, "stored.first", 10, { status: "succeeded", rows: [rowOf("old")] });
    await storeRun(
      tenantA,
      "stored.first",
      12,
      { status: "succeeded", rows: [rowOf("new")], complete: false },
      ["a.example"],
    );
    const run = await stored(tenantA).latestRun("stored.first");
    expect(run?.rows.map((item) => item.key)).toEqual(["new"]);
    expect(run).toMatchObject({
      finishedAt: "2026-10-12T10:00:00.000Z",
      complete: false,
      scope: ["a.example"],
    });
  });

  it("is not hidden by a newer failed run, and is null for an audit that never succeeded", async () => {
    await storeRun(tenantA, "stored.second", 10, { status: "succeeded", rows: [rowOf("kept")] });
    await storeRun(tenantA, "stored.second", 14, { status: "failed" });
    expect(
      (await stored(tenantA).latestRun("stored.second"))?.rows.map((item) => item.key),
    ).toEqual(["kept"]);
    await storeRun(tenantA, "stored.failed-only", 14, { status: "failed" });
    expect(await stored(tenantA).latestRun("stored.failed-only")).toBeNull();
    expect(await stored(tenantA).latestRun("stored.never")).toBeNull();
  });

  it("reads only its own tenant", async () => {
    await storeRun(tenantB, "stored.tenant", 11, { status: "succeeded", rows: [rowOf("b-only")] });
    expect(await stored(tenantA).latestRun("stored.tenant")).toBeNull();
    expect(
      (await stored(tenantB).latestRun("stored.tenant"))?.rows.map((item) => item.key),
    ).toEqual(["b-only"]);
  });
});

describe("openFindings", () => {
  it("returns open findings of one module, with what an audit may see and no more", async () => {
    await addFinding(tenantA, "stored.signins", { severity: "critical", title: "Spray" });
    const [found] = await stored(tenantA).openFindings("stored.signins");
    expect(found).toEqual({
      ruleId: "signins.synthetic",
      severity: "critical",
      subjectKind: "person",
      subjectId: "u-ada",
      title: "Spray",
      lastSeen: "2026-10-22T12:00:00.000Z",
    });
  });

  it("leaves out resolved and acknowledged findings, and other modules", async () => {
    const resolved = await addFinding(tenantA, "stored.filtered");
    const acknowledged = await addFinding(tenantA, "stored.filtered");
    await addFinding(tenantA, "stored.filtered", { title: "Still open" });
    await addFinding(tenantA, "stored.elsewhere", { title: "Another module" });
    await database.tenant(tenantA, async (sql) => {
      await setFindingState(sql, tenantA, resolved.id, "resolved", actor.id);
      await setFindingState(sql, tenantA, acknowledged.id, "acknowledged", actor.id);
    });
    expect(
      (await stored(tenantA).openFindings("stored.filtered")).map((item) => item.title),
    ).toEqual(["Still open"]);
  });

  it("reads only its own tenant", async () => {
    await addFinding(tenantB, "stored.isolated", { title: "Tenant B" });
    expect(await stored(tenantA).openFindings("stored.isolated")).toEqual([]);
    expect(
      (await stored(tenantB).openFindings("stored.isolated")).map((item) => item.title),
    ).toEqual(["Tenant B"]);
  });
});

describe("the accessor", () => {
  it("offers two reads and nothing else, and changes nothing", async () => {
    const results = stored(tenantA);
    expect(Object.keys(results).sort()).toEqual(["latestRun", "openFindings"]);
    const count = async () =>
      database.tenant(tenantA, async (sql) => {
        const runs = await sql.query<{ n: number }>(
          "SELECT count(*)::int AS n FROM prism365_audit_run WHERE tenant_id=$1",
          [tenantA],
        );
        const findings = await sql.query<{ n: number }>(
          "SELECT count(*)::int AS n FROM prism365_finding WHERE tenant_id=$1",
          [tenantA],
        );
        return [runs.rows[0]?.n, findings.rows[0]?.n];
      });
    const before = await count();
    await results.latestRun("stored.first");
    await results.openFindings("stored.signins");
    expect(await count()).toEqual(before);
  });

  it("reads nothing once the run has been stopped", async () => {
    const controller = new AbortController();
    const results = createStoredResults(database, tenantA, controller.signal);
    controller.abort(new Error("stopped"));
    await expect(results.latestRun("stored.first")).rejects.toThrow("stopped");
    await expect(results.openFindings("stored.signins")).rejects.toThrow("stopped");
  });
});

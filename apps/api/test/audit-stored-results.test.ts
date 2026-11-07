import { randomUUID } from "node:crypto";
import { type Audit, type AuditContext, finishRun, insertRun } from "@prism365/audits";
import { auditRowSchema } from "@prism365/contracts";
import { GraphHttpClient } from "@prism365/graph";
import { PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { AuditRunner } from "../src/audits.ts";

// The runner gives stored results to an audit that declares it uses them, and to no other.
// Synthetic tenants and accounts only.

const tenantA = "11111111-1111-4111-8111-111111111111";
const tenantB = "99999999-9999-4999-8999-999999999999";
const actor = { id: "22222222-2222-4222-8222-222222222222", name: "Synthetic admin" };
const client = new GraphHttpClient({
  getAccessToken: async () => "synthetic",
  fetchImplementation: async () => {
    throw new Error("These audits read nothing from Microsoft.");
  },
});

let database: PGliteModuleDatabase;
beforeAll(async () => {
  database = await PGliteModuleDatabase.open();
  for (const [tenant, key] of [
    [tenantA, "a-only"],
    [tenantB, "b-only"],
  ] as const) {
    const id = randomUUID();
    await database.tenant(tenant, async (sql) => {
      await insertRun(sql, tenant, {
        id,
        auditId: "stored.source",
        auditVersion: 1,
        startedAt: new Date("2026-10-20T09:00:00Z"),
        startedBy: actor,
      });
      await finishRun(sql, tenant, id, new Date("2026-10-20T10:00:00Z"), {
        status: "succeeded",
        complete: true,
        summary: [],
        notes: [],
        rows: [auditRowSchema.parse({ key, personId: key, tags: [], cells: {} })],
      });
    });
  }
}, 60_000);
afterAll(async () => {
  await database?.close();
});

function probe(id: string, usesStoredResults: boolean) {
  const seen: { stored: AuditContext["stored"]; keys: string[] | null }[] = [];
  const audit: Audit = {
    definition: {
      id,
      version: 1,
      title: "Probe",
      description: "Records what the runner gave it.",
      category: "Test",
      subject: { one: "row", many: "rows" },
      columns: [{ key: "value", label: "Value", kind: "text" }],
      views: [{ id: "all", label: "All", tags: [] }],
      defaultView: "all",
      defaultSort: { column: "value", direction: "asc" },
      requirements: [],
    },
    scoped: false,
    packs: [],
    optionalPacks: [],
    forbidden: "Not used.",
    ...(usesStoredResults ? { usesStoredResults } : {}),
    async run(context) {
      const run = (await context.stored?.latestRun("stored.source")) ?? null;
      seen.push({ stored: context.stored, keys: run ? run.rows.map((item) => item.key) : null });
      return { rows: [], summary: [], notes: [], complete: true };
    },
  };
  return { audit, seen };
}

async function runAudit(tenantId: string, audit: Audit) {
  const runner = new AuditRunner({ database });
  const started = await runner.start({ tenantId, actor, audit, client, packs: [], granted: [] });
  await started.done;
  return runner;
}

describe("stored results in the runner", () => {
  it("are given to an audit that declares it uses them", async () => {
    const { audit, seen } = probe("stored.uses", true);
    await runAudit(tenantA, audit);
    expect(seen).toHaveLength(1);
    expect(seen[0]?.stored).toBeDefined();
  });

  it("are given to no other audit", async () => {
    const { audit, seen } = probe("stored.plain", false);
    await runAudit(tenantA, audit);
    expect(seen).toHaveLength(1);
    expect(seen[0]?.stored).toBeUndefined();
    expect(seen[0]?.keys).toBeNull();
  });

  it("read the run's own tenant and never another", async () => {
    const a = probe("stored.tenant-a", true);
    const b = probe("stored.tenant-b", true);
    await runAudit(tenantA, a.audit);
    await runAudit(tenantB, b.audit);
    expect(a.seen[0]?.keys).toEqual(["a-only"]);
    expect(b.seen[0]?.keys).toEqual(["b-only"]);
  });

  it("stop reading when the run is stopped", async () => {
    const runner = new AuditRunner({ database, timeoutMilliseconds: 20 });
    let outcome = "not asked";
    const slow = probe("stored.slow", true).audit;
    const audit: Audit = {
      ...slow,
      async run(context) {
        // Past the time limit: the run is already closed as timed out when this asks.
        await new Promise((resolve) => setTimeout(resolve, 80));
        outcome = await (context.stored?.latestRun("stored.source") ?? Promise.resolve(null)).then(
          () => "read",
          () => "refused",
        );
        return { rows: [], summary: [], notes: [], complete: true };
      },
    };
    const started = await runner.start({
      tenantId: tenantA,
      actor,
      audit,
      client,
      packs: [],
      granted: [],
    });
    await started.done;
    await new Promise((resolve) => setTimeout(resolve, 200));
    expect(outcome).toBe("refused");
  });
});

import type { PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { FindingPolicyError, listFindings, setFindingState, upsertFinding } from "../src/index.ts";
import { actor, finding, openRuntime, probeModule, tenantA, tenantB } from "./fixtures.ts";

let database: PGliteModuleDatabase;

beforeAll(async () => {
  ({ database } = await openRuntime(probeModule({ run: async () => undefined })));
}, 60_000);
afterAll(async () => {
  await database?.close();
});

const upsert = (tenantId: string, input = finding()) =>
  database.tenant(tenantId, (sql) => upsertFinding(sql, tenantId, "test.probe", input));
const list = (tenantId: string) =>
  database.tenant(tenantId, (sql) => listFindings(sql, tenantId, { module: "test.probe" }));

describe("findings", () => {
  it("creates once, folds repeats into one finding, and keeps first-seen", async () => {
    const first = await upsert(tenantA, finding({ dedupeKey: "fold" }));
    expect(first.change).toBe("created");
    const later = new Date("2026-10-22T13:00:00Z");
    const second = await upsert(tenantA, finding({ dedupeKey: "fold", observedAt: later }));
    expect(second).toEqual({ id: first.id, change: "updated" });
    const [record] = (await list(tenantA)).filter((item) => item.id === first.id);
    expect(record).toMatchObject({
      occurrences: 2,
      state: "open",
      firstSeen: "2026-10-22T12:00:00.000Z",
      lastSeen: later.toISOString(),
    });
  });

  it("reopens a resolved finding when it recurs, and escalates on higher severity", async () => {
    const { id } = await upsert(tenantA, finding({ dedupeKey: "lifecycle" }));
    await database.tenant(tenantA, (sql) => setFindingState(sql, tenantA, id, "resolved", actor));
    expect((await upsert(tenantA, finding({ dedupeKey: "lifecycle" }))).change).toBe("reopened");

    await database.tenant(tenantA, (sql) =>
      setFindingState(sql, tenantA, id, "acknowledged", actor),
    );
    expect((await upsert(tenantA, finding({ dedupeKey: "lifecycle" }))).change).toBe("updated");
    const escalated = await upsert(tenantA, finding({ dedupeKey: "lifecycle", severity: "high" }));
    expect(escalated.change).toBe("escalated");
    const [record] = (await list(tenantA)).filter((item) => item.id === id);
    expect(record).toMatchObject({ severity: "high", state: "open" });

    // A lower severity later does not downgrade the stored finding.
    await upsert(tenantA, finding({ dedupeKey: "lifecycle", severity: "low" }));
    const [after] = (await list(tenantA)).filter((item) => item.id === id);
    expect(after?.severity).toBe("high");
  });

  it("keeps findings and state changes inside their tenant", async () => {
    const { id } = await upsert(tenantA, finding({ dedupeKey: "isolated" }));
    expect((await list(tenantB)).map((item) => item.id)).not.toContain(id);
    const fromB = await database.tenant(tenantB, (sql) =>
      setFindingState(sql, tenantB, id, "resolved", actor),
    );
    expect(fromB).toBeNull();
    const [still] = (await list(tenantA)).filter((item) => item.id === id);
    expect(still?.state).toBe("open");
  });

  it("refuses productivity or performance language anywhere in a finding", async () => {
    await expect(
      upsert(tenantA, finding({ dedupeKey: "p1", title: "Low performer sign-in pattern" })),
    ).rejects.toBeInstanceOf(FindingPolicyError);
    await expect(
      upsert(
        tenantA,
        finding({
          dedupeKey: "p2",
          evidence: [
            {
              source: "Synthetic fixture",
              ref: "e",
              observedAt: "2026-10-22T12:00:00Z",
              fields: { note: "employee productivity dropped" },
            },
          ],
        }),
      ),
    ).rejects.toBeInstanceOf(FindingPolicyError);
  });

  it("refuses findings without evidence, limitations, or a coherent window", async () => {
    await expect(
      upsert(tenantA, finding({ dedupeKey: "e1", evidence: [] })),
    ).rejects.toBeInstanceOf(FindingPolicyError);
    await expect(
      upsert(tenantA, finding({ dedupeKey: "e2", limitations: [] })),
    ).rejects.toBeInstanceOf(FindingPolicyError);
    await expect(
      upsert(
        tenantA,
        finding({
          dedupeKey: "e3",
          window: {
            start: new Date("2026-10-22T12:00:00Z"),
            end: new Date("2026-10-21T12:00:00Z"),
          },
        }),
      ),
    ).rejects.toBeInstanceOf(FindingPolicyError);
  });
});

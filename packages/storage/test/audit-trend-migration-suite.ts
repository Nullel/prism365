import { randomUUID } from "node:crypto";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { loadCoreMigrations } from "../src/index.ts";

// Migration 005 (audit trends) on a database that already holds audit runs. Shared by the
// PGlite run (audit-trend-migration.test.ts) and the PostgreSQL run (postgres.test.ts), so both
// backends prove the same backfill with the same assertions. Synthetic tenants and labels only.

export interface TrendMigrationFixture {
  // Runs one statement as the database owner, outside any tenant transaction, so row security
  // does not hide other tenants' rows from the assertions.
  owner(sql: string, params?: unknown[]): Promise<Record<string, unknown>[]>;
  // Applies every core migration the way the backend does on start (PGlite) or migrate
  // (PostgreSQL's runPostgresMigrations).
  migrateAll(): Promise<void>;
  close(): Promise<void>;
}

// The core migrations before 005, which a fixture applies before handing the database over.
export async function migrationsBeforeTrends() {
  return (await loadCoreMigrations()).filter((migration) => migration.id < "005");
}

async function trendMigration(): Promise<string> {
  const migration = (await loadCoreMigrations()).find((item) => item.id === "005-audit-trends");
  if (!migration) throw new Error("Migration 005-audit-trends is not loaded.");
  return migration.sql;
}

const tenantA = randomUUID();
const tenantB = randomUUID();
const actor = randomUUID();
const ids = {
  complete: randomUUID(),
  partial: randomUUID(),
  empty: randomUUID(),
  failed: randomUUID(),
  running: randomUUID(),
  interrupted: randomUUID(),
  otherTenant: randomUUID(),
  later: randomUUID(),
};

const summary = [
  { label: "Accounts", value: 214, tone: "neutral", view: "all" },
  { label: "Inactive 30+ days", value: 47, tone: "warning", view: "d30" },
  { label: "Attempts, no success", value: 2, tone: "critical", view: "no-success" },
  // Letters outside ASCII become separators, whatever the database's locale.
  { label: "  Ünïcode --- label  ", value: 1, tone: "neutral", view: null },
  // The same id as an earlier tile: the first one is kept.
  { label: "ACCOUNTS", value: 999, tone: "warning", view: null },
  // Nothing left of the label: the id falls back to "tile".
  { label: "—", value: 3, tone: "neutral", view: null },
  { label: `${"x".repeat(70)} tail`, value: 4, tone: "good", view: null },
  // Not a count: left out.
  { label: "Kept as text", value: "7", tone: "neutral", view: null },
];
const expectedTiles = [
  { id: "accounts", label: "Accounts", value: 214, tone: "neutral" },
  { id: "inactive-30-days", label: "Inactive 30+ days", value: 47, tone: "warning" },
  { id: "attempts-no-success", label: "Attempts, no success", value: 2, tone: "critical" },
  { id: "n-code-label", label: "  Ünïcode --- label  ", value: 1, tone: "neutral" },
  { id: "tile", label: "—", value: 3, tone: "neutral" },
  { id: "x".repeat(60), label: `${"x".repeat(70)} tail`, value: 4, tone: "good" },
];

export function trendMigrationSuite(
  name: string,
  open: () => Promise<TrendMigrationFixture>,
  enabled = true,
) {
  describe.skipIf(!enabled)(`${name}: audit trend migration`, () => {
    let fixture: TrendMigrationFixture;
    const owner = (sql: string, params?: unknown[]) => fixture.owner(sql, params);
    const insertRun = (run: {
      id: string;
      tenant: string;
      status: "running" | "succeeded" | "failed";
      complete?: boolean;
      summary?: unknown;
      scope?: string[] | null;
      errorCode?: string | null;
      day: number;
    }) =>
      owner(
        `INSERT INTO prism365_audit_run (id, tenant_id, audit_id, audit_version, status, complete,
           started_at, finished_at, started_by, summary, row_count, run_scope, error_code)
         VALUES ($1,$2,'accounts.inactive',1,$3,$4,$5,$6,$7,$8::jsonb,$9,$10::jsonb,$11)`,
        [
          run.id,
          run.tenant,
          run.status,
          run.complete ?? false,
          new Date(Date.UTC(2026, 9, run.day, 9)).toISOString(),
          run.status === "running" ? null : new Date(Date.UTC(2026, 9, run.day, 10)).toISOString(),
          actor,
          JSON.stringify(run.summary ?? []),
          run.status === "succeeded" ? 214 : 0,
          run.scope ? JSON.stringify(run.scope) : null,
          run.errorCode ?? null,
        ],
      );
    const points = () =>
      owner(
        `SELECT tenant_id::text AS tenant, run_id::text AS run, audit_id, audit_version,
                finished_at, complete, run_scope, tiles, row_count, trigger
           FROM prism365_audit_trend ORDER BY finished_at, run_id`,
      );

    beforeAll(async () => {
      fixture = await open();
      // Runs stored before migration 005: every outcome a run can have.
      await insertRun({
        id: ids.complete,
        tenant: tenantA,
        status: "succeeded",
        complete: true,
        summary,
        day: 1,
      });
      await insertRun({
        id: ids.partial,
        tenant: tenantA,
        status: "succeeded",
        complete: false,
        summary: summary.slice(0, 2),
        scope: ["a.example"],
        day: 2,
      });
      // A summary that is not a list still makes a point, with no tiles.
      await insertRun({ id: ids.empty, tenant: tenantA, status: "succeeded", summary: {}, day: 3 });
      await insertRun({
        id: ids.failed,
        tenant: tenantA,
        status: "failed",
        errorCode: "forbidden",
        day: 4,
      });
      await insertRun({ id: ids.running, tenant: tenantA, status: "running", day: 5 });
      await insertRun({
        id: ids.interrupted,
        tenant: tenantA,
        status: "failed",
        errorCode: "interrupted",
        day: 6,
      });
      await insertRun({
        id: ids.otherTenant,
        tenant: tenantB,
        status: "succeeded",
        complete: true,
        summary: summary.slice(1, 2),
        day: 7,
      });
      await owner(await trendMigration());
    }, 60_000);
    afterAll(async () => {
      await fixture?.close();
    }, 30_000);

    it("turns each succeeded run, complete or not, into one point and nothing else", async () => {
      const rows = await points();
      expect(rows.map((row) => row.run)).toEqual([
        ids.complete,
        ids.partial,
        ids.empty,
        ids.otherTenant,
      ]);
      const [complete, partial, empty, other] = rows;
      expect(complete).toMatchObject({
        tenant: tenantA,
        audit_id: "accounts.inactive",
        audit_version: 1,
        complete: true,
        run_scope: null,
        row_count: 214,
        trigger: "manual",
        tiles: expectedTiles,
      });
      expect(new Date(complete?.finished_at as Date | string).toISOString()).toBe(
        "2026-10-01T10:00:00.000Z",
      );
      expect(partial).toMatchObject({
        complete: false,
        run_scope: ["a.example"],
        tiles: expectedTiles.slice(0, 2),
      });
      expect(empty).toMatchObject({ complete: false, tiles: [] });
      expect(other).toMatchObject({ tenant: tenantB, tiles: [expectedTiles[1]] });
    });

    it("marks every run made before schedules as manual", async () => {
      const runs = await owner("SELECT DISTINCT trigger FROM prism365_audit_run");
      expect(runs).toEqual([{ trigger: "manual" }]);
    });

    it("puts the trend table under the one tenant policy", async () => {
      const [table] = await owner(
        `SELECT c.relrowsecurity AS rls, c.relforcerowsecurity AS forced,
                (SELECT count(*)::int FROM pg_policy p WHERE p.polrelid = c.oid) AS policies
           FROM pg_class c WHERE c.relname = 'prism365_audit_trend'`,
      );
      expect(table).toEqual({ rls: true, forced: true, policies: 1 });
    });

    it("changes nothing when applied again, and never replaces a point", async () => {
      const before = await points();
      // A point already there is kept as it is, even when its run says otherwise.
      await owner("UPDATE prism365_audit_trend SET tiles='[]'::jsonb WHERE run_id=$1", [
        ids.complete,
      ]);
      await owner(await trendMigration());
      await owner(await trendMigration());
      const after = await points();
      expect(after).toHaveLength(before.length);
      expect(after.find((row) => row.run === ids.complete)?.tiles).toEqual([]);
      expect(after.filter((row) => row.run !== ids.complete)).toEqual(
        before.filter((row) => row.run !== ids.complete),
      );
      // Re-running the ADD COLUMN does not stack up check constraints.
      const checks = await owner(
        `SELECT count(*)::int AS checks FROM pg_constraint
          WHERE conrelid = 'prism365_audit_run'::regclass AND contype = 'c'
            AND pg_get_constraintdef(oid) LIKE '%trigger%'`,
      );
      expect(checks).toEqual([{ checks: 1 }]);
    });

    it("gives a run that succeeded since the last start its point on the next one", async () => {
      await insertRun({
        id: ids.later,
        tenant: tenantA,
        status: "succeeded",
        complete: true,
        summary: summary.slice(0, 1),
        day: 8,
      });
      await owner(await trendMigration());
      const later = (await points()).find((row) => row.run === ids.later);
      expect(later).toMatchObject({ tiles: [expectedTiles[0]], trigger: "manual" });
    });

    it("keeps every point when the whole migrator runs again", async () => {
      const before = await points();
      await fixture.migrateAll();
      await fixture.migrateAll();
      expect(await points()).toEqual(before);
    });

    it("refuses a trigger that is neither manual nor schedule", async () => {
      await expect(
        owner("UPDATE prism365_audit_run SET trigger='cron' WHERE id=$1", [ids.complete]),
      ).rejects.toThrow();
      await expect(
        owner("UPDATE prism365_audit_trend SET trigger='cron' WHERE run_id=$1", [ids.partial]),
      ).rejects.toThrow();
    });
  });
}

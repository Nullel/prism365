import { randomUUID } from "node:crypto";
import { readFile } from "node:fs/promises";
import { type AuditSummaryItem, auditTileId } from "@prism365/contracts";
import { PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import {
  auditDue,
  auditTrendTiles,
  buildAuditTrend,
  builtInAudits,
  finishRun,
  insertRun,
  interruptRuns,
  KEEP_RUNS,
  knownAuditSchedule,
  latestRunStart,
  listRuns,
  pruneRuns,
  readAuditSchedule,
  readStoredAuditSchedule,
  readTrendPoints,
  type StoredTrendPoint,
  scheduledAuditCount,
  writeAuditSchedule,
} from "../src/index.ts";

// Trend points and schedules on a real (PGlite) database. Synthetic tenants and labels only.
const tenantId = "11111111-1111-4111-8111-111111111111";
const otherTenantId = "99999999-9999-4999-8999-999999999999";
const actor = { id: "22222222-2222-4222-8222-222222222222", name: "Synthetic admin" };
const otherActor = "44444444-4444-4444-8444-444444444444";
const at = (day: number, hour = 9) => new Date(Date.UTC(2026, 9, day, hour));
const epoch = new Date(0);

const summary: AuditSummaryItem[] = [
  { label: "Accounts", value: 214, tone: "neutral", view: "all" },
  { label: "Inactive 30+ days", value: 47, tone: "warning", view: "d30" },
  { label: "Attempts, no success", value: 2, tone: "critical", view: "no-success" },
];

let database: PGliteModuleDatabase;
beforeAll(async () => {
  database = await PGliteModuleDatabase.open();
}, 60_000);
afterAll(async () => {
  await database?.close();
});

// Stores a run and finishes it the way the runner does: finish and prune in one transaction.
async function run(
  auditId: string,
  day: number,
  outcome:
    | { status: "succeeded"; complete?: boolean; summary?: AuditSummaryItem[] }
    | { status: "failed" },
  options: { scope?: string[]; trigger?: "manual" | "schedule"; version?: number } = {},
): Promise<string> {
  const id = randomUUID();
  await database.tenant(tenantId, async (sql) => {
    await insertRun(sql, tenantId, {
      id,
      auditId,
      auditVersion: options.version ?? 1,
      startedAt: at(day),
      startedBy: actor,
      scope: options.scope ?? null,
      trigger: options.trigger ?? "manual",
    });
  });
  await database.tenant(tenantId, async (sql) => {
    await finishRun(
      sql,
      tenantId,
      id,
      at(day, 10),
      outcome.status === "succeeded"
        ? {
            status: "succeeded",
            complete: outcome.complete ?? true,
            summary: outcome.summary ?? summary,
            notes: [],
            rows: [],
          }
        : { status: "failed", code: "forbidden", message: "Microsoft refused." },
    );
    await pruneRuns(sql, tenantId, auditId);
  });
  return id;
}

const points = (auditId: string) =>
  database.tenant(tenantId, (sql) => readTrendPoints(sql, tenantId, auditId, epoch));

describe("trend points", () => {
  it("writes one point for a succeeded run, with its tiles, scope and trigger", async () => {
    const id = await run(
      "accounts.inactive",
      1,
      { status: "succeeded", complete: false },
      {
        scope: ["b.example", "a.example"],
        trigger: "schedule",
      },
    );
    const { points: stored, truncated } = await points("accounts.inactive");
    expect(truncated).toBe(false);
    expect(stored).toEqual([
      {
        runId: id,
        finishedAt: at(1, 10).toISOString(),
        version: 1,
        complete: false,
        scope: ["a.example", "b.example"],
        trigger: "schedule",
        tiles: [
          { id: "accounts", label: "Accounts", value: 214, tone: "neutral" },
          { id: "inactive-30-days", label: "Inactive 30+ days", value: 47, tone: "warning" },
          { id: "attempts-no-success", label: "Attempts, no success", value: 2, tone: "critical" },
        ],
      },
    ]);
  });

  it("writes the point in the result's transaction: no result, no point", async () => {
    const id = randomUUID();
    await database.tenant(tenantId, (sql) =>
      insertRun(sql, tenantId, {
        id,
        auditId: "licenses.capacity",
        auditVersion: 1,
        startedAt: at(2),
        startedBy: actor,
      }),
    );
    await expect(
      database.tenant(tenantId, async (sql) => {
        await finishRun(sql, tenantId, id, at(2, 10), {
          status: "succeeded",
          complete: true,
          summary,
          notes: [],
          rows: [],
        });
        // Pruning or anything else after the result fails: the whole result is rolled back.
        throw new Error("storage refused");
      }),
    ).rejects.toThrow("storage refused");
    expect((await points("licenses.capacity")).points).toEqual([]);
    // Finishing it for real then leaves exactly one point, and a second finish none more.
    for (let attempt = 0; attempt < 2; attempt += 1)
      await database.tenant(tenantId, (sql) =>
        finishRun(sql, tenantId, id, at(2, 10), {
          status: "succeeded",
          complete: true,
          summary,
          notes: [],
          rows: [],
        }),
      );
    expect((await points("licenses.capacity")).points.map((point) => point.runId)).toEqual([id]);
  });

  it("writes no point for a failed run or an interrupted one", async () => {
    await run("guests.review", 3, { status: "failed" });
    const interrupted = randomUUID();
    await database.tenant(tenantId, async (sql) => {
      await insertRun(sql, tenantId, {
        id: interrupted,
        auditId: "guests.review",
        auditVersion: 1,
        startedAt: at(3, 11),
        startedBy: actor,
      });
      await interruptRuns(sql, tenantId, [], at(3, 12));
    });
    const runs = await database.tenant(tenantId, (sql) => listRuns(sql, tenantId, "guests.review"));
    expect(runs.map((item) => item.status)).toEqual(["failed", "failed"]);
    expect(runs.map((item) => item.trigger)).toEqual(["manual", "manual"]);
    expect((await points("guests.review")).points).toEqual([]);
  });

  it("keeps every point when old runs are pruned", async () => {
    const ids: string[] = [];
    for (let day = 1; day <= KEEP_RUNS + 3; day += 1)
      ids.push(await run("mfa.registration", day, { status: "succeeded" }));
    const runs = await database.tenant(tenantId, (sql) =>
      listRuns(sql, tenantId, "mfa.registration"),
    );
    expect(runs).toHaveLength(KEEP_RUNS);
    expect((await points("mfa.registration")).points.map((point) => point.runId)).toEqual(ids);
  });

  it("reads one audit, one tenant and one window, oldest first, and says when it cut", async () => {
    await database.tenant(otherTenantId, async (sql) => {
      const id = randomUUID();
      await insertRun(sql, otherTenantId, {
        id,
        auditId: "mfa.registration",
        auditVersion: 1,
        startedAt: at(20),
        startedBy: actor,
      });
      await finishRun(sql, otherTenantId, id, at(20, 10), {
        status: "succeeded",
        complete: true,
        summary,
        notes: [],
        rows: [],
      });
    });
    const all = (await points("mfa.registration")).points;
    expect(all).toHaveLength(KEEP_RUNS + 3);
    const recent = await database.tenant(tenantId, (sql) =>
      readTrendPoints(sql, tenantId, "mfa.registration", at(11), 100),
    );
    expect(recent.points.map((point) => point.finishedAt)).toEqual(
      all.slice(-3).map((point) => point.finishedAt),
    );
    const cut = await database.tenant(tenantId, (sql) =>
      readTrendPoints(sql, tenantId, "mfa.registration", epoch, 4),
    );
    expect(cut.truncated).toBe(true);
    // The newest are kept.
    expect(cut.points).toEqual(all.slice(-4));
  });

  it("derives tile ids in the migration's backfill exactly as auditTileId does", async () => {
    const labels = [
      "Accounts",
      "Inactive 90+ days",
      "Admins without MFA outside the scope",
      "Licenses held",
      "  Ünïcode --- label  ",
      "İstanbul office",
      "Café 😀 seats",
      "—",
      "-leading and trailing-",
      `${"y".repeat(59)}-z and more`,
    ];
    const id = randomUUID();
    // A run stored as an older version would have stored it, with no point, then the migration
    // applied again as it is on every start.
    await database.ownerExecForTests(
      `INSERT INTO prism365_audit_run (id, tenant_id, audit_id, audit_version, status, complete,
         started_at, finished_at, started_by, summary, row_count)
       VALUES ('${id}', '${tenantId}', 'parity.check', 1, 'succeeded', true,
         '2026-10-25T09:00:00Z', '2026-10-25T10:00:00Z', '${actor.id}',
         '${JSON.stringify(labels.map((label, index) => ({ label, value: index, tone: "neutral", view: null }))).replaceAll("'", "''")}'::jsonb,
         0)`,
    );
    await database.ownerExecForTests(
      await readFile(
        new URL("../../storage/migrations/005-audit-trends.sql", import.meta.url),
        "utf8",
      ),
    );
    const [point] = (await points("parity.check")).points;
    const expected = auditTrendTiles(
      labels.map((label, index) => ({ label, value: index, tone: "neutral", view: null })),
    );
    expect(point?.tiles).toEqual(expected);
    expect(point?.tiles.map((tile) => tile.id)).toEqual([...new Set(labels.map(auditTileId))]);
  });
});

describe("tile ids", () => {
  it("keep ASCII letters and digits, one hyphen for anything else, at most 60 characters", () => {
    expect(auditTileId("Inactive 30+ days")).toBe("inactive-30-days");
    expect(auditTileId("Admins without MFA")).toBe("admins-without-mfa");
    expect(auditTileId("  Phone or text only  ")).toBe("phone-or-text-only");
    expect(auditTileId("İstanbul")).toBe("stanbul");
    expect(auditTileId("—")).toBe("tile");
    expect(auditTileId(`${"a".repeat(59)} b`)).toBe("a".repeat(59));
    expect(auditTileId("a".repeat(80))).toHaveLength(60);
  });

  it("keep the first of two tiles whose labels give the same id", () => {
    expect(
      auditTrendTiles([
        { label: "Accounts", value: 1, tone: "neutral", view: null },
        { label: "accounts!", value: 2, tone: "warning", view: null },
      ]),
    ).toEqual([{ id: "accounts", label: "Accounts", value: 1, tone: "neutral" }]);
  });
});

const point = (overrides: Partial<StoredTrendPoint> = {}): StoredTrendPoint => ({
  runId: randomUUID(),
  finishedAt: at(1, 10).toISOString(),
  version: 2,
  complete: true,
  scope: null,
  trigger: "manual",
  tiles: [
    { id: "accounts", label: "Accounts", value: 10, tone: "neutral" },
    { id: "inactive-30-days", label: "Inactive 30+ days", value: 3, tone: "warning" },
  ],
  ...overrides,
});

describe("the trend the API returns", () => {
  it("compares every point with the newest: version, scope and partial results", () => {
    const points = [
      point({ version: 1 }),
      point({ scope: ["a.example"] }),
      point({ complete: false }),
      point({ version: 1, scope: ["a.example"], complete: false }),
      point({ scope: ["b.example"] }),
      point({ scope: ["a.example", "b.example"] }),
      point({ scope: ["a.example", "b.example"] }),
    ];
    const trend = buildAuditTrend({
      auditId: "accounts.inactive",
      days: 365,
      points,
      truncated: false,
    });
    expect(trend.points.map((item) => item.notComparable)).toEqual([
      ["version", "scope"],
      ["scope"],
      ["scope", "incomplete"],
      ["version", "scope", "incomplete"],
      ["scope"],
      [],
      [],
    ]);
    expect(trend.points.map((item) => item.scopeLimited)).toEqual([
      false,
      true,
      false,
      true,
      true,
      true,
      true,
    ]);
  });

  it("marks the newest point itself only when it is partial", () => {
    const trend = buildAuditTrend({
      auditId: "accounts.inactive",
      days: 30,
      points: [point(), point({ complete: false })],
      truncated: false,
    });
    expect(trend.points.map((item) => item.notComparable)).toEqual([[], ["incomplete"]]);
  });

  it("maps each point's values by tile id and lists the newest point's tiles", () => {
    const newest = point({
      trigger: "schedule",
      tiles: [
        { id: "accounts", label: "Accounts", value: 12, tone: "neutral" },
        { id: "never-signed-in", label: "Never signed in", value: 1, tone: "critical" },
      ],
    });
    const trend = buildAuditTrend({
      auditId: "accounts.inactive",
      days: 365,
      points: [point(), newest],
      truncated: true,
    });
    expect(trend.points[0]?.tiles).toEqual({ accounts: 10, "inactive-30-days": 3 });
    expect(trend.points[1]).toMatchObject({
      runId: newest.runId,
      trigger: "schedule",
      tiles: { accounts: 12, "never-signed-in": 1 },
    });
    expect(trend.tiles).toEqual([
      { id: "accounts", label: "Accounts", tone: "neutral" },
      { id: "never-signed-in", label: "Never signed in", tone: "critical" },
    ]);
    expect(trend.truncated).toBe(true);
  });

  it("is empty with no points", () => {
    expect(
      buildAuditTrend({ auditId: "accounts.inactive", days: 1, points: [], truncated: false }),
    ).toEqual({ auditId: "accounts.inactive", days: 1, points: [], tiles: [], truncated: false });
  });
});

describe("audit schedules", () => {
  it("reads every audit as off until one is saved", async () => {
    const schedule = await database.tenant(tenantId, (sql) =>
      readAuditSchedule(sql, tenantId, builtInAudits),
    );
    expect(Object.keys(schedule)).toEqual(builtInAudits.map((audit) => audit.definition.id));
    expect(Object.values(schedule).every((frequency) => frequency === "off")).toBe(true);
  });

  it("saves the audits named, drops unknown ones from a request and keeps tenants apart", async () => {
    const saved = await database.tenant(tenantId, (sql) =>
      writeAuditSchedule(
        sql,
        tenantId,
        {
          audits: {
            "accounts.inactive": "daily",
            "unknown.audit": "weekly",
            "guests.review": "weekly",
          },
        },
        actor.id,
        builtInAudits,
      ),
    );
    expect(saved).toMatchObject({
      "accounts.inactive": "daily",
      "guests.review": "weekly",
      "mfa.registration": "off",
    });
    expect(saved).not.toHaveProperty("unknown.audit");
    expect(scheduledAuditCount(saved)).toBe(2);
    const read = await database.tenant(tenantId, (sql) =>
      readAuditSchedule(sql, tenantId, builtInAudits),
    );
    expect(read).toEqual(saved);
    const other = await database.tenant(otherTenantId, (sql) =>
      readAuditSchedule(sql, otherTenantId, builtInAudits),
    );
    expect(scheduledAuditCount(other)).toBe(0);
    // Saving again changes only what it names: an audit left out keeps its schedule.
    const next = await database.tenant(tenantId, (sql) =>
      writeAuditSchedule(
        sql,
        tenantId,
        { audits: { "guests.review": "daily" } },
        actor.id,
        builtInAudits,
      ),
    );
    expect(next).toMatchObject({ "accounts.inactive": "daily", "guests.review": "daily" });
    const stored = await database.tenant(tenantId, (sql) => readStoredAuditSchedule(sql, tenantId));
    expect(stored.scheduledBy).toEqual({
      "accounts.inactive": actor.id,
      "guests.review": actor.id,
    });
  });

  it("keeps the entries a newer build saved for audits this build does not know", async () => {
    const tenant = randomUUID();
    await database.tenant(tenant, (sql) =>
      sql.query(
        `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
         VALUES ($1,'audits.schedule',$2::jsonb,now(),$3)`,
        [
          tenant,
          JSON.stringify({
            audits: { "devices.compliance": "daily", "accounts.inactive": "weekly" },
            scheduledBy: { "devices.compliance": actor.id, "accounts.inactive": actor.id },
          }),
          actor.id,
        ],
      ),
    );
    const saved = await database.tenant(tenant, (sql) =>
      writeAuditSchedule(
        sql,
        tenant,
        { audits: { "mfa.registration": "daily" } },
        otherActor,
        builtInAudits,
      ),
    );
    // This build neither shows nor runs it, but it is still there for the build that knows it.
    expect(saved).not.toHaveProperty("devices.compliance");
    const stored = await database.tenant(tenant, (sql) => readStoredAuditSchedule(sql, tenant));
    expect(stored).toEqual({
      audits: {
        "accounts.inactive": "weekly",
        "mfa.registration": "daily",
        "devices.compliance": "daily",
      },
      scheduledBy: {
        "accounts.inactive": actor.id,
        "mfa.registration": otherActor,
        "devices.compliance": actor.id,
      },
    });
  });

  it("reads a damaged value, or an audit that no longer exists, as off", async () => {
    expect(knownAuditSchedule(null, builtInAudits)["accounts.inactive"]).toBe("off");
    expect(knownAuditSchedule({ audits: [] }, builtInAudits)["accounts.inactive"]).toBe("off");
    const mixed = knownAuditSchedule(
      {
        audits: { "accounts.inactive": "hourly", "guests.review": "weekly", "gone.audit": "daily" },
      },
      builtInAudits,
    );
    expect(mixed["accounts.inactive"]).toBe("off");
    expect(mixed["guests.review"]).toBe("weekly");
    expect(mixed).not.toHaveProperty("gone.audit");
    // An inherited name is never read as a saved value.
    expect(
      knownAuditSchedule(
        { audits: Object.create({ "accounts.inactive": "daily" }) },
        builtInAudits,
      )["accounts.inactive"],
    ).toBe("off");
    await database.tenant(tenantId, (sql) =>
      sql.query(
        `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
         VALUES ($1,'audits.schedule','"not an object"'::jsonb,now(),$2)
         ON CONFLICT (tenant_id, key) DO UPDATE SET value=EXCLUDED.value`,
        [tenantId, actor.id],
      ),
    );
    const damaged = await database.tenant(tenantId, (sql) =>
      readAuditSchedule(sql, tenantId, builtInAudits),
    );
    expect(scheduledAuditCount(damaged)).toBe(0);
  });

  it("is due a day or a week after the newest start, never when off", () => {
    const now = at(10, 12);
    const hours = (count: number) => new Date(now.getTime() - count * 3_600_000);
    expect(auditDue("daily", null, now)).toBe(true);
    expect(auditDue("weekly", null, now)).toBe(true);
    expect(auditDue("off", null, now)).toBe(false);
    expect(auditDue("off", hours(24 * 30), now)).toBe(false);
    expect(auditDue("daily", hours(24), now)).toBe(false);
    expect(auditDue("daily", hours(24.01), now)).toBe(true);
    expect(auditDue("weekly", hours(24 * 7 - 1), now)).toBe(false);
    expect(auditDue("weekly", hours(24 * 7 + 1), now)).toBe(true);
    // A start in the future (the clock went back) is not due.
    expect(auditDue("daily", hours(-48), now)).toBe(false);
  });

  it("counts from the newest start of any run, failed and scheduled ones included", async () => {
    expect(
      await database.tenant(tenantId, (sql) => latestRunStart(sql, tenantId, "never.ran")),
    ).toBeNull();
    await run("licenses.reclaim", 4, { status: "succeeded" });
    await run("licenses.reclaim", 6, { status: "failed" }, { trigger: "schedule" });
    expect(
      await database.tenant(tenantId, (sql) => latestRunStart(sql, tenantId, "licenses.reclaim")),
    ).toEqual(at(6));
  });
});

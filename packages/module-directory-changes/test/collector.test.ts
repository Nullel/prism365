import { randomUUID } from "node:crypto";
import { GraphHttpClient } from "@prism365/graph";
import {
  type CollectionCredential,
  latestRuns,
  listFindings,
  ModuleRegistry,
  ModuleScheduler,
} from "@prism365/modules";
import { PGliteModuleDatabase } from "@prism365/storage";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  createDirectoryChangesModule,
  type DirectoryChangesConfig,
  defaultDirectoryChangesConfig,
  MAX_SPAN_MS,
  PAGE_INTERVAL_MS,
} from "../src/index.ts";

// Each scenario opens PGlite inside the test; on a loaded machine that cold start alone can pass
// the default 5 s budget.
vi.setConfig({ testTimeout: 30_000 });

// Synthetic tenant and accounts only.
const tenantId = randomUUID();
const actorId = randomUUID();
const adminId = "0a1b2c3d-0000-4000-8000-00000000a001";
const personId = "0a1b2c3d-0000-4000-8000-00000000b002";
const start = new Date("2026-10-22T12:00:00Z");
const MINUTE = 60_000;
const HOUR = 60 * MINUTE;
const DAY = 24 * HOUR;

interface FakeAudit {
  id?: unknown;
  activityDateTime?: unknown;
  // Milliseconds after the change Microsoft makes it visible; default 0.
  publishDelay?: number;
  [key: string]: unknown;
}

let records: FakeAudit[];
let clock: Date;
// The collector's own clock for its time budget and pacing, in milliseconds. Sleeping moves it,
// so a run's budget is spent by pacing and by the request costs a test sets.
let elapsed: number;
let sleeps: number[];
let requestTimes: number[];
let requests: URL[];
let graphStatus: number | null;
let serverPageSize: number;
let maxTop: number | null;
let onRequest: ((url: URL) => void) | null;
// A moment Microsoft never answers for in time, at any window size.
let slowAt: number | null;
let database: PGliteModuleDatabase;
let scheduler: ModuleScheduler;

function audit(at: Date, overrides: Record<string, unknown> = {}): FakeAudit {
  return {
    id: `Directory_${randomUUID()}`,
    activityDateTime: at.toISOString(),
    activityDisplayName: "Update user",
    category: "UserManagement",
    correlationId: randomUUID(),
    result: "success",
    resultReason: "",
    loggedByService: "Core Directory",
    operationType: "Update",
    initiatedBy: {
      user: {
        id: adminId,
        displayName: "Synthetic Admin",
        userPrincipalName: "admin@example.invalid",
        ipAddress: "192.0.2.10",
      },
      app: null,
    },
    targetResources: [
      {
        id: personId,
        displayName: "Synthetic Person",
        type: "User",
        userPrincipalName: "person@example.invalid",
        modifiedProperties: [
          { displayName: "DisplayName", oldValue: '["Old name"]', newValue: '["New name"]' },
        ],
      },
    ],
    additionalDetails: [{ key: "UserType", value: "Member" }],
    ...overrides,
  };
}

const ago = (milliseconds: number) => new Date(clock.getTime() - milliseconds);

// 100-nanosecond ticks since the epoch for an ISO timestamp with up to seven fraction digits.
function ticks(value: string): bigint {
  const match = /^(.+?)(?:\.(\d+))?Z$/.exec(value);
  if (!match?.[1]) throw new Error(`Unreadable timestamp ${value}`);
  const fraction = (match[2] ?? "").padEnd(7, "0").slice(0, 7);
  return BigInt(Date.parse(`${match[1]}Z`)) * 10_000n + BigInt(fraction);
}

async function fakeFetch(input: RequestInfo | URL): Promise<Response> {
  const url = new URL(String(input));
  requests.push(url);
  requestTimes.push(elapsed);
  onRequest?.(url);
  if (graphStatus)
    return new Response("{}", {
      status: graphStatus,
      headers: { "Content-Type": "application/json" },
    });
  const top = Number(url.searchParams.get("$top") ?? "100");
  if (maxTop !== null && top > maxTop)
    return new Response("{}", { status: 400, headers: { "Content-Type": "application/json" } });
  const match = /^activityDateTime ge (\S+) and activityDateTime le (\S+)$/.exec(
    url.searchParams.get("$filter") ?? "",
  );
  if (!match?.[1] || !match[2]) return new Response("{}", { status: 400 });
  const from = ticks(match[1]);
  const to = ticks(match[2]);
  if (slowAt !== null && Number(from / 10_000n) <= slowAt && Number(to / 10_000n) >= slowAt)
    await new Promise((resolve) => setTimeout(resolve, 300));
  const value = records
    .filter((record) => {
      // A record with no readable time matches every window, as nothing can place it.
      if (typeof record.activityDateTime !== "string" || !/Z$/.test(record.activityDateTime))
        return true;
      const at = ticks(record.activityDateTime);
      const published =
        Date.parse(record.activityDateTime) + (record.publishDelay ?? 0) <= clock.getTime();
      return published && at >= from && at <= to;
    })
    // Newest first, as Graph usually answers; the collector must not depend on it.
    .sort((a, b) => String(b.activityDateTime).localeCompare(String(a.activityDateTime)))
    .map(({ publishDelay: _delay, ...record }) => record);
  const size = Math.min(top, serverPageSize);
  const skip = Number(url.searchParams.get("$skiptoken") ?? "0");
  const next = new URL(url);
  next.searchParams.set("$skiptoken", String(skip + size));
  return Response.json({
    value: value.slice(skip, skip + size),
    ...(skip + size < value.length ? { "@odata.nextLink": next.href } : {}),
  });
}

async function setup(
  overrides: Partial<DirectoryChangesConfig> = {},
  graphOptions: { requestTimeoutMilliseconds?: number } = {},
) {
  const config = { ...defaultDirectoryChangesConfig, ...overrides };
  const registry = new ModuleRegistry().register(
    createDirectoryChangesModule(config, {
      clock: () => elapsed,
      sleep: async (milliseconds) => {
        sleeps.push(milliseconds);
        elapsed += milliseconds;
      },
    }),
  );
  database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  scheduler = new ModuleScheduler({
    enabledPacks: ["audit.directory.read"],
    registry,
    database,
    now: () => clock,
    graphClient: (getAccessToken) =>
      new GraphHttpClient({
        getAccessToken,
        fetchImplementation: fakeFetch,
        sleep: async () => {},
        ...graphOptions,
      }),
  });
  const credential: CollectionCredential = {
    kind: "delegated",
    tenantId,
    actorId,
    grantedPermissions: async () => ["User.Read", "AuditLog.Read.All"],
    accessToken: async () => "synthetic-token",
  };
  scheduler.bind(credential);
}

const runs = () => database.tenant(tenantId, (sql) => latestRuns(sql, tenantId));
// Every run's status, oldest first.
const statuses = async () =>
  (
    await database.tenant(tenantId, (sql) =>
      sql.query<{ status: string }>(
        "SELECT status FROM prism365_module_run WHERE tenant_id = $1 ORDER BY started_at, id",
        [tenantId],
      ),
    )
  ).rows.map((row) => row.status);
const findings = () =>
  database.tenant(tenantId, (sql) => listFindings(sql, tenantId, { module: "directory.changes" }));
interface StoredCursor {
  phase?: string;
  start?: string;
  through?: string;
  span?: number;
  trying?: string;
  strikes?: { at: string; count: number };
}
const cursor = async (): Promise<StoredCursor> =>
  JSON.parse(
    (
      await database.tenant(tenantId, (sql) =>
        sql.query<{ value: string }>(
          "SELECT value FROM prism365_cursor WHERE module='directory.changes'",
        ),
      )
    ).rows[0]?.value ?? "{}",
  ) as StoredCursor;
const setCursor = (value: string) =>
  database.tenant(tenantId, (sql) =>
    sql.query("UPDATE prism365_cursor SET value=$1 WHERE module='directory.changes'", [value]),
  );
const count = async (table = "event", where = "true") =>
  (
    await database.tenant(tenantId, (sql) =>
      sql.query<{ n: number }>(
        `SELECT count(*)::int AS n FROM mod_directory_changes_${table} WHERE ${where}`,
      ),
    )
  ).rows[0]?.n;

// Runs until history is loaded. The scheduler runs a collector again at once when it asks for
// more, so this needs no clock movement.
async function loadHistory() {
  for (let attempt = 0; attempt < 20; attempt += 1) {
    await scheduler.tick();
    if ((await cursor()).phase === "live") return;
  }
  throw new Error("History never finished loading.");
}

// The scheduled run ten minutes after the last, and any runs it asks for to catch up.
async function poll() {
  clock = new Date(clock.getTime() + 10 * MINUTE);
  for (let attempt = 0; attempt < 10; attempt += 1) {
    await scheduler.tick();
    if ((await cursor()).through === clock.toISOString()) return;
  }
  throw new Error("Collection never caught up.");
}

beforeEach(() => {
  clock = start;
  elapsed = 1_000_000;
  sleeps = [];
  requestTimes = [];
  requests = [];
  records = [];
  graphStatus = null;
  serverPageSize = 1_000;
  maxTop = null;
  onRequest = null;
  slowAt = null;
});

afterEach(async () => {
  await scheduler?.stop();
  await database?.close();
});

describe("directory change collector", () => {
  it("loads 30 days of history over several runs, asking for more while it remains", async () => {
    for (let day = 29; day >= 1; day -= 2) records.push(audit(ago(day * DAY + 3 * HOUR)));
    await setup({ runBudgetSeconds: 10 });

    await scheduler.tick();
    const first = await cursor();
    expect(first.phase).toBe("history");
    expect(first.start).toBe(new Date(start.getTime() - 30 * DAY).toISOString());
    const partly = await count();
    expect(partly).toBeGreaterThan(0);
    expect(partly).toBeLessThan(records.length);

    // The clock has not moved and the cadence is ten minutes: only a run that asked for more is
    // due again now.
    await scheduler.tick();
    expect(await count()).toBeGreaterThan(partly ?? 0);

    await loadHistory();
    expect(await count()).toBe(records.length);
    expect(await cursor()).toMatchObject({ phase: "live", through: start.toISOString() });
    // Every request of the history window asked for the documented filter and a page size.
    for (const url of requests) {
      expect(url.pathname).toBe("/v1.0/auditLogs/directoryAudits");
      expect(url.searchParams.get("$top")).toBe("999");
    }
    expect((await runs())[0]?.status).toBe("succeeded");
  });

  it("starts requests at least two seconds apart", async () => {
    // A budget long enough to load the whole history in one run.
    await setup({ runBudgetSeconds: 100 });
    await loadHistory();
    expect(requests.length).toBeGreaterThan(5);
    for (let index = 1; index < requestTimes.length; index += 1)
      expect((requestTimes[index] ?? 0) - (requestTimes[index - 1] ?? 0)).toBeGreaterThanOrEqual(
        PAGE_INTERVAL_MS,
      );
    expect(sleeps.length).toBeGreaterThanOrEqual(requests.length - 1);
  });

  it("picks up records Microsoft publishes late, and stores repeats once", async () => {
    records.push(audit(ago(2 * HOUR)));
    await setup();
    await loadHistory();
    await poll();
    await poll();
    expect(await count()).toBe(1);

    // Happened 40 minutes ago, published only now.
    const late = audit(ago(40 * MINUTE), { publishDelay: 40 * MINUTE + 5 * MINUTE });
    records.push(late);
    await poll();
    expect(await count("event", `id = '${String(late.id)}'`)).toBe(1);

    // Every later run re-reads the margin: the same records come back and change nothing.
    await poll();
    await poll();
    expect(await count()).toBe(2);
    expect(await count("subject")).toBe(4);
    const [run] = await runs();
    expect(run?.status).toBe("succeeded");
    expect(run?.fetched).toBeGreaterThan(0);
    expect(run?.inserted).toBe(0);
  });

  it("does not move past unread time when the budget runs out mid-window", async () => {
    // Ten records in one hour, two to a page: the day they are in needs five pages.
    const busy = new Date(start.getTime() - 5 * DAY);
    const midnight = `${busy.toISOString().slice(0, 10)}T00:00:00.000Z`;
    for (let index = 0; index < 10; index += 1)
      records.push(audit(new Date(busy.getTime() + index * 5 * MINUTE)));
    serverPageSize = 2;
    // Budget enough to reach that day in the first run.
    await setup({ runBudgetSeconds: 100 });
    // The budget runs out while the second page of that day is read.
    let pages = 0;
    onRequest = (url) => {
      if (!(url.searchParams.get("$filter") ?? "").startsWith(`activityDateTime ge ${midnight}`))
        return;
      pages += 1;
      if (pages === 2) elapsed += 60_000;
    };
    for (let attempt = 0; attempt < 10 && pages < 2; attempt += 1) await scheduler.tick();
    expect(pages).toBe(2);
    const stopped = await cursor();
    expect(Date.parse(stopped.through ?? "")).toBeLessThanOrEqual(Date.parse(midnight));
    expect(stopped.phase).toBe("history");
    // What was read before the budget ran out is kept.
    expect(await count()).toBe(4);

    onRequest = null;
    await loadHistory();
    expect(await count()).toBe(10);
  });

  it("splits a window that holds more records than the limit, and keeps every record", async () => {
    for (let index = 0; index < 9; index += 1)
      records.push(audit(new Date(start.getTime() - 3 * DAY + index * 7 * MINUTE)));
    serverPageSize = 2;
    await setup({ windowRecordLimit: 4 });
    await loadHistory();
    expect(await count()).toBe(9);
    expect(await findings()).toEqual([]);
    expect((await runs())[0]?.status).toBe("succeeded");
    // The busy day was read in smaller windows than a day.
    const filters = requests.map((url) => url.searchParams.get("$filter") ?? "");
    const spans = filters.map((filter) => {
      const [, from = "", to = ""] = /ge (\S+) and activityDateTime le (\S+)/.exec(filter) ?? [];
      return Date.parse(to) - Date.parse(from);
    });
    expect(Math.min(...spans)).toBeLessThan(MAX_SPAN_MS / 4);
  });

  it("records a gap when a minute holds more records than one window may read", async () => {
    // Seven changes within one second, 30 minutes ago: inside the late-arrival margin, so later
    // runs read that minute again.
    const minute = Math.floor((start.getTime() - 30 * MINUTE) / MINUTE) * MINUTE;
    for (let index = 0; index < 7; index += 1)
      records.push(audit(new Date(minute + 20_000 + index * 100)));
    records.push(audit(new Date(start.getTime() - DAY)));
    serverPageSize = 2;
    await setup({ windowRecordLimit: 4 });
    await loadHistory();
    const [gap] = await findings();
    expect(gap).toMatchObject({
      ruleId: "collection.gap",
      purpose: "operations",
      severity: "medium",
      subject: { kind: "service", id: "directory.changes" },
      title: "Some directory changes could not be read",
    });
    expect(gap?.summary).toContain("4 were read and kept");
    // What was read is kept, and collection carried on past the gap.
    expect(await count()).toBe(5);
    expect(await cursor()).toMatchObject({ phase: "live" });
    expect(await statuses()).toContain("partial");

    // The settle lane reads that minute again once it is two hours old: still one gap finding.
    for (let index = 0; index < 10; index += 1) await poll();
    const again = await findings();
    expect(again).toHaveLength(1);
    expect(again[0]?.occurrences).toBeGreaterThan(1);
  });

  it("leaves out unreadable records, counts them, and keeps collecting", async () => {
    records.push(audit(ago(3 * DAY)));
    records.push(audit(ago(2 * DAY), { id: "" }));
    records.push(audit(ago(2 * DAY), { activityDateTime: "yesterday" }));
    records.push(audit(ago(DAY), { targetResources: "not a list", initiatedBy: 7 }));
    await setup();
    await loadHistory();
    expect(await count()).toBe(2);
    const [run] = await runs();
    expect(run).toMatchObject({ status: "partial" });
    // The record with no readable date matches every window's fake filter, so it is counted on
    // each read; the one with no id is read once.
    expect(run?.skipped).toBeGreaterThanOrEqual(2);
    expect(await cursor()).toMatchObject({ phase: "live" });
  });

  it("recovers from a cursor it cannot read by loading history again", async () => {
    records.push(audit(ago(4 * DAY)));
    await setup();
    await loadHistory();
    for (const garbage of ["not json", JSON.stringify({ v: 9, phase: "live" }), "{}"]) {
      await setCursor(garbage);
      clock = new Date(clock.getTime() + 10 * MINUTE);
      await loadHistory();
      expect(await cursor()).toMatchObject({ phase: "live", through: clock.toISOString() });
    }
    expect(await count()).toBe(1);
    expect((await runs())[0]?.status).toBe("succeeded");
  });

  it("records an access failure without moving its place", async () => {
    await setup();
    graphStatus = 403;
    await scheduler.tick();
    expect((await runs())[0]).toMatchObject({ status: "failed", errorCategory: "access_denied" });
    const kept = await cursor();
    expect(kept.phase).toBe("history");
    expect(kept.through).toBe(kept.start);
    expect(kept.trying).toBeUndefined();
    graphStatus = null;
    records.push(audit(ago(DAY)));
    clock = new Date(clock.getTime() + 10 * MINUTE);
    await loadHistory();
    expect(await count()).toBe(1);
  });

  it("asks for smaller pages when Microsoft refuses the larger size", async () => {
    records.push(audit(ago(DAY)));
    maxTop = 100;
    await setup();
    await loadHistory();
    expect(await count()).toBe(1);
    expect(requests.some((url) => url.searchParams.get("$top") === "100")).toBe(true);
    expect((await runs())[0]?.status).toBe("succeeded");
  });

  it("treats a window a stopped run was reading as failed: smaller, then a gap", async () => {
    records.push(audit(ago(DAY)));
    await setup();
    await loadHistory();
    const live = await cursor();
    // A run stopped while reading a window: the next one tries a smaller one, without a gap.
    await setCursor(JSON.stringify({ ...live, trying: live.through }));
    await poll();
    expect((await cursor()).span).toBeLessThan(MAX_SPAN_MS);
    expect(await findings()).toEqual([]);

    // A minimum-size window that has failed twice before is given up and recorded.
    const at = new Date(Math.floor(clock.getTime() / MINUTE) * MINUTE - 10 * MINUTE);
    await setCursor(
      JSON.stringify({
        ...(await cursor()),
        span: MINUTE,
        trying: at.toISOString(),
        strikes: { at: at.toISOString(), count: 2 },
      }),
    );
    await poll();
    const gaps = await findings();
    expect(gaps).toHaveLength(1);
    expect(gaps[0]?.summary).toContain("did not answer in time");
    expect((await cursor()).strikes).toBeUndefined();
  });

  it("judges each stored change once, including ones a stopped run never judged", async () => {
    const assignment = audit(ago(3 * HOUR), {
      activityDisplayName: "Add member to role",
      category: "RoleManagement",
      operationType: "Assign",
      targetResources: [
        {
          id: personId,
          displayName: "Synthetic Person",
          type: "User",
          userPrincipalName: "person@example.invalid",
          modifiedProperties: [
            { displayName: "Role.DisplayName", oldValue: null, newValue: '"Global Administrator"' },
            {
              displayName: "Role.TemplateId",
              oldValue: null,
              newValue: '"62e90394-69f5-4237-9190-012177145e10"',
            },
          ],
        },
      ],
    });
    records.push(assignment);
    // Inside the late-arrival margin, so every poll reads it again.
    assignment.activityDateTime = ago(50 * MINUTE).toISOString();
    await setup();
    await loadHistory();
    await poll();
    await poll();
    const [flagged] = await findings();
    expect(flagged).toMatchObject({
      ruleId: "role.privileged-assigned",
      severity: "critical",
      occurrences: 1,
      subject: { kind: "person", id: personId, label: "Synthetic Person" },
    });

    // A change stored by a run that stopped before its rules ran is judged by the next one.
    await database.tenant(tenantId, (sql) =>
      sql.query(
        `INSERT INTO mod_directory_changes_event (tenant_id, id, activity_at, category, activity,
           result, initiated_by_kind, targets, details, processed, collected_at)
         VALUES ($1, 'Directory_pending', $2, 'Policy', 'Delete Conditional Access policy',
           'success', 'unknown', $3::jsonb, '[]'::jsonb, false, $2)`,
        [
          tenantId,
          ago(HOUR).toISOString(),
          JSON.stringify([
            {
              id: "policy-1",
              name: "Require MFA",
              type: "Policy",
              principal: null,
              properties: [],
            },
          ]),
        ],
      ),
    );
    await poll();
    const rules = (await findings()).map((item) => item.ruleId).sort();
    expect(rules).toEqual(["policy.conditional-access-changed", "role.privileged-assigned"]);
    expect(await count("event", "processed = false")).toBe(0);
  });

  it("records changes Microsoft no longer keeps after collection was off too long", async () => {
    await setup();
    await loadHistory();
    clock = new Date(clock.getTime() + 45 * DAY);
    records.push(audit(ago(DAY)));
    for (let attempt = 0; attempt < 10; attempt += 1) {
      await scheduler.tick();
      if ((await cursor()).through === clock.toISOString()) break;
    }
    expect((await cursor()).through).toBe(clock.toISOString());
    const [missed] = await findings();
    expect(missed?.title).toBe("Directory changes were missed while collection was off");
    expect(await count()).toBe(1);
  });

  it("collects a late record on the settle lane even with no time to re-read", async () => {
    records.push(audit(ago(5 * HOUR)));
    // Three seconds: the main and settle lanes get one request each, the best-effort re-read none.
    await setup({ runBudgetSeconds: 3 });
    for (let attempt = 0; attempt < 40 && (await cursor()).phase !== "live"; attempt += 1)
      await scheduler.tick();
    expect((await cursor()).phase).toBe("live");
    // Happened 40 minutes ago, published 45 minutes after it happened.
    const late = audit(ago(40 * MINUTE), { publishDelay: 45 * MINUTE });
    records.push(late);
    await poll();
    // Visible now, but not yet two hours old: no lane has read it again yet.
    expect(await count()).toBe(1);
    for (let index = 0; index < 9; index += 1) await poll();
    expect(await count("event", `id = '${String(late.id)}'`)).toBe(1);
    const settled = await cursor();
    expect(Date.parse((settled as { settle?: { at: string } }).settle?.at ?? "")).toBe(
      clock.getTime() - 120 * MINUTE,
    );
  });

  it("makes a window Microsoft never answers smaller, then records it as a gap and moves on", async () => {
    slowAt = start.getTime() - 3 * DAY + 90 * 1_000;
    records.push(audit(new Date(start.getTime() - 5 * DAY)));
    records.push(audit(new Date(start.getTime() - DAY)));
    await setup({}, { requestTimeoutMilliseconds: 100 });
    for (let attempt = 0; attempt < 40 && (await cursor()).phase !== "live"; attempt += 1) {
      await scheduler.tick();
      clock = new Date(clock.getTime() + 10 * MINUTE);
    }
    expect((await cursor()).phase).toBe("live");
    expect(await count()).toBe(2);
    const [gap] = await findings();
    expect(gap?.summary).toContain("did not answer in time");
    // The window given up is at most a minute long and holds the moment that never answers.
    const from = Date.parse(gap?.windowStart ?? "");
    const to = Date.parse(gap?.windowEnd ?? "");
    expect(from).toBeLessThanOrEqual(slowAt);
    expect(to).toBeGreaterThanOrEqual(slowAt);
    expect(to - from).toBeLessThanOrEqual(MINUTE);
    expect(await statuses()).toContain("failed");
  });

  it("records a finding without names when a directory name trips the language policy", async () => {
    records.push(
      audit(ago(3 * HOUR), {
        activityDisplayName: "Delete Conditional Access policy",
        category: "Policy",
        targetResources: [
          {
            id: "policy-7",
            displayName: "Productivity apps",
            type: "Policy",
            modifiedProperties: [],
          },
        ],
      }),
    );
    await setup();
    await loadHistory();
    const [finding] = await findings();
    expect(finding).toMatchObject({
      ruleId: "policy.conditional-access-changed",
      severity: "high",
      subject: { kind: "resource", id: "policy-7", label: null },
    });
    expect(finding?.summary).toContain("Names were left out");
    expect(JSON.stringify(finding)).not.toContain("Productivity");
  });
});

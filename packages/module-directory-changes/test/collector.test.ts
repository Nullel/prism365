import { randomUUID } from "node:crypto";
import { GraphHttpClient } from "@prism365/graph";
import {
  type CollectionCredential,
  latestRuns,
  listFindings,
  ModuleRegistry,
  ModuleScheduler,
  setFindingState,
} from "@prism365/modules";
import { PGliteModuleDatabase } from "@prism365/storage";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  createDirectoryChangesModule,
  type DirectoryChangesConfig,
  defaultDirectoryChangesConfig,
  HEALED_TITLE,
  MAX_SPAN_MS,
  PAGE_INTERVAL_MS,
  type Pacer,
} from "../src/index.ts";

// Each scenario opens PGlite inside the test; on a loaded machine that cold start alone can pass
// the default 5 s budget.
vi.setConfig({ testTimeout: 60_000 });

// Synthetic tenant and accounts only.
const tenantId = randomUUID();
const actorId = randomUUID();
const adminId = "0a1b2c3d-0000-4000-8000-00000000a001";
const personId = "0a1b2c3d-0000-4000-8000-00000000b002";
const start = new Date("2026-10-22T12:00:00Z");
const MINUTE = 60_000;
const HOUR = 60 * MINUTE;
const DAY = 24 * HOUR;
// The margin most scenarios use; one test checks the deployed default.
const MARGIN_MINUTES = 120;

interface FakeAudit {
  id?: unknown;
  activityDateTime?: unknown;
  // Milliseconds after the change Microsoft makes it visible; default 0.
  publishDelay?: number;
  [key: string]: unknown;
}

interface Sent {
  url: URL;
  from: number;
  to: number;
  // On the collector's clock: when the request started and when its answer was back.
  start: number;
  end: number;
  status: number;
}

let records: FakeAudit[];
let clock: Date;
// The collector's own clock for its time budget and pacing, in milliseconds. Sleeping moves it,
// so a run's budget is spent by pacing and by the request costs a test sets.
let elapsed: number;
let sent: Sent[];
let serverPageSize: number;
// Collector-clock time each request takes.
let requestCost: number;
// Answers a request with this status instead of the log, when it returns one.
let failWith: ((request: { from: number; to: number; url: URL }) => number | null) | null;
// Answers a request with this response before anything else, when it returns one. The window is
// NaN for a request without a time filter (a continuation link).
let answerWith: ((request: { from: number; to: number; url: URL }) => Response | null) | null;
// Answers a request with an empty page instead of the log, when it returns true.
let emptyWhen: ((request: { from: number; to: number }) => boolean) | null;
let onRequest: ((url: URL) => void) | null;
// A moment Microsoft never answers for in time, at any window size.
let slowAt: number | null;
// Every wait the collector asked for between requests.
let slept: number[];
let database: PGliteModuleDatabase;
let scheduler: ModuleScheduler;
let credential: CollectionCredential;

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
const json = (status: number) =>
  new Response("{}", { status, headers: { "Content-Type": "application/json" } });

// 100-nanosecond ticks since the epoch for an ISO timestamp with up to seven fraction digits.
function ticks(value: string): bigint {
  const match = /^(.+?)(?:\.(\d+))?Z$/.exec(value);
  if (!match?.[1]) throw new Error(`Unreadable timestamp ${value}`);
  const fraction = (match[2] ?? "").padEnd(7, "0").slice(0, 7);
  return BigInt(Date.parse(`${match[1]}Z`)) * 10_000n + BigInt(fraction);
}

async function fakeFetch(input: RequestInfo | URL): Promise<Response> {
  const url = new URL(String(input));
  const match = /^activityDateTime ge (\S+) and activityDateTime le (\S+)$/.exec(
    url.searchParams.get("$filter") ?? "",
  );
  const begun = elapsed;
  elapsed += requestCost;
  const log = (status: number) =>
    sent.push({
      url,
      from: match?.[1] ? Date.parse(match[1]) : Number.NaN,
      to: match?.[2] ? Date.parse(match[2]) : Number.NaN,
      start: begun,
      end: elapsed,
      status,
    });
  onRequest?.(url);
  const custom = answerWith?.({
    from: match?.[1] ? Date.parse(match[1]) : Number.NaN,
    to: match?.[2] ? Date.parse(match[2]) : Number.NaN,
    url,
  });
  if (custom) {
    log(custom.status);
    return custom;
  }
  if (!match?.[1] || !match[2]) {
    log(400);
    return json(400);
  }
  const from = ticks(match[1]);
  const to = ticks(match[2]);
  const fromMs = Number(from / 10_000n);
  const toMs = Number(to / 10_000n);
  const status = failWith?.({ from: fromMs, to: toMs, url }) ?? null;
  if (status) {
    log(status);
    return json(status);
  }
  log(200);
  if (slowAt !== null && fromMs <= slowAt && toMs >= slowAt)
    await new Promise((resolve) => setTimeout(resolve, 300));
  if (emptyWhen?.({ from: fromMs, to: toMs })) return Response.json({ value: [] });
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
  const top = Number(url.searchParams.get("$top") ?? "100");
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
  pacer?: Pacer,
) {
  const config = {
    ...defaultDirectoryChangesConfig,
    lateArrivalMinutes: MARGIN_MINUTES,
    ...overrides,
  };
  const registry = new ModuleRegistry().register(
    createDirectoryChangesModule(config, {
      clock: () => elapsed,
      sleep: async (milliseconds) => {
        slept.push(milliseconds);
        elapsed += milliseconds;
      },
      ...(pacer ? { pacer } : {}),
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
  credential = {
    kind: "delegated",
    tenantId,
    actorId,
    grantedPermissions: async () => ["User.Read", "AuditLog.Read.All"],
    accessToken: async () => "synthetic-token",
  };
  scheduler.bind(credential);
}

const runs = () => database.tenant(tenantId, (sql) => latestRuns(sql, tenantId));
// Every run, oldest first.
const allRuns = async () =>
  (
    await database.tenant(tenantId, (sql) =>
      sql.query<{ status: string; error_category: string | null }>(
        "SELECT status, error_category FROM prism365_module_run WHERE tenant_id = $1 ORDER BY started_at, id",
        [tenantId],
      ),
    )
  ).rows;
const findings = () =>
  database.tenant(tenantId, (sql) =>
    listFindings(sql, tenantId, { module: "directory.changes", limit: 200 }),
  );
interface StoredCursor {
  phase?: string;
  start?: string;
  through?: string;
  span?: number;
  hold?: string;
  trying?: string;
  strikes?: { at: string; count: number };
  settle?: { at: string; span?: number; trying?: string };
  missed?: { from: string; to: string };
  unsettled?: { from: string; to: string };
  fresh?: string;
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
const storedIds = async () =>
  new Set(
    (
      await database.tenant(tenantId, (sql) =>
        sql.query<{ id: string }>("SELECT id FROM mod_directory_changes_event"),
      )
    ).rows.map((row) => row.id),
  );

// Runs until history is loaded. The scheduler runs a collector again at once when it asks for
// more, so this needs no clock movement.
async function loadHistory() {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    await scheduler.tick();
    if ((await cursor()).phase === "live") return;
  }
  throw new Error("History never finished loading.");
}

// Time passes for the scheduler and the collector alike.
function wait(milliseconds: number) {
  clock = new Date(clock.getTime() + milliseconds);
  elapsed += milliseconds;
}

// The scheduled run ten minutes after the last, and the runs it asks for to catch up, five
// seconds apart as the scheduler starts them.
async function poll(minutes = 10) {
  wait(minutes * MINUTE);
  for (let attempt = 0; attempt < 20; attempt += 1) {
    await scheduler.tick();
    if ((await cursor()).through === clock.toISOString()) return;
    wait(5_000);
  }
  throw new Error("Collection never caught up.");
}

// Runs, five seconds apart, until the settle lane has read everything old enough.
async function settleUp() {
  for (let attempt = 0; attempt < 60; attempt += 1) {
    const at = Date.parse((await cursor()).settle?.at ?? "");
    if (at >= clock.getTime() - MARGIN_MINUTES * MINUTE) return;
    wait(5_000);
    await scheduler.tick();
  }
  throw new Error("The settle lane never caught up.");
}

// Scheduled runs ten minutes apart, each followed by the runs it asks for, five seconds apart,
// until the main lane is current or a run asks for nothing more. Runs may fail.
async function scheduled(polls: number, after?: () => Promise<void>) {
  for (let index = 0; index < polls; index += 1) {
    wait(10 * MINUTE);
    await scheduler.tick();
    for (let attempt = 0; attempt < 20; attempt += 1) {
      if ((await cursor()).through === clock.toISOString()) break;
      wait(5_000);
      await scheduler.tick();
    }
    await after?.();
  }
}

// Scheduled runs that may fail, ten minutes apart.
async function runFor(polls: number, after?: () => Promise<void>) {
  for (let index = 0; index < polls; index += 1) {
    wait(10 * MINUTE);
    await scheduler.tick();
    await after?.();
  }
}

beforeEach(() => {
  clock = start;
  elapsed = 1_000_000;
  sent = [];
  records = [];
  serverPageSize = 1_000;
  requestCost = 0;
  failWith = null;
  answerWith = null;
  emptyWhen = null;
  slept = [];
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
    for (const request of sent) {
      expect(request.url.pathname).toBe("/v1.0/auditLogs/directoryAudits");
      expect(request.url.searchParams.get("$top")).toBe("999");
    }
    expect((await runs())[0]?.status).toBe("succeeded");
  });

  it("keeps requests 2.2 s apart after each answer, on every lane and across runs", async () => {
    // Pages, late records for the settle lane and the re-read, a refused page size on every run,
    // and runs the scheduler starts again at once.
    serverPageSize = 2;
    requestCost = 300;
    for (let hours = 70; hours >= 1; hours -= 3) records.push(audit(ago(hours * HOUR)));
    failWith = ({ url }) => (Number(url.searchParams.get("$top")) > 100 ? 400 : null);
    await setup({ runBudgetSeconds: 10 });
    await loadHistory();
    for (let index = 0; index < 6; index += 1) {
      for (let k = 0; k < 6; k += 1)
        records.push(
          audit(new Date(clock.getTime() + k * MINUTE), { publishDelay: k * 30 * MINUTE }),
        );
      await poll();
    }
    const ordered = [...sent].sort((a, b) => a.start - b.start);
    expect(ordered.length).toBeGreaterThan(40);
    for (let index = 1; index < ordered.length; index += 1)
      expect((ordered[index]?.start ?? 0) - (ordered[index - 1]?.end ?? 0)).toBeGreaterThanOrEqual(
        2_200,
      );
    expect(PAGE_INTERVAL_MS).toBe(2_200);
    // Never more than five starts in any closed ten seconds.
    for (const request of ordered)
      expect(
        ordered.filter(
          (other) => other.start >= request.start && other.start <= request.start + 10_000,
        ).length,
      ).toBeLessThanOrEqual(5);
    // The settle lane and the re-read both made requests, older than the main lane's.
    expect(sent.some((request) => request.to <= clock.getTime() - MARGIN_MINUTES * MINUTE)).toBe(
      true,
    );
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

    // Later runs read the same records again: they change nothing.
    await poll();
    await poll();
    expect(await count()).toBe(2);
    expect(await count("subject")).toBe(4);
    const [run] = await runs();
    expect(run?.status).toBe("succeeded");
    expect(run?.fetched).toBeGreaterThan(0);
    expect(run?.inserted).toBe(0);
  });

  it("collects a late record on the settle lane even with no time to re-read", async () => {
    records.push(audit(ago(5 * HOUR)));
    // Three seconds: the main and settle lanes get one request each, the best-effort re-read none.
    await setup({ runBudgetSeconds: 3 });
    for (let attempt = 0; attempt < 40 && (await cursor()).phase !== "live"; attempt += 1)
      await scheduler.tick();
    expect((await cursor()).phase).toBe("live");
    // After the history load the settle lane reads all of it again.
    await settleUp();
    // Happened 40 minutes ago, published 45 minutes after it happened.
    const late = audit(ago(40 * MINUTE), { publishDelay: 45 * MINUTE });
    records.push(late);
    await poll();
    // Visible now, but not yet two hours old: no lane has read it again yet.
    expect(await count()).toBe(1);
    for (let index = 0; index < 9; index += 1) await poll();
    expect(await count("event", `id = '${String(late.id)}'`)).toBe(1);
    expect(Date.parse((await cursor()).settle?.at ?? "")).toBe(
      clock.getTime() - MARGIN_MINUTES * MINUTE,
    );
  });

  it("re-reads by default four hours behind, not two", async () => {
    await setup({ lateArrivalMinutes: defaultDirectoryChangesConfig.lateArrivalMinutes });
    await loadHistory();
    for (let index = 0; index < 4; index += 1) await poll();
    expect(defaultDirectoryChangesConfig.lateArrivalMinutes).toBe(240);
    const settled = Date.parse((await cursor()).settle?.at ?? "");
    expect(settled).toBeLessThanOrEqual(clock.getTime() - 240 * MINUTE);
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
      if (pages === 2) elapsed += 200_000;
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
    const spans = sent.map((request) => request.to - request.from);
    expect(Math.min(...spans)).toBeLessThan(MAX_SPAN_MS / 4);
  });

  it("reads every record of a burst inside one minute by splitting below a minute", async () => {
    // Twelve changes in half a second, four allowed per window.
    const minute = Math.floor((start.getTime() - DAY) / MINUTE) * MINUTE;
    for (let index = 0; index < 12; index += 1)
      records.push(audit(new Date(minute + 20_000 + index * 40)));
    serverPageSize = 2;
    await setup({ windowRecordLimit: 4, runBudgetSeconds: 100 });
    await loadHistory();
    expect(await count()).toBe(12);
    expect(await findings()).toEqual([]);
    expect(Math.min(...sent.map((request) => request.to - request.from))).toBeLessThan(1_000);
  });

  it("records a gap only for a millisecond that holds more than one window may read", async () => {
    // Seven changes in the same millisecond, 30 minutes ago: inside the margin, so later runs
    // read that millisecond again.
    const at = new Date(Math.floor((start.getTime() - 30 * MINUTE) / MINUTE) * MINUTE + 20_000);
    for (let index = 0; index < 7; index += 1) records.push(audit(at));
    records.push(audit(new Date(start.getTime() - DAY)));
    serverPageSize = 2;
    await setup({ windowRecordLimit: 4, runBudgetSeconds: 100 });
    await loadHistory();
    const [gap] = await findings();
    expect(gap).toMatchObject({
      ruleId: "collection.gap",
      purpose: "operations",
      severity: "medium",
      subject: { kind: "service", id: "directory.changes" },
      title: "Some directory changes could not be read",
    });
    expect(gap?.summary).toContain("cannot be split further");
    expect(gap?.summary).toContain("4 were read and kept");
    // Both bounds are inclusive, so a record exactly on a millisecond is in the windows on
    // either side of it: at most two milliseconds.
    expect(
      Date.parse(gap?.windowEnd ?? "") - Date.parse(gap?.windowStart ?? ""),
    ).toBeLessThanOrEqual(2);
    expect(await count()).toBe(5);
    expect(await cursor()).toMatchObject({ phase: "live" });
    expect((await allRuns()).map((run) => run.status)).toContain("partial");

    // The settle lane reads that millisecond again once it is two hours old: still one finding.
    for (let index = 0; index < 12; index += 1) await poll();
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

  it("raises no gap while access is refused before anything was read, then loads history", async () => {
    records.push(audit(ago(2 * DAY)));
    await setup();
    failWith = () => 403;
    // Six hours of refused reads at the normal cadence.
    await runFor(36);
    expect((await runs())[0]).toMatchObject({ status: "failed", errorCategory: "access_denied" });
    expect(await findings()).toEqual([]);
    const refused = await cursor();
    expect(refused.phase).toBe("history");
    expect(refused.through).toBe(refused.start);
    expect(refused.trying).toBeUndefined();
    failWith = null;
    await poll();
    await loadHistory();
    expect(await count()).toBe(1);
    expect(await findings()).toEqual([]);
    // History starts from what Microsoft kept when reading first worked, not six hours earlier.
    const begun = Date.parse((await cursor()).start ?? "");
    expect(begun).toBeLessThanOrEqual(clock.getTime() - 30 * DAY);
    expect(begun).toBeGreaterThan(clock.getTime() - 30 * DAY - 2 * MINUTE);
  });

  it("asks for smaller pages when Microsoft refuses the larger size", async () => {
    records.push(audit(ago(DAY)));
    failWith = ({ url }) => (Number(url.searchParams.get("$top")) > 100 ? 400 : null);
    await setup();
    await loadHistory();
    expect(await count()).toBe(1);
    expect(sent.some((request) => request.url.searchParams.get("$top") === "100")).toBe(true);
    expect((await runs())[0]?.status).toBe("succeeded");
  });

  it("treats a window a run was stopped in as failed and tries it smaller", async () => {
    records.push(audit(ago(DAY)));
    await setup();
    await loadHistory();
    // Collection stops while a request is in flight (a sign-out or a restart): nothing the run
    // does afterwards reaches storage.
    let stopping = true;
    onRequest = () => {
      if (!stopping) return;
      stopping = false;
      scheduler.unbind("shutdown");
    };
    clock = new Date(clock.getTime() + 10 * MINUTE);
    await scheduler.tick();
    expect((await cursor()).trying).toBeTruthy();
    onRequest = null;
    scheduler.bind(credential);
    await poll();
    const after = await cursor();
    expect(after.trying).toBeUndefined();
    expect(after.span).toBeLessThan(MAX_SPAN_MS);
    expect(await findings()).toEqual([]);
  });

  it("clears a window's strikes once it is read", async () => {
    records.push(audit(ago(DAY)));
    await setup();
    await loadHistory();
    const live = await cursor();
    const at = new Date(clock.getTime() + 5 * MINUTE).toISOString();
    await setCursor(JSON.stringify({ ...live, strikes: { at, count: 2 } }));
    await poll();
    expect((await cursor()).strikes).toBeUndefined();
  });

  it("narrows a window Microsoft never answers to a millisecond at once and gives it up on the third try, ten minutes on", async () => {
    slowAt = start.getTime() - 3 * DAY + 90 * 1_000;
    const slow = slowAt;
    records.push(audit(new Date(start.getTime() - 5 * DAY)));
    records.push(audit(new Date(start.getTime() - DAY)));
    await setup({}, { requestTimeoutMilliseconds: 100 });
    // The first scheduled run and the runs it asks for narrow the window down and try the
    // millisecond twice; its last try waits for the next scheduled run.
    const seen: number[] = [];
    await scheduled(2, async () => {
      seen.push((await cursor()).strikes?.count ?? 0);
    });
    expect(seen[0]).toBe(2);
    expect((await cursor()).phase).toBe("live");
    expect(await count()).toBe(2);
    const [gap] = await findings();
    expect(gap?.summary).toContain("did not answer in time, on 3 tries");
    const from = Date.parse(gap?.windowStart ?? "");
    const to = Date.parse(gap?.windowEnd ?? "");
    expect(from).toBeLessThanOrEqual(slow);
    expect(to).toBeGreaterThanOrEqual(slow);
    expect(to - from).toBeLessThanOrEqual(2);
    // Each millisecond window around it was tried three times, then given up.
    const tries = new Map<string, number>();
    for (const request of sent)
      if (request.to - request.from <= 1 && request.from <= slow && request.to >= slow) {
        const key = `${request.from}/${request.to}`;
        tries.set(key, (tries.get(key) ?? 0) + 1);
      }
    expect(tries.size).toBeGreaterThan(0);
    for (const count of tries.values()) expect(count).toBe(3);
    // The last try came at least ten minutes after the first.
    const times = sent
      .filter(
        (request) => request.to - request.from <= 1 && request.from <= slow && request.to >= slow,
      )
      .map((request) => request.start);
    expect(Math.max(...times) - Math.min(...times)).toBeGreaterThanOrEqual(10 * MINUTE);
    // Microsoft answered everything else, so no run failed.
    expect((await allRuns()).map((run) => run.status)).not.toContain("failed");
  });

  it("shrinks windows Microsoft answers with a gateway error and catches up in the same poll", async () => {
    for (let days = 25; days >= 1; days -= 3) records.push(audit(ago(days * DAY)));
    await setup();
    await loadHistory();
    // Collection is off for six hours; Microsoft then answers 504 for any window longer than an
    // hour, as Learn advises shrinking the range.
    const off = clock.getTime();
    clock = new Date(off + 6 * HOUR);
    for (let minutes = 30; minutes < 360; minutes += 30)
      records.push(audit(new Date(off + minutes * MINUTE)));
    failWith = ({ from, to }) => (to - from > HOUR ? 504 : null);
    // One scheduled run and the runs it asks for, five seconds apart.
    await poll();
    expect(await storedIds()).toEqual(new Set(records.map((record) => String(record.id))));
    expect(await findings()).toEqual([]);
    expect((await allRuns()).map((run) => run.status)).not.toContain("failed");
    expect(sent.some((request) => request.status === 504)).toBe(true);
  });

  it("gives up a millisecond Microsoft keeps failing for, says why, and says so again once its change is read", async () => {
    const failAt = start.getTime() - 3 * DAY + 7_000;
    for (let days = 5; days >= 1; days -= 2) records.push(audit(ago(days * DAY)));
    // A change in that very millisecond.
    const inside = audit(new Date(failAt));
    records.push(inside);
    let failing = true;
    failWith = ({ from, to }) => (failing && from <= failAt && to >= failAt ? 500 : null);
    await setup();
    await scheduled(2);
    expect((await cursor()).phase).toBe("live");
    expect(await count()).toBe(3);
    const [gap] = await findings();
    expect(gap?.summary).toContain(
      "kept answering with an error or asking Prism365 to wait, on 3 tries",
    );
    expect(
      Date.parse(gap?.windowEnd ?? "") - Date.parse(gap?.windowStart ?? ""),
    ).toBeLessThanOrEqual(2);
    // Microsoft recovers; the settle lane reads the whole span later, with its change, and the
    // finding says so.
    failing = false;
    await poll();
    await settleUp();
    expect(await count("event", `id = '${String(inside.id)}'`)).toBe(1);
    const healed = await findings();
    expect(healed).toHaveLength(1);
    expect(healed[0]?.title).toBe(HEALED_TITLE);
    expect(healed[0]?.summary).toContain("nothing is missing");
  });

  it("leaves a gap open when the later read of it comes back empty", async () => {
    records.push(audit(ago(DAY)));
    await setup();
    await loadHistory();
    await settleUp();
    // Two changes in a millisecond Microsoft fails for while the main lane reads it. Once the
    // main lane has given it up, reads of it come back empty, as Learn says this API can answer
    // under load.
    const bad = clock.getTime() + 3 * MINUTE + 456;
    const lost = [audit(new Date(bad)), audit(new Date(bad))];
    records.push(...lost);
    let passed = false;
    failWith = ({ from, to }) => (!passed && from <= bad && to >= bad ? 500 : null);
    emptyWhen = ({ from, to }) => passed && from <= bad && to >= bad;
    await scheduled(2);
    passed = Date.parse((await cursor()).through ?? "") > bad;
    expect(passed).toBe(true);
    for (let index = 0; index < 14; index += 1) await poll();
    expect(Date.parse((await cursor()).settle?.at ?? "")).toBeGreaterThan(bad);
    for (const change of lost) expect(await count("event", `id = '${String(change.id)}'`)).toBe(0);
    const [gap] = await findings();
    expect(gap?.title).toBe("Some directory changes could not be read");
    expect(gap?.state).toBe("open");
  });

  it("heals a gap only when the settle lane reads it, not when the early re-read does", async () => {
    records.push(audit(ago(DAY)));
    await setup();
    await loadHistory();
    await settleUp();
    const failAt = clock.getTime() + 3 * MINUTE + 456;
    const inside = audit(new Date(failAt));
    records.push(inside);
    let failing = true;
    failWith = ({ from, to }) => (failing && from <= failAt && to >= failAt ? 500 : null);
    await scheduled(2);
    expect(Date.parse((await cursor()).through ?? "")).toBeGreaterThan(failAt);
    expect((await findings())[0]?.title).toBe("Some directory changes could not be read");
    // Microsoft recovers. The best-effort re-read of recent changes reads that span again and
    // stores the change, but only the settle lane's read says nothing is missing.
    failing = false;
    await poll();
    expect(await count("event", `id = '${String(inside.id)}'`)).toBe(1);
    expect(Date.parse((await cursor()).settle?.at ?? "")).toBeLessThan(failAt);
    expect((await findings())[0]?.title).toBe("Some directory changes could not be read");
    for (let index = 0; index < 14; index += 1) await poll();
    expect((await findings())[0]?.title).toBe(HEALED_TITLE);
  });

  it("changes nothing and leaves no gap while Microsoft fails every request, then catches up", async () => {
    for (let days = 5; days >= 1; days -= 2) records.push(audit(ago(days * DAY)));
    await setup();
    await loadHistory();
    await settleUp();
    const before = await cursor();
    const off = clock.getTime();
    for (let minutes = 20; minutes < 6 * 60; minutes += 40)
      records.push(audit(new Date(off + minutes * MINUTE)));
    // An outage: every request fails, whatever its window.
    failWith = () => 500;
    const sentBefore = sent.length;
    await runFor(36);
    expect(await findings()).toEqual([]);
    const held = await cursor();
    expect(held.through).toBe(before.through);
    expect(held.settle?.at).toBe(before.settle?.at);
    expect(held.span).toBeUndefined();
    expect(held.strikes).toBeUndefined();
    // Each run tried one window and checked whether Microsoft answers anything, then waited.
    expect(sent.length - sentBefore).toBeLessThanOrEqual(36 * 2);
    const outage = (await allRuns()).slice(-36);
    expect(outage.every((run) => run.status === "failed")).toBe(true);
    expect(outage.map((run) => run.error_category)).toContain("provider_unavailable");
    failWith = null;
    await poll();
    expect(await storedIds()).toEqual(new Set(records.map((record) => String(record.id))));
    expect(await findings()).toEqual([]);
  });

  it("counts a strike against a millisecond only while Microsoft still answers other requests", async () => {
    await setup();
    await loadHistory();
    await settleUp();
    // The main lane is reading millisecond windows again after a failure, and Microsoft stops
    // answering anything after the first request of the next run (throttling, say).
    const live = await cursor();
    const through = Date.parse(live.through ?? "");
    await setCursor(
      JSON.stringify({ ...live, span: 1, hold: new Date(through + 2).toISOString() }),
    );
    let answered = 0;
    failWith = () => (answered++ < 1 ? null : 500);
    await runFor(1);
    expect(await findings()).toEqual([]);
    expect((await cursor()).strikes?.count ?? 0).toBeLessThanOrEqual(1);
    expect((await allRuns()).at(-1)?.status).toBe("failed");
  });

  it("does not give up a millisecond for a burst of failures shorter than ten minutes", async () => {
    await setup();
    await loadHistory();
    await settleUp();
    const live = await cursor();
    const through = Date.parse(live.through ?? "");
    await setCursor(
      JSON.stringify({ ...live, span: 1, hold: new Date(through + 2).toISOString() }),
    );
    // A change in the next millisecond, which Microsoft fails for during a short burst.
    const target = through + 1;
    const change = audit(new Date(target));
    records.push(change);
    let burst = true;
    failWith = ({ from, to }) => (burst && from <= target && to >= target ? 500 : null);
    await runFor(1);
    expect((await cursor()).strikes?.count).toBe(2);
    burst = false;
    await poll();
    expect(await findings()).toEqual([]);
    expect(await count("event", `id = '${String(change.id)}'`)).toBe(1);
  });

  it.each([
    ["too large at both page sizes", "could not be read (too large"],
    ["refused at both page sizes", "refused the request"],
    ["paging in a loop", "could not be read (too large, malformed or paging in a loop)"],
  ])("gives up a millisecond whose answer is %s, and moves on", async (kind, why) => {
    const poison = start.getTime() - 3 * DAY + 90_000 + 456;
    for (let days = 5; days >= 1; days -= 2) records.push(audit(ago(days * DAY)));
    const loop = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?$skiptoken=loop";
    const covers = ({ from, to }: { from: number; to: number }) => from <= poison && to >= poison;
    if (kind === "refused at both page sizes")
      failWith = (request) => (covers(request) ? 400 : null);
    else if (kind === "too large at both page sizes")
      answerWith = (request) =>
        covers(request)
          ? new Response("{}", {
              status: 200,
              headers: { "Content-Type": "application/json", "Content-Length": "99999999" },
            })
          : null;
    else
      answerWith = (request) =>
        covers(request) || request.url.searchParams.get("$skiptoken") === "loop"
          ? Response.json({ value: [], "@odata.nextLink": loop })
          : null;
    await setup();
    await scheduled(2);
    expect((await cursor()).phase).toBe("live");
    expect(await count()).toBe(3);
    const [gap] = await findings();
    expect(gap?.summary).toContain(why);
    expect(Date.parse(gap?.windowStart ?? "")).toBeLessThanOrEqual(poison);
    expect(Date.parse(gap?.windowEnd ?? "")).toBeGreaterThanOrEqual(poison);
    expect(
      Date.parse(gap?.windowEnd ?? "") - Date.parse(gap?.windowStart ?? ""),
    ).toBeLessThanOrEqual(2);
    expect((await allRuns()).map((run) => run.status)).not.toContain("failed");
  });

  it("loses nothing to empty answers during the history load or a catch-up", async () => {
    for (let hours = 29 * 24; hours >= 1; hours -= 7) records.push(audit(ago(hours * HOUR)));
    // Under load Microsoft answers empty: the first request that would return a record comes
    // back with nothing, later ones with the record.
    const hidden = new Set<unknown>();
    emptyWhen = ({ from, to }) => {
      const fresh = records.filter((record) => {
        const at = Date.parse(String(record.activityDateTime));
        return at >= from && at <= to && !hidden.has(record.id);
      });
      for (const record of fresh) hidden.add(record.id);
      return fresh.length > 0;
    };
    await setup({ runBudgetSeconds: 100 });
    await loadHistory();
    // Off for eight hours, with changes meanwhile.
    const off = clock.getTime();
    for (let minutes = 15; minutes < 8 * 60; minutes += 15)
      records.push(audit(new Date(off + minutes * MINUTE)));
    clock = new Date(off + 8 * HOUR);
    for (let index = 0; index < 30; index += 1) await poll();
    expect(await storedIds()).toEqual(new Set(records.map((record) => String(record.id))));
  });

  it("re-reads what it may have skipped after the clock ran ahead and came back", async () => {
    await setup();
    await loadHistory();
    const live = await cursor();
    // The lanes ran two days ahead while the clock was wrong.
    const ahead = clock.getTime() + 2 * DAY;
    await setCursor(
      JSON.stringify({
        ...live,
        through: new Date(ahead).toISOString(),
        settle: { at: new Date(ahead - MARGIN_MINUTES * MINUTE).toISOString() },
      }),
    );
    const meanwhile = audit(new Date(clock.getTime() + 5 * MINUTE));
    records.push(meanwhile);
    await poll();
    for (let index = 0; index < 3; index += 1) await poll();
    expect(await count("event", `id = '${String(meanwhile.id)}'`)).toBe(1);
  });

  it("raises one widening gap for a span Microsoft dropped before it was read, and one for what was not read again", async () => {
    records.push(audit(ago(DAY)));
    await setup();
    await loadHistory();
    clock = new Date(clock.getTime() + 45 * DAY);
    records.push(audit(ago(DAY)));
    // Reads keep being refused for a while, then work.
    failWith = () => 403;
    await runFor(6);
    const titles = async () => (await findings()).map((finding) => finding.title).sort();
    // The main lane never read the span; the settle lane never read again what the main lane
    // read in the last two hours before collection stopped. Neither moves on silently.
    expect(await titles()).toEqual([
      "Directory changes were missed while they could not be read",
      "Late directory changes may be missing",
    ]);
    const missed = (await findings()).find((finding) => finding.title.startsWith("Directory"));
    failWith = null;
    await poll();
    for (let index = 0; index < 5; index += 1) await poll();
    expect(await titles()).toEqual([
      "Directory changes were missed while they could not be read",
      "Late directory changes may be missing",
    ]);
    const widened = (await findings()).find((finding) => finding.id === missed?.id);
    expect(Date.parse(widened?.windowEnd ?? "")).toBeGreaterThan(
      Date.parse(missed?.windowEnd ?? ""),
    );
    expect(await count()).toBe(2);
  });

  it("keeps one missed-span finding while a lane recovering at the retention edge falls behind it again", async () => {
    await setup();
    await loadHistory();
    const now = clock.getTime();
    for (let minutes = 0; minutes < 600; minutes += 7)
      records.push(audit(new Date(now - 30 * DAY + minutes * MINUTE)));
    // The main lane is five minutes behind what Microsoft keeps, reading millisecond windows
    // after a long run of failures.
    const through = now - 30 * DAY - 5 * MINUTE;
    const iso = (at: number) => new Date(at).toISOString();
    await setCursor(
      JSON.stringify({
        v: 1,
        phase: "live",
        start: iso(now - 40 * DAY),
        through: iso(through),
        span: 1,
        hold: iso(through + 2),
        settle: { at: iso(through - 4 * HOUR), span: 1, hold: iso(through - 4 * HOUR + 2) },
        fresh: iso(now - 40 * DAY),
      }),
    );
    // Runs six seconds apart: each reads a little, and the edge moves on meanwhile.
    for (let index = 0; index < 12; index += 1) {
      wait(6_000);
      await scheduler.tick();
    }
    const missed = (await findings()).filter(
      (finding) => finding.title === "Directory changes were missed while they could not be read",
    );
    expect(missed).toHaveLength(1);
    expect(missed[0]?.occurrences).toBeGreaterThan(1);
    expect(
      (await findings()).filter(
        (finding) => finding.title === "Late directory changes may be missing",
      ),
    ).toHaveLength(1);
  });

  it("waits at most one interval before a request after the clock stepped back", async () => {
    // The last request finished an hour ahead of the collector's clock: the clock has since
    // stepped back an hour.
    await setup({}, {}, { last: elapsed + HOUR });
    await scheduler.tick();
    expect(slept.length).toBeGreaterThan(0);
    expect(Math.max(...slept)).toBeLessThanOrEqual(PAGE_INTERVAL_MS);
  });

  it("keeps a permanent assignment made through PIM out of an activation someone acknowledged", async () => {
    await setup();
    await loadHistory();
    const target = {
      id: personId,
      displayName: "Synthetic Person",
      type: "User",
      userPrincipalName: "person@example.invalid",
      modifiedProperties: [
        {
          displayName: "Role.TemplateId",
          oldValue: null,
          newValue: '"62e90394-69f5-4237-9190-012177145e10"',
        },
      ],
    };
    const msPim = {
      user: null,
      app: {
        appId: "0a1b2c3d-0000-4000-8000-0000000000f2",
        displayName: "MS-PIM",
        servicePrincipalId: "0a1b2c3d-0000-4000-8000-0000000000f1",
        servicePrincipalName: "MS-PIM",
      },
    };
    const byPerson = {
      user: {
        id: personId,
        displayName: "Synthetic Person",
        userPrincipalName: "person@example.invalid",
        ipAddress: "192.0.2.5",
      },
      app: null,
    };
    // Morning: the person activates an eligible Global Administrator role. PIM records it, and
    // the directory records the role addition PIM made, under the same correlation id.
    const morning = clock.getTime() + 5 * MINUTE;
    const activationId = randomUUID();
    records.push(
      audit(new Date(morning), {
        activityDisplayName: "Add member to role completed (PIM activation)",
        category: "RoleManagement",
        loggedByService: "PIM",
        correlationId: activationId,
        initiatedBy: byPerson,
        targetResources: [target],
      }),
      audit(new Date(morning + 800), {
        activityDisplayName: "Add member to role",
        category: "RoleManagement",
        correlationId: activationId,
        initiatedBy: msPim,
        targetResources: [target],
      }),
    );
    await poll();
    const activated = await findings();
    expect(activated).toHaveLength(1);
    expect(activated[0]).toMatchObject({
      title: "Global Administrator activated",
      severity: "critical",
      occurrences: 2,
    });
    const activationFinding = String(activated[0]?.id);
    await database.tenant(tenantId, (sql) =>
      setFindingState(sql, tenantId, activationFinding, "acknowledged", actorId),
    );
    // Afternoon: another administrator gives the same person a permanent active assignment in
    // PIM, which the directory records the same way.
    const afternoon = clock.getTime() + 2 * HOUR;
    const assignmentId = randomUUID();
    records.push(
      audit(new Date(afternoon), {
        activityDisplayName: "Add member to role in PIM completed (permanent)",
        category: "RoleManagement",
        loggedByService: "PIM",
        correlationId: assignmentId,
        targetResources: [target],
      }),
      audit(new Date(afternoon + 900), {
        activityDisplayName: "Add member to role",
        category: "RoleManagement",
        correlationId: assignmentId,
        initiatedBy: msPim,
        targetResources: [target],
      }),
    );
    for (let index = 0; index < 14; index += 1) await poll();
    const role = (await findings()).filter(
      (finding) => finding.ruleId === "role.privileged-assigned",
    );
    const raised = role.filter((finding) => finding.state === "open");
    expect(raised).toHaveLength(2);
    for (const finding of raised)
      expect(finding).toMatchObject({
        severity: "critical",
        title: "Global Administrator assigned",
      });
    expect(role.find((finding) => finding.id === activationFinding)).toMatchObject({
      state: "acknowledged",
      occurrences: 2,
    });
  });

  it("gives the settle lane a third of the run while the main lane has a backlog", async () => {
    serverPageSize = 2;
    await setup();
    await loadHistory();
    for (let index = 0; index < 3; index += 1) await poll();
    // Off for twelve hours with a dense stretch: the main lane needs many pages.
    const off = clock.getTime();
    for (let index = 0; index < 120; index += 1)
      records.push(audit(new Date(off + 10 * MINUTE + index * 5 * MINUTE)));
    clock = new Date(off + 12 * HOUR);
    const before = sent.length;
    await scheduler.tick();
    const thisRun = sent.slice(before);
    // Some requests of the run were the settle lane's, behind where the main lane started.
    expect(thisRun.some((request) => request.to <= off)).toBe(true);
    expect((await cursor()).through).not.toBe(clock.toISOString());
  });

  it("judges each stored change once, including ones a stopped run never judged", async () => {
    const assignment = audit(ago(50 * MINUTE), {
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
           result, initiated_by_kind, targets, details, search_text, processed, collected_at)
         VALUES ($1, 'Directory_pending', $2, 'Policy', 'Delete Conditional Access policy',
           'success', 'unknown', $3::jsonb, '[]'::jsonb, '', false, $2)`,
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

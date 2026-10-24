import { randomUUID } from "node:crypto";
import { GraphHttpClient } from "@prism365/graph";
import {
  type CollectionCredential,
  latestRuns,
  listFindings,
  ModuleRegistry,
  ModuleScheduler,
  type PrismModule,
  setFindingState,
} from "@prism365/modules";
import { PGliteModuleDatabase } from "@prism365/storage";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  createSigninsAuditModule,
  defaultSigninsAuditConfig,
  type SigninsAuditConfig,
  summarizeSignins,
  writeCanaries,
} from "../src/index.ts";

// Each scenario opens PGlite inside the test; on a loaded machine that cold start alone can pass
// the default 5 s budget.
vi.setConfig({ testTimeout: 30_000 });

// Synthetic tenant and accounts; addresses are documentation ranges.
const tenantId = randomUUID();
const actorId = randomUUID();
const start = new Date("2026-10-22T12:00:00Z");

interface FakeSignIn {
  id: string;
  createdDateTime: string;
  [key: string]: unknown;
}

let records: FakeSignIn[];
let graphStatus: number | null;
let clock: Date;
let database: PGliteModuleDatabase;
let scheduler: ModuleScheduler;

const chicago = {
  city: "Chicago",
  state: "Illinois",
  countryOrRegion: "US",
  geoCoordinates: { latitude: 41.88, longitude: -87.63 },
};
const amsterdam = {
  city: "Amsterdam",
  state: "Noord-Holland",
  countryOrRegion: "NL",
  geoCoordinates: { latitude: 52.37, longitude: 4.9 },
};
const paris = {
  city: "Paris",
  state: "Ile-de-France",
  countryOrRegion: "FR",
  geoCoordinates: { latitude: 48.86, longitude: 2.35 },
};

function signIn(at: Date, overrides: Record<string, unknown> = {}): FakeSignIn {
  return {
    id: randomUUID(),
    createdDateTime: at.toISOString(),
    userId: "user-1",
    userPrincipalName: "Synthetic.User@example.invalid",
    userDisplayName: "Synthetic User",
    appDisplayName: "Office 365 Exchange Online",
    ipAddress: "192.0.2.10",
    clientAppUsed: "Browser",
    conditionalAccessStatus: "success",
    isInteractive: true,
    riskLevelDuringSignIn: "hidden",
    status: { errorCode: 0, failureReason: null },
    deviceDetail: { deviceId: "device-1", operatingSystem: "Windows", browser: "Edge" },
    location: chicago,
    ...overrides,
  };
}

const minutesAgo = (minutes: number) => new Date(clock.getTime() - minutes * 60_000);

async function fakeFetch(input: RequestInfo | URL): Promise<Response> {
  if (graphStatus)
    return new Response("{}", {
      status: graphStatus,
      headers: { "Content-Type": "application/json" },
    });
  const url = new URL(String(input));
  const filter = url.searchParams.get("$filter") ?? "";
  const match = /createdDateTime ge (\S+) and createdDateTime (le|lt) (\S+)/.exec(filter);
  if (!match?.[1] || !match[3]) return new Response("{}", { status: 400 });
  // Microsoft filters and sorts at 100-nanosecond precision, finer than a JavaScript date.
  const from = ticks(match[1]);
  const before = match[2] === "lt";
  const to = ticks(match[3]);
  // Simulates Microsoft failing for every read that starts before a moment.
  if (failBefore !== null && Number(from / 10_000n) < failBefore)
    return new Response("{}", { status: 500, headers: { "Content-Type": "application/json" } });
  requested.push({ from: match[1], to: match[3] });
  // Simulates Microsoft answering wide ranges too slowly.
  if (slowFrom !== null && Number((to - from) / 10_000n) > slowFrom)
    await new Promise((resolve) => setTimeout(resolve, 400));

  // Simulates one moment Microsoft can never answer in time, at any window size.
  if (slowAt !== null && Number(from / 10_000n) <= slowAt && Number(to / 10_000n) >= slowAt)
    await new Promise((resolve) => setTimeout(resolve, 400));
  const value = records
    .filter((record) => {
      const at = ticks(record.createdDateTime);
      // A record is visible once Microsoft has published it, publishDelay after it happened.
      const delay = typeof record.publishDelay === "number" ? record.publishDelay : publishDelay;
      const published = Date.parse(record.createdDateTime) + delay <= clock.getTime();
      return published && at >= from && (before ? at < to : at <= to);
    })
    .sort((a, b) => {
      const difference = ticks(b.createdDateTime) - ticks(a.createdDateTime);
      return difference > 0n ? 1 : difference < 0n ? -1 : 0;
    });
  // Pages of $top records, continued with $skiptoken, as Graph pages sign-ins.
  const top = Number(url.searchParams.get("$top") ?? "100");
  const skip = Number(url.searchParams.get("$skiptoken") ?? "0");
  if (requestDelay > 0 && (delayEmptyPages || value.length > 0))
    await new Promise((resolve) => setTimeout(resolve, requestDelay));
  await onRequest?.();
  const next = new URL(url);
  next.searchParams.set("$skiptoken", String(skip + top));
  return Response.json({
    value: value.slice(skip, skip + top),
    ...(skip + top < value.length ? { "@odata.nextLink": next.href } : {}),
  });
}

// 100-nanosecond ticks since the epoch for an ISO timestamp with up to seven fraction digits.
function ticks(value: string): bigint {
  const match = /^(.+?)(?:\.(\d+))?Z$/.exec(value);
  if (!match?.[1]) throw new Error(`Unreadable timestamp ${value}`);
  const fraction = (match[2] ?? "").padEnd(7, "0").slice(0, 7);
  return BigInt(Date.parse(`${match[1]}Z`)) * 10_000n + BigInt(fraction);
}

let slowFrom: number | null = null;
let requestDelay = 0;
let slowAt: number | null = null;
let publishDelay = 0;
let requested: { from: string; to: string }[] = [];
let onRequest: (() => Promise<void>) | null = null;
let delayEmptyPages = true;
let failBefore: number | null = null;

async function setup(
  overrides: Partial<SigninsAuditConfig> = {},
  graphOptions: { requestTimeoutMilliseconds?: number } = {},
  decorate: (module: PrismModule) => PrismModule = (module) => module,
) {
  const config: SigninsAuditConfig = {
    ...defaultSigninsAuditConfig,
    backfillDays: 3,
    ...overrides,
  };
  const registry = new ModuleRegistry().register(decorate(createSigninsAuditModule(config)));
  database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  scheduler = new ModuleScheduler({
    enabledPacks: registry.list().flatMap((module) => [...module.packs]),
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

async function poll() {
  await scheduler.tick();
  clock = new Date(clock.getTime() + 5 * 60_000);
}

const findings = () =>
  database.tenant(tenantId, (sql) => listFindings(sql, tenantId, { module: "signins.audit" }));
const runs = () => database.tenant(tenantId, (sql) => latestRuns(sql, tenantId));
const cursor = async () =>
  JSON.parse(
    (
      await database.tenant(tenantId, (sql) =>
        sql.query<{ value: string }>(
          "SELECT value FROM prism365_cursor WHERE module='signins.audit'",
        ),
      )
    ).rows[0]?.value ?? "{}",
  ) as { phase?: string; span?: number };
const setCursor = (value: Record<string, unknown>) =>
  database.tenant(tenantId, (sql) =>
    sql.query("UPDATE prism365_cursor SET value=$1 WHERE module='signins.audit'", [
      JSON.stringify(value),
    ]),
  );
const storedIds = async () =>
  new Set(
    (
      await database.tenant(tenantId, (sql) =>
        sql.query<{ id: string }>("SELECT id FROM mod_signins_audit_event"),
      )
    ).rows.map((row) => row.id),
  );
const count = async (where = "true") =>
  (
    await database.tenant(tenantId, (sql) =>
      sql.query<{ n: number }>(
        `SELECT count(*)::int AS n FROM mod_signins_audit_event WHERE ${where}`,
      ),
    )
  ).rows[0]?.n;

beforeEach(() => {
  clock = start;
  records = [];
  graphStatus = null;
  slowFrom = null;
  requestDelay = 0;
  slowAt = null;
  publishDelay = 0;
  requested = [];
  onRequest = null;
  delayEmptyPages = true;
  failBefore = null;
  // Three days of ordinary history: one account, one office, one device, every four hours.
  for (let hours = 70; hours >= 2; hours -= 4)
    records.push(signIn(new Date(start.getTime() - hours * 3_600_000)));
});

afterEach(async () => {
  await scheduler?.stop();
  await database?.close();
});

describe("sign-in audit collector", () => {
  it("learns history on the first run without flagging it, then collects live", async () => {
    await setup();
    await poll();
    expect(await count()).toBe(18);
    expect(await count("processed = false")).toBe(0);
    expect(await findings()).toEqual([]);
    const cursor = await database.tenant(tenantId, (sql) =>
      sql.query<{ value: string }>(
        "SELECT value FROM prism365_cursor WHERE module='signins.audit'",
      ),
    );
    expect(JSON.parse(cursor.rows[0]?.value ?? "{}")).toMatchObject({ phase: "live" });
    // Graph's date filter is inclusive at both ends, so boundary records come back twice and are
    // de-duplicated on insert: more fetched than inserted is expected.
    const [run] = await runs();
    expect(run).toMatchObject({ status: "succeeded", inserted: 18 });
    expect(run?.fetched).toBeGreaterThanOrEqual(18);
  });

  it("flags an unfamiliar sign-in with impossible travel and keeps it unfamiliar until resolved", async () => {
    await setup();
    await poll();
    records.push(
      signIn(minutesAgo(2), {
        ipAddress: "203.0.113.47",
        location: amsterdam,
        deviceDetail: { deviceId: "", operatingSystem: "Linux", browser: "Chrome" },
      }),
    );
    await poll();
    const [flagged] = await findings();
    expect(flagged).toMatchObject({
      ruleId: "signins.unfamiliar_signin",
      severity: "critical",
      subject: { kind: "person", id: "user-1", label: "Synthetic User" },
      occurrences: 1,
    });
    expect(flagged?.summary).toContain("Amsterdam");
    const evidence = (flagged?.evidence ?? []) as { fields: Record<string, unknown> }[];
    const fields = evidence[0]?.fields;
    expect(fields).toMatchObject({ network: "203.0.113.0/24", country: "NL" });
    expect(JSON.stringify(flagged)).not.toContain("203.0.113.47");

    // Same place again before anyone looked: still unfamiliar, same finding.
    records.push(
      signIn(minutesAgo(2), {
        ipAddress: "203.0.113.47",
        location: amsterdam,
        deviceDetail: { deviceId: "", operatingSystem: "Linux", browser: "Chrome" },
      }),
    );
    await poll();
    expect((await findings()).map((item) => item.occurrences)).toEqual([2]);

    // An administrator resolves it (a real trip): the place becomes normal for this account.
    await database.tenant(tenantId, (sql) =>
      setFindingState(sql, tenantId, flagged?.id ?? "", "resolved", actorId),
    );
    records.push(
      signIn(minutesAgo(2), {
        ipAddress: "203.0.113.47",
        location: amsterdam,
        deviceDetail: { deviceId: "", operatingSystem: "Linux", browser: "Chrome" },
      }),
    );
    await poll();
    expect(await findings()).toMatchObject([{ state: "resolved", occurrences: 2 }]);
  });

  it("does not flag a new country from a device the account already uses", async () => {
    await setup();
    await poll();
    records.push(signIn(minutesAgo(2), { ipAddress: "198.51.100.30", location: paris }));
    await poll();
    expect(await findings()).toEqual([]);
  });

  it("flags a failure burst and a password spray, and ignores sign-in interrupts", async () => {
    await setup();
    await poll();
    for (let minute = 10; minute >= 4; minute -= 1)
      records.push(
        signIn(minutesAgo(minute), {
          userId: "user-2",
          userDisplayName: "Second Synthetic",
          ipAddress: "203.0.113.9",
          status: { errorCode: 50126 },
          location: paris,
        }),
      );
    for (const user of ["a", "b", "c", "d", "e"])
      records.push(
        signIn(minutesAgo(30), {
          userId: `user-${user}`,
          ipAddress: "198.51.100.7",
          status: { errorCode: 50126 },
          location: paris,
        }),
      );
    for (let minute = 12; minute >= 3; minute -= 1)
      records.push(
        signIn(minutesAgo(minute), { ipAddress: "192.0.2.77", status: { errorCode: 50074 } }),
      );
    await poll();
    const rules = (await findings()).map((item) => `${item.ruleId}:${item.subject.id}`).sort();
    expect(rules).toEqual([
      "signins.failed_burst_account:user-2",
      "signins.failed_burst_ip:198.51.100.7",
      "signins.failed_burst_ip:203.0.113.9",
      "signins.password_spray:198.51.100.7",
    ]);
  });

  it("picks up records that Microsoft publishes late", async () => {
    await setup();
    await poll();
    await poll();
    records.push(signIn(minutesAgo(40), { ipAddress: "192.0.2.11" }));
    await poll();
    expect(await count("ip = '192.0.2.11'")).toBe(1);
  });

  it("splits a time window that returns more records than the cap", async () => {
    records = [];
    for (let index = 0; index < 7; index += 1)
      records.push(signIn(new Date(start.getTime() - (20 + index * 10) * 60_000)));
    await setup({ backfillDays: 1, windowRecordLimit: 2 });
    await poll();
    expect(await count()).toBe(7);
    expect((await runs())[0]?.status).toBe("succeeded");
  });

  it("records an access failure without losing its place", async () => {
    await setup();
    graphStatus = 403;
    await poll();
    expect((await runs())[0]).toMatchObject({ status: "failed", errorCategory: "access_denied" });
    expect(await count()).toBe(0);
    graphStatus = null;
    clock = new Date(clock.getTime() + 5 * 60_000);
    await poll();
    expect(await count()).toBe(18);
  });

  it("leaves out an unreadable record and keeps collecting past it", async () => {
    records.push({ ...signIn(new Date(start.getTime() - 3 * 3_600_000)), id: "" });
    await setup();
    await poll();
    expect(await count()).toBe(18);
    expect((await runs())[0]).toMatchObject({ status: "partial", skipped: 1 });
    const cursor = await database.tenant(tenantId, (sql) =>
      sql.query<{ value: string }>(
        "SELECT value FROM prism365_cursor WHERE module='signins.audit'",
      ),
    );
    expect(JSON.parse(cursor.rows[0]?.value ?? "{}")).toMatchObject({ phase: "live" });
  });

  it("collects every record of a dense burst that exceeds the per-request cap", async () => {
    records = [];
    const burst = new Date(start.getTime() - 3 * 3_600_000);
    for (let index = 0; index < 300; index += 1)
      records.push(
        signIn(new Date(burst.getTime() + index * 400), {
          userId: `user-${index % 40}`,
          ipAddress: "203.0.113.66",
          status: { errorCode: 50126 },
        }),
      );
    await setup({ backfillDays: 1, windowRecordLimit: 100 });
    await poll();
    expect(await count()).toBe(300);
    expect((await runs())[0]?.status).toBe("succeeded");
  });

  it("splits a range Microsoft answers too slowly instead of failing the run", async () => {
    slowFrom = 30 * 60_000;
    await setup({ backfillDays: 1 }, { requestTimeoutMilliseconds: 150 });
    await poll();
    expect(await count()).toBe(6);
    expect((await runs())[0]?.status).toBe("succeeded");
  });

  it("summarizes by country and masks failing networks", async () => {
    await setup();
    for (let minute = 10; minute >= 4; minute -= 1)
      records.push(
        signIn(new Date(start.getTime() - minute * 60_000), {
          userId: "user-2",
          ipAddress: "203.0.113.9",
          status: { errorCode: 50126 },
          location: paris,
        }),
      );
    await poll();
    const summary = await database.tenant(tenantId, (sql) =>
      summarizeSignins(sql, tenantId, clock, 7),
    );
    expect(summary.totals).toMatchObject({ successes: 18, failures: 7, accounts: 1 });
    expect(summary.countries).toEqual([
      { country: "US", successes: 18, failures: 0 },
      { country: "FR", successes: 0, failures: 7 },
    ]);
    expect(summary.failingNetworks).toEqual([
      { network: "203.0.113.0/24", failures: 7, accounts: 1 },
    ]);
    expect(summary.coverage.events).toBe(25);
  });
  it("shrinks a window too busy for one run and still reads every sign-in", async () => {
    records = [];
    const busy = start.getTime() - 20 * 3_600_000;
    for (let index = 0; index < 300; index += 1)
      records.push(
        signIn(new Date(busy + index * 12_000), {
          userId: `user-${index % 5}`,
          ipAddress: "198.51.100.20",
        }),
      );
    requestDelay = 25;
    await setup({ backfillDays: 1, windowRecordLimit: 10, runBudgetSeconds: 0.2 });
    const spans: number[] = [];
    let historyPhase = "";
    for (let round = 0; round < 80; round += 1) {
      await poll();
      const position = await cursor();
      if (position.span) spans.push(position.span);
      if (round === 0)
        historyPhase = (
          await database.tenant(tenantId, (sql) => summarizeSignins(sql, tenantId, clock, 7))
        ).history.phase;
      if (position.phase === "live") break;
    }
    expect(historyPhase).toBe("backfill");
    expect((await cursor()).phase).toBe("live");
    expect(Math.min(...spans)).toBeLessThan(6 * 3_600_000);
    expect(await count()).toBe(300);
    expect((await findings()).filter((item) => item.ruleId === "signins.collection_gap")).toEqual(
      [],
    );
    expect((await runs())[0]?.status).not.toBe("failed");
  }, 60_000);

  it("keeps what it read of an unreadable burst, records the gap, and moves on", async () => {
    records = [];
    const at = new Date(start.getTime() - 5 * 3_600_000);
    for (let index = 0; index < 8; index += 1)
      records.push(signIn(at, { userId: `user-${index}`, ipAddress: "203.0.113.80" }));
    await setup({ backfillDays: 1, windowRecordLimit: 5 });
    await poll();
    expect(await count()).toBe(5);
    const gaps = (await findings()).filter((item) => item.ruleId === "signins.collection_gap");
    expect(gaps).toHaveLength(1);
    expect(gaps[0]).toMatchObject({ purpose: "operations", subject: { kind: "service" } });
    expect(gaps[0]?.summary).toContain("5 were read");
    expect((await runs())[0]?.status).toBe("partial");
    expect((await cursor()).phase).toBe("live");
  });

  it("keeps sign-ins that share a millisecond with a page boundary", async () => {
    records = [];
    const base = start.getTime() - 5 * 3_600_000;
    for (let index = 0; index < 20; index += 1) {
      const moment = new Date(base + index * 3_000);
      for (const fraction of ["1000", "5000"])
        records.push({
          ...signIn(moment, { ipAddress: "192.0.2.30" }),
          createdDateTime: moment.toISOString().replace("Z", `${fraction}Z`),
        });
    }
    await setup({ backfillDays: 1, windowRecordLimit: 9 });
    await poll();
    expect(await count()).toBe(40);
    expect((await runs())[0]?.status).toBe("succeeded");
  });

  it("reports history as live once the first load finishes", async () => {
    await setup();
    await poll();
    const summary = await database.tenant(tenantId, (sql) =>
      summarizeSignins(sql, tenantId, clock, 7),
    );
    expect(summary.history).toMatchObject({ phase: "live", from: null });
    expect(summary.history.through).toEqual(expect.any(String));
  });
  it("keeps live collection moving forward under steady heavy volume", async () => {
    records = [];
    await setup({ backfillDays: 1, windowRecordLimit: 20, runBudgetSeconds: 0.25 });
    await poll();
    expect((await cursor()).phase).toBe("live");
    requestDelay = 10;
    let added = 0;
    let previous = "";
    for (let round = 0; round < 12; round += 1) {
      // 150 sign-ins in the five minutes since the last run.
      for (let index = 0; index < 150; index += 1)
        records.push(
          signIn(new Date(clock.getTime() - index * 2_000), { userId: `user-${index % 9}` }),
        );
      added += 150;
      await poll();
      const position = await cursor();
      const through = (position as { through?: string }).through ?? "";
      expect(through >= previous, `round ${round}: ${through} < ${previous}`).toBe(true);
      previous = through;
    }
    expect(await count()).toBe(added);
  }, 60_000);

  it("retries a busy window the run reached late instead of recording a gap", async () => {
    records = [];
    const early = start.getTime() - 23 * 3_600_000;
    for (let index = 0; index < 100; index += 1)
      records.push(signIn(new Date(early + index * 30_000), { userId: `user-${index % 4}` }));
    const burst = start.getTime() - 15 * 3_600_000;
    for (let index = 0; index < 150; index += 1)
      records.push(signIn(new Date(burst + index * 600), { userId: `user-${index % 7}` }));
    requestDelay = 20;
    await setup({ backfillDays: 1, windowRecordLimit: 10, runBudgetSeconds: 0.5 });
    for (let round = 0; round < 40 && (await cursor()).phase !== "live"; round += 1) await poll();
    expect((await cursor()).phase).toBe("live");
    expect((await findings()).filter((item) => item.ruleId === "signins.collection_gap")).toEqual(
      [],
    );
    expect(await count()).toBe(250);
  }, 60_000);

  it("records a gap for a moment Microsoft never answers in time, then moves on", async () => {
    records = [];
    slowAt = start.getTime() - 10 * 3_600_000;
    for (let index = 0; index < 10; index += 1)
      records.push(signIn(new Date(start.getTime() - (20 - index) * 3_600_000)));
    await setup({ backfillDays: 1 }, { requestTimeoutMilliseconds: 150 });
    for (let round = 0; round < 40 && (await cursor()).phase !== "live"; round += 1) await poll();
    expect((await cursor()).phase).toBe("live");
    const gaps = (await findings()).filter((item) => item.ruleId === "signins.collection_gap");
    expect(gaps).toHaveLength(1);
    // Nothing was read there, so the finding says Microsoft did not answer, not that it was busy.
    expect(gaps[0]?.summary).toContain("Microsoft did not answer in time");
    expect(gaps[0]?.summary).not.toContain("burst");
    // Every record outside the slow five minutes was still collected.
    expect(await count()).toBe(10);
  }, 60_000);

  it("loses nothing Microsoft publishes late, even under steady heavy volume", async () => {
    records = [];
    await setup({
      backfillDays: 1,
      windowRecordLimit: 20,
      runBudgetSeconds: 0.5,
      lateArrivalMinutes: 5,
    });
    await poll();
    requestDelay = 10;
    // Every sign-in reaches Graph a minute after it happens, so the forward read of each run
    // misses the newest ones and only the re-read behind it can collect them.
    publishDelay = 60_000;
    for (let round = 0; round < 12; round += 1) {
      for (let index = 0; index < 150; index += 1)
        records.push(
          signIn(new Date(clock.getTime() - index * 2_000), { userId: `user-${index % 9}` }),
        );
      await poll();
    }
    const ids = await storedIds();
    const settled = records.filter(
      (record) => Date.parse(record.createdDateTime) <= clock.getTime() - 15 * 60_000,
    );
    expect(settled.length).toBeGreaterThan(1_000);
    expect(settled.filter((record) => !ids.has(record.id))).toEqual([]);
  }, 60_000);

  it("re-reads the minutes before a long pause once collection resumes", async () => {
    records = [];
    await setup({ backfillDays: 1, windowRecordLimit: 20, runBudgetSeconds: 0.3 });
    await poll();
    publishDelay = 2 * 60_000;
    requestDelay = 5;
    const trickle = (minutes: number) => {
      for (let second = 0; second < minutes * 60; second += 20)
        records.push(signIn(new Date(clock.getTime() - second * 1_000), { userId: "user-3" }));
    };
    for (let round = 0; round < 3; round += 1) {
      trickle(5);
      await poll();
    }
    // Two hours without a run (sign-out, restart, an outage) while sign-ins continue.
    for (let gap = 0; gap < 24; gap += 1) {
      clock = new Date(clock.getTime() + 5 * 60_000);
      trickle(5);
    }
    for (let round = 0; round < 20; round += 1) {
      trickle(5);
      await poll();
    }
    const ids = await storedIds();
    const settled = records.filter(
      (record) => Date.parse(record.createdDateTime) <= clock.getTime() - 65 * 60_000,
    );
    expect(settled.filter((record) => !ids.has(record.id))).toEqual([]);
  }, 60_000);

  it("treats a window a stopped run was reading as failed, and gives up on it after three", async () => {
    await setup({ backfillDays: 1 });
    // The last run was stopped while reading its first six-hour window of history.
    const begin = new Date(start.getTime() - 86_400_000).toISOString();
    await poll();
    await setCursor({ v: 1, phase: "backfill", start: begin, next: begin, trying: begin });
    requested = [];
    await poll();
    const first = requested[0];
    expect(first?.from).toBe(begin);
    expect(Date.parse(first?.to ?? "") - Date.parse(begin)).toBe(3 * 3_600_000);

    // A five-minute window that has stopped two runs already is passed on the third, with a gap.
    const through = new Date(clock.getTime() - 5 * 60_000).toISOString();
    await setCursor({
      v: 1,
      phase: "live",
      through,
      trying: through,
      strikes: { at: through, count: 2 },
      settle: { at: through },
    });
    await poll();
    const gaps = (await findings()).filter((item) => item.ruleId === "signins.collection_gap");
    expect(gaps).toHaveLength(1);
    expect(gaps[0]?.summary).toContain("stopped while reading");
    expect(Date.parse(((await cursor()) as { through: string }).through)).toBeGreaterThan(
      Date.parse(through),
    );
  });

  it("keeps the older sign-ins of a window with a millisecond too dense to read", async () => {
    records = [];
    const at = new Date(start.getTime() - 5 * 3_600_000);
    for (let index = 0; index < 20; index += 1)
      records.push(
        signIn(new Date(at.getTime() - (index + 1) * 60_000), { userId: "user-sparse" }),
      );
    for (let index = 0; index < 8; index += 1)
      records.push(signIn(at, { userId: `user-${index}`, ipAddress: "203.0.113.80" }));
    await setup({ backfillDays: 1, windowRecordLimit: 5 });
    await poll();
    expect(await count("user_id = 'user-sparse'")).toBe(20);
    expect(await count()).toBe(25);
    const gaps = (await findings()).filter((item) => item.ruleId === "signins.collection_gap");
    expect(gaps).toHaveLength(1);
    expect(gaps[0]?.summary).toContain("share the millisecond");
  });

  it("keeps what it read of a dense five-minute window when the budget ends mid-page, and moves on", async () => {
    records = [];
    const burst = start.getTime() - 5 * 3_600_000;
    for (let index = 0; index < 400; index += 1)
      records.push(signIn(new Date(burst + index * 400), { userId: `user-${index % 11}` }));
    // Each page of 100 takes 60 ms and a run may spend 100 ms: no run can page through 400.
    requestDelay = 60;
    delayEmptyPages = false;
    await setup({ backfillDays: 1, windowRecordLimit: 500, runBudgetSeconds: 0.1 });
    for (let round = 0; round < 60 && (await cursor()).phase !== "live"; round += 1) {
      await poll();
    }
    expect((await cursor()).phase).toBe("live");
    // Most of the burst is kept; the part no run could page through is named by one gap.
    const kept = (await count()) ?? 0;
    expect(kept).toBeGreaterThan(200);
    expect(kept).toBeLessThan(400);
    const gaps = (await findings()).filter((item) => item.ruleId === "signins.collection_gap");
    expect(gaps).toHaveLength(1);
    expect(gaps[0]?.summary).toMatch(/\d+ were read and kept/);
  }, 60_000);

  it("collects records Microsoft publishes 45 minutes late when no run has time to re-read them early", async () => {
    records = [];
    await setup({ backfillDays: 1, windowRecordLimit: 20, runBudgetSeconds: 0.35 });
    await poll();
    requestDelay = 25;
    const late: FakeSignIn[] = [];
    for (let round = 0; round < 24; round += 1) {
      // Most sign-ins reach Graph at once, so re-reading recent minutes costs a run its time;
      // a few reach it 45 minutes late, and only the settle lane goes back that far.
      for (let index = 0; index < 60; index += 1)
        records.push(
          signIn(new Date(clock.getTime() - index * 5_000), { userId: `user-${index % 7}` }),
        );
      for (let index = 0; index < 10; index += 1) {
        const record = signIn(new Date(clock.getTime() - index * 30_000), {
          userId: "user-late",
          publishDelay: 45 * 60_000,
        });
        late.push(record);
        records.push(record);
      }
      await poll();
    }
    const ids = await storedIds();
    const due = late.filter(
      (record) => Date.parse(record.createdDateTime) <= clock.getTime() - 70 * 60_000,
    );
    expect(due.length).toBeGreaterThan(100);
    expect(due.filter((record) => !ids.has(record.id))).toEqual([]);
  }, 90_000);

  it("marks the window it is reading before each read, so a stopped run can be detected", async () => {
    await setup();
    await poll();
    records.push(signIn(minutesAgo(1)));
    const seen: Record<string, unknown>[] = [];
    onRequest = async () => {
      seen.push(await cursor());
    };
    await poll();
    onRequest = null;
    expect(
      seen.some(
        (value) =>
          typeof value.trying === "string" &&
          (value.trying === value.through ||
            value.trying === (value.settle as { at?: string } | undefined)?.at),
      ),
    ).toBe(true);
    // A normal run leaves no marker behind.
    const after = await cursor();
    expect((after as { trying?: string }).trying).toBeUndefined();
  });

  it("moves the live position past a burst no run can page through, keeping what it read", async () => {
    await setup({ backfillDays: 1, windowRecordLimit: 300, runBudgetSeconds: 0.2 });
    await poll();
    // A thousand sign-ins in the next five minutes: each request pages 100 at a time, 40 ms a
    // page, and a run may spend 200 ms, so every run stops partway through its page sequence.
    requestDelay = 40;
    delayEmptyPages = false;
    for (let index = 0; index < 1_000; index += 1)
      records.push(
        signIn(new Date(clock.getTime() + index * 250), { userId: `user-${index % 13}` }),
      );
    const before = Date.parse(((await cursor()) as { through: string }).through);
    for (let round = 0; round < 4; round += 1) await poll();
    const through = Date.parse(((await cursor()) as { through: string }).through);
    expect(through).toBeGreaterThan(before + 5 * 60_000);
    const kept = (await count("user_id LIKE 'user-%' AND user_id <> 'user-1'")) ?? 0;
    expect(kept).toBeGreaterThan(100);
    const gaps = (await findings()).filter((item) => item.ruleId === "signins.collection_gap");
    expect(gaps.length).toBeGreaterThan(0);
  }, 60_000);

  it("lets the settle lane read first when a dense window needs the whole budget", async () => {
    await setup({ backfillDays: 1, windowRecordLimit: 300, runBudgetSeconds: 0.2 });
    await poll();
    requestDelay = 40;
    delayEmptyPages = false;
    // A dense minute that Microsoft publishes 50 minutes late: only the settle lane reads it,
    // and no run that reads new sign-ins first has time for all of it.
    const burst = clock.getTime() + 60_000;
    for (let index = 0; index < 1_000; index += 1)
      records.push(
        signIn(new Date(burst + index * 50), {
          userId: `user-${index % 13}`,
          publishDelay: 50 * 60_000,
        }),
      );
    for (let round = 0; round < 24; round += 1) await poll();
    const settled = Date.parse(
      ((await cursor()) as { settle?: { at: string } }).settle?.at ?? "1970-01-01T00:00:00Z",
    );
    expect(settled).toBeGreaterThan(burst + 5 * 60_000);
    const gaps = (await findings()).filter((item) => item.ruleId === "signins.collection_gap");
    expect(gaps.length).toBeGreaterThan(0);
  }, 90_000);

  it("keeps the settle lane moving under traffic that fills every run", async () => {
    await setup({ backfillDays: 1, windowRecordLimit: 300, runBudgetSeconds: 0.2 });
    await poll();
    requestDelay = 40;
    delayEmptyPages = false;
    const startedAt = Date.parse(
      ((await cursor()) as { settle?: { at: string } }).settle?.at ?? "1970-01-01T00:00:00Z",
    );
    for (let round = 0; round < 24; round += 1) {
      // More new sign-ins each cadence than one run can read.
      for (let index = 0; index < 700; index += 1)
        records.push(
          signIn(new Date(clock.getTime() - index * 420), { userId: `user-${index % 17}` }),
        );
      await poll();
    }
    const settled = Date.parse(
      ((await cursor()) as { settle?: { at: string } }).settle?.at ?? "1970-01-01T00:00:00Z",
    );
    // Two hours of traffic: the re-read must not sit still while the main lane reads.
    expect(settled - startedAt).toBeGreaterThan(30 * 60_000);
  }, 120_000);

  it("retries a busy half of a window it split after a timeout, instead of recording a gap", async () => {
    records = [];
    await setup({ backfillDays: 1, runBudgetSeconds: 0.35 }, { requestTimeoutMilliseconds: 150 });
    await poll();
    // Ten minutes to read. Microsoft answers anything over five minutes too slowly, so the window
    // is split; its newer half holds more sign-ins than the time left after the timeout allows.
    const from = new Date(clock.getTime() - 10 * 60_000);
    await setCursor({
      v: 1,
      phase: "live",
      through: from.toISOString(),
      settle: { at: from.toISOString() },
    });
    slowFrom = 5 * 60_000;
    requestDelay = 60;
    delayEmptyPages = false;
    for (let index = 0; index < 400; index += 1)
      records.push(
        signIn(new Date(clock.getTime() - index * 500), { userId: `user-${index % 9}` }),
      );
    await poll();
    // That run did not have the whole budget for the busy half, so it may not call it a gap.
    expect((await findings()).filter((item) => item.ruleId === "signins.collection_gap")).toEqual(
      [],
    );
    for (let round = 0; round < 6; round += 1) await poll();
    expect(await count(`user_id <> 'user-1'`)).toBeGreaterThan(0);
  }, 60_000);

  it("does not let a failing re-read block new sign-ins", async () => {
    await setup({ backfillDays: 1 });
    await poll();
    // The settle lane is owed the next run and every read behind the live position fails.
    const position = (await cursor()) as { through: string; settle?: { at: string } };
    await setCursor({
      ...position,
      settle: { at: position.settle?.at ?? position.through, owed: true },
    });
    failBefore = Date.parse(position.through);
    for (let round = 0; round < 3; round += 1) {
      records.push(signIn(minutesAgo(1), { userId: `user-new-${round}` }));
      await poll();
    }
    expect(await count("user_id LIKE 'user-new-%'")).toBeGreaterThanOrEqual(2);
  });

  describe("canary accounts", () => {
    const decoy = {
      userId: "canary-1",
      principal: "decoy.account@example.invalid",
      addedAt: "2026-10-22T11:00:00.000Z",
      addedBy: actorId,
    };
    const mark = (...accounts: (typeof decoy)[]) =>
      database.tenant(tenantId, (sql) => writeCanaries(sql, tenantId, { accounts }, actorId));
    const used = async () =>
      (await findings()).filter((item) => item.ruleId === "signin.canary-used");
    const attempt = (at: Date, overrides: Record<string, unknown> = {}) =>
      signIn(at, {
        userId: "canary-1",
        userPrincipalName: "Decoy.Account@example.invalid",
        userDisplayName: "Decoy Account",
        ipAddress: "203.0.113.44",
        location: amsterdam,
        status: { errorCode: 50126, failureReason: "Invalid username or password." },
        ...overrides,
      });

    it("opens a critical finding in the run that reads the attempt, and leaves other accounts alone", async () => {
      await setup();
      await poll();
      await mark(decoy);
      records.push(attempt(minutesAgo(3)), signIn(minutesAgo(2)));
      await poll();
      const [found, ...others] = await used();
      expect(others).toEqual([]);
      expect(found).toMatchObject({
        severity: "critical",
        state: "open",
        occurrences: 1,
        subject: { kind: "person", id: "canary-1", label: "Decoy Account" },
      });
      const evidence = (found?.evidence ?? []) as { fields: Record<string, unknown> }[];
      expect(evidence[0]?.fields).toMatchObject({
        result: "failure",
        errorCode: 50126,
        ip: "203.0.113.44",
        app: "Office 365 Exchange Online",
      });
      // The ordinary account that signed in during the same run has no finding.
      expect((await findings()).filter((item) => item.subject.id === "user-1")).toEqual([]);
    });

    it("counts up within a UTC day, opens a new finding the next day, and fires on a success", async () => {
      clock = new Date("2026-10-22T23:50:00Z");
      await setup();
      await poll();
      await mark(decoy);
      records.push(attempt(minutesAgo(3)));
      await poll();
      records.push(attempt(minutesAgo(2), { status: { errorCode: 0, failureReason: null } }));
      await poll();
      let found = await used();
      expect(found).toHaveLength(1);
      expect(found[0]).toMatchObject({ occurrences: 2, severity: "critical" });
      const evidence = (found[0]?.evidence ?? []) as { fields: Record<string, unknown> }[];
      expect(evidence[0]?.fields).toMatchObject({ result: "success", errorCode: 0 });

      // After midnight UTC the same account opens a new finding.
      records.push(attempt(minutesAgo(2)));
      await poll();
      found = await used();
      expect(found.map((item) => item.occurrences).sort()).toEqual([1, 2]);
    });

    it("also reads the history loaded on the first run", async () => {
      await setup();
      await mark(decoy);
      records.push(attempt(new Date(start.getTime() - 30 * 3_600_000)));
      await poll();
      expect(await used()).toMatchObject([{ occurrences: 1, severity: "critical" }]);
    });

    it("matches a sign-in with no user id by the account's sign-in name", async () => {
      await setup();
      await poll();
      await mark({ ...decoy, userId: "canary-2", principal: "Old.Mailbox@example.invalid" });
      records.push(
        attempt(minutesAgo(3), { userId: "", userPrincipalName: "OLD.MAILBOX@example.invalid" }),
      );
      await poll();
      expect(await used()).toMatchObject([{ subject: { id: "canary-2" } }]);
    });

    it("does not look back: an attempt already read before the account was marked stays quiet", async () => {
      await setup();
      await poll();
      records.push(attempt(minutesAgo(3)));
      await poll();
      await mark(decoy);
      await poll();
      expect(await used()).toEqual([]);
    });

    it("reads the list once per run, however many windows the run reads", async () => {
      let runsStarted = 0;
      let reads = 0;
      await setup({}, {}, (module) => ({
        ...module,
        collectors: module.collectors.map((collector) => ({
          ...collector,
          run: (context) => {
            runsStarted += 1;
            return collector.run({
              ...context,
              settings: {
                get: (key) => {
                  reads += 1;
                  return context.settings.get(key);
                },
              },
            });
          },
        })),
      }));
      // Three days of history are read in many windows.
      await poll();
      expect(await count()).toBe(18);
      expect(runsStarted).toBeGreaterThan(0);
      expect(reads).toBe(runsStarted);
    });

    it("keeps collecting when the stored list is damaged", async () => {
      await setup();
      await poll();
      await database.tenant(tenantId, (sql) =>
        sql.query(
          `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
           VALUES ($1,'signins.canaries','{"accounts":"no"}'::jsonb,now(),$2)`,
          [tenantId, actorId],
        ),
      );
      records.push(attempt(minutesAgo(3)));
      await poll();
      expect(await used()).toEqual([]);
      const [run] = await runs();
      expect(run).toMatchObject({ status: "succeeded" });
    });
  });
});

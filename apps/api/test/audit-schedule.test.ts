import { randomUUID } from "node:crypto";
import {
  type Audit,
  auditDefinition,
  builtInAudits,
  finishRun,
  insertRun,
  listRuns,
  writeAuditSchedule,
} from "@prism365/audits";
import { type AuditAvailability, type AuditSchedule, auditTrendSchema } from "@prism365/contracts";
import { GraphHttpClient } from "@prism365/graph";
import { PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import { AuditScheduler, type ScheduleCredential } from "../src/audit-schedule.ts";
import type { AuditRunner } from "../src/audits.ts";
import { actorId, auditHarness, manualTimers, tenantId } from "./audit-harness.ts";

// Scheduled audits: the scheduler on a fake clock with a fake runner, then wired into the app with
// the real runner. Synthetic tenant and administrator only.

const HOUR = 3_600_000;
type StartInput = Parameters<AuditRunner["start"]>[0];

// Records each start, stores the run as the real runner would (so the next due check sees it),
// and finishes it on its own or when the test says.
function fakeRunner(database: PGliteModuleDatabase, clock: () => Date) {
  const starts: StartInput[] = [];
  const running = new Map<string, string>();
  const waiting: (() => Promise<void>)[] = [];
  let outcome: "succeeded" | "failed" | "hold" = "succeeded";
  const finish = async (input: StartInput, id: string, status: "succeeded" | "failed") => {
    await database.tenant(input.tenantId, (sql) =>
      finishRun(
        sql,
        input.tenantId,
        id,
        clock(),
        status === "succeeded"
          ? { status, complete: true, summary: [], notes: [], rows: [] }
          : { status, code: "throttled", message: "Microsoft is limiting requests right now." },
      ),
    );
    running.delete(`${input.tenantId}:${input.audit.definition.id}`);
  };
  const runner: Pick<AuditRunner, "start" | "runningId"> = {
    runningId: (tenant, auditId) => running.get(`${tenant}:${auditId}`) ?? null,
    async start(input) {
      const id = randomUUID();
      const run = await database.tenant(input.tenantId, (sql) =>
        insertRun(sql, input.tenantId, {
          id,
          auditId: input.audit.definition.id,
          auditVersion: input.audit.definition.version,
          startedAt: clock(),
          startedBy: input.actor,
          trigger: input.trigger ?? "manual",
        }),
      );
      starts.push(input);
      running.set(`${input.tenantId}:${input.audit.definition.id}`, id);
      await input.beforeRead?.(run);
      let done: Promise<void>;
      if (outcome === "hold")
        done = new Promise<void>((resolve) => {
          waiting.push(async () => {
            await finish(input, id, "succeeded");
            resolve();
          });
        });
      else done = finish(input, id, outcome);
      return { runId: id, started: true, run, done };
    },
  };
  return {
    runner,
    starts,
    running,
    started: () => starts.map((input) => input.audit.definition.id),
    finishWith: (value: typeof outcome) => {
      outcome = value;
    },
    // Lets the oldest held run finish.
    release: async () => {
      await waiting.shift()?.();
    },
  };
}

// The manual route's rule, reduced to what these tests need: every required pack enabled and its
// permission granted.
const somePacks = ["directory.users.read", "audit.signins.read"];
const morePacks = [...somePacks, "reports.authentication-methods.read"];
const availability =
  (enabledPacks: readonly string[]) =>
  (audit: Audit, granted: readonly string[]): AuditAvailability => {
    const off = audit.packs.filter((pack) => !enabledPacks.includes(pack));
    if (off.length > 0)
      return { state: "not_enabled", missing: [], grantFeature: null, enable: off };
    const missing = auditDefinition(audit).permissions.filter((item) => !granted.includes(item));
    return missing.length > 0
      ? { state: "missing_permission", missing, grantFeature: null, enable: [] }
      : { state: "available", missing: [], grantFeature: null, enable: [] };
  };

// One database for the scheduler tests; each test works in a tenant of its own.
let database: PGliteModuleDatabase;
beforeAll(async () => {
  database = await PGliteModuleDatabase.open();
}, 60_000);
afterAll(async () => {
  await database?.close();
});

async function setup(schedule: AuditSchedule["audits"], enabledPacks = somePacks) {
  const tenantId = randomUUID();
  let now = new Date("2026-10-22T12:00:00Z");
  const clock = () => now;
  const fake = fakeRunner(database, clock);
  let granted = ["User.Read", "User.Read.All", "AuditLog.Read.All"];
  let credential: ScheduleCredential | null = {
    tenantId,
    actorId,
    displayName: "Synthetic admin",
    grantedPermissions: async () => granted,
    accessToken: async () => "synthetic",
  };
  const ledger: { tenantId: string; actorId: string; event: Record<string, unknown> }[] = [];
  const logs: string[] = [];
  const timers = manualTimers();
  const scheduler = new AuditScheduler({
    runner: fake.runner,
    database,
    audits: builtInAudits,
    credential: () => credential,
    enabledPacks,
    availability: availability(enabledPacks),
    graphClient: (getAccessToken) => new GraphHttpClient({ getAccessToken }),
    ledger: async (tenant, actor, event) => {
      ledger.push({ tenantId: tenant, actorId: actor, event });
    },
    now: clock,
    timers: timers.timers,
    log: (message) => logs.push(message),
  });
  const save = (audits: AuditSchedule["audits"]) =>
    database.tenant(tenantId, (sql) =>
      writeAuditSchedule(sql, tenantId, { audits }, actorId, builtInAudits),
    );
  await save(schedule);
  return {
    tenantId,
    database,
    scheduler,
    fake,
    ledger,
    logs,
    timers,
    save,
    advance: (milliseconds: number) => {
      now = new Date(now.getTime() + milliseconds);
    },
    setCredential: (value: ScheduleCredential | null) => {
      credential = value;
    },
    credential: () => credential,
    grant: (scopes: string[]) => {
      granted = scopes;
    },
  };
}

describe("audit scheduler", () => {
  it("starts nothing while collection is off", async () => {
    const s = await setup({ "accounts.inactive": "daily" });
    s.setCredential(null);
    await s.scheduler.tick();
    expect(s.fake.starts).toEqual([]);
  });

  it("runs a daily audit that never ran, then once a day from its last start", async () => {
    const s = await setup({ "accounts.inactive": "daily" });
    await s.scheduler.tick();
    expect(s.fake.started()).toEqual(["accounts.inactive"]);
    await s.scheduler.tick();
    s.advance(23 * HOUR);
    await s.scheduler.tick();
    s.advance(HOUR);
    // Exactly 24 hours is not more than a day.
    await s.scheduler.tick();
    expect(s.fake.starts).toHaveLength(1);
    s.advance(60_000);
    await s.scheduler.tick();
    expect(s.fake.started()).toEqual(["accounts.inactive", "accounts.inactive"]);
  });

  it("counts a weekly audit from the newest start of any run, manual and failed included", async () => {
    const s = await setup({ "guests.review": "weekly" });
    const id = randomUUID();
    await s.database.tenant(s.tenantId, async (sql) => {
      await insertRun(sql, s.tenantId, {
        id,
        auditId: "guests.review",
        auditVersion: 1,
        startedAt: new Date("2026-10-15T13:00:00Z"),
        startedBy: { id: actorId, name: null },
      });
      await finishRun(sql, s.tenantId, id, new Date("2026-10-15T13:01:00Z"), {
        status: "failed",
        code: "forbidden",
        message: "Microsoft refused.",
      });
    });
    // Six days and 23 hours since that manual run started.
    await s.scheduler.tick();
    expect(s.fake.starts).toEqual([]);
    s.advance(2 * HOUR);
    await s.scheduler.tick();
    expect(s.fake.started()).toEqual(["guests.review"]);
  });

  it("does not retry a failed scheduled run before its next turn", async () => {
    const s = await setup({ "accounts.inactive": "daily" });
    s.fake.finishWith("failed");
    await s.scheduler.tick();
    for (let hour = 0; hour < 24; hour += 1) {
      s.advance(HOUR - 1);
      await s.scheduler.tick();
    }
    expect(s.fake.starts).toHaveLength(1);
    s.advance(60_000);
    await s.scheduler.tick();
    expect(s.fake.starts).toHaveLength(2);
    const runs = await s.database.tenant(s.tenantId, (sql) =>
      listRuns(sql, s.tenantId, "accounts.inactive"),
    );
    expect(runs.map((run) => [run.status, run.trigger])).toEqual([
      ["failed", "schedule"],
      ["failed", "schedule"],
    ]);
  });

  it("skips an audit whose packs are not enabled or granted, and logs only a category", async () => {
    const s = await setup({
      "accounts.inactive": "daily",
      "mfa.registration": "daily",
      "guests.review": "daily",
    });
    s.grant(["User.Read", "User.Read.All"]);
    await s.scheduler.tick();
    expect(s.fake.started()).toEqual(["guests.review"]);
    expect(s.logs).toContain(
      "Prism365 scheduled audit accounts.inactive skipped (missing_permission).",
    );
    // The MFA audit's own pack is not enabled here, whatever the session granted.
    expect(s.logs).toContain("Prism365 scheduled audit mfa.registration skipped (not_enabled).");
    // Skipped audits stay due: granted later, they run on the next tick.
    s.grant(["User.Read", "User.Read.All", "AuditLog.Read.All"]);
    await s.scheduler.tick();
    expect(s.fake.started()).toEqual(["guests.review", "accounts.inactive"]);
    const logged = s.logs.join("\n");
    expect(logged).not.toContain(s.tenantId);
    expect(logged).not.toContain(actorId);
    expect(logged).not.toContain("Synthetic admin");
  });

  it("runs one audit at a time, waiting for each to be recorded", async () => {
    const s = await setup({ "accounts.inactive": "daily", "licenses.reclaim": "daily" });
    s.fake.finishWith("hold");
    const pass = s.scheduler.tick();
    await vi.waitFor(() => expect(s.fake.starts).toHaveLength(1));
    // A tick while a pass runs joins it.
    expect(s.scheduler.tick()).toBe(pass);
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(s.fake.started()).toEqual(["accounts.inactive"]);
    await s.fake.release();
    await vi.waitFor(() => expect(s.fake.starts).toHaveLength(2));
    expect(s.fake.started()).toEqual(["accounts.inactive", "licenses.reclaim"]);
    await s.fake.release();
    await pass;
  });

  it("stops the pass when collection stops, or a schedule is turned off, during a run", async () => {
    const s = await setup(
      {
        "accounts.inactive": "daily",
        "licenses.reclaim": "daily",
        "mfa.registration": "daily",
        "guests.review": "daily",
      },
      morePacks,
    );
    s.fake.finishWith("hold");
    const pass = s.scheduler.tick();
    await vi.waitFor(() => expect(s.fake.starts).toHaveLength(1));
    await s.save({
      "accounts.inactive": "daily",
      "mfa.registration": "daily",
      "guests.review": "daily",
    });
    await s.fake.release();
    // Reclaim was turned off while inactive accounts ran: MFA is next.
    await vi.waitFor(() => expect(s.fake.starts).toHaveLength(2));
    expect(s.fake.started()).toEqual(["accounts.inactive", "mfa.registration"]);
    // Collection stops while MFA runs: guests does not start.
    s.setCredential(null);
    await s.fake.release();
    await pass;
    expect(s.fake.started()).toEqual(["accounts.inactive", "mfa.registration"]);
  });

  it("leaves an audit alone while a run of it is in progress", async () => {
    const s = await setup({ "accounts.inactive": "daily" });
    s.fake.running.set(`${s.tenantId}:accounts.inactive`, randomUUID());
    await s.scheduler.tick();
    expect(s.fake.starts).toEqual([]);
  });

  it("runs as the collecting administrator and writes the ledger entry before reading", async () => {
    const s = await setup({ "guests.review": "daily" });
    await s.scheduler.tick();
    const [input] = s.fake.starts;
    // Its own pack and the enabled optional one, with only the permissions those carry; the
    // license pack is not enabled here, so it is not read.
    expect(input).toMatchObject({
      tenantId: s.tenantId,
      actor: { id: actorId, name: "Synthetic admin" },
      trigger: "schedule",
      packs: ["directory.users.read", "audit.signins.read"],
      granted: ["User.Read.All", "AuditLog.Read.All"],
    });
    const [run] = await s.database.tenant(s.tenantId, (sql) =>
      listRuns(sql, s.tenantId, "guests.review"),
    );
    expect(s.ledger).toEqual([
      {
        tenantId: s.tenantId,
        actorId,
        event: {
          type: "audit.run.started",
          auditId: "guests.review",
          runId: run?.id,
          // Required, then the optional ones this sign-in granted.
          datasets: ["User.Read.All", "AuditLog.Read.All"],
          trigger: "schedule",
          occurredAt: "2026-10-22T12:00:00.000Z",
        },
      },
    ]);
  });

  it("pauses the pass when the collecting sign-in is gone", async () => {
    const s = await setup({ "accounts.inactive": "daily" });
    s.setCredential({
      tenantId: s.tenantId,
      actorId,
      displayName: null,
      grantedPermissions: async () => {
        throw new Error("The collection credential is no longer available.");
      },
      accessToken: async () => "synthetic",
    });
    await s.scheduler.tick();
    expect(s.fake.starts).toEqual([]);
    expect(s.logs).toEqual([
      "Prism365 scheduled audits paused: the collecting sign-in is not available.",
    ]);
  });

  it("ticks on its interval and once shortly after collection starts, and stops both", async () => {
    const s = await setup({ "accounts.inactive": "daily" });
    s.scheduler.start();
    s.scheduler.start();
    s.scheduler.soon();
    s.scheduler.soon();
    expect(s.timers.active().map((timer) => timer.kind)).toEqual(["every", "after"]);
    s.timers.fire("after");
    await vi.waitFor(() => expect(s.fake.starts).toHaveLength(1));
    await s.scheduler.settled();
    s.advance(25 * HOUR);
    s.scheduler.stop();
    expect(s.timers.active()).toEqual([]);
    // A callback that fires anyway, and every later request, starts nothing.
    for (const timer of s.timers.registered) timer.run();
    s.scheduler.start();
    s.scheduler.soon();
    await s.scheduler.tick();
    expect(s.timers.active()).toEqual([]);
    expect(s.fake.starts).toHaveLength(1);
  });
});

describe("scheduled audits in the app", () => {
  it("clear their timers when the app closes", async () => {
    const h = await auditHarness();
    expect(h.timers.active().map((timer) => timer.kind)).toEqual(["every"]);
    await h.post("/v1/collection/start");
    expect(h.timers.active().map((timer) => timer.kind)).toEqual(["every", "after"]);
    await h.app.close();
    expect(h.timers.active()).toEqual([]);
  }, 60_000);

  it("never run while collection is off", async () => {
    const h = await auditHarness();
    try {
      await h.post("/v1/settings/audit-schedule", { audits: { "licenses.reclaim": "daily" } });
      h.timers.fire("every");
      await new Promise((resolve) => setTimeout(resolve, 50));
      await h.runner.idle();
      expect((await h.get("/v1/audits/licenses.reclaim/runs")).json().runs).toEqual([]);
      expect(h.reads).toEqual([]);
    } finally {
      await h.app.close();
    }
  }, 60_000);

  it("run a due audit on the collecting sign-in, shortly after collection starts", async () => {
    const h = await auditHarness();
    try {
      await h.post("/v1/settings/audit-schedule", { audits: { "licenses.reclaim": "daily" } });
      await h.post("/v1/collection/start");
      await h.scheduler.tick();
      h.timers.fire("after");
      await vi.waitFor(async () => {
        const runs = (await h.get("/v1/audits/licenses.reclaim/runs")).json().runs;
        expect(runs.map((run: { status: string }) => run.status)).toEqual(["succeeded"]);
      });
      const [run] = (await h.get("/v1/audits/licenses.reclaim/runs")).json().runs;
      expect(run).toMatchObject({
        trigger: "schedule",
        startedBy: { id: actorId, name: "Synthetic admin" },
      });
      const started = h.store.audit.filter((entry) => entry.event.type === "audit.run.started");
      expect(started).toEqual([
        {
          tenantId,
          actorId,
          event: {
            type: "audit.run.started",
            auditId: "licenses.reclaim",
            runId: run.id,
            datasets: ["User.Read.All", "AuditLog.Read.All"],
            trigger: "schedule",
            occurredAt: expect.any(String),
          },
        },
      ]);
      // Every Graph read of the run came after its ledger entry.
      const reads = h.reads.filter((read) => read.path === "/v1.0/users");
      expect(reads.length).toBeGreaterThan(0);
      expect(reads.every((read) => read.ledgerStarts === 1)).toBe(true);
      const trend = auditTrendSchema.parse(
        (await h.get("/v1/audits/licenses.reclaim/trend")).json(),
      );
      expect(trend.points.map((point) => [point.runId, point.trigger])).toEqual([
        [run.id, "schedule"],
      ]);
      // Not due again until a day after it started.
      h.timers.fire("every");
      await new Promise((resolve) => setTimeout(resolve, 50));
      await h.runner.idle();
      expect((await h.get("/v1/audits/licenses.reclaim/runs")).json().runs).toHaveLength(1);
    } finally {
      await h.app.close();
    }
  }, 60_000);

  it("fail a run whose sign-in Microsoft stops accepting, the way a manual run fails", async () => {
    const h = await auditHarness();
    try {
      await h.post("/v1/settings/audit-schedule", { audits: { "licenses.reclaim": "daily" } });
      await h.post("/v1/collection/start");
      await h.scheduler.tick();
      h.refuseTokens(true);
      h.timers.fire("after");
      await vi.waitFor(async () => {
        const runs = (await h.get("/v1/audits/licenses.reclaim/runs")).json().runs;
        expect(runs.map((run: { status: string }) => run.status)).toEqual(["failed"]);
      });
      const [run] = (await h.get("/v1/audits/licenses.reclaim/runs")).json().runs;
      expect(run).toMatchObject({ trigger: "schedule", error: { code: "session" } });
      expect(h.reads.filter((read) => read.path === "/v1.0/users")).toEqual([]);
    } finally {
      await h.app.close();
    }
  }, 60_000);
});

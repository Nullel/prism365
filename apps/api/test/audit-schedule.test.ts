import { randomUUID } from "node:crypto";
import {
  type Audit,
  auditDefinition,
  auditReadScope,
  builtInAudits,
  finishRun,
  findAudit,
  insertRun,
  listRuns,
  storedAuditSchedule,
  writeAuditSchedule,
} from "@prism365/audits";
import { type AuditAvailability, type AuditSchedule, auditTrendSchema } from "@prism365/contracts";
import { GraphHttpClient } from "@prism365/graph";
import { CredentialUnavailableError } from "@prism365/modules";
import { PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import {
  AuditScheduler,
  type ScheduleCredential,
  scheduledReadScope,
  scheduleLimits,
} from "../src/audit-schedule.ts";
import type { AuditRunner } from "../src/audits.ts";
import { actorId, auditHarness, manualTimers, otherActorId, tenantId } from "./audit-harness.ts";

// Scheduled audits: the scheduler on a fake clock with a fake runner, then wired into the app with
// the real runner. Synthetic tenant and administrators only.

const HOUR = 3_600_000;
type StartInput = Parameters<AuditRunner["start"]>[0];
type Outcome = "succeeded" | "failed" | "hold";

// Records each start, stores the run as the real runner would (so the next due check sees it),
// and finishes it on its own or when the test says. With `read`, a run asks for a token first, as
// a real run does before its first Graph request, and fails as a session failure if refused.
function fakeRunner(
  database: PGliteModuleDatabase,
  clock: () => Date,
  token: () => () => Promise<string>,
) {
  const starts: StartInput[] = [];
  const stopped: string[] = [];
  const running = new Map<string, string>();
  const held = new Map<string, (status: "succeeded" | "failed", code?: string) => Promise<void>>();
  let outcome: Outcome = "succeeded";
  let read = false;
  const finish = async (
    input: StartInput,
    id: string,
    status: "succeeded" | "failed",
    code = "throttled",
  ) => {
    await database.tenant(input.tenantId, (sql) =>
      finishRun(
        sql,
        input.tenantId,
        id,
        clock(),
        status === "succeeded"
          ? { status, complete: true, summary: [], notes: [], rows: [] }
          : { status, code, message: "Synthetic failure." },
      ),
    );
    running.delete(`${input.tenantId}:${input.audit.definition.id}`);
  };
  const runner: Pick<AuditRunner, "start" | "runningId" | "stopRun"> = {
    runningId: (tenant, auditId) => running.get(`${tenant}:${auditId}`) ?? null,
    stopRun: (runId) => {
      stopped.push(runId);
      const stop = held.get(runId);
      if (!stop) return false;
      void stop("failed", "collection_stopped");
      return true;
    },
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
      if (read) {
        const refused = await token()().then(
          () => false,
          () => true,
        );
        if (refused)
          return { runId: id, started: true, run, done: finish(input, id, "failed", "session") };
      }
      let done: Promise<void>;
      if (outcome === "hold")
        done = new Promise<void>((resolve) => {
          held.set(id, async (status, code) => {
            held.delete(id);
            await finish(input, id, status, code);
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
    stopped,
    running,
    started: () => starts.map((input) => input.audit.definition.id),
    finishWith: (value: Outcome) => {
      outcome = value;
    },
    reading: (value: boolean) => {
      read = value;
    },
    // Lets the oldest held run finish.
    release: async () => {
      const [first] = held.values();
      await first?.("succeeded");
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
  // The token functions the scheduler handed to each run's Graph client, newest last.
  const tokens: (() => Promise<string>)[] = [];
  const fake = fakeRunner(database, clock, () => tokens.at(-1) ?? (async () => "none"));
  let granted = ["User.Read", "User.Read.All", "AuditLog.Read.All"];
  const credentialFor = (overrides: Partial<ScheduleCredential> = {}): ScheduleCredential => ({
    tenantId,
    actorId,
    displayName: "Synthetic admin",
    grantedPermissions: async () => granted,
    accessToken: async () => "synthetic",
    ...overrides,
  });
  let credential: ScheduleCredential | null = credentialFor();
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
    graphClient: (getAccessToken) => {
      tokens.push(getAccessToken);
      return new GraphHttpClient({ getAccessToken });
    },
    ledger: async (tenant, actor, event) => {
      ledger.push({ tenantId: tenant, actorId: actor, event });
    },
    now: clock,
    timers: timers.timers,
    log: (message) => logs.push(message),
  });
  // The other administrator saves the schedule, with every permission granted unless a test
  // says otherwise; null saves the way a build without stored grants did.
  const save = (
    audits: AuditSchedule["audits"],
    saverScopes: string[] | null = ["User.Read", "User.Read.All", "AuditLog.Read.All"],
  ) =>
    database.tenant(tenantId, (sql) =>
      writeAuditSchedule(
        sql,
        tenantId,
        { audits },
        otherActorId,
        builtInAudits,
        saverScopes
          ? {
              name: "Other admin",
              grants: (audit) => auditReadScope(audit, enabledPacks, saverScopes).granted,
            }
          : undefined,
      ),
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
    tokens,
    save,
    credentialFor,
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

  it("logs an audit that stays unavailable once, and again only when the reason changes", async () => {
    const s = await setup({ "mfa.registration": "daily", "accounts.inactive": "daily" });
    s.grant(["User.Read", "User.Read.All"]);
    for (let tick = 0; tick < 3; tick += 1) await s.scheduler.tick();
    const skips = () => s.logs.filter((line) => line.includes("skipped"));
    expect(skips()).toEqual([
      "Prism365 scheduled audit accounts.inactive skipped (missing_permission).",
      "Prism365 scheduled audit mfa.registration skipped (not_enabled).",
    ]);
    // Granted, it runs; withdrawn again a day later, it is logged again.
    s.grant(["User.Read", "User.Read.All", "AuditLog.Read.All"]);
    await s.scheduler.tick();
    expect(s.fake.started()).toEqual(["accounts.inactive"]);
    s.advance(25 * HOUR);
    s.grant(["User.Read", "User.Read.All"]);
    await s.scheduler.tick();
    await s.scheduler.tick();
    expect(skips()).toHaveLength(3);
    expect(skips().at(-1)).toBe(
      "Prism365 scheduled audit accounts.inactive skipped (missing_permission).",
    );
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

  it("stops the run in progress and the pass when collection stops, and respects a schedule turned off", async () => {
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
    await s.save({ "licenses.reclaim": "off" });
    await s.fake.release();
    // Reclaim was turned off while inactive accounts ran: MFA is next.
    await vi.waitFor(() => expect(s.fake.starts).toHaveLength(2));
    expect(s.fake.started()).toEqual(["accounts.inactive", "mfa.registration"]);
    // Collection stops while MFA runs: the app tells the scheduler, which stops that run at once.
    s.setCredential(null);
    s.scheduler.collectionEnded();
    await pass;
    const mfa = await s.database.tenant(s.tenantId, (sql) =>
      listRuns(sql, s.tenantId, "mfa.registration"),
    );
    expect(s.fake.stopped).toEqual([mfa[0]?.id]);
    expect(mfa.map((run) => [run.status, run.error?.code])).toEqual([
      ["failed", "collection_stopped"],
    ]);
    // And guests does not start.
    expect(s.fake.started()).toEqual(["accounts.inactive", "mfa.registration"]);
    // Only runs the scheduler started are ever stopped this way.
    s.scheduler.collectionEnded();
    expect(s.fake.stopped).toHaveLength(1);
  });

  it("refuses a token once its credential no longer collects, or the scheduler stopped", async () => {
    const s = await setup({ "accounts.inactive": "daily" });
    s.fake.finishWith("hold");
    const pass = s.scheduler.tick();
    await vi.waitFor(() => expect(s.fake.starts).toHaveLength(1));
    const token = s.tokens.at(-1);
    if (!token) throw new Error("The run got no Graph client.");
    await expect(token()).resolves.toBe("synthetic");
    // Another administrator's credential now collects: the old one gets no token.
    s.setCredential(s.credentialFor({ actorId: otherActorId }));
    await expect(token()).rejects.toBeInstanceOf(CredentialUnavailableError);
    s.setCredential(null);
    await expect(token()).rejects.toBeInstanceOf(CredentialUnavailableError);
    s.scheduler.collectionEnded();
    await pass;
    // Stopped, the scheduler hands out nothing even for the credential that collects.
    const t = await setup({ "accounts.inactive": "daily" });
    t.fake.finishWith("hold");
    const second = t.scheduler.tick();
    await vi.waitFor(() => expect(t.fake.starts).toHaveLength(1));
    t.scheduler.stop();
    await expect(t.tokens.at(-1)?.()).rejects.toBeInstanceOf(CredentialUnavailableError);
    t.scheduler.collectionEnded();
    await second;
  });

  it("starts nothing when collection stops while the audit is being checked", async () => {
    const s = await setup({ "accounts.inactive": "daily" });
    // Collection stops while the scheduler asks the sign-in what it may read.
    s.setCredential(
      s.credentialFor({
        grantedPermissions: async () => {
          s.setCredential(null);
          return ["User.Read", "User.Read.All", "AuditLog.Read.All"];
        },
      }),
    );
    await s.scheduler.tick();
    expect(s.fake.starts).toEqual([]);
    expect(s.ledger).toEqual([]);
  });

  it("ends the pass when Microsoft refuses the sign-in, without starting the other audits", async () => {
    const s = await setup({
      "accounts.inactive": "daily",
      "licenses.reclaim": "daily",
      "guests.review": "daily",
    });
    s.fake.reading(true);
    const refusing = s.credentialFor({
      accessToken: async () => {
        throw new CredentialUnavailableError();
      },
    });
    s.setCredential(refusing);
    await s.scheduler.tick();
    expect(s.fake.started()).toEqual(["accounts.inactive"]);
    expect(s.ledger).toHaveLength(1);
    expect(s.logs.at(-1)).toBe(
      "Prism365 scheduled audits paused: the collecting sign-in is not available.",
    );
    // A credential already marked refused starts nothing at all.
    const t = await setup({ "accounts.inactive": "daily" });
    t.setCredential(t.credentialFor({ refused: true }));
    await t.scheduler.tick();
    expect(t.fake.starts).toEqual([]);
    expect(t.ledger).toEqual([]);
  });

  it("reads with no more than the administrator who scheduled it had granted", async () => {
    // Scheduled by someone whose own sign-in had not granted AuditLog.Read.All; collected by
    // someone whose sign-in has.
    const s = await setup({});
    await s.save({ "guests.review": "daily" }, ["User.Read", "User.Read.All"]);
    await s.scheduler.tick();
    const [input] = s.fake.starts;
    expect(input).toMatchObject({
      packs: ["directory.users.read"],
      granted: ["User.Read.All"],
    });
    expect(s.ledger[0]?.event).toMatchObject({ datasets: ["User.Read.All"] });
  });

  it("runs an entry saved without grants with its required permissions only", async () => {
    const s = await setup({});
    await s.save({ "guests.review": "daily" }, null);
    await s.scheduler.tick();
    expect(s.fake.starts[0]).toMatchObject({
      packs: ["directory.users.read"],
      granted: ["User.Read.All"],
    });
  });

  it("skips an audit whose scheduler's grants lack a required permission", async () => {
    const s = await setup({});
    await s.save({ "accounts.inactive": "daily" }, ["User.Read", "User.Read.All"]);
    await s.scheduler.tick();
    expect(s.fake.starts).toEqual([]);
    expect(s.logs).toContain(
      "Prism365 scheduled audit accounts.inactive skipped (missing_permission).",
    );
  });

  it("leaves an audit alone while a run of it is in progress", async () => {
    const s = await setup({ "accounts.inactive": "daily" });
    s.fake.running.set(`${s.tenantId}:accounts.inactive`, randomUUID());
    await s.scheduler.tick();
    expect(s.fake.starts).toEqual([]);
  });

  it("runs as the collecting administrator and names who scheduled it, before reading", async () => {
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
        // Whose sign-in the run reads with.
        actorId,
        event: {
          type: "audit.run.started",
          auditId: "guests.review",
          runId: run?.id,
          // Required, then the optional ones this sign-in granted.
          datasets: ["User.Read.All", "AuditLog.Read.All"],
          trigger: "schedule",
          scheduledBy: otherActorId,
          occurredAt: "2026-10-22T12:00:00.000Z",
        },
      },
    ]);
  });

  it("pauses the pass when the collecting sign-in is gone", async () => {
    const s = await setup({ "accounts.inactive": "daily" });
    s.setCredential(
      s.credentialFor({
        displayName: null,
        grantedPermissions: async () => {
          throw new Error("The collection credential is no longer available.");
        },
      }),
    );
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

describe("what a scheduled run may read", () => {
  const audit = (id: string) => {
    const found = findAudit(id);
    if (!found) throw new Error(`No audit ${id}.`);
    return found;
  };
  const everything = ["User.Read", "User.Read.All", "AuditLog.Read.All"];

  it("narrows the collecting sign-in to the scheduler's grants, and drops packs left without", () => {
    const guests = audit("guests.review");
    const narrow = storedAuditSchedule({
      audits: { "guests.review": "daily" },
      grants: { "guests.review": ["User.Read.All"] },
    });
    expect(scheduledReadScope(guests, somePacks, everything, narrow)).toEqual({
      scopes: ["User.Read.All"],
      packs: ["directory.users.read"],
      granted: ["User.Read.All"],
    });
    const full = storedAuditSchedule({
      audits: { "guests.review": "daily" },
      grants: { "guests.review": ["User.Read.All", "AuditLog.Read.All"] },
    });
    expect(scheduledReadScope(guests, somePacks, everything, full)).toEqual({
      scopes: ["User.Read.All", "AuditLog.Read.All"],
      packs: ["directory.users.read", "audit.signins.read"],
      granted: ["User.Read.All", "AuditLog.Read.All"],
    });
    // Neither sign-in can widen the other: the collecting one lacks what the scheduler had.
    expect(scheduledReadScope(guests, somePacks, ["User.Read.All"], full).granted).toEqual([
      "User.Read.All",
    ]);
    // An entry saved without grants reads with required permissions only, and damaged grants
    // count as none.
    for (const value of [
      { audits: { "guests.review": "daily" } },
      { audits: { "guests.review": "daily" }, grants: { "guests.review": "AuditLog.Read.All" } },
    ])
      expect(
        scheduledReadScope(guests, somePacks, everything, storedAuditSchedule(value)).granted,
      ).toEqual(["User.Read.All"]);
  });

  it("says which scheduled audits read less than a fully granted run, and by whose grants", () => {
    const stored = storedAuditSchedule({
      audits: {
        "guests.review": "daily",
        "accounts.inactive": "weekly",
        "licenses.reclaim": "off",
      },
      names: { "guests.review": "Other admin" },
      grants: {
        "guests.review": ["User.Read.All"],
        "accounts.inactive": ["User.Read.All", "AuditLog.Read.All"],
        "licenses.reclaim": ["User.Read.All"],
      },
    });
    expect(scheduleLimits(stored, builtInAudits, somePacks)).toEqual({
      "guests.review": { by: "Other admin", without: ["AuditLog.Read.All"] },
    });
    expect(
      scheduleLimits(
        storedAuditSchedule({ audits: { "guests.review": "daily" } }),
        builtInAudits,
        somePacks,
      ),
    ).toEqual({ "guests.review": { by: null, without: ["AuditLog.Read.All"] } });
  });
});

const runsOf = async (h: Awaited<ReturnType<typeof auditHarness>>, auditId: string) =>
  (await h.get(`/v1/audits/${auditId}/runs`)).json().runs as {
    id: string;
    status: string;
    trigger: string;
    error: { code: string; message: string } | null;
    startedBy: { id: string; name: string | null };
  }[];
const settle = async (h: Awaited<ReturnType<typeof auditHarness>>) => {
  await new Promise((resolve) => setTimeout(resolve, 50));
  await h.runner.idle();
};

describe("scheduled audits in the app", () => {
  it("clear their timers when the app closes", async () => {
    const h = await auditHarness();
    expect(h.timers.active().map((timer) => timer.kind)).toEqual(["every"]);
    await h.post("/v1/collection/start");
    expect(h.timers.active().map((timer) => timer.kind)).toEqual(["every", "after"]);
    await h.app.close();
    expect(h.timers.active()).toEqual([]);
  }, 60_000);

  it("never run while collection is off, including after it was stopped", async () => {
    const h = await auditHarness();
    try {
      await h.post("/v1/settings/audit-schedule", { audits: { "licenses.reclaim": "daily" } });
      h.timers.fire("every");
      await settle(h);
      expect(await runsOf(h, "licenses.reclaim")).toEqual([]);
      await h.post("/v1/settings/audit-schedule", { audits: { "licenses.reclaim": "off" } });
      // Collection runs and the scheduler checks once, with nothing due, so it has seen the
      // collecting sign-in. Then collection stops, and only then is an audit scheduled.
      await h.post("/v1/collection/start");
      await h.scheduler.tick();
      h.timers.fire("after");
      await settle(h);
      await h.post("/v1/collection/stop");
      await h.post("/v1/settings/audit-schedule", { audits: { "licenses.reclaim": "daily" } });
      h.timers.fire("every");
      await settle(h);
      // The stopped sign-in is never used for a scheduled run.
      expect(await runsOf(h, "licenses.reclaim")).toEqual([]);
      expect(h.reads.filter((read) => read.path.startsWith("/v1.0/users"))).toEqual([]);
    } finally {
      await h.app.close();
    }
  }, 60_000);

  it("skip an audit the collecting sign-in has not granted, whoever scheduled it", async () => {
    // The other administrator may run the inactive accounts audit and schedules it; the one who
    // collects has not granted AuditLog.Read.All.
    const h = await auditHarness({
      scopes: ["User.Read", "User.Read.All"],
      otherScopes: ["User.Read", "User.Read.All", "AuditLog.Read.All"],
    });
    try {
      const other = await h.signIn(otherActorId);
      const saved = await h.post(
        "/v1/settings/audit-schedule",
        { audits: { "accounts.inactive": "daily" } },
        other,
      );
      expect(saved.statusCode).toBe(200);
      await h.post("/v1/collection/start");
      await h.scheduler.tick();
      h.timers.fire("after");
      await settle(h);
      expect(await runsOf(h, "accounts.inactive")).toEqual([]);
      expect(h.logs).toContain(
        "Prism365 scheduled audit accounts.inactive skipped (missing_permission).",
      );
    } finally {
      await h.app.close();
    }
  }, 60_000);

  it("run a due audit on the collecting sign-in, naming who scheduled it", async () => {
    const h = await auditHarness();
    try {
      const other = await h.signIn(otherActorId);
      await h.post(
        "/v1/settings/audit-schedule",
        { audits: { "licenses.reclaim": "daily" } },
        other,
      );
      await h.post("/v1/collection/start");
      await h.scheduler.tick();
      h.timers.fire("after");
      await vi.waitFor(async () => {
        expect((await runsOf(h, "licenses.reclaim")).map((run) => run.status)).toEqual([
          "succeeded",
        ]);
      });
      const [run] = await runsOf(h, "licenses.reclaim");
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
            runId: run?.id,
            datasets: ["User.Read.All", "AuditLog.Read.All"],
            trigger: "schedule",
            scheduledBy: otherActorId,
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
        [run?.id, "schedule"],
      ]);
      // Not due again until a day after it started.
      h.timers.fire("every");
      await settle(h);
      expect(await runsOf(h, "licenses.reclaim")).toHaveLength(1);
    } finally {
      await h.app.close();
    }
  }, 60_000);

  it("stop reading the moment collection is stopped", async () => {
    const h = await auditHarness();
    try {
      await h.post("/v1/settings/audit-schedule", { audits: { "licenses.reclaim": "daily" } });
      await h.post("/v1/collection/start");
      await h.scheduler.tick();
      const hold = h.holdUsers();
      h.timers.fire("after");
      await vi.waitFor(() => expect(hold.waiting()).toBe(1));
      const stop = await h.post("/v1/collection/stop");
      expect(stop.json()).toMatchObject({ state: "stopped" });
      const reads = h.reads.length;
      const tokens = h.tokens.length;
      hold.release();
      await settle(h);
      // Nothing more was read, and no token was asked for, after Stop.
      expect(h.reads.slice(reads)).toEqual([]);
      expect(h.tokens.slice(tokens)).toEqual([]);
      const [run] = await runsOf(h, "licenses.reclaim");
      expect(run).toMatchObject({
        status: "failed",
        trigger: "schedule",
        error: {
          code: "collection_stopped",
          message: "Collection stopped while this scheduled run was reading.",
        },
      });
    } finally {
      await h.app.close();
    }
  }, 60_000);

  it("stop reading on the old sign-in when another administrator starts collecting", async () => {
    const h = await auditHarness();
    try {
      const other = await h.signIn(otherActorId);
      await h.post("/v1/settings/audit-schedule", { audits: { "licenses.reclaim": "daily" } });
      await h.post("/v1/collection/start");
      await h.scheduler.tick();
      const hold = h.holdUsers();
      h.timers.fire("after");
      await vi.waitFor(() => expect(hold.waiting()).toBe(1));
      const takeover = await h.post("/v1/collection/start", undefined, other);
      expect(takeover.statusCode).toBe(200);
      const reads = h.reads.length;
      const tokens = h.tokens.length;
      hold.release();
      await settle(h);
      await h.scheduler.tick();
      // The run on the first administrator's sign-in read nothing more; collection itself now
      // reads with the other administrator's.
      expect(h.reads.slice(reads).filter((read) => read.path.startsWith("/v1.0/users"))).toEqual(
        [],
      );
      expect(h.tokens.slice(tokens).filter((actor) => actor === actorId)).toEqual([]);
      const [run] = await runsOf(h, "licenses.reclaim");
      expect(run).toMatchObject({
        status: "failed",
        startedBy: { id: actorId },
        error: { code: "collection_stopped" },
      });
    } finally {
      await h.app.close();
    }
  }, 60_000);

  it("never read more than the administrator who scheduled the audit could", async () => {
    // The other administrator has not granted AuditLog.Read.All: their own guest review reads no
    // sign-in dates, so the one they schedule must not either, on anyone's sign-in.
    const h = await auditHarness({ otherScopes: ["User.Read", "User.Read.All"] });
    try {
      const other = await h.signIn(otherActorId);
      const saved = await h.post(
        "/v1/settings/audit-schedule",
        { audits: { "guests.review": "daily" } },
        other,
      );
      expect(saved.statusCode).toBe(200);
      expect(saved.json().limits).toEqual({
        "guests.review": { by: "Other admin", without: ["AuditLog.Read.All"] },
      });
      await h.post("/v1/collection/start");
      await h.scheduler.tick();
      h.timers.fire("after");
      await vi.waitFor(async () => {
        expect((await runsOf(h, "guests.review")).map((run) => run.status)).toEqual(["succeeded"]);
      });
      const users = h.reads.filter((read) => read.path === "/v1.0/users");
      expect(users.length).toBeGreaterThan(0);
      expect(users.every((read) => !(read.select ?? "").includes("signInActivity"))).toBe(true);
      const started = h.store.audit.find((entry) => entry.event.type === "audit.run.started");
      expect(started?.event).toMatchObject({
        datasets: ["User.Read.All"],
        trigger: "schedule",
        scheduledBy: otherActorId,
      });
      // Scheduled by the administrator collecting, who has granted it, it reads them.
      expect(
        (
          await h.post("/v1/settings/audit-schedule", { audits: { "guests.review": "weekly" } })
        ).json().limits,
      ).toEqual({});
    } finally {
      await h.app.close();
    }
  }, 60_000);

  it("stop only the scheduled run when collection stops, never a manual one", async () => {
    const h = await auditHarness();
    try {
      await h.post("/v1/settings/audit-schedule", { audits: { "licenses.reclaim": "daily" } });
      await h.post("/v1/collection/start");
      await h.scheduler.tick();
      const hold = h.holdUsers();
      const manual = await h.post("/v1/audits/guests.review/runs");
      expect(manual.statusCode).toBe(202);
      h.timers.fire("after");
      await vi.waitFor(() => expect(hold.waiting()).toBe(2));
      await h.post("/v1/collection/stop");
      hold.release();
      await settle(h);
      expect(await runsOf(h, "guests.review")).toMatchObject([
        { status: "succeeded", trigger: "manual" },
      ]);
      expect(await runsOf(h, "licenses.reclaim")).toMatchObject([
        { status: "failed", trigger: "schedule", error: { code: "collection_stopped" } },
      ]);
    } finally {
      await h.app.close();
    }
  }, 60_000);

  it("keep reading when the collecting administrator presses Start again", async () => {
    const h = await auditHarness();
    try {
      await h.post("/v1/settings/audit-schedule", { audits: { "licenses.reclaim": "daily" } });
      await h.post("/v1/collection/start");
      await h.scheduler.tick();
      const boundAt = (await h.get("/v1/modules")).json().collection.boundAt;
      const hold = h.holdUsers();
      h.timers.fire("after");
      await vi.waitFor(() => expect(hold.waiting()).toBe(1));
      const again = await h.post("/v1/collection/start");
      expect(again.statusCode).toBe(200);
      hold.release();
      await settle(h);
      expect(await runsOf(h, "licenses.reclaim")).toMatchObject([
        { status: "succeeded", trigger: "schedule" },
      ]);
      // The same binding, and the start recorded as before.
      expect((await h.get("/v1/modules")).json().collection).toMatchObject({
        state: "collecting",
        boundAt,
      });
      expect(
        h.store.audit.filter((entry) => entry.event.type === "collection.started"),
      ).toHaveLength(2);
    } finally {
      await h.app.close();
    }
  }, 60_000);

  it("tell an administrator whose Run joined a scheduled run that it stops with collection", async () => {
    const h = await auditHarness();
    try {
      await h.post("/v1/settings/audit-schedule", { audits: { "licenses.reclaim": "daily" } });
      await h.post("/v1/collection/start");
      await h.scheduler.tick();
      const hold = h.holdUsers();
      h.timers.fire("after");
      await vi.waitFor(() => expect(hold.waiting()).toBe(1));
      const other = await h.signIn(otherActorId);
      const joined = await h.post("/v1/audits/licenses.reclaim/runs", undefined, other);
      expect(joined.statusCode).toBe(200);
      expect(joined.json()).toMatchObject({
        started: false,
        message:
          "This audit was already running on its schedule, so your Run joined that run. It stops if collection stops; run it again then.",
      });
      hold.release();
      await settle(h);
      // A manual run started by itself says nothing of the kind.
      const own = await h.post("/v1/audits/guests.review/runs", undefined, other);
      expect(own.json()).toMatchObject({ started: true, message: null });
      await settle(h);
    } finally {
      await h.app.close();
    }
  }, 60_000);

  it("fail a run whose sign-in Microsoft stops accepting, the way a manual run fails", async () => {
    const h = await auditHarness();
    try {
      await h.post("/v1/settings/audit-schedule", {
        audits: { "licenses.reclaim": "daily", "guests.review": "daily" },
      });
      await h.post("/v1/collection/start");
      await h.scheduler.tick();
      h.refuseTokens(true);
      h.timers.fire("after");
      await vi.waitFor(async () => {
        expect((await runsOf(h, "licenses.reclaim")).map((run) => run.status)).toEqual(["failed"]);
      });
      await settle(h);
      const [run] = await runsOf(h, "licenses.reclaim");
      expect(run).toMatchObject({ trigger: "schedule", error: { code: "session" } });
      expect(h.reads.filter((read) => read.path.startsWith("/v1.0/users"))).toEqual([]);
      // The refusal ends the pass: the next scheduled audit is neither started nor ledgered.
      expect(await runsOf(h, "guests.review")).toEqual([]);
      expect(
        h.store.audit.filter(
          (entry) =>
            entry.event.type === "audit.run.started" && entry.event.auditId === "guests.review",
        ),
      ).toEqual([]);
    } finally {
      await h.app.close();
    }
  }, 60_000);
});

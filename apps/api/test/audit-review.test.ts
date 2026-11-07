import { randomUUID } from "node:crypto";
import {
  type Audit,
  auditDefinition,
  auditReadScope,
  builtInAudits,
  finishRun,
  insertRun,
  MAX_REPORT_DOWNLOADS,
  REPORT_WINDOW_MILLISECONDS,
  REQUEST_SPACING_MILLISECONDS,
  REVIEW_ORDER,
} from "@prism365/audits";
import { type AuditAvailability, COMBINED_RISKS_AUDIT_ID } from "@prism365/contracts";
import { GraphHttpClient } from "@prism365/graph";
import { PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import {
  AuditReviewer,
  interrupted,
  readLastReview,
  readReviewSchedule,
  type ReviewStartInput,
  ReviewRunningError,
  reviewDue,
  reviewSteps,
  storedReviewSchedule,
  writeLastReview,
  writeReviewSchedule,
} from "../src/audit-review.ts";
import { AuditScheduler, type ScheduleCredential } from "../src/audit-schedule.ts";
import type { AuditRunner } from "../src/audits.ts";
import { actorId, manualTimers, otherActorId } from "./audit-harness.ts";

// The full review on a fake clock with a fake runner, over a real in-memory database. Synthetic
// tenants and administrators only.

const MINUTE = 60_000;
type StartInput = Parameters<AuditRunner["start"]>[0];

let database: PGliteModuleDatabase;
beforeAll(async () => {
  database = await PGliteModuleDatabase.open();
}, 60_000);
afterAll(async () => {
  await database?.close();
});

// Records each start, stores the run as the real runner would, and finishes it after the time the
// test gives it on the fake clock. An audit named in `hold` waits until released or stopped.
function fakeRunner(clock: { now: Date; advance(ms: number): void }) {
  const starts: { audit: string; at: number; trigger: string | undefined }[] = [];
  const stopped: { runId: string; reason: string | undefined }[] = [];
  const outcomes = new Map<string, "succeeded" | "failed">();
  const hold = new Set<string>();
  const held = new Map<string, () => void>();
  let running = 0;
  let mostAtOnce = 0;
  let takes = 0;
  const finish = async (
    input: StartInput,
    id: string,
    status: "succeeded" | "failed",
    message = "Synthetic failure.",
  ) => {
    clock.advance(takes);
    await database.tenant(input.tenantId, (sql) =>
      finishRun(
        sql,
        input.tenantId,
        id,
        clock.now,
        status === "succeeded"
          ? { status, complete: true, summary: [], notes: [], rows: [] }
          : { status, code: "failed", message },
      ),
    );
    running -= 1;
  };
  const runner: Pick<AuditRunner, "start" | "stopRun"> = {
    stopRun: (runId, reason) => {
      stopped.push({ runId, reason });
      const release = held.get(runId);
      if (!release) return false;
      held.delete(runId);
      release();
      return true;
    },
    async start(input) {
      const id = randomUUID();
      const run = await database.tenant(input.tenantId, (sql) =>
        insertRun(sql, input.tenantId, {
          id,
          auditId: input.audit.definition.id,
          auditVersion: input.audit.definition.version,
          startedAt: clock.now,
          startedBy: input.actor,
          trigger: input.trigger ?? "manual",
        }),
      );
      running += 1;
      mostAtOnce = Math.max(mostAtOnce, running);
      starts.push({
        audit: input.audit.definition.id,
        at: clock.now.getTime(),
        trigger: input.trigger,
      });
      await input.beforeRead?.(run);
      const auditId = input.audit.definition.id;
      let done: Promise<void>;
      if (hold.has(auditId))
        done = new Promise<void>((resolve) => {
          held.set(id, () => {
            void finish(input, id, "failed", "The full review was stopped.").then(resolve);
          });
        });
      else done = finish(input, id, outcomes.get(auditId) ?? "succeeded");
      return { runId: id, started: true, run, done };
    },
  };
  return {
    runner,
    starts,
    stopped,
    hold,
    held,
    outcomes,
    ids: () => starts.map((item) => item.audit),
    mostAtOnce: () => mostAtOnce,
    takes: (milliseconds: number) => {
      takes = milliseconds;
    },
  };
}

function setup(options: { audits?: readonly Audit[] } = {}) {
  const tenantId = randomUUID();
  const clock = {
    now: new Date("2026-10-26T06:00:00Z"),
    advance(ms: number) {
      this.now = new Date(this.now.getTime() + ms);
    },
  };
  const fake = fakeRunner(clock);
  const waits: number[] = [];
  const reviewer = new AuditReviewer({
    runner: fake.runner,
    database,
    audits: options.audits ?? builtInAudits,
    now: () => clock.now,
    pause: async (milliseconds, signal) => {
      signal.throwIfAborted();
      waits.push(milliseconds);
      clock.advance(milliseconds);
    },
    log: () => {},
  });
  const ledger: Record<string, unknown>[] = [];
  const events: string[] = [];
  const input = (over: Partial<ReviewStartInput> = {}): ReviewStartInput => ({
    tenantId,
    actor: { id: actorId, name: "Synthetic admin" },
    trigger: "manual",
    unavailable: () => null,
    reading: async () => ({
      client: new GraphHttpClient({ getAccessToken: async () => "synthetic" }),
      packs: [],
      granted: [],
      beforeRead: async (run) => {
        events.push(`run ${run.auditId}`);
      },
    }),
    ledger: async (event) => {
      ledger.push(event);
      events.push(String(event.type));
    },
    ...over,
  });
  return { tenantId, clock, fake, reviewer, waits, ledger, events, input };
}

const ORDER = REVIEW_ORDER.map((step) => step.auditId);

describe("the order of a review", () => {
  it("runs the directory audits, then the paced ones, then the reports, and Combined risks last", async () => {
    const s = setup();
    const { review, done } = await s.reviewer.start(s.input());
    await done;
    expect(s.fake.ids()).toEqual(ORDER);
    expect(s.fake.ids().at(-1)).toBe(COMBINED_RISKS_AUDIT_ID);
    const stages = REVIEW_ORDER.map((step) => step.stage);
    expect(stages).toEqual(
      [...stages].sort(
        (a, b) =>
          ["directory", "paced", "reports", "summary"].indexOf(a) -
          ["directory", "paced", "reports", "summary"].indexOf(b),
      ),
    );
    const last = await s.reviewer.last(s.tenantId);
    expect(last?.id).toBe(review.id);
    expect(last?.state).toBe("finished");
    expect(last?.audits.map((item) => item.state)).toEqual(ORDER.map(() => "done"));
    expect(last?.audits.map((item) => item.auditId)).toEqual(ORDER);
  });

  it("runs the audits one at a time", async () => {
    const s = setup();
    const { done } = await s.reviewer.start(s.input());
    await done;
    expect(s.fake.mostAtOnce()).toBe(1);
  });

  it("starts every run with the review's trigger and the starting administrator", async () => {
    const s = setup();
    const { done } = await s.reviewer.start(s.input({ trigger: "schedule" }));
    await done;
    expect(new Set(s.fake.starts.map((item) => item.trigger))).toEqual(new Set(["schedule"]));
  });

  it("puts an audit the order does not list with the directory audits", () => {
    const extra: Audit = {
      ...builtInAudits[0],
      definition: { ...(builtInAudits[0] as Audit).definition, id: "extra.audit" },
    } as Audit;
    const ids = reviewSteps([...builtInAudits, extra]).map(({ audit }) => audit.definition.id);
    expect(ids.indexOf("extra.audit")).toBeLessThan(ids.indexOf("identity.conditional-access"));
    expect(ids.at(-1)).toBe(COMBINED_RISKS_AUDIT_ID);
  });
});

describe("audits that cannot run", () => {
  it("skips them with the reason and runs the rest", async () => {
    const s = setup();
    const refused = new Set(["identity.risky-users", "mail.mailbox-storage", "devices.managed"]);
    const { review, done } = await s.reviewer.start(
      s.input({
        unavailable: (audit) =>
          refused.has(audit.definition.id) ? `No access to ${audit.definition.id}.` : null,
      }),
    );
    expect(review.audits.filter((item) => item.state === "skipped")).toHaveLength(3);
    await done;
    expect(s.fake.ids()).toEqual(ORDER.filter((id) => !refused.has(id)));
    const last = await s.reviewer.last(s.tenantId);
    const skipped = last?.audits.filter((item) => item.state === "skipped") ?? [];
    expect(skipped.map((item) => [item.auditId, item.reason])).toEqual(
      ORDER.filter((id) => refused.has(id)).map((id) => [id, `No access to ${id}.`]),
    );
    expect(last?.state).toBe("finished");
  });

  it("records how many audits will run and how many are left out in the ledger", async () => {
    const s = setup();
    const { done } = await s.reviewer.start(
      s.input({ unavailable: (audit) => (audit.definition.id === "guests.review" ? "No." : null) }),
    );
    await done;
    expect(s.ledger[0]).toMatchObject({
      type: "audit.review.started",
      audits: ORDER.length - 1,
      skipped: 1,
      trigger: "manual",
    });
    expect(JSON.stringify(s.ledger[0])).not.toContain("guests.review");
  });
});

describe("one review per tenant", () => {
  it("refuses a second review while one runs, and allows another tenant's", async () => {
    const s = setup();
    s.fake.hold.add("accounts.inactive");
    const first = await s.reviewer.start(s.input());
    await expect(s.reviewer.start(s.input())).rejects.toBeInstanceOf(ReviewRunningError);
    expect(s.reviewer.running(s.tenantId)).toBe(true);
    const other = await s.reviewer.start(s.input({ tenantId: randomUUID() }));
    s.fake.hold.clear();
    for (const release of s.fake.held.values()) release();
    await Promise.all([first.done, other.done]);
    expect(s.reviewer.running(s.tenantId)).toBe(false);
    const again = await s.reviewer.start(s.input());
    await again.done;
  });
});

describe("the access ledger", () => {
  it("is written before any audit is run", async () => {
    const s = setup();
    const { done } = await s.reviewer.start(s.input());
    await done;
    expect(s.events[0]).toBe("audit.review.started");
    expect(s.events.indexOf("audit.review.started")).toBeLessThan(
      s.events.findIndex((item) => item.startsWith("run ")),
    );
    expect(s.events).toHaveLength(1 + ORDER.length);
  });

  it("holds counts and ids, never a name", async () => {
    const s = setup();
    const { done } = await s.reviewer.start(s.input());
    await done;
    expect(Object.keys(s.ledger[0] ?? {}).sort()).toEqual(
      ["audits", "occurredAt", "reviewId", "skipped", "trigger", "type"].sort(),
    );
  });

  it("starts nothing, and frees the tenant, when the entry cannot be written", async () => {
    const s = setup();
    await expect(
      s.reviewer.start(
        s.input({
          ledger: async () => {
            throw new Error("Ledger down.");
          },
        }),
      ),
    ).rejects.toThrow("Ledger down.");
    expect(s.fake.starts).toEqual([]);
    expect(s.reviewer.running(s.tenantId)).toBe(false);
    expect(await s.reviewer.last(s.tenantId)).toBeNull();
  });
});

describe("stopping a review", () => {
  it("stops the audit reading now, skips the rest and does not run Combined risks", async () => {
    const s = setup();
    s.fake.hold.add("licenses.capacity");
    const { done } = await s.reviewer.start(s.input());
    await vi_waitFor(() => s.fake.held.size === 1);
    expect(s.reviewer.stop(s.tenantId)).toBe(true);
    expect((await s.reviewer.last(s.tenantId))?.state).toBe("stopping");
    await done;
    expect(s.fake.stopped).toHaveLength(1);
    expect(s.fake.stopped[0]?.reason).toBe("review");
    expect(s.fake.ids().at(-1)).toBe("licenses.capacity");
    const last = await s.reviewer.last(s.tenantId);
    expect(last?.state).toBe("stopped");
    expect(last?.finishedAt).not.toBeNull();
    const byId = new Map(last?.audits.map((item) => [item.auditId, item]));
    expect(byId.get("licenses.capacity")?.state).toBe("failed");
    expect(byId.get(COMBINED_RISKS_AUDIT_ID)).toMatchObject({
      state: "skipped",
      reason: "The review was stopped before this audit started.",
    });
    expect(byId.get("accounts.inactive")?.state).toBe("done");
  });

  it("says false when no review is running", () => {
    const s = setup();
    expect(s.reviewer.stop(s.tenantId)).toBe(false);
  });

  it("stops a scheduled review when collection ends, and leaves an administrator's alone", async () => {
    const s = setup();
    s.fake.hold.add("accounts.inactive");
    const manual = await s.reviewer.start(s.input());
    await vi_waitFor(() => s.fake.held.size === 1);
    s.reviewer.stopScheduled(s.tenantId);
    expect(s.fake.stopped).toEqual([]);
    s.fake.hold.clear();
    for (const release of s.fake.held.values()) release();
    await manual.done;

    s.fake.hold.add("accounts.inactive");
    const scheduled = await s.reviewer.start(s.input({ trigger: "schedule" }));
    await vi_waitFor(() => s.fake.held.size === 1);
    s.reviewer.stopScheduled(s.tenantId);
    await scheduled.done;
    expect(s.fake.stopped).toHaveLength(1);
    expect((await s.reviewer.last(s.tenantId))?.state).toBe("stopped");
  });

  it("ends the review as stopped when the sign-in it reads with is gone", async () => {
    const s = setup();
    let asked = 0;
    const base = s.input();
    const { done } = await s.reviewer.start({
      ...base,
      reading: async (audit) => (++asked > 2 ? null : base.reading(audit)),
    });
    await done;
    expect(s.fake.ids()).toEqual(ORDER.slice(0, 2));
    const last = await s.reviewer.last(s.tenantId);
    expect(last?.state).toBe("stopped");
    expect(last?.audits[2]).toMatchObject({
      state: "skipped",
      reason: "The sign-in this review used ended.",
    });
  });
});

describe("a failed audit", () => {
  it("is shown with its reason, and the review carries on to Combined risks", async () => {
    const s = setup();
    s.fake.outcomes.set("accounts.passwords", "failed");
    const { done } = await s.reviewer.start(s.input());
    await done;
    const last = await s.reviewer.last(s.tenantId);
    const failed = last?.audits.find((item) => item.auditId === "accounts.passwords");
    expect(failed).toMatchObject({ state: "failed", reason: "Synthetic failure." });
    expect(s.fake.ids().at(-1)).toBe(COMBINED_RISKS_AUDIT_ID);
    expect(last?.state).toBe("finished");
  });
});

describe("Microsoft's limits", () => {
  it("waits out the gap between the paced audits and before Combined risks", async () => {
    const s = setup();
    const { done } = await s.reviewer.start(
      s.input({
        unavailable: (a) =>
          REVIEW_ORDER.find((step) => step.auditId === a.definition.id)?.stage === "reports"
            ? "Off."
            : null,
      }),
    );
    await done;
    // Conditional Access and Risky users end at the same fake instant: each following audit that
    // reads a paced endpoint waits the whole spacing.
    expect(s.waits).toEqual([REQUEST_SPACING_MILLISECONDS, REQUEST_SPACING_MILLISECONDS]);
    const at = (id: string) => s.fake.starts.find((item) => item.audit === id)?.at ?? 0;
    expect(at("identity.risky-users") - at("identity.conditional-access")).toBe(
      REQUEST_SPACING_MILLISECONDS,
    );
    expect(at(COMBINED_RISKS_AUDIT_ID) - at("identity.risky-users")).toBe(
      REQUEST_SPACING_MILLISECONDS,
    );
  });

  it("needs no wait for Combined risks when the reports took longer than the gap", async () => {
    const s = setup();
    s.fake.takes(MINUTE);
    const { done } = await s.reviewer.start(s.input());
    await done;
    // Only the gap after Conditional Access is owed: Risky users follows it at once, and by the
    // time Combined risks starts the four report audits have taken minutes.
    expect(s.waits).toEqual([REQUEST_SPACING_MILLISECONDS]);
  });

  it("starts no more than 12 report downloads in any ten minutes, across reviews", async () => {
    const s = setup();
    s.fake.takes(10_000);
    for (let round = 0; round < 3; round += 1) {
      const { done } = await s.reviewer.start(s.input());
      await done;
    }
    const cost = new Map(REVIEW_ORDER.map((step) => [step.auditId, step.reportDownloads]));
    const reports = s.fake.starts.filter((item) => (cost.get(item.audit) ?? 0) > 0);
    expect(reports).toHaveLength(12);
    for (const start of reports) {
      const inWindow = reports.filter(
        (other) => other.at <= start.at && other.at > start.at - REPORT_WINDOW_MILLISECONDS,
      );
      const total = inWindow.reduce((sum, other) => sum + (cost.get(other.audit) ?? 0), 0);
      expect(total, `at ${new Date(start.at).toISOString()}`).toBeLessThanOrEqual(
        MAX_REPORT_DOWNLOADS,
      );
    }
    // A second review straight after the first had to wait for the first one's reports to age.
    expect(s.waits.some((wait) => wait > MINUTE)).toBe(true);
  });

  it("does not wait on the first review of a tenant, whose reports fit the limit", async () => {
    const s = setup();
    const { done } = await s.reviewer.start(s.input());
    await done;
    expect(s.waits.every((wait) => wait === REQUEST_SPACING_MILLISECONDS)).toBe(true);
  });
});

describe("the review as stored", () => {
  it("is read back by a reviewer that never ran it, as a restarted server would", async () => {
    const s = setup();
    const { review, done } = await s.reviewer.start(s.input());
    await done;
    const fresh = new AuditReviewer({
      runner: s.fake.runner,
      database,
      audits: builtInAudits,
      log: () => {},
    });
    const read = await fresh.last(s.tenantId);
    expect(read?.id).toBe(review.id);
    expect(read?.state).toBe("finished");
    expect((await fresh.lastStartedAt(s.tenantId))?.toISOString()).toBe(review.startedAt);
  });

  it("shows a review its process left running as stopped, with what it never finished marked", async () => {
    const s = setup();
    const { review, done } = await s.reviewer.start(s.input());
    await done;
    const crashed = {
      ...review,
      state: "running" as const,
      finishedAt: null,
      audits: review.audits.map((item, index) => ({
        ...item,
        state:
          index === 0
            ? ("done" as const)
            : index === 1
              ? ("running" as const)
              : ("waiting" as const),
        reason: null,
        finishedAt: index === 0 ? "2026-10-26T06:01:00.000Z" : null,
      })),
    };
    await database.tenant(s.tenantId, (sql) => writeLastReview(sql, s.tenantId, crashed, actorId));
    const read = await new AuditReviewer({
      runner: s.fake.runner,
      database,
      audits: builtInAudits,
      log: () => {},
    }).last(s.tenantId);
    expect(read?.state).toBe("stopped");
    expect(read?.finishedAt).toBe("2026-10-26T06:01:00.000Z");
    expect(read?.audits[1]).toMatchObject({
      state: "failed",
      reason: "Prism365 stopped while this audit was running.",
    });
    expect(read?.audits[2]).toMatchObject({ state: "skipped" });
    expect(interrupted(read as never)).toEqual(read);
  });

  it("reads nothing back from a value that is not a review", async () => {
    const s = setup();
    await database.tenant(s.tenantId, (sql) =>
      sql.query(
        `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
         VALUES ($1,'audits.review.last','{"state":"finished"}'::jsonb,now(),$2)`,
        [s.tenantId, actorId],
      ),
    );
    expect(await database.tenant(s.tenantId, (sql) => readLastReview(sql, s.tenantId))).toBeNull();
  });
});

describe("the weekly schedule", () => {
  const weekly = (over: Partial<ReturnType<typeof storedReviewSchedule>> = {}) => ({
    ...storedReviewSchedule({ frequency: "weekly", weekday: 0, hour: 6 }),
    ...over,
  });
  // Monday 2026-10-26 06:00 in Toronto is 10:00 UTC (daylight time ends on 11-01).
  const zone = "America/Toronto";

  it("is never due while off", () => {
    expect(reviewDue(storedReviewSchedule({}), null, new Date("2026-10-27T00:00:00Z"), zone)).toBe(
      false,
    );
    expect(
      reviewDue(weekly({ frequency: "off" }), null, new Date("2026-10-27T00:00:00Z"), zone),
    ).toBe(false);
  });

  it("is due once the chosen day and hour in the time zone have passed since the last review", () => {
    const slot = new Date("2026-10-26T10:00:00Z");
    expect(reviewDue(weekly(), null, new Date(slot.getTime() - 1), zone)).toBe(true);
    // Before the slot, the previous Monday's slot is the one that counts, and no review since.
    expect(
      reviewDue(weekly(), new Date("2026-10-19T10:00:00Z"), new Date(slot.getTime() - 1), zone),
    ).toBe(false);
    expect(
      reviewDue(weekly(), new Date("2026-10-19T09:59:00Z"), new Date(slot.getTime() - 1), zone),
    ).toBe(true);
    expect(reviewDue(weekly(), new Date("2026-10-19T10:00:00Z"), slot, zone)).toBe(true);
    expect(reviewDue(weekly(), slot, new Date(slot.getTime() + 6 * 86_400_000), zone)).toBe(false);
    // The next Monday 06:00 is an hour later in UTC, daylight time having ended on the Sunday.
    expect(reviewDue(weekly(), slot, new Date("2026-11-02T11:00:00Z"), zone)).toBe(true);
  });

  it("counts a review an administrator started after the slot as that week's", () => {
    const now = new Date("2026-10-28T12:00:00Z");
    expect(reviewDue(weekly(), new Date("2026-10-27T09:00:00Z"), now, zone)).toBe(false);
  });

  it("does not run last Monday's review at once for a schedule saved afterwards", () => {
    const now = new Date("2026-10-28T12:00:00Z");
    expect(reviewDue(weekly({ savedAt: "2026-10-28T11:00:00.000Z" }), null, now, zone)).toBe(false);
    expect(reviewDue(weekly({ savedAt: "2026-10-26T09:00:00.000Z" }), null, now, zone)).toBe(true);
  });

  it("follows the daylight saving change when the hour is on the other side of it", () => {
    // 06:00 on Monday 2026-11-02 in Toronto is 11:00 UTC, once daylight time has ended.
    expect(reviewDue(weekly(), null, new Date("2026-11-02T10:59:00Z"), zone)).toBe(true);
    expect(
      reviewDue(weekly(), new Date("2026-10-26T10:00:00Z"), new Date("2026-11-02T10:59:00Z"), zone),
    ).toBe(false);
    expect(
      reviewDue(weekly(), new Date("2026-10-26T10:00:00Z"), new Date("2026-11-02T11:00:00Z"), zone),
    ).toBe(true);
  });

  it("saves and reads the schedule with the permissions of whoever saved it", async () => {
    const tenant = randomUUID();
    const saved = await database.tenant(tenant, (sql) =>
      writeReviewSchedule(
        sql,
        tenant,
        { frequency: "weekly", weekday: 2, hour: 22 },
        otherActorId,
        {
          name: "Other admin",
          grants: { "guests.review": ["User.Read.All"] },
          now: new Date("2026-10-26T00:00:00Z"),
        },
      ),
    );
    expect(saved).toMatchObject({
      frequency: "weekly",
      weekday: 2,
      hour: 22,
      scheduledBy: otherActorId,
      name: "Other admin",
      grants: { "guests.review": ["User.Read.All"] },
    });
    expect(await database.tenant(tenant, (sql) => readReviewSchedule(sql, tenant))).toEqual(saved);
    const off = await database.tenant(tenant, (sql) =>
      writeReviewSchedule(sql, tenant, { frequency: "off", weekday: 2, hour: 22 }, otherActorId, {
        name: "Other admin",
        grants: {},
        now: new Date(),
      }),
    );
    expect(off).toMatchObject({ frequency: "off", scheduledBy: null, grants: {} });
  });

  it("reads a damaged value as off", () => {
    for (const value of [
      null,
      "x",
      { frequency: "weekly" },
      { frequency: "weekly", weekday: 9, hour: 1 },
      { frequency: "daily", weekday: 1, hour: 1 },
    ])
      expect(storedReviewSchedule(value).frequency).toBe("off");
  });
});

// A scheduled review through the scheduler, with the real reviewer over the fake runner.
describe("the scheduler and the full review", () => {
  const availabilityOf =
    (granted: string[]) =>
    (audit: Audit, scopes: readonly string[]): AuditAvailability => {
      const missing = auditDefinition(audit).permissions.filter(
        (permission) => !scopes.includes(permission) && !granted.includes(permission),
      );
      return missing.length > 0
        ? { state: "missing_permission", missing, grantFeature: null, enable: [] }
        : { state: "available", missing: [], grantFeature: null, enable: [] };
    };

  async function scheduled(options: { savedAt?: string; weekday?: number } = {}) {
    const s = setup();
    const credential: ScheduleCredential = {
      tenantId: s.tenantId,
      actorId,
      displayName: "Synthetic admin",
      grantedPermissions: async () => ["User.Read", "User.Read.All", "AuditLog.Read.All"],
      accessToken: async () => "synthetic",
    };
    const ledger: Record<string, unknown>[] = [];
    const logs: string[] = [];
    const timers = manualTimers();
    const scheduler = new AuditScheduler({
      runner: { ...s.fake.runner, runningId: () => null },
      database,
      audits: builtInAudits,
      credential: () => credential,
      enabledPacks: ["directory.users.read", "audit.signins.read"],
      availability: availabilityOf([]),
      graphClient: (getAccessToken) => new GraphHttpClient({ getAccessToken }),
      ledger: async (_tenant, _actor, event) => {
        ledger.push(event);
      },
      now: () => s.clock.now,
      timers: timers.timers,
      log: (message) => logs.push(message),
      review: { reviewer: s.reviewer, timeZone: "UTC" },
    });
    await database.tenant(s.tenantId, (sql) =>
      writeReviewSchedule(
        sql,
        s.tenantId,
        { frequency: "weekly", weekday: options.weekday ?? 0, hour: 5 },
        otherActorId,
        {
          name: "Other admin",
          grants: Object.fromEntries(
            builtInAudits.map((audit) => [
              audit.definition.id,
              auditReadScope(
                audit,
                ["directory.users.read", "audit.signins.read"],
                ["User.Read", "User.Read.All", "AuditLog.Read.All"],
              ).granted,
            ]),
          ),
          now: new Date(options.savedAt ?? "2026-10-01T00:00:00Z"),
        },
      ),
    );
    return { ...s, scheduler, ledger, logs, credential };
  }

  it("runs a review when the slot has passed, as a scheduled review, and not again that week", async () => {
    const s = await scheduled();
    await s.scheduler.tick();
    expect(s.fake.starts.length).toBeGreaterThan(0);
    expect(new Set(s.fake.starts.map((item) => item.trigger))).toEqual(new Set(["schedule"]));
    expect(s.ledger[0]).toMatchObject({
      type: "audit.review.started",
      trigger: "schedule",
      scheduledBy: otherActorId,
    });
    // The audits this installation cannot run are listed and skipped with a reason.
    const last = await s.reviewer.last(s.tenantId);
    expect(last?.trigger).toBe("schedule");
    expect(last?.audits.some((item) => item.state === "skipped" && item.reason)).toBe(true);
    expect(s.fake.ids().at(-1)).toBe(COMBINED_RISKS_AUDIT_ID);
    const before = s.fake.starts.length;
    await s.scheduler.tick();
    expect(s.fake.starts).toHaveLength(before);
  });

  it("does not run before the slot", async () => {
    // It is Monday 06:00; the schedule is for Wednesday and was saved after the previous one.
    const s = await scheduled({ weekday: 2, savedAt: "2026-10-22T00:00:00Z" });
    await s.scheduler.tick();
    expect(s.fake.starts).toEqual([]);
  });

  it("is skipped while a review an administrator started is running", async () => {
    const s = await scheduled();
    // The administrator's review began at 04:30, before this week's 05:00 slot, and is still going
    // at 05:30: a review is due, and it is not run on top of the one in progress.
    s.clock.now = new Date("2026-10-26T04:30:00Z");
    s.fake.hold.add("accounts.inactive");
    const manual = await s.reviewer.start(s.input());
    await vi_waitFor(() => s.fake.held.size === 1);
    s.clock.advance(60 * MINUTE);
    await s.scheduler.tick();
    expect(s.fake.starts).toHaveLength(1);
    expect(s.ledger).toEqual([]);
    expect(s.reviewer.running(s.tenantId)).toBe(true);
    s.fake.hold.clear();
    for (const release of s.fake.held.values()) release();
    await manual.done;
    // Once it has ended, the review that was due runs, as a scheduled one.
    await s.scheduler.tick();
    expect(s.ledger.map((event) => event.type)).toContain("audit.review.started");
  });

  it("reads each audit with no more than the permissions the saving administrator had", async () => {
    const s = await scheduled();
    await database.tenant(s.tenantId, (sql) =>
      writeReviewSchedule(
        sql,
        s.tenantId,
        { frequency: "weekly", weekday: 0, hour: 5 },
        otherActorId,
        {
          name: "Other admin",
          grants: { "accounts.inactive": ["User.Read.All"] },
          now: new Date("2026-10-01T00:00:00Z"),
        },
      ),
    );
    await s.scheduler.tick();
    const last = await s.reviewer.last(s.tenantId);
    const inactive = last?.audits.find((item) => item.auditId === "accounts.inactive");
    // The collecting sign-in can read sign-ins, but the saver's grants for this audit could not.
    expect(inactive?.state).toBe("skipped");
    expect(inactive?.reason).toContain("AuditLog.Read.All");
    expect(s.fake.ids()).not.toContain("accounts.inactive");
    // An audit the saver has no entry for reads with its required permissions only.
    expect(s.fake.ids()).toContain("licenses.reclaim");
  });

  it("stops a scheduled review when collection ends", async () => {
    const s = await scheduled();
    s.fake.hold.add("accounts.inactive");
    const tick = s.scheduler.tick();
    await vi_waitFor(() => s.fake.held.size === 1);
    s.scheduler.collectionEnded();
    await tick;
    expect(s.fake.stopped).toHaveLength(1);
    expect((await s.reviewer.last(s.tenantId))?.state).toBe("stopped");
  });
});

// Waits for a condition that the fake runner reaches on its own after a few turns of the loop.
async function vi_waitFor(condition: () => boolean): Promise<void> {
  for (let turn = 0; turn < 2_000; turn += 1) {
    if (condition()) return;
    await new Promise((resolve) => setTimeout(resolve, 2));
  }
  throw new Error("The condition was never reached.");
}

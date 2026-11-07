import { randomUUID } from "node:crypto";
import {
  type Audit,
  abortablePause,
  listRuns,
  type Pause,
  REQUEST_SPACING_MILLISECONDS,
  REVIEW_ORDER,
  type ReviewStep,
} from "@prism365/audits";
import {
  type AuditReview,
  type AuditReviewAudit,
  type AuditReviewSchedule,
  type AuditRun,
  type AuditRunTrigger,
  auditReviewScheduleSchema,
  auditReviewSchema,
  cleanText,
} from "@prism365/contracts";
import type { GraphFeatureId, GraphHttpClient } from "@prism365/graph";
import type { ModuleDatabase, SqlClient } from "@prism365/storage";
import { z } from "zod";
import { weeklySlot } from "./alert-rules.ts";
import type { AuditRunner } from "./audits.ts";
import { ReportBudget } from "./report-budget.ts";

// The full review (docs/audits.md, section 22): every audit the administrator can run, one at a
// time through the ordinary runner, in an order that keeps Microsoft's limits, finishing with
// Combined risks so it joins results from this review. The review lives in memory while it runs
// and its state is kept in the preference store (no table of its own), so the page can show the
// last review after a restart.

const REVIEW_KEY = "audits.review.last";
const SCHEDULE_KEY = "audits.review.schedule";

export class ReviewRunningError extends Error {
  constructor() {
    super("A full review is already running.");
    this.name = "ReviewRunningError";
  }
}

// What one audit of a review may read, asked for just before it starts. Different for a
// session's review (the administrator's sign-in) and a scheduled one (the collecting sign-in).
export interface ReviewReading {
  client: GraphHttpClient;
  packs: readonly GraphFeatureId[];
  granted: readonly string[];
  // The access-ledger entry for the audit's run, written once it is stored and before it reads.
  beforeRead: (run: AuditRun) => Promise<void>;
}

export interface ReviewStartInput {
  tenantId: string;
  actor: { id: string; name: string | null };
  trigger: AuditRunTrigger;
  // Why an audit cannot run on this sign-in, in plain words, or null when it can.
  unavailable(audit: Audit): string | null;
  // null ends the review: the sign-in it runs on is gone.
  reading(audit: Audit): Promise<ReviewReading | null>;
  // The access ledger, for the entry that says a review started.
  ledger(event: Record<string, unknown>): Promise<void>;
}

export interface AuditReviewerOptions {
  runner: Pick<AuditRunner, "start" | "stopRun">;
  database: ModuleDatabase;
  audits: readonly Audit[];
  now?: () => Date;
  // How the review waits. Tests replace it.
  pause?: Pause;
  log?: (message: string) => void;
}

interface Live {
  tenantId: string;
  review: AuditReview;
  steps: ReviewStep[];
  controller: AbortController;
  trigger: AuditRunTrigger;
  actorId: string;
  runId: string | null;
  // The run in progress was started by this review, so Stop may end it.
  ownsRun: boolean;
  stopRequested: boolean;
  stopReason: "review" | "shutdown";
  persisting: Promise<void>;
  done: Promise<void>;
}

const copy = (review: AuditReview): AuditReview => structuredClone(review);

// The audits in the order a review runs them. An audit the order does not list (a new one the
// order was not updated for, which a test catches) runs with the directory audits.
export function reviewSteps(audits: readonly Audit[]): { audit: Audit; step: ReviewStep }[] {
  const known = new Set(REVIEW_ORDER.map((step) => step.auditId));
  const byId = new Map(audits.map((audit) => [audit.definition.id, audit]));
  const unknown = audits
    .filter((audit) => !known.has(audit.definition.id))
    .map((audit) => ({
      audit,
      step: { auditId: audit.definition.id, stage: "directory", reportDownloads: 0 } as ReviewStep,
    }));
  const listed = REVIEW_ORDER.flatMap((step) => {
    const audit = byId.get(step.auditId);
    return audit ? [{ audit, step }] : [];
  });
  const firstLater = listed.findIndex(({ step }) => step.stage !== "directory");
  const at = firstLater < 0 ? listed.length : firstLater;
  return [...listed.slice(0, at), ...unknown, ...listed.slice(at)];
}

export class AuditReviewer {
  readonly #runner: AuditReviewerOptions["runner"];
  readonly #database: ModuleDatabase;
  readonly #audits: readonly Audit[];
  readonly #now: () => Date;
  readonly #pause: Pause;
  readonly #log: (message: string) => void;
  readonly #live = new Map<string, Live>();
  readonly #budgets = new Map<string, ReportBudget>();
  #closed = false;

  constructor(options: AuditReviewerOptions) {
    this.#runner = options.runner;
    this.#database = options.database;
    this.#audits = options.audits;
    this.#now = options.now ?? (() => new Date());
    this.#pause = options.pause ?? abortablePause;
    this.#log = options.log ?? ((message) => console.warn(message));
  }

  running(tenantId: string): boolean {
    return this.#live.has(tenantId);
  }

  // The review in progress, or the last one that ended: from memory while this process has it,
  // otherwise as stored. A stored review that was still running when its process stopped is shown
  // as stopped, with its unfinished audits marked, because nothing is running it now.
  async last(tenantId: string): Promise<AuditReview | null> {
    const live = this.#live.get(tenantId);
    if (live) return copy(live.review);
    const stored = await this.#database.tenant(tenantId, (sql) => readLastReview(sql, tenantId));
    return stored ? interrupted(stored) : null;
  }

  // When the newest review started, for the scheduled review's due check.
  async lastStartedAt(tenantId: string): Promise<Date | null> {
    const review = await this.last(tenantId);
    return review ? new Date(review.startedAt) : null;
  }

  // Starts a review and returns at once with its first state; `done` resolves when it has ended
  // and its final state is stored. Throws ReviewRunningError when this tenant has one running.
  async start(input: ReviewStartInput): Promise<{ review: AuditReview; done: Promise<void> }> {
    if (this.#closed) throw new Error("Prism365 is stopping.");
    if (this.#live.has(input.tenantId)) throw new ReviewRunningError();
    const steps = reviewSteps(this.#audits);
    const at = this.#now().toISOString();
    const audits: AuditReviewAudit[] = steps.map(({ audit }) => {
      const reason = input.unavailable(audit);
      return {
        auditId: audit.definition.id,
        state: reason === null ? "waiting" : "skipped",
        reason,
        runId: null,
        startedAt: null,
        finishedAt: null,
      };
    });
    const review: AuditReview = {
      id: randomUUID(),
      state: "running",
      trigger: input.trigger,
      startedAt: at,
      finishedAt: null,
      startedBy: { id: input.actor.id, name: input.actor.name },
      audits,
    };
    let release: () => void = () => {};
    const live: Live = {
      tenantId: input.tenantId,
      review,
      steps: steps.map(({ step }) => step),
      controller: new AbortController(),
      trigger: input.trigger,
      actorId: input.actor.id,
      runId: null,
      ownsRun: false,
      stopRequested: false,
      stopReason: "review",
      persisting: Promise.resolve(),
      done: new Promise<void>((resolve) => {
        release = resolve;
      }),
    };
    // Claimed before the first await, so a second request for this tenant is refused.
    this.#live.set(input.tenantId, live);
    try {
      // Before anything runs: the ledger names how many audits will run and how many are left out.
      await input.ledger({
        type: "audit.review.started",
        reviewId: review.id,
        trigger: input.trigger,
        audits: audits.filter((item) => item.state === "waiting").length,
        skipped: audits.filter((item) => item.state === "skipped").length,
        occurredAt: at,
      });
    } catch (error) {
      this.#live.delete(input.tenantId);
      release();
      throw error;
    }
    this.#persist(live);
    void this.#run(input, live)
      .catch((error: unknown) => {
        this.#log(
          `Prism365 full review ended early (${error instanceof Error ? error.name : typeof error}).`,
        );
        this.#end(live, "stopped", "The review ended early. The server log has the category.");
      })
      .then(async () => {
        await live.persisting;
        this.#live.delete(input.tenantId);
        release();
      });
    return { review: copy(review), done: live.done };
  }

  // Asks the review to stop: the audit reading now is stopped the way a scheduled run is when
  // collection ends, and the audits after it are skipped. false when none is running.
  stop(tenantId: string): boolean {
    const live = this.#live.get(tenantId);
    if (!live) return false;
    if (!live.stopRequested) this.#halt(live, "review");
    return true;
  }

  // Collection ended: a review a schedule started stops; a review an administrator started with
  // their own sign-in does not.
  stopScheduled(tenantId: string): void {
    const live = this.#live.get(tenantId);
    if (live?.trigger === "schedule") this.#halt(live, "review");
  }

  async close(): Promise<void> {
    this.#closed = true;
    const live = [...this.#live.values()];
    for (const item of live) this.#halt(item, "shutdown");
    await Promise.all(live.map((item) => item.done));
  }

  // Resolves when no review is running. For tests.
  async idle(): Promise<void> {
    await Promise.all([...this.#live.values()].map((item) => item.done));
  }

  #halt(live: Live, reason: "review" | "shutdown"): void {
    live.stopRequested = true;
    live.stopReason = reason;
    if (live.review.state === "running") live.review.state = "stopping";
    live.controller.abort();
    if (live.runId && live.ownsRun) this.#runner.stopRun(live.runId, reason);
  }

  #budget(tenantId: string): ReportBudget {
    let budget = this.#budgets.get(tenantId);
    if (!budget) {
      budget = new ReportBudget({ now: () => this.#now().getTime(), pause: this.#pause });
      this.#budgets.set(tenantId, budget);
    }
    return budget;
  }

  async #run(input: ReviewStartInput, live: Live): Promise<void> {
    const { tenantId } = input;
    const budget = this.#budget(tenantId);
    const signal = live.controller.signal;
    const byId = new Map(this.#audits.map((audit) => [audit.definition.id, audit]));
    // When the last Conditional Access or Risky users audit ended: those endpoints allow one
    // request a second for the tenant, so the next audit that reads one waits out the gap.
    let pacedEnd: number | null = null;
    let ending: { state: "finished" | "stopped"; reason: string } | null = null;
    for (const [index, step] of live.steps.entries()) {
      const item = live.review.audits[index];
      const audit = byId.get(step.auditId);
      if (!item || !audit || item.state !== "waiting") continue;
      if (signal.aborted) break;
      try {
        if ((step.stage === "paced" || step.stage === "summary") && pacedEnd !== null) {
          const wait = REQUEST_SPACING_MILLISECONDS - (this.#now().getTime() - pacedEnd);
          if (wait > 0) await this.#pause(wait, signal);
        }
        if (step.reportDownloads > 0) await budget.waitFor(step.reportDownloads, signal);
      } catch {
        break;
      }
      let reading: ReviewReading | null;
      try {
        reading = await input.reading(audit);
      } catch {
        reading = null;
      }
      if (!reading) {
        ending = { state: "stopped", reason: "The sign-in this review used ended." };
        break;
      }
      if (signal.aborted) break;
      item.state = "running";
      item.startedAt = this.#now().toISOString();
      this.#persist(live);
      let started: Awaited<ReturnType<AuditRunner["start"]>>;
      try {
        started = await this.#runner.start({
          tenantId,
          actor: input.actor,
          audit,
          client: reading.client,
          packs: reading.packs,
          granted: reading.granted,
          trigger: input.trigger,
          beforeRead: reading.beforeRead,
        });
      } catch {
        item.state = "failed";
        item.reason = "The audit could not be started. Run it on its own to see why.";
        item.finishedAt = this.#now().toISOString();
        this.#persist(live);
        continue;
      }
      item.runId = started.runId;
      live.runId = started.runId;
      live.ownsRun = started.started;
      // Stop was asked while the run was being stored.
      if (live.stopRequested && started.started)
        this.#runner.stopRun(started.runId, live.stopReason);
      // A run that was already going (someone started this audit on its own) is joined, not
      // stopped: Stop only stops waiting for it.
      await (started.started ? started.done : Promise.race([started.done, abortion(signal)]));
      live.runId = null;
      live.ownsRun = false;
      const outcome = await this.#outcome(tenantId, audit, started.runId);
      item.finishedAt = this.#now().toISOString();
      if (outcome?.status === "succeeded") {
        item.state = "done";
        item.reason = null;
      } else {
        item.state = "failed";
        item.reason =
          outcome?.error?.message ??
          (signal.aborted
            ? "The review stopped before this audit finished."
            : "The audit did not finish.");
      }
      budget.record(step.reportDownloads);
      if (step.stage === "paced") pacedEnd = this.#now().getTime();
      this.#persist(live);
    }
    this.#end(
      live,
      live.stopRequested ? "stopped" : (ending?.state ?? "finished"),
      ending?.reason ??
        (live.stopRequested
          ? "The review was stopped before this audit started."
          : "The review ended before this audit started."),
    );
  }

  async #outcome(tenantId: string, audit: Audit, runId: string): Promise<AuditRun | null> {
    try {
      const runs = await this.#database.tenant(tenantId, (sql) =>
        listRuns(sql, tenantId, audit.definition.id),
      );
      return runs.find((run) => run.id === runId) ?? null;
    } catch {
      return null;
    }
  }

  // Closes the review: audits that never started are skipped with the reason, and the final
  // state is stored.
  #end(live: Live, state: "finished" | "stopped", reason: string): void {
    const { review } = live;
    const at = this.#now().toISOString();
    for (const item of review.audits) {
      if (item.state === "waiting") {
        item.state = "skipped";
        item.reason = reason;
      } else if (item.state === "running") {
        item.state = "failed";
        item.reason = "The review ended while this audit was running.";
        item.finishedAt = at;
      }
    }
    review.state = state;
    review.finishedAt = at;
    this.#persist(live);
  }

  // Stores the review's present state. Writes follow each other in order and a failed one is
  // logged, not thrown: the review goes on, and the page shows it from memory.
  #persist(live: Live): void {
    const snapshot = copy(live.review);
    live.persisting = live.persisting.then(async () => {
      try {
        await this.#database.tenant(live.tenantId, (sql) =>
          writeLastReview(sql, live.tenantId, snapshot, live.actorId),
        );
      } catch {
        this.#log("Prism365 could not store the state of the full review.");
      }
    });
  }
}

// Resolves when the signal aborts.
function abortion(signal: AbortSignal): Promise<void> {
  return new Promise((resolve) => {
    if (signal.aborted) resolve();
    else signal.addEventListener("abort", () => resolve(), { once: true });
  });
}

// A stored review whose process is gone: nothing is running it, so it is stopped.
export function interrupted(review: AuditReview): AuditReview {
  if (review.state === "finished" || review.state === "stopped") return review;
  const ended = review.audits
    .map((item) => item.finishedAt ?? item.startedAt)
    .filter((value): value is string => value !== null)
    .sort();
  const at = ended.at(-1) ?? review.startedAt;
  return {
    ...review,
    state: "stopped",
    finishedAt: review.finishedAt ?? at,
    audits: review.audits.map((item) =>
      item.state === "waiting"
        ? { ...item, state: "skipped", reason: "Prism365 stopped before this audit started." }
        : item.state === "running"
          ? {
              ...item,
              state: "failed",
              reason: "Prism365 stopped while this audit was running.",
              finishedAt: item.finishedAt ?? at,
            }
          : item,
    ),
  };
}

// Storage ---------------------------------------------------------------------------------------

function parsedValue(value: unknown): unknown {
  if (typeof value !== "string") return value ?? null;
  try {
    return JSON.parse(value) as unknown;
  } catch {
    return null;
  }
}

export async function readLastReview(
  sql: SqlClient,
  tenantId: string,
): Promise<AuditReview | null> {
  const result = await sql.query<{ value: unknown }>(
    "SELECT value FROM prism365_preference WHERE tenant_id=$1 AND key=$2",
    [tenantId, REVIEW_KEY],
  );
  const parsed = auditReviewSchema.safeParse(parsedValue(result.rows[0]?.value));
  return parsed.success ? parsed.data : null;
}

export async function writeLastReview(
  sql: SqlClient,
  tenantId: string,
  review: AuditReview,
  actorId: string,
): Promise<void> {
  await sql.query(
    `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
     VALUES ($1,$2,$3::jsonb,now(),$4)
     ON CONFLICT (tenant_id, key) DO UPDATE SET value=EXCLUDED.value, updated_at=EXCLUDED.updated_at,
       updated_by=EXCLUDED.updated_by`,
    [tenantId, REVIEW_KEY, JSON.stringify(review), actorId],
  );
}

// The scheduled review ----------------------------------------------------------------------------

// Stored as the schedule plus who saved it, their name, when, and the permissions their sign-in
// had granted among each audit's packs then. A scheduled review reads each audit with no more
// than those, whoever's sign-in it uses, the way a scheduled single audit does.
export interface StoredReviewSchedule extends AuditReviewSchedule {
  scheduledBy: string | null;
  name: string | null;
  savedAt: string | null;
  grants: Record<string, string[]>;
}

export const OFF_REVIEW_SCHEDULE: StoredReviewSchedule = {
  frequency: "off",
  weekday: 0,
  hour: 6,
  scheduledBy: null,
  name: null,
  savedAt: null,
  grants: {},
};

const permissionSchema = z.string().regex(/^[A-Za-z][A-Za-z0-9]*(?:\.[A-Za-z][A-Za-z0-9]*){0,4}$/);

// Read leniently: a damaged value is the schedule being off, so it can stop a review from
// running and never start one.
export function storedReviewSchedule(value: unknown): StoredReviewSchedule {
  const base = auditReviewScheduleSchema.safeParse(value);
  if (!base.success) return OFF_REVIEW_SCHEDULE;
  const raw = value as Record<string, unknown>;
  const by = z.uuid().safeParse(raw.scheduledBy);
  const savedAt = z.string().datetime({ offset: true }).safeParse(raw.savedAt);
  const grants: Record<string, string[]> = {};
  const savedGrants =
    raw.grants !== null && typeof raw.grants === "object" && !Array.isArray(raw.grants)
      ? (raw.grants as Record<string, unknown>)
      : {};
  for (const id of Object.keys(savedGrants).slice(0, 100)) {
    const list = z.array(permissionSchema).max(20).safeParse(savedGrants[id]);
    if (list.success) grants[id] = [...new Set(list.data)];
  }
  return {
    ...base.data,
    scheduledBy: by.success ? by.data : null,
    name: typeof raw.name === "string" ? cleanText(raw.name).trim().slice(0, 200) || null : null,
    savedAt: savedAt.success ? savedAt.data : null,
    grants,
  };
}

export async function readReviewSchedule(
  sql: SqlClient,
  tenantId: string,
): Promise<StoredReviewSchedule> {
  const result = await sql.query<{ value: unknown }>(
    "SELECT value FROM prism365_preference WHERE tenant_id=$1 AND key=$2",
    [tenantId, SCHEDULE_KEY],
  );
  return storedReviewSchedule(parsedValue(result.rows[0]?.value));
}

export async function writeReviewSchedule(
  sql: SqlClient,
  tenantId: string,
  schedule: AuditReviewSchedule,
  actorId: string,
  saver: { name: string | null; grants: Record<string, string[]>; now: Date },
): Promise<StoredReviewSchedule> {
  const request = auditReviewScheduleSchema.parse(schedule);
  const value: StoredReviewSchedule =
    request.frequency === "off"
      ? { ...OFF_REVIEW_SCHEDULE, weekday: request.weekday, hour: request.hour }
      : {
          ...request,
          scheduledBy: z.uuid().parse(actorId),
          name: saver.name ? cleanText(saver.name).trim().slice(0, 200) || null : null,
          savedAt: saver.now.toISOString(),
          grants: saver.grants,
        };
  await sql.query(
    `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
     VALUES ($1,$2,$3::jsonb,now(),$4)
     ON CONFLICT (tenant_id, key) DO UPDATE SET value=EXCLUDED.value, updated_at=EXCLUDED.updated_at,
       updated_by=EXCLUDED.updated_by`,
    [tenantId, SCHEDULE_KEY, JSON.stringify(value), actorId],
  );
  return storedReviewSchedule(value);
}

// Whether a scheduled review is due: the latest slot (the chosen weekday and hour in the time
// zone) at or before now came after the newest review started, and after the schedule was saved,
// so saving a schedule on a Wednesday does not run last Monday's review at once. A review an
// administrator started after the slot counts as that week's.
export function reviewDue(
  schedule: StoredReviewSchedule,
  lastStart: Date | null,
  now: Date,
  timeZone: string,
): boolean {
  if (schedule.frequency !== "weekly") return false;
  const slot = weeklySlot(now, timeZone, schedule.weekday, schedule.hour).getTime();
  if (lastStart !== null && lastStart.getTime() >= slot) return false;
  if (schedule.savedAt !== null && Date.parse(schedule.savedAt) >= slot) return false;
  return true;
}

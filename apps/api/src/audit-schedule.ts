import {
  type Audit,
  type AuditScheduleMap,
  auditDefinition,
  auditDue,
  auditReadScope,
  latestRunStart,
  readAuditSchedule,
  scheduledAuditCount,
} from "@prism365/audits";
import type { AuditAvailability } from "@prism365/contracts";
import type { GraphHttpClient } from "@prism365/graph";
import type { ModuleDatabase } from "@prism365/storage";
import type { AuditRunner } from "./audits.ts";

// Scheduled audits (docs/audits.md, section 11). An administrator can set an audit to run daily
// or weekly; this runs the audits that are due, one at a time, on the sign-in that background
// collection uses, and only while collection is on. A scheduled run is an ordinary run with
// trigger "schedule": same runner, same access-ledger entry before anything is read, same
// history, same failures.
//
// Nothing here retries. An audit is due a day (a week) after its newest run of any outcome
// started, so a failed run waits for its next turn and a manual run counts as this period's.

export const SCHEDULE_TICK_MILLISECONDS = 5 * 60_000;
// After collection starts, due audits run this soon instead of waiting up to a full tick.
export const SCHEDULE_START_DELAY_MILLISECONDS = 30_000;

// What the scheduler uses of the collecting credential (DelegatedSessionCredential): who it is,
// what it may read, and its tokens. It never sees the session handle.
export interface ScheduleCredential {
  readonly tenantId: string;
  readonly actorId: string;
  readonly displayName: string | null;
  grantedPermissions(): Promise<readonly string[]>;
  accessToken(permissions: readonly string[]): Promise<string>;
}

// Timers, injectable so tests can drive them. Each returns its own cancel.
export interface ScheduleTimers {
  every(milliseconds: number, run: () => void): () => void;
  after(milliseconds: number, run: () => void): () => void;
}

// Unreferenced: a pending tick never keeps the process alive.
export const systemTimers: ScheduleTimers = {
  every(milliseconds, run) {
    const timer = setInterval(run, milliseconds);
    timer.unref?.();
    return () => clearInterval(timer);
  },
  after(milliseconds, run) {
    const timer = setTimeout(run, milliseconds);
    timer.unref?.();
    return () => clearTimeout(timer);
  },
};

export interface AuditSchedulerOptions {
  runner: Pick<AuditRunner, "start" | "runningId">;
  database: ModuleDatabase;
  audits: readonly Audit[];
  // The credential collection runs on, or null while collection is off.
  credential: () => ScheduleCredential | null;
  // The read packs this installation enabled (ENABLED_READ_FEATURES).
  enabledPacks: readonly string[];
  // The rule the manual Run route uses, given the permissions the credential has granted.
  availability: (audit: Audit, granted: readonly string[]) => AuditAvailability;
  // A read-only Graph client with the audits' budgets.
  graphClient: (getAccessToken: () => Promise<string>) => GraphHttpClient;
  // The access ledger.
  ledger: (tenantId: string, actorId: string, event: Record<string, unknown>) => Promise<void>;
  now?: () => Date;
  intervalMilliseconds?: number;
  startDelayMilliseconds?: number;
  timers?: ScheduleTimers;
  // Categories and audit ids only: never a tenant value, an account or a provider message.
  log?: (message: string) => void;
}

const errorName = (error: unknown) => (error instanceof Error ? error.name : typeof error);

export class AuditScheduler {
  readonly #runner: AuditSchedulerOptions["runner"];
  readonly #database: ModuleDatabase;
  readonly #audits: readonly Audit[];
  readonly #credential: () => ScheduleCredential | null;
  readonly #enabledPacks: readonly string[];
  readonly #availability: AuditSchedulerOptions["availability"];
  readonly #graphClient: AuditSchedulerOptions["graphClient"];
  readonly #ledger: AuditSchedulerOptions["ledger"];
  readonly #now: () => Date;
  readonly #interval: number;
  readonly #startDelay: number;
  readonly #timers: ScheduleTimers;
  readonly #log: (message: string) => void;
  #cancelInterval: (() => void) | null = null;
  #cancelSoon: (() => void) | null = null;
  #inFlight: Promise<void> | null = null;
  #stopped = false;

  constructor(options: AuditSchedulerOptions) {
    this.#runner = options.runner;
    this.#database = options.database;
    this.#audits = options.audits;
    this.#credential = options.credential;
    this.#enabledPacks = options.enabledPacks;
    this.#availability = options.availability;
    this.#graphClient = options.graphClient;
    this.#ledger = options.ledger;
    this.#now = options.now ?? (() => new Date());
    this.#interval = options.intervalMilliseconds ?? SCHEDULE_TICK_MILLISECONDS;
    this.#startDelay = options.startDelayMilliseconds ?? SCHEDULE_START_DELAY_MILLISECONDS;
    this.#timers = options.timers ?? systemTimers;
    this.#log = options.log ?? ((message) => console.log(message));
  }

  start(): void {
    if (this.#stopped || this.#cancelInterval) return;
    this.#cancelInterval = this.#timers.every(this.#interval, () => void this.tick());
  }

  // One tick shortly from now: collection has just started.
  soon(): void {
    if (this.#stopped || this.#cancelSoon) return;
    this.#cancelSoon = this.#timers.after(this.#startDelay, () => {
      this.#cancelSoon = null;
      void this.tick();
    });
  }

  // Clears both timers and starts nothing more. A run in progress is stopped by the runner's
  // own close; the tick waiting on it then returns, which settled() waits for.
  stop(): void {
    this.#stopped = true;
    this.#cancelInterval?.();
    this.#cancelInterval = null;
    this.#cancelSoon?.();
    this.#cancelSoon = null;
  }

  async settled(): Promise<void> {
    await this.#inFlight;
  }

  // One pass at a time: a tick that arrives while a pass runs joins it. Never rejects.
  tick(): Promise<void> {
    if (this.#stopped) return Promise.resolve();
    if (this.#inFlight) return this.#inFlight;
    this.#inFlight = this.#pass()
      .catch((error: unknown) => {
        this.#log(`Prism365 scheduled audits stopped for this round (${errorName(error)}).`);
      })
      .finally(() => {
        this.#inFlight = null;
      });
    return this.#inFlight;
  }

  async #pass(): Promise<void> {
    const first = this.#credential();
    if (!first) return;
    const schedule = await this.#database.tenant(first.tenantId, (sql) =>
      readAuditSchedule(sql, first.tenantId, this.#audits),
    );
    if (scheduledAuditCount(schedule) === 0) return;
    for (const audit of this.#audits) {
      if (schedule[audit.definition.id] === "off") continue;
      if (this.#stopped) return;
      // Asked again before every audit: collection can stop, or pass to another administrator,
      // while an earlier audit runs.
      const credential = this.#credential();
      if (!credential) return;
      if (!(await this.#runIfDue(audit, credential))) return;
    }
  }

  // Starts the audit when it is due and waits for it to be recorded. false when the pass should
  // end: the credential is gone or the runner has stopped.
  async #runIfDue(audit: Audit, credential: ScheduleCredential): Promise<boolean> {
    const id = audit.definition.id;
    const { tenantId, actorId } = credential;
    // The schedule and the last start are read again here, so a schedule turned off, or a manual
    // run started, while an earlier audit ran is respected.
    const { current, lastStart } = await this.#database.tenant(tenantId, async (sql) => ({
      current: (await readAuditSchedule(sql, tenantId, this.#audits)) as AuditScheduleMap,
      lastStart: await latestRunStart(sql, tenantId, id),
    }));
    if (!auditDue(current[id] ?? "off", lastStart, this.#now())) return true;
    if (this.#runner.runningId(tenantId, id)) return true;
    let scopes: readonly string[];
    try {
      scopes = await credential.grantedPermissions();
    } catch {
      this.#log("Prism365 scheduled audits paused: the collecting sign-in is not available.");
      return false;
    }
    // Skipped unless every pack the audit needs is enabled and granted, by the manual route's
    // rule; a skipped audit stays due and is tried again on the next tick.
    const availability = this.#availability(audit, scopes);
    if (availability.state !== "available") {
      this.#log(`Prism365 scheduled audit ${id} skipped (${availability.state}).`);
      return true;
    }
    // The run reads the audit's packs and only the optional packs this installation enabled,
    // with the permissions those packs carry, exactly as a manual run does.
    const read = auditReadScope(audit, this.#enabledPacks, scopes);
    const definition = auditDefinition(audit);
    // The ledger names what this run will read, as the manual route's entry does.
    const datasets = [
      ...definition.permissions,
      ...definition.optionalPermissions.filter((permission) => read.granted.includes(permission)),
    ];
    let started: Awaited<ReturnType<AuditSchedulerOptions["runner"]["start"]>>;
    try {
      started = await this.#runner.start({
        tenantId,
        actor: { id: actorId, name: credential.displayName },
        audit,
        client: this.#graphClient(() => credential.accessToken(read.granted)),
        packs: read.packs,
        granted: read.granted,
        trigger: "schedule",
        beforeRead: (run) =>
          this.#ledger(tenantId, actorId, {
            type: "audit.run.started",
            auditId: id,
            runId: run.id,
            datasets,
            trigger: "schedule",
            occurredAt: this.#now().toISOString(),
          }),
      });
    } catch (error) {
      // A run whose ledger entry failed is already stored as failed, so it waits for its next
      // turn; one that could not be stored at all is tried again on the next tick.
      this.#log(`Prism365 scheduled audit ${id} did not start (${errorName(error)}).`);
      return !this.#stopped;
    }
    if (!started.started) return true;
    this.#log(`Prism365 scheduled audit ${id} started.`);
    await started.done;
    return !this.#stopped;
  }
}

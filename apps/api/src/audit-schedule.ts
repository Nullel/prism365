import {
  type Audit,
  auditDefinition,
  auditDue,
  auditReadScope,
  knownAuditSchedule,
  latestRunStart,
  readAuditSchedule,
  readStoredAuditSchedule,
  type StoredAuditSchedule,
  scheduledAuditCount,
} from "@prism365/audits";
import type { AuditAvailability } from "@prism365/contracts";
import { type GraphFeatureId, type GraphHttpClient, permissionPacks } from "@prism365/graph";
import { CredentialUnavailableError } from "@prism365/modules";
import type { ModuleDatabase } from "@prism365/storage";
import type { AuditRunner } from "./audits.ts";

// Scheduled audits (docs/audits.md, section 15). An administrator can set an audit to run daily
// or weekly; this runs the audits that are due, one at a time, on the sign-in that background
// collection uses, and only while collection is on. A scheduled run is an ordinary run with
// trigger "schedule": same runner, same access-ledger entry before anything is read, same
// history, same failures.
//
// A scheduled run reads only while the credential it started with is still the one collecting:
// every token request checks, and a run in progress is stopped the moment collection stops or
// passes to another administrator (collectionEnded), so Stop stops scheduled reads too.
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
  // Set once Microsoft refused this sign-in a token: only signing in again cures it.
  readonly refused?: boolean;
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

const packPermission = (pack: string) =>
  permissionPacks.find((candidate) => candidate.id === pack)?.permission ?? null;

// What a scheduled run of this audit may read. The permissions allowed are the ones the
// administrator who scheduled it had granted among its packs when they saved it, or, for an entry
// saved before those were kept, the audit's required permissions only. `scopes` is the
// collecting sign-in's grants narrowed to those; `granted` and `packs` are what auditReadScope
// allows with them, without any pack whose permission did not survive. A scheduled run therefore
// never reads more than the administrator who scheduled it could have read themselves.
export function scheduledReadScope(
  audit: Audit,
  enabledPacks: readonly string[],
  collecting: readonly string[],
  stored: StoredAuditSchedule,
): { scopes: string[]; packs: GraphFeatureId[]; granted: string[] } {
  const id = audit.definition.id;
  const allowed = new Set(
    Object.hasOwn(stored.grants, id)
      ? (stored.grants[id] ?? [])
      : auditDefinition(audit).permissions,
  );
  const scopes = collecting.filter((scope) => allowed.has(scope));
  const read = auditReadScope(audit, enabledPacks, scopes);
  return {
    scopes,
    packs: read.packs.filter((pack) => read.granted.includes(packPermission(pack) ?? "")),
    granted: read.granted,
  };
}

// The scheduled audits that read less than a fully granted manual run would here: the
// permissions their enabled packs carry that the administrator who scheduled them had not
// granted. Shown on the audit's page.
export function scheduleLimits(
  stored: StoredAuditSchedule,
  audits: readonly Audit[],
  enabledPacks: readonly string[],
): Record<string, { by: string | null; without: string[] }> {
  const known = knownAuditSchedule(stored, audits);
  const limits: Record<string, { by: string | null; without: string[] }> = {};
  for (const audit of audits) {
    const id = audit.definition.id;
    if (known[id] === "off") continue;
    const { permissions, optionalPermissions } = auditDefinition(audit);
    const everything = auditReadScope(audit, enabledPacks, [
      ...permissions,
      ...optionalPermissions,
    ]).granted;
    const allowed = scheduledReadScope(audit, enabledPacks, everything, stored).granted;
    const without = everything.filter((permission) => !allowed.includes(permission));
    if (without.length > 0) limits[id] = { by: stored.names[id] ?? null, without };
  }
  return limits;
}

export interface AuditSchedulerOptions {
  runner: Pick<AuditRunner, "start" | "runningId" | "stopRun">;
  database: ModuleDatabase;
  audits: readonly Audit[];
  // The credential collection runs on, or null while collection is off. Asked on every token
  // request, so it must return the live value, never a copy.
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
  // Scheduled runs this scheduler started that are still in progress.
  readonly #active = new Set<string>();
  // The reason each skipped audit was last logged with, so an audit that stays unavailable is
  // logged once, not on every tick.
  readonly #skipped = new Map<string, string>();
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

  // Collection stopped, or passed to another administrator: every scheduled run in progress is
  // stopped now, its request in flight included. Manual runs are not touched.
  collectionEnded(): void {
    for (const runId of this.#active) this.#runner.stopRun(runId);
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

  // Whether this credential is still the one collecting, and the scheduler still running.
  #live(credential: ScheduleCredential): boolean {
    return !this.#stopped && this.#credential() === credential;
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
  // end: the credential is gone, stopped collecting or refused, or the scheduler has stopped.
  async #runIfDue(audit: Audit, credential: ScheduleCredential): Promise<boolean> {
    const id = audit.definition.id;
    const { tenantId, actorId } = credential;
    if (credential.refused) return this.#paused();
    // The schedule and the last start are read again here, so a schedule turned off, or a manual
    // run started, while an earlier audit ran is respected.
    const { stored, lastStart } = await this.#database.tenant(tenantId, async (sql) => ({
      stored: await readStoredAuditSchedule(sql, tenantId),
      lastStart: await latestRunStart(sql, tenantId, id),
    }));
    const frequency = knownAuditSchedule(stored, this.#audits)[id] ?? "off";
    if (!auditDue(frequency, lastStart, this.#now())) return true;
    if (this.#runner.runningId(tenantId, id)) return true;
    let scopes: readonly string[];
    try {
      scopes = await credential.grantedPermissions();
    } catch {
      return this.#paused();
    }
    // The run may read only what both sign-ins allow: the collecting one's grants, narrowed to
    // the permissions the administrator who scheduled it had granted then.
    const read = scheduledReadScope(audit, this.#enabledPacks, scopes, stored);
    // Skipped unless every pack the audit needs is enabled and granted, by the manual route's
    // rule; a skipped audit stays due and is tried again on the next tick.
    const availability = this.#availability(audit, read.scopes);
    if (availability.state !== "available") {
      const reason = [availability.state, ...availability.enable, ...availability.missing].join(
        " ",
      );
      if (this.#skipped.get(id) !== reason)
        this.#log(`Prism365 scheduled audit ${id} skipped (${availability.state}).`);
      this.#skipped.set(id, reason);
      return true;
    }
    this.#skipped.delete(id);
    const definition = auditDefinition(audit);
    // The ledger names what this run will read, as the manual route's entry does.
    const datasets = [
      ...definition.permissions,
      ...definition.optionalPermissions.filter((permission) => read.granted.includes(permission)),
    ];
    // Every token request checks the credential is still the one collecting. A refused or
    // ended sign-in is noted, so the pass stops instead of spending every audit's turn on it.
    let unavailable = false;
    const getAccessToken = async () => {
      if (!this.#live(credential)) {
        unavailable = true;
        throw new CredentialUnavailableError("Collection stopped.");
      }
      try {
        return await credential.accessToken(read.granted);
      } catch (error) {
        if (error instanceof CredentialUnavailableError) unavailable = true;
        throw error;
      }
    };
    // Checked once more, after the awaits above: collection may have stopped meanwhile.
    if (!this.#live(credential)) return false;
    let started: Awaited<ReturnType<AuditSchedulerOptions["runner"]["start"]>>;
    try {
      started = await this.#runner.start({
        tenantId,
        actor: { id: actorId, name: credential.displayName },
        audit,
        client: this.#graphClient(getAccessToken),
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
            // Who set this audit to run on its own; the entry's actor is whose sign-in it uses.
            scheduledBy: stored.scheduledBy[id] ?? null,
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
    this.#active.add(started.runId);
    // Collection may have ended between the start and now; the run stops before it reads.
    if (!this.#live(credential)) this.#runner.stopRun(started.runId);
    this.#log(`Prism365 scheduled audit ${id} started.`);
    try {
      await started.done;
    } finally {
      this.#active.delete(started.runId);
    }
    if (!this.#live(credential)) return false;
    if (unavailable || credential.refused) return this.#paused();
    return true;
  }

  #paused(): false {
    this.#log("Prism365 scheduled audits paused: the collecting sign-in is not available.");
    return false;
  }
}

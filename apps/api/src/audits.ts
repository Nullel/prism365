import { randomUUID } from "node:crypto";
import {
  type Audit,
  AuditError,
  cleanText,
  finishRun,
  insertRun,
  interruptRuns,
  pruneRuns,
  type RunOutcome,
  readAuditScope,
} from "@prism365/audits";
import {
  type AuditAvailability,
  type AuditRun,
  type AuditRunTrigger,
  auditRowSchema,
  auditSummaryItemSchema,
} from "@prism365/contracts";
import {
  GraphBoundaryError,
  type GraphFeatureId,
  type GraphHttpClient,
  GraphReportDownloadError,
  GraphReportHostError,
  GraphRequestError,
  GraphRetryDeferredError,
} from "@prism365/graph";
import { CredentialUnavailableError, scopeGraph } from "@prism365/modules";
import type { ModuleDatabase } from "@prism365/storage";
import { z } from "zod";

// Runs audits in the API process, one at a time per audit and tenant, with the delegated
// session of the administrator who started them. A run is stored as "running" before it reads
// anything and closed exactly once, by whichever comes first: its result, its failure, its time
// limit or shutdown. A run left "running" by a stopped process is closed as interrupted the
// next time this tenant's audits are read.

export const AUDIT_TIMEOUT_MILLISECONDS = 10 * 60_000;

const resultSchema = z.object({
  rows: z.array(auditRowSchema).max(20_000),
  summary: z.array(auditSummaryItemSchema).max(20),
  notes: z.array(z.string().transform((value) => value.slice(0, 500))).max(20),
  complete: z.boolean(),
  domains: z
    .array(z.object({ domain: z.string().max(253), accounts: z.number().int().nonnegative() }))
    .max(1_000)
    .optional(),
});

// Every string in a result, cleaned the way the audits clean provider text, so a value an audit
// forgot to clean cannot make the result impossible to store.
function cleaned(value: unknown): unknown {
  if (typeof value === "string") return cleanText(value);
  if (Array.isArray(value)) return value.map(cleaned);
  if (value && typeof value === "object")
    return Object.fromEntries(
      Object.entries(value).map(([key, item]) => [cleanText(key), cleaned(item)]),
    );
  return value;
}

// Why a run was stopped before it finished: its time limit, the process stopping, or (for a
// scheduled run only) collection stopping or passing to another administrator.
type StopReason = "timeout" | "shutdown" | "collection";

class AuditStoppedError extends Error {
  readonly reason: StopReason;
  constructor(reason: StopReason) {
    super(
      reason === "timeout"
        ? "The audit exceeded its time limit."
        : reason === "collection"
          ? "Collection stopped."
          : "Prism365 is stopping.",
    );
    this.name = "AuditStoppedError";
    this.reason = reason;
  }
}

// What the Run route says, and the schedule route says, when an audit cannot run on a session.
export function auditUnavailableMessage(availability: AuditAvailability): string {
  return availability.state === "not_enabled"
    ? `This installation has not enabled ${availability.enable.join(", ") || availability.missing.join(", ")}. Add it to ENABLED_READ_FEATURES and restart.`
    : `Your session has not granted ${availability.missing.join(", ")}. Grant access and run it again.`;
}

// What the administrator is told. Fixed text per category: provider messages and tenant values
// never reach the page or the log.
export function describeAuditFailure(
  error: unknown,
  audit: Audit,
  timeoutMilliseconds = AUDIT_TIMEOUT_MILLISECONDS,
): { code: string; message: string } {
  if (error instanceof AuditError) return { code: error.code, message: error.message };
  if (error instanceof AuditStoppedError)
    return error.reason === "timeout"
      ? {
          code: "timeout",
          message: `The audit took longer than ${Math.round(timeoutMilliseconds / 60_000)} minutes and was stopped.`,
        }
      : error.reason === "collection"
        ? {
            code: "collection_stopped",
            message: "Collection stopped while this scheduled run was reading.",
          }
        : {
            code: "interrupted",
            message: "Prism365 stopped while this audit was running. Run it again.",
          };
  const reference = (requestId: string | undefined) =>
    requestId ? ` Microsoft request ID: ${requestId}.` : "";
  // A usage report's download link (ADR-0010). The host is printable ASCII of at most 253
  // characters, never the link's path or query, which carry its token.
  if (error instanceof GraphReportHostError)
    return {
      code: "report_host",
      message:
        error.host === null
          ? "Microsoft sent a report download link Prism365 could not read."
          : error.knownHost
            ? `Microsoft sent a report download link on ${error.host} that Prism365 does not accept: ${error.problem}.`
            : `Microsoft sent the report from a host Prism365 does not recognise: ${error.host}.`,
    };
  // The short-lived link failed. Not a sign-in or role problem, whatever the status says.
  if (
    error instanceof GraphReportDownloadError ||
    (error instanceof GraphBoundaryError && error.reason === "download")
  )
    return {
      code: "report_download",
      message: "Microsoft's report download link failed. Run the audit again.",
    };
  if (error instanceof GraphRetryDeferredError)
    return {
      code: "throttled",
      message: `Microsoft is limiting requests right now. Try again in a few minutes.${reference(error.requestId)}`,
    };
  if (error instanceof GraphRequestError) {
    if (error.status === 401)
      return {
        code: "session",
        message:
          "Microsoft did not accept the session. Sign out, sign in again, and rerun the audit.",
      };
    if (error.status === 403)
      return { code: "forbidden", message: `${audit.forbidden}${reference(error.requestId)}` };
    if (error.status === 400)
      return {
        code: "rejected",
        message: `Microsoft rejected the request (HTTP 400).${reference(error.requestId)}`,
      };
    return {
      code: "microsoft_error",
      message: `Microsoft returned HTTP ${error.status}. Try again shortly.${reference(error.requestId)}`,
    };
  }
  if (error instanceof GraphBoundaryError) {
    if (error.reason === "timeout")
      return { code: "slow", message: "Microsoft took too long to answer. Try again." };
    if (error.reason === "credential")
      return {
        code: "session",
        message: "The Microsoft session ended. Sign in again and rerun the audit.",
      };
    return {
      code: "graph",
      message: `Microsoft's answer could not be read safely (${error.reason}). Try again; if it repeats, the server log has the category.`,
    };
  }
  // A scheduled run's sign-in ended (signed out, expired, or refused by Microsoft): the same
  // category a manual run gets when its session ends.
  if (error instanceof CredentialUnavailableError)
    return {
      code: "session",
      message: "The Microsoft session ended. Sign in again and rerun the audit.",
    };
  const name = error instanceof Error ? error.name : "";
  if (name === "InteractionRequiredAuthError")
    return {
      code: "session",
      message: "Microsoft needs you to sign in again. Sign out, sign in, and rerun the audit.",
    };
  if (error instanceof z.ZodError)
    return {
      code: "invalid_result",
      message: "The audit produced a result Prism365 could not store. This is a bug in the audit.",
    };
  return { code: "failed", message: "The audit failed. The server log has the category." };
}

// The server log line for a failed run, or null when the page already says all there is. It holds
// fixed codes and error class names, and for a refused report link its host, so an operator can
// report a host Microsoft moved to; never a provider message or the link itself.
export function auditFailureLogLine(auditId: string, code: string, error: unknown): string | null {
  if (error instanceof GraphReportHostError)
    return `Prism365 audit ${auditId} failed (report_host, ${error.host === null ? "no host" : `host ${error.host}`}).`;
  if (code === "failed" || code === "invalid_result")
    return `Prism365 audit ${auditId} failed (${error instanceof Error ? error.name : typeof error}).`;
  return null;
}

interface Running {
  runId: string;
  controller: AbortController;
  done: Promise<void>;
}

// A run whose result is stored, and with it the audit's new trend point.
export interface AuditStoredEvent {
  tenantId: string;
  auditId: string;
  runId: string;
  trigger: AuditRunTrigger;
}

export interface AuditRunnerOptions {
  database: ModuleDatabase;
  timeoutMilliseconds?: number;
  now?: () => Date;
}

export class AuditRunner {
  readonly #database: ModuleDatabase;
  readonly #timeout: number;
  readonly #now: () => Date;
  readonly #running = new Map<string, Running>();
  readonly #recovered = new Map<string, Promise<void>>();
  readonly #listeners = new Set<(event: AuditStoredEvent) => void>();
  #closed = false;

  constructor(options: AuditRunnerOptions) {
    this.#database = options.database;
    this.#timeout = options.timeoutMilliseconds ?? AUDIT_TIMEOUT_MILLISECONDS;
    this.#now = options.now ?? (() => new Date());
  }

  // Called after each successful run is stored. Listeners must not throw or block: a failing one
  // is ignored, and nothing waits for what it starts.
  subscribe(listener: (event: AuditStoredEvent) => void): () => void {
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  }

  runningId(tenantId: string, auditId: string): string | null {
    return this.#running.get(`${tenantId}:${auditId}`)?.runId ?? null;
  }

  // Closes runs a stopped process left open, once per tenant per process. Every insert waits
  // for this, so it can never close a run this process has just started.
  recover(tenantId: string): Promise<void> {
    let recovery = this.#recovered.get(tenantId);
    if (!recovery) {
      const live = [...this.#running.values()].map((entry) => entry.runId);
      recovery = this.#database
        .tenant(tenantId, async (sql) => {
          await interruptRuns(sql, tenantId, live, this.#now());
        })
        .catch((error: unknown) => {
          this.#recovered.delete(tenantId);
          throw error;
        });
      this.#recovered.set(tenantId, recovery);
    }
    return recovery;
  }

  // Starts a run, or returns the one already running for this audit. `done` resolves once the
  // run (this one or the one joined) has been recorded, however it ended.
  async start(input: {
    tenantId: string;
    actor: { id: string; name: string | null };
    audit: Audit;
    client: GraphHttpClient;
    // The packs this run may read and the permissions they carry, from auditReadScope: the
    // audit's own packs plus only the optional packs the operator enabled.
    packs: readonly GraphFeatureId[];
    granted: readonly string[];
    // An administrator's request unless the audit's schedule started it.
    trigger?: AuditRunTrigger;
    // Runs once, only for a run this call started, after it is stored and before anything is
    // read: the access-ledger entry. If it fails, the run is closed as failed and reads nothing.
    beforeRead?: (run: AuditRun) => Promise<void>;
  }): Promise<{ runId: string; started: boolean; run: AuditRun | null; done: Promise<void> }> {
    if (this.#closed) throw new AuditStoppedError("shutdown");
    const key = `${input.tenantId}:${input.audit.definition.id}`;
    const existing = this.#running.get(key);
    if (existing) return { runId: existing.runId, started: false, run: null, done: existing.done };
    const runId = randomUUID();
    const controller = new AbortController();
    let release: () => void = () => {};
    const entry: Running = {
      runId,
      controller,
      done: new Promise<void>((resolve) => {
        release = resolve;
      }),
    };
    // Claimed before the first await, so a second request for the same audit joins this run.
    this.#running.set(key, entry);
    let run: AuditRun;
    // The account domains the administrator limited the audits to, read once here: a scope saved
    // while the run is reading Microsoft applies from the next run on.
    let scope: ReadonlySet<string> | null = null;
    try {
      await this.recover(input.tenantId);
      // One transaction for the scope and the new run, so the run is stored as soon as it was
      // before the scope existed, and a second request that joins it can already read it.
      run = await this.#database.tenant(input.tenantId, async (sql) => {
        const saved = input.audit.scoped ? await readAuditScope(sql, input.tenantId) : null;
        scope = saved && saved.domains.length > 0 ? new Set(saved.domains) : null;
        return insertRun(sql, input.tenantId, {
          id: runId,
          auditId: input.audit.definition.id,
          auditVersion: input.audit.definition.version,
          startedAt: this.#now(),
          startedBy: input.actor,
          scope: saved?.domains ?? null,
          trigger: input.trigger ?? "manual",
        });
      });
    } catch (error) {
      this.#running.delete(key);
      release();
      throw error;
    }
    try {
      await input.beforeRead?.(run);
    } catch (error) {
      await this.#record(input.tenantId, runId, input.audit, {
        status: "failed",
        code: "ledger",
        message: "The access ledger could not be written, so the audit read nothing. Try again.",
      });
      this.#running.delete(key);
      release();
      throw error;
    }
    void this.#execute({ ...input, scope }, entry)
      .catch(() => false)
      .then((stored) => {
        // Out of the running set first: a recovery that starts after this sees the run as dead.
        this.#running.delete(key);
        if (!stored) this.#recovered.delete(input.tenantId);
        release();
      });
    return { runId, started: true, run, done: entry.done };
  }

  // Records how a run ended and prunes old runs. false when storage refused.
  async #record(tenantId: string, runId: string, audit: Audit, outcome: RunOutcome) {
    try {
      await this.#database.tenant(
        tenantId,
        async (sql) => {
          await finishRun(sql, tenantId, runId, this.#now(), outcome);
          await pruneRuns(sql, tenantId, audit.definition.id);
        },
        { timeoutMilliseconds: 120_000 },
      );
      return true;
    } catch {
      return false;
    }
  }

  async #execute(
    input: {
      tenantId: string;
      audit: Audit;
      client: GraphHttpClient;
      packs: readonly GraphFeatureId[];
      granted: readonly string[];
      scope: ReadonlySet<string> | null;
      trigger?: AuditRunTrigger;
    },
    entry: Running,
  ): Promise<boolean> {
    const { audit, tenantId } = input;
    const { controller } = entry;
    let timer: ReturnType<typeof setTimeout> | undefined;
    let outcome: RunOutcome;
    try {
      const stopped = new Promise<never>((_resolve, reject) => {
        const stop = () =>
          reject(
            controller.signal.reason instanceof AuditStoppedError
              ? controller.signal.reason
              : new AuditStoppedError("shutdown"),
          );
        if (controller.signal.aborted) stop();
        controller.signal.addEventListener("abort", stop, { once: true });
      });
      stopped.catch(() => {});
      timer = setTimeout(() => controller.abort(new AuditStoppedError("timeout")), this.#timeout);
      // The signal also stops paging and the request in flight, so a stopped run stops reading.
      // Only the packs this run was given, which always include the audit's required ones.
      const allowed = new Set<string>([...audit.packs, ...audit.optionalPacks]);
      const packs = input.packs.filter((pack) => allowed.has(pack));
      if (audit.packs.some((pack) => !packs.includes(pack)))
        throw new AuditError("not_enabled", "This audit's read packs are not all enabled.");
      const graph = scopeGraph(
        input.client,
        packs,
        () => controller.signal.throwIfAborted(),
        controller.signal,
      );
      const work = audit.run({
        graph,
        now: this.#now(),
        granted: new Set(input.granted),
        signal: controller.signal,
        scope: input.scope,
        tenantId,
      });
      work.catch(() => {});
      const result = resultSchema.parse(cleaned(await Promise.race([work, stopped])));
      const { domains, ...rest } = result;
      outcome = { status: "succeeded", ...rest, ...(domains ? { domains } : {}) };
    } catch (error) {
      const cause =
        controller.signal.aborted && controller.signal.reason instanceof AuditStoppedError
          ? controller.signal.reason
          : error;
      const failure = describeAuditFailure(cause, audit, this.#timeout);
      outcome = { status: "failed", ...failure };
      const line = auditFailureLogLine(audit.definition.id, failure.code, cause);
      if (line) console.warn(line);
    } finally {
      if (timer !== undefined) clearTimeout(timer);
    }
    if (await this.#record(tenantId, entry.runId, audit, outcome)) {
      if (outcome.status === "succeeded")
        for (const listener of this.#listeners)
          try {
            listener({
              tenantId,
              auditId: audit.definition.id,
              runId: entry.runId,
              trigger: input.trigger ?? "manual",
            });
          } catch {
            // Observers, such as alert rules, must not be able to fail a run.
          }
      return true;
    }
    console.warn(`Prism365 audit ${audit.definition.id} result could not be stored.`);
    // A result storage refused still closes the run, as its own failure rather than an
    // interruption. If even that fails, the next recovery closes it.
    return (
      outcome.status === "succeeded" &&
      this.#record(tenantId, entry.runId, audit, {
        status: "failed",
        code: "store_failed",
        message:
          "The result could not be saved. Run the audit again; if it repeats, the server log has the category.",
      })
    );
  }

  // Stops one scheduled run because collection stopped or passed to another administrator: the
  // cancel reaches the request in flight and the run closes as failed. Only the audit scheduler
  // calls this, with ids of runs it started itself, so a manual run is never stopped this way.
  // false when the run is no longer in progress.
  stopRun(runId: string): boolean {
    for (const entry of this.#running.values())
      if (entry.runId === runId) {
        entry.controller.abort(new AuditStoppedError("collection"));
        return true;
      }
    return false;
  }

  // Stops every run and waits for each to record that it was interrupted.
  async close(): Promise<void> {
    this.#closed = true;
    const entries = [...this.#running.values()];
    for (const entry of entries) entry.controller.abort(new AuditStoppedError("shutdown"));
    await Promise.all(entries.map((entry) => entry.done));
  }

  // Resolves when every run in progress has been recorded. For tests.
  async idle(): Promise<void> {
    await Promise.all([...this.#running.values()].map((entry) => entry.done));
  }
}

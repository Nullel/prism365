import { randomUUID } from "node:crypto";
import type { Severity } from "@prism365/contracts";
import {
  GraphBoundaryError,
  type GraphHttpClient,
  GraphRequestError,
  GraphRetryDeferredError,
} from "@prism365/graph";
import {
  type ModuleDatabase,
  type SqlClient,
  SqlGuardError,
  type TableAccess,
  TransactionAbortedError,
} from "@prism365/storage";
import {
  AttentionInputError,
  readAttentionPreferences,
  retireAttention,
  upsertAttention,
} from "./attention.ts";
import type {
  Collector,
  CollectorContext,
  CredentialKind,
  FindingChange,
  FindingInput,
  PrismModule,
} from "./contract.ts";
import { FindingPolicyError, severityAtLeast, upsertFinding } from "./findings.ts";
import { ModuleBoundaryError, scopeGraph } from "./graph-scope.ts";
import { type ModuleRegistry, modulePermissions, moduleTablePrefix } from "./registry.ts";

export class CredentialUnavailableError extends Error {
  constructor(message = "The collection credential is no longer available.") {
    super(message);
    this.name = "CredentialUnavailableError";
  }
}

class RunEndedError extends Error {
  constructor() {
    super("The collector run has ended; late work is discarded.");
    this.name = "RunEndedError";
  }
}

// A source of Microsoft tokens for background collection. The delegated implementation wraps an
// operator's signed-in session and lives only in process memory; the application one (later)
// wraps a certificate. The scheduler never sees a token cache, secret, or session handle.
export interface CollectionCredential {
  readonly kind: CredentialKind;
  readonly tenantId: string;
  readonly actorId: string | null;
  grantedPermissions(): Promise<readonly string[]>;
  accessToken(permissions: readonly string[]): Promise<string>;
}

export interface FindingNotification {
  tenantId: string;
  module: string;
  findingId: string;
  ruleId: string;
  severity: Severity;
  title: string;
  summary: string;
  subjectLabel: string | null;
  change: Exclude<FindingChange, "updated">;
  observedAt: string;
}

export interface FindingSink {
  readonly id: string;
  readonly minimumSeverity: Severity;
  deliver(notification: FindingNotification): Promise<void>;
}

export type StopReason =
  | "operator_stopped"
  | "signed_out"
  | "credential_unavailable"
  | "replaced"
  | "shutdown";

export type RunStatus = "succeeded" | "partial" | "failed";

export type ErrorCategory =
  | "credential_unavailable"
  | "access_denied"
  | "throttled"
  | "provider_unavailable"
  | "provider_error"
  | "boundary"
  | `boundary.${string}`
  | "invalid_finding"
  | "invalid_attention"
  | "sql_guard"
  | "storage_denied"
  | "storage_error"
  | "timeout"
  | "collector_error";

export type SchedulerEvent =
  | { type: "collection.bound"; tenantId: string; credentialKind: CredentialKind }
  | {
      type: "collection.unbound";
      tenantId: string;
      actorId: string | null;
      reason: StopReason;
    }
  | {
      type: "run.finished";
      tenantId: string;
      module: string;
      collector: string;
      status: RunStatus;
      category: ErrorCategory | null;
      fetched: number;
      inserted: number;
      findings: number;
      skipped: number;
    }
  | { type: "module.skipped"; tenantId: string; module: string; missing: string[] }
  | { type: "sink.failed"; sink: string }
  | { type: "maintenance.failed"; tenantId: string; task: "recover" | "retention" };

export interface CollectionStatus {
  state: "stopped" | "collecting";
  credentialKind: CredentialKind | null;
  boundAt: string | null;
  lastStopReason: StopReason | null;
  skipped: { module: string; missing: string[] }[];
}

export interface SchedulerOptions {
  registry: ModuleRegistry;
  database: ModuleDatabase;
  graphClient: (getAccessToken: () => Promise<string>) => GraphHttpClient;
  // Read packs the operator enabled (ENABLED_READ_FEATURES). A module runs only when every pack it
  // declares is enabled: several packs share one Microsoft permission (AuditLog.Read.All), so a
  // consented permission must not switch on a module whose own pack the operator left off.
  enabledPacks: readonly string[];
  sinks?: readonly FindingSink[];
  now?: () => Date;
  tickMilliseconds?: number;
  retentionIntervalMilliseconds?: number;
  sinkTimeoutMilliseconds?: number;
  runHistoryDays?: number;
  // How long stop() waits for a running collector before giving up on it. Default 15 seconds.
  stopTimeoutMilliseconds?: number;
  // Override how long a collector may run, for tests and unusually slow environments.
  timeoutMilliseconds?: (collector: Collector) => number;
  // How soon a collector that reported more to read runs again. Default 5 seconds.
  catchUpDelayMilliseconds?: number;
  onEvent?: (event: SchedulerEvent) => void;
}

const identifier = /^[a-z][a-z0-9_]{0,40}$/;

export function classifyCollectorError(error: unknown): ErrorCategory {
  if (error instanceof CredentialUnavailableError) return "credential_unavailable";
  if (error instanceof RunEndedError || error instanceof TransactionAbortedError) return "timeout";
  if (error instanceof GraphRetryDeferredError) return "throttled";
  if (error instanceof GraphRequestError) {
    if (error.status === 401 || error.status === 403) return "access_denied";
    if (error.status >= 500) return "provider_unavailable";
    return "provider_error";
  }
  // The reason is a fixed code chosen by the Graph client, never provider text, so it is safe to
  // store and log, and it is what an operator needs to tell a timeout from an oversized page.
  if (error instanceof GraphBoundaryError) return `boundary.${error.reason}`;
  if (error instanceof ModuleBoundaryError) return "boundary.module";
  if (error instanceof FindingPolicyError) return "invalid_finding";
  if (error instanceof AttentionInputError) return "invalid_attention";
  if (error instanceof SqlGuardError) return "sql_guard";
  const code =
    typeof error === "object" && error !== null && "code" in error ? String(error.code) : "";
  if (code === "42501") return "storage_denied";
  if (/^[0-9A-Z]{5}$/.test(code)) return "storage_error";
  return "collector_error";
}

function boundedCount(value: number): number {
  if (!Number.isInteger(value) || value < 0)
    throw new ModuleBoundaryError("Counts must be non-negative integers.");
  return value;
}

export class ModuleScheduler {
  readonly #registry: ModuleRegistry;
  readonly #database: ModuleDatabase;
  readonly #graphClient: SchedulerOptions["graphClient"];
  readonly #enabledPacks: ReadonlySet<string>;
  readonly #sinks: readonly FindingSink[];
  readonly #now: () => Date;
  readonly #tickMilliseconds: number;
  readonly #retentionInterval: number;
  readonly #sinkTimeout: number;
  readonly #runHistoryDays: number;
  readonly #stopTimeout: number;
  readonly #timeoutFor: (collector: Collector) => number;
  readonly #catchUpDelay: number;
  readonly #listeners = new Set<(event: SchedulerEvent) => void>();
  #credential: CollectionCredential | null = null;
  #boundAt: Date | null = null;
  #lastStopReason: StopReason | null = null;
  #timer: ReturnType<typeof setInterval> | null = null;
  #soon: ReturnType<typeof setTimeout> | null = null;
  #inFlight: Promise<void> | null = null;
  // Collectors that stopped at their own budget with more to read; due on the next cycle.
  readonly #catchingUp = new Set<string>();
  readonly #recovered = new Set<string>();
  readonly #lastRetention = new Map<string, number>();
  readonly #skipped = new Map<string, string[]>();

  constructor(options: SchedulerOptions) {
    this.#registry = options.registry;
    this.#database = options.database;
    this.#graphClient = options.graphClient;
    this.#enabledPacks = new Set(options.enabledPacks);
    this.#sinks = options.sinks ?? [];
    this.#now = options.now ?? (() => new Date());
    this.#tickMilliseconds = options.tickMilliseconds ?? 60_000;
    this.#retentionInterval = options.retentionIntervalMilliseconds ?? 24 * 60 * 60_000;
    this.#sinkTimeout = options.sinkTimeoutMilliseconds ?? 10_000;
    this.#runHistoryDays = options.runHistoryDays ?? 30;
    this.#stopTimeout = options.stopTimeoutMilliseconds ?? 15_000;
    this.#timeoutFor =
      options.timeoutMilliseconds ?? ((collector) => collector.timeoutSeconds * 1_000);
    this.#catchUpDelay = options.catchUpDelayMilliseconds ?? 5_000;
    if (options.onEvent) this.#listeners.add(options.onEvent);
  }

  bind(credential: CollectionCredential): void {
    if (this.#credential && this.#credential !== credential) this.unbind("replaced");
    this.#credential = credential;
    this.#boundAt = this.#now();
    this.#skipped.clear();
    this.#emit({
      type: "collection.bound",
      tenantId: credential.tenantId,
      credentialKind: credential.kind,
    });
  }

  unbind(reason: StopReason): void {
    const credential = this.#credential;
    if (!credential) return;
    this.#credential = null;
    this.#boundAt = null;
    this.#lastStopReason = reason;
    this.#catchingUp.clear();
    this.#emit({
      type: "collection.unbound",
      tenantId: credential.tenantId,
      actorId: credential.actorId,
      reason,
    });
  }

  isBound(predicate: (credential: CollectionCredential) => boolean): boolean {
    return this.#credential !== null && predicate(this.#credential);
  }

  status(): CollectionStatus {
    return {
      state: this.#credential ? "collecting" : "stopped",
      credentialKind: this.#credential?.kind ?? null,
      boundAt: this.#boundAt?.toISOString() ?? null,
      lastStopReason: this.#lastStopReason,
      skipped: [...this.#skipped].map(([module, missing]) => ({ module, missing })),
    };
  }

  start(): void {
    if (this.#timer) return;
    this.#timer = setInterval(() => {
      void this.tick().catch(() => {});
    }, this.#tickMilliseconds);
    this.#timer.unref?.();
  }

  async stop(): Promise<void> {
    if (this.#timer) clearInterval(this.#timer);
    this.#timer = null;
    if (this.#soon) clearTimeout(this.#soon);
    this.#soon = null;
    this.unbind("shutdown");
    // Unbinding already makes every further call by a running collector fail. Wait a bounded
    // time for it to notice; shutdown must not hang on a collector that never returns.
    const running = this.#inFlight;
    if (running) {
      let timer: ReturnType<typeof setTimeout> | undefined;
      await Promise.race([
        running.catch(() => {}),
        new Promise<void>((resolve) => {
          timer = setTimeout(resolve, this.#stopTimeout);
        }),
      ]);
      if (timer !== undefined) clearTimeout(timer);
    }
  }

  // One cycle at a time. A tick that arrives while a cycle runs joins it instead of overlapping.
  tick(): Promise<void> {
    if (this.#inFlight) return this.#inFlight;
    const credential = this.#credential;
    if (!credential) return Promise.resolve();
    this.#inFlight = this.#cycle(credential).finally(() => {
      this.#inFlight = null;
      this.#scheduleCatchUp();
    });
    return this.#inFlight;
  }

  // A collector with more to read runs again shortly, but only once the scheduler is started:
  // tests and callers that drive tick() themselves keep full control of timing.
  #scheduleCatchUp(): void {
    if (!this.#timer || this.#soon || this.#catchingUp.size === 0 || !this.#credential) return;
    this.#soon = setTimeout(() => {
      this.#soon = null;
      void this.tick().catch(() => {});
    }, this.#catchUpDelay);
    this.#soon.unref?.();
  }

  #key(tenantId: string, module: PrismModule, collector: Collector): string {
    return `${tenantId}\u0000${module.id}\u0000${collector.id}`;
  }

  subscribe(listener: (event: SchedulerEvent) => void): () => void {
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  }

  #emit(event: SchedulerEvent): void {
    for (const listener of this.#listeners) {
      try {
        listener(event);
      } catch {
        // Event observers must not be able to stop collection.
      }
    }
  }

  #current(credential: CollectionCredential): boolean {
    return this.#credential === credential;
  }

  async #cycle(credential: CollectionCredential): Promise<void> {
    const tenantId = credential.tenantId;
    if (!this.#recovered.has(tenantId)) {
      try {
        // Runs still marked running belong to a process that stopped mid-run.
        await this.#database.tenant(tenantId, (sql) =>
          sql.query(
            "UPDATE prism365_module_run SET status='abandoned', finished_at=$2 WHERE tenant_id=$1 AND status='running'",
            [tenantId, this.#now().toISOString()],
          ),
        );
        this.#recovered.add(tenantId);
      } catch {
        this.#emit({ type: "maintenance.failed", tenantId, task: "recover" });
        return;
      }
    }

    let granted: readonly string[];
    try {
      granted = await credential.grantedPermissions();
    } catch {
      if (this.#current(credential)) this.unbind("credential_unavailable");
      return;
    }

    for (const module of this.#registry.list()) {
      if (!this.#current(credential)) return;
      const needed = modulePermissions(module);
      const notEnabled = module.packs.filter((pack) => !this.#enabledPacks.has(pack));
      const missing =
        notEnabled.length > 0
          ? notEnabled
          : needed.filter((permission) => !granted.includes(permission));
      if (missing.length > 0) {
        const previous = this.#skipped.get(module.id)?.join(",");
        this.#skipped.set(module.id, missing);
        if (previous !== missing.join(","))
          this.#emit({ type: "module.skipped", tenantId, module: module.id, missing });
        continue;
      }
      this.#skipped.delete(module.id);
      for (const collector of module.collectors) {
        if (!this.#current(credential)) return;
        if (!(await this.#due(tenantId, module, collector))) continue;
        await this.#run(credential, module, collector, needed);
      }
    }
    if (this.#current(credential)) await this.#retention(tenantId);
  }

  async #due(tenantId: string, module: PrismModule, collector: Collector): Promise<boolean> {
    if (this.#catchingUp.delete(this.#key(tenantId, module, collector))) return true;
    const result = await this.#database.tenant(tenantId, (sql) =>
      sql.query<{ started_at: Date | string }>(
        `SELECT started_at FROM prism365_module_run
          WHERE tenant_id=$1 AND module=$2 AND collector=$3 ORDER BY started_at DESC LIMIT 1`,
        [tenantId, module.id, collector.id],
      ),
    );
    const last = result.rows[0];
    if (!last) return true;
    const started = new Date(last.started_at).getTime();
    return this.#now().getTime() - started >= collector.cadenceSeconds * 1_000;
  }

  async #run(
    credential: CollectionCredential,
    module: PrismModule,
    collector: Collector,
    permissions: readonly string[],
  ): Promise<void> {
    const tenantId = credential.tenantId;
    const runId = randomUUID();
    const startedAt = this.#now();
    await this.#database.tenant(tenantId, (sql) =>
      sql.query(
        `INSERT INTO prism365_module_run (id, tenant_id, module, collector, credential_kind, started_at, status)
         VALUES ($1,$2,$3,$4,$5,$6,'running')`,
        [runId, tenantId, module.id, collector.id, credential.kind, startedAt.toISOString()],
      ),
    );

    const ended = new AbortController();
    let timedOut = false;
    let inTransaction = false;
    const counters = { fetched: 0, inserted: 0, findings: 0, skipped: 0 };
    const notifications: FindingNotification[] = [];
    const prefix = moduleTablePrefix(module.id);

    // Every context method checks this first: once the run is over (finished, failed, timed
    // out) or collection was stopped, nothing the collector does can reach Graph or storage.
    const live = () => {
      if (ended.signal.aborted) throw new RunEndedError();
      if (!this.#current(credential)) throw new CredentialUnavailableError("Collection stopped.");
    };
    const outsideTransaction = () => {
      if (inTransaction)
        throw new ModuleBoundaryError("This call is not allowed inside a module transaction.");
    };
    // Module SQL sees only the module's own tables, plus read access to findings. Core calls on
    // its behalf (cursor, findings) use the core's access. Either transaction is rolled back the
    // moment the run ends, so a collector stuck inside one cannot hold the database.
    const tables = this.#registry.tables(module.id);
    const moduleAccess: TableAccess = {
      read: new Set([...tables, "prism365_finding"]),
      write: new Set(tables),
    };
    const transaction = async <T>(
      operation: (sql: SqlClient) => Promise<T>,
      access?: TableAccess,
    ): Promise<T> => {
      live();
      outsideTransaction();
      inTransaction = true;
      try {
        return await this.#database.tenant(
          tenantId,
          async (sql) => {
            const result = await operation(sql);
            live();
            return result;
          },
          { signal: ended.signal, ...(access ? { access } : {}) },
        );
      } finally {
        inTransaction = false;
      }
    };
    // Promises handed to collector code are marked handled, so a call the collector never awaits
    // cannot become an unhandled rejection that ends the process. Awaiting still sees the error.
    const handled = <T>(promise: Promise<T>): Promise<T> => {
      promise.catch(() => {});
      return promise;
    };

    const client = this.#graphClient(async () => {
      live();
      return credential.accessToken(permissions);
    });
    const context: CollectorContext = {
      tenantId,
      moduleId: module.id,
      collectorId: collector.id,
      credentialKind: credential.kind,
      now: startedAt,
      signal: ended.signal,
      graph: (() => {
        const scoped = scopeGraph(
          client,
          module.packs,
          () => {
            live();
            outsideTransaction();
          },
          ended.signal,
        );
        return {
          get: (path, graphContext, schema) => handled(scoped.get(path, graphContext, schema)),
          getCollection: (path, graphContext, options) =>
            handled(scoped.getCollection(path, graphContext, options)),
          getReport: (path, graphContext, options) =>
            handled(scoped.getReport(path, graphContext, options)),
        } satisfies CollectorContext["graph"];
      })(),
      table: (name) => {
        if (!identifier.test(name)) throw new ModuleBoundaryError("Invalid module table name.");
        return `"${prefix}${name}"`;
      },
      db: (operation) => handled(transaction(operation, moduleAccess)),
      cursor: {
        get: () =>
          handled(
            (async () => {
              const result = await transaction((sql) =>
                sql.query<{ value: string }>(
                  "SELECT value FROM prism365_cursor WHERE tenant_id=$1 AND module=$2 AND collector=$3",
                  [tenantId, module.id, collector.id],
                ),
              );
              return result.rows[0]?.value ?? null;
            })(),
          ),
        set: (value) =>
          handled(
            (async () => {
              if (typeof value !== "string" || value.length === 0 || value.length > 2_048)
                throw new ModuleBoundaryError("A cursor is a string of 1 to 2048 characters.");
              await transaction((sql) =>
                sql.query(
                  `INSERT INTO prism365_cursor (tenant_id, module, collector, value, updated_at)
               VALUES ($1,$2,$3,$4,$5)
               ON CONFLICT (tenant_id, module, collector) DO UPDATE SET value=EXCLUDED.value, updated_at=EXCLUDED.updated_at`,
                  [tenantId, module.id, collector.id, value, this.#now().toISOString()],
                ),
              );
            })(),
          ),
      },
      findings: {
        upsert: (input: FindingInput) =>
          handled(
            (async () => {
              const { id, change } = await transaction((sql) =>
                upsertFinding(sql, tenantId, module.id, input),
              );
              if (change !== "updated") {
                counters.findings += 1;
                notifications.push({
                  tenantId,
                  module: module.id,
                  findingId: id,
                  ruleId: input.ruleId,
                  severity: input.severity,
                  title: input.title,
                  summary: input.summary,
                  subjectLabel: input.subject.label ?? null,
                  change,
                  observedAt: input.observedAt.toISOString(),
                });
              }
              return change;
            })(),
          ),
      },
      attention: {
        upsert: (input) =>
          handled(
            (async () => {
              const { change } = await transaction(async (sql) =>
                upsertAttention(
                  sql,
                  tenantId,
                  module.id,
                  input,
                  await readAttentionPreferences(sql, tenantId),
                ),
              );
              return change;
            })(),
          ),
        retire: (kind, keep, note) =>
          handled(
            transaction((sql) =>
              retireAttention(sql, tenantId, module.id, kind, keep, this.#now(), note),
            ),
          ),
      },
      count: {
        fetched: (count) => {
          counters.fetched = Math.min(counters.fetched + boundedCount(count), 2_147_483_647);
        },
        inserted: (count) => {
          counters.inserted = Math.min(counters.inserted + boundedCount(count), 2_147_483_647);
        },
        skipped: (count) => {
          counters.skipped = Math.min(counters.skipped + boundedCount(count), 2_147_483_647);
        },
      },
    };

    let status: RunStatus = "succeeded";
    let category: ErrorCategory | null = null;
    let more = false;
    const timer = setTimeout(() => {
      timedOut = true;
      ended.abort();
    }, this.#timeoutFor(collector));
    try {
      const result = await Promise.race([
        collector.run(context),
        new Promise<never>((_resolve, reject) => {
          ended.signal.addEventListener("abort", () => reject(new RunEndedError()), {
            once: true,
          });
        }),
      ]);
      if (result && result.complete === false) status = "partial";
      more = result?.more === true;
    } catch (error) {
      status = "failed";
      category = timedOut ? "timeout" : classifyCollectorError(error);
    } finally {
      clearTimeout(timer);
      ended.abort();
    }

    try {
      await this.#database.tenant(tenantId, (sql) =>
        sql.query(
          `UPDATE prism365_module_run SET status=$3, finished_at=$4, fetched=$5, inserted=$6,
             findings=$7, error_category=$8, skipped=$9 WHERE tenant_id=$1 AND id=$2`,
          [
            tenantId,
            runId,
            status,
            this.#now().toISOString(),
            counters.fetched,
            counters.inserted,
            counters.findings,
            category,
            counters.skipped,
          ],
        ),
      );
    } catch {
      this.#emit({ type: "maintenance.failed", tenantId, task: "recover" });
    }
    this.#emit({
      type: "run.finished",
      tenantId,
      module: module.id,
      collector: collector.id,
      status,
      category,
      ...counters,
    });
    if (category === "credential_unavailable" && this.#current(credential))
      this.unbind("credential_unavailable");
    else if (more && this.#current(credential))
      this.#catchingUp.add(this.#key(tenantId, module, collector));
    await this.#deliver(notifications);
  }

  async #deliver(notifications: readonly FindingNotification[]): Promise<void> {
    for (const sink of this.#sinks)
      for (const notification of notifications) {
        if (!severityAtLeast(notification.severity, sink.minimumSeverity)) continue;
        let timer: ReturnType<typeof setTimeout> | undefined;
        try {
          await Promise.race([
            sink.deliver(notification),
            new Promise<never>((_resolve, reject) => {
              timer = setTimeout(() => reject(new Error("sink timeout")), this.#sinkTimeout);
            }),
          ]);
        } catch {
          this.#emit({ type: "sink.failed", sink: sink.id });
        } finally {
          if (timer !== undefined) clearTimeout(timer);
        }
      }
  }

  // Deletes data past each module's declared retention, plus old run history, for one tenant.
  async #retention(tenantId: string): Promise<void> {
    const now = this.#now().getTime();
    const last = this.#lastRetention.get(tenantId);
    if (last !== undefined && now - last < this.#retentionInterval) return;
    try {
      await this.#database.tenant(tenantId, async (sql) => {
        for (const module of this.#registry.list())
          for (const rule of module.retention) {
            if (!identifier.test(rule.table) || !identifier.test(rule.timeColumn)) continue;
            const table = `"${moduleTablePrefix(module.id)}${rule.table}"`;
            await sql.query(
              `DELETE FROM ${table} WHERE tenant_id=$1 AND "${rule.timeColumn}" < $2`,
              [tenantId, new Date(now - rule.days * 86_400_000).toISOString()],
            );
          }
        // Notification-center items the provider closed or dropped are kept 90 days.
        await sql.query(
          "DELETE FROM prism365_attention WHERE tenant_id=$1 AND active=false AND last_seen < $2",
          [tenantId, new Date(now - 90 * 86_400_000).toISOString()],
        );
        await sql.query("DELETE FROM prism365_module_run WHERE tenant_id=$1 AND started_at < $2", [
          tenantId,
          new Date(now - this.#runHistoryDays * 86_400_000).toISOString(),
        ]);
      });
      this.#lastRetention.set(tenantId, now);
    } catch {
      this.#emit({ type: "maintenance.failed", tenantId, task: "retention" });
    }
  }
}

export interface ModuleRunRecord {
  module: string;
  collector: string;
  status: string;
  credentialKind: CredentialKind;
  startedAt: string;
  finishedAt: string | null;
  fetched: number;
  inserted: number;
  findings: number;
  skipped: number;
  errorCategory: string | null;
}

export async function latestRuns(sql: SqlClient, tenantId: string): Promise<ModuleRunRecord[]> {
  const result = await sql.query<{
    module: string;
    collector: string;
    status: string;
    credential_kind: CredentialKind;
    started_at: Date | string;
    finished_at: Date | string | null;
    fetched: number;
    inserted: number;
    findings: number;
    skipped: number;
    error_category: string | null;
  }>(
    `SELECT DISTINCT ON (module, collector) module, collector, status, credential_kind, started_at,
            finished_at, fetched, inserted, findings, skipped, error_category
       FROM prism365_module_run WHERE tenant_id=$1
      ORDER BY module, collector, started_at DESC`,
    [tenantId],
  );
  return result.rows.map((row) => ({
    module: row.module,
    collector: row.collector,
    status: row.status,
    credentialKind: row.credential_kind,
    startedAt: new Date(row.started_at).toISOString(),
    finishedAt: row.finished_at ? new Date(row.finished_at).toISOString() : null,
    fetched: row.fetched,
    inserted: row.inserted,
    findings: row.findings,
    skipped: row.skipped ?? 0,
    errorCategory: row.error_category,
  }));
}

import { createHmac, randomBytes, randomUUID } from "node:crypto";
import {
  auditDefinition,
  builtInAudits,
  findAudit,
  getRun,
  readTrendPoints,
} from "@prism365/audits";
import {
  type AlertChoices,
  type AlertOverview,
  type AlertRule,
  type AlertSettings,
  auditCsv,
  type alertRuleInputSchema,
  type Digest,
  type DeliveryError,
  type DeliveryRecord,
  type DeliverySource,
  type Destination,
  type DestinationCreated,
  type destinationCreateSchema,
  MAX_ALERT_RULES,
  MAX_DESTINATIONS,
  MAX_EXPORT_ATTACHMENT_BYTES,
  scopeFileSuffix,
} from "@prism365/contracts";
import {
  type CollectionStatus,
  type FindingNotification,
  type FindingSink,
  readCollectionIntent,
  type SchedulerEvent,
  type StopReason,
} from "@prism365/modules";
import {
  type ModuleDatabase,
  type SqlClient,
  StateCipher,
  type StateStore,
} from "@prism365/storage";
import { z } from "zod";
import {
  auditAlert,
  auditFires,
  auditMovement,
  COLLECTION_STOPPED_SUBJECT,
  type CollectionState,
  collectionAlert,
  collectionStoppedTooLong,
  exportAlert,
  findingAlert,
  findingMatches,
  weeklySummaryAlert,
  weeklySummarySlot,
} from "./alert-rules.ts";
import type { AuditStoredEvent } from "./audits.ts";
import { type ScheduleTimers, systemTimers } from "./audit-schedule.ts";
import { type HostLookup, parseWebhookUrl } from "./delivery-address.ts";
import {
  type AlertMessage,
  MAX_ALERT_ITEMS,
  minimise,
  renderEmail,
  renderWebhook,
} from "./delivery-payload.ts";
import {
  createSmtpSender,
  type DeliveryConfig,
  type MailSender,
  recipientAllowed,
  sendEmail,
} from "./delivery-smtp.ts";
import {
  changeAlertRules,
  type DestinationRow,
  deleteDestination,
  getDestination,
  insertDelivery,
  insertDestination,
  interruptDeliveries,
  listDestinations,
  pruneDeliveries,
  quietSubjects,
  readAlertExports,
  readAlertRules,
  readStoredRules,
  readWeekly,
  recentDeliveries,
  setAlertExport,
  setWeekly,
  subjectDigest,
  type StoredRules,
  summarySentSince,
  updateDelivery,
} from "./delivery-store.ts";
import {
  httpsConnector,
  newSigningSecret,
  sendWebhook,
  signedHeaders,
  systemHostLookup,
  type WebhookConnector,
} from "./delivery-webhook.ts";

// Alert delivery (docs/alerts.md, ADR-0011). Runs in the API process, off every request path:
//
// - The scheduler's finding sink only appends to a bounded buffer and returns, so a stuck
//   destination can never hold up collection; the buffer is evaluated a moment later.
// - Rules are evaluated one batch at a time, on their own lane, and each delivery they make is
//   written to the access ledger and the delivery log before anything leaves the server.
// - Sending happens on a small worker pool with its own deadlines. A failed attempt that may
//   succeed later is tried again after a backoff, three attempts in all, then marked failed.
//   Nothing waits for a send except the administrator's own Send test.
//
// The message is held in memory only until it is sent; a delivery a restart interrupts is marked
// so when the service starts again.

export const MAX_ATTEMPTS = 3;
export const RETRY_DELAYS_MILLISECONDS = [30_000, 120_000] as const;
// Workers in all; one destination never has more than one, so a destination that hangs holds up
// only itself.
const CONCURRENCY = 3;
const MAX_BUFFER = 1_000;
const MAX_QUEUE = 200;
const FLUSH_DELAY_MILLISECONDS = 2_000;
const CHECK_INTERVAL_MILLISECONDS = 5 * 60_000;
// One attempt, end to end. Webhooks have their own 10 second deadline inside this.
const ATTEMPT_TIMEOUT_MILLISECONDS = 45_000;
const TEST_WAIT_MILLISECONDS = 15_000;
// Teams refuses a card over 28 KB; every format keeps under it.
const MAX_BODY_BYTES = 27_000;

// Worth trying again: the destination or the network may answer later.
const retryable: ReadonlySet<DeliveryError> = new Set([
  "dns_failed",
  "connection_failed",
  "timeout",
  "rate_limited",
  "server_error",
  "error",
]);

export class DeliveryInputError extends Error {
  readonly code: string;
  readonly status: number;
  constructor(code: string, message: string, status = 400) {
    super(message);
    this.name = "DeliveryInputError";
    this.code = code;
    this.status = status;
  }
}

export interface DeliveryServiceOptions {
  tenantId: string;
  database: ModuleDatabase;
  store: Pick<StateStore, "appendAudit">;
  origin: string;
  config: DeliveryConfig;
  // STATE_ENCRYPTION_KEY. Without one, a key is made for this process only.
  encryptionKey?: string | null;
  // The scheduler's collection status, for collection rules.
  collection?: () => CollectionStatus;
  // The installed modules, which a finding rule may name.
  modules?: () => readonly { id: string; title: string }[];
  lookup?: HostLookup;
  connect?: WebhookConnector;
  // undefined: a sender on the configured relay. null: none.
  mailSender?: MailSender | null;
  now?: () => Date;
  timers?: ScheduleTimers;
  retryDelaysMilliseconds?: readonly number[];
  flushDelayMilliseconds?: number;
  checkIntervalMilliseconds?: number;
  attemptTimeoutMilliseconds?: number;
  log?: (message: string) => void;
  // The time zone the weekly summary goes out in, Monday 07:00 (SIGNINS_TIME_ZONE). UTC unset.
  timeZone?: string;
  // Builds the This week digest for the weekly summary, with the datasets it reads for the
  // ledger entry. Without it there is no weekly summary.
  weeklyDigest?: {
    build: (sql: SqlClient, now: Date) => Promise<Digest>;
    datasets: () => readonly string[];
  };
}

type Outcome = { ok: true } | { ok: false; error: DeliveryError };

interface Job {
  id: string;
  destinationId: string;
  message: AlertMessage;
  attachment?: { filename: string; content: Buffer };
  attempts: number;
  maxAttempts: number;
  settle: (result: { status: "sent" | "failed"; lastError: DeliveryError | null }) => void;
}

const secretSchema = z.object({
  url: z.string().min(1).max(2_048),
  signingSecret: z.string().min(20).max(200),
});

type DestinationInput = z.output<typeof destinationCreateSchema>;

export class DeliveryService {
  readonly #tenantId: string;
  readonly #database: ModuleDatabase;
  readonly #store: Pick<StateStore, "appendAudit">;
  readonly #origin: string;
  readonly #config: DeliveryConfig;
  readonly #cipher: StateCipher;
  readonly #keyId: string;
  readonly #temporaryKey: boolean;
  readonly #collection: (() => CollectionStatus) | null;
  readonly #modules: (() => readonly { id: string; title: string }[]) | null;
  readonly #lookup: HostLookup;
  readonly #connect: WebhookConnector;
  readonly #mailSender: MailSender | null;
  readonly #now: () => Date;
  readonly #timers: ScheduleTimers;
  readonly #retryDelays: readonly number[];
  readonly #flushDelay: number;
  readonly #checkInterval: number;
  readonly #attemptTimeout: number;
  readonly #log: (message: string) => void;
  readonly #timeZone: string;
  readonly #weeklyDigest: DeliveryServiceOptions["weeklyDigest"] | null;
  readonly #startedAt: Date;
  readonly #buffer: FindingNotification[] = [];
  // Destinations with an attempt in progress.
  readonly #busy = new Set<string>();
  // Findings left out of a full buffer since the last batch, by the rule that wanted them.
  readonly #dropped = new Map<string, number>();
  // The rules, as last read: what the sink takes in. null until first read, when it takes all.
  #rules: AlertRule[] | null = null;
  // Counts rule saves. A read that began before the latest save never replaces the cache, so a
  // slow read cannot hide a rule just saved.
  #rulesGeneration = 0;
  readonly #queue: Job[] = [];
  // Jobs waiting for their next attempt, with the timer that will queue them.
  readonly #waiting = new Map<Job, () => void>();
  readonly #work = new Set<Promise<unknown>>();
  #active = 0;
  #lane: Promise<void> = Promise.resolve();
  #cancelFlush: (() => void) | null = null;
  #cancelCheck: (() => void) | null = null;
  #lastPrune = 0;
  #started = false;
  #stopped = false;
  // When collection last stopped, and why. A process starts with nothing collecting.
  #stoppedSince: Date | null;
  #stopReason: StopReason | null = "shutdown";

  constructor(options: DeliveryServiceOptions) {
    this.#tenantId = options.tenantId.toLowerCase();
    this.#database = options.database;
    this.#store = options.store;
    this.#origin = options.origin;
    this.#config = options.config;
    const key = options.encryptionKey ?? randomBytes(32).toString("hex");
    this.#temporaryKey = !options.encryptionKey;
    this.#cipher = new StateCipher(key);
    // Names the key without revealing it, so a destination written under another key is known
    // to be unreadable without trying to decrypt it.
    this.#keyId = createHmac("sha256", Buffer.from(key, "hex"))
      .update("prism365.alerts.destination-key")
      .digest("hex")
      .slice(0, 16);
    this.#collection = options.collection ?? null;
    this.#modules = options.modules ?? null;
    this.#lookup = options.lookup ?? systemHostLookup;
    this.#connect = options.connect ?? httpsConnector();
    this.#mailSender =
      options.mailSender !== undefined
        ? options.mailSender
        : options.config.email.smtp
          ? createSmtpSender(options.config.email.smtp)
          : null;
    this.#now = options.now ?? (() => new Date());
    this.#timers = options.timers ?? systemTimers;
    this.#retryDelays = options.retryDelaysMilliseconds ?? RETRY_DELAYS_MILLISECONDS;
    this.#flushDelay = options.flushDelayMilliseconds ?? FLUSH_DELAY_MILLISECONDS;
    this.#checkInterval = options.checkIntervalMilliseconds ?? CHECK_INTERVAL_MILLISECONDS;
    this.#attemptTimeout = options.attemptTimeoutMilliseconds ?? ATTEMPT_TIMEOUT_MILLISECONDS;
    this.#log = options.log ?? ((message) => console.warn(message));
    this.#timeZone = options.timeZone ?? "UTC";
    this.#weeklyDigest = options.weeklyDigest ?? null;
    this.#startedAt = this.#now();
    this.#stoppedSince = this.#startedAt;
  }

  // The scheduler awaits a sink with its own timeout, so this only buffers.
  readonly sink: FindingSink = {
    id: "alerts",
    minimumSeverity: "info",
    deliver: async (notification) => {
      this.#accept(notification);
    },
  };

  start(): void {
    if (this.#started || this.#stopped) return;
    this.#started = true;
    void this.#track(this.#refreshRules().catch(() => []));
    void this.#track(
      this.#db(async (sql) => {
        const interrupted = await interruptDeliveries(sql, this.#tenantId, this.#startedAt);
        await pruneDeliveries(sql, this.#tenantId, this.#now());
        const destinations = await listDestinations(sql, this.#tenantId);
        return { interrupted, refused: destinations.filter((row) => !this.#allowed(row)).length };
      }).then(
        ({ interrupted, refused }) => {
          this.#lastPrune = this.#now().getTime();
          if (interrupted > 0)
            this.#log(`Prism365 alerts: ${interrupted} deliveries were interrupted by a restart.`);
          // Saved under an older, wider rule or before a domain left ALERT_EMAIL_DOMAINS: each
          // is refused at every send and marked on the page until it is added again.
          if (refused > 0)
            this.#log(
              `Prism365 alerts: ${refused} email destinations have recipients this server refuses; nothing is sent to them.`,
            );
        },
        () => this.#log("Prism365 alerts could not tidy the delivery log (storage)."),
      ),
    );
    this.#cancelCheck = this.#timers.every(this.#checkInterval, () => {
      void this.#evaluate(() => this.#periodic());
    });
  }

  // Stops taking work. Deliveries still waiting are marked interrupted; attempts in progress get
  // a few seconds to finish.
  async stop(): Promise<void> {
    if (this.#stopped) return;
    this.#stopped = true;
    this.#cancelFlush?.();
    this.#cancelFlush = null;
    this.#cancelCheck?.();
    this.#cancelCheck = null;
    const waiting = [...this.#queue, ...this.#waiting.keys()];
    for (const cancel of this.#waiting.values()) cancel();
    this.#waiting.clear();
    this.#queue.length = 0;
    for (const job of waiting) {
      await this.#record(job, "failed", "interrupted");
      job.settle({ status: "failed", lastError: "interrupted" });
    }
    await Promise.race([
      Promise.allSettled([...this.#work]),
      new Promise((resolve) => setTimeout(resolve, 5_000).unref()),
    ]);
  }

  // Resolves once nothing is buffered, queued, waiting or in progress. For tests.
  async idle(): Promise<void> {
    for (let round = 0; round < 1_000; round += 1) {
      if (
        this.#work.size === 0 &&
        this.#cancelFlush === null &&
        this.#queue.length === 0 &&
        this.#waiting.size === 0 &&
        this.#active === 0
      )
        return;
      await Promise.allSettled([...this.#work]);
      await new Promise((resolve) => setTimeout(resolve, 1));
    }
  }

  settings(): AlertSettings {
    return {
      email: { configured: this.#config.email.smtp !== null, domains: this.#config.email.domains },
      webhookPorts: this.#config.webhookPorts,
      temporaryKey: this.#temporaryKey,
    };
  }

  async overview(): Promise<AlertOverview> {
    const { destinations, stored, deliveries, exports, weekly, audits } = await this.#db(
      async (sql) => {
        const audits: AlertChoices["audits"] = [];
        // The tiles each audit's newest run had, which is what an audit rule can watch.
        for (const audit of builtInAudits.slice(0, 50)) {
          const { points } = await readTrendPoints(
            sql,
            this.#tenantId,
            audit.definition.id,
            new Date(0),
            1,
          );
          audits.push({
            id: audit.definition.id,
            title: audit.definition.title,
            tiles: (points.at(-1)?.tiles ?? []).map(({ id, label }) => ({ id, label })),
          });
        }
        return {
          destinations: await listDestinations(sql, this.#tenantId),
          stored: await readStoredRules(sql, this.#tenantId),
          deliveries: await recentDeliveries(sql, this.#tenantId, 50),
          exports: await readAlertExports(sql, this.#tenantId),
          weekly: await readWeekly(sql, this.#tenantId),
          audits,
        };
      },
    );
    return {
      settings: this.settings(),
      exports: exports.map(({ auditId, destinationId }) => ({ auditId, destinationId })),
      weekly: { destinationId: weekly?.destinationId ?? null, timeZone: this.#timeZone },
      choices: { modules: [...(this.#modules?.() ?? [])].slice(0, 50), audits },
      destinations: destinations.slice(0, MAX_DESTINATIONS).map((row) => this.#public(row)),
      rules: stored.rules,
      otherRules: stored.others,
      deliveries,
    };
  }

  async createDestination(
    actorId: string,
    input: DestinationInput,
    beforeSave: (id: string) => Promise<void>,
  ): Promise<DestinationCreated> {
    const id = randomUUID();
    const createdAt = this.#now().toISOString();
    let row: DestinationRow;
    let signingSecret: string | null = null;
    if (input.kind === "email") {
      const { smtp, domains } = this.#config.email;
      if (!smtp)
        throw new DeliveryInputError(
          "email_not_configured",
          "Email alerts need SMTP_URL, SMTP_FROM and ALERT_EMAIL_DOMAINS on the server.",
          409,
        );
      if (!input.recipients.every((address) => recipientAllowed(address, domains)))
        throw new DeliveryInputError(
          "recipient_not_allowed",
          `Recipients must be in ${domains.join(", ")}, the domains ALERT_EMAIL_DOMAINS allows.`,
        );
      row = {
        id,
        kind: "email",
        label: input.label,
        host: null,
        recipients: input.recipients,
        secret: null,
        keyId: null,
        includeDetails: input.includeDetails,
        createdAt,
        createdBy: actorId,
      };
    } else {
      const parsed = parseWebhookUrl(input.url, this.#config.webhookPorts);
      if (!parsed.ok)
        throw new DeliveryInputError(
          parsed.problem,
          parsed.problem === "port_not_allowed"
            ? "That port is not allowed. Webhooks use port 443 unless WEBHOOK_ALLOWED_PORTS lists another."
            : "Use the webhook's full HTTPS address with a host name, without a user name or password.",
        );
      signingSecret = newSigningSecret();
      row = {
        id,
        kind: input.kind,
        label: input.label,
        host: parsed.url.host,
        recipients: [],
        secret: this.#cipher.encrypt({ url: parsed.url.href, signingSecret }, this.#binding(id)),
        keyId: this.#keyId,
        includeDetails: input.includeDetails,
        createdAt,
        createdBy: actorId,
      };
    }
    const existing = await this.#db((sql) => listDestinations(sql, this.#tenantId));
    if (existing.length >= MAX_DESTINATIONS)
      throw new DeliveryInputError(
        "too_many_destinations",
        `A tenant can have at most ${MAX_DESTINATIONS} destinations. Delete one first.`,
      );
    await beforeSave(id);
    await this.#db((sql) => insertDestination(sql, this.#tenantId, row));
    return { destination: this.#public(row), signingSecret };
  }

  // false when there is no such destination. A destination a rule still sends to is not
  // deleted: the rule would go quiet without anyone choosing that.
  async deleteDestination(id: string, beforeDelete: () => Promise<void>): Promise<boolean> {
    const { destination, rules, exports, weekly } = await this.#db(async (sql) => ({
      destination: await getDestination(sql, this.#tenantId, id),
      rules: await readAlertRules(sql, this.#tenantId),
      exports: await readAlertExports(sql, this.#tenantId),
      weekly: await readWeekly(sql, this.#tenantId),
    }));
    if (!destination) return false;
    const using = rules.filter((rule) => rule.destinations.includes(id));
    if (using.length > 0)
      throw new DeliveryInputError(
        "destination_in_use",
        `${using.length === 1 ? "A rule sends" : `${using.length} rules send`} to this destination. Remove it from ${using.length === 1 ? "that rule" : "those rules"} first.`,
        409,
      );
    if (exports.some((entry) => entry.destinationId === id) || weekly?.destinationId === id)
      throw new DeliveryInputError(
        "destination_in_use",
        "A scheduled export or the weekly summary goes to this destination. Choose another for it first.",
        409,
      );
    await beforeDelete();
    return this.#db((sql) => deleteDestination(sql, this.#tenantId, id));
  }

  // Sends a fixed test message once, without retries, and waits a while for the outcome. The
  // ledger entry is written before anything is sent, as for every delivery. null when there is
  // no such destination.
  async sendTest(actorId: string, id: string): Promise<DeliveryRecord | null> {
    const destination = await this.#db((sql) => getDestination(sql, this.#tenantId, id));
    if (!destination) return null;
    const queued = await this.#enqueue({
      source: "test",
      ruleId: null,
      actorId,
      destination,
      message: {
        kind: "test",
        title: "Test alert from Prism365",
        severity: null,
        count: 0,
        items: [],
        lines: [
          "If you can read this, Prism365 alerts can reach this destination. Nothing needs doing.",
        ],
        link: "/#/alerts",
      },
      subjects: [],
      items: 0,
      maxAttempts: 1,
    });
    await Promise.race([
      queued.done,
      new Promise((resolve) => setTimeout(resolve, TEST_WAIT_MILLISECONDS).unref()),
    ]);
    const records = await this.#db((sql) => recentDeliveries(sql, this.#tenantId, 100));
    return records.find((record) => record.id === queued.id) ?? null;
  }

  async createRule(
    actorId: string,
    input: z.output<typeof alertRuleInputSchema>,
    beforeSave: (id: string) => Promise<void>,
  ): Promise<AlertRule[]> {
    const id = randomUUID();
    const destinations = await this.#db((sql) => listDestinations(sql, this.#tenantId));
    const known = new Set(destinations.map((row) => row.id));
    if (!input.destinations.every((destination) => known.has(destination)))
      throw new DeliveryInputError(
        "unknown_destination",
        "One of those destinations does not exist. Refresh the page and choose again.",
      );
    if (input.kind === "audit" && !findAudit(input.auditId))
      throw new DeliveryInputError("unknown_audit", "There is no audit with that id.");
    if (input.kind === "finding" && this.#modules) {
      const modules = new Set(this.#modules().map((module) => module.id));
      if (!input.modules.every((module) => modules.has(module)))
        throw new DeliveryInputError("unknown_module", "One of those modules is not installed.");
    }
    const rule = {
      ...input,
      id,
      enabled: true,
      createdAt: this.#now().toISOString(),
      createdBy: actorId,
    } as AlertRule;
    await beforeSave(id);
    const saved = await this.#saveRules((sql) =>
      changeAlertRules(sql, this.#tenantId, actorId, { add: rule }),
    );
    if (!saved)
      throw new DeliveryInputError(
        "too_many_rules",
        `A tenant can have at most ${MAX_ALERT_RULES} alert rules. Delete one first.`,
      );
    return saved;
  }

  // null when there is no such rule.
  async changeRule(
    actorId: string,
    id: string,
    change: { enabled: boolean } | "delete",
    beforeSave: () => Promise<void>,
  ): Promise<AlertRule[] | null> {
    // Any stored rule with this id, including one this version keeps but does not run.
    const current = await this.#db((sql) => readStoredRules(sql, this.#tenantId));
    if (![...current.rules, ...current.others].some((rule) => rule.id === id)) return null;
    await beforeSave();
    return this.#saveRules((sql) =>
      changeAlertRules(
        sql,
        this.#tenantId,
        actorId,
        change === "delete" ? { remove: id } : { id, enabled: change.enabled },
      ),
    );
  }

  // Saves a change to the rules and puts the result in the cache, marking every read that began
  // before it as stale.
  async #saveRules(
    save: (sql: SqlClient) => Promise<StoredRules | null>,
  ): Promise<AlertRule[] | null> {
    const saved = await this.#db(save);
    if (!saved) return null;
    this.#rulesGeneration += 1;
    this.#rules = saved.rules;
    return saved.rules;
  }

  // Reads the rules and caches them, unless a save happened while reading.
  async #refreshRules(): Promise<AlertRule[]> {
    const generation = this.#rulesGeneration;
    const rules = await this.#db((sql) => readAlertRules(sql, this.#tenantId));
    if (generation === this.#rulesGeneration) this.#rules = rules;
    return rules;
  }

  onSchedulerEvent(event: SchedulerEvent): void {
    if (event.type === "collection.bound" && event.tenantId.toLowerCase() === this.#tenantId) {
      this.#stoppedSince = null;
      this.#stopReason = null;
    }
    if (event.type === "collection.unbound" && event.tenantId.toLowerCase() === this.#tenantId) {
      this.#stoppedSince = this.#now();
      this.#stopReason = event.reason;
    }
  }

  // Where an audit's scheduled runs are sent, or null to stop sending them.
  async setExport(
    actorId: string,
    auditId: string,
    destinationId: string | null,
    beforeSave: () => Promise<void>,
  ): Promise<void> {
    if (!findAudit(auditId))
      throw new DeliveryInputError("unknown_audit", "There is no audit with that id.");
    await this.#knownDestination(destinationId);
    await beforeSave();
    const saved = await this.#db((sql) =>
      setAlertExport(sql, this.#tenantId, actorId, auditId, destinationId),
    );
    if (!saved)
      throw new DeliveryInputError(
        "too_many_exports",
        "This installation keeps at most 50 scheduled exports. Stop one first.",
      );
  }

  // Where the weekly summary goes, or null for none.
  async setWeekly(
    actorId: string,
    destinationId: string | null,
    beforeSave: () => Promise<void>,
  ): Promise<void> {
    if (destinationId && !this.#weeklyDigest)
      throw new DeliveryInputError(
        "summary_unavailable",
        "The weekly summary is not available in this installation.",
        409,
      );
    await this.#knownDestination(destinationId);
    await beforeSave();
    await this.#db((sql) => setWeekly(sql, this.#tenantId, actorId, destinationId));
  }

  async #knownDestination(id: string | null): Promise<void> {
    if (id === null) return;
    const destination = await this.#db((sql) => getDestination(sql, this.#tenantId, id));
    if (!destination)
      throw new DeliveryInputError(
        "unknown_destination",
        "That destination does not exist. Refresh the page and choose again.",
      );
  }

  // An audit run was stored with a new trend point: audit rules for it are checked, and a
  // scheduled run is sent where its audit's export goes.
  onAuditStored(event: AuditStoredEvent): void {
    if (this.#stopped || event.tenantId.toLowerCase() !== this.#tenantId) return;
    void this.#evaluate(() => this.#auditStored(event));
    if (event.trigger === "schedule") void this.#evaluate(() => this.#exportRun(event));
  }

  // Takes in only what an enabled rule wants. A full buffer gives up its lowest-severity, oldest
  // finding for a new one, or the new one when it is the lowest; either way the loss is counted
  // for each rule that wanted it, and recorded with the next batch.
  #accept(notification: FindingNotification): void {
    if (this.#stopped || notification.tenantId.toLowerCase() !== this.#tenantId) return;
    const rules = this.#rules;
    if (rules !== null && !rules.some((rule) => findingMatches(rule, notification))) return;
    if (this.#buffer.length >= MAX_BUFFER) {
      let lowest = 0;
      for (let index = 1; index < this.#buffer.length; index += 1)
        if (rank(this.#buffer[index]) < rank(this.#buffer[lowest])) lowest = index;
      const keepNew = rank(notification) >= rank(this.#buffer[lowest]);
      const lost = keepNew ? this.#buffer.splice(lowest, 1)[0] : notification;
      if (lost)
        for (const rule of rules ?? [])
          if (findingMatches(rule, lost))
            this.#dropped.set(rule.id, (this.#dropped.get(rule.id) ?? 0) + 1);
      if (!keepNew) return;
    }
    this.#buffer.push(notification);
    if (this.#cancelFlush) return;
    this.#cancelFlush = this.#timers.after(this.#flushDelay, () => {
      this.#cancelFlush = null;
      void this.#evaluate(() => this.#flush());
    });
  }

  #track<T>(promise: Promise<T>): Promise<T> {
    this.#work.add(promise);
    void promise.then(
      () => this.#work.delete(promise),
      () => this.#work.delete(promise),
    );
    return promise;
  }

  // Rule evaluations run one at a time, so two batches can never both pass one quiet period.
  #evaluate(work: () => Promise<void>): Promise<void> {
    const next = this.#lane.then(work).catch(() => {
      this.#log("Prism365 alerts could not evaluate rules (storage).");
    });
    this.#lane = next;
    return this.#track(next);
  }

  #db<T>(operation: (sql: SqlClient) => Promise<T>): Promise<T> {
    return this.#database.tenant(this.#tenantId, operation, { timeoutMilliseconds: 15_000 });
  }

  #binding(id: string): string {
    return `${this.#tenantId}:alerts.destination:${id}`;
  }

  #public(row: DestinationRow): Destination {
    return {
      id: row.id,
      kind: row.kind,
      label: row.label,
      host: row.host,
      recipients: row.recipients,
      includeDetails: row.includeDetails,
      readable: row.kind === "email" || row.keyId === this.#keyId,
      recipientsAllowed: this.#allowed(row),
      createdAt: row.createdAt,
    };
  }

  #allowed(row: DestinationRow): boolean {
    return (
      row.kind !== "email" ||
      row.recipients.every((address) => recipientAllowed(address, this.#config.email.domains))
    );
  }

  #secret(row: DestinationRow): z.infer<typeof secretSchema> | null {
    if (!row.secret || row.keyId !== this.#keyId) return null;
    try {
      const parsed = secretSchema.safeParse(
        this.#cipher.decrypt(row.secret, this.#binding(row.id)),
      );
      return parsed.success ? parsed.data : null;
    } catch {
      return null;
    }
  }

  async #flush(): Promise<void> {
    const batch = this.#buffer.splice(0);
    const dropped = new Map(this.#dropped);
    this.#dropped.clear();
    const rules = await this.#refreshRules();
    for (const [ruleId, count] of dropped) {
      const rule = rules.find((candidate) => candidate.id === ruleId);
      if (!rule) continue;
      this.#log(
        `Prism365 alerts left out ${count} findings for one rule that arrived all at once.`,
      );
      try {
        await this.#recordDropped(rule, count);
      } catch {
        this.#log("Prism365 alerts could not record left-out findings (storage).");
      }
    }
    if (batch.length === 0) return;
    for (const rule of rules) {
      if (!rule.enabled || rule.kind === "audit") continue;
      // The newest notification for each finding.
      const matched = new Map<string, FindingNotification>();
      for (const notification of batch)
        if (findingMatches(rule, notification)) matched.set(notification.findingId, notification);
      if (matched.size === 0) continue;
      // One rule failing to send leaves the others to run.
      try {
        await this.#fire(rule, [...matched.keys()], async (fresh, withDetails) => {
          const chosen = [...matched.values()].filter((item) => fresh.has(item.findingId));
          const details = withDetails
            ? await this.#findingDetails(chosen.map((item) => item.findingId))
            : new Map();
          return findingAlert(rule, chosen, details);
        });
      } catch {
        this.#log("Prism365 alerts could not evaluate a rule (storage).");
      }
    }
  }

  // Findings a full buffer left out, as one failed delivery row per destination of the rule, so the
  // page shows them, and one ledger entry with the count.
  async #recordDropped(rule: AlertRule, count: number): Promise<void> {
    const destinations = (await this.#db((sql) => listDestinations(sql, this.#tenantId))).filter(
      (row) => rule.destinations.includes(row.id),
    );
    const at = this.#now().toISOString();
    await this.#store.appendAudit(this.#tenantId, rule.createdBy, {
      type: "alerts.dropped",
      ruleId: rule.id,
      items: count,
      occurredAt: at,
    });
    for (const destination of destinations) {
      const id = randomUUID();
      await this.#db(async (sql) => {
        await insertDelivery(sql, this.#tenantId, {
          id,
          destinationId: destination.id,
          source: "rule",
          ruleId: rule.id,
          subjects: [],
          items: count,
          createdAt: at,
        });
        await updateDelivery(sql, this.#tenantId, id, {
          status: "failed",
          attempts: 0,
          lastError: "dropped",
          at,
        });
      });
    }
  }

  async #auditStored(event: AuditStoredEvent): Promise<void> {
    const rules = (await this.#db((sql) => readAlertRules(sql, this.#tenantId))).filter(
      (rule): rule is Extract<AlertRule, { kind: "audit" }> =>
        rule.enabled && rule.kind === "audit" && rule.auditId === event.auditId,
    );
    if (rules.length === 0) return;
    const { points } = await this.#db((sql) =>
      readTrendPoints(sql, this.#tenantId, event.auditId, new Date(0), 20),
    );
    // A later run already wrote a newer point; its own event checks it.
    const newest = points.at(-1);
    if (newest?.runId !== event.runId) return;
    const title = findAudit(event.auditId)?.definition.title ?? event.auditId;
    for (const rule of rules) {
      const movement = auditMovement(points, rule.tileId);
      if (!movement || !auditFires(rule, movement)) continue;
      const tileLabel = newest.tiles.find((tile) => tile.id === rule.tileId)?.label ?? rule.tileId;
      await this.#fire(rule, [`${rule.auditId}\n${rule.tileId}`], async () =>
        auditAlert(rule, { title, tileLabel }, movement),
      );
    }
  }

  // The run's CSV, as the export route writes it, attached only to email with details on (the
  // file carries every column, names and addresses included) and only up to 5 MB. Everyone else
  // gets the counts and a link.
  async #exportRun(event: AuditStoredEvent): Promise<void> {
    const entry = (await this.#db((sql) => readAlertExports(sql, this.#tenantId))).find(
      (candidate) => candidate.auditId === event.auditId,
    );
    const audit = findAudit(event.auditId);
    if (!entry || !audit) return;
    const { destination, run } = await this.#db(async (sql) => ({
      destination: await getDestination(sql, this.#tenantId, entry.destinationId),
      run: await getRun(sql, this.#tenantId, event.runId),
    }));
    if (!destination || !run || run.status !== "succeeded") return;
    const content =
      destination.kind === "email" && destination.includeDetails
        ? Buffer.from(auditCsv(auditDefinition(audit), run.rows))
        : null;
    const withheld = exportWithheld(destination, content?.length ?? 0);
    const attachment =
      content && withheld === null
        ? {
            filename: `prism365-${run.auditId}-${run.startedAt.slice(0, 10)}${scopeFileSuffix(run.scope)}.csv`,
            content,
          }
        : undefined;
    await this.#enqueue({
      source: "export",
      ruleId: null,
      actorId: entry.setBy,
      destination,
      message: exportAlert({
        auditId: run.auditId,
        title: audit.definition.title,
        rows: run.rows.length,
        complete: run.complete,
        scoped: run.scope !== null && run.scope.length > 0,
        attached: attachment !== undefined,
        withheld,
      }),
      ...(attachment ? { attachment } : {}),
      subjects: [],
      items: run.rows.length,
      maxAttempts: MAX_ATTEMPTS,
      ledger: { auditId: run.auditId, runId: run.id },
    });
  }

  // The weekly summary, once per week from Monday 07:00 in the configured time zone. A server that
  // was down then sends it when it comes back the same day; later than that, the week is skipped.
  async #weekly(now: Date): Promise<void> {
    if (!this.#weeklyDigest) return;
    const slot = weeklySummarySlot(now, this.#timeZone);
    if (now.getTime() - slot.getTime() > 86_400_000) return;
    const { weekly, sent } = await this.#db(async (sql) => ({
      weekly: await readWeekly(sql, this.#tenantId),
      sent: await summarySentSince(sql, this.#tenantId, slot),
    }));
    if (!weekly || sent) return;
    const destination = await this.#db((sql) =>
      getDestination(sql, this.#tenantId, weekly.destinationId),
    );
    if (!destination) return;
    const builder = this.#weeklyDigest;
    const digest = await this.#db((sql) => builder.build(sql, now));
    await this.#enqueue({
      source: "summary",
      ruleId: null,
      actorId: weekly.setBy,
      destination,
      message: weeklySummaryAlert(digest),
      subjects: [],
      items: 0,
      maxAttempts: MAX_ATTEMPTS,
      ledger: { datasets: [...builder.datasets()] },
    });
  }

  async #periodic(): Promise<void> {
    const now = this.#now();
    await this.#refreshRules();
    if (now.getTime() - this.#lastPrune >= 86_400_000) {
      this.#lastPrune = now.getTime();
      await this.#db((sql) => pruneDeliveries(sql, this.#tenantId, now));
    }
    await this.#weekly(now);
    const status = this.#collection?.();
    if (!status) return;
    const collecting = status.state === "collecting";
    if (collecting) return;
    const rules = (await this.#db((sql) => readAlertRules(sql, this.#tenantId))).filter(
      (rule): rule is Extract<AlertRule, { kind: "collection" }> =>
        rule.enabled && rule.kind === "collection" && rule.stoppedHours !== null,
    );
    if (rules.length === 0) return;
    const intent = await this.#db((sql) => readCollectionIntent(sql, this.#tenantId));
    const state: CollectionState = {
      intended: intent !== null,
      collecting,
      stoppedSince: this.#stoppedSince,
      reason: status.lastStopReason ?? this.#stopReason,
    };
    for (const rule of rules)
      if (collectionStoppedTooLong(rule, state, now))
        await this.#fire(rule, [COLLECTION_STOPPED_SUBJECT], async () =>
          collectionAlert(rule, state, now),
        );
  }

  // Sends one alert for a rule to each of its destinations, for the subjects outside the rule's
  // quiet period. `build` gets those subjects and whether any destination wants details.
  async #fire(
    rule: AlertRule,
    subjects: readonly string[],
    build: (fresh: ReadonlySet<string>, withDetails: boolean) => Promise<AlertMessage>,
  ): Promise<void> {
    const now = this.#now();
    const { quiet, destinations } = await this.#db(async (sql) => ({
      quiet: await quietSubjects(
        sql,
        this.#tenantId,
        rule.id,
        new Date(now.getTime() - rule.quietMinutes * 60_000),
      ),
      destinations: await listDestinations(sql, this.#tenantId),
    }));
    const fresh = subjects.filter((subject) => !quiet.has(subjectDigest(rule.id, subject)));
    if (fresh.length === 0) return;
    const targets = destinations.filter((row) => rule.destinations.includes(row.id));
    if (targets.length === 0) return;
    const message = await build(
      new Set(fresh),
      targets.some((row) => row.includeDetails),
    );
    for (const destination of targets)
      await this.#enqueue({
        source: "rule",
        ruleId: rule.id,
        actorId: rule.createdBy,
        destination,
        message,
        subjects: fresh.map((subject) => subjectDigest(rule.id, subject)),
        items: message.count,
        maxAttempts: MAX_ATTEMPTS,
      });
  }

  // Evidence fields of the findings, for destinations with details on: the first evidence
  // record's plain values, at most twelve.
  async #findingDetails(
    ids: readonly string[],
  ): Promise<Map<string, Record<string, string | number | boolean | null>>> {
    const details = new Map<string, Record<string, string | number | boolean | null>>();
    if (ids.length === 0) return details;
    const result = await this.#db((sql) =>
      sql.query<{ id: string; evidence: unknown }>(
        "SELECT id, evidence FROM prism365_finding WHERE tenant_id=$1 AND id = ANY($2::uuid[])",
        [this.#tenantId, ids.slice(0, 10)],
      ),
    );
    for (const row of result.rows) {
      const evidence = typeof row.evidence === "string" ? safeJson(row.evidence) : row.evidence;
      const first = Array.isArray(evidence) ? (evidence[0] as { fields?: unknown }) : null;
      const fields =
        first?.fields && typeof first.fields === "object" && !Array.isArray(first.fields)
          ? (first.fields as Record<string, unknown>)
          : {};
      const kept: Record<string, string | number | boolean | null> = {};
      for (const [key, value] of Object.entries(fields)) {
        if (Object.keys(kept).length >= 12) break;
        if (!/^[A-Za-z][A-Za-z0-9_]{0,40}$/.test(key)) continue;
        if (
          value === null ||
          typeof value === "boolean" ||
          (typeof value === "number" && Number.isFinite(value))
        )
          kept[key] = value;
        else if (typeof value === "string") kept[key] = value.slice(0, 200);
      }
      details.set(row.id, kept);
    }
    return details;
  }

  // Writes the ledger entry and the delivery row, then queues the send. The ledger entry comes
  // first: nothing leaves the server unrecorded. It holds ids and a count, never the message.
  async #enqueue(input: {
    source: DeliverySource;
    ruleId: string | null;
    actorId: string;
    destination: DestinationRow;
    message: AlertMessage;
    attachment?: { filename: string; content: Buffer };
    subjects: readonly string[];
    items: number;
    maxAttempts: number;
    // More ids for the ledger entry, such as the audit run an export sends.
    ledger?: Record<string, string | readonly string[]>;
  }): Promise<{ id: string; done: Promise<unknown> }> {
    const id = randomUUID();
    const createdAt = this.#now().toISOString();
    let recorded = true;
    try {
      await this.#store.appendAudit(this.#tenantId, input.actorId, {
        type: "alerts.delivery",
        source: input.source,
        ruleId: input.ruleId,
        destinationId: input.destination.id,
        deliveryId: id,
        items: input.items,
        ...input.ledger,
        occurredAt: createdAt,
      });
    } catch {
      recorded = false;
    }
    await this.#db((sql) =>
      insertDelivery(sql, this.#tenantId, {
        id,
        destinationId: input.destination.id,
        source: input.source,
        ruleId: input.ruleId,
        subjects: input.subjects,
        items: input.items,
        createdAt,
      }),
    );
    const refuse = async (error: DeliveryError) => {
      await this.#db((sql) =>
        updateDelivery(sql, this.#tenantId, id, {
          status: "failed",
          attempts: 0,
          lastError: error,
          at: this.#now().toISOString(),
        }),
      );
      return { id, done: Promise.resolve() };
    };
    if (!recorded) return refuse("ledger_failed");
    if (this.#stopped) return refuse("interrupted");
    if (this.#queue.length >= MAX_QUEUE) return refuse("queue_full");
    let settle: Job["settle"] = () => {};
    const done = new Promise<unknown>((resolve) => {
      settle = resolve;
    });
    this.#queue.push({
      id,
      destinationId: input.destination.id,
      message: minimise(input.message, input.destination),
      ...(input.attachment ? { attachment: input.attachment } : {}),
      attempts: 0,
      maxAttempts: input.maxAttempts,
      settle,
    });
    this.#pump();
    return { id, done };
  }

  // Starts the oldest job whose destination has nothing in progress, while workers are free.
  #pump(): void {
    while (!this.#stopped && this.#active < CONCURRENCY) {
      const index = this.#queue.findIndex((job) => !this.#busy.has(job.destinationId));
      if (index < 0) return;
      const [job] = this.#queue.splice(index, 1);
      if (!job) return;
      this.#active += 1;
      this.#busy.add(job.destinationId);
      void this.#track(
        this.#attempt(job)
          .catch(() => this.#log("Prism365 alerts could not record a delivery (storage)."))
          .finally(() => {
            this.#active -= 1;
            this.#busy.delete(job.destinationId);
            this.#pump();
          }),
      );
    }
  }

  async #record(
    job: Job,
    status: "pending" | "sent" | "failed",
    lastError: DeliveryError | null,
  ): Promise<void> {
    try {
      await this.#db((sql) =>
        updateDelivery(sql, this.#tenantId, job.id, {
          status,
          attempts: job.attempts,
          lastError,
          at: this.#now().toISOString(),
        }),
      );
    } catch {
      this.#log("Prism365 alerts could not record a delivery (storage).");
    }
  }

  async #attempt(job: Job): Promise<void> {
    job.attempts += 1;
    let outcome: Outcome;
    try {
      // Read again for every attempt: the destination may have been deleted in between.
      const destination = await this.#db((sql) =>
        getDestination(sql, this.#tenantId, job.destinationId),
      );
      outcome = destination
        ? await bounded((signal) => this.#send(destination, job, signal), this.#attemptTimeout)
        : { ok: false, error: "destination_deleted" };
    } catch {
      outcome = { ok: false, error: "error" };
    }
    if (outcome.ok) {
      await this.#record(job, "sent", null);
      job.settle({ status: "sent", lastError: null });
      return;
    }
    const delay = this.#retryDelays[job.attempts - 1] ?? this.#retryDelays.at(-1) ?? 60_000;
    if (retryable.has(outcome.error) && job.attempts < job.maxAttempts && !this.#stopped) {
      await this.#record(job, "pending", outcome.error);
      const cancel = this.#timers.after(delay, () => {
        this.#waiting.delete(job);
        if (this.#stopped) return;
        this.#queue.push(job);
        this.#pump();
      });
      this.#waiting.set(job, cancel);
      return;
    }
    await this.#record(job, "failed", outcome.error);
    job.settle({ status: "failed", lastError: outcome.error });
  }

  async #send(destination: DestinationRow, job: Job, signal: AbortSignal): Promise<Outcome> {
    if (destination.kind === "email") {
      const { subject, text } = renderEmail(job.message, this.#origin);
      return sendEmail(
        this.#config.email,
        this.#mailSender,
        {
          recipients: destination.recipients,
          subject,
          text,
          ...(job.attachment ? { attachment: job.attachment } : {}),
        },
        signal,
      );
    }
    const secret = this.#secret(destination);
    if (!secret) return { ok: false, error: "secret_unreadable" };
    const now = this.#now();
    // A message too large for the destination drops listed items until it fits; the count and
    // the "and more" line still say how many there were.
    let body = "";
    for (const keep of [MAX_ALERT_ITEMS, 5, 2, 0]) {
      body = renderWebhook(
        destination.kind,
        { ...job.message, items: job.message.items.slice(0, keep) },
        this.#origin,
        { deliveryId: job.id, sentAt: now.toISOString() },
      );
      if (Buffer.byteLength(body) <= MAX_BODY_BYTES) break;
    }
    if (Buffer.byteLength(body) > MAX_BODY_BYTES) return { ok: false, error: "too_large" };
    const result = await sendWebhook({
      url: secret.url,
      body,
      headers: signedHeaders({ secret: secret.signingSecret, body, deliveryId: job.id, now }),
      allowedPorts: this.#config.webhookPorts,
      lookup: this.#lookup,
      connect: this.#connect,
      signal,
    });
    return result.ok ? { ok: true } : result;
  }
}

// Why a scheduled export's CSV is not attached, or null when it is: only email with details on
// gets the file, because it carries every column, and only up to 5 MB.
export function exportWithheld(
  destination: { kind: DestinationRow["kind"]; includeDetails: boolean },
  bytes: number,
): "webhook" | "details_off" | "too_large" | null {
  if (destination.kind !== "email") return "webhook";
  if (!destination.includeDetails) return "details_off";
  return bytes > MAX_EXPORT_ATTACHMENT_BYTES ? "too_large" : null;
}

function safeJson(value: string): unknown {
  try {
    return JSON.parse(value) as unknown;
  } catch {
    return null;
  }
}

// One attempt with a time limit. At the limit the attempt is aborted, which closes its connection,
// and its own answer to the abort is the outcome: email says whether the message may already have
// left (outcome_unknown, never tried again), a webhook reports a timeout. Nothing keeps running
// after an attempt ends.
async function bounded(
  work: (signal: AbortSignal) => Promise<Outcome>,
  milliseconds: number,
): Promise<Outcome> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), milliseconds);
  try {
    return await work(controller.signal);
  } finally {
    clearTimeout(timer);
    controller.abort();
  }
}

const severityRank: Record<FindingNotification["severity"], number> = {
  info: 0,
  low: 1,
  medium: 2,
  high: 3,
  critical: 4,
};
const rank = (notification: FindingNotification | undefined) =>
  notification ? severityRank[notification.severity] : -1;

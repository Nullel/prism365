import { randomUUID } from "node:crypto";
import { finishRun, insertRun } from "@prism365/audits";
import type { AlertRuleInput, DeliveryError } from "@prism365/contracts";
import {
  type CollectionStatus,
  type FindingNotification,
  upsertFinding,
  writeCollectionIntent,
} from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase } from "@prism365/storage";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { finding as findingInput } from "../../../packages/modules/test/fixtures.ts";
import type { ScheduleTimers } from "../src/audit-schedule.ts";
import type { HostLookup } from "../src/delivery-address.ts";
import { DeliveryService } from "../src/delivery.ts";
import type { DeliveryConfig, MailSender } from "../src/delivery-smtp.ts";
import type { WebhookConnector, WebhookTarget } from "../src/delivery-webhook.ts";
import { verifyWebhookSignature } from "../src/delivery-webhook.ts";

vi.setConfig({ testTimeout: 30_000 });

// The delivery service on PGlite with fake transports: rule evaluation per kind, quiet periods,
// retries, the ledger, and what is stored. Synthetic tenant, administrator, hosts and addresses.

const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "https://prism365.example";
const hookUrl = "https://hooks.example/workflows/synthetic?sig=secret-token-value";
const config: DeliveryConfig = {
  email: {
    smtp: {
      host: "relay.example",
      port: 465,
      implicitTls: true,
      user: "alerts",
      password: "synthetic-password",
      from: "alerts@example.invalid",
    },
    domains: ["example.invalid"],
  },
  webhookPorts: [],
};

interface Sent {
  target: WebhookTarget;
  headers: Record<string, string>;
  body: string;
}

let database: PGliteModuleDatabase;
let store: MemoryStateStore;
let clock: Date;
let statuses: number[];
let sent: Sent[];
let mail: Record<string, unknown>[];
let delays: number[];
let periodic: (() => void) | null;
let collection: CollectionStatus;
let service: DeliveryService;
let logged: string[];

const timers: ScheduleTimers = {
  every(_milliseconds, run) {
    periodic = run;
    return () => {
      periodic = null;
    };
  },
  after(milliseconds, run) {
    delays.push(milliseconds);
    const timer = setTimeout(run, 0);
    return () => clearTimeout(timer);
  },
};
const lookup: HostLookup = async () => [{ address: "100.128.0.1", family: 4 }];
const connect: WebhookConnector = async (target, request) => {
  sent.push({ target, headers: request.headers, body: request.body });
  return { status: statuses.shift() ?? 200 };
};
const mailSender: MailSender = {
  sendMail: async (message) => {
    mail.push(message);
    return {};
  },
};

function build(overrides: Partial<ConstructorParameters<typeof DeliveryService>[0]> = {}) {
  return new DeliveryService({
    tenantId,
    database,
    store,
    origin,
    config,
    encryptionKey: "a".repeat(64),
    collection: () => collection,
    lookup,
    connect,
    mailSender,
    now: () => clock,
    timers,
    log: (line) => logged.push(line),
    ...overrides,
  });
}

beforeEach(async () => {
  database = await PGliteModuleDatabase.open();
  store = new MemoryStateStore();
  clock = new Date("2026-10-24T12:00:00.000Z");
  statuses = [];
  sent = [];
  mail = [];
  delays = [];
  periodic = null;
  logged = [];
  collection = {
    state: "collecting",
    credentialKind: "delegated",
    boundAt: clock.toISOString(),
    lastStopReason: null,
    skipped: [],
  };
  service = build();
  service.start();
});

afterEach(async () => {
  await service.stop();
  await database.close();
});

const noop = async () => {};

async function webhook(kind: "teams" | "slack" | "generic" = "generic", includeDetails = false) {
  const created = await service.createDestination(
    actorId,
    { kind, label: "Synthetic hook", url: hookUrl, includeDetails },
    noop,
  );
  return created;
}

async function emailDestination(includeDetails = false) {
  return service.createDestination(
    actorId,
    { kind: "email", label: "Ops", recipients: ["ops@example.invalid"], includeDetails },
    noop,
  );
}

async function rule(input: AlertRuleInput) {
  const rules = await service.createRule(actorId, input as never, noop);
  return rules.at(-1) as NonNullable<(typeof rules)[number]>;
}

function notification(overrides: Partial<FindingNotification> = {}): FindingNotification {
  return {
    tenantId,
    module: "signins.audit",
    findingId: randomUUID(),
    ruleId: "signins.failed_burst_ip",
    severity: "high",
    title: "Repeated failed sign-ins from one address",
    summary: "40 failed sign-ins from 203.0.113.45 within 10 minutes.",
    subjectLabel: "203.0.113.45",
    change: "created",
    observedAt: clock.toISOString(),
    ...overrides,
  };
}

const rows = () =>
  database.tenant(tenantId, (sql) =>
    sql.query<Record<string, unknown>>("SELECT * FROM prism365_delivery ORDER BY created_at, id"),
  );
const ledger = () => store.audit.map((entry) => entry.event);

describe("destinations at rest", () => {
  it("keeps a webhook's address and signing secret only encrypted, and shows only the host", async () => {
    const created = await webhook();
    expect(created.signingSecret).toMatch(/^[A-Za-z0-9_-]{43}$/);
    expect(created.destination).toMatchObject({ host: "hooks.example", readable: true });
    expect(JSON.stringify(created.destination)).not.toContain("secret-token-value");
    const stored = await database.tenant(tenantId, (sql) =>
      sql.query<{ host: string; secret: Uint8Array; recipients: unknown }>(
        "SELECT host, secret, recipients FROM prism365_destination",
      ),
    );
    const raw = Buffer.from(stored.rows[0]?.secret ?? []).toString("latin1");
    expect(stored.rows[0]?.host).toBe("hooks.example");
    expect(raw).not.toContain("secret-token-value");
    expect(raw).not.toContain("hooks.example");
    expect(raw).not.toContain(created.signingSecret ?? "missing");
    const overview = await service.overview();
    expect(JSON.stringify(overview)).not.toContain("secret-token-value");
    expect(JSON.stringify(overview)).not.toContain(created.signingSecret ?? "missing");
  });

  it("cannot use a stored address under another key, and says so", async () => {
    const created = await webhook();
    await service.stop();
    service = build({ encryptionKey: "b".repeat(64) });
    service.start();
    const overview = await service.overview();
    expect(overview.destinations[0]).toMatchObject({ id: created.destination.id, readable: false });
    const record = await service.sendTest(actorId, created.destination.id);
    expect(record).toMatchObject({ status: "failed", lastError: "secret_unreadable", attempts: 1 });
    expect(sent).toEqual([]);
  });

  it("binds the encrypted address to its destination, so rows cannot be swapped", async () => {
    const first = await webhook();
    const second = await service.createDestination(
      actorId,
      { kind: "generic", label: "Other", url: "https://other.example/hook", includeDetails: false },
      noop,
    );
    // Copy the first destination's secret onto the second, as someone with database access might.
    await database.tenant(tenantId, (sql) =>
      sql.query(
        "UPDATE prism365_destination SET secret = (SELECT secret FROM prism365_destination WHERE id=$1) WHERE id=$2",
        [first.destination.id, second.destination.id],
      ),
    );
    expect(await service.sendTest(actorId, second.destination.id)).toMatchObject({
      status: "failed",
      lastError: "secret_unreadable",
    });
    expect(sent).toEqual([]);
  });

  it("refuses email recipients outside the allowed domains, and email without a relay", async () => {
    await expect(
      service.createDestination(
        actorId,
        {
          kind: "email",
          label: "Ops",
          recipients: ["ops@sub.example.invalid"],
          includeDetails: false,
        },
        noop,
      ),
    ).rejects.toMatchObject({ code: "recipient_not_allowed" });
    await service.stop();
    service = build({ config: { email: { smtp: null, domains: [] }, webhookPorts: [] } });
    await expect(emailDestination()).rejects.toMatchObject({ code: "email_not_configured" });
  });
});

describe("finding rules", () => {
  it("sends one signed alert per batch and records it in the log and the ledger", async () => {
    const hook = await webhook();
    const watching = await rule({
      kind: "finding",
      name: "High findings",
      destinations: [hook.destination.id],
      quietMinutes: 60,
      minimumSeverity: "high",
      modules: [],
    });
    const deliver = service.sink.deliver;
    await deliver(notification());
    await deliver(notification({ severity: "critical" }));
    await deliver(notification({ severity: "low" }));
    await service.idle();
    expect(sent).toHaveLength(1);
    const body = JSON.parse(sent[0]?.body ?? "{}");
    expect(body).toMatchObject({ version: 1, kind: "finding", count: 2, severity: "critical" });
    expect(sent[0]?.target).toMatchObject({ address: "100.128.0.1", host: "hooks.example" });
    expect(
      verifyWebhookSignature({
        secret: hook.signingSecret ?? "",
        timestamp: sent[0]?.headers["x-prism365-timestamp"] ?? "",
        body: sent[0]?.body ?? "",
        signature: sent[0]?.headers["x-prism365-signature"] ?? "",
        now: clock,
      }),
    ).toBe(true);
    const logged = (await rows()).rows;
    expect(logged).toHaveLength(1);
    expect(logged[0]).toMatchObject({
      source: "rule",
      rule_id: watching.id,
      destination_id: hook.destination.id,
      status: "sent",
      attempts: 1,
      items: 2,
      last_error: null,
    });
    // The log row holds no part of the message.
    expect(JSON.stringify(logged[0])).not.toContain("failed sign-ins");
    expect(JSON.stringify(logged[0])).not.toContain("203.0.113");
    const entry = ledger().find((event) => event.type === "alerts.delivery");
    expect(entry).toEqual({
      type: "alerts.delivery",
      source: "rule",
      ruleId: watching.id,
      destinationId: hook.destination.id,
      deliveryId: logged[0]?.id,
      items: 2,
      occurredAt: clock.toISOString(),
    });
    expect(store.audit.find((item) => item.event.type === "alerts.delivery")?.actorId).toBe(
      actorId,
    );
  });

  it("sends nothing for a subject inside the quiet period, and again after it", async () => {
    const hook = await webhook();
    await rule({
      kind: "finding",
      name: "Quiet",
      destinations: [hook.destination.id],
      quietMinutes: 60,
      minimumSeverity: "medium",
      modules: [],
    });
    const repeated = notification();
    await service.sink.deliver(repeated);
    await service.idle();
    clock = new Date(clock.getTime() + 30 * 60_000);
    await service.sink.deliver({ ...repeated, change: "escalated", severity: "critical" });
    await service.idle();
    expect(sent).toHaveLength(1);
    // Another subject in the same period still goes out.
    await service.sink.deliver(notification());
    await service.idle();
    expect(sent).toHaveLength(2);
    clock = new Date(clock.getTime() + 31 * 60_000);
    await service.sink.deliver({ ...repeated, change: "reopened" });
    await service.idle();
    expect(sent).toHaveLength(3);
  });

  it("counts a failed delivery towards the quiet period", async () => {
    const hook = await webhook();
    await rule({
      kind: "finding",
      name: "Quiet",
      destinations: [hook.destination.id],
      quietMinutes: 60,
      minimumSeverity: "medium",
      modules: [],
    });
    statuses = [400];
    const repeated = notification();
    await service.sink.deliver(repeated);
    await service.idle();
    await service.sink.deliver({ ...repeated, change: "escalated" });
    await service.idle();
    expect(sent).toHaveLength(1);
  });

  it("matches directory changes by rule id and gaps by their rules", async () => {
    const hook = await webhook();
    await rule({
      kind: "directory",
      name: "Roles",
      destinations: [hook.destination.id],
      quietMinutes: 60,
      ruleIds: ["role.privileged-assigned"],
    });
    await rule({
      kind: "collection",
      name: "Gaps",
      destinations: [hook.destination.id],
      quietMinutes: 60,
      stoppedHours: null,
      gaps: true,
    });
    await service.sink.deliver(
      notification({ module: "directory.changes", ruleId: "app.consent-granted" }),
    );
    await service.idle();
    expect(sent).toHaveLength(0);
    await service.sink.deliver(
      notification({ module: "directory.changes", ruleId: "role.privileged-assigned" }),
    );
    await service.sink.deliver(
      notification({ ruleId: "signins.collection_gap", severity: "medium" }),
    );
    await service.idle();
    expect(sent.map((item) => JSON.parse(item.body).kind).sort()).toEqual([
      "collection",
      "directory",
    ]);
  });

  it("adds evidence only for a destination that asked, and cuts addresses for webhooks", async () => {
    const plain = await webhook("generic", false);
    const detailed = await webhook("generic", true);
    const email = await emailDestination(true);
    await rule({
      kind: "finding",
      name: "Details",
      destinations: [plain.destination.id, detailed.destination.id, email.destination.id],
      quietMinutes: 60,
      minimumSeverity: "medium",
      modules: [],
    });
    const stored = await database.tenant(tenantId, (sql) =>
      upsertFinding(
        sql,
        tenantId,
        "signins.audit",
        findingInput({
          severity: "high",
          subject: { kind: "resource", id: "203.0.113.45", label: "203.0.113.45" },
          evidence: [
            {
              source: "Synthetic fixture",
              ref: "event-1",
              observedAt: clock.toISOString(),
              fields: { ip: "203.0.113.45", country: "XA" },
            },
          ],
        }),
      ),
    );
    await service.sink.deliver(notification({ findingId: stored.id }));
    await service.idle();
    expect(sent).toHaveLength(2);
    const bodies = sent.map((item) => JSON.parse(item.body));
    const withDetails = bodies.find((body) => Object.keys(body.items[0].details).length > 0);
    const withoutDetails = bodies.find((body) => Object.keys(body.items[0].details).length === 0);
    expect(withDetails.items[0].details).toMatchObject({ ip: "203.0.113.0/24", country: "XA" });
    expect(withoutDetails.items[0].subject).toBe("203.0.113.0/24");
    for (const item of sent) expect(item.body).not.toContain("203.0.113.45");
    expect(mail).toHaveLength(1);
    expect(String(mail[0]?.text)).toContain("ip: 203.0.113.45");
    expect(mail[0]).toMatchObject({ to: [{ name: "", address: "ops@example.invalid" }] });
  });

  it("returns at once from the sink even when a destination never answers", async () => {
    const hook = await webhook();
    await rule({
      kind: "finding",
      name: "Stuck",
      destinations: [hook.destination.id],
      quietMinutes: 60,
      minimumSeverity: "info",
      modules: [],
    });
    await service.stop();
    service = build({
      connect: () => new Promise<never>(() => {}),
      attemptTimeoutMilliseconds: 200,
    });
    service.start();
    const started = Date.now();
    for (let index = 0; index < 50; index += 1) await service.sink.deliver(notification());
    expect(Date.now() - started).toBeLessThan(100);
  });
});

describe("retries", () => {
  async function hookRule() {
    const hook = await webhook();
    await rule({
      kind: "finding",
      name: "Retries",
      destinations: [hook.destination.id],
      quietMinutes: 60,
      minimumSeverity: "info",
      modules: [],
    });
    return hook;
  }
  const lastRow = async () => (await rows()).rows.at(-1);

  it("tries again with backoff and succeeds on the third attempt", async () => {
    await hookRule();
    statuses = [503, 503, 200];
    await service.sink.deliver(notification());
    await service.idle();
    expect(sent).toHaveLength(3);
    expect(await lastRow()).toMatchObject({ status: "sent", attempts: 3, last_error: null });
    // The flush delay, then the two backoffs.
    expect(delays.slice(-2)).toEqual([30_000, 120_000]);
  });

  it("marks a delivery failed after three attempts and shows why", async () => {
    await hookRule();
    statuses = [503, 503, 503, 200];
    await service.sink.deliver(notification());
    await service.idle();
    expect(sent).toHaveLength(3);
    expect(await lastRow()).toMatchObject({
      status: "failed",
      attempts: 3,
      last_error: "server_error",
    });
    const overview = await service.overview();
    expect(overview.deliveries[0]).toMatchObject({
      status: "failed",
      attempts: 3,
      lastError: "server_error",
    });
  });

  it("does not retry what cannot succeed later", async () => {
    await hookRule();
    for (const [status, error] of [
      [400, "rejected"],
      [302, "redirect"],
    ] as [number, DeliveryError][]) {
      sent = [];
      statuses = [status, 200];
      // Each delivery a minute after the last, so the newest row is the one just made.
      clock = new Date(clock.getTime() + 60_000);
      await service.sink.deliver(notification());
      await service.idle();
      expect(sent, error).toHaveLength(1);
      expect(await lastRow()).toMatchObject({ status: "failed", attempts: 1, last_error: error });
    }
  });

  it("refuses a blocked address without connecting or retrying", async () => {
    await service.stop();
    service = build({ lookup: async () => [{ address: "169.254.169.254", family: 4 }] });
    service.start();
    await hookRule();
    await service.sink.deliver(notification());
    await service.idle();
    expect(sent).toEqual([]);
    expect(await lastRow()).toMatchObject({
      status: "failed",
      attempts: 1,
      last_error: "blocked_address",
    });
  });

  it("stops retrying a destination deleted in between", async () => {
    const hook = await hookRule();
    statuses = [503];
    let release: () => void = () => {};
    const held: (() => void)[] = [];
    await service.stop();
    service = build({
      timers: {
        every: timers.every,
        after(milliseconds, run) {
          delays.push(milliseconds);
          if (milliseconds >= 30_000) {
            held.push(run);
            return () => {};
          }
          const timer = setTimeout(run, 0);
          return () => clearTimeout(timer);
        },
      },
    });
    service.start();
    await service.sink.deliver(notification());
    for (let wait = 0; wait < 200 && held.length === 0; wait += 1)
      await new Promise((resolve) => setTimeout(resolve, 5));
    await database.tenant(tenantId, (sql) =>
      sql.query("DELETE FROM prism365_preference WHERE key='alerts.rules'"),
    );
    await service.deleteDestination(hook.destination.id, noop);
    release = held[0] ?? release;
    release();
    await service.idle();
    expect(sent).toHaveLength(1);
    expect(await lastRow()).toMatchObject({ status: "failed", last_error: "destination_deleted" });
  });

  it("marks what a restart interrupted", async () => {
    await database.tenant(tenantId, (sql) =>
      sql.query(
        `INSERT INTO prism365_delivery (tenant_id, id, destination_id, source, rule_id, subjects,
           items, status, attempts, created_at, updated_at)
         VALUES ($1,$2,$3,'test',NULL,'[]'::jsonb,0,'pending',1,$4,$4)`,
        [tenantId, randomUUID(), randomUUID(), "2026-10-24T11:00:00.000Z"],
      ),
    );
    await service.stop();
    service = build();
    service.start();
    await service.idle();
    expect(await lastRow()).toMatchObject({ status: "failed", last_error: "interrupted" });
  });

  it("sends nothing when the ledger cannot be written", async () => {
    const hook = await webhook();
    store.appendAudit = async () => {
      throw new Error("synthetic ledger failure");
    };
    const record = await service.sendTest(actorId, hook.destination.id);
    expect(record).toMatchObject({ status: "failed", lastError: "ledger_failed", attempts: 0 });
    expect(sent).toEqual([]);
  });
});

describe("audit rules", () => {
  const auditId = "accounts.inactive";
  async function storeRun(value: number, options: { complete?: boolean; version?: number } = {}) {
    const runId = randomUUID();
    await database.tenant(tenantId, async (sql) => {
      await insertRun(sql, tenantId, {
        id: runId,
        auditId,
        auditVersion: options.version ?? 1,
        startedAt: clock,
        startedBy: { id: actorId, name: "Synthetic admin" },
        trigger: "schedule",
      });
      await finishRun(sql, tenantId, runId, clock, {
        status: "succeeded",
        complete: options.complete ?? true,
        summary: [{ label: "Inactive accounts", value, tone: "warning", view: null }],
        notes: [],
        rows: [],
      });
    });
    service.onAuditStored({ tenantId, auditId, runId, trigger: "schedule" });
    await service.idle();
    clock = new Date(clock.getTime() + 86_400_000);
  }

  it("fires when a tile crosses the threshold against its previous comparable point", async () => {
    const hook = await webhook();
    await rule({
      kind: "audit",
      name: "Inactive over 10",
      destinations: [hook.destination.id],
      quietMinutes: 5,
      auditId,
      tileId: "inactive-accounts",
      condition: "above",
      value: 10,
    });
    await storeRun(4);
    expect(sent).toHaveLength(0);
    // Partial runs and another audit version are not compared.
    await storeRun(30, { complete: false });
    await storeRun(30, { version: 2 });
    expect(sent).toHaveLength(1);
    expect(JSON.parse(sent[0]?.body ?? "{}").lines).toContain(
      "There is no earlier comparable run to compare with.",
    );
    await storeRun(12);
    expect(sent).toHaveLength(2);
    expect(JSON.parse(sent[1]?.body ?? "{}")).toMatchObject({
      kind: "audit",
      title: "Inactive accounts: Inactive accounts is 12",
      link: "https://prism365.example/#/audits/accounts.inactive",
    });
    // Staying above is not a new crossing.
    await storeRun(15);
    expect(sent).toHaveLength(2);
  });

  it("fires a change rule only on a large enough move", async () => {
    const hook = await webhook();
    await rule({
      kind: "audit",
      name: "Moves by 5",
      destinations: [hook.destination.id],
      quietMinutes: 5,
      auditId,
      tileId: "inactive-accounts",
      condition: "change",
      value: 5,
    });
    await storeRun(10);
    await storeRun(13);
    expect(sent).toHaveLength(0);
    await storeRun(19);
    expect(sent).toHaveLength(1);
  });
});

describe("collection rules", () => {
  it("alerts once when collection meant to be on stays stopped, then waits out the quiet period", async () => {
    const hook = await webhook();
    await rule({
      kind: "collection",
      name: "Stopped",
      destinations: [hook.destination.id],
      quietMinutes: 24 * 60,
      stoppedHours: 6,
      gaps: false,
    });
    await database.tenant(tenantId, (sql) =>
      writeCollectionIntent(sql, tenantId, { actorId, since: clock.toISOString() }),
    );
    collection = { ...collection, state: "stopped", lastStopReason: "signed_out" };
    service.onSchedulerEvent({
      type: "collection.unbound",
      tenantId,
      actorId,
      reason: "signed_out",
    });
    const tick = async () => {
      periodic?.();
      await service.idle();
    };
    clock = new Date(clock.getTime() + 5 * 3_600_000);
    await tick();
    expect(sent).toHaveLength(0);
    clock = new Date(clock.getTime() + 2 * 3_600_000);
    await tick();
    expect(sent).toHaveLength(1);
    expect(JSON.parse(sent[0]?.body ?? "{}").lines[1]).toBe(
      "The administrator it collects for signed out.",
    );
    clock = new Date(clock.getTime() + 3_600_000);
    await tick();
    expect(sent).toHaveLength(1);
    clock = new Date(clock.getTime() + 24 * 3_600_000);
    await tick();
    expect(sent).toHaveLength(2);
  });

  it("stays quiet when collection was stopped on purpose or is running", async () => {
    const hook = await webhook();
    await rule({
      kind: "collection",
      name: "Stopped",
      destinations: [hook.destination.id],
      quietMinutes: 60,
      stoppedHours: 1,
      gaps: false,
    });
    collection = { ...collection, state: "stopped", lastStopReason: "operator_stopped" };
    service.onSchedulerEvent({
      type: "collection.unbound",
      tenantId,
      actorId,
      reason: "operator_stopped",
    });
    clock = new Date(clock.getTime() + 5 * 3_600_000);
    periodic?.();
    await service.idle();
    expect(sent).toHaveLength(0);
  });
});

describe("scheduled exports", () => {
  const auditId = "accounts.inactive";
  async function scheduledRun(trigger: "schedule" | "manual" = "schedule") {
    const runId = randomUUID();
    await database.tenant(tenantId, async (sql) => {
      await insertRun(sql, tenantId, {
        id: runId,
        auditId,
        auditVersion: 1,
        startedAt: clock,
        startedBy: { id: actorId, name: "Synthetic admin" },
        scope: ["example.invalid"],
        trigger,
      });
      await finishRun(sql, tenantId, runId, clock, {
        status: "succeeded",
        complete: true,
        summary: [{ label: "Inactive accounts", value: 1, tone: "warning", view: null }],
        notes: [],
        rows: [
          {
            key: "user-3",
            personId: "00000003-aaaa-4aaa-8aaa-000000000003",
            tags: ["d30"],
            cells: {
              person: { name: "Synthetic Person 3", principal: "person3@example.invalid" },
              days: 40,
            },
          },
        ],
      });
    });
    service.onAuditStored({ tenantId, auditId, runId, trigger });
    await service.idle();
    return runId;
  }

  it("emails the run's CSV, named with its scope, to email with details on", async () => {
    const email = await emailDestination(true);
    await service.setExport(actorId, auditId, email.destination.id, noop);
    const runId = await scheduledRun();
    expect(mail).toHaveLength(1);
    const attachments = mail[0]?.attachments as { filename: string; content: Buffer }[];
    expect(attachments[0]?.filename).toBe(
      "prism365-accounts.inactive-2026-10-24-limited-to-1-domain.csv",
    );
    expect(attachments[0]?.content.toString()).toContain("person3@example.invalid");
    expect(String(mail[0]?.text)).toContain("The CSV is attached.");
    expect(ledger().find((event) => event.source === "export")).toMatchObject({
      type: "alerts.delivery",
      auditId,
      runId,
      items: 1,
    });
  });

  it("sends webhooks and email without details a summary and a link, never the file", async () => {
    const hook = await webhook();
    await service.setExport(actorId, auditId, hook.destination.id, noop);
    await scheduledRun();
    expect(sent).toHaveLength(1);
    expect(sent[0]?.body).not.toContain("person3@example.invalid");
    expect(JSON.parse(sent[0]?.body ?? "{}")).toMatchObject({
      kind: "export",
      count: 1,
      link: "https://prism365.example/#/audits/accounts.inactive",
    });
    const email = await emailDestination(false);
    await service.setExport(actorId, auditId, email.destination.id, noop);
    await scheduledRun();
    expect(mail).toHaveLength(1);
    expect(mail[0]).not.toHaveProperty("attachments");
    expect(String(mail[0]?.text)).toContain("does not include details");
  });

  it("sends nothing for a run an administrator started", async () => {
    const email = await emailDestination(true);
    await service.setExport(actorId, auditId, email.destination.id, noop);
    await scheduledRun("manual");
    expect(mail).toEqual([]);
  });

  it("keeps a destination an export uses", async () => {
    const email = await emailDestination(true);
    await service.setExport(actorId, auditId, email.destination.id, noop);
    await expect(service.deleteDestination(email.destination.id, noop)).rejects.toMatchObject({
      code: "destination_in_use",
    });
    await expect(service.setExport(actorId, "audit.unknown", null, noop)).rejects.toMatchObject({
      code: "unknown_audit",
    });
  });
});

describe("the weekly summary", () => {
  const digest = {
    signInFindings: {
      state: "ok",
      opened: { critical: 0, high: 1, medium: 0, low: 0, info: 0 },
      openTotal: 1,
      resolved: 0,
    },
    directoryChanges: { state: "ok", total: 3, flaggedTotal: 0 },
    quietAccounts: { state: "ok", total: 0 },
    auditDrift: { compared: 0, audits: [] },
    microsoft: { state: "ok", posts: { total: 0 }, incidents: { opened: 0 } },
    collection: { collecting: true, failedRuns: 0, gaps: 0 },
  };

  it("goes out once from Monday 07:00, and not for a week the server missed", async () => {
    await service.stop();
    let builds = 0;
    service = build({
      timeZone: "UTC",
      weeklyDigest: {
        build: async () => {
          builds += 1;
          return digest as never;
        },
        datasets: () => ["signInFindings", "directoryChanges"],
      },
    });
    service.start();
    const hook = await webhook();
    const tick = async (at: string) => {
      clock = new Date(at);
      periodic?.();
      await service.idle();
    };
    await tick("2026-10-26T06:55:00.000Z");
    expect(sent).toHaveLength(0);
    await service.setWeekly(actorId, hook.destination.id, noop);
    await tick("2026-10-26T07:05:00.000Z");
    await tick("2026-10-26T07:10:00.000Z");
    expect(sent).toHaveLength(1);
    expect(JSON.parse(sent[0]?.body ?? "{}")).toMatchObject({
      kind: "summary",
      title: "This week in Prism365",
      link: "https://prism365.example/#/week",
    });
    expect(ledger().find((event) => event.source === "summary")).toMatchObject({
      datasets: ["signInFindings", "directoryChanges"],
    });
    // The server was down all of the next Monday: that week is skipped, not sent on Wednesday.
    await tick("2026-11-04T09:00:00.000Z");
    expect(sent).toHaveLength(1);
    await tick("2026-11-09T07:00:00.000Z");
    expect(sent).toHaveLength(2);
    expect(builds).toBe(2);
    await service.setWeekly(actorId, null, noop);
    await tick("2026-11-16T07:00:00.000Z");
    expect(sent).toHaveLength(2);
  });
});

describe("a burst of findings", () => {
  const burst = async (severity: "info" | "low", count: number) => {
    for (let index = 0; index < count; index += 1)
      await service.sink.deliver(notification({ severity, ruleId: "signins.burst" }));
  };
  const canary = () =>
    notification({ severity: "critical", ruleId: "signin.canary-used", title: "Canary" });

  it("never pushes out a critical finding with findings no rule wants", async () => {
    const hook = await webhook();
    await rule({
      kind: "finding",
      name: "Critical only",
      destinations: [hook.destination.id],
      quietMinutes: 5,
      minimumSeverity: "critical",
      modules: [],
    });
    await burst("low", 1_000);
    await service.sink.deliver(canary());
    await service.idle();
    expect(sent).toHaveLength(1);
    expect(JSON.parse(sent[0]?.body ?? "{}")).toMatchObject({ severity: "critical", count: 1 });
    expect((await rows()).rows.map((row) => row.last_error)).toEqual([null]);
  });

  it("takes in only what a rule wants, so unwanted findings never crowd out a wanted one", async () => {
    const hook = await webhook();
    await rule({
      kind: "directory",
      name: "Roles",
      destinations: [hook.destination.id],
      quietMinutes: 5,
      ruleIds: ["role.privileged-assigned"],
    });
    // Medium sign-in findings no rule wants, then a low directory change one rule does.
    for (let index = 0; index < 1_000; index += 1)
      await service.sink.deliver(notification({ severity: "medium" }));
    await service.sink.deliver(
      notification({
        module: "directory.changes",
        ruleId: "role.privileged-assigned",
        severity: "low",
      }),
    );
    await service.idle();
    expect(sent).toHaveLength(1);
    expect(JSON.parse(sent[0]?.body ?? "{}")).toMatchObject({ kind: "directory", count: 1 });
  });

  it("gives up the lowest, oldest finding for a worse one, and records what it left out", async () => {
    const hook = await webhook();
    const wanting = await rule({
      kind: "finding",
      name: "Low and up",
      destinations: [hook.destination.id],
      quietMinutes: 5,
      minimumSeverity: "low",
      modules: [],
    });
    await burst("low", 1_000);
    await service.sink.deliver(canary());
    await service.idle();
    const body = JSON.parse(sent[0]?.body ?? "{}");
    expect(body).toMatchObject({ severity: "critical", count: 1_000 });
    expect(body.items[0]).toMatchObject({ severity: "critical" });
    const logged = (await rows()).rows;
    expect(logged.find((row) => row.last_error === "dropped")).toMatchObject({
      status: "failed",
      items: 1,
      rule_id: wanting.id,
      destination_id: hook.destination.id,
    });
    expect(ledger().find((event) => event.type === "alerts.dropped")).toMatchObject({
      ruleId: wanting.id,
      items: 1,
    });
    const overview = await service.overview();
    expect(overview.deliveries.some((item) => item.lastError === "dropped")).toBe(true);
  });

  it("keeps the newest of equals and leaves out the oldest", async () => {
    const hook = await webhook();
    await rule({
      kind: "finding",
      name: "Low and up",
      destinations: [hook.destination.id],
      quietMinutes: 5,
      minimumSeverity: "low",
      modules: [],
    });
    const first = notification({ severity: "low", title: "First of the burst" });
    await service.sink.deliver(first);
    await burst("low", 999);
    const last = notification({ severity: "low", title: "Last of the burst" });
    await service.sink.deliver(last);
    await service.idle();
    expect(sent[0]?.body).not.toContain("First of the burst");
    expect(JSON.parse(sent[0]?.body ?? "{}").count).toBe(1_000);
  });
});

describe("one destination at a time", () => {
  it("lets a hanging destination hold one worker, not delay the others", async () => {
    await service.stop();
    service = build({
      connect: async (target, request) => {
        if (target.host === "stuck.example") return new Promise<never>(() => {});
        sent.push({ target, headers: request.headers, body: request.body });
        return { status: 200 };
      },
      attemptTimeoutMilliseconds: 1_000,
    });
    service.start();
    const stuck = await service.createDestination(
      actorId,
      { kind: "generic", label: "Stuck", url: "https://stuck.example/hook", includeDetails: false },
      noop,
    );
    const healthy = await service.createDestination(
      actorId,
      {
        kind: "generic",
        label: "Healthy",
        url: "https://healthy.example/hook",
        includeDetails: false,
      },
      noop,
    );
    const waiting = Array.from({ length: 3 }, () =>
      service.sendTest(actorId, stuck.destination.id),
    );
    await new Promise((resolve) => setTimeout(resolve, 50));
    const started = Date.now();
    const record = await service.sendTest(actorId, healthy.destination.id);
    expect(record).toMatchObject({ status: "sent" });
    expect(Date.now() - started).toBeLessThan(700);
    // The stuck tests end one after another, each at its own time limit.
    expect((await Promise.all(waiting)).map((item) => item?.lastError)).toEqual([
      "timeout",
      "timeout",
      "timeout",
    ]);
  });
});

describe("a send that times out", () => {
  it("closes the attempt, and never resends mail the relay may already have", async () => {
    const signals: AbortSignal[] = [];
    await service.stop();
    service = build({
      mailSender: {
        sendMail: (_message, control) => {
          if (control) signals.push(control.signal);
          control?.dataStarted();
          return new Promise(() => {});
        },
      },
      attemptTimeoutMilliseconds: 100,
    });
    service.start();
    const email = await emailDestination();
    await rule({
      kind: "finding",
      name: "Mail",
      destinations: [email.destination.id],
      quietMinutes: 5,
      minimumSeverity: "info",
      modules: [],
    });
    await service.sink.deliver(notification());
    await service.idle();
    expect(signals).toHaveLength(1);
    expect(signals[0]?.aborted).toBe(true);
    expect((await rows()).rows.at(-1)).toMatchObject({
      status: "failed",
      attempts: 1,
      last_error: "outcome_unknown",
    });
  });

  it("tries again only once the earlier attempt is closed, when nothing was handed over", async () => {
    const signals: AbortSignal[] = [];
    const openAtStart: number[] = [];
    await service.stop();
    service = build({
      mailSender: {
        sendMail: (_message, control) => {
          openAtStart.push(signals.filter((signal) => !signal.aborted).length);
          if (control) signals.push(control.signal);
          return new Promise(() => {});
        },
      },
      attemptTimeoutMilliseconds: 50,
    });
    service.start();
    const email = await emailDestination();
    await rule({
      kind: "finding",
      name: "Mail",
      destinations: [email.destination.id],
      quietMinutes: 5,
      minimumSeverity: "info",
      modules: [],
    });
    await service.sink.deliver(notification());
    await service.idle();
    expect(signals).toHaveLength(3);
    expect(openAtStart).toEqual([0, 0, 0]);
    expect((await rows()).rows.at(-1)).toMatchObject({
      status: "failed",
      attempts: 3,
      last_error: "timeout",
    });
  });
});

describe("settings this version cannot read", () => {
  const preference = (key: string) =>
    database.tenant(tenantId, async (sql) => {
      const result = await sql.query<{ value: unknown }>(
        "SELECT value FROM prism365_preference WHERE key=$1",
        [key],
      );
      const value = result.rows[0]?.value;
      return (typeof value === "string" ? JSON.parse(value) : value) as Record<string, unknown[]>;
    });
  const write = (key: string, value: unknown) =>
    database.tenant(tenantId, (sql) =>
      sql.query(
        `INSERT INTO prism365_preference (tenant_id, key, value, updated_at)
         VALUES ($1,$2,$3::jsonb,now())`,
        [tenantId, key, JSON.stringify(value)],
      ),
    );
  const future = { id: "rule-from-a-newer-version", kind: "anomaly", sensitivity: 3 };

  it("keeps a rule it cannot parse through every save, and counts it", async () => {
    const hook = await webhook();
    await write("alerts.rules", { rules: [future] });
    const input = (name: string) => ({
      kind: "finding" as const,
      name,
      destinations: [hook.destination.id],
      quietMinutes: 5,
      minimumSeverity: "high" as const,
      modules: [],
    });
    const created = await rule(input("Kept beside it"));
    await service.changeRule(actorId, created.id, { enabled: false }, noop);
    expect((await preference("alerts.rules")).rules).toContainEqual(future);
    for (let index = 0; index < 48; index += 1) await rule(input(`Rule ${index}`));
    await expect(rule(input("One too many"))).rejects.toMatchObject({ code: "too_many_rules" });
    await service.changeRule(actorId, created.id, "delete", noop);
    expect((await preference("alerts.rules")).rules).toContainEqual(future);
  });

  it("keeps an export entry it cannot parse", async () => {
    const email = await emailDestination(true);
    const odd = { auditId: "accounts.inactive", destinationId: "not-a-uuid", setBy: actorId };
    await write("alerts.exports", { exports: [odd] });
    await service.setExport(actorId, "guests.review", email.destination.id, noop);
    await service.setExport(actorId, "guests.review", null, noop);
    expect((await preference("alerts.exports")).exports).toEqual([odd]);
  });
});

describe("recipients this server refuses now", () => {
  it("marks them and sends nothing to them", async () => {
    const email = await emailDestination();
    // Saved by an earlier, wider rule: a routing address in an allowed domain.
    const routed = randomUUID();
    await database.tenant(tenantId, (sql) =>
      sql.query(
        `INSERT INTO prism365_destination (tenant_id, id, kind, label, host, recipients, secret,
           key_id, include_details, created_at, created_by)
         VALUES ($1,$2,'email','Routed',NULL,$3::jsonb,NULL,NULL,false,now(),$4)`,
        [tenantId, routed, JSON.stringify(["ops%attacker.example@example.invalid"]), actorId],
      ),
    );
    let overview = await service.overview();
    expect(overview.destinations.find((item) => item.id === routed)?.recipientsAllowed).toBe(false);
    expect(
      overview.destinations.find((item) => item.id === email.destination.id)?.recipientsAllowed,
    ).toBe(true);
    expect(await service.sendTest(actorId, routed)).toMatchObject({
      status: "failed",
      lastError: "recipient_not_allowed",
    });
    // A domain taken off ALERT_EMAIL_DOMAINS refuses what was saved under it.
    await service.stop();
    service = build({
      config: { ...config, email: { ...config.email, domains: ["other.invalid"] } },
    });
    service.start();
    overview = await service.overview();
    expect(
      overview.destinations.find((item) => item.id === email.destination.id)?.recipientsAllowed,
    ).toBe(false);
    expect(mail).toEqual([]);
  });
});

describe("a rule from a newer version", () => {
  const stored = () =>
    database.tenant(tenantId, async (sql) => {
      const result = await sql.query<{ value: unknown }>(
        "SELECT value FROM prism365_preference WHERE key='alerts.rules'",
      );
      const value = result.rows[0]?.value;
      return (typeof value === "string" ? JSON.parse(value) : value) as {
        rules: Record<string, unknown>[];
      };
    });

  it("is never run, and keeps every field through any save", async () => {
    const hook = await webhook();
    const newer = {
      id: "44444444-4444-4444-8444-444444444444",
      kind: "finding",
      name: "Newer rule",
      enabled: true,
      destinations: [hook.destination.id],
      quietMinutes: 60,
      minimumSeverity: "high",
      modules: [],
      createdAt: "2026-10-24T00:00:00.000Z",
      createdBy: actorId,
      // A field a later version added, which narrows what the rule means.
      onlyWhenUnresolvedFor: 120,
    };
    await database.tenant(tenantId, (sql) =>
      sql.query(
        `INSERT INTO prism365_preference (tenant_id, key, value, updated_at)
         VALUES ($1,'alerts.rules',$2::jsonb,now())`,
        [tenantId, JSON.stringify({ rules: [newer], futureSetting: true })],
      ),
    );
    await service.stop();
    service = build();
    service.start();
    await service.idle();
    // A finding it would match without its extra field goes nowhere.
    await service.sink.deliver(notification({ severity: "critical" }));
    await service.idle();
    expect(sent).toEqual([]);
    const overview = await service.overview();
    expect(overview.rules).toEqual([]);
    expect(overview.otherRules).toEqual([{ id: newer.id, name: "Newer rule", newer: true }]);
    // A save about another rule leaves it, and the setting's other fields, exactly as stored.
    const other = await rule({
      kind: "finding",
      name: "This version's rule",
      destinations: [hook.destination.id],
      quietMinutes: 60,
      minimumSeverity: "low",
      modules: [],
    });
    await service.changeRule(actorId, other.id, { enabled: false }, noop);
    let value = await stored();
    expect(value.rules[0]).toEqual(newer);
    expect(value).toMatchObject({ futureSetting: true });
    // Turning it off changes that field alone; deleting it removes it.
    await service.changeRule(actorId, newer.id, { enabled: false }, noop);
    value = await stored();
    expect(value.rules[0]).toEqual({ ...newer, enabled: false });
    await service.changeRule(actorId, newer.id, "delete", noop);
    value = await stored();
    expect(value.rules.map((item) => item.id)).toEqual([other.id]);
  });

  it("keeps an export entry with a field it does not know, and does not use it", async () => {
    const email = await emailDestination(true);
    const newer = {
      auditId: "accounts.inactive",
      destinationId: email.destination.id,
      setBy: actorId,
      format: "xlsx",
    };
    await database.tenant(tenantId, (sql) =>
      sql.query(
        `INSERT INTO prism365_preference (tenant_id, key, value, updated_at)
         VALUES ($1,'alerts.exports',$2::jsonb,now())`,
        [tenantId, JSON.stringify({ exports: [newer] })],
      ),
    );
    expect((await service.overview()).exports).toEqual([]);
    await service.setExport(actorId, "guests.review", email.destination.id, noop);
    const value = await database.tenant(tenantId, async (sql) => {
      const result = await sql.query<{ value: unknown }>(
        "SELECT value FROM prism365_preference WHERE key='alerts.exports'",
      );
      const raw = result.rows[0]?.value;
      return (typeof raw === "string" ? JSON.parse(raw) : raw) as { exports: unknown[] };
    });
    expect(value.exports).toContainEqual(newer);
  });
});

describe("the rule cache", () => {
  it("is never overwritten by a read that began before a save", async () => {
    await service.stop();
    let holdNext: Promise<void> | null = null;
    // A read that has finished in the database but whose answer arrives late.
    const slow = {
      kind: database.kind,
      durable: database.durable,
      verify: () => database.verify(),
      close: async () => {},
      async tenant<T>(
        tenant: string,
        operation: Parameters<typeof database.tenant<T>>[1],
        options?: Parameters<typeof database.tenant<T>>[2],
      ): Promise<T> {
        const result = await database.tenant(tenant, operation, options);
        const hold = holdNext;
        holdNext = null;
        if (hold) await hold;
        return result;
      },
    };
    service = build({ database: slow });
    service.start();
    await service.idle();
    const hook = await webhook();
    let release: () => void = () => {};
    holdNext = new Promise<void>((resolve) => {
      release = resolve;
    });
    // The five-minute refresh: its read of the rules (none yet) is held.
    periodic?.();
    await new Promise((resolve) => setTimeout(resolve, 20));
    await rule({
      kind: "finding",
      name: "Saved meanwhile",
      destinations: [hook.destination.id],
      quietMinutes: 5,
      minimumSeverity: "critical",
      modules: [],
    });
    release();
    await service.idle();
    await service.sink.deliver(notification({ severity: "critical" }));
    await service.idle();
    expect(sent).toHaveLength(1);
  });
});

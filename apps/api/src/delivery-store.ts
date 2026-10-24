import { createHash } from "node:crypto";
import {
  type AlertRule,
  alertRuleSchema,
  type DeliveryError,
  type DeliveryRecord,
  type DeliverySource,
  type DestinationKind,
  deliveryErrorSchema,
  deliveryRecordSchema,
  destinationKindSchema,
  MAX_ALERT_RULES,
} from "@prism365/contracts";
import type { SqlClient } from "@prism365/storage";

// Destinations and the delivery log in prism365_destination and prism365_delivery
// (007-delivery.sql), and the alert rules in prism365_preference, key "alerts.rules".
// Deliveries never store the message: only where it went, what sent it, how many items it
// covered and how it went.

const RULES_KEY = "alerts.rules";
// How long the delivery log is kept.
export const DELIVERY_RETENTION_DAYS = 90;
// The most subjects one delivery row remembers for quiet periods.
const MAX_SUBJECTS = 200;

export interface DestinationRow {
  id: string;
  kind: DestinationKind;
  label: string;
  host: string | null;
  recipients: string[];
  secret: Buffer | null;
  keyId: string | null;
  includeDetails: boolean;
  createdAt: string;
  createdBy: string;
}

interface DestinationRecord {
  id: string;
  kind: string;
  label: string;
  host: string | null;
  recipients: unknown;
  secret: Buffer | Uint8Array | null;
  key_id: string | null;
  include_details: boolean;
  created_at: Date | string;
  created_by: string;
}

const parsedJson = (value: unknown): unknown => {
  if (typeof value !== "string") return value;
  try {
    return JSON.parse(value) as unknown;
  } catch {
    return null;
  }
};

function destinationRow(row: DestinationRecord): DestinationRow | null {
  const kind = destinationKindSchema.safeParse(row.kind);
  if (!kind.success) return null;
  const recipients = parsedJson(row.recipients);
  return {
    id: row.id,
    kind: kind.data,
    label: row.label,
    host: row.host,
    recipients: Array.isArray(recipients)
      ? recipients.filter((item): item is string => typeof item === "string").slice(0, 10)
      : [],
    secret: row.secret ? Buffer.from(row.secret) : null,
    keyId: row.key_id,
    includeDetails: row.include_details,
    createdAt: new Date(row.created_at).toISOString(),
    createdBy: row.created_by,
  };
}

const destinationColumns =
  "id, kind, label, host, recipients, secret, key_id, include_details, created_at, created_by";

export async function insertDestination(
  sql: SqlClient,
  tenantId: string,
  row: DestinationRow,
): Promise<void> {
  await sql.query(
    `INSERT INTO prism365_destination (tenant_id, ${destinationColumns})
     VALUES ($1,$2,$3,$4,$5,$6::jsonb,$7,$8,$9,$10,$11)`,
    [
      tenantId,
      row.id,
      row.kind,
      row.label,
      row.host,
      row.kind === "email" ? JSON.stringify(row.recipients) : null,
      row.secret,
      row.keyId,
      row.includeDetails,
      row.createdAt,
      row.createdBy,
    ],
  );
}

export async function listDestinations(
  sql: SqlClient,
  tenantId: string,
): Promise<DestinationRow[]> {
  const result = await sql.query<DestinationRecord>(
    `SELECT ${destinationColumns} FROM prism365_destination WHERE tenant_id=$1
      ORDER BY created_at, id LIMIT 100`,
    [tenantId],
  );
  return result.rows.map(destinationRow).filter((row): row is DestinationRow => row !== null);
}

export async function getDestination(
  sql: SqlClient,
  tenantId: string,
  id: string,
): Promise<DestinationRow | null> {
  const result = await sql.query<DestinationRecord>(
    `SELECT ${destinationColumns} FROM prism365_destination WHERE tenant_id=$1 AND id=$2`,
    [tenantId, id],
  );
  const row = result.rows[0];
  return row ? destinationRow(row) : null;
}

export async function deleteDestination(
  sql: SqlClient,
  tenantId: string,
  id: string,
): Promise<boolean> {
  const result = await sql.query("DELETE FROM prism365_destination WHERE tenant_id=$1 AND id=$2", [
    tenantId,
    id,
  ]);
  return result.rowCount === 1;
}

// A short digest of a subject, so a quiet period can be checked without keeping its name.
export function subjectDigest(ruleId: string, subject: string): string {
  return createHash("sha256").update(`${ruleId}\n${subject}`).digest("hex").slice(0, 24);
}

export interface NewDelivery {
  id: string;
  destinationId: string;
  source: DeliverySource;
  ruleId: string | null;
  subjects: readonly string[];
  items: number;
  createdAt: string;
}

export async function insertDelivery(
  sql: SqlClient,
  tenantId: string,
  delivery: NewDelivery,
): Promise<void> {
  await sql.query(
    `INSERT INTO prism365_delivery (tenant_id, id, destination_id, source, rule_id, subjects,
       items, status, attempts, last_error, created_at, updated_at)
     VALUES ($1,$2,$3,$4,$5,$6::jsonb,$7,'pending',0,NULL,$8,$8)`,
    [
      tenantId,
      delivery.id,
      delivery.destinationId,
      delivery.source,
      delivery.ruleId,
      JSON.stringify(delivery.subjects.slice(0, MAX_SUBJECTS)),
      Math.min(delivery.items, 100_000),
      delivery.createdAt,
    ],
  );
}

export async function updateDelivery(
  sql: SqlClient,
  tenantId: string,
  id: string,
  update: {
    status: "pending" | "sent" | "failed";
    attempts: number;
    lastError: DeliveryError | null;
    at: string;
  },
): Promise<void> {
  await sql.query(
    `UPDATE prism365_delivery SET status=$3, attempts=$4, last_error=$5, updated_at=$6
      WHERE tenant_id=$1 AND id=$2`,
    [tenantId, id, update.status, update.attempts, update.lastError, update.at],
  );
}

interface DeliveryRow {
  id: string;
  destination_id: string;
  source: string;
  rule_id: string | null;
  items: number;
  status: string;
  attempts: number;
  last_error: string | null;
  created_at: Date | string;
  updated_at: Date | string;
}

export async function recentDeliveries(
  sql: SqlClient,
  tenantId: string,
  limit = 50,
): Promise<DeliveryRecord[]> {
  const result = await sql.query<DeliveryRow>(
    `SELECT id, destination_id, source, rule_id, items, status, attempts, last_error, created_at,
            updated_at
       FROM prism365_delivery WHERE tenant_id=$1 ORDER BY created_at DESC, id DESC LIMIT $2`,
    [tenantId, limit],
  );
  return result.rows.flatMap((row) => {
    const parsed = deliveryRecordSchema.safeParse({
      id: row.id,
      destinationId: row.destination_id,
      source: row.source,
      ruleId: row.rule_id,
      items: row.items,
      status: row.status,
      attempts: row.attempts,
      lastError: deliveryErrorSchema.nullable().catch("error").parse(row.last_error),
      createdAt: new Date(row.created_at).toISOString(),
      updatedAt: new Date(row.updated_at).toISOString(),
    });
    return parsed.success ? [parsed.data] : [];
  });
}

// The subjects a rule has already alerted on since `since`, whatever became of those deliveries:
// a destination that keeps failing is not tried again for the same subject until the quiet
// period ends.
export async function quietSubjects(
  sql: SqlClient,
  tenantId: string,
  ruleId: string,
  since: Date,
): Promise<Set<string>> {
  const result = await sql.query<{ subjects: unknown }>(
    `SELECT subjects FROM prism365_delivery
      WHERE tenant_id=$1 AND rule_id=$2 AND created_at >= $3
      ORDER BY created_at DESC LIMIT 500`,
    [tenantId, ruleId, since.toISOString()],
  );
  const seen = new Set<string>();
  for (const row of result.rows) {
    const subjects = parsedJson(row.subjects);
    if (Array.isArray(subjects))
      for (const subject of subjects) if (typeof subject === "string") seen.add(subject);
  }
  return seen;
}

// Deliveries a stopped process left pending can never finish: the message was never stored.
export async function interruptDeliveries(
  sql: SqlClient,
  tenantId: string,
  before: Date,
): Promise<number> {
  const result = await sql.query(
    `UPDATE prism365_delivery SET status='failed', last_error='interrupted', updated_at=now()
      WHERE tenant_id=$1 AND status='pending' AND created_at < $2`,
    [tenantId, before.toISOString()],
  );
  return result.rowCount;
}

export async function pruneDeliveries(sql: SqlClient, tenantId: string, now: Date): Promise<void> {
  await sql.query("DELETE FROM prism365_delivery WHERE tenant_id=$1 AND created_at < $2", [
    tenantId,
    new Date(now.getTime() - DELIVERY_RETENTION_DAYS * 86_400_000).toISOString(),
  ]);
}

// Stored settings are kept exactly as stored. A save changes only what it is about (one rule
// added, one rule's "enabled", one rule removed, one audit's export) and writes every other item
// back untouched, so a version of Prism365 that does not know a field can never drop it.
//
// An item is used only when this version understands all of it. One that does not parse, or
// carries a field this version does not know (a newer version may have added a field that
// narrows what it means), is not evaluated at all: a damaged or newer setting can only stop an
// alert, never widen one. Such items still count towards the limit.

const asObject = (value: unknown): Record<string, unknown> | null =>
  value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;

// True when the stored object has a key the parsed one does not: a field this version drops.
const unknownFields = (raw: unknown, parsed: object) =>
  Object.keys(asObject(raw) ?? {}).some((key) => !Object.hasOwn(parsed, key));

function storedList(value: unknown, field: string, limit: number): unknown[] {
  const list = asObject(parsedJson(value))?.[field];
  return Array.isArray(list) ? list.slice(0, limit) : [];
}

// A rule this version keeps but does not run, as the page lists it.
export interface OtherRule {
  id: string | null;
  name: string | null;
  // Made by a newer version (it parses, with fields this one does not know), or unreadable.
  newer: boolean;
}

export interface StoredRules {
  rules: AlertRule[];
  others: OtherRule[];
}

export function storedAlertRules(value: unknown): StoredRules {
  const rules: AlertRule[] = [];
  const others: OtherRule[] = [];
  const ids = new Set<string>();
  for (const raw of storedList(value, "rules", MAX_ALERT_RULES)) {
    const parsed = alertRuleSchema.safeParse(raw);
    const object = asObject(raw);
    const id = typeof object?.id === "string" && object.id.length <= 100 ? object.id : null;
    const name =
      typeof object?.name === "string" ? object.name.replace(/[\p{Cc}]/gu, " ").slice(0, 80) : null;
    if (parsed.success && !unknownFields(raw, parsed.data) && !ids.has(parsed.data.id)) {
      ids.add(parsed.data.id);
      rules.push(parsed.data);
    } else others.push({ id, name, newer: parsed.success && unknownFields(raw, parsed.data) });
  }
  return { rules, others };
}

export async function readStoredRules(sql: SqlClient, tenantId: string): Promise<StoredRules> {
  const result = await sql.query<{ value: unknown }>(
    "SELECT value FROM prism365_preference WHERE tenant_id=$1 AND key=$2",
    [tenantId, RULES_KEY],
  );
  return storedAlertRules(result.rows[0]?.value);
}

// The rules this version runs.
export async function readAlertRules(sql: SqlClient, tenantId: string): Promise<AlertRule[]> {
  return (await readStoredRules(sql, tenantId)).rules;
}

export type RuleChange = { add: AlertRule } | { id: string; enabled: boolean } | { remove: string };

// Applies one change under a row lock, so two administrators saving at once apply one after the
// other. null when there is no room for a new rule, or no stored rule with that id.
export async function changeAlertRules(
  sql: SqlClient,
  tenantId: string,
  actorId: string,
  change: RuleChange,
): Promise<StoredRules | null> {
  await sql.query(
    `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
     VALUES ($1,$2,'{"rules":[]}'::jsonb,now(),$3)
     ON CONFLICT (tenant_id, key) DO NOTHING`,
    [tenantId, RULES_KEY, actorId],
  );
  const locked = await sql.query<{ value: unknown }>(
    "SELECT value FROM prism365_preference WHERE tenant_id=$1 AND key=$2 FOR UPDATE",
    [tenantId, RULES_KEY],
  );
  const stored = asObject(parsedJson(locked.rows[0]?.value)) ?? {};
  const items = storedList(stored, "rules", MAX_ALERT_RULES);
  if ("add" in change) {
    if (items.length >= MAX_ALERT_RULES) return null;
    items.push(alertRuleSchema.parse(change.add));
  } else {
    const id = "remove" in change ? change.remove : change.id;
    const index = items.findIndex((item) => asObject(item)?.id === id);
    if (index < 0) return null;
    if ("remove" in change) items.splice(index, 1);
    else items[index] = { ...asObject(items[index]), enabled: change.enabled };
  }
  await sql.query(
    `UPDATE prism365_preference SET value=$3::jsonb, updated_at=now(), updated_by=$4
      WHERE tenant_id=$1 AND key=$2`,
    [tenantId, RULES_KEY, JSON.stringify({ ...stored, rules: items }), actorId],
  );
  return storedAlertRules({ rules: items });
}

// Reads one preference under a row lock (creating it first when missing), applies `change`, and
// writes the result back. null from `change` leaves the row as it was.
async function changePreference<T>(
  sql: SqlClient,
  tenantId: string,
  key: string,
  actorId: string,
  read: (value: unknown) => T,
  change: (current: T) => T | null,
  write: (next: T) => unknown,
): Promise<T | null> {
  await sql.query(
    `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
     VALUES ($1,$2,'{}'::jsonb,now(),$3)
     ON CONFLICT (tenant_id, key) DO NOTHING`,
    [tenantId, key, actorId],
  );
  const locked = await sql.query<{ value: unknown }>(
    "SELECT value FROM prism365_preference WHERE tenant_id=$1 AND key=$2 FOR UPDATE",
    [tenantId, key],
  );
  const next = change(read(parsedJson(locked.rows[0]?.value)));
  if (next === null) return null;
  await sql.query(
    `UPDATE prism365_preference SET value=$3::jsonb, updated_at=now(), updated_by=$4
      WHERE tenant_id=$1 AND key=$2`,
    [tenantId, key, JSON.stringify(write(next)), actorId],
  );
  return next;
}

const EXPORTS_KEY = "alerts.exports";
const MAX_EXPORTS = 50;
const WEEKLY_KEY = "alerts.weekly";
const uuidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const auditIdPattern = /^[a-z][a-z0-9-]*(?:\.[a-z][a-z0-9-]*)+$/;

// Which audit's scheduled runs go to which destination, and who chose it. Read leniently: an
// entry that does not parse is left out.
export interface StoredExport {
  auditId: string;
  destinationId: string;
  setBy: string;
}

const exportKeys = new Set(["auditId", "destinationId", "setBy"]);

// The entries this version uses: well formed, with no field it does not know, the first for each
// audit. Every other entry is kept as stored and left alone.
function storedExports(value: unknown): StoredExport[] {
  const exports: StoredExport[] = [];
  const seen = new Set<string>();
  for (const item of storedList(value, "exports", MAX_EXPORTS)) {
    const entry = asObject(item);
    if (
      !entry ||
      Object.keys(entry).some((key) => !exportKeys.has(key)) ||
      typeof entry.auditId !== "string" ||
      entry.auditId.length > 80 ||
      !auditIdPattern.test(entry.auditId) ||
      typeof entry.destinationId !== "string" ||
      !uuidPattern.test(entry.destinationId) ||
      typeof entry.setBy !== "string" ||
      !uuidPattern.test(entry.setBy) ||
      seen.has(entry.auditId)
    )
      continue;
    seen.add(entry.auditId);
    exports.push({
      auditId: entry.auditId,
      destinationId: entry.destinationId,
      setBy: entry.setBy,
    });
  }
  return exports;
}

export async function readAlertExports(sql: SqlClient, tenantId: string): Promise<StoredExport[]> {
  const result = await sql.query<{ value: unknown }>(
    "SELECT value FROM prism365_preference WHERE tenant_id=$1 AND key=$2",
    [tenantId, EXPORTS_KEY],
  );
  return storedExports(result.rows[0]?.value);
}

// Replaces the entries for one audit, whatever they were, and leaves every other entry as stored.
// null when the change would pass the limit.
export async function setAlertExport(
  sql: SqlClient,
  tenantId: string,
  actorId: string,
  auditId: string,
  destinationId: string | null,
): Promise<StoredExport[] | null> {
  const saved = await changePreference(
    sql,
    tenantId,
    EXPORTS_KEY,
    actorId,
    (value) => asObject(value) ?? {},
    (stored) => {
      const items = storedList(stored, "exports", MAX_EXPORTS).filter(
        (item) => asObject(item)?.auditId !== auditId,
      );
      if (destinationId) items.push({ auditId, destinationId, setBy: actorId });
      return items.length > MAX_EXPORTS ? null : { ...stored, exports: items };
    },
    (next) => next,
  );
  return saved ? storedExports(saved) : null;
}

export interface StoredWeekly {
  destinationId: string;
  setBy: string;
}

function storedWeekly(value: unknown): StoredWeekly | null {
  const entry = value as Partial<StoredWeekly> | null;
  if (
    typeof entry?.destinationId !== "string" ||
    !uuidPattern.test(entry.destinationId) ||
    typeof entry.setBy !== "string" ||
    !uuidPattern.test(entry.setBy)
  )
    return null;
  return { destinationId: entry.destinationId, setBy: entry.setBy };
}

export async function readWeekly(sql: SqlClient, tenantId: string): Promise<StoredWeekly | null> {
  const result = await sql.query<{ value: unknown }>(
    "SELECT value FROM prism365_preference WHERE tenant_id=$1 AND key=$2",
    [tenantId, WEEKLY_KEY],
  );
  return storedWeekly(parsedJson(result.rows[0]?.value));
}

// Replaces the setting whole, so no lock is needed: the last save wins.
export async function setWeekly(
  sql: SqlClient,
  tenantId: string,
  actorId: string,
  destinationId: string | null,
): Promise<StoredWeekly | null> {
  const value = destinationId ? { destinationId, setBy: actorId } : null;
  await sql.query(
    `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
     VALUES ($1,$2,$3::jsonb,now(),$4)
     ON CONFLICT (tenant_id, key) DO UPDATE SET value=EXCLUDED.value, updated_at=EXCLUDED.updated_at,
       updated_by=EXCLUDED.updated_by`,
    [tenantId, WEEKLY_KEY, JSON.stringify(value ?? {}), actorId],
  );
  return value;
}

// Whether a weekly summary was already sent (or tried) since `since`.
export async function summarySentSince(
  sql: SqlClient,
  tenantId: string,
  since: Date,
): Promise<boolean> {
  const result = await sql.query(
    "SELECT 1 FROM prism365_delivery WHERE tenant_id=$1 AND source='summary' AND created_at >= $2 LIMIT 1",
    [tenantId, since.toISOString()],
  );
  // Counted by rows: PGlite reports no row count for a SELECT.
  return result.rows.length > 0;
}

import {
  AUDIT_SCHEDULE_PERIOD_MILLISECONDS,
  type AuditSchedule,
  type AuditScheduleFrequency,
  auditIdSchema,
  auditScheduleFrequencySchema,
  auditScheduleSchema,
  cleanText,
  MAX_AUDIT_SCHEDULE_ENTRIES,
} from "@prism365/contracts";
import type { SqlClient } from "@prism365/storage";
import { z } from "zod";
import type { Audit } from "./definition.ts";

// Audit schedules: which audits run on their own, daily or weekly. One setting per tenant in
// prism365_preference, saved by an administrator and read by the API's audit scheduler
// (apps/api/src/audit-schedule.ts), which runs due audits on the collecting administrator's
// sign-in while collection is on.
//
// Stored as { audits: { <auditId>: "off" | "daily" | "weekly" }, scheduledBy: { <auditId>:
// <actor id> }, names: { <auditId>: <name> }, grants: { <auditId>: [<permission>] } }. For each
// scheduled audit: who last set it to its cadence, their name as the page shows it, and the
// permissions their own sign-in had granted among the audit's packs then. A scheduled run reads
// with no more than those, whoever's sign-in it uses, and its ledger entry names the scheduler
// as well as the administrator whose sign-in it reads with. An entry saved before grants were
// kept has none, and its runs read with the audit's required permissions only.

const PREFERENCE_KEY = "audits.schedule";

export type AuditScheduleMap = Record<string, AuditScheduleFrequency>;

export interface StoredAuditSchedule {
  // Every well-formed entry, including audits this build does not know: a newer build may have
  // saved them, and saving here must not erase them.
  audits: AuditScheduleMap;
  scheduledBy: Record<string, string>;
  names: Record<string, string>;
  grants: Record<string, string[]>;
}

// Who saves a schedule: their name, and for each audit the permissions their own session granted
// among its packs (auditReadScope's `granted`).
export interface ScheduleSaver {
  name: string | null;
  grants: (audit: Audit) => readonly string[];
}

const actorIdSchema = z.uuid();
const permissionSchema = z.string().regex(/^[A-Za-z][A-Za-z0-9]*(?:\.[A-Za-z][A-Za-z0-9]*){0,4}$/);
const grantsSchema = z.array(permissionSchema).max(20);
const nameOf = (value: unknown) =>
  typeof value === "string" ? cleanText(value).trim().slice(0, 200) || null : null;

const record = (value: unknown): Record<string, unknown> =>
  value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : {};

// The stored value, read leniently: an entry whose id is not an audit id or whose value is not a
// frequency is left out, so a damaged setting can only stop audits from running, never start one.
// Own keys only, and at most MAX_AUDIT_SCHEDULE_ENTRIES.
export function storedAuditSchedule(value: unknown): StoredAuditSchedule {
  const saved = record(record(value).audits);
  const by = record(record(value).scheduledBy);
  const savedNames = record(record(value).names);
  const savedGrants = record(record(value).grants);
  const stored: StoredAuditSchedule = { audits: {}, scheduledBy: {}, names: {}, grants: {} };
  for (const id of Object.keys(saved)) {
    if (Object.keys(stored.audits).length >= MAX_AUDIT_SCHEDULE_ENTRIES) break;
    const frequency = auditScheduleFrequencySchema.safeParse(saved[id]);
    if (!auditIdSchema.safeParse(id).success || !frequency.success) continue;
    stored.audits[id] = frequency.data;
    if (frequency.data === "off") continue;
    const actor = Object.hasOwn(by, id) ? actorIdSchema.safeParse(by[id]) : null;
    if (actor?.success) stored.scheduledBy[id] = actor.data;
    const name = Object.hasOwn(savedNames, id) ? nameOf(savedNames[id]) : null;
    if (name) stored.names[id] = name;
    // Grants that do not parse are dropped whole: the entry then reads with required permissions
    // only, never with more.
    const grants = Object.hasOwn(savedGrants, id) ? grantsSchema.safeParse(savedGrants[id]) : null;
    if (grants?.success) stored.grants[id] = [...new Set(grants.data)];
  }
  return stored;
}

// Every known audit with its frequency, "off" unless the value names it with a frequency. An id
// that is not a known audit is left out here, so nothing this build does not know ever runs.
export function knownAuditSchedule(value: unknown, audits: readonly Audit[]): AuditScheduleMap {
  const saved = record(record(value).audits);
  const schedule: AuditScheduleMap = {};
  for (const audit of audits) {
    const id = audit.definition.id;
    const frequency = Object.hasOwn(saved, id)
      ? auditScheduleFrequencySchema.safeParse(saved[id])
      : null;
    schedule[id] = frequency?.success ? frequency.data : "off";
  }
  return schedule;
}

// How many audits run on their own.
export function scheduledAuditCount(schedule: AuditScheduleMap): number {
  return Object.values(schedule).filter((frequency) => frequency !== "off").length;
}

function parsedValue(value: unknown): unknown {
  if (typeof value !== "string") return value ?? null;
  try {
    return JSON.parse(value) as unknown;
  } catch {
    return null;
  }
}

export async function readStoredAuditSchedule(
  sql: SqlClient,
  tenantId: string,
): Promise<StoredAuditSchedule> {
  const result = await sql.query<{ value: unknown }>(
    "SELECT value FROM prism365_preference WHERE tenant_id=$1 AND key=$2",
    [tenantId, PREFERENCE_KEY],
  );
  return storedAuditSchedule(parsedValue(result.rows[0]?.value));
}

export async function readAuditSchedule(
  sql: SqlClient,
  tenantId: string,
  audits: readonly Audit[],
): Promise<AuditScheduleMap> {
  return knownAuditSchedule(await readStoredAuditSchedule(sql, tenantId), audits);
}

// The known audits a request would change, with the cadence each would get. Entries for unknown
// audits, and entries that match what is saved, are not changes.
export function auditScheduleChanges(
  current: AuditScheduleMap,
  request: AuditSchedule,
): { auditId: string; frequency: AuditScheduleFrequency }[] {
  return Object.entries(request.audits)
    .filter(([id, frequency]) => Object.hasOwn(current, id) && current[id] !== frequency)
    .map(([auditId, frequency]) => ({ auditId, frequency }));
}

// Saves the audits the request names and leaves every other entry as it is, including audits
// this build does not know. A request entry for an unknown audit is dropped. The administrator
// saving becomes the one who scheduled each audit set to daily or weekly, with their name and the
// permissions `saver` says they had for it; without `saver`, the entry keeps no grants and runs
// with required permissions only. Returns the known audits as saved.
//
// The row is locked for the transaction (created first when missing), so two saves at the same
// moment apply one after the other: neither can undo the other's audits.
export async function writeAuditSchedule(
  sql: SqlClient,
  tenantId: string,
  schedule: AuditSchedule,
  actorId: string,
  audits: readonly Audit[],
  saver?: ScheduleSaver,
): Promise<AuditScheduleMap> {
  const request = auditScheduleSchema.parse(schedule);
  const actor = actorIdSchema.parse(actorId);
  await sql.query(
    `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
     VALUES ($1,$2,'{"audits":{}}'::jsonb,now(),$3)
     ON CONFLICT (tenant_id, key) DO NOTHING`,
    [tenantId, PREFERENCE_KEY, actor],
  );
  const locked = await sql.query<{ value: unknown }>(
    "SELECT value FROM prism365_preference WHERE tenant_id=$1 AND key=$2 FOR UPDATE",
    [tenantId, PREFERENCE_KEY],
  );
  const stored = storedAuditSchedule(parsedValue(locked.rows[0]?.value));
  // Known audits first, so the entry limit can only ever leave out an unknown one.
  const next: StoredAuditSchedule = { audits: {}, scheduledBy: {}, names: {}, grants: {} };
  const carry = (id: string) => {
    next.audits[id] = stored.audits[id] ?? "off";
    for (const field of ["scheduledBy", "names", "grants"] as const) {
      const value = stored[field][id];
      if (value !== undefined) (next[field] as Record<string, unknown>)[id] = value;
    }
  };
  for (const audit of audits) {
    const id = audit.definition.id;
    if (Object.hasOwn(request.audits, id)) {
      const frequency = request.audits[id] ?? "off";
      next.audits[id] = frequency;
      if (frequency === "off") continue;
      next.scheduledBy[id] = actor;
      const name = nameOf(saver?.name);
      if (name) next.names[id] = name;
      if (saver) next.grants[id] = [...new Set(saver.grants(audit))];
    } else if (Object.hasOwn(stored.audits, id)) carry(id);
  }
  for (const id of Object.keys(stored.audits)) if (!Object.hasOwn(next.audits, id)) carry(id);
  await sql.query(
    `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
     VALUES ($1,$2,$3::jsonb,now(),$4)
     ON CONFLICT (tenant_id, key) DO UPDATE SET value=EXCLUDED.value, updated_at=EXCLUDED.updated_at,
       updated_by=EXCLUDED.updated_by`,
    [tenantId, PREFERENCE_KEY, JSON.stringify(storedAuditSchedule(next)), actor],
  );
  return knownAuditSchedule(next, audits);
}

// Whether a scheduled audit is due: it never ran, or its newest run of any outcome, manual or
// scheduled, started more than a day (a week) ago. Counting from the start of any run means a
// failed scheduled run waits for its next turn instead of retrying, and a manual run counts as
// this period's run. A start in the future (the clock went back) is not due until a full period
// after it, so a clock change can delay a run but never start one every few minutes; the audit's
// page says so.
export function auditDue(
  frequency: AuditScheduleFrequency,
  lastStart: Date | null,
  now: Date,
): boolean {
  if (frequency === "off") return false;
  if (lastStart === null) return true;
  return now.getTime() - lastStart.getTime() > AUDIT_SCHEDULE_PERIOD_MILLISECONDS[frequency];
}

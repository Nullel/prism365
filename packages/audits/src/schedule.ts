import {
  AUDIT_SCHEDULE_PERIOD_MILLISECONDS,
  type AuditSchedule,
  type AuditScheduleFrequency,
  auditIdSchema,
  auditScheduleFrequencySchema,
  auditScheduleSchema,
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
// <actor id> } }: who last set each scheduled audit to its cadence, so a scheduled run's ledger
// entry can name them as well as the administrator whose sign-in it reads with.

const PREFERENCE_KEY = "audits.schedule";

export type AuditScheduleMap = Record<string, AuditScheduleFrequency>;

export interface StoredAuditSchedule {
  // Every well-formed entry, including audits this build does not know: a newer build may have
  // saved them, and saving here must not erase them.
  audits: AuditScheduleMap;
  scheduledBy: Record<string, string>;
}

const actorIdSchema = z.uuid();

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
  const audits: AuditScheduleMap = {};
  const scheduledBy: Record<string, string> = {};
  for (const id of Object.keys(saved)) {
    if (Object.keys(audits).length >= MAX_AUDIT_SCHEDULE_ENTRIES) break;
    const frequency = auditScheduleFrequencySchema.safeParse(saved[id]);
    if (!auditIdSchema.safeParse(id).success || !frequency.success) continue;
    audits[id] = frequency.data;
    const actor = Object.hasOwn(by, id) ? actorIdSchema.safeParse(by[id]) : null;
    if (frequency.data !== "off" && actor?.success) scheduledBy[id] = actor.data;
  }
  return { audits, scheduledBy };
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

export async function readStoredAuditSchedule(
  sql: SqlClient,
  tenantId: string,
): Promise<StoredAuditSchedule> {
  const result = await sql.query<{ value: unknown }>(
    "SELECT value FROM prism365_preference WHERE tenant_id=$1 AND key=$2",
    [tenantId, PREFERENCE_KEY],
  );
  let value: unknown = result.rows[0]?.value ?? null;
  if (typeof value === "string")
    try {
      value = JSON.parse(value) as unknown;
    } catch {
      value = null;
    }
  return storedAuditSchedule(value);
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
// saving becomes the one who scheduled each audit set to daily or weekly. Returns the known
// audits as saved.
export async function writeAuditSchedule(
  sql: SqlClient,
  tenantId: string,
  schedule: AuditSchedule,
  actorId: string,
  audits: readonly Audit[],
): Promise<AuditScheduleMap> {
  const request = auditScheduleSchema.parse(schedule);
  const actor = actorIdSchema.parse(actorId);
  const stored = await readStoredAuditSchedule(sql, tenantId);
  // Known audits first, so the entry limit can only ever leave out an unknown one.
  const next: StoredAuditSchedule = { audits: {}, scheduledBy: {} };
  const keep = (id: string, frequency: AuditScheduleFrequency, by: string | undefined) => {
    next.audits[id] = frequency;
    if (frequency !== "off" && by) next.scheduledBy[id] = by;
  };
  for (const audit of audits) {
    const id = audit.definition.id;
    if (Object.hasOwn(request.audits, id)) {
      const frequency = request.audits[id] ?? "off";
      keep(id, frequency, actor);
    } else if (Object.hasOwn(stored.audits, id))
      keep(id, stored.audits[id] ?? "off", stored.scheduledBy[id]);
  }
  for (const [id, frequency] of Object.entries(stored.audits))
    if (!Object.hasOwn(next.audits, id)) keep(id, frequency, stored.scheduledBy[id]);
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

import {
  type AuditSchedule,
  type AuditScheduleFrequency,
  auditScheduleFrequencySchema,
  auditScheduleSchema,
} from "@prism365/contracts";
import type { SqlClient } from "@prism365/storage";
import type { Audit } from "./definition.ts";

// Audit schedules: which audits run on their own, daily or weekly. One setting per tenant in
// prism365_preference, saved by an administrator and read by the API's audit scheduler
// (apps/api/src/audit-schedule.ts), which runs due audits on the collecting administrator's
// sign-in while collection is on.

const PREFERENCE_KEY = "audits.schedule";

const DAY_MILLISECONDS = 24 * 60 * 60_000;
export const SCHEDULE_PERIOD_MILLISECONDS: Record<
  Exclude<AuditScheduleFrequency, "off">,
  number
> = {
  daily: DAY_MILLISECONDS,
  weekly: 7 * DAY_MILLISECONDS,
};

export type AuditScheduleMap = Record<string, AuditScheduleFrequency>;

// Every known audit with its frequency, "off" unless the value names it. An id that is not a
// known audit is dropped, and so is a value that is not a frequency: either reads as off, so a
// damaged setting can only stop audits from running, never start one.
export function knownAuditSchedule(value: unknown, audits: readonly Audit[]): AuditScheduleMap {
  const saved =
    value !== null && typeof value === "object" && !Array.isArray(value)
      ? (value as { audits?: unknown }).audits
      : undefined;
  const entries =
    saved !== null && typeof saved === "object" && !Array.isArray(saved)
      ? (saved as Record<string, unknown>)
      : {};
  const schedule: AuditScheduleMap = {};
  for (const audit of audits) {
    const id = audit.definition.id;
    const frequency = Object.hasOwn(entries, id)
      ? auditScheduleFrequencySchema.safeParse(entries[id])
      : null;
    schedule[id] = frequency?.success ? frequency.data : "off";
  }
  return schedule;
}

// How many audits run on their own. The access ledger records this number, never which audits.
export function scheduledAuditCount(schedule: AuditScheduleMap): number {
  return Object.values(schedule).filter((frequency) => frequency !== "off").length;
}

export async function readAuditSchedule(
  sql: SqlClient,
  tenantId: string,
  audits: readonly Audit[],
): Promise<AuditScheduleMap> {
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
  return knownAuditSchedule(value, audits);
}

// Saves the whole schedule: an audit the body does not name is off. Unknown audit ids are
// dropped before saving. Returns what was saved.
export async function writeAuditSchedule(
  sql: SqlClient,
  tenantId: string,
  schedule: AuditSchedule,
  actorId: string,
  audits: readonly Audit[],
): Promise<AuditScheduleMap> {
  const saved = knownAuditSchedule(auditScheduleSchema.parse(schedule), audits);
  await sql.query(
    `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
     VALUES ($1,$2,$3::jsonb,now(),$4)
     ON CONFLICT (tenant_id, key) DO UPDATE SET value=EXCLUDED.value, updated_at=EXCLUDED.updated_at,
       updated_by=EXCLUDED.updated_by`,
    [tenantId, PREFERENCE_KEY, JSON.stringify({ audits: saved }), actorId],
  );
  return saved;
}

// Whether a scheduled audit is due: it never ran, or its newest run of any outcome, manual or
// scheduled, started more than a day (a week) ago. Counting from the start of any run means a
// failed scheduled run waits for its next turn instead of retrying, and a manual run counts as
// this period's run. A start in the future (the clock went back) is not due until a full period
// after it, so a clock change can delay a run but never start one every few minutes.
export function auditDue(
  frequency: AuditScheduleFrequency,
  lastStart: Date | null,
  now: Date,
): boolean {
  if (frequency === "off") return false;
  if (lastStart === null) return true;
  return now.getTime() - lastStart.getTime() > SCHEDULE_PERIOD_MILLISECONDS[frequency];
}

import { z } from "zod";
import { auditIdSchema, auditRunTriggerSchema, COMBINED_RISKS_AUDIT_ID } from "./audits.ts";

// The full review (docs/audits.md, section 22) and the families the Audits page groups the audits
// into. Both live here so the page, the API and the tests share one list.

const text = z.string().transform((value) => value.slice(0, 500));
const time = z.string().datetime({ offset: true });

// Families -----------------------------------------------------------------------------------------

export interface AuditFamilyEntry {
  auditId: string;
  // "main" is the audit that gives the whole picture of the family; "narrower" ones are views of
  // part of it and are shown indented under it.
  role?: "main" | "narrower";
  // Said under the audit's name on the page.
  note?: string;
}

export interface AuditFamily {
  id: string;
  title: string;
  audits: readonly AuditFamilyEntry[];
}

// In the order the page shows them. Every built-in audit is in exactly one family; a test in the
// audits package fails when a new audit is in none or in two.
export const AUDIT_FAMILIES: readonly AuditFamily[] = [
  {
    id: "summary",
    title: "Summary",
    audits: [{ auditId: COMBINED_RISKS_AUDIT_ID, role: "main" }],
  },
  {
    id: "account-lifecycle",
    title: "Account lifecycle",
    audits: [
      { auditId: "accounts.activity", role: "main", note: "The full picture of each account." },
      { auditId: "accounts.inactive", role: "narrower" },
      { auditId: "licenses.reclaim", role: "narrower" },
      { auditId: "usage.m365-apps", role: "narrower" },
    ],
  },
  {
    id: "identity-protection",
    title: "Identity protection",
    audits: [
      { auditId: "mfa.registration" },
      { auditId: "identity.conditional-access" },
      { auditId: "identity.risky-users" },
      { auditId: "accounts.passwords" },
    ],
  },
  {
    id: "access-ownership",
    title: "Access and ownership",
    audits: [
      { auditId: "directory.privileged-roles" },
      { auditId: "guests.review" },
      { auditId: "groups.teams" },
      { auditId: "apps.consents" },
      { auditId: "apps.credentials" },
    ],
  },
  {
    id: "devices-mail",
    title: "Devices and mail",
    audits: [{ auditId: "devices.managed" }, { auditId: "mail.mailbox-storage" }],
  },
  {
    id: "tenant-setup",
    title: "Tenant setup",
    audits: [{ auditId: "licenses.capacity" }],
  },
];

// The family an audit belongs to, or undefined for one no family lists.
export function auditFamilyOf(auditId: string): AuditFamily | undefined {
  return AUDIT_FAMILIES.find((family) => family.audits.some((entry) => entry.auditId === auditId));
}

// The audits of a catalog in family order, with each family's entries. An audit no family lists
// is not dropped: it goes in a last group, so a new audit is never hidden from the page.
export function groupByFamily<T extends { definition: { id: string } }>(
  items: readonly T[],
): { family: { id: string; title: string }; entries: { item: T; entry: AuditFamilyEntry }[] }[] {
  const byId = new Map(items.map((item) => [item.definition.id, item]));
  const groups = AUDIT_FAMILIES.map((family) => ({
    family: { id: family.id, title: family.title },
    entries: family.audits.flatMap((entry) => {
      const item = byId.get(entry.auditId);
      return item ? [{ item, entry }] : [];
    }),
  })).filter((group) => group.entries.length > 0);
  const listed = new Set(AUDIT_FAMILIES.flatMap((family) => family.audits.map((e) => e.auditId)));
  const other = items.filter((item) => !listed.has(item.definition.id));
  if (other.length > 0)
    groups.push({
      family: { id: "other", title: "Other audits" },
      entries: other.map((item) => ({ item, entry: { auditId: item.definition.id } })),
    });
  return groups;
}

// A review ------------------------------------------------------------------------------------------

export const auditReviewAuditStateSchema = z.enum([
  "waiting",
  "running",
  "done",
  "failed",
  "skipped",
]);
export type AuditReviewAuditState = z.infer<typeof auditReviewAuditStateSchema>;

// running: audits are being run. stopping: Stop was asked and the audit in progress is ending.
// finished: every audit was run or skipped. stopped: it ended early (stopped, or collection ended
// during a scheduled review, or the server restarted).
export const auditReviewStateSchema = z.enum(["running", "stopping", "finished", "stopped"]);
export type AuditReviewState = z.infer<typeof auditReviewStateSchema>;

export const auditReviewAuditSchema = z.object({
  auditId: auditIdSchema,
  state: auditReviewAuditStateSchema,
  // Why an audit was skipped or failed, in plain words. null otherwise.
  reason: text.nullable(),
  runId: z.uuid().nullable(),
  startedAt: time.nullable(),
  finishedAt: time.nullable(),
});
export type AuditReviewAudit = z.infer<typeof auditReviewAuditSchema>;

export const auditReviewSchema = z.object({
  id: z.uuid(),
  state: auditReviewStateSchema,
  trigger: auditRunTriggerSchema,
  startedAt: time,
  finishedAt: time.nullable(),
  startedBy: z.object({ id: z.uuid(), name: text.nullable() }),
  // In the order they run.
  audits: z.array(auditReviewAuditSchema).max(40),
});
export type AuditReview = z.infer<typeof auditReviewSchema>;

// The current or last review; null when none has run.
export const auditReviewResponseSchema = z.object({ review: auditReviewSchema.nullable() });
export type AuditReviewResponse = z.infer<typeof auditReviewResponseSchema>;
export const auditReviewStartSchema = z.object({ review: auditReviewSchema });

// How far through a review is: audits that ended (done, failed or skipped) out of all of them.
export function auditReviewProgress(review: AuditReview): { ended: number; total: number } {
  return {
    ended: review.audits.filter((item) => item.state !== "waiting" && item.state !== "running")
      .length,
    total: review.audits.length,
  };
}

// The schedule -------------------------------------------------------------------------------------

export const auditReviewFrequencySchema = z.enum(["off", "weekly"]);
export type AuditReviewFrequency = z.infer<typeof auditReviewFrequencySchema>;

export const REVIEW_WEEKDAYS = [
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
  "Sunday",
] as const;

// What an administrator saves. weekday is 0 for Monday through 6 for Sunday; hour is the hour of
// the day (0 to 23) in the installation's time zone.
export const auditReviewScheduleSchema = z.object({
  frequency: auditReviewFrequencySchema,
  weekday: z.number().int().min(0).max(6),
  hour: z.number().int().min(0).max(23),
});
export type AuditReviewSchedule = z.infer<typeof auditReviewScheduleSchema>;

export const auditReviewScheduleSettingsSchema = auditReviewScheduleSchema.extend({
  timeZone: z.string().max(64),
  // Collection must be on for a scheduled review to run.
  collecting: z.boolean(),
  // The administrator who saved it, whose permissions it reads with.
  by: text.nullable(),
});
export type AuditReviewScheduleSettings = z.infer<typeof auditReviewScheduleSettingsSchema>;

import { z } from "zod";
import { auditIdSchema, auditRunTriggerSchema, auditSummaryItemSchema } from "./audits.ts";

// Audit trends and audit schedules (docs/audits.md, sections 10 and 11). A trend is each
// audit's summary tiles over time, one point per successful run, kept after the runs themselves
// are pruned. A schedule lets an audit run daily or weekly on the collecting administrator's
// sign-in.

const text = z.string().transform((value) => value.slice(0, 500));
const time = z.string().datetime({ offset: true });

// Trends ---------------------------------------------------------------------------------------

export const AUDIT_TILE_ID_LENGTH = 60;

// A tile's id within its audit, from its label: "Inactive 30+ days" is "inactive-30-days". Tiles
// carry no id of their own, and the label is what an administrator recognises from one run to
// the next. Only ASCII letters and digits survive, so the result never depends on a locale.
// The backfill in packages/storage/migrations/005-audit-trends.sql derives the same id in SQL:
// change both or neither.
export function auditTileId(label: string): string {
  const id = label
    .replace(/[^A-Za-z0-9]+/g, "-")
    .toLowerCase()
    .slice(0, AUDIT_TILE_ID_LENGTH);
  let start = 0;
  let end = id.length;
  while (start < end && id[start] === "-") start += 1;
  while (end > start && id[end - 1] === "-") end -= 1;
  return id.slice(start, end) || "tile";
}

export const auditTileIdSchema = z
  .string()
  .max(AUDIT_TILE_ID_LENGTH)
  .regex(/^[a-z0-9](?:[a-z0-9-]*[a-z0-9])?$/);

export const auditTileToneSchema = auditSummaryItemSchema.shape.tone;

export const auditTrendTileSchema = z.object({
  id: auditTileIdSchema,
  label: text,
  tone: auditTileToneSchema,
});
export type AuditTrendTile = z.infer<typeof auditTrendTileSchema>;

// Why a point does not compare with the newest one: it was made by another version of the
// audit, it covered another audit scope, or Microsoft answered only in part.
export const auditTrendDifferenceSchema = z.enum(["version", "scope", "incomplete"]);
export type AuditTrendDifference = z.infer<typeof auditTrendDifferenceSchema>;

export const auditTrendPointSchema = z.object({
  runId: z.uuid(),
  finishedAt: time,
  version: z.number().int().positive(),
  complete: z.boolean(),
  // The run was limited to some account domains (docs/audits.md, section 9).
  scopeLimited: z.boolean(),
  trigger: auditRunTriggerSchema,
  // Each tile's value by tile id. A tile the run did not have is missing.
  tiles: z.record(auditTileIdSchema, z.number().int().nonnegative()),
  // Empty when the point compares with the newest one.
  notComparable: z.array(auditTrendDifferenceSchema).max(3),
});
export type AuditTrendPoint = z.infer<typeof auditTrendPointSchema>;

export const MAX_AUDIT_TREND_DAYS = 3_650;
export const MAX_AUDIT_TREND_POINTS = 1_000;

export const auditTrendQuerySchema = z.object({
  days: z.coerce.number().int().min(1).max(MAX_AUDIT_TREND_DAYS).default(365),
});

export const auditTrendSchema = z.object({
  auditId: auditIdSchema,
  days: z.number().int().min(1).max(MAX_AUDIT_TREND_DAYS),
  // Oldest first.
  points: z.array(auditTrendPointSchema).max(MAX_AUDIT_TREND_POINTS),
  // The newest point's tiles, in the audit's order: the values that can be followed.
  tiles: z.array(auditTrendTileSchema).max(20),
  // More points fell in the window than one answer holds; the oldest of them were left out.
  truncated: z.boolean(),
});
export type AuditTrend = z.infer<typeof auditTrendSchema>;

// Schedules ------------------------------------------------------------------------------------

export const auditScheduleFrequencySchema = z.enum(["off", "daily", "weekly"]);
export type AuditScheduleFrequency = z.infer<typeof auditScheduleFrequencySchema>;

// Far more than there are audits; it only bounds what one request can make the server check.
export const MAX_AUDIT_SCHEDULE_ENTRIES = 100;

// The saved schedule (prism365_preference, key audits.schedule) and the body that saves it. An
// audit missing from the list does not run on its own.
export const auditScheduleSchema = z.object({
  audits: z
    .record(auditIdSchema, auditScheduleFrequencySchema)
    .refine((audits) => Object.keys(audits).length <= MAX_AUDIT_SCHEDULE_ENTRIES, {
      message: `At most ${MAX_AUDIT_SCHEDULE_ENTRIES} audits.`,
    }),
});
export type AuditSchedule = z.infer<typeof auditScheduleSchema>;

// What the settings route returns: every audit with its frequency, and whether collection is on,
// because nothing runs on a schedule while it is off.
export const auditScheduleSettingsSchema = auditScheduleSchema.extend({
  collecting: z.boolean(),
});
export type AuditScheduleSettings = z.infer<typeof auditScheduleSettingsSchema>;

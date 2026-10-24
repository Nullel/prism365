import { z } from "zod";
import { attentionSeveritySchema } from "./attention.ts";
import { auditIdSchema } from "./audits.ts";
import { changeCoverageSchema } from "./changes.ts";
import {
  auditScheduleFrequencySchema,
  auditTileIdSchema,
  auditTileToneSchema,
} from "./insights.ts";
import { personIdSchema } from "./people.ts";

// The "This week" page: what Prism365 collected and stored in the last 1, 7, 14 or 30 days,
// section by section. It is built from stored data only and never asks Microsoft anything
// (docs/this-week.md). Strings are provider text or built from it, so they are cut to size
// here; the lists are bounded where they are read, and again here so one odd row cannot turn the
// page into an error.

const text = z.string().transform((value) => value.slice(0, 500));
const nullableText = text.nullable();
const time = z.string().datetime({ offset: true });
const count = z.number().int().nonnegative();

export const DIGEST_DAYS = [1, 7, 14, 30] as const;
export const digestDaysSchema = z.union([z.literal(1), z.literal(7), z.literal(14), z.literal(30)]);
export type DigestDays = z.infer<typeof digestDaysSchema>;

export const digestQuerySchema = z.object({
  days: z.coerce.number().pipe(digestDaysSchema).default(7),
});

// How many rows each list holds. The totals next to a list say how many there were.
export const DIGEST_OPEN_FINDINGS = 10;
export const DIGEST_FLAGGED_CHANGES = 20;
export const DIGEST_QUIET_ACCOUNTS = 50;
export const DIGEST_POSTS = 20;
export const DIGEST_INCIDENTS = 20;
// Message center posts are looked for this far ahead.
export const DIGEST_DEADLINE_DAYS = 30;

// "ok": there is something to show. "empty": the data was read and holds nothing for this range.
// "unavailable": it could not be read, and the note says why. An unavailable section is never
// shown as empty, because "nothing happened" and "nothing was looked at" are different things.
export const digestSectionStateSchema = z.enum(["ok", "empty", "unavailable"]);
export type DigestSectionState = z.infer<typeof digestSectionStateSchema>;

const section = {
  state: digestSectionStateSchema,
  // One line in plain words: what the section holds, or why it could not be read.
  note: text,
};

export const digestFindingSchema = z.object({
  id: z.uuid(),
  title: text,
  severity: attentionSeveritySchema,
  // The finding's subject as a person would read it: its label, else its id.
  subjectLabel: text,
  // Set when the subject is a person, to open their page.
  personId: personIdSchema.nullable(),
  firstSeen: time,
});
export type DigestFinding = z.infer<typeof digestFindingSchema>;

export const digestSeverityCountsSchema = z.object({
  critical: count,
  high: count,
  medium: count,
  low: count,
  info: count,
});
export type DigestSeverityCounts = z.infer<typeof digestSeverityCountsSchema>;

export const digestSignInFindingsSchema = z.object({
  ...section,
  // Opened in the range, by severity.
  opened: digestSeverityCountsSchema,
  // The most severe findings still open, whenever they were opened.
  open: z.array(digestFindingSchema).max(DIGEST_OPEN_FINDINGS),
  openTotal: count,
  // Resolved in the range.
  resolved: count,
});
export type DigestSignInFindings = z.infer<typeof digestSignInFindingsSchema>;

export const digestDirectoryChangesSchema = z.object({
  ...section,
  // Every change stored for the range, and by category.
  total: count,
  categories: z.array(z.object({ category: nullableText, count })).max(50),
  // Changes the rules flagged in the range: the findings of the directory change module,
  // without its collection gaps (those are counted under Collection).
  flagged: z.array(digestFindingSchema).max(DIGEST_FLAGGED_CHANGES),
  flaggedTotal: count,
  coverage: changeCoverageSchema,
});
export type DigestDirectoryChanges = z.infer<typeof digestDirectoryChangesSchema>;

// A stored audit result a section was worked out from.
export const digestRunRefSchema = z.object({ runId: z.uuid(), at: time });
export type DigestRunRef = z.infer<typeof digestRunRefSchema>;

// Why two results do not compare directly: the same three reasons an audit's trend gives.
export const digestDifferenceSchema = z.enum(["version", "scope", "incomplete"]);
export type DigestDifference = z.infer<typeof digestDifferenceSchema>;

export const digestQuietAccountSchema = z.object({
  personId: personIdSchema.nullable(),
  name: nullableText,
  principal: nullableText,
  // Days since the last sign-in as the audit counted them; null when it could not tell.
  days: count.nullable(),
});
export type DigestQuietAccount = z.infer<typeof digestQuietAccountSchema>;

export const digestQuietAccountsSchema = z.object({
  ...section,
  // The audit and view the accounts can be looked at in.
  auditId: auditIdSchema,
  view: z.string().max(40),
  // Every account that went quiet; `items` holds the first 50.
  total: count,
  items: z.array(digestQuietAccountSchema).max(DIGEST_QUIET_ACCOUNTS),
  // The two results compared: the newest, and the newest from before the range began. null when
  // there is nothing to compare.
  latest: digestRunRefSchema.nullable(),
  earlier: digestRunRefSchema.nullable(),
  differences: z.array(digestDifferenceSchema).max(3),
});
export type DigestQuietAccounts = z.infer<typeof digestQuietAccountsSchema>;

export const digestDriftTileSchema = z.object({
  id: auditTileIdSchema,
  label: text,
  tone: auditTileToneSchema,
  from: count,
  to: count,
  change: z.number().int(),
});
export type DigestDriftTile = z.infer<typeof digestDriftTileSchema>;

export const digestDriftAuditSchema = z.object({
  auditId: auditIdSchema,
  title: text,
  from: digestRunRefSchema,
  to: digestRunRefSchema,
  differences: z.array(digestDifferenceSchema).max(3),
  tiles: z.array(digestDriftTileSchema).max(20),
});
export type DigestDriftAudit = z.infer<typeof digestDriftAuditSchema>;

export const digestAuditDriftSchema = z.object({
  ...section,
  // Audits with a result from before the range and a newer one: the ones that could be compared.
  compared: count,
  // Audits that have results but could not be compared.
  uncompared: count,
  // Only the audits with a tile that changed.
  audits: z.array(digestDriftAuditSchema).max(30),
});
export type DigestAuditDrift = z.infer<typeof digestAuditDriftSchema>;

export const digestPostSchema = z.object({
  id: z.uuid(),
  title: text,
  actBy: time,
  services: z.array(text).max(20),
});
export type DigestPost = z.infer<typeof digestPostSchema>;

export const digestIncidentSchema = z.object({
  id: z.uuid(),
  title: text,
  severity: attentionSeveritySchema,
  status: nullableText,
  services: z.array(text).max(20),
  firstSeen: time,
});
export type DigestIncident = z.infer<typeof digestIncidentSchema>;

export const digestMicrosoftSchema = z.object({
  ...section,
  // Message center posts with an act-by date in the next 30 days.
  posts: z.object({
    ...section,
    total: count,
    items: z.array(digestPostSchema).max(DIGEST_POSTS),
  }),
  // Service incidents opened or resolved in the range, and the ones still open.
  incidents: z.object({
    ...section,
    opened: count,
    resolved: count,
    openTotal: count,
    open: z.array(digestIncidentSchema).max(DIGEST_INCIDENTS),
  }),
});
export type DigestMicrosoft = z.infer<typeof digestMicrosoftSchema>;

export const digestCollectorSchema = z.object({
  module: z.string().max(80),
  moduleTitle: text,
  collector: z.string().max(80),
  collectorTitle: text,
  // When it last finished a run without errors, within the 30 days of run history Prism365 keeps.
  lastSuccess: time.nullable(),
  // Its newest run of any outcome.
  lastRun: z.object({ status: z.string().max(40), at: time }).nullable(),
});
export type DigestCollector = z.infer<typeof digestCollectorSchema>;

export const digestScheduledAuditSchema = z.object({
  auditId: auditIdSchema,
  title: text,
  frequency: auditScheduleFrequencySchema.exclude(["off"]),
  lastRun: z.object({ status: z.string().max(40), at: time }).nullable(),
});
export type DigestScheduledAudit = z.infer<typeof digestScheduledAuditSchema>;

export const digestCollectionSchema = z.object({
  ...section,
  collecting: z.boolean(),
  collectors: z.array(digestCollectorSchema).max(50),
  // Collector runs that failed in the range, and the categories of the newest failures.
  failedRuns: count,
  errors: z.array(z.string().max(80)).max(5),
  // Gaps the collectors recorded in the range: spans of the sign-in or audit log they could not
  // read.
  gaps: count,
  audits: z.array(digestScheduledAuditSchema).max(50),
});
export type DigestCollection = z.infer<typeof digestCollectionSchema>;

export const digestSchema = z.object({
  days: digestDaysSchema,
  // The range is from `from` (included) to `generatedAt` (not included).
  from: time,
  generatedAt: time,
  signInFindings: digestSignInFindingsSchema,
  directoryChanges: digestDirectoryChangesSchema,
  quietAccounts: digestQuietAccountsSchema,
  auditDrift: digestAuditDriftSchema,
  microsoft: digestMicrosoftSchema,
  collection: digestCollectionSchema,
});
export type Digest = z.infer<typeof digestSchema>;

// The sections, under the names the access ledger records for what a read touched.
export const DIGEST_SECTIONS = [
  "signInFindings",
  "directoryChanges",
  "quietAccounts",
  "auditDrift",
  "microsoft",
  "collection",
] as const;
export type DigestSectionName = (typeof DIGEST_SECTIONS)[number];

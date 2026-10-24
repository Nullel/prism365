import { z } from "zod";

// The directory change feed: stored directory audit events and the feed page. Values were cut
// and cleaned when Prism365 collected them; they are cut again here so one long value can never
// turn a page into an error. List lengths are bounded where rows are read (changeFromRow), not
// here, so a row from a build with other bounds cannot either.

const text = z.string().transform((value) => value.slice(0, 500));
const nullableText = text.nullable();
const time = z.string().datetime({ offset: true });
const count = z.number().int().nonnegative();

// The form the feed's search compares, applied to the stored text and to what a person types
// alike: compatibility-decomposed (NFKD) with the combining marks removed, lower-cased, and the
// Greek final sigma written as an ordinary sigma. "ilker" then finds "İlker", "elodie" finds
// "Élodie" and "ΟΔΥΣ" finds "ΟΔΥΣΣΕΥΣ", whatever the database's own locale would do.
export function searchKey(value: string): string {
  return value.normalize("NFKD").replace(/\p{M}/gu, "").toLowerCase().replaceAll("ς", "σ");
}

export const changeResultFilterSchema = z.enum(["success", "failure"]);
export type ChangeResultFilter = z.infer<typeof changeResultFilterSchema>;

export const changePropertySchema = z.object({
  name: nullableText,
  oldValue: nullableText,
  newValue: nullableText,
  // The value was longer than Prism365 keeps (400 characters) and was cut when collected.
  oldCut: z.boolean(),
  newCut: z.boolean(),
});
export type ChangeProperty = z.infer<typeof changePropertySchema>;

export const changeTargetSchema = z.object({
  id: nullableText,
  name: nullableText,
  type: nullableText,
  principal: nullableText,
  properties: z.array(changePropertySchema),
});
export type ChangeTarget = z.infer<typeof changeTargetSchema>;

export const changeEventSchema = z.object({
  id: z.string().min(1).max(200),
  activityAt: time,
  activity: nullableText,
  category: nullableText,
  operationType: nullableText,
  result: nullableText,
  resultReason: nullableText,
  loggedByService: nullableText,
  correlationId: nullableText,
  initiatedBy: z.object({
    kind: z.enum(["user", "app", "unknown"]),
    id: nullableText,
    name: nullableText,
    // A user's sign-in name, or an application's app ID.
    principal: nullableText,
  }),
  targets: z.array(changeTargetSchema),
  details: z.array(z.object({ key: nullableText, value: nullableText })),
});
export type ChangeEvent = z.infer<typeof changeEventSchema>;

// How far collection has read Microsoft's log. While the first load runs ("history"), changes
// from `from` to `through` are in and newer ones are still coming.
export const changeCoverageSchema = z.object({
  from: time.nullable(),
  through: time.nullable(),
  lastCollected: time.nullable(),
  phase: z.enum(["none", "history", "live"]),
});
export type ChangeCoverage = z.infer<typeof changeCoverageSchema>;

export const changeFeedSchema = z.object({
  days: z.number().int().min(1).max(3_650),
  items: z.array(changeEventSchema),
  // Pass back as `before` for the next, older page; null when there is none.
  next: z.string().max(600).nullable(),
  // Every change in the time range by category, whatever the other filters are.
  categories: z.array(z.object({ category: nullableText, count })),
  coverage: changeCoverageSchema,
});
export type ChangeFeed = z.infer<typeof changeFeedSchema>;

export const personChangeSchema = changeEventSchema.extend({
  // Whether the person made the change, had it made to them, or both.
  involvement: z.object({ actor: z.boolean(), target: z.boolean() }),
});
export type PersonChange = z.infer<typeof personChangeSchema>;

export const personChangePageSchema = z.object({
  items: z.array(personChangeSchema),
  next: z.string().max(600).nullable(),
});
export type PersonChangePage = z.infer<typeof personChangePageSchema>;

// Query strings the feed routes accept.
export const changeFeedQuerySchema = z.object({
  days: z.coerce.number().int().min(1).max(3_650).default(7),
  category: z.string().trim().min(1).max(100).optional(),
  result: changeResultFilterSchema.optional(),
  search: z.string().max(200).optional(),
  before: z.string().min(1).max(600).optional(),
  limit: z.coerce.number().int().min(1).max(200).default(100),
});

export const personChangeQuerySchema = z.object({
  before: z.string().min(1).max(600).optional(),
  limit: z.coerce.number().int().min(1).max(200).default(20),
});

import { z } from "zod";
import { auditIdSchema } from "./audits.ts";
import { accessAuditCellSchema, combinedRiskRuleSchema } from "./person-access.ts";

// What the stored audits know about one group, role, Conditional Access policy, managed device or
// consented app, in one answer (docs/audits.md, "Entity panel"). It is read from stored results
// only and asks Microsoft nothing. Every section stands alone and names the audit it comes from,
// so an audit that never ran, or that does not list this thing, says so and the rest still shows.

const text = z.string().max(500);
const time = z.string().datetime({ offset: true });
const count = z.number().int().nonnegative();

export const ENTITY_KINDS = ["group", "role", "policy", "device", "app"] as const;
export const entityKindSchema = z.enum(ENTITY_KINDS);
export type EntityKind = z.infer<typeof entityKindSchema>;

// Every id these five kinds have is a GUID: a directory object id, a role definition id, a policy
// id or an Intune device id. Lower case is the one spelling the stored audits keep.
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export const entityIdSchema = z
  .string()
  .regex(GUID)
  .transform((id) => id.toLowerCase());

export const entityParamsSchema = z.object({ kind: entityKindSchema, id: entityIdSchema });

// How many rows a section sends, and how many things a list names, before it says "and more".
export const MAX_ENTITY_ROWS = 25;
export const MAX_ENTITY_LIST = 20;
export const MAX_ENTITY_SEARCH = 20;
export const ENTITY_SEARCH_MIN = 2;
export const ENTITY_SEARCH_MAX = 100;

// A person, group, role, policy, device or app that another thing names. With an id it can be
// opened; without a name the page says "Unknown <kind>" and keeps the id behind a copy button.
export const entityMentionSchema = z.object({
  kind: z.enum(["person", ...ENTITY_KINDS, "other"]),
  id: z.string().max(200).nullable(),
  name: text.nullable(),
  // A person's sign-in name, when the stored audits have one.
  principal: text.nullable(),
});
export type EntityMention = z.infer<typeof entityMentionSchema>;

export const entityListSchema = z.object({
  label: text,
  items: z.array(entityMentionSchema).max(MAX_ENTITY_LIST),
  // How many more the source holds than the list names.
  more: count,
});
export type EntityList = z.infer<typeof entityListSchema>;

// One row of a source audit, as far as this panel shows it. `subject` is the thing the row is
// about when that is another thing (the holder of a role, the policy that names a group).
export const entityRowSchema = z.object({
  key: z.string().max(200),
  subject: entityMentionSchema.nullable(),
  // Short facts about why the row is here: "Includes it", "Leaves it out", "Held through it".
  tags: z.array(z.string().max(80)).max(10),
  cells: z.array(accessAuditCellSchema).max(40),
  lists: z.array(entityListSchema).max(8),
  // The Combined risks rules the subject trips in the latest run, when the subject is an account.
  risks: z.array(combinedRiskRuleSchema).max(20),
});
export type EntityRow = z.infer<typeof entityRowSchema>;

const sectionBase = {
  id: z.string().regex(/^[a-z][a-z-]{0,39}$/),
  title: text,
  auditId: auditIdSchema,
  auditName: text,
};

export const entitySectionSchema = z.discriminatedUnion("status", [
  z.object({
    status: z.literal("ok"),
    ...sectionBase,
    finishedAt: time.nullable(),
    // False when the run could not check everything.
    complete: z.boolean(),
    rows: z.array(entityRowSchema).max(MAX_ENTITY_ROWS),
    // The run held more matching rows than were sent.
    more: z.boolean(),
    // A plain line for a section with nothing in it that is still an answer.
    note: text.nullable(),
  }),
  // The audit has never finished a run.
  z.object({ status: z.literal("no_run"), ...sectionBase }),
  // The latest run was made before the audit kept the ids this section reads, so it cannot say
  // anything, not even that there is nothing. Running the audit again fixes it.
  z.object({ status: z.literal("outdated"), ...sectionBase, finishedAt: time.nullable() }),
  // The latest run says nothing about this thing.
  z.object({
    status: z.literal("not_in_run"),
    ...sectionBase,
    finishedAt: time.nullable(),
    complete: z.boolean(),
  }),
  z.object({ status: z.literal("failed"), ...sectionBase, message: text }),
]);
export type EntitySection = z.infer<typeof entitySectionSchema>;

export const entitySchema = z.object({
  kind: entityKindSchema,
  id: z.string().max(200),
  // null when no stored audit names it.
  name: text.nullable(),
  // The key facts, shown as tags under the name.
  tags: z.array(z.string().max(80)).max(8),
  // Plain lines about what this panel cannot say, such as why members are not listed.
  notes: z.array(text).max(5),
  sections: z.array(entitySectionSchema).max(10),
});
export type Entity = z.infer<typeof entitySchema>;

// The columns of each audit the panel shows for the thing itself, in the order they are sent. The
// route sends only these (and only the ones a row has a value for), so this is the one place that
// says what a panel shows. Hidden columns may be among them: a detail view is the place for what a
// table hides by default. The internal id columns, which only hold machine data, never are.
export const ENTITY_AUDIT_FIELDS: Readonly<Record<string, readonly string[]>> = {
  "groups.teams": [
    "kind",
    "status",
    "visibility",
    "owners",
    "enabledOwners",
    "created",
    "expires",
    "renewed",
    "roleAssignable",
    "synced",
    "lastActivity",
    "activeUsers",
    "guests",
    "activeChannels",
  ],
  "directory.privileged-roles": [
    "holderType",
    "role",
    "privileged",
    "held",
    "through",
    "scope",
    "starts",
    "ends",
    "enabled",
    "mfa",
  ],
  "identity.conditional-access": [
    "state",
    "appliesTo",
    "excludes",
    "apps",
    "conditions",
    "grant",
    "session",
    "modified",
    "created",
  ],
  "devices.managed": [
    "os",
    "osVersion",
    "compliance",
    "ownership",
    "lastCheckIn",
    "days",
    "encrypted",
    "jailbroken",
    "enrolled",
    "model",
    "manufacturer",
    "serial",
    "wifiMac",
  ],
  "apps.consents": [
    "publisher",
    "owner",
    "enabled",
    "highRisk",
    "appPermissions",
    "delegated",
    "consentedUsers",
    "userScopes",
    "appId",
  ],
  "apps.credentials": [
    "status",
    "audience",
    "created",
    "secrets",
    "certificates",
    "nextExpiry",
    "expiresIn",
    "expired",
    "oldestAge",
    "owners",
    // The owners themselves follow as links, so their names are not repeated as a field.
    "redirects",
    "appId",
  ],
};

export const entitySearchQuerySchema = z.object({
  q: z.string().trim().min(ENTITY_SEARCH_MIN).max(ENTITY_SEARCH_MAX),
});

export const entitySearchResultSchema = z.object({
  kind: entityKindSchema,
  id: z.string().max(200),
  name: text,
  // Kind of group, device operating system, publisher: one muted word or two.
  detail: z.string().max(100).nullable(),
});
export type EntitySearchResult = z.infer<typeof entitySearchResultSchema>;

export const entitySearchSchema = z.object({
  results: z.array(entitySearchResultSchema).max(MAX_ENTITY_SEARCH),
  // More things matched than were sent.
  more: z.boolean(),
  // At least one source could not be searched, so the list may be short.
  partial: z.boolean(),
});
export type EntitySearch = z.infer<typeof entitySearchSchema>;

// The Groups and Teams audit keeps each owner as one text item in its hidden `ownerList` cell,
// because a cell holds text lists and not records. One item is this record as JSON, in the order
// of the `ownerIds` cell, so the page can link an owner without pairing two lists.
export const OWNER_KINDS = ["user", "servicePrincipal", "other"] as const;
export const ownerEntrySchema = z.object({
  id: z.string().min(1).max(200),
  name: z.string().max(200).nullable(),
  kind: z.enum(OWNER_KINDS),
});
export type OwnerEntry = z.infer<typeof ownerEntrySchema>;

const CELL_ITEM_MAX = 500;

// One owner as a cell item. An item longer than a cell item may be would make the whole stored
// result unreadable, so the name gives way: JSON escaping can nearly double a name full of quotes.
export function encodeOwnerEntry(entry: OwnerEntry): string {
  let name = entry.name;
  let encoded = JSON.stringify({ id: entry.id, name, kind: entry.kind });
  while (name !== null && encoded.length > CELL_ITEM_MAX) {
    name = name.length > 20 ? name.slice(0, Math.floor(name.length / 2)) : null;
    encoded = JSON.stringify({ id: entry.id, name, kind: entry.kind });
  }
  return encoded.slice(0, CELL_ITEM_MAX);
}

// The owners a stored `ownerList` cell holds. An item that is not a valid record is left out
// rather than guessed at, and a cell that is not a list holds none.
export function decodeOwnerList(cell: unknown): OwnerEntry[] {
  if (!Array.isArray(cell)) return [];
  const owners: OwnerEntry[] = [];
  for (const item of cell.slice(0, 200)) {
    if (typeof item !== "string") continue;
    try {
      const parsed = ownerEntrySchema.safeParse(JSON.parse(item));
      if (parsed.success) owners.push(parsed.data);
    } catch {
      // Not JSON: not one of ours.
    }
  }
  return owners;
}

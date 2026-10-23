import { cleanText } from "@prism365/contracts";
import { z } from "zod";

// Microsoft Graph v1.0, verified 2026-10-22:
// https://learn.microsoft.com/en-us/graph/api/directoryaudit-list?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/resources/directoryaudit?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/resources/appidentity?view=graph-rest-1.0
// - GET /auditLogs/directoryAudits with AuditLog.Read.All, delegated or application. Delegated
//   callers also need Reports Reader, Security Reader or Security Administrator.
// - $filter supports eq, ge, le and startswith; activityDateTime supports eq, ge, le and $orderby.
//   `lt` is not documented, so windows are bounded with `ge` and `le` and the shared boundary
//   millisecond is read twice and de-duplicated by id.
// - $top is supported but its maximum is not documented (see PAGE_SIZES in collector.ts).
// - Microsoft keeps directory audits 7 days without Entra ID P1 or P2 and 30 days with it.
// - Target `type` appears as `Type` in older samples, so it is read case-insensitively.

export const directoryAuditContext = {
  feature: "audit.directory.read",
  permission: "AuditLog.Read.All",
} as const;

// Bounds on what one stored event may hold. Microsoft's own sizes are not documented, so these
// are Prism365's: enough for the feed and the rules, small enough that one event cannot crowd out
// the rest.
export const NAME_LIMIT = 256;
export const VALUE_LIMIT = 400;
export const TARGET_LIMIT = 50;
export const PROPERTY_LIMIT = 50;
export const DETAIL_LIMIT = 20;

// A directory id that names nobody. Microsoft fills it in for system actions.
const EMPTY_ID = "00000000-0000-0000-0000-000000000000";

// Cut to length before cleaning, never after (see cleanText).
const clip = (value: string, max: number) => cleanText(value.slice(0, max));

// Optional text tolerates any shape (it becomes null) and long values are cut.
const text = (max: number) =>
  z
    .string()
    .transform((value) => clip(value, max))
    .nullable()
    .optional()
    .catch(null);

// Property values are kept as Microsoft sent them until normalizeDirectoryAudit unquotes and
// cuts them, so it can tell whether a value was cut. The Graph client's response size budget
// bounds them until then. Anything that is not a string is left out.
const rawValue = z.string().nullable().optional().catch(null);

const userSchema = z
  .object({
    id: text(200),
    displayName: text(NAME_LIMIT),
    userPrincipalName: text(NAME_LIMIT),
    ipAddress: text(64),
  })
  .nullable()
  .optional()
  .catch(null);

const appSchema = z
  .object({
    appId: text(200),
    displayName: text(NAME_LIMIT),
    servicePrincipalId: text(200),
    servicePrincipalName: text(NAME_LIMIT),
  })
  .nullable()
  .optional()
  .catch(null);

const propertySchema = z.object({
  displayName: text(NAME_LIMIT),
  oldValue: rawValue,
  newValue: rawValue,
});

const targetSchema = z.preprocess(
  // Older records spell the field `Type`.
  (value) =>
    value && typeof value === "object" && !("type" in value) && "Type" in value
      ? { ...value, type: (value as { Type: unknown }).Type }
      : value,
  z.object({
    id: text(200),
    displayName: text(NAME_LIMIT),
    type: text(64),
    userPrincipalName: text(NAME_LIMIT),
    modifiedProperties: z.array(z.unknown()).nullable().optional().catch(null),
  }),
);

const detailSchema = z.object({ key: text(NAME_LIMIT), value: rawValue });

// One record as Graph returns it. Identity and time are strict: a record without them cannot be
// stored or ordered, so it is left out and counted. Everything else is tolerant, and unknown
// fields are dropped: the module never stores more of a record than it shows or judges.
export const graphDirectoryAuditSchema = z.object({
  id: z
    .string()
    .max(200)
    .transform((value) => cleanText(value).trim())
    .pipe(z.string().min(1)),
  activityDateTime: z.string().datetime({ offset: true }),
  activityDisplayName: text(NAME_LIMIT),
  category: text(NAME_LIMIT),
  correlationId: text(200),
  loggedByService: text(NAME_LIMIT),
  operationType: text(NAME_LIMIT),
  result: text(40),
  resultReason: text(VALUE_LIMIT),
  initiatedBy: z.object({ user: userSchema, app: appSchema }).nullable().optional().catch(null),
  // Arrays are cut before their elements are read, so a record with thousands of targets costs
  // no more than one with fifty.
  targetResources: z.array(z.unknown()).nullable().optional().catch(null),
  additionalDetails: z.array(z.unknown()).nullable().optional().catch(null),
});

export type GraphDirectoryAudit = z.infer<typeof graphDirectoryAuditSchema>;

export interface ChangeProperty {
  name: string | null;
  oldValue: string | null;
  newValue: string | null;
  // Set when the value was longer than VALUE_LIMIT and was cut.
  oldCut?: true;
  newCut?: true;
}

export interface ChangeTarget {
  id: string | null;
  name: string | null;
  type: string | null;
  principal: string | null;
  properties: ChangeProperty[];
}

export interface ChangeInitiator {
  kind: "user" | "app" | "unknown";
  id: string | null;
  name: string | null;
  // A user's sign-in name, or an application's app ID.
  principal: string | null;
}

export interface DirectoryChangeEvent {
  id: string;
  activityAt: Date;
  activity: string | null;
  category: string | null;
  operationType: string | null;
  result: string | null;
  resultReason: string | null;
  loggedByService: string | null;
  correlationId: string | null;
  initiatedBy: ChangeInitiator;
  targets: ChangeTarget[];
  details: { key: string | null; value: string | null }[];
}

const blank = (value: string | null | undefined) => {
  const trimmed = value?.trim();
  return trimmed ? trimmed : null;
};

const identifier = (value: string | null | undefined) => {
  const id = blank(value);
  return id && id !== EMPTY_ID ? id : null;
};

// Microsoft writes most property values as JSON: a quoted string ("\"Sales\""), or an array of
// plain values ("[true]", "[\"a\",\"b\"]"). Those are shown as the text they hold. Anything
// else, such as a list of objects, is kept as Microsoft wrote it.
export function unquote(value: string): string {
  const trimmed = value.trim();
  if (!trimmed.startsWith('"') && !trimmed.startsWith("[")) return value;
  try {
    const parsed: unknown = JSON.parse(trimmed);
    if (typeof parsed === "string") return parsed;
    if (
      Array.isArray(parsed) &&
      parsed.every((item) => ["string", "number", "boolean"].includes(typeof item))
    )
      return parsed.map(String).join(", ");
  } catch {
    // Not JSON: shown as written.
  }
  return value;
}

function boundedValue(value: string | null | undefined): { value: string | null; cut: boolean } {
  if (value === null || value === undefined) return { value: null, cut: false };
  const plain = unquote(value);
  return { value: clip(plain, VALUE_LIMIT), cut: plain.length > VALUE_LIMIT };
}

function normalizeProperty(raw: unknown): ChangeProperty | null {
  const parsed = propertySchema.safeParse(raw);
  if (!parsed.success) return null;
  const before = boundedValue(parsed.data.oldValue);
  const after = boundedValue(parsed.data.newValue);
  return {
    name: blank(parsed.data.displayName),
    oldValue: before.value,
    newValue: after.value,
    ...(before.cut ? { oldCut: true as const } : {}),
    ...(after.cut ? { newCut: true as const } : {}),
  };
}

function normalizeTarget(raw: unknown): ChangeTarget | null {
  const parsed = targetSchema.safeParse(raw);
  if (!parsed.success) return null;
  const target = parsed.data;
  return {
    id: identifier(target.id),
    name: blank(target.displayName),
    type: blank(target.type),
    principal: blank(target.userPrincipalName),
    properties: (target.modifiedProperties ?? [])
      .slice(0, PROPERTY_LIMIT)
      .map(normalizeProperty)
      .filter((property): property is ChangeProperty => property !== null),
  };
}

function initiator(record: GraphDirectoryAudit): ChangeInitiator {
  const user = record.initiatedBy?.user;
  const userId = identifier(user?.id);
  if (user && (userId || blank(user.userPrincipalName) || blank(user.displayName)))
    return {
      kind: "user",
      id: userId,
      name: blank(user.displayName),
      principal: blank(user.userPrincipalName),
    };
  const app = record.initiatedBy?.app;
  const appId = identifier(app?.appId);
  const servicePrincipal = identifier(app?.servicePrincipalId);
  const appName = blank(app?.displayName) ?? blank(app?.servicePrincipalName);
  if (app && (appId || servicePrincipal || appName))
    return { kind: "app", id: servicePrincipal ?? appId, name: appName, principal: appId };
  return { kind: "unknown", id: null, name: null, principal: null };
}

// A bounded event: what the feed shows and the rules judge. The time keeps millisecond
// precision, the same as the feed's page positions, so paging never skips a record that shares
// a millisecond with the last one shown.
export function normalizeDirectoryAudit(record: GraphDirectoryAudit): DirectoryChangeEvent {
  return {
    id: record.id,
    activityAt: new Date(record.activityDateTime),
    activity: blank(record.activityDisplayName),
    category: blank(record.category),
    operationType: blank(record.operationType),
    result: blank(record.result)?.toLowerCase() ?? null,
    resultReason: blank(record.resultReason),
    loggedByService: blank(record.loggedByService),
    correlationId: blank(record.correlationId),
    initiatedBy: initiator(record),
    targets: (record.targetResources ?? [])
      .slice(0, TARGET_LIMIT)
      .map(normalizeTarget)
      .filter((target): target is ChangeTarget => target !== null),
    details: (record.additionalDetails ?? [])
      .slice(0, DETAIL_LIMIT)
      .map((raw) => detailSchema.safeParse(raw))
      .filter((parsed) => parsed.success)
      .map((parsed) => ({
        key: blank(parsed.data.key),
        value: boundedValue(parsed.data.value).value,
      })),
  };
}

// Reads one record. null when it cannot be stored: no id, no readable time.
export function readDirectoryAudit(raw: unknown): DirectoryChangeEvent | null {
  const parsed = graphDirectoryAuditSchema.safeParse(raw);
  if (!parsed.success) return null;
  const event = normalizeDirectoryAudit(parsed.data);
  return Number.isFinite(event.activityAt.getTime()) ? event : null;
}

// Accounts and objects an event is about, for "changes involving this person". Ids are compared
// without case: Microsoft writes directory GUIDs in lower case, but nothing promises it.
export interface ChangeSubject {
  id: string;
  kind: "actor" | "target";
  type: string | null;
}

export function changeSubjects(event: DirectoryChangeEvent): ChangeSubject[] {
  const subjects = new Map<string, ChangeSubject>();
  const add = (id: string | null, kind: ChangeSubject["kind"], type: string | null) => {
    if (!id) return;
    const key = `${kind}|${id.toLowerCase()}`;
    if (!subjects.has(key))
      subjects.set(key, { id: id.toLowerCase(), kind, type: type?.toLowerCase() ?? null });
  };
  if (event.initiatedBy.kind !== "unknown")
    add(event.initiatedBy.id, "actor", event.initiatedBy.kind);
  for (const target of event.targets) add(target.id, "target", target.type);
  return [...subjects.values()];
}

import { z } from "zod";
import {
  type AuditCell,
  type AuditDefinition,
  type AuditQuery,
  type AuditRow,
  auditIdSchema,
} from "./audits.ts";
import { personIdSchema } from "./people.ts";
import { cleanText } from "./text.ts";

// Handoffs and receipts (docs/handoff.md). Prism365 sees; a separate action tool acts. They meet
// through files: Prism365 exports a handoff (the accounts in an audit's view, with the reasons
// and evidence), the tool sends back a receipt (what it did to whom), and Prism365 keeps the
// receipt and shows it as the tool's report, never as something Prism365 checked. Both formats
// are versioned and strict: an unknown field anywhere means a document for another version, so
// it is refused rather than half read.

export const HANDOFF_FORMAT = "prism365.handoff";
export const RECEIPT_FORMAT = "prism365.receipt";
export const MAX_HANDOFF_ITEMS = 20_000;
export const MAX_HANDOFF_REASONS = 20;
export const MAX_HANDOFF_EVIDENCE = 40;
export const MAX_RECEIPT_ITEMS = 20_000;
// The import route's body limit. A larger receipt is split into several, each with its own id.
export const MAX_RECEIPT_BYTES = 4 * 1024 * 1024;

const time = z.string().datetime({ offset: true });
const count = z.number().int().nonnegative();
const label = z.string().min(1).max(500);
const viewId = z.string().regex(/^[a-z0-9][a-z0-9-]{0,39}$/);
const columnKey = z.string().regex(/^[a-zA-Z][a-zA-Z0-9]{0,39}$/);
const direction = z.enum(["asc", "desc"]);
// Ids from another tool: any GUID PostgreSQL accepts, in lower case, so one id written in
// either case is one receipt.
const guid = z.guid().transform((value) => value.toLowerCase());
const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// A directory object id in the form Prism365 stores it: a GUID in lower case, anything else as
// given. Graph returns lower case; a tool may not.
export function normalizeUserId(value: string): string {
  return guidPattern.test(value) ? value.toLowerCase() : value;
}

// ---------------------------------------------------------------------------------------------
// Handoff v1: written by Prism365
// ---------------------------------------------------------------------------------------------

export const handoffEvidenceValueSchema = z.union([
  z.string().max(2_000),
  z.number(),
  z.boolean(),
  z.null(),
]);
export type HandoffEvidenceValue = z.infer<typeof handoffEvidenceValueSchema>;

export const handoffItemSchema = z.strictObject({
  userId: personIdSchema,
  userPrincipalName: z.string().max(500).nullable(),
  displayName: z.string().max(500).nullable(),
  // The labels of the audit's views the account is in, such as "90+ days".
  reasons: z.array(label).max(MAX_HANDOFF_REASONS),
  // The visible columns of the row, by their label on the page.
  evidence: z
    .record(label, handoffEvidenceValueSchema)
    .refine((value) => Object.keys(value).length <= MAX_HANDOFF_EVIDENCE, {
      message: `At most ${MAX_HANDOFF_EVIDENCE} evidence values.`,
    }),
});
export type HandoffItem = z.infer<typeof handoffItemSchema>;

export const handoffSourceSchema = z.strictObject({
  kind: z.literal("audit"),
  auditId: auditIdSchema,
  auditTitle: label,
  // The version of the audit that produced the rows.
  auditVersion: z.number().int().positive(),
  runId: z.uuid(),
  runFinishedAt: time,
  // The account domains the run was limited to, as the audit scope said when it ran; null when
  // it covered every domain. A limited list is not the whole tenant.
  scope: z.array(z.string().min(1).max(253)).min(1).max(100).nullable(),
  // false when Microsoft returned less than everything to the run: the list may miss accounts.
  complete: z.boolean(),
  // The view the rows were taken from.
  view: z.strictObject({ id: viewId, label }),
  search: z.string().max(200),
  facets: z
    .record(columnKey, z.string().min(1).max(200))
    .refine((value) => Object.keys(value).length <= 30, { message: "At most 30 filters." }),
  // null when the export named a column the audit does not have: the rows keep the audit's order.
  sort: z.strictObject({ column: columnKey, direction }).nullable(),
});
export type HandoffSource = z.infer<typeof handoffSourceSchema>;

export const handoffSchema = z.strictObject({
  format: z.literal(HANDOFF_FORMAT),
  version: z.literal(1),
  id: z.uuid(),
  generatedAt: time,
  tenantId: z.uuid(),
  generatedBy: z.strictObject({ actorId: z.uuid(), name: z.string().max(500).nullable() }),
  source: handoffSourceSchema,
  items: z.array(handoffItemSchema).max(MAX_HANDOFF_ITEMS),
});
export type Handoff = z.infer<typeof handoffSchema>;

function isPerson(
  cell: AuditCell | undefined,
): cell is { name: string | null; principal: string | null } {
  return typeof cell === "object" && cell !== null && !Array.isArray(cell);
}

// Stored text as it goes into a handoff: cut to its bound first, then cleaned, so a cut that
// splits a surrogate pair is repaired (see cleanText).
function bounded(value: string | null | undefined, max: number): string | null {
  if (typeof value !== "string") return null;
  return cleanText(value.slice(0, max));
}

// What an export applied, resolved exactly as queryAuditRows resolves it, so the document says
// which rows it holds: the view (the default when none was named), the filters on filterable
// columns, the trimmed search and the sort.
export function handoffSelection(
  definition: AuditDefinition,
  query: AuditQuery,
): Pick<HandoffSource, "view" | "search" | "facets" | "sort"> {
  const view = handoffView(definition, query.view ?? null);
  if (!view) throw new RangeError("The audit has no such view.");
  const facets = Object.fromEntries(
    Object.entries(query.facets ?? {}).filter(
      ([key, value]) =>
        value && definition.columns.some((column) => column.key === key && column.facet),
    ),
  );
  const sort = query.sort ?? definition.defaultSort;
  return {
    view: { id: view.id, label: view.label },
    search: (query.search ?? "").trim(),
    facets,
    sort: definition.columns.some((column) => column.key === sort.column)
      ? { column: sort.column, direction: sort.direction }
      : null,
  };
}

// The view an export names, or the audit's default view when it names none; null for a view the
// audit does not have. queryAuditRows selects every row for an unknown view, so an export must
// refuse one rather than hand off accounts nobody chose.
export function handoffView(
  definition: AuditDefinition,
  view: string | null,
): AuditDefinition["views"][number] | null {
  return definition.views.find((item) => item.id === (view ?? definition.defaultView)) ?? null;
}

// The columns whose values go into each item's evidence: the ones named (the columns the page
// shows), or the ones the audit shows by default. Person columns are left out: the account is
// the item itself.
export function handoffColumns(
  definition: AuditDefinition,
  shown: readonly string[] | null,
): AuditDefinition["columns"] {
  const named = shown ? new Set(shown) : null;
  return definition.columns
    .filter((column) => column.kind !== "person")
    .filter((column) => (named ? named.has(column.key) : !column.hidden))
    .slice(0, MAX_HANDOFF_EVIDENCE);
}

function evidenceValue(cell: AuditCell | undefined): HandoffEvidenceValue {
  if (cell === null || cell === undefined) return null;
  if (typeof cell === "boolean") return cell;
  if (typeof cell === "number") return Number.isFinite(cell) ? cell : null;
  if (typeof cell === "string") return bounded(cell, 2_000);
  // A list as the page reads it, one value after another; an empty one is no value.
  if (Array.isArray(cell)) return cell.length > 0 ? bounded(cell.join("; "), 2_000) : null;
  return bounded(cell.name ?? cell.principal, 2_000);
}

// One item per row that is an account, in the order of the rows. Reasons are the labels of the
// single-tag views the row is in, in the order the audit lists its views; a tag no view names on
// its own (an account's state, its type) is an attribute rather than a reason, and the evidence
// carries it instead.
export function handoffItems(
  definition: AuditDefinition,
  rows: readonly AuditRow[],
  shown: readonly string[] | null,
): HandoffItem[] {
  const person = definition.columns.find((column) => column.kind === "person");
  if (!person) return [];
  const columns = handoffColumns(definition, shown);
  const keys = new Map<string, string>();
  for (const column of columns) {
    // Two columns with one label keep both values.
    const taken = [...keys.values()].includes(column.label);
    keys.set(column.key, taken ? `${column.label} (${column.key})` : column.label);
  }
  const reasonViews = definition.views.filter((view) => view.tags.length === 1);
  const items: HandoffItem[] = [];
  for (const row of rows) {
    if (items.length >= MAX_HANDOFF_ITEMS) break;
    if (!row.personId || !personIdSchema.safeParse(row.personId).success) continue;
    const cell = row.cells[person.key];
    const account = isPerson(cell) ? cell : { name: null, principal: null };
    const reasons = [
      ...new Set(
        reasonViews
          .filter((view) => view.tags.every((tag) => row.tags.includes(tag)))
          .map((view) => bounded(view.label, 500) ?? ""),
      ),
    ]
      .filter(Boolean)
      .slice(0, MAX_HANDOFF_REASONS);
    const evidence: Record<string, HandoffEvidenceValue> = {};
    for (const column of columns)
      evidence[keys.get(column.key) ?? column.label] = evidenceValue(row.cells[column.key]);
    items.push({
      userId: row.personId,
      userPrincipalName: bounded(account.principal, 500),
      displayName: bounded(account.name, 500),
      reasons,
      evidence,
    });
  }
  return items;
}

export function hasPersonColumn(definition: AuditDefinition): boolean {
  return definition.columns.some((column) => column.kind === "person");
}

// The whole document for one export. `rows` are the rows the export selected, in order.
export function buildHandoff(input: {
  id: string;
  generatedAt: string;
  tenantId: string;
  generatedBy: { actorId: string; name: string | null };
  definition: AuditDefinition;
  run: {
    id: string;
    auditVersion: number;
    startedAt: string;
    finishedAt: string | null;
    complete: boolean;
    scope?: readonly string[] | null;
  };
  query: AuditQuery;
  rows: readonly AuditRow[];
  // Column keys the page shows; null for the audit's default columns.
  columns: readonly string[] | null;
}): Handoff {
  const { definition, run } = input;
  return handoffSchema.parse({
    format: HANDOFF_FORMAT,
    version: 1,
    id: input.id,
    generatedAt: input.generatedAt,
    tenantId: input.tenantId,
    generatedBy: {
      actorId: input.generatedBy.actorId,
      name: bounded(input.generatedBy.name, 500),
    },
    source: {
      kind: "audit",
      auditId: definition.id,
      auditTitle: bounded(definition.title, 500),
      auditVersion: run.auditVersion,
      runId: run.id,
      runFinishedAt: run.finishedAt ?? run.startedAt,
      scope: run.scope && run.scope.length > 0 ? [...run.scope] : null,
      complete: run.complete,
      ...handoffSelection(definition, input.query),
    },
    items: handoffItems(definition, input.rows, input.columns),
  });
}

// "-limited-to-2-domains": a file of a run limited to some account domains says so in its name,
// so a file passed on is never mistaken for the whole tenant. The CSV export uses it too.
export function scopeFileSuffix(scope: readonly string[] | null | undefined): string {
  if (!scope || scope.length === 0) return "";
  return `-limited-to-${scope.length}-domain${scope.length === 1 ? "" : "s"}`;
}

// "prism365-handoff-accounts.inactive-2026-10-22.json": the audit and the day of its run, as the
// CSV export names its file, and the same suffix for a limited run.
export function handoffFileName(
  auditId: string,
  runStartedAt: string,
  scope: readonly string[] | null = null,
): string {
  return `prism365-handoff-${auditId}-${runStartedAt.slice(0, 10)}${scopeFileSuffix(scope)}.json`;
}

// ---------------------------------------------------------------------------------------------
// Receipt v1: written by the action tool
// ---------------------------------------------------------------------------------------------

// Characters that reorder how the text around them is displayed. In a tool's name or a detail
// they could make one thing read as another, and plain text never needs them.
const BIDI_CONTROLS = /[\u202A-\u202E\u2066-\u2069\u200E\u200F]/g;

// Text from the tool as it is kept: NUL removed and broken UTF-16 replaced (cleanText), and
// direction controls removed.
export function cleanReceiptText(value: string): string {
  return cleanText(value).replace(BIDI_CONTROLS, "");
}

// Text from the tool: trimmed, bounded, then cleaned. The bound is checked on the text as sent,
// never cut, so nothing the tool wrote is silently shortened.
const receiptText = (min: number, max: number) =>
  z
    .string()
    .trim()
    .max(max)
    .transform((value) => cleanReceiptText(value).trim())
    .pipe(z.string().min(min));

// A principal name as Prism365 matches it: compatibility-normalised (NFKC) and in lower case, in
// JavaScript, so "K" and the Kelvin sign are one letter whatever the database's locale.
export function principalKey(value: string): string {
  return value.normalize("NFKC").toLowerCase();
}

const reserved = (value: string) => value.normalize("NFKC").toLowerCase();

// The day a receipt may be ahead of the server's clock: two clocks never agree exactly, but a
// receipt dated further ahead would sort first on every page for as long as it is kept.
export const RECEIPT_CLOCK_SKEW_MS = 86_400_000;

// When the tool did the work, kept as UTC. PostgreSQL refuses time zone offsets beyond 15:59
// (the format allows 23:59) and some dates the format allows (year 0), and it moves others past
// year 9999, which no ISO timestamp can show; each would fail only once stored.
const performedAt = time
  .refine(
    (value) => {
      const offset = /([+-])(\d{2}):(\d{2})$/.exec(value);
      return !offset || Number(offset[2]) <= 15;
    },
    { message: "has a time zone offset beyond 15:59, which no time zone uses" },
  )
  .refine(
    (value) => {
      const at = Date.parse(value);
      return at >= Date.UTC(2000, 0, 1) && at < Date.UTC(2100, 0, 1);
    },
    { message: "is not between the years 2000 and 2099" },
  )
  .refine((value) => Date.parse(value) <= Date.now() + RECEIPT_CLOCK_SKEW_MS, {
    message: "is more than a day in the future. Check the clock of the computer the tool ran on",
  })
  .transform((value) => new Date(value).toISOString());

export const receiptActionSchema = z
  .string()
  .regex(/^[a-z][a-z0-9.-]{1,79}$/)
  .refine((value) => !value.startsWith("prism365."), {
    message: "starts with prism365., which is kept for Prism365's own names",
  });
export const receiptStatusSchema = z.enum(["done", "failed", "skipped"]);
export type ReceiptStatus = z.infer<typeof receiptStatusSchema>;

export const receiptItemSchema = z
  .strictObject({
    userId: personIdSchema.transform(normalizeUserId).nullable(),
    userPrincipalName: receiptText(1, 500)
      .refine((value) => principalKey(value).length <= 2_000, {
        message: "is not a principal name",
      })
      .nullable(),
    action: receiptActionSchema,
    status: receiptStatusSchema,
    detail: z.string().max(500).transform(cleanReceiptText),
  })
  .refine((item) => item.userId !== null || item.userPrincipalName !== null, {
    message: "needs a userId or a userPrincipalName",
    path: ["userId"],
  });
export type ReceiptItem = z.infer<typeof receiptItemSchema>;

export const receiptSchema = z.strictObject({
  format: z.literal(RECEIPT_FORMAT),
  version: z.literal(1),
  // Chosen by the tool, and the key for idempotency: the same receipt sent twice is kept once.
  id: guid,
  // The handoff this receipt answers, or null for actions taken without one.
  handoffId: guid.nullable(),
  tool: z.strictObject({
    // Another tool's report can never pass for Prism365's own.
    name: receiptText(1, 80).refine((value) => reserved(value) !== "prism365", {
      message: "cannot be Prism365: a receipt reports what another tool did",
    }),
    version: receiptText(1, 40),
  }),
  performedAt,
  // The operator, as the tool names them.
  performedBy: receiptText(1, 200),
  items: z.array(receiptItemSchema).min(1).max(MAX_RECEIPT_ITEMS),
});
export type Receipt = z.infer<typeof receiptSchema>;

// A value as one string that does not depend on key order or spacing: object keys sorted, arrays
// in order, no whitespace. The SHA-256 of a validated receipt in this form is its digest.
export function canonicalJson(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(canonicalJson).join(",")}]`;
  if (value !== null && typeof value === "object") {
    const object = value as Record<string, unknown>;
    return `{${Object.keys(object)
      .filter((key) => object[key] !== undefined)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${canonicalJson(object[key])}`)
      .join(",")}}`;
  }
  return JSON.stringify(value) ?? "null";
}

function pathText(path: readonly PropertyKey[]): string {
  let text = "";
  for (const part of path)
    text += typeof part === "number" ? `[${part}]` : `${text ? "." : ""}${String(part)}`;
  return text;
}

function valueAt(input: unknown, path: readonly PropertyKey[]): unknown {
  let current = input;
  for (const part of path) {
    if (current === null || typeof current !== "object") return undefined;
    current = (current as Record<PropertyKey, unknown>)[part];
  }
  return current;
}

const formatHints: Record<string, string> = {
  guid: "is not a UUID",
  uuid: "is not a UUID",
  datetime: "is not a date and time with a time zone, such as 2026-10-22T14:30:00Z",
};
const nullableFields = new Set(["userId", "userPrincipalName", "handoffId"]);

// The first problem with a receipt, in plain words, naming where it is ("items[3].action").
export function receiptProblem(input: unknown, error: z.ZodError): string {
  const issue = error.issues[0];
  if (!issue) return "The receipt is not valid.";
  const path = issue.path;
  const where = path.length > 0 ? pathText(path) : "The receipt";
  const field = String(path.at(-1) ?? "");
  const value = valueAt(input, path);
  if (path.length === 0 && issue.code === "invalid_type")
    return "The receipt is not a JSON object. Choose a receipt file written by the action tool.";
  if (issue.code === "invalid_type" && value === undefined) return `${where} is missing.`;
  if (issue.code === "invalid_type")
    return `${where} should be a ${issue.expected}${nullableFields.has(field) ? " or null" : ""}.`;
  // The tool's own text is echoed only short, so a message never grows with what was sent.
  if (issue.code === "unrecognized_keys")
    return `${where} has a field Prism365 does not read: "${cleanText((issue.keys[0] ?? "").slice(0, 80))}". Check the receipt's format and version.`;
  if (issue.code === "too_big")
    return issue.origin === "array"
      ? `${where} has more than ${issue.maximum} entries.`
      : `${where} is longer than ${issue.maximum} characters.`;
  if (issue.code === "too_small")
    return issue.origin === "array"
      ? `${where} is empty. A receipt reports at least one action.`
      : `${where} is empty.`;
  if (issue.code === "invalid_value") {
    if (field === "format")
      return `This is not a Prism365 receipt: format should be "${RECEIPT_FORMAT}".`;
    if (field === "version")
      return typeof value === "number" && Number.isFinite(value)
        ? `This receipt is version ${value}. This Prism365 reads version 1.`
        : "This receipt has no version Prism365 reads. This Prism365 reads version 1.";
    if (field === "status") return `${where} should be done, failed or skipped.`;
    return `${where} has a value Prism365 does not read.`;
  }
  if (issue.code === "invalid_format") {
    if (field === "action")
      return `${where} is not an action name such as account.sign-in-blocked: 2 to 80 lower-case letters, digits, dots and dashes, starting with a letter.`;
    if (field === "userId") return `${where} is not a directory object id.`;
    return `${where} ${formatHints[issue.format] ?? "is not in the expected form"}.`;
  }
  if (issue.code === "custom" && field === "userId")
    return `${pathText(path.slice(0, -1))} needs a userId or a userPrincipalName.`;
  // The schema's own refinements are written to follow the field's name.
  if (issue.code === "custom") return `${where} ${issue.message}.`;
  return `${where} is not valid.`;
}

// "account.sign-in-blocked" as "Account sign in blocked".
export function actionLabel(action: string): string {
  const words = action.replaceAll(/[.-]+/g, " ").trim();
  return words ? `${words.charAt(0).toUpperCase()}${words.slice(1)}` : action;
}

// ---------------------------------------------------------------------------------------------
// API responses
// ---------------------------------------------------------------------------------------------

export const handoffRecordSchema = z.object({
  id: z.uuid(),
  auditId: auditIdSchema,
  auditTitle: z.string().max(500),
  runId: z.uuid(),
  view: z.object({ id: z.string().max(40), label: z.string().max(500) }).nullable(),
  // A search or a filter narrowed the view.
  filtered: z.boolean(),
  // The run was limited to these account domains; null when it covered every one.
  scope: z.array(z.string().max(253)).max(100).nullable(),
  // false when the run was partial.
  complete: z.boolean(),
  items: count,
  generatedAt: time,
  generatedBy: z.object({ id: z.uuid(), name: z.string().max(500).nullable() }),
});
export type HandoffRecord = z.infer<typeof handoffRecordSchema>;

export const receiptSummarySchema = z.object({
  id: z.guid(),
  // The handoff this receipt answers, as Prism365 recorded it; null when it answers none.
  handoff: z
    .object({ id: z.uuid(), auditId: auditIdSchema, generatedAt: time, items: count })
    .nullable(),
  tool: z.object({ name: z.string().max(80), version: z.string().max(40) }),
  performedAt: time,
  performedBy: z.string().max(200),
  importedAt: time,
  items: count,
  done: count,
  failed: count,
  skipped: count,
});
export type ReceiptSummary = z.infer<typeof receiptSummarySchema>;

export const handoffListSchema = z.object({
  handoffs: z.array(handoffRecordSchema).max(100),
  receipts: z.array(receiptSummarySchema).max(100),
});
export type HandoffList = z.infer<typeof handoffListSchema>;

// The answer to an import: imported is false when this exact receipt was already kept.
export const receiptImportSchema = z.object({
  receipt: receiptSummarySchema,
  imported: z.boolean(),
});
export type ReceiptImport = z.infer<typeof receiptImportSchema>;

export const storedReceiptItemSchema = z.object({
  position: count,
  userId: z.string().max(200).nullable(),
  userPrincipalName: z.string().max(500).nullable(),
  action: z.string().max(80),
  status: receiptStatusSchema,
  detail: z.string().max(500),
});
export type StoredReceiptItem = z.infer<typeof storedReceiptItemSchema>;

export const receiptItemPageSchema = z.object({
  receipt: receiptSummarySchema,
  items: z.array(storedReceiptItemSchema).max(500),
  // The position to continue after, or null on the last page.
  next: count.nullable(),
});
export type ReceiptItemPage = z.infer<typeof receiptItemPageSchema>;

export const personReceiptItemSchema = storedReceiptItemSchema.extend({
  receiptId: z.guid(),
  tool: z.object({ name: z.string().max(80), version: z.string().max(40) }),
  performedAt: time,
  performedBy: z.string().max(200),
  importedAt: time,
  // How the item was matched to this person: by the account's id, or by its principal name when
  // the tool gave only that.
  matchedBy: z.enum(["userId", "userPrincipalName"]),
});
export type PersonReceiptItem = z.infer<typeof personReceiptItemSchema>;

export const personReceiptListSchema = z.object({
  items: z.array(personReceiptItemSchema).max(100),
  // More than 100 matched; the newest 100 are shown.
  truncated: z.boolean(),
});
export type PersonReceiptList = z.infer<typeof personReceiptListSchema>;

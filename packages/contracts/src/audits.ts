import { z } from "zod";
import { personIdSchema } from "./people.ts";

// Audits: checks an administrator runs on demand against the live tenant. Each audit describes
// its own table (columns, preset views, a default sort), so the page, the CSV export and the
// API work for any audit without knowing what it checks. Rows carry tags, and a view is the set
// of tags a row must have.

const text = z.string().transform((value) => value.slice(0, 500));
const time = z.string().datetime({ offset: true });
const slug = z.string().regex(/^[a-z0-9][a-z0-9-]{0,39}$/);
const columnKey = z.string().regex(/^[a-zA-Z][a-zA-Z0-9]{0,39}$/);

export const auditIdSchema = z
  .string()
  .max(80)
  .regex(/^[a-z][a-z0-9-]*(?:\.[a-z][a-z0-9-]*)+$/);

export const auditColumnKindSchema = z.enum([
  // { name, principal }, linked to the person's page.
  "person",
  "text",
  // A short value from a fixed set; can be offered as a filter.
  "enum",
  // An ISO timestamp, or null for never.
  "date",
  // A whole number of days.
  "days",
  "number",
  "boolean",
  "list",
]);
export type AuditColumnKind = z.infer<typeof auditColumnKindSchema>;

export const auditColumnSchema = z.object({
  key: columnKey,
  label: text,
  kind: auditColumnKindSchema,
  // Offered as a filter above the table. Only meaningful for enum columns.
  facet: z.boolean().default(false),
  // Left out of the table unless the administrator shows it. Always exported.
  hidden: z.boolean().default(false),
  // How the value is worked out, shown on the column header.
  help: text.nullable().default(null),
  // What an empty value means here ("None recorded", "Unknown"). Defaults by kind.
  empty: text.nullable().default(null),
});
export type AuditColumn = z.infer<typeof auditColumnSchema>;

export const auditViewSchema = z.object({
  id: slug,
  label: text,
  // A row is in the view when it has every one of these tags. No tags means every row.
  tags: z.array(slug).max(8),
});
export type AuditView = z.infer<typeof auditViewSchema>;

export const auditDefinitionSchema = z.object({
  id: auditIdSchema,
  version: z.number().int().positive(),
  title: text,
  description: text,
  category: text,
  // What one row is, for counts: "account", "accounts".
  subject: z.object({ one: text, many: text }),
  columns: z.array(auditColumnSchema).min(1).max(30),
  views: z.array(auditViewSchema).min(1).max(20),
  defaultView: slug,
  defaultSort: z.object({ column: columnKey, direction: z.enum(["asc", "desc"]) }),
  // Microsoft permissions the audit cannot run without, and ones that only add detail.
  permissions: z.array(text).max(10),
  optionalPermissions: z.array(text).max(10),
  // Roles and licenses Microsoft also requires, in plain words.
  requirements: z.array(text).max(10),
  // The audit only covers the account domains the administrator chose (see auditScopeSchema).
  scoped: z.boolean().default(false),
});
export type AuditDefinition = z.infer<typeof auditDefinitionSchema>;
export type AuditDefinitionInput = z.input<typeof auditDefinitionSchema>;

export const auditPersonCellSchema = z.object({
  name: text.nullable(),
  principal: text.nullable(),
});
export const auditCellSchema = z.union([
  z.null(),
  z.boolean(),
  z.number(),
  z.string().max(2_000),
  z.array(z.string().max(500)).max(200),
  auditPersonCellSchema,
]);
export type AuditCell = z.infer<typeof auditCellSchema>;

export const auditRowSchema = z.object({
  // Stable within a run: the directory object ID or the product's SKU ID.
  key: z.string().min(1).max(200),
  // Opens the person's page when the row is an account.
  personId: personIdSchema.nullable(),
  tags: z.array(slug).max(40),
  cells: z.record(columnKey, auditCellSchema),
});
export type AuditRow = z.infer<typeof auditRowSchema>;

export const auditSummaryItemSchema = z.object({
  label: text,
  value: z.number().int().nonnegative(),
  tone: z.enum(["neutral", "good", "warning", "critical"]),
  // Selecting the tile shows this view.
  view: slug.nullable(),
});
export type AuditSummaryItem = z.infer<typeof auditSummaryItemSchema>;

export const auditRunStatusSchema = z.enum(["running", "succeeded", "failed"]);
export type AuditRunStatus = z.infer<typeof auditRunStatusSchema>;

export const auditRunSchema = z.object({
  id: z.uuid(),
  auditId: auditIdSchema,
  auditVersion: z.number().int().positive(),
  status: auditRunStatusSchema,
  // false when Microsoft returned less than everything (a record cap, unreadable records).
  complete: z.boolean(),
  startedAt: time,
  finishedAt: time.nullable(),
  startedBy: z.object({ id: z.uuid(), name: text.nullable() }),
  error: z.object({ code: z.string().max(60), message: text }).nullable(),
  summary: z.array(auditSummaryItemSchema).max(20),
  notes: z.array(text).max(20),
  rowCount: z.number().int().nonnegative(),
  // The account domains this run was limited to; null when it covered every domain.
  scope: z.array(z.string().max(253)).max(100).nullable().default(null),
  // Every domain the run saw before its scope was applied, with its number of accounts.
  domains: z
    .array(z.object({ domain: z.string().max(253), accounts: z.number().int().nonnegative() }))
    .max(1_000)
    .default([]),
});
export type AuditRun = z.infer<typeof auditRunSchema>;

export const auditRunDetailSchema = auditRunSchema.extend({
  rows: z.array(auditRowSchema).max(20_000),
});
export type AuditRunDetail = z.infer<typeof auditRunDetailSchema>;

export const auditAvailabilitySchema = z.object({
  state: z.enum(["available", "missing_permission", "not_enabled"]),
  // Microsoft permissions that are missing.
  missing: z.array(text),
  // The read pack to ask consent for, when the session has not granted a permission.
  grantFeature: z.string().max(80).nullable(),
  // Read packs the operator has to add to ENABLED_READ_FEATURES first.
  enable: z.array(z.string().max(80)),
});
export type AuditAvailability = z.infer<typeof auditAvailabilitySchema>;

export const auditCatalogEntrySchema = z.object({
  definition: auditDefinitionSchema,
  availability: auditAvailabilitySchema,
  lastRun: auditRunSchema.nullable(),
  runningId: z.uuid().nullable(),
});
export type AuditCatalogEntry = z.infer<typeof auditCatalogEntrySchema>;

export const auditCatalogSchema = z.object({
  mode: z.enum(["demo", "live"]),
  audits: z.array(auditCatalogEntrySchema),
});
export type AuditCatalog = z.infer<typeof auditCatalogSchema>;

// The account domains the account audits cover. A domain is the part of a principal name after
// the @, in lower case. No domains means every domain.
export const MAX_AUDIT_SCOPE_DOMAINS = 100;
// Domains offered for choosing. A tenant with more has far more than anyone would tick by hand.
export const MAX_AUDIT_SETTINGS_DOMAINS = 1_000;
export const auditDomainSchema = z
  .string()
  .trim()
  .toLowerCase()
  .max(253)
  .regex(/^(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)*[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?$/);
export const auditScopeSchema = z.object({
  // Kept once each and in order, so the same choice is always saved the same way.
  domains: z
    .array(auditDomainSchema)
    .max(MAX_AUDIT_SCOPE_DOMAINS)
    .transform((domains) => [...new Set(domains)].sort()),
});
export type AuditScope = z.infer<typeof auditScopeSchema>;

// The saved scope and the domains to choose from, with how many accounts each holds in the
// latest results of the audits that use the scope.
export const auditSettingsSchema = z.object({
  scope: auditScopeSchema,
  domains: z
    .array(z.object({ domain: auditDomainSchema, accounts: z.number().int().nonnegative() }))
    .max(MAX_AUDIT_SETTINGS_DOMAINS),
});
export type AuditSettings = z.infer<typeof auditSettingsSchema>;

export const auditRunListSchema = z.object({ runs: z.array(auditRunSchema) });
export const auditStartSchema = z.object({ run: auditRunSchema, started: z.boolean() });

// Filtering, sorting and export, shared by the page and the CSV route so both show the same rows.

export interface AuditQuery {
  view?: string | null;
  facets?: Readonly<Record<string, string>>;
  search?: string;
  sort?: { column: string; direction: "asc" | "desc" } | null;
}

function isPerson(cell: AuditCell | undefined): cell is z.infer<typeof auditPersonCellSchema> {
  return typeof cell === "object" && cell !== null && !Array.isArray(cell);
}

export function auditFacetValue(cell: AuditCell | undefined): string | null {
  if (typeof cell === "string") return cell;
  if (typeof cell === "boolean") return cell ? "Yes" : "No";
  return null;
}

// The values each filter column has in these rows, sorted.
export function auditFacetOptions(
  definition: AuditDefinition,
  rows: readonly AuditRow[],
): Record<string, string[]> {
  const options: Record<string, string[]> = {};
  for (const column of definition.columns.filter((item) => item.facet)) {
    const values = new Set<string>();
    for (const row of rows) {
      const value = auditFacetValue(row.cells[column.key]);
      if (value) values.add(value);
    }
    options[column.key] = [...values].sort((a, b) => a.localeCompare(b));
  }
  return options;
}

function searchText(definition: AuditDefinition, row: AuditRow): string {
  const parts: string[] = [];
  for (const column of definition.columns) {
    const cell = row.cells[column.key];
    if (isPerson(cell)) parts.push(cell.name ?? "", cell.principal ?? "");
    else if (Array.isArray(cell)) parts.push(...cell);
    else if (typeof cell === "string" && column.kind !== "date") parts.push(cell);
  }
  return parts.join("\n").toLowerCase();
}

function sortValue(kind: AuditColumnKind, cell: AuditCell | undefined): number | string | null {
  if (cell === null || cell === undefined) return null;
  if (isPerson(cell)) return (cell.name ?? cell.principal ?? "").toLowerCase() || null;
  if (Array.isArray(cell)) return cell.length;
  if (typeof cell === "boolean") return cell ? 1 : 0;
  if (typeof cell === "number") return cell;
  if (kind === "date") {
    const at = Date.parse(cell);
    return Number.isNaN(at) ? null : at;
  }
  return cell.toLowerCase();
}

// Rows in the chosen view, filters and search, in the chosen order. Empty values sort last in
// both directions; ties keep the order the audit produced.
export function queryAuditRows(
  definition: AuditDefinition,
  rows: readonly AuditRow[],
  query: AuditQuery = {},
): AuditRow[] {
  const view =
    definition.views.find((item) => item.id === (query.view ?? definition.defaultView)) ?? null;
  const facets = Object.entries(query.facets ?? {}).filter(
    ([key, value]) =>
      value && definition.columns.some((column) => column.key === key && column.facet),
  );
  const needle = (query.search ?? "").trim().toLowerCase();
  const filtered = rows.filter(
    (row) =>
      (!view || view.tags.every((tag) => row.tags.includes(tag))) &&
      facets.every(([key, value]) => auditFacetValue(row.cells[key]) === value) &&
      (!needle || searchText(definition, row).includes(needle)),
  );
  const sort = query.sort ?? definition.defaultSort;
  const column = definition.columns.find((item) => item.key === sort.column);
  if (!column) return filtered;
  const sign = sort.direction === "asc" ? 1 : -1;
  return filtered
    .map((row, index) => ({ row, index, value: sortValue(column.kind, row.cells[column.key]) }))
    .sort((a, b) => {
      if (a.value === null || b.value === null)
        return a.value === b.value ? a.index - b.index : a.value === null ? 1 : -1;
      const order =
        typeof a.value === "number" && typeof b.value === "number"
          ? a.value - b.value
          : String(a.value).localeCompare(String(b.value), undefined, { numeric: true });
      return order === 0 ? a.index - b.index : order * sign;
    })
    .map((entry) => entry.row);
}

// One CSV field. Text is always quoted, and text a spreadsheet would read as a formula gets a
// leading apostrophe, so a display name such as "=HYPERLINK(...)" stays text.
function csvField(value: string | number | boolean | null): string {
  if (value === null) return "";
  if (typeof value === "number") return Number.isFinite(value) ? String(value) : "";
  if (typeof value === "boolean") return value ? "Yes" : "No";
  const guarded = /^[=+\-@\t\r]/.test(value) ? `'${value}` : value;
  return `"${guarded.replaceAll('"', '""')}"`;
}

// RFC 4180 CSV with a byte-order mark, so spreadsheet apps read names with accents correctly.
// A person column becomes two: the name and the user principal name.
export function auditCsv(definition: AuditDefinition, rows: readonly AuditRow[]): string {
  const header: string[] = [];
  for (const column of definition.columns) {
    header.push(csvField(column.label));
    if (column.kind === "person") header.push(csvField("User principal name"));
  }
  const lines = [header.join(",")];
  for (const row of rows) {
    const fields: string[] = [];
    for (const column of definition.columns) {
      const cell = row.cells[column.key] ?? null;
      if (column.kind === "person") {
        const person = isPerson(cell) ? cell : { name: null, principal: null };
        fields.push(csvField(person.name), csvField(person.principal));
      } else if (Array.isArray(cell)) fields.push(csvField(cell.join("; ")));
      else if (isPerson(cell)) fields.push(csvField(cell.name ?? cell.principal));
      else fields.push(csvField(cell));
    }
    lines.push(fields.join(","));
  }
  return `﻿${lines.join("\r\n")}\r\n`;
}

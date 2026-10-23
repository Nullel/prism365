import type {
  ChangeCoverage,
  ChangeEvent,
  ChangeFeed,
  PersonChange,
  PersonChangePage,
} from "@prism365/contracts";
import { moduleTablePrefix } from "@prism365/modules";
import type { SqlClient } from "@prism365/storage";
import { z } from "zod";
import { parseCursor } from "./cursor.ts";
import type { ChangeTarget, DirectoryChangeEvent } from "./source.ts";

// Reads over stored directory changes: the feed with its filters and the changes involving one
// person. Values were bounded and cleaned when they were collected; nothing here looks inside the
// targets JSON except the free-text search, which reads it as text.

export const DIRECTORY_CHANGES_MODULE = "directory.changes";
const prefix = moduleTablePrefix(DIRECTORY_CHANGES_MODULE);
export const eventTable = `"${prefix}event"`;
export const subjectTable = `"${prefix}subject"`;

const iso = (value: Date | string) => new Date(value).toISOString();
const isoOrNull = (value: Date | string | null | undefined) => (value ? iso(value) : null);

export const eventColumns = `id, activity_at, category, activity, operation_type, result, result_reason,
  logged_by_service, correlation_id, initiated_by_kind, initiated_by_id, initiated_by_name,
  initiated_by_principal, targets, details`;

export interface EventRow {
  id: string;
  activity_at: Date | string;
  category: string | null;
  activity: string | null;
  operation_type: string | null;
  result: string | null;
  result_reason: string | null;
  logged_by_service: string | null;
  correlation_id: string | null;
  initiated_by_kind: "user" | "app" | "unknown";
  initiated_by_id: string | null;
  initiated_by_name: string | null;
  initiated_by_principal: string | null;
  targets: unknown;
  details: unknown;
}

const nullableString = z.string().nullable().catch(null);
const storedTargets = z
  .array(
    z.object({
      id: nullableString,
      name: nullableString,
      type: nullableString,
      principal: nullableString,
      properties: z
        .array(
          z.object({
            name: nullableString,
            oldValue: nullableString,
            newValue: nullableString,
            oldCut: z.literal(true).optional().catch(undefined),
            newCut: z.literal(true).optional().catch(undefined),
          }),
        )
        .catch([]),
    }),
  )
  .catch([]);
const storedDetails = z.array(z.object({ key: nullableString, value: nullableString })).catch([]);

const json = (value: unknown) =>
  typeof value === "string" ? (JSON.parse(value) as unknown) : value;

// A stored row as the rules see it.
export function eventFromRow(row: EventRow): DirectoryChangeEvent {
  return {
    id: row.id,
    activityAt: new Date(row.activity_at),
    activity: row.activity,
    category: row.category,
    operationType: row.operation_type,
    result: row.result,
    resultReason: row.result_reason,
    loggedByService: row.logged_by_service,
    correlationId: row.correlation_id,
    initiatedBy: {
      kind: row.initiated_by_kind,
      id: row.initiated_by_id,
      name: row.initiated_by_name,
      principal: row.initiated_by_principal,
    },
    targets: storedTargets.parse(json(row.targets)) as ChangeTarget[],
    details: storedDetails.parse(json(row.details)),
  };
}

// A stored row as the feed shows it.
export function changeFromRow(row: EventRow): ChangeEvent {
  const event = eventFromRow(row);
  return {
    ...event,
    activityAt: iso(event.activityAt),
    targets: event.targets.map((target) => ({
      ...target,
      properties: target.properties.map((property) => ({
        name: property.name,
        oldValue: property.oldValue,
        newValue: property.newValue,
        oldCut: property.oldCut === true,
        newCut: property.newCut === true,
      })),
    })),
  };
}

// Page positions are "<ISO time>|<event id>", the last row of the previous page. Microsoft's ids
// never hold "|" after the time, which has none, so the first one separates them.
export function parsePosition(value: string): { at: string; id: string } {
  const split = value.indexOf("|");
  const at = Date.parse(split > 0 ? value.slice(0, split) : "");
  const id = split > 0 ? value.slice(split + 1) : "";
  if (!Number.isFinite(at) || !id) throw new RangeError("Invalid page position.");
  return { at: new Date(at).toISOString(), id };
}

const position = (row: { activity_at: Date | string; id: string }) =>
  `${iso(row.activity_at)}|${row.id}`;

export interface ChangeFilter {
  now: Date;
  days: number;
  // Compared without case.
  category?: string | null;
  // "failure" is every result that is not success: failure, timeout and values Microsoft adds.
  result?: "success" | "failure" | null;
  // Activity, actor or target text, compared without case.
  search?: string | null;
  before?: string | null;
  limit?: number;
}

export async function changeCoverage(sql: SqlClient, tenantId: string): Promise<ChangeCoverage> {
  const stored = await sql.query<{
    first_event: Date | string | null;
    last_collected: Date | string | null;
  }>(
    `SELECT min(activity_at) AS first_event, max(collected_at) AS last_collected
       FROM ${eventTable} WHERE tenant_id = $1`,
    [tenantId],
  );
  const saved = await sql.query<{ value: string }>(
    "SELECT value FROM prism365_cursor WHERE tenant_id = $1 AND module = $2 AND collector = $3",
    [tenantId, DIRECTORY_CHANGES_MODULE, "audits"],
  );
  const cursor = parseCursor(saved.rows[0]?.value);
  const first = stored.rows[0]?.first_event ?? null;
  // History reaches back to where collection started reading, or to the oldest change stored
  // before a reset, whichever is older.
  const from = [cursor?.start ?? null, isoOrNull(first)]
    .filter((value): value is string => value !== null)
    .sort()[0];
  return {
    from: from ?? null,
    through: cursor?.through ?? null,
    lastCollected: isoOrNull(stored.rows[0]?.last_collected),
    phase: cursor?.phase ?? "none",
  };
}

export async function listChanges(
  sql: SqlClient,
  tenantId: string,
  filter: ChangeFilter,
): Promise<ChangeFeed> {
  const limit = Math.min(Math.max(filter.limit ?? 100, 1), 200);
  const since = new Date(filter.now.getTime() - filter.days * 86_400_000).toISOString();
  const page = filter.before ? parsePosition(filter.before) : null;
  const search = (filter.search ?? "").trim().toLowerCase().slice(0, 200);
  const category = (filter.category ?? "").trim().toLowerCase();
  const rows = await sql.query<EventRow>(
    `SELECT ${eventColumns} FROM ${eventTable}
      WHERE tenant_id = $1 AND activity_at >= $2
        AND ($3 = '' OR lower(COALESCE(category, '')) = $3)
        AND ($4 = '' OR ($4 = 'success' AND result = 'success')
             OR ($4 = 'failure' AND COALESCE(result, '') <> 'success'))
        AND ($5 = '' OR strpos(lower(COALESCE(activity, '')), $5) > 0
             OR strpos(lower(COALESCE(initiated_by_name, '')), $5) > 0
             OR strpos(lower(COALESCE(initiated_by_principal, '')), $5) > 0
             OR strpos(lower(targets::text), $5) > 0)
        AND ($6::timestamptz IS NULL OR (activity_at, id) < ($6::timestamptz, $7))
      ORDER BY activity_at DESC, id DESC
      LIMIT $8`,
    [
      tenantId,
      since,
      category,
      filter.result ?? "",
      search,
      page?.at ?? null,
      page?.id ?? "",
      limit + 1,
    ],
  );
  const counts = await sql.query<{ category: string | null; n: number }>(
    `SELECT category, count(*)::int AS n FROM ${eventTable}
      WHERE tenant_id = $1 AND activity_at >= $2
      GROUP BY category ORDER BY count(*) DESC, category LIMIT 50`,
    [tenantId, since],
  );
  const shown = rows.rows.slice(0, limit);
  const last = shown.at(-1);
  return {
    days: filter.days,
    items: shown.map(changeFromRow),
    next: rows.rows.length > limit && last ? position(last) : null,
    categories: counts.rows.map((row) => ({ category: row.category, count: row.n })),
    coverage: await changeCoverage(sql, tenantId),
  };
}

// Changes the person made or that were made to them, newest first. A change where they are both
// (someone editing their own account) is listed once.
export async function personChanges(
  sql: SqlClient,
  tenantId: string,
  userId: string,
  options: { before?: string | null; limit?: number } = {},
): Promise<PersonChangePage> {
  const limit = Math.min(Math.max(options.limit ?? 20, 1), 200);
  const page = options.before ? parsePosition(options.before) : null;
  const rows = await sql.query<EventRow & { as_actor: boolean; as_target: boolean }>(
    `SELECT e.id, e.activity_at, e.category, e.activity, e.operation_type, e.result,
            e.result_reason, e.logged_by_service, e.correlation_id, e.initiated_by_kind,
            e.initiated_by_id, e.initiated_by_name, e.initiated_by_principal, e.targets, e.details,
            bool_or(s.subject_kind = 'actor') AS as_actor,
            bool_or(s.subject_kind = 'target') AS as_target
       FROM ${subjectTable} AS s
       JOIN ${eventTable} AS e ON e.tenant_id = s.tenant_id AND e.id = s.event_id
      WHERE s.tenant_id = $1 AND s.subject_id = $2
        AND ($3::timestamptz IS NULL OR (s.activity_at, s.event_id) < ($3::timestamptz, $4))
      GROUP BY e.tenant_id, e.id
      ORDER BY e.activity_at DESC, e.id DESC
      LIMIT $5`,
    [tenantId, userId.toLowerCase(), page?.at ?? null, page?.id ?? "", limit + 1],
  );
  const shown = rows.rows.slice(0, limit);
  const last = shown.at(-1);
  return {
    items: shown.map(
      (row): PersonChange => ({
        ...changeFromRow(row),
        involvement: { actor: row.as_actor, target: row.as_target },
      }),
    ),
    next: rows.rows.length > limit && last ? position(last) : null,
  };
}

import {
  type AuditRow,
  type AuditRun,
  type AuditRunDetail,
  type AuditRunTrigger,
  type AuditSummaryItem,
  auditRunTriggerSchema,
} from "@prism365/contracts";
import type { SqlClient } from "@prism365/storage";
import { joinKey } from "./activity-reports.ts";
import { HASHED } from "./m365-app-usage.ts";
import { auditTrendTiles } from "./trend.ts";

// Audit runs in prism365_audit_run. Every query is tenant-scoped twice: by the transaction's
// row-security setting and by the tenant_id predicate.

// How many runs to keep per audit, counted apart for successes and failures. Older ones are
// deleted when a run finishes.
export const KEEP_RUNS = 10;
export const KEEP_FAILED_RUNS = 5;

const summaryColumns = `id, audit_id, audit_version, status, complete, started_at, finished_at,
  started_by, started_by_name, error_code, error_message, summary, notes, row_count, run_scope,
  domain_counts, trigger`;

interface RunRecord {
  id: string;
  audit_id: string;
  audit_version: number;
  status: AuditRun["status"];
  complete: boolean;
  started_at: Date | string;
  finished_at: Date | string | null;
  started_by: string;
  started_by_name: string | null;
  error_code: string | null;
  error_message: string | null;
  summary: unknown;
  notes: unknown;
  row_count: number;
  run_scope: unknown;
  domain_counts: unknown;
  trigger: string | null;
  result_rows?: unknown;
}

const iso = (value: Date | string) => new Date(value).toISOString();
const json = (value: unknown): unknown[] => {
  const parsed = typeof value === "string" ? (JSON.parse(value) as unknown) : value;
  return Array.isArray(parsed) ? parsed : [];
};

function toRun(row: RunRecord): AuditRun {
  return {
    id: row.id,
    auditId: row.audit_id,
    auditVersion: row.audit_version,
    status: row.status,
    complete: row.complete,
    startedAt: iso(row.started_at),
    finishedAt: row.finished_at === null ? null : iso(row.finished_at),
    startedBy: { id: row.started_by, name: row.started_by_name },
    error:
      row.error_code === null ? null : { code: row.error_code, message: row.error_message ?? "" },
    summary: json(row.summary) as AuditSummaryItem[],
    notes: json(row.notes).filter((note): note is string => typeof note === "string"),
    rowCount: row.row_count,
    scope: scopeOf(row.run_scope),
    domains: json(row.domain_counts)
      .filter(
        (item): item is { domain: string; accounts: number } =>
          typeof item === "object" &&
          item !== null &&
          typeof (item as { domain?: unknown }).domain === "string" &&
          typeof (item as { accounts?: unknown }).accounts === "number",
      )
      .slice(0, 1_000),
    trigger: auditRunTriggerSchema.catch("manual").parse(row.trigger),
  };
}

function scopeOf(value: unknown): string[] | null {
  const parsed = typeof value === "string" ? (JSON.parse(value) as unknown) : value;
  if (!Array.isArray(parsed)) return null;
  const domains = parsed.filter((item): item is string => typeof item === "string");
  return domains.length > 0 ? domains.slice(0, 100) : null;
}

export async function insertRun(
  sql: SqlClient,
  tenantId: string,
  run: {
    id: string;
    auditId: string;
    auditVersion: number;
    startedAt: Date;
    startedBy: { id: string; name: string | null };
    // The account domains this run is limited to; null or empty for every domain.
    scope?: readonly string[] | null;
    // Who asked for the run: an administrator, or the audit's schedule. Manual by default.
    trigger?: AuditRunTrigger;
  },
): Promise<AuditRun> {
  const result = await sql.query<RunRecord>(
    `INSERT INTO prism365_audit_run (id, tenant_id, audit_id, audit_version, status, started_at,
       started_by, started_by_name, run_scope, trigger)
     VALUES ($1,$2,$3,$4,'running',$5,$6,$7,$8::jsonb,$9)
     RETURNING ${summaryColumns}`,
    [
      run.id,
      tenantId,
      run.auditId,
      run.auditVersion,
      run.startedAt.toISOString(),
      run.startedBy.id,
      run.startedBy.name?.slice(0, 500) ?? null,
      run.scope && run.scope.length > 0 ? JSON.stringify([...run.scope].sort()) : null,
      run.trigger ?? "manual",
    ],
  );
  const row = result.rows[0];
  if (!row) throw new Error("The audit run was not stored.");
  return toRun(row);
}

export type RunOutcome =
  | {
      status: "succeeded";
      complete: boolean;
      summary: AuditSummaryItem[];
      notes: string[];
      rows: AuditRow[];
      domains?: { domain: string; accounts: number }[];
    }
  | { status: "failed"; code: string; message: string };

// Records how a run ended. Only a run still marked running is changed, so a run that was
// already closed (timed out, or marked interrupted) keeps its first outcome. A run that
// succeeded, complete or not, also leaves its trend point, in the same transaction as its
// result: a result storage refuses leaves no point, and a failed run never has one.
export async function finishRun(
  sql: SqlClient,
  tenantId: string,
  id: string,
  finishedAt: Date,
  outcome: RunOutcome,
): Promise<AuditRun | null> {
  const succeeded = outcome.status === "succeeded";
  const result = await sql.query<RunRecord>(
    `UPDATE prism365_audit_run
        SET status=$3, complete=$4, finished_at=$5, error_code=$6, error_message=$7,
            summary=$8::jsonb, notes=$9::jsonb, result_rows=$10::jsonb, row_count=$11,
            domain_counts=$12::jsonb
      WHERE tenant_id=$1 AND id=$2 AND status='running'
      RETURNING ${summaryColumns}`,
    [
      tenantId,
      id,
      outcome.status,
      succeeded ? outcome.complete : false,
      finishedAt.toISOString(),
      succeeded ? null : outcome.code.slice(0, 60),
      succeeded ? null : outcome.message.slice(0, 500),
      JSON.stringify(succeeded ? outcome.summary : []),
      JSON.stringify(succeeded ? outcome.notes : []),
      JSON.stringify(succeeded ? outcome.rows : []),
      succeeded ? outcome.rows.length : 0,
      JSON.stringify(succeeded ? (outcome.domains ?? []).slice(0, 1_000) : []),
    ],
  );
  const row = result.rows[0];
  if (!row) return null;
  if (outcome.status === "succeeded") {
    const scope = scopeOf(row.run_scope);
    await sql.query(
      `INSERT INTO prism365_audit_trend (tenant_id, run_id, audit_id, audit_version, finished_at,
         complete, run_scope, tiles, row_count, trigger)
       VALUES ($1,$2,$3,$4,$5,$6,$7::jsonb,$8::jsonb,$9,$10)
       ON CONFLICT (tenant_id, run_id) DO NOTHING`,
      [
        tenantId,
        row.id,
        row.audit_id,
        row.audit_version,
        iso(row.finished_at ?? finishedAt),
        row.complete,
        scope ? JSON.stringify(scope) : null,
        JSON.stringify(auditTrendTiles(outcome.summary)),
        row.row_count,
        auditRunTriggerSchema.catch("manual").parse(row.trigger),
      ],
    );
  }
  return toRun(row);
}

// The newest start of any run of one audit, whatever its outcome or trigger: what a schedule
// counts from. null when the audit never ran.
export async function latestRunStart(
  sql: SqlClient,
  tenantId: string,
  auditId: string,
): Promise<Date | null> {
  const result = await sql.query<{ started_at: Date | string }>(
    `SELECT started_at FROM prism365_audit_run
      WHERE tenant_id=$1 AND audit_id=$2
      ORDER BY started_at DESC, id DESC
      LIMIT 1`,
    [tenantId, auditId],
  );
  const row = result.rows[0];
  return row ? new Date(row.started_at) : null;
}

// Deletes old runs of one audit: succeeded runs beyond the newest `keep`, failed runs beyond the
// newest `keepFailed`. The two are counted apart, so a string of failures (throttling, an expired
// session) can never push out the last good result. Running runs are never deleted, and neither
// are trend points: they live in their own table and outlast the runs they came from.
export async function pruneRuns(
  sql: SqlClient,
  tenantId: string,
  auditId: string,
  keep = KEEP_RUNS,
  keepFailed = KEEP_FAILED_RUNS,
): Promise<number> {
  let deleted = 0;
  for (const [status, limit] of [
    ["succeeded", keep],
    ["failed", keepFailed],
  ] as const) {
    const old = await sql.query<{ id: string }>(
      `SELECT id FROM prism365_audit_run
        WHERE tenant_id=$1 AND audit_id=$2 AND status=$3
        ORDER BY started_at DESC, id DESC
        OFFSET $4`,
      [tenantId, auditId, status, limit],
    );
    if (old.rows.length === 0) continue;
    const result = await sql.query(
      `DELETE FROM prism365_audit_run
        WHERE tenant_id=$1 AND audit_id=$2 AND status=$3 AND id = ANY($4::uuid[])`,
      [tenantId, auditId, status, old.rows.map((row) => row.id)],
    );
    deleted += result.rowCount;
  }
  return deleted;
}

// Runs still marked running that this process is not running: the process stopped or crashed
// mid-run. Prism365 runs one API process per installation, so a run it does not hold is dead.
export async function interruptRuns(
  sql: SqlClient,
  tenantId: string,
  live: readonly string[],
  at: Date,
): Promise<number> {
  const result = await sql.query(
    `UPDATE prism365_audit_run
        SET status='failed', finished_at=$3, error_code='interrupted',
            error_message='Prism365 stopped while this audit was running. Run it again.'
      WHERE tenant_id=$1 AND status='running' AND NOT (id = ANY($2::uuid[]))`,
    [tenantId, [...live], at.toISOString()],
  );
  return result.rowCount;
}

export async function listRuns(
  sql: SqlClient,
  tenantId: string,
  auditId: string,
  limit = KEEP_RUNS + KEEP_FAILED_RUNS + 1,
): Promise<AuditRun[]> {
  const result = await sql.query<RunRecord>(
    `SELECT ${summaryColumns} FROM prism365_audit_run
      WHERE tenant_id=$1 AND audit_id=$2
      ORDER BY started_at DESC, id DESC
      LIMIT $3`,
    [tenantId, auditId, limit],
  );
  return result.rows.map(toRun);
}

// The newest run of each audit.
export async function latestRuns(sql: SqlClient, tenantId: string): Promise<AuditRun[]> {
  const result = await sql.query<RunRecord>(
    `SELECT DISTINCT ON (audit_id) ${summaryColumns} FROM prism365_audit_run
      WHERE tenant_id=$1
      ORDER BY audit_id, started_at DESC, id DESC`,
    [tenantId],
  );
  return result.rows.map(toRun);
}

// The newest run of one audit that finished successfully, with its rows. A newer failed or
// running run does not hide it.
export async function latestSucceededRun(
  sql: SqlClient,
  tenantId: string,
  auditId: string,
): Promise<AuditRunDetail | null> {
  const result = await sql.query<RunRecord>(
    `SELECT ${summaryColumns}, result_rows FROM prism365_audit_run
      WHERE tenant_id=$1 AND audit_id=$2 AND status='succeeded'
      ORDER BY started_at DESC, id DESC
      LIMIT 1`,
    [tenantId, auditId],
  );
  const row = result.rows[0];
  return row ? { ...toRun(row), rows: json(row.result_rows) as AuditRow[] } : null;
}

export async function getRun(
  sql: SqlClient,
  tenantId: string,
  id: string,
): Promise<AuditRunDetail | null> {
  const result = await sql.query<RunRecord>(
    `SELECT ${summaryColumns}, result_rows FROM prism365_audit_run
      WHERE tenant_id=$1 AND id=$2`,
    [tenantId, id],
  );
  const row = result.rows[0];
  if (!row) return null;
  return { ...toRun(row), rows: json(row.result_rows) as AuditRow[] };
}

// Which rows of a run to pick.
//
// - personId: the rows whose account id is this one.
// - name: the rows whose cell at `path` is this sign-in name, compared the way the audits compare
//   names (NFKC, trimmed, lower case). A cell that is a hash is Microsoft hiding the name, and is
//   never matched against anything.
// - list: the rows whose cell at `path` is a list that holds `value` exactly (the owners of a group).
// - keys: the rows whose key is one of `keys`, or whose cell at `cell.path` is one of
//   `cell.values` (an application registration is found by its application id, which the row
//   carries as a cell). Ids compare in lower case.
export type RowPick =
  | { by: "personId"; personId: string }
  | { by: "name"; path: readonly string[]; principal: string }
  | { by: "list"; path: readonly string[]; value: string }
  | {
      by: "keys";
      keys: readonly string[];
      cell?: { path: readonly string[]; values: readonly string[] };
    };

function cellAt(row: AuditRow, path: readonly string[]): unknown {
  let value: unknown = row.cells;
  for (const step of path) {
    if (typeof value !== "object" || value === null || Array.isArray(value)) return undefined;
    value = (value as Record<string, unknown>)[step];
  }
  return value;
}

// The one definition of "this row belongs to that pick". The SQL below only narrows the rows it
// hands back to this: it never decides alone.
export function rowMatches(row: AuditRow, pick: RowPick): boolean {
  switch (pick.by) {
    case "personId":
      return row.personId === pick.personId;
    case "name": {
      const cell = cellAt(row, pick.path);
      if (typeof cell !== "string") return false;
      const name = cell.trim();
      if (name === "" || HASHED.test(name)) return false;
      return joinKey(name) === joinKey(pick.principal);
    }
    case "list": {
      const cell = cellAt(row, pick.path);
      return Array.isArray(cell) && cell.includes(pick.value);
    }
    case "keys": {
      if (pick.keys.some((key) => key.toLowerCase() === row.key.toLowerCase())) return true;
      if (!pick.cell) return false;
      const cell = cellAt(row, pick.cell.path);
      return (
        typeof cell === "string" &&
        pick.cell.values.some((value) => value.toLowerCase() === cell.toLowerCase())
      );
    }
  }
}

export interface PickedRows {
  run: AuditRun;
  // At most `limit` rows, in the order the run stores them.
  rows: AuditRow[];
  // The run held more matching rows than `limit`.
  more: boolean;
}

// The most rows of a name comparison the database hands back for JavaScript to settle. A cell with
// a character outside printable ASCII could equal the name only after NFKC folding (a full-width
// letter) or because of white space other than a space, which the database cannot do, so every
// such cell is a candidate whatever name is sought. A name equal to an ASCII cell is found by the
// comparison itself. More candidates than this and the run is read whole instead.
const NAME_CANDIDATES = 200;

// The rows of the newest succeeded run of one audit that match `pick`, found by the database in
// one statement (the run's header and its matching rows come from one snapshot), so a page view
// never loads a whole run of up to 20,000 rows. null when the audit never finished a run.
//
// The database only narrows the rows; rowMatches decides, so the answer is the same as filtering
// the whole run in memory. A name comparison is narrowed to the cells that equal the name in
// plain ASCII plus the few that hold any other character, and falls back to reading the run when
// there are too many of those.
export async function pickRows(
  sql: SqlClient,
  tenantId: string,
  auditId: string,
  pick: RowPick,
  limit: number,
): Promise<PickedRows | null> {
  const wholeRun = async (): Promise<PickedRows | null> => {
    const run = await latestSucceededRun(sql, tenantId, auditId);
    if (!run) return null;
    const matching = run.rows.filter((row) => rowMatches(row, pick));
    const { rows: _rows, ...header } = run;
    return { run: header, rows: matching.slice(0, limit), more: matching.length > limit };
  };
  const fetch = pick.by === "name" ? NAME_CANDIDATES : limit;
  const where =
    pick.by === "personId"
      ? "e.elem->>'personId' = $3"
      : pick.by === "name"
        ? `jsonb_typeof(e.elem #> $3::text[]) = 'string' AND (lower(btrim(e.elem #>> $3::text[]) COLLATE "C") = $4 OR (e.elem #>> $3::text[]) ~ '[^ -~]')`
        : pick.by === "list"
          ? "jsonb_typeof(e.elem #> $3::text[]) = 'array' AND (e.elem #> $3::text[]) @> $4::jsonb"
          : `lower((e.elem->>'key') COLLATE "C") = ANY($3::text[]) OR lower((e.elem #>> $4::text[]) COLLATE "C") = ANY($5::text[])`;
  const params: unknown[] =
    pick.by === "personId"
      ? [pick.personId]
      : pick.by === "name"
        ? [["cells", ...pick.path], joinKey(pick.principal)]
        : pick.by === "list"
          ? [["cells", ...pick.path], JSON.stringify([pick.value])]
          : [
              pick.keys.map((key) => key.toLowerCase()),
              ["cells", ...(pick.cell?.path ?? [])],
              (pick.cell?.values ?? []).map((value) => value.toLowerCase()),
            ];
  const result = await sql.query<RunRecord & { elem: unknown }>(
    `WITH latest AS (
       SELECT ${summaryColumns}, result_rows FROM prism365_audit_run
        WHERE tenant_id=$1 AND audit_id=$2 AND status='succeeded'
        ORDER BY started_at DESC, id DESC
        LIMIT 1)
     SELECT l.id, l.audit_id, l.audit_version, l.status, l.complete, l.started_at, l.finished_at,
            l.started_by, l.started_by_name, l.error_code, l.error_message, l.summary, l.notes,
            l.row_count, l.run_scope, l.domain_counts, l.trigger, m.elem
       FROM latest l
       LEFT JOIN LATERAL (
         SELECT e.elem, e.ord FROM jsonb_array_elements(
           CASE WHEN jsonb_typeof(l.result_rows) = 'array' THEN l.result_rows ELSE '[]'::jsonb END)
           WITH ORDINALITY AS e(elem, ord)
          WHERE ${where}
          ORDER BY e.ord
          LIMIT $${params.length + 3}) m ON true
      ORDER BY m.ord`,
    [tenantId, auditId, ...params, fetch + 1],
  );
  const first = result.rows[0];
  if (!first) return null;
  const candidates = result.rows.flatMap((row) =>
    row.elem === null ? [] : [row.elem as AuditRow],
  );
  if (pick.by === "name" && candidates.length > fetch) return wholeRun();
  const found = candidates.filter((row) => rowMatches(row, pick));
  return { run: toRun(first), rows: found.slice(0, limit), more: found.length > limit };
}

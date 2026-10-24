import type { AuditRow, AuditRun, AuditRunDetail, AuditSummaryItem } from "@prism365/contracts";
import type { SqlClient } from "@prism365/storage";

// Audit runs in prism365_audit_run. Every query is tenant-scoped twice: by the transaction's
// row-security setting and by the tenant_id predicate.

// How many runs to keep per audit, counted apart for successes and failures. Older ones are
// deleted when a run finishes.
export const KEEP_RUNS = 10;
export const KEEP_FAILED_RUNS = 5;

const summaryColumns = `id, audit_id, audit_version, status, complete, started_at, finished_at,
  started_by, started_by_name, error_code, error_message, summary, notes, row_count, run_scope,
  domain_counts`;

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
  },
): Promise<AuditRun> {
  const result = await sql.query<RunRecord>(
    `INSERT INTO prism365_audit_run (id, tenant_id, audit_id, audit_version, status, started_at,
       started_by, started_by_name, run_scope)
     VALUES ($1,$2,$3,$4,'running',$5,$6,$7,$8::jsonb)
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
// already closed (timed out, or marked interrupted) keeps its first outcome.
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
  return row ? toRun(row) : null;
}

// Deletes old runs of one audit: succeeded runs beyond the newest `keep`, failed runs beyond the
// newest `keepFailed`. The two are counted apart, so a string of failures (throttling, an expired
// session) can never push out the last good result. Running runs are never deleted.
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

import { severitySchema } from "@prism365/contracts";
import type { ModuleDatabase } from "@prism365/storage";
import {
  MAX_STORED_FINDINGS,
  type StoredFinding,
  type StoredResults,
  type StoredRun,
} from "./definition.ts";
import { latestSucceededRun } from "./store.ts";

// What the runner hands an audit that joins other audits' results. Every call is one read-only
// query in the run's own tenant: the tenant is fixed here, so an audit cannot name another, and
// the only statements are SELECTs. The run's signal is checked before each read, so a stopped run
// reads nothing more.

interface FindingRecord {
  rule_id: string;
  severity: string;
  subject_kind: string;
  subject_id: string;
  title: string;
  last_seen: Date | string;
}

export function createStoredResults(
  database: ModuleDatabase,
  tenantId: string,
  signal: AbortSignal,
): StoredResults {
  return {
    async latestRun(auditId) {
      signal.throwIfAborted();
      const run = await database.tenant(tenantId, (sql) =>
        latestSucceededRun(sql, tenantId, auditId),
      );
      if (!run?.finishedAt) return null;
      const stored: StoredRun = {
        finishedAt: run.finishedAt,
        complete: run.complete,
        scope: run.scope,
        rows: run.rows,
      };
      return stored;
    },
    async openFindings(moduleId) {
      signal.throwIfAborted();
      const result = await database.tenant(tenantId, (sql) =>
        sql.query<FindingRecord>(
          `SELECT rule_id, severity, subject_kind, subject_id, title, last_seen
             FROM prism365_finding
            WHERE tenant_id=$1 AND module=$2 AND state='open'
            ORDER BY last_seen DESC, id
            LIMIT $3`,
          [tenantId, moduleId, MAX_STORED_FINDINGS],
        ),
      );
      const findings: StoredFinding[] = [];
      for (const row of result.rows) {
        const severity = severitySchema.safeParse(row.severity);
        if (!severity.success) continue;
        findings.push({
          ruleId: row.rule_id,
          severity: severity.data,
          subjectKind: row.subject_kind,
          subjectId: row.subject_id,
          title: row.title,
          lastSeen: new Date(row.last_seen).toISOString(),
        });
      }
      return findings;
    },
  };
}

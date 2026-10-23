import { randomUUID } from "node:crypto";
import { type Severity, severitySchema } from "@prism365/contracts";
import { containsUnsupportedWorkforceClaim, evaluateFinding } from "@prism365/policy";
import type { SqlClient } from "@prism365/storage";
import { z } from "zod";
import type { FindingChange, FindingInput } from "./contract.ts";

export class FindingPolicyError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "FindingPolicyError";
  }
}

const severityRank: Record<Severity, number> = {
  info: 0,
  low: 1,
  medium: 2,
  high: 3,
  critical: 4,
};

export function severityAtLeast(value: Severity, minimum: Severity): boolean {
  return severityRank[value] >= severityRank[minimum];
}

const fieldValue = z.union([z.string().max(500), z.number().finite(), z.boolean(), z.null()]);

const findingInputSchema = z
  .object({
    ruleId: z.string().regex(/^[a-z][a-z0-9_.-]{1,60}$/),
    ruleVersion: z.number().int().positive(),
    dedupeKey: z.string().min(1).max(200),
    purpose: z.enum(["security", "operations", "compliance", "adoption"]),
    severity: severitySchema,
    score: z.number().int().min(0).max(1_000),
    subject: z.object({
      kind: z.enum(["tenant", "service", "resource", "person"]),
      id: z.string().min(1).max(200),
      label: z.string().min(1).max(200).optional(),
    }),
    title: z.string().trim().min(1).max(160),
    summary: z.string().trim().min(1).max(1_000),
    evidence: z
      .array(
        z.object({
          source: z.string().min(1).max(120),
          ref: z.string().min(1).max(200),
          observedAt: z.string().datetime({ offset: true }),
          fields: z.record(z.string().regex(/^[A-Za-z][A-Za-z0-9_]{0,40}$/), fieldValue),
        }),
      )
      .min(1)
      .max(20),
    limitations: z.array(z.string().trim().min(1).max(300)).min(1).max(10),
    window: z.object({ start: z.date(), end: z.date() }),
    observedAt: z.date(),
  })
  .refine((value) => value.window.start.getTime() <= value.window.end.getTime(), {
    message: "The finding window must not end before it starts.",
  })
  .refine((value) => JSON.stringify(value.evidence).length <= 16_000, {
    message: "Finding evidence exceeds 16 KB.",
  });

// Validates shape and language before anything is stored. The language policy is the same one
// every Prism365 finding passes: no claims about productivity, effort, or performance.
export function checkFinding(input: FindingInput): FindingInput {
  const parsed = findingInputSchema.safeParse(input);
  if (!parsed.success) throw new FindingPolicyError("The finding does not match its contract.");
  const decision = evaluateFinding({
    title: input.title,
    summary: input.summary,
    purpose: input.purpose,
    subject: input.subject.kind,
    hasSource: input.evidence.length > 0,
    hasTimeWindow: true,
    hasLimitations: input.limitations.length > 0,
  });
  if (!decision.allowed) throw new FindingPolicyError(decision.reason);
  for (const item of input.evidence)
    for (const value of Object.values(item.fields))
      if (typeof value === "string" && containsUnsupportedWorkforceClaim(value))
        throw new FindingPolicyError("Evidence fields make an unsupported workforce claim.");
  if (input.subject.label && containsUnsupportedWorkforceClaim(input.subject.label))
    throw new FindingPolicyError("The subject label makes an unsupported workforce claim.");
  return input;
}

interface ExistingFinding {
  id: string;
  severity: Severity;
  state: "open" | "acknowledged" | "resolved";
}

// Inserts a new finding or folds a repeat observation into the existing one. A resolved finding
// that recurs is reopened, and an acknowledged one reopens when its severity rises, so marking
// something handled never silences it for good. A severity rise on an open finding is reported
// as "escalated" so sinks can announce it, without changing its state.
export async function upsertFinding(
  sql: SqlClient,
  tenantId: string,
  moduleId: string,
  raw: FindingInput,
): Promise<{ id: string; change: FindingChange }> {
  const input = checkFinding(raw);
  const values = {
    evidence: JSON.stringify(input.evidence),
    limitations: JSON.stringify(input.limitations),
    label: input.subject.label ?? null,
    start: input.window.start.toISOString(),
    end: input.window.end.toISOString(),
    seen: input.observedAt.toISOString(),
  };
  const inserted = await sql.query<{ id: string }>(
    `INSERT INTO prism365_finding (id, tenant_id, module, rule_id, rule_version, dedupe_key, purpose,
       severity, score, subject_kind, subject_id, subject_label, title, summary, evidence,
       limitations, window_start, window_end, state, occurrences, first_seen, last_seen)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15::jsonb,$16::jsonb,$17,$18,'open',1,$19,$19)
     ON CONFLICT (tenant_id, module, rule_id, dedupe_key) DO NOTHING
     RETURNING id`,
    [
      randomUUID(),
      tenantId,
      moduleId,
      input.ruleId,
      input.ruleVersion,
      input.dedupeKey,
      input.purpose,
      input.severity,
      input.score,
      input.subject.kind,
      input.subject.id,
      values.label,
      input.title,
      input.summary,
      values.evidence,
      values.limitations,
      values.start,
      values.end,
      values.seen,
    ],
  );
  const created = inserted.rows[0];
  if (created) return { id: created.id, change: "created" };

  const current = await sql.query<ExistingFinding>(
    `SELECT id, severity, state FROM prism365_finding
      WHERE tenant_id=$1 AND module=$2 AND rule_id=$3 AND dedupe_key=$4 FOR UPDATE`,
    [tenantId, moduleId, input.ruleId, input.dedupeKey],
  );
  const existing = current.rows[0];
  if (!existing) throw new Error("The finding disappeared during its update.");
  const escalated = severityRank[input.severity] > severityRank[existing.severity];
  const change: FindingChange =
    existing.state === "resolved" ? "reopened" : escalated ? "escalated" : "updated";
  const severity = escalated ? input.severity : existing.severity;
  const reopen =
    change === "reopened" || (change === "escalated" && existing.state === "acknowledged");
  await sql.query(
    `UPDATE prism365_finding SET
       rule_version=$3, severity=$4, score=GREATEST(score,$5), subject_label=COALESCE($6, subject_label),
       title=$7, summary=$8, evidence=$9::jsonb, limitations=$10::jsonb,
       window_start=LEAST(window_start,$11), window_end=GREATEST(window_end,$12),
       last_seen=GREATEST(last_seen,$13), occurrences=occurrences+1,
       state=CASE WHEN $14 THEN 'open' ELSE state END,
       state_changed_at=CASE WHEN $14 THEN $13::timestamptz ELSE state_changed_at END,
       state_changed_by=CASE WHEN $14 THEN NULL ELSE state_changed_by END
     WHERE tenant_id=$1 AND id=$2`,
    [
      tenantId,
      existing.id,
      input.ruleVersion,
      severity,
      input.score,
      values.label,
      input.title,
      input.summary,
      values.evidence,
      values.limitations,
      values.start,
      values.end,
      values.seen,
      reopen,
    ],
  );
  return { id: existing.id, change };
}

export interface FindingRecord {
  id: string;
  module: string;
  ruleId: string;
  ruleVersion: number;
  purpose: string;
  severity: Severity;
  score: number;
  subject: { kind: string; id: string; label: string | null };
  title: string;
  summary: string;
  evidence: unknown;
  limitations: unknown;
  windowStart: string;
  windowEnd: string;
  state: "open" | "acknowledged" | "resolved";
  occurrences: number;
  firstSeen: string;
  lastSeen: string;
  stateChangedAt: string | null;
}

interface FindingRow {
  id: string;
  module: string;
  rule_id: string;
  rule_version: number;
  purpose: string;
  severity: Severity;
  score: number;
  subject_kind: string;
  subject_id: string;
  subject_label: string | null;
  title: string;
  summary: string;
  evidence: unknown;
  limitations: unknown;
  window_start: Date | string;
  window_end: Date | string;
  state: "open" | "acknowledged" | "resolved";
  occurrences: number;
  first_seen: Date | string;
  last_seen: Date | string;
  state_changed_at: Date | string | null;
}

function iso(value: Date | string): string {
  return (value instanceof Date ? value : new Date(value)).toISOString();
}

function toRecord(row: FindingRow): FindingRecord {
  return {
    id: row.id,
    module: row.module,
    ruleId: row.rule_id,
    ruleVersion: row.rule_version,
    purpose: row.purpose,
    severity: row.severity,
    score: row.score,
    subject: { kind: row.subject_kind, id: row.subject_id, label: row.subject_label },
    title: row.title,
    summary: row.summary,
    evidence: typeof row.evidence === "string" ? JSON.parse(row.evidence) : row.evidence,
    limitations:
      typeof row.limitations === "string" ? JSON.parse(row.limitations) : row.limitations,
    windowStart: iso(row.window_start),
    windowEnd: iso(row.window_end),
    state: row.state,
    occurrences: row.occurrences,
    firstSeen: iso(row.first_seen),
    lastSeen: iso(row.last_seen),
    stateChangedAt: row.state_changed_at ? iso(row.state_changed_at) : null,
  };
}

const findingColumns = `id, module, rule_id, rule_version, purpose, severity, score, subject_kind,
  subject_id, subject_label, title, summary, evidence, limitations, window_start, window_end,
  state, occurrences, first_seen, last_seen, state_changed_at`;

export async function listFindings(
  sql: SqlClient,
  tenantId: string,
  filter: {
    module?: string;
    state?: FindingRecord["state"];
    // Only findings about this person (subject kind "person").
    personId?: string;
    limit?: number;
  } = {},
): Promise<FindingRecord[]> {
  const limit = Math.min(Math.max(filter.limit ?? 50, 1), 200);
  const result = await sql.query<FindingRow>(
    `SELECT ${findingColumns} FROM prism365_finding
      WHERE tenant_id=$1 AND ($2::text IS NULL OR module=$2) AND ($3::text IS NULL OR state=$3)
        AND ($5::text IS NULL OR (subject_kind='person' AND subject_id=$5))
      ORDER BY last_seen DESC, id LIMIT $4`,
    [tenantId, filter.module ?? null, filter.state ?? null, limit, filter.personId ?? null],
  );
  return result.rows.map(toRecord);
}

export async function setFindingState(
  sql: SqlClient,
  tenantId: string,
  findingId: string,
  state: FindingRecord["state"],
  actorId: string,
): Promise<FindingRecord | null> {
  const result = await sql.query<FindingRow>(
    `UPDATE prism365_finding SET state=$3, state_changed_at=now(), state_changed_by=$4
      WHERE tenant_id=$1 AND id=$2 RETURNING ${findingColumns}`,
    [tenantId, findingId, state, actorId],
  );
  const row = result.rows[0];
  return row ? toRecord(row) : null;
}

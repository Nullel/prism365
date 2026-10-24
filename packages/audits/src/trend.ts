import {
  type AuditRunTrigger,
  type AuditSummaryItem,
  type AuditTrend,
  type AuditTrendDifference,
  auditRunTriggerSchema,
  auditTileId,
  auditTileIdSchema,
  auditTileToneSchema,
  auditTrendSchema,
  MAX_AUDIT_TREND_POINTS,
} from "@prism365/contracts";
import type { SqlClient } from "@prism365/storage";
import { z } from "zod";

// Audit trends: one point per successful run in prism365_audit_trend, written by finishRun in
// the same transaction as the run's result and never deleted. A point holds the run's summary
// tiles, so each tile's value can be followed after the run itself has been pruned.

export interface TrendTile {
  id: string;
  label: string;
  value: number;
  tone: AuditSummaryItem["tone"];
}

// The tiles a point keeps: the run's summary without its views, each with an id from its label.
// A label that gives the same id as an earlier tile is dropped, so a point never holds two
// values for one id. The migration's backfill does the same in SQL.
export function auditTrendTiles(summary: readonly AuditSummaryItem[]): TrendTile[] {
  const seen = new Set<string>();
  const tiles: TrendTile[] = [];
  for (const item of summary) {
    const id = auditTileId(item.label);
    if (seen.has(id)) continue;
    seen.add(id);
    tiles.push({ id, label: item.label, value: item.value, tone: item.tone });
  }
  return tiles;
}

export interface StoredTrendPoint {
  runId: string;
  finishedAt: string;
  version: number;
  complete: boolean;
  // The domains the run was limited to, sorted; null for every domain.
  scope: string[] | null;
  trigger: AuditRunTrigger;
  tiles: TrendTile[];
}

interface TrendRecord {
  run_id: string;
  audit_version: number;
  finished_at: Date | string;
  complete: boolean;
  run_scope: unknown;
  tiles: unknown;
  trigger: string | null;
}

const storedTileSchema = z.object({
  id: auditTileIdSchema,
  label: z.string().transform((value) => value.slice(0, 500)),
  value: z.number().int().nonnegative(),
  tone: auditTileToneSchema.catch("neutral"),
});

function parsedJson(value: unknown): unknown {
  if (typeof value !== "string") return value;
  try {
    return JSON.parse(value) as unknown;
  } catch {
    return null;
  }
}

// Stored tiles, read leniently: one that does not parse is left out rather than failing the
// whole trend.
export function storedTiles(value: unknown): TrendTile[] {
  const parsed = parsedJson(value);
  if (!Array.isArray(parsed)) return [];
  const seen = new Set<string>();
  const tiles: TrendTile[] = [];
  for (const item of parsed) {
    const tile = storedTileSchema.safeParse(item);
    if (!tile.success || seen.has(tile.data.id)) continue;
    seen.add(tile.data.id);
    tiles.push(tile.data);
    if (tiles.length === 20) break;
  }
  return tiles;
}

export function storedScope(value: unknown): string[] | null {
  const parsed = parsedJson(value);
  if (!Array.isArray(parsed)) return null;
  const domains = parsed.filter((item): item is string => typeof item === "string");
  return domains.length > 0 ? [...new Set(domains)].sort().slice(0, 100) : null;
}

// The points of one audit that finished at or after `since`, oldest first. At most `limit`: when
// more qualify, the newest are kept and `truncated` says so.
export async function readTrendPoints(
  sql: SqlClient,
  tenantId: string,
  auditId: string,
  since: Date,
  limit = MAX_AUDIT_TREND_POINTS,
): Promise<{ points: StoredTrendPoint[]; truncated: boolean }> {
  const result = await sql.query<TrendRecord>(
    `SELECT run_id, audit_version, finished_at, complete, run_scope, tiles, trigger
       FROM prism365_audit_trend
      WHERE tenant_id=$1 AND audit_id=$2 AND finished_at >= $3
      ORDER BY finished_at DESC, run_id DESC
      LIMIT $4`,
    [tenantId, auditId, since.toISOString(), limit + 1],
  );
  const points = result.rows.slice(0, limit).map((row) => ({
    runId: row.run_id,
    finishedAt: new Date(row.finished_at).toISOString(),
    version: row.audit_version,
    complete: row.complete,
    scope: storedScope(row.run_scope),
    trigger: auditRunTriggerSchema.catch("manual").parse(row.trigger),
    tiles: storedTiles(row.tiles),
  }));
  return { points: points.reverse(), truncated: result.rows.length > limit };
}

const scopeKey = (scope: readonly string[] | null) => (scope ?? []).join("\n");

// Why a point does not compare with the reference point: another version of the audit, another
// scope (a different set of domains, or limited against unlimited), or a partial result.
export function trendDifferences(
  point: StoredTrendPoint,
  reference: StoredTrendPoint,
): AuditTrendDifference[] {
  const differences: AuditTrendDifference[] = [];
  if (point.version !== reference.version) differences.push("version");
  if (scopeKey(point.scope) !== scopeKey(reference.scope)) differences.push("scope");
  if (!point.complete) differences.push("incomplete");
  return differences;
}

// The trend as the API returns it. Every point is compared with the newest one, which is what
// the administrator sees on the audit's page today; the tiles that can be followed are the
// newest point's.
export function buildAuditTrend(input: {
  auditId: string;
  days: number;
  points: readonly StoredTrendPoint[];
  truncated: boolean;
}): AuditTrend {
  const newest = input.points.at(-1);
  return auditTrendSchema.parse({
    auditId: input.auditId,
    days: input.days,
    points: input.points.map((point) => ({
      runId: point.runId,
      finishedAt: point.finishedAt,
      version: point.version,
      complete: point.complete,
      scopeLimited: point.scope !== null,
      trigger: point.trigger,
      tiles: Object.fromEntries(point.tiles.map((tile) => [tile.id, tile.value])),
      notComparable: newest ? trendDifferences(point, newest) : [],
    })),
    tiles: (newest?.tiles ?? []).map(({ id, label, tone }) => ({ id, label, tone })),
    truncated: input.truncated,
  });
}

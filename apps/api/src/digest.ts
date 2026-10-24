import {
  builtInAudits,
  getRun,
  latestRuns as latestAuditRuns,
  listRuns,
  readAuditSchedule,
  storedScope,
  storedTiles,
  type TrendTile,
} from "@prism365/audits";
import {
  type AuditRow,
  auditPersonCellSchema,
  cleanText,
  DIGEST_DEADLINE_DAYS,
  DIGEST_FLAGGED_CHANGES,
  DIGEST_SECTIONS,
  DIGEST_INCIDENTS,
  DIGEST_OPEN_FINDINGS,
  DIGEST_POSTS,
  DIGEST_QUIET_ACCOUNTS,
  type Digest,
  type DigestAuditDrift,
  type DigestCollection,
  type DigestDays,
  type DigestDifference,
  type DigestDirectoryChanges,
  type DigestDriftAudit,
  type DigestDriftTile,
  type DigestFinding,
  type DigestMicrosoft,
  type DigestQuietAccount,
  type DigestQuietAccounts,
  type DigestSectionName,
  type DigestSeverityCounts,
  type DigestSignInFindings,
  personIdSchema,
  type Severity,
} from "@prism365/contracts";
import {
  changeCoverage,
  DIRECTORY_CHANGES_MODULE,
  eventTable,
  GAP_RULE as DIRECTORY_GAP_RULE,
} from "@prism365/module-directory-changes";
import {
  HEALTH_MODULE,
  INCIDENT_KIND,
  MESSAGE_KIND,
  MESSAGES_MODULE,
} from "@prism365/module-service-health";
import {
  COLLECTION_GAP_RULE as SIGNINS_GAP_RULE,
  MODULE_ID as SIGNINS_MODULE,
} from "@prism365/module-signins-audit";
import { latestRuns as latestModuleRuns } from "@prism365/modules";
import type { SqlClient } from "@prism365/storage";

// The "This week" page (docs/this-week.md): what Prism365 stored in the last 1, 7, 14 or 30
// days. Every function here reads stored rows through one SqlClient and nothing else. None calls
// Microsoft, and none writes. A range runs from `from` (included) to `to` (not included), so a
// row stamped exactly at the start counts and one stamped at the moment of the request does not.

const DAY = 86_400_000;

export interface DigestWindow {
  days: DigestDays;
  from: Date;
  to: Date;
}

export function digestWindow(now: Date, days: DigestDays): DigestWindow {
  return { days, from: new Date(now.getTime() - days * DAY), to: now };
}

// Whether a section's source can be read, and whether it ever ran. A source that is off says why.
// One that never ran and has stored nothing is unavailable rather than empty: nothing happened
// and nothing was looked at must not read the same.
export interface DigestSource {
  reason: string | null;
  collected: boolean;
}

const iso = (value: Date | string) => new Date(value).toISOString();
const plural = (n: number, one: string, many: string) => (n === 1 ? one : many);
const sum = (counts: Record<string, number>) => Object.values(counts).reduce((a, b) => a + b, 0);

// "this week", "in the last 14 days": how a note names the range it describes.
export function rangeWords(days: DigestDays): string {
  if (days === 7) return "this week";
  return days === 1 ? "in the last day" : `in the last ${days} days`;
}

const severityRank = `CASE severity WHEN 'critical' THEN 4 WHEN 'high' THEN 3 WHEN 'medium' THEN 2
  WHEN 'low' THEN 1 ELSE 0 END`;

const noSeverities = (): DigestSeverityCounts => ({
  critical: 0,
  high: 0,
  medium: 0,
  low: 0,
  info: 0,
});

// Findings ------------------------------------------------------------------------------------

interface FindingRow {
  id: string;
  severity: Severity;
  subject_kind: string;
  subject_id: string;
  subject_label: string | null;
  title: string;
  first_seen: Date | string;
}

const findingColumns = "id, severity, subject_kind, subject_id, subject_label, title, first_seen";

function toFinding(row: FindingRow): DigestFinding {
  const personId =
    row.subject_kind === "person" && personIdSchema.safeParse(row.subject_id).success
      ? row.subject_id
      : null;
  return {
    id: row.id,
    title: cleanText(row.title),
    severity: row.severity,
    subjectLabel: cleanText(row.subject_label ?? row.subject_id),
    personId,
    firstSeen: iso(row.first_seen),
  };
}

// Findings of one module that the section is about. A module's collection gaps are findings too,
// but they say something about collection, not about the tenant, and are counted under
// Collection instead.
async function countOpenedBySeverity(
  sql: SqlClient,
  tenantId: string,
  module: string,
  gapRule: string,
  window: DigestWindow,
): Promise<DigestSeverityCounts> {
  const result = await sql.query<{ severity: Severity; n: number }>(
    `SELECT severity, count(*)::int AS n FROM prism365_finding
      WHERE tenant_id=$1 AND module=$2 AND rule_id <> $3 AND first_seen >= $4 AND first_seen < $5
      GROUP BY severity`,
    [tenantId, module, gapRule, window.from.toISOString(), window.to.toISOString()],
  );
  const counts = noSeverities();
  for (const row of result.rows) if (row.severity in counts) counts[row.severity] = row.n;
  return counts;
}

async function scalar(sql: SqlClient, text: string, params: readonly unknown[]): Promise<number> {
  const result = await sql.query<{ n: number }>(text, params);
  return result.rows[0]?.n ?? 0;
}

// Sign-in findings -----------------------------------------------------------------------------

export async function signInFindingsSection(
  sql: SqlClient,
  tenantId: string,
  window: DigestWindow,
  source: DigestSource,
): Promise<DigestSignInFindings> {
  const unavailable = (note: string): DigestSignInFindings => ({
    state: "unavailable",
    note,
    opened: noSeverities(),
    open: [],
    openTotal: 0,
    resolved: 0,
  });
  if (source.reason) return unavailable(source.reason);
  const opened = await countOpenedBySeverity(
    sql,
    tenantId,
    SIGNINS_MODULE,
    SIGNINS_GAP_RULE,
    window,
  );
  const open = await sql.query<FindingRow>(
    `SELECT ${findingColumns} FROM prism365_finding
      WHERE tenant_id=$1 AND module=$2 AND rule_id <> $3 AND state='open'
      ORDER BY ${severityRank} DESC, score DESC, first_seen DESC, id
      LIMIT $4`,
    [tenantId, SIGNINS_MODULE, SIGNINS_GAP_RULE, DIGEST_OPEN_FINDINGS],
  );
  const openTotal = await scalar(
    sql,
    `SELECT count(*)::int AS n FROM prism365_finding
      WHERE tenant_id=$1 AND module=$2 AND rule_id <> $3 AND state='open'`,
    [tenantId, SIGNINS_MODULE, SIGNINS_GAP_RULE],
  );
  const resolved = await scalar(
    sql,
    `SELECT count(*)::int AS n FROM prism365_finding
      WHERE tenant_id=$1 AND module=$2 AND rule_id <> $3 AND state='resolved'
        AND state_changed_at >= $4 AND state_changed_at < $5`,
    [
      tenantId,
      SIGNINS_MODULE,
      SIGNINS_GAP_RULE,
      window.from.toISOString(),
      window.to.toISOString(),
    ],
  );
  const openedTotal = sum(opened);
  const range = rangeWords(window.days);
  if (openedTotal === 0 && openTotal === 0 && resolved === 0) {
    if (!source.collected)
      return unavailable(
        "Sign-ins have not been collected yet. Start collecting on the Sign-in audit page.",
      );
    return {
      state: "empty",
      note: `No new sign-in findings ${range}.`,
      opened,
      open: [],
      openTotal,
      resolved,
    };
  }
  const first =
    openedTotal > 0
      ? `${openedTotal} sign-in ${plural(openedTotal, "finding", "findings")} opened ${range}.`
      : `No new sign-in findings ${range}.`;
  const rest = [
    openTotal > 0 ? `${openTotal} still open` : null,
    resolved > 0 ? `${resolved} resolved ${range}` : null,
  ].filter((part): part is string => part !== null);
  return {
    state: "ok",
    note: rest.length > 0 ? `${first} ${rest.join(", ")}.` : first,
    opened,
    open: open.rows.map(toFinding),
    openTotal,
    resolved,
  };
}

// Directory changes ----------------------------------------------------------------------------

export async function directoryChangesSection(
  sql: SqlClient,
  tenantId: string,
  window: DigestWindow,
  source: Pick<DigestSource, "reason">,
): Promise<DigestDirectoryChanges> {
  const nothing = { total: 0, categories: [], flagged: [], flaggedTotal: 0 };
  const unavailable = (note: string): DigestDirectoryChanges => ({
    state: "unavailable",
    note,
    ...nothing,
    coverage: { from: null, through: null, lastCollected: null, phase: "none" },
  });
  if (source.reason) return unavailable(source.reason);
  const bounds = [tenantId, window.from.toISOString(), window.to.toISOString()];
  const total = await scalar(
    sql,
    `SELECT count(*)::int AS n FROM ${eventTable}
      WHERE tenant_id=$1 AND activity_at >= $2 AND activity_at < $3`,
    bounds,
  );
  const categories = await sql.query<{ category: string | null; n: number }>(
    `SELECT category, count(*)::int AS n FROM ${eventTable}
      WHERE tenant_id=$1 AND activity_at >= $2 AND activity_at < $3
      GROUP BY category ORDER BY count(*) DESC, category LIMIT 50`,
    bounds,
  );
  const flagged = await sql.query<FindingRow>(
    `SELECT ${findingColumns} FROM prism365_finding
      WHERE tenant_id=$1 AND module=$4 AND rule_id <> $5 AND first_seen >= $2 AND first_seen < $3
      ORDER BY ${severityRank} DESC, first_seen DESC, id
      LIMIT $6`,
    [...bounds, DIRECTORY_CHANGES_MODULE, DIRECTORY_GAP_RULE, DIGEST_FLAGGED_CHANGES],
  );
  const flaggedTotal = await scalar(
    sql,
    `SELECT count(*)::int AS n FROM prism365_finding
      WHERE tenant_id=$1 AND module=$4 AND rule_id <> $5 AND first_seen >= $2 AND first_seen < $3`,
    [...bounds, DIRECTORY_CHANGES_MODULE, DIRECTORY_GAP_RULE],
  );
  const coverage = await changeCoverage(sql, tenantId);
  // Nothing stored and no place in the log: collection never read it, which is not the same as a
  // range with no changes in it.
  if (total === 0 && flaggedTotal === 0 && coverage.from === null)
    return unavailable(
      "Directory changes have not been collected yet. Start collecting with the directory audit log enabled.",
    );
  const range = rangeWords(window.days);
  const caveats = [
    coverage.from !== null && Date.parse(coverage.from) > window.from.getTime()
      ? "Collection started partway through this range."
      : null,
    coverage.phase === "history"
      ? "History is still loading, so newer changes may be missing."
      : null,
  ].filter((part): part is string => part !== null);
  const counted = {
    total,
    categories: categories.rows.map((row) => ({
      category: row.category === null ? null : cleanText(row.category),
      count: row.n,
    })),
    flagged: flagged.rows.map(toFinding),
    flaggedTotal,
    coverage,
  };
  if (total === 0 && flaggedTotal === 0)
    return {
      state: "empty",
      note: [`No directory changes ${range}.`, ...caveats].join(" "),
      ...counted,
    };
  return {
    state: "ok",
    note: [
      `${total} ${plural(total, "change", "changes")} ${range}${flaggedTotal > 0 ? `, ${flaggedTotal} flagged` : ""}.`,
      ...caveats,
    ].join(" "),
    ...counted,
  };
}

// Accounts that went quiet ---------------------------------------------------------------------

const INACTIVE_AUDIT = "accounts.inactive";
// The audit's "30+ days" view, which is also the view it opens on. Its rows carry this tag once
// they have been quiet that long, and every longer range's tag as well.
const QUIET_VIEW = "d30";

// Two results compare directly when they come from the same version of the audit, cover the same
// account domains, and Microsoft answered in full for both. The same three reasons a trend point
// is drawn hollow.
export function resultDifferences(
  a: { version: number; scope: readonly string[] | null; complete: boolean },
  b: { version: number; scope: readonly string[] | null; complete: boolean },
): DigestDifference[] {
  const key = (scope: readonly string[] | null) => [...(scope ?? [])].sort().join("\n");
  const differences: DigestDifference[] = [];
  if (a.version !== b.version) differences.push("version");
  if (key(a.scope) !== key(b.scope)) differences.push("scope");
  if (!a.complete || !b.complete) differences.push("incomplete");
  return differences;
}

// Rows were stored by the audit as it was when they were written, so what they hold is read with
// care: a cell that is not what this build expects is left out, never an error.
function quietAccount(row: AuditRow): DigestQuietAccount {
  const person = auditPersonCellSchema.safeParse(row.cells?.person);
  const days = row.cells?.days;
  return {
    personId: typeof row.personId === "string" ? row.personId : null,
    name: person.success && person.data.name ? cleanText(person.data.name) : null,
    principal: person.success && person.data.principal ? cleanText(person.data.principal) : null,
    days: typeof days === "number" && Number.isInteger(days) && days >= 0 ? days : null,
  };
}

const isQuiet = (row: AuditRow) => Array.isArray(row.tags) && row.tags.includes(QUIET_VIEW);

// Longest quiet first; accounts without a day count, or a name, last.
function byQuiet(a: DigestQuietAccount & { key: string }, b: DigestQuietAccount & { key: string }) {
  if (a.days !== b.days) return a.days === null ? 1 : b.days === null ? -1 : b.days - a.days;
  if (a.name !== b.name)
    return a.name === null ? 1 : b.name === null ? -1 : a.name < b.name ? -1 : 1;
  return a.key < b.key ? -1 : a.key > b.key ? 1 : 0;
}

export async function quietAccountsSection(
  sql: SqlClient,
  tenantId: string,
  window: DigestWindow,
): Promise<DigestQuietAccounts> {
  const base = { auditId: INACTIVE_AUDIT, view: QUIET_VIEW };
  const unavailable = (note: string): DigestQuietAccounts => ({
    ...base,
    state: "unavailable",
    note,
    total: 0,
    items: [],
    latest: null,
    earlier: null,
    differences: [],
  });
  // Newest first. A result that has not started yet (a clock that stepped back) is not one.
  const finished = (await listRuns(sql, tenantId, INACTIVE_AUDIT, 50)).filter(
    (run) => run.status === "succeeded" && Date.parse(run.startedAt) < window.to.getTime(),
  );
  const latest = finished[0];
  if (!latest)
    return unavailable(
      "Inactive accounts has not finished a run yet. Run it, or schedule it daily or weekly, to see accounts that went quiet.",
    );
  const earlier = finished.find((run) => Date.parse(run.startedAt) <= window.from.getTime());
  if (!earlier)
    return unavailable(
      "Inactive accounts has no kept result from before this range to compare with. Schedule it daily or weekly: Prism365 keeps the newest 10 results, so a daily schedule reaches back 10 days and a weekly one 10 weeks.",
    );
  if (earlier.id === latest.id)
    return unavailable(
      "The newest Inactive accounts result is older than this range, so there is nothing newer to compare it with. Run it again, or schedule it daily or weekly.",
    );
  const newest = await getRun(sql, tenantId, latest.id);
  const before = await getRun(sql, tenantId, earlier.id);
  if (!newest || !before)
    return unavailable(
      "A stored Inactive accounts result was removed while this page was being built. Reload to try again.",
    );
  const wasQuiet = new Set(before.rows.filter(isQuiet).map((row) => row.key));
  const quiet = newest.rows
    .filter((row) => isQuiet(row) && !wasQuiet.has(row.key))
    .map((row) => ({ key: String(row.key), ...quietAccount(row) }))
    .sort(byQuiet);
  const differences = resultDifferences(
    { version: earlier.auditVersion, scope: earlier.scope, complete: earlier.complete },
    { version: latest.auditVersion, scope: latest.scope, complete: latest.complete },
  );
  const found = {
    ...base,
    total: quiet.length,
    items: quiet.slice(0, DIGEST_QUIET_ACCOUNTS).map(({ key: _key, ...account }) => account),
    latest: { runId: latest.id, at: latest.startedAt },
    earlier: { runId: earlier.id, at: earlier.startedAt },
    differences,
  };
  const caveat =
    differences.length > 0
      ? " The two results do not compare directly, so some accounts may be listed only because of that."
      : "";
  return quiet.length === 0
    ? {
        ...found,
        state: "empty",
        note: `No accounts passed 30 days without a sign-in since the earlier result.${caveat}`,
      }
    : {
        ...found,
        state: "ok",
        note: `${quiet.length} ${plural(quiet.length, "account", "accounts")} passed 30 days without a sign-in since the earlier result.${caveat}`,
      };
}

// Audit drift ----------------------------------------------------------------------------------

interface TrendRow {
  audit_id: string;
  run_id: string;
  audit_version: number;
  finished_at: Date | string;
  complete: boolean;
  run_scope: unknown;
  tiles: unknown;
}

interface Point {
  runId: string;
  at: string;
  version: number;
  complete: boolean;
  scope: string[] | null;
  tiles: TrendTile[];
}

function toPoint(row: TrendRow): Point {
  return {
    runId: row.run_id,
    at: iso(row.finished_at),
    version: row.audit_version,
    complete: row.complete,
    scope: storedScope(row.run_scope),
    tiles: storedTiles(row.tiles),
  };
}

// One point per audit: the newest that finished before (or, with `inclusive`, at) the bound.
async function newestPoints(
  sql: SqlClient,
  tenantId: string,
  bound: Date,
  inclusive: boolean,
): Promise<Map<string, Point>> {
  const result = await sql.query<TrendRow>(
    `SELECT DISTINCT ON (audit_id) audit_id, run_id, audit_version, finished_at, complete,
            run_scope, tiles
       FROM prism365_audit_trend
      WHERE tenant_id=$1 AND finished_at ${inclusive ? "<=" : "<"} $2
      ORDER BY audit_id, finished_at DESC, run_id DESC`,
    [tenantId, bound.toISOString()],
  );
  return new Map(result.rows.map((row) => [row.audit_id, toPoint(row)]));
}

const toneRank = { critical: 0, warning: 1, good: 2, neutral: 3 } as const;

export async function auditDriftSection(
  sql: SqlClient,
  tenantId: string,
  window: DigestWindow,
): Promise<DigestAuditDrift> {
  const latestPoints = await newestPoints(sql, tenantId, window.to, false);
  const startPoints = await newestPoints(sql, tenantId, window.from, true);
  let compared = 0;
  let uncompared = 0;
  const audits: (DigestDriftAudit & { rank: number; order: number })[] = [];
  for (const [order, audit] of builtInAudits.entries()) {
    const id = audit.definition.id;
    const to = latestPoints.get(id);
    if (!to) continue;
    const from = startPoints.get(id);
    // No point from before the range, or no newer one: there is no change to measure.
    if (!from || from.runId === to.runId) {
      uncompared += 1;
      continue;
    }
    compared += 1;
    const before = new Map(from.tiles.map((tile) => [tile.id, tile.value]));
    const tiles: DigestDriftTile[] = [];
    for (const tile of to.tiles) {
      const was = before.get(tile.id);
      // A tile only one of the two results has is not compared: it exists because the audit's
      // wording or scope changed, not because a number moved.
      if (was === undefined || was === tile.value) continue;
      tiles.push({
        id: tile.id,
        label: cleanText(tile.label),
        tone: tile.tone,
        from: was,
        to: tile.value,
        change: tile.value - was,
      });
    }
    if (tiles.length === 0) continue;
    tiles.sort((a, b) => toneRank[a.tone] - toneRank[b.tone]);
    audits.push({
      auditId: id,
      title: audit.definition.title,
      from: { runId: from.runId, at: from.at },
      to: { runId: to.runId, at: to.at },
      differences: resultDifferences(from, to),
      tiles,
      rank: Math.min(...tiles.map((tile) => toneRank[tile.tone])),
      order,
    });
  }
  audits.sort((a, b) => a.rank - b.rank || a.order - b.order);
  const shown = audits.map(({ rank: _rank, order: _order, ...audit }) => audit);
  const counts = { compared, uncompared, audits: shown };
  if (compared === 0)
    return {
      state: "unavailable",
      note:
        uncompared === 0
          ? "No audit has a kept result yet, so there is nothing to compare. Run audits, or schedule them daily or weekly, to see how their numbers change."
          : "No audit has a kept result from before this range and a newer one to compare. Schedule audits daily or weekly to see how their numbers change.",
      ...counts,
    };
  const left = uncompared > 0 ? ` ${uncompared} could not be compared.` : "";
  const range = rangeWords(window.days);
  return shown.length === 0
    ? { state: "empty", note: `No audit tile changed ${range}.${left}`, ...counts }
    : {
        state: "ok",
        note: `${shown.length} ${plural(shown.length, "audit", "audits")} changed ${range}.${left}`,
        ...counts,
      };
}

// Microsoft ------------------------------------------------------------------------------------

interface AttentionRow {
  id: string;
  title: string;
  services: unknown;
  severity: Severity;
  provider_status: string | null;
  provider_deadline: Date | string | null;
  first_seen: Date | string;
}

function services(value: unknown): string[] {
  let parsed = value;
  if (typeof value === "string") {
    try {
      parsed = JSON.parse(value) as unknown;
    } catch {
      return [];
    }
  }
  return Array.isArray(parsed)
    ? parsed
        .filter((item): item is string => typeof item === "string")
        .map((item) => cleanText(item))
        .slice(0, 20)
    : [];
}

// What the notification list calls open: still active at Microsoft, not marked done, not muted.
const openAttention = "active AND review_state <> 'done' AND NOT muted";

type PostsPart = DigestMicrosoft["posts"];
type IncidentsPart = DigestMicrosoft["incidents"];

async function postsPart(
  sql: SqlClient,
  tenantId: string,
  window: DigestWindow,
  source: DigestSource,
): Promise<PostsPart> {
  const unavailable = (note: string): PostsPart => ({
    state: "unavailable",
    note,
    total: 0,
    items: [],
  });
  if (source.reason) return unavailable(source.reason);
  const reach = new Date(window.to.getTime() + DIGEST_DEADLINE_DAYS * DAY);
  const bounds = [
    tenantId,
    MESSAGES_MODULE,
    MESSAGE_KIND,
    window.to.toISOString(),
    reach.toISOString(),
  ];
  const rows = await sql.query<AttentionRow>(
    `SELECT id, title, services, severity, provider_status, provider_deadline, first_seen
       FROM prism365_attention
      WHERE tenant_id=$1 AND module=$2 AND kind=$3 AND ${openAttention}
        AND provider_deadline >= $4 AND provider_deadline < $5
      ORDER BY provider_deadline, id
      LIMIT ${DIGEST_POSTS}`,
    bounds,
  );
  const total = await scalar(
    sql,
    `SELECT count(*)::int AS n FROM prism365_attention
      WHERE tenant_id=$1 AND module=$2 AND kind=$3 AND ${openAttention}
        AND provider_deadline >= $4 AND provider_deadline < $5`,
    bounds,
  );
  if (total === 0 && !source.collected) {
    const stored = await scalar(
      sql,
      "SELECT count(*)::int AS n FROM prism365_attention WHERE tenant_id=$1 AND module=$2",
      [tenantId, MESSAGES_MODULE],
    );
    if (stored === 0)
      return unavailable(
        "Message center has not been checked yet. Start collecting with Message center enabled.",
      );
  }
  return {
    state: total > 0 ? "ok" : "empty",
    note:
      total > 0
        ? `${total} Message center ${plural(total, "post has", "posts have")} an act-by date in the next ${DIGEST_DEADLINE_DAYS} days.`
        : `No Message center posts have an act-by date in the next ${DIGEST_DEADLINE_DAYS} days.`,
    total,
    items: rows.rows.flatMap((row) =>
      row.provider_deadline
        ? [
            {
              id: row.id,
              title: cleanText(row.title),
              actBy: iso(row.provider_deadline),
              services: services(row.services),
            },
          ]
        : [],
    ),
  };
}

async function incidentsPart(
  sql: SqlClient,
  tenantId: string,
  window: DigestWindow,
  source: DigestSource,
): Promise<IncidentsPart> {
  const unavailable = (note: string): IncidentsPart => ({
    state: "unavailable",
    note,
    opened: 0,
    resolved: 0,
    openTotal: 0,
    open: [],
  });
  if (source.reason) return unavailable(source.reason);
  const kind = [tenantId, HEALTH_MODULE, INCIDENT_KIND];
  const bounds = [...kind, window.from.toISOString(), window.to.toISOString()];
  // Opened is when Prism365 first saw the incident; resolved is when Microsoft closed it, or
  // stopped listing it, after that.
  const opened = await scalar(
    sql,
    `SELECT count(*)::int AS n FROM prism365_attention
      WHERE tenant_id=$1 AND module=$2 AND kind=$3 AND first_seen >= $4 AND first_seen < $5`,
    bounds,
  );
  const resolved = await scalar(
    sql,
    `SELECT count(*)::int AS n FROM prism365_attention
      WHERE tenant_id=$1 AND module=$2 AND kind=$3 AND NOT active
        AND changed_at >= $4 AND changed_at < $5`,
    bounds,
  );
  const open = await sql.query<AttentionRow>(
    `SELECT id, title, services, severity, provider_status, provider_deadline, first_seen
       FROM prism365_attention
      WHERE tenant_id=$1 AND module=$2 AND kind=$3 AND ${openAttention}
      ORDER BY ${severityRank} DESC, first_seen DESC, id
      LIMIT ${DIGEST_INCIDENTS}`,
    kind,
  );
  const openTotal = await scalar(
    sql,
    `SELECT count(*)::int AS n FROM prism365_attention
      WHERE tenant_id=$1 AND module=$2 AND kind=$3 AND ${openAttention}`,
    kind,
  );
  const stored = await scalar(
    sql,
    "SELECT count(*)::int AS n FROM prism365_attention WHERE tenant_id=$1 AND module=$2 AND kind=$3",
    kind,
  );
  if (stored === 0 && !source.collected)
    return unavailable(
      "Service health has not been checked yet. Start collecting with Service health enabled.",
    );
  const any = opened > 0 || resolved > 0 || openTotal > 0;
  return {
    state: any ? "ok" : "empty",
    note: any
      ? `${opened} opened and ${resolved} resolved ${rangeWords(window.days)}${openTotal > 0 ? `, ${openTotal} open now` : ""}.`
      : `Microsoft reported no incidents ${rangeWords(window.days)}.`,
    opened,
    resolved,
    openTotal,
    open: open.rows.map((row) => ({
      id: row.id,
      title: cleanText(row.title),
      severity: row.severity,
      status: row.provider_status === null ? null : cleanText(row.provider_status),
      services: services(row.services),
      firstSeen: iso(row.first_seen),
    })),
  };
}

export async function microsoftSection(
  sql: SqlClient,
  tenantId: string,
  window: DigestWindow,
  sources: { health: DigestSource; messages: DigestSource },
): Promise<DigestMicrosoft> {
  const posts = await postsPart(sql, tenantId, window, sources.messages);
  const incidents = await incidentsPart(sql, tenantId, window, sources.health);
  const state =
    posts.state === "unavailable" && incidents.state === "unavailable"
      ? "unavailable"
      : posts.state === "ok" || incidents.state === "ok"
        ? "ok"
        : "empty";
  return { state, note: [incidents.note, posts.note].join(" "), posts, incidents };
}

// Collection -----------------------------------------------------------------------------------

export interface DigestModule {
  id: string;
  title: string;
  collectors: readonly { id: string; title: string }[];
}

export async function collectionSection(
  sql: SqlClient,
  tenantId: string,
  window: DigestWindow,
  input: { collecting: boolean; modules: readonly DigestModule[] },
): Promise<DigestCollection> {
  const ids = input.modules.map((module) => module.id);
  const runs = await latestModuleRuns(sql, tenantId);
  const successes = await sql.query<{ module: string; collector: string; at: Date | string }>(
    `SELECT module, collector, max(COALESCE(finished_at, started_at)) AS at
       FROM prism365_module_run
      WHERE tenant_id=$1 AND status='succeeded'
      GROUP BY module, collector`,
    [tenantId],
  );
  const collectors = input.modules.flatMap((module) =>
    module.collectors.map((collector) => {
      const run = runs.find(
        (candidate) => candidate.module === module.id && candidate.collector === collector.id,
      );
      const success = successes.rows.find(
        (row) => row.module === module.id && row.collector === collector.id,
      );
      return {
        module: module.id,
        moduleTitle: module.title,
        collector: collector.id,
        collectorTitle: collector.title,
        lastSuccess: success ? iso(success.at) : null,
        lastRun: run ? { status: run.status, at: run.startedAt } : null,
      };
    }),
  );
  const bounds = [tenantId, ids, window.from.toISOString(), window.to.toISOString()];
  const failedRuns = await scalar(
    sql,
    `SELECT count(*)::int AS n FROM prism365_module_run
      WHERE tenant_id=$1 AND status='failed' AND module = ANY($2::text[])
        AND started_at >= $3 AND started_at < $4`,
    bounds,
  );
  const categories = await sql.query<{ error_category: string }>(
    `SELECT error_category FROM prism365_module_run
      WHERE tenant_id=$1 AND status='failed' AND module = ANY($2::text[])
        AND started_at >= $3 AND started_at < $4 AND error_category IS NOT NULL
      GROUP BY error_category
      ORDER BY max(started_at) DESC, error_category
      LIMIT 5`,
    bounds,
  );
  const gaps = await scalar(
    sql,
    `SELECT count(*)::int AS n FROM prism365_finding
      WHERE tenant_id=$1 AND first_seen >= $6 AND first_seen < $7
        AND ((module=$2 AND rule_id=$3) OR (module=$4 AND rule_id=$5))`,
    [
      tenantId,
      DIRECTORY_CHANGES_MODULE,
      DIRECTORY_GAP_RULE,
      SIGNINS_MODULE,
      SIGNINS_GAP_RULE,
      window.from.toISOString(),
      window.to.toISOString(),
    ],
  );
  const schedule = await readAuditSchedule(sql, tenantId, builtInAudits);
  const lastRuns = await latestAuditRuns(sql, tenantId);
  const audits = builtInAudits.flatMap((audit) => {
    const frequency = schedule[audit.definition.id];
    if (frequency !== "daily" && frequency !== "weekly") return [];
    const run = lastRuns.find((candidate) => candidate.auditId === audit.definition.id);
    return [
      {
        auditId: audit.definition.id,
        title: audit.definition.title,
        frequency,
        lastRun: run ? { status: run.status, at: run.startedAt } : null,
      },
    ];
  });
  const range = rangeWords(window.days);
  const note = [
    input.collecting
      ? "Collecting."
      : "Collection is off. Nothing new is gathered until an administrator starts it.",
    failedRuns > 0
      ? `${failedRuns} collector ${plural(failedRuns, "run", "runs")} failed ${range}.`
      : null,
    gaps > 0 ? `${gaps} ${plural(gaps, "gap", "gaps")} recorded ${range}.` : null,
  ]
    .filter((part): part is string => part !== null)
    .join(" ");
  const nothing = collectors.length === 0 && audits.length === 0;
  return {
    state: nothing ? "empty" : "ok",
    note: nothing ? "Nothing is set up to collect." : note,
    collecting: input.collecting,
    collectors,
    failedRuns,
    errors: categories.rows.map((row) => row.error_category),
    gaps,
    audits,
  };
}

// The whole page ---------------------------------------------------------------------------------

export type DigestSourceName = "signins" | "changes" | "health" | "messages";

export interface DigestInput {
  tenantId: string;
  now: Date;
  days: DigestDays;
  collecting: boolean;
  modules: readonly DigestModule[];
  // Why a module's section cannot be read, by source. A source not named here is available.
  unavailable: Partial<Record<DigestSourceName, string>>;
}

export async function buildDigest(sql: SqlClient, input: DigestInput): Promise<Digest> {
  const window = digestWindow(input.now, input.days);
  const { tenantId } = input;
  const runs = await latestModuleRuns(sql, tenantId);
  const source = (name: DigestSourceName, module: string): DigestSource => ({
    reason: input.unavailable[name] ?? null,
    collected: runs.some((run) => run.module === module),
  });
  return {
    days: input.days,
    from: window.from.toISOString(),
    generatedAt: window.to.toISOString(),
    signInFindings: await signInFindingsSection(
      sql,
      tenantId,
      window,
      source("signins", SIGNINS_MODULE),
    ),
    directoryChanges: await directoryChangesSection(
      sql,
      tenantId,
      window,
      source("changes", DIRECTORY_CHANGES_MODULE),
    ),
    quietAccounts: await quietAccountsSection(sql, tenantId, window),
    auditDrift: await auditDriftSection(sql, tenantId, window),
    microsoft: await microsoftSection(sql, tenantId, window, {
      health: source("health", HEALTH_MODULE),
      messages: source("messages", MESSAGES_MODULE),
    }),
    collection: await collectionSection(sql, tenantId, window, {
      collecting: input.collecting,
      modules: input.modules,
    }),
  };
}

// The sections a read touches, named as the access ledger records them. A section whose source is
// off is not read; the two Microsoft sources share one section, which is read if either is on.
export function digestSectionsRead(
  unavailable: Partial<Record<DigestSourceName, string>>,
): DigestSectionName[] {
  return DIGEST_SECTIONS.filter((name) => {
    if (name === "signInFindings") return !unavailable.signins;
    if (name === "directoryChanges") return !unavailable.changes;
    if (name === "microsoft") return !unavailable.health || !unavailable.messages;
    return true;
  });
}

// What the page gets when there is no module storage to read: every section, none readable.
export function unavailableDigest(now: Date, days: DigestDays, reason: string): Digest {
  const window = digestWindow(now, days);
  const gone = { state: "unavailable", note: reason } as const;
  return {
    days,
    from: window.from.toISOString(),
    generatedAt: window.to.toISOString(),
    signInFindings: {
      ...gone,
      opened: noSeverities(),
      open: [],
      openTotal: 0,
      resolved: 0,
    },
    directoryChanges: {
      ...gone,
      total: 0,
      categories: [],
      flagged: [],
      flaggedTotal: 0,
      coverage: { from: null, through: null, lastCollected: null, phase: "none" },
    },
    quietAccounts: {
      ...gone,
      auditId: INACTIVE_AUDIT,
      view: QUIET_VIEW,
      total: 0,
      items: [],
      latest: null,
      earlier: null,
      differences: [],
    },
    auditDrift: { ...gone, compared: 0, uncompared: 0, audits: [] },
    microsoft: {
      ...gone,
      posts: { ...gone, total: 0, items: [] },
      incidents: { ...gone, opened: 0, resolved: 0, openTotal: 0, open: [] },
    },
    collection: {
      ...gone,
      collecting: false,
      collectors: [],
      failedRuns: 0,
      errors: [],
      gaps: 0,
      audits: [],
    },
  };
}

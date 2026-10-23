import { GraphBoundaryError, GraphRequestError, GraphRetryDeferredError } from "@prism365/graph";
import {
  type CollectorContext,
  type CollectorResult,
  type FindingInput,
  FindingPolicyError,
} from "@prism365/modules";
import { z } from "zod";
import type { DirectoryChangesConfig } from "./config.ts";
import {
  DAY,
  type DirectoryChangesCursor,
  HISTORY_DAYS,
  MAX_SPAN_MS,
  MIN_SPAN_MS,
  MINUTE,
  parseCursor,
} from "./cursor.ts";
import { DIRECTORY_CHANGES_MODULE, type EventRow, eventColumns, eventFromRow } from "./query.ts";
import { judgeChange, RULES_VERSION } from "./rules.ts";
import {
  changeSubjects,
  type DirectoryChangeEvent,
  directoryAuditContext,
  readDirectoryAudit,
} from "./source.ts";

export const MODULE_ID = DIRECTORY_CHANGES_MODULE;
export const COLLECTOR_ID = "audits";
export const GAP_RULE = "collection.gap";

// $top for each request. Microsoft documents $top for directory audits without a maximum, so
// the first size stays under the 1,000 records the Graph client accepts per page. A page Microsoft
// refuses (HTTP 400) or one larger than the client's size budget is read again with the second
// size for the rest of the run.
export const PAGE_SIZES = [999, 100] as const;
// Microsoft allows 5 directory audit requests per 10 seconds per app and tenant; requests in a run
// start at least this far apart. 429 answers are retried by the Graph client.
export const PAGE_INTERVAL_MS = 2_000;
// Pages one window may take before it counts as too busy, whatever the record count.
const MAX_PAGES = 100;
// A minimum-size window that cannot be read is retried on this many runs, then recorded as a gap.
const TIMEOUT_STRIKES = 3;
const DETECTION_BATCH = 500;

const iso = (value: Date | number) => new Date(value).toISOString();

const pageSchema = z.object({
  value: z.array(z.unknown()).max(1_000),
  "@odata.nextLink": z.string().max(16_384).optional(),
});

export interface CollectorOptions {
  // Waits between requests. Tests pass one that returns at once.
  sleep?: (milliseconds: number) => Promise<void>;
  // Epoch milliseconds for the run's time budget and request pacing.
  clock?: () => number;
}

interface Run {
  context: CollectorContext;
  config: DirectoryChangesConfig;
  clock: () => number;
  sleep: (milliseconds: number) => Promise<void>;
  deadline: number;
  lastRequest: number | null;
  size: number;
  complete: boolean;
}

interface Position {
  phase: "history" | "live";
  start: number;
  through: number;
  span: number;
  trying: number | null;
  strikes: { at: string; count: number } | undefined;
}

type WindowStatus = "complete" | "overflow" | "stopped";

// Windows sit on a grid of their own size, so a busy minute is the same window on every run and
// a gap recorded for it is one finding, not one per run.
function windowEnd(from: number, span: number, now: number): number {
  return Math.min(Math.floor(from / span) * span + span, now);
}

async function pace(run: Run): Promise<void> {
  if (run.lastRequest !== null) {
    const wait = run.lastRequest + PAGE_INTERVAL_MS - run.clock();
    if (wait > 0) await run.sleep(wait);
  }
  run.lastRequest = run.clock();
}

async function insertEvents(
  context: CollectorContext,
  events: readonly DirectoryChangeEvent[],
): Promise<number> {
  if (events.length === 0) return 0;
  const eventsTable = context.table("event");
  const subjectsTable = context.table("subject");
  return context.db(async (sql) => {
    let inserted = 0;
    for (const event of events) {
      const result = await sql.query(
        `INSERT INTO ${eventsTable} (tenant_id, id, activity_at, category, activity, operation_type,
           result, result_reason, logged_by_service, correlation_id, initiated_by_kind,
           initiated_by_id, initiated_by_name, initiated_by_principal, targets, details,
           processed, collected_at)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15::jsonb,$16::jsonb,false,$17)
         ON CONFLICT (tenant_id, id) DO NOTHING`,
        [
          context.tenantId,
          event.id,
          iso(event.activityAt),
          event.category,
          event.activity,
          event.operationType,
          event.result,
          event.resultReason,
          event.loggedByService,
          event.correlationId,
          event.initiatedBy.kind,
          event.initiatedBy.id,
          event.initiatedBy.name,
          event.initiatedBy.principal,
          JSON.stringify(event.targets),
          JSON.stringify(event.details),
          iso(context.now),
        ],
      );
      // A record already stored is a repeat: it changes nothing, including its subjects.
      if (result.rowCount === 0) continue;
      inserted += 1;
      for (const subject of changeSubjects(event))
        await sql.query(
          `INSERT INTO ${subjectsTable} (tenant_id, event_id, subject_id, subject_kind, subject_type, activity_at)
           VALUES ($1,$2,$3,$4,$5,$6)
           ON CONFLICT (tenant_id, event_id, subject_id, subject_kind) DO NOTHING`,
          [
            context.tenantId,
            event.id,
            subject.id,
            subject.kind,
            subject.type,
            iso(event.activityAt),
          ],
        );
    }
    return inserted;
  });
}

// Reads one window page by page. Every readable record is stored as it arrives, so what a
// stopped or too-busy window read is kept; the caller decides whether the window counts as read.
// No request starts after the run's deadline, and requests are paced.
async function pageThrough(
  run: Run,
  from: number,
  to: number,
): Promise<{ status: WindowStatus; read: number }> {
  const query = new URLSearchParams({
    $filter: `activityDateTime ge ${iso(from)} and activityDateTime le ${iso(to)}`,
    $top: String(run.size),
  });
  let next: string | undefined = `auditLogs/directoryAudits?${query}`;
  const seen = new Set<string>();
  let read = 0;
  for (let page = 0; next; page += 1) {
    if (page >= MAX_PAGES) return { status: "overflow", read };
    if (run.clock() > run.deadline) return { status: "stopped", read };
    await pace(run);
    if (run.clock() > run.deadline) return { status: "stopped", read };
    if (seen.has(next))
      throw new GraphBoundaryError("Graph pagination repeated a continuation link.", "pagination");
    seen.add(next);
    const body: z.infer<typeof pageSchema> = await run.context.graph.get(
      next,
      directoryAuditContext,
      pageSchema,
    );
    const events: DirectoryChangeEvent[] = [];
    for (const item of body.value) {
      const event = readDirectoryAudit(item);
      if (event) events.push(event);
    }
    read += body.value.length;
    run.context.count.fetched(body.value.length);
    // An unreadable record is left out and counted; it can never stall collection.
    const skipped = body.value.length - events.length;
    if (skipped > 0) {
      run.context.count.skipped(skipped);
      run.complete = false;
    }
    run.context.count.inserted(await insertEvents(run.context, events));
    next = body["@odata.nextLink"];
    if (next && read >= run.config.windowRecordLimit) return { status: "overflow", read };
  }
  return { status: "complete", read };
}

const pageRefused = (error: unknown) =>
  (error instanceof GraphBoundaryError && error.reason === "size") ||
  (error instanceof GraphRequestError &&
    !(error instanceof GraphRetryDeferredError) &&
    error.status === 400);

async function readWindow(run: Run, from: number, to: number) {
  for (;;) {
    try {
      return await pageThrough(run, from, to);
    } catch (error) {
      const smaller = PAGE_SIZES.find((size) => size < run.size);
      if (!pageRefused(error) || smaller === undefined) throw error;
      run.size = smaller;
    }
  }
}

// Names come from the directory and are not Prism365's to vet. If one trips the finding language
// policy (a group called "Productivity Team", say), the finding is recorded without names
// rather than failing the run on the same change forever.
async function upsertSafely(context: CollectorContext, finding: FindingInput): Promise<void> {
  try {
    await context.findings.upsert(finding);
  } catch (error) {
    if (!(error instanceof FindingPolicyError)) throw error;
    const { label: _label, ...subject } = finding.subject;
    await context.findings.upsert({
      ...finding,
      subject,
      summary:
        "A directory change matched this rule. Names were left out because one conflicted with the finding language policy; open the change in the feed for details.",
      evidence: finding.evidence.map((item) => ({
        ...item,
        fields: Object.fromEntries(
          Object.entries(item.fields).filter(
            ([key, value]) =>
              typeof value !== "string" ||
              ["activity", "category", "result", "actorKind", "role"].includes(key),
          ),
        ),
      })),
    });
  }
}

// Judges stored changes that have not been judged yet, oldest first, then marks them. A run that
// stops between storing and judging leaves them for the next one.
async function judgePending(context: CollectorContext): Promise<void> {
  const table = context.table("event");
  for (;;) {
    const pending = await context.db((sql) =>
      sql.query<EventRow>(
        `SELECT ${eventColumns} FROM ${table}
          WHERE tenant_id = $1 AND processed = false
          ORDER BY activity_at, id LIMIT ${DETECTION_BATCH}`,
        [context.tenantId],
      ),
    );
    if (pending.rows.length === 0) return;
    for (const row of pending.rows)
      for (const finding of judgeChange(eventFromRow(row), context.tenantId))
        await upsertSafely(context, finding);
    const ids = pending.rows.map((row) => row.id);
    await context.db((sql) =>
      sql.query(
        `UPDATE ${table} SET processed = true
          WHERE tenant_id = $1 AND id IN (${ids.map((_id, index) => `$${index + 2}`).join(",")})`,
        [context.tenantId, ...ids],
      ),
    );
    if (pending.rows.length < DETECTION_BATCH) return;
  }
}

type GapKind = "overflow" | "unanswered" | "missed";

// A span Prism365 could not read in full. It is recorded as a finding so a hole in the history
// is never silent. Keys use the minute the span starts in: the same busy minute is one finding.
export function gapFinding(
  kind: GapKind,
  from: number,
  to: number,
  detail: { read?: number; limit?: number } = {},
): FindingInput {
  const span = `between ${iso(from)} and ${iso(to)}`;
  const summary = {
    overflow: `Microsoft recorded more directory changes ${span} than Prism365 reads in one window (${detail.limit ?? 0}). ${detail.read ?? 0} were read and kept; the rest are missing from the feed and from findings. Check the Entra audit log for that span: a burst this dense can be a bulk change or an attack.`,
    unanswered: `Prism365 could not read the directory changes ${span}: Microsoft did not answer in time, or the run ran out of time reading them, on ${TIMEOUT_STRIKES} runs in a row. Collection moved past that span, so any changes in it are missing from the feed and from findings. Check the Entra audit log for it.`,
    missed: `Collection was not running ${span}, and Microsoft keeps directory changes for 30 days at most. Changes in that span are missing from the feed and from findings.`,
  }[kind];
  return {
    ruleId: GAP_RULE,
    ruleVersion: RULES_VERSION,
    dedupeKey: `${kind}:${iso(Math.floor(from / MINUTE) * MINUTE)}`,
    purpose: "operations",
    severity: "medium",
    score: 40,
    subject: { kind: "service", id: MODULE_ID, label: "Directory change collection" },
    title:
      kind === "missed"
        ? "Directory changes were missed while collection was off"
        : "Some directory changes could not be read",
    summary,
    evidence: [
      {
        source: "Prism365 directory change collector",
        ref: `${iso(from)}/${iso(to)}`,
        observedAt: iso(to),
        fields: { read: detail.read ?? 0, windowStart: iso(from), windowEnd: iso(to) },
      },
    ],
    limitations: [
      kind === "missed"
        ? "Without Entra ID P1 or P2 Microsoft keeps only 7 days, so more may be missing than this span."
        : "The span may hold fewer changes than it seems; Prism365 could not tell, because it could not read all of them.",
    ],
    window: { start: new Date(from), end: new Date(to) },
    observedAt: new Date(to),
  };
}

// A window that timed out or that a run could not finish: a larger one is halved for the next
// try; a minimum-size one is retried a few runs, then given up, so one stretch cannot stall
// collection.
function windowFailed(position: Position, from: number, end: number): "retry" | "give_up" {
  if (end - from > MIN_SPAN_MS) {
    position.span = Math.max(MIN_SPAN_MS, Math.floor((end - from) / 2));
    return "retry";
  }
  const count = position.strikes?.at === iso(from) ? position.strikes.count + 1 : 1;
  if (count < TIMEOUT_STRIKES) {
    position.strikes = { at: iso(from), count };
    return "retry";
  }
  position.strikes = undefined;
  return "give_up";
}

function toCursor(position: Position): DirectoryChangesCursor {
  return {
    v: 1,
    phase: position.phase,
    start: iso(position.start),
    through: iso(position.through),
    ...(position.span < MAX_SPAN_MS ? { span: position.span } : {}),
    ...(position.trying !== null ? { trying: iso(position.trying) } : {}),
    ...(position.strikes ? { strikes: position.strikes } : {}),
  };
}

const isTimeout = (error: unknown) =>
  error instanceof GraphBoundaryError && error.reason === "timeout";

// One collector run. The first loads history, oldest first, from 30 days back (Microsoft's
// longest retention), asking to run again within seconds while history remains. Later runs read
// from where the last one stopped, less the late-arrival margin, up to now; records already
// stored change nothing. The position moves only past windows read completely, and is saved
// after each one, so a slow or stopped run leaves a backlog for the next, never a hole.
export async function collectDirectoryChanges(
  context: CollectorContext,
  config: DirectoryChangesConfig,
  options: CollectorOptions = {},
): Promise<CollectorResult> {
  const clock = options.clock ?? Date.now;
  const run: Run = {
    context,
    config,
    clock,
    sleep:
      options.sleep ?? ((milliseconds) => new Promise((done) => setTimeout(done, milliseconds))),
    deadline: clock() + config.runBudgetSeconds * 1_000,
    lastRequest: null,
    size: PAGE_SIZES[0],
    complete: true,
  };
  const now = context.now.getTime();
  const oldest = now - HISTORY_DAYS * DAY;

  // Changes a previous run stored but did not get to judge.
  await judgePending(context);

  const saved = parseCursor(await context.cursor.get());
  const position: Position = saved
    ? {
        phase: saved.phase,
        start: Date.parse(saved.start),
        // A position in the future (a clock that ran ahead) is brought back to now.
        through: Math.min(Date.parse(saved.through), now),
        span: saved.span ?? MAX_SPAN_MS,
        trying: saved.trying ? Date.parse(saved.trying) : null,
        strikes: saved.strikes,
      }
    : {
        phase: "history",
        start: oldest,
        through: oldest,
        span: MAX_SPAN_MS,
        trying: null,
        strikes: undefined,
      };
  position.start = Math.min(position.start, position.through);
  const persist = () => context.cursor.set(JSON.stringify(toCursor(position)));
  const advance = (end: number) => {
    position.through = Math.max(position.through, end);
    if (position.phase === "history" && position.through >= now) position.phase = "live";
    if (position.strikes && Date.parse(position.strikes.at) < end) position.strikes = undefined;
  };

  // Collection was off for longer than Microsoft keeps the log: what fell out of it is gone.
  if (position.through < oldest) {
    await upsertSafely(context, gapFinding("missed", position.through, oldest));
    run.complete = false;
    position.through = oldest;
  }

  // The last run was stopped while reading a window: that window failed.
  if (position.trying !== null) {
    const from = position.trying;
    const end = windowEnd(from, position.span, now);
    position.trying = null;
    if (from < now && windowFailed(position, from, end) === "give_up") {
      await upsertSafely(context, gapFinding("unanswered", from, end));
      run.complete = false;
      advance(end);
    }
  }
  await persist();

  let place =
    position.phase === "history"
      ? position.through
      : Math.max(position.through - config.lateArrivalMinutes * MINUTE, oldest);
  let first = true;
  while (place < now) {
    const from = place;
    const end = windowEnd(from, position.span, now);
    // Saved before reading, so a run stopped mid-window is visible to the next one.
    position.trying = from;
    await persist();
    let outcome: { status: WindowStatus | "given_up"; read: number };
    try {
      outcome = await readWindow(run, from, end);
    } catch (error) {
      position.trying = null;
      if (!isTimeout(error) || windowFailed(position, from, end) === "retry") {
        await persist();
        throw error;
      }
      await upsertSafely(context, gapFinding("unanswered", from, end));
      run.complete = false;
      outcome = { status: "given_up", read: 0 };
    }
    position.trying = null;

    if (outcome.status === "stopped") {
      // The run's budget ran out inside this window. When it was the run's first window, the
      // window cannot be read in one run at this size, so the next run tries a smaller one.
      if (!first || windowFailed(position, from, end) === "retry") {
        await persist();
        await judgePending(context);
        return { complete: run.complete, more: true };
      }
      await upsertSafely(context, gapFinding("unanswered", from, end));
      run.complete = false;
    } else if (outcome.status === "overflow") {
      if (end - from > MIN_SPAN_MS) {
        // Too many records for one window: read the same place again in a smaller one.
        position.span = Math.max(MIN_SPAN_MS, Math.floor((end - from) / 2));
        await persist();
        continue;
      }
      await upsertSafely(
        context,
        gapFinding("overflow", from, end, { read: outcome.read, limit: config.windowRecordLimit }),
      );
      run.complete = false;
    } else if (outcome.status === "complete" && outcome.read * 2 < config.windowRecordLimit) {
      // A quiet window: grow back toward full-size windows.
      position.span = Math.min(MAX_SPAN_MS, position.span * 2);
    }

    place = end;
    first = false;
    advance(end);
    await persist();
    await judgePending(context);
  }
  advance(now);
  await persist();
  return { complete: run.complete, more: false };
}

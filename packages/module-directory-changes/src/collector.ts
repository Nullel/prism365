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
// Window size of the best-effort re-read of recent records that have not settled yet.
const FRESH_WINDOW_MS = 30 * MINUTE;

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

// One lane's reading position and window state (see cursor.ts).
interface Lane {
  at: number;
  span: number;
  // Windows stay at `span` until the lane passes this time: the end of a window that was too
  // busy or too slow, so the next window does not grow straight back into it.
  hold: number | null;
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

// The next window at this place is half of one that could not be read whole, and windows stay
// that small until the lane is past it.
function shrink(lane: Lane, from: number, end: number): void {
  lane.span = Math.max(MIN_SPAN_MS, Math.floor((end - from) / 2));
  lane.hold = end;
}

// A window that timed out or that a run could not finish: a larger one is halved for the next
// try; a minimum-size one is retried a few runs, then given up, so one stretch cannot stall
// collection.
function windowFailed(lane: Lane, from: number, end: number): "retry" | "give_up" {
  if (end - from > MIN_SPAN_MS) {
    shrink(lane, from, end);
    return "retry";
  }
  const count = lane.strikes?.at === iso(from) ? lane.strikes.count + 1 : 1;
  if (count < TIMEOUT_STRIKES) {
    lane.strikes = { at: iso(from), count };
    return "retry";
  }
  lane.strikes = undefined;
  return "give_up";
}

type LaneState = Omit<DirectoryChangesCursor["settle"], "at">;

function toLane(at: string, state: LaneState): Lane {
  return {
    at: Date.parse(at),
    span: state.span ?? MAX_SPAN_MS,
    hold: state.hold ? Date.parse(state.hold) : null,
    trying: state.trying ? Date.parse(state.trying) : null,
    strikes: state.strikes,
  };
}

function laneState(lane: Lane): LaneState {
  return {
    ...(lane.span < MAX_SPAN_MS ? { span: lane.span } : {}),
    ...(lane.hold !== null ? { hold: iso(lane.hold) } : {}),
    ...(lane.trying !== null ? { trying: iso(lane.trying) } : {}),
    ...(lane.strikes ? { strikes: lane.strikes } : {}),
  };
}

const freshLane = (at: number): Lane => ({
  at,
  span: MAX_SPAN_MS,
  hold: null,
  trying: null,
  strikes: undefined,
});

const isTimeout = (error: unknown) =>
  error instanceof GraphBoundaryError && error.reason === "timeout";

// One collector run. The first loads history, oldest first, from 30 days back (Microsoft's
// longest retention), asking to run again within seconds while history remains. Later runs read
// new records on the main lane, then re-read behind it on the settle lane once records are the
// late-arrival margin old, so records Microsoft publishes late are collected; records already
// stored change nothing. A lane moves only past windows read completely, or past a window
// recorded as a gap, and saves its place after each, so a slow or stopped run leaves a backlog
// for the next, never a silent hole.
export async function collectDirectoryChanges(
  context: CollectorContext,
  config: DirectoryChangesConfig,
  options: CollectorOptions = {},
): Promise<CollectorResult> {
  const clock = options.clock ?? Date.now;
  const began = clock();
  const budget = config.runBudgetSeconds * 1_000;
  const run: Run = {
    context,
    config,
    clock,
    sleep:
      options.sleep ?? ((milliseconds) => new Promise((done) => setTimeout(done, milliseconds))),
    deadline: began + budget,
    lastRequest: null,
    size: PAGE_SIZES[0],
    complete: true,
  };
  const now = context.now.getTime();
  const oldest = now - HISTORY_DAYS * DAY;
  // Everything before this has been published for at least the late-arrival margin.
  const settledBy = now - config.lateArrivalMinutes * MINUTE;

  // Changes a previous run stored but did not get to judge.
  await judgePending(context);

  const saved = parseCursor(await context.cursor.get());
  let phase = saved?.phase ?? "history";
  let start = saved ? Date.parse(saved.start) : oldest;
  const main = saved ? toLane(saved.through, saved) : freshLane(oldest);
  const settle = saved ? toLane(saved.settle.at, saved.settle) : freshLane(oldest);
  // A position in the future (a clock that ran ahead) is brought back to now.
  main.at = Math.min(main.at, now);
  settle.at = Math.min(settle.at, main.at);
  const persist = () =>
    context.cursor.set(
      JSON.stringify({
        v: 1,
        phase,
        start: iso(start),
        through: iso(main.at),
        ...laneState(main),
        settle: { at: iso(settle.at), ...laneState(settle) },
      } satisfies DirectoryChangesCursor),
    );

  // Collection was off for longer than Microsoft keeps the log: what fell out of it is gone.
  if (main.at < oldest) {
    await upsertSafely(context, gapFinding("missed", main.at, oldest));
    run.complete = false;
    main.at = oldest;
  }
  // Nothing older than Microsoft's retention can be read again, late or not.
  settle.at = Math.max(settle.at, Math.min(oldest, main.at));
  start = Math.min(start, settle.at);

  const advance = (lane: Lane, end: number) => {
    lane.at = Math.max(lane.at, end);
    if (lane.strikes && Date.parse(lane.strikes.at) < end) lane.strikes = undefined;
    if (lane.hold !== null && end >= lane.hold) lane.hold = null;
    if (phase === "history" && main.at >= now) phase = "live";
  };

  // A lane still marked as reading a window was stopped there by the last run: that window failed.
  for (const lane of [main, settle]) {
    if (lane.trying === null) continue;
    const from = lane.trying;
    const end = windowEnd(from, lane.span, now);
    lane.trying = null;
    if (from < now && windowFailed(lane, from, end) === "give_up") {
      await upsertSafely(context, gapFinding("unanswered", from, end));
      run.complete = false;
      if (from <= lane.at) advance(lane, end);
    }
  }
  settle.at = Math.min(settle.at, main.at);
  await persist();
  // Where the main lane stood when this run began: earlier runs read everything before it.
  const savedThrough = main.at;

  // Reads one lane forward to `until`, one window at a time, saving its place after each.
  // Returns false when the run's time ran out first.
  const walk = async (
    lane: Lane,
    until: number,
    deadline: number,
    afterRead?: (from: number, end: number) => void,
  ): Promise<boolean> => {
    run.deadline = deadline;
    while (lane.at < until) {
      const from = lane.at;
      const end = windowEnd(from, lane.span, until);
      // A window that had less than half the run's budget and still could not be finished
      // proves nothing about its size; only one that had the time counts as failed.
      const hadTime = deadline - clock() >= budget / 2;
      // Saved before reading, so a run stopped mid-window is visible to the next one.
      lane.trying = from;
      await persist();
      let outcome: { status: WindowStatus | "given_up"; read: number };
      try {
        outcome = await readWindow(run, from, end);
      } catch (error) {
        lane.trying = null;
        if (!isTimeout(error) || windowFailed(lane, from, end) === "retry") {
          await persist();
          throw error;
        }
        await upsertSafely(context, gapFinding("unanswered", from, end));
        run.complete = false;
        outcome = { status: "given_up", read: 0 };
      }
      lane.trying = null;

      if (outcome.status === "stopped") {
        // The time ran out inside this window. When the window had the time to itself, it
        // cannot be read in one run at this size, so the next run tries a smaller one.
        if (!hadTime || windowFailed(lane, from, end) === "retry") {
          await persist();
          return false;
        }
        await upsertSafely(context, gapFinding("unanswered", from, end));
        run.complete = false;
      } else if (outcome.status === "overflow") {
        if (end - from > MIN_SPAN_MS) {
          // Too many records for one window: read the same place again in a smaller one.
          shrink(lane, from, end);
          await persist();
          continue;
        }
        await upsertSafely(
          context,
          gapFinding("overflow", from, end, {
            read: outcome.read,
            limit: config.windowRecordLimit,
          }),
        );
        run.complete = false;
      } else if (outcome.status === "complete") {
        afterRead?.(from, end);
        // Past the busy stretch and quiet again: grow back toward full-size windows.
        if (lane.hold === null && outcome.read * 2 < config.windowRecordLimit)
          lane.span = Math.min(MAX_SPAN_MS, lane.span * 2);
      }
      advance(lane, end);
      await persist();
      await judgePending(context);
    }
    return true;
  };

  // A window the main lane read once it was already old enough is settled by that read: the
  // settle lane does not read it again when it is contiguous with what is already settled.
  const settleBehind = (from: number, end: number) => {
    if (settle.at >= from) settle.at = Math.max(settle.at, Math.min(end, settledBy));
  };
  const backlog = () => Math.min(main.at, settledBy) - settle.at;

  // New records first. When the settle lane is behind, the main lane leaves it a third of the
  // run.
  const mainDeadline =
    phase === "live" && backlog() > 0 ? began + Math.floor((budget * 2) / 3) : began + budget;
  const caughtUp = await walk(main, now, mainDeadline, settleBehind);
  await persist();
  if (phase === "history") return { complete: run.complete, more: !caughtUp };

  // Whether the settle lane ran out of time with more to read now.
  let settleMore = false;
  try {
    settleMore = !(await walk(settle, Math.min(main.at, settledBy), began + budget));
  } catch (error) {
    // A slow re-read is already recorded on the lane (a smaller window, or a strike) and is tried
    // again on the next scheduled run; the new records this run read are kept and the run stays
    // a success.
    if (!isTimeout(error)) throw error;
  }
  await persist();

  // Best effort, newest first: what earlier runs read that has not settled yet is read again
  // with the time left, so a record Microsoft published a few minutes late shows up now rather
  // than when it settles. It saves no position; the settle lane is what guarantees every such
  // record is read once it is the margin old.
  run.deadline = began + budget;
  const floor = Math.max(settle.at, oldest);
  for (let to = savedThrough; to > floor && clock() <= run.deadline; ) {
    const from = Math.max(floor, to - FRESH_WINDOW_MS);
    try {
      if ((await readWindow(run, from, to)).status === "stopped") break;
    } catch (error) {
      if (isTimeout(error)) break;
      throw error;
    }
    to = from;
  }
  await judgePending(context);
  return { complete: run.complete, more: !caughtUp || settleMore };
}

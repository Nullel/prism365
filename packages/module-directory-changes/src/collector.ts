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
import { addedByPim, judgeChange, PIM_ACTIVATION, RULES_VERSION } from "./rules.ts";
import {
  changeSubjects,
  type DirectoryChangeEvent,
  directoryAuditContext,
  readDirectoryAudit,
  searchText,
} from "./source.ts";

export const MODULE_ID = DIRECTORY_CHANGES_MODULE;
export const COLLECTOR_ID = "audits";
export const GAP_RULE = "collection.gap";

// $top for each request. Microsoft documents $top for directory audits without a maximum, so
// the first size stays under the 1,000 records the Graph client accepts per page. A page Microsoft
// refuses (HTTP 400) or one larger than the client's size budget is read again with the second
// size for the rest of the run.
export const PAGE_SIZES = [999, 100] as const;
// Microsoft allows 5 directory audit requests per 10 seconds per app and tenant. Each request the
// module makes starts at least this long after the previous one finished, across runs, so no 10
// seconds hold more than 5 of them. When Microsoft answers 429, the Graph client retries inside
// that one request after the Retry-After Microsoft gives; those retries are not paced here.
export const PAGE_INTERVAL_MS = 2_200;
// Pages one window may take before it counts as too busy, whatever the record count.
const MAX_PAGES = 100;
// A window that cannot be split and still cannot be read is tried this many times, over at least
// STRIKE_SPREAD_MS, then recorded as a gap. The spread keeps a burst of failures that has nothing
// to do with the window (Microsoft failing a third of requests for a few minutes, say) from
// using up all three tries within seconds.
const STRIKES = 3;
const STRIKE_SPREAD_MS = 10 * MINUTE;
const DETECTION_BATCH = 500;
// Window size of the best-effort re-read of recent records that have not settled yet.
const FRESH_WINDOW_MS = 30 * MINUTE;
// A gap finding that a later complete read covered is retitled to this, so it is not retitled
// again on every run.
export const HEALED_TITLE = "A gap in directory changes was read later";

const iso = (value: Date | number) => new Date(value).toISOString();

const pageSchema = z.object({
  value: z.array(z.unknown()).max(1_000),
  "@odata.nextLink": z.string().max(16_384).optional(),
});

// When the last request finished, kept across runs of one module so a run the scheduler starts
// again at once still waits its turn.
export interface Pacer {
  last: number | null;
}

export interface CollectorOptions {
  // Waits between requests. Tests pass one that returns at once.
  sleep?: (milliseconds: number) => Promise<void>;
  // Epoch milliseconds for the run's time budget and request pacing.
  clock?: () => number;
  pacer?: Pacer;
}

interface Run {
  context: CollectorContext;
  config: DirectoryChangesConfig;
  clock: () => number;
  sleep: (milliseconds: number) => Promise<void>;
  deadline: number;
  pacer: Pacer;
  size: number;
  complete: boolean;
  // Pages Microsoft has answered in this run.
  answers: number;
}

// One lane's reading position and window state (see cursor.ts).
interface Lane {
  at: number;
  span: number;
  // Windows stay at `span` until the lane passes this time: the end of a window that was too
  // busy or too slow, so the next window does not grow straight back into it.
  hold: number | null;
  trying: number | null;
  strikes: { at: string; count: number; since?: string | undefined } | undefined;
}

type WindowStatus = "complete" | "overflow" | "stopped";

// What one window read: how it ended, how many records came back, and when the readable ones
// happened.
interface WindowRead {
  status: WindowStatus;
  read: number;
  times: number[];
}

// Why a window could not be read: Microsoft did not answer in time; kept answering with a server
// error or asking to wait longer than the client waits; refused the request; sent an answer that
// cannot be read (too large, malformed, or paging in a loop); or the run ran out of time.
type Failure = "slow" | "failing" | "refused" | "unreadable" | "stopped";

// Windows sit on a grid of their own size, so a busy minute is the same window on every run and
// a gap recorded for it is one finding, not one per run.
function windowEnd(from: number, span: number, now: number): number {
  return Math.min(Math.floor(from / span) * span + span, now);
}

async function paced<T>(run: Run, request: () => Promise<T>): Promise<T> {
  if (run.pacer.last !== null) {
    // Never longer than one interval: a clock that stepped back must not hold collection.
    const wait = Math.min(PAGE_INTERVAL_MS, run.pacer.last + PAGE_INTERVAL_MS - run.clock());
    if (wait > 0) await run.sleep(wait);
  }
  try {
    return await request();
  } finally {
    run.pacer.last = run.clock();
  }
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
           search_text, processed, collected_at)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15::jsonb,$16::jsonb,$17,false,$18)
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
          searchText(event),
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
async function pageThrough(run: Run, from: number, to: number): Promise<WindowRead> {
  const query = new URLSearchParams({
    $filter: `activityDateTime ge ${iso(from)} and activityDateTime le ${iso(to)}`,
    $top: String(run.size),
  });
  let next: string | undefined = `auditLogs/directoryAudits?${query}`;
  const seen = new Set<string>();
  let read = 0;
  const times: number[] = [];
  for (let page = 0; next; page += 1) {
    if (page >= MAX_PAGES) return { status: "overflow", read, times };
    if (run.clock() > run.deadline) return { status: "stopped", read, times };
    if (seen.has(next))
      throw new GraphBoundaryError("Graph pagination repeated a continuation link.", "pagination");
    seen.add(next);
    const link: string = next;
    const body = await paced(
      run,
      async (): Promise<z.infer<typeof pageSchema> | null> =>
        // The wait can carry the run past its deadline; then nothing is requested.
        run.clock() > run.deadline
          ? null
          : run.context.graph.get(link, directoryAuditContext, pageSchema),
    );
    if (!body) return { status: "stopped", read, times };
    run.answers += 1;
    const events: DirectoryChangeEvent[] = [];
    for (const item of body.value) {
      const event = readDirectoryAudit(item);
      if (event) {
        events.push(event);
        times.push(event.activityAt.getTime());
      }
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
    if (next && read >= run.config.windowRecordLimit) return { status: "overflow", read, times };
  }
  return { status: "complete", read, times };
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

// Answers that cannot be read as a page of records, whatever the window.
const unreadableAnswers: ReadonlySet<string> = new Set([
  "size",
  "shape",
  "item",
  "pagination",
  "continuation",
  "json",
  "content_type",
  "empty",
  "page_limit",
  "redirect",
]);

// A failure that belongs to the window or the moment rather than to access: the window is retried
// smaller, then given up with a gap. Only what says Prism365 may not read (401, 403), has no
// credential, or is not Microsoft's answer at all (storage, the module's own boundary) holds the
// lane where it is and fails the run.
function windowFailure(error: unknown): Exclude<Failure, "stopped"> | null {
  if (error instanceof GraphBoundaryError) {
    if (error.reason === "timeout") return "slow";
    if (error.reason === "retry_budget") return "failing";
    return unreadableAnswers.has(error.reason) ? "unreadable" : null;
  }
  if (error instanceof GraphRetryDeferredError) return "failing";
  if (error instanceof GraphRequestError) {
    if (error.status === 401 || error.status === 403) return null;
    if (error.status >= 500) return "failing";
    if (error.status >= 400) return "refused";
  }
  return null;
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
    const events = pending.rows.map(eventFromRow);
    // A directory record PIM wrote counts as an activation only beside a stored PIM activation
    // record with the same correlation id. PIM writes both within seconds, so the lookup stays
    // within a day of them and uses the time index.
    const byPim = events.filter((event) => addedByPim(event) && event.correlationId !== null);
    const correlations = [...new Set(byPim.map((event) => event.correlationId as string))];
    const activationCorrelations = new Set<string>();
    if (correlations.length > 0) {
      const times = byPim.map((event) => event.activityAt.getTime());
      const found = await context.db((sql) =>
        sql.query<{ correlation_id: string }>(
          `SELECT DISTINCT correlation_id FROM ${table}
            WHERE tenant_id = $1 AND activity_at >= $2 AND activity_at <= $3
              AND lower(activity) = $4
              AND correlation_id IN (${correlations.map((_id, index) => `$${index + 5}`).join(",")})`,
          [
            context.tenantId,
            iso(Math.min(...times) - DAY),
            iso(Math.max(...times) + DAY),
            PIM_ACTIVATION,
            ...correlations,
          ],
        ),
      );
      for (const row of found.rows) activationCorrelations.add(row.correlation_id);
    }
    for (const event of events)
      for (const finding of judgeChange(event, context.tenantId, { activationCorrelations }))
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

type GapKind = "overflow" | "unanswered" | "missed" | "unsettled";

const minuteKey = (at: number) => iso(Math.floor(at / MINUTE) * MINUTE);

// A span Prism365 could not read in full. It is recorded as a finding so a hole in the history
// is never silent. Keys use the minute the span starts in, so the same busy minute is one
// finding; a span that fell out of Microsoft's retention is keyed by where it began, so it is one
// finding that widens.
export function gapFinding(
  kind: GapKind,
  from: number,
  to: number,
  detail: { read?: number; limit?: number; failure?: Failure } = {},
): FindingInput {
  const span = `between ${iso(from)} and ${iso(to)}`;
  const why = {
    slow: "Microsoft did not answer in time",
    failing: "Microsoft kept answering with an error or asking Prism365 to wait",
    refused: "Microsoft refused the request for them",
    unreadable:
      "Microsoft's answer for them could not be read (too large, malformed or paging in a loop)",
    stopped: "the run ran out of time reading them",
  }[detail.failure ?? "slow"];
  const summary = {
    overflow: `Microsoft recorded more directory changes ${span} than Prism365 reads in one window (${detail.limit ?? 0}), and the span cannot be split further. ${detail.read ?? 0} were read and kept; the rest are missing from the feed and from findings. Check the Entra audit log for that span: a burst this dense can be a bulk change or an attack.`,
    unanswered: `Prism365 could not read the directory changes ${span}: ${why}, on ${STRIKES} tries. Collection moved past that span, so any changes in it are missing from the feed and from findings until a later read returns them. Check the Entra audit log for it.`,
    missed: `Prism365 had not read the directory changes ${span} when Microsoft stopped keeping them (30 days at most), because collection was stopped or every read failed. Changes in that span are missing from the feed and from findings.`,
    unsettled: `Prism365 read the directory changes ${span} once, but Microsoft stopped keeping them before Prism365 could read them again, because collection was stopped or every read failed. A change Microsoft published late in that span may be missing from the feed and from findings.`,
  }[kind];
  return {
    ruleId: GAP_RULE,
    ruleVersion: RULES_VERSION,
    dedupeKey: `${kind}:${kind === "missed" || kind === "unsettled" ? iso(from) : minuteKey(from)}`,
    purpose: "operations",
    severity: "medium",
    score: 40,
    subject: { kind: "service", id: MODULE_ID, label: "Directory change collection" },
    title:
      kind === "missed"
        ? "Directory changes were missed while they could not be read"
        : kind === "unsettled"
          ? "Late directory changes may be missing"
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
      kind === "missed" || kind === "unsettled"
        ? "Without Entra ID P1 or P2 Microsoft keeps only 7 days, so more may be missing than this span."
        : "The span may hold fewer changes than it seems; Prism365 could not tell, because it could not read all of them.",
    ],
    window: { start: new Date(from), end: new Date(to) },
    observedAt: new Date(to),
  };
}

interface OpenGap {
  dedupeKey: string;
  start: number;
  end: number;
}

// A gap that a later complete read covered: the same finding, saying nothing is missing. The
// module cannot change a finding's state, so it says so and leaves resolving it to a person.
function healedFinding(gap: OpenGap, at: number): FindingInput {
  const span = `between ${iso(gap.start)} and ${iso(gap.end)}`;
  return {
    ruleId: GAP_RULE,
    ruleVersion: RULES_VERSION,
    dedupeKey: gap.dedupeKey,
    purpose: "operations",
    severity: "medium",
    score: 40,
    subject: { kind: "service", id: MODULE_ID, label: "Directory change collection" },
    title: HEALED_TITLE,
    summary: `Prism365 could not read the directory changes ${span} at first. It read the whole span later (${iso(at)}), so nothing is missing from it. Resolve this finding.`,
    evidence: [
      {
        source: "Prism365 directory change collector",
        ref: `${iso(gap.start)}/${iso(gap.end)}`,
        observedAt: iso(at),
        fields: { readLater: true, windowStart: iso(gap.start), windowEnd: iso(gap.end) },
      },
    ],
    limitations: ["Read later; nothing is missing from this span."],
    window: { start: new Date(gap.start), end: new Date(gap.end) },
    observedAt: new Date(at),
  };
}

async function openGaps(context: CollectorContext): Promise<OpenGap[]> {
  const result = await context.db((sql) =>
    sql.query<{
      dedupe_key: string;
      window_start: Date | string;
      window_end: Date | string;
    }>(
      `SELECT dedupe_key, window_start, window_end FROM prism365_finding
        WHERE tenant_id = $1 AND module = $2 AND rule_id = $3 AND state <> 'resolved'
          AND title <> $4
        ORDER BY first_seen LIMIT 200`,
      [context.tenantId, MODULE_ID, GAP_RULE, HEALED_TITLE],
    ),
  );
  // Only a span that was never read can be read whole later. A span that held too many records
  // or fell out of Microsoft's retention cannot.
  return result.rows
    .filter((row) => row.dedupe_key.startsWith("unanswered:"))
    .map((row) => ({
      dedupeKey: row.dedupe_key,
      start: new Date(row.window_start).getTime(),
      end: new Date(row.window_end).getTime(),
    }));
}

// The next window at this place is half of one that could not be read whole, and windows stay
// that small until the lane is past it.
function shrink(lane: Lane, from: number, end: number): void {
  lane.span = Math.max(MIN_SPAN_MS, Math.floor((end - from) / 2));
  lane.hold = end;
}

// A window that failed: one that can still be split is halved and tried again; one that cannot
// is tried again until its strikes run out, then given up, so one stretch cannot stall
// collection.
function windowFailed(lane: Lane, from: number, end: number, now: number): "retry" | "give_up" {
  if (end - from > MIN_SPAN_MS) {
    shrink(lane, from, end);
    return "retry";
  }
  const earlier = lane.strikes?.at === iso(from) ? lane.strikes : undefined;
  const count = (earlier?.count ?? 0) + 1;
  const since = earlier?.since ?? iso(now);
  if (count < STRIKES) {
    lane.strikes = { at: iso(from), count, since };
    return "retry";
  }
  // Both bounds are inclusive, so the next window starts on this one's last millisecond, which
  // failed on every try: it has been failing since the same time.
  lane.strikes = { at: iso(end), count: 0, since };
  return "give_up";
}

// Whether the last try of a window waits: its first try was less than the spread ago.
function lastTryWaits(lane: Lane, from: number, now: number): boolean {
  const strikes = lane.strikes;
  if (strikes?.at !== iso(from) || strikes.count < STRIKES - 1 || !strikes.since) return false;
  return now - Date.parse(strikes.since) < STRIKE_SPREAD_MS;
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

const newLane = (at: number): Lane => ({
  at,
  span: MAX_SPAN_MS,
  hold: null,
  trying: null,
  strikes: undefined,
});

// How a walk ended: at its end, or with more to read. `waiting` means it stopped at a window
// whose last try is for a later scheduled run, so the run does not ask to run again for it.
interface WalkResult {
  done: boolean;
  waiting?: boolean;
}

// A span recorded as a gap finding that widens while it grows (see cursor.ts).
interface SpanMark {
  from: number;
  to: number;
}

const toMark = (value: { from: string; to: string } | undefined): SpanMark | null =>
  value ? { from: Date.parse(value.from), to: Date.parse(value.to) } : null;
const fromMark = (mark: SpanMark | null) =>
  mark ? { from: iso(mark.from), to: iso(mark.to) } : undefined;

// One collector run. The first loads history, oldest first, from 30 days back (Microsoft's
// longest retention), asking to run again within seconds while history remains. Later runs read
// new records on the main lane, then read every window again on the settle lane once it is the
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
    pacer: options.pacer ?? { last: null },
    size: PAGE_SIZES[0],
    complete: true,
    answers: 0,
  };
  const now = context.now.getTime();
  const oldest = now - HISTORY_DAYS * DAY;
  const margin = config.lateArrivalMinutes * MINUTE;
  // Everything before this has been published for at least the late-arrival margin.
  const settledBy = now - margin;

  // Changes a previous run stored but did not get to judge.
  await judgePending(context);

  const saved = parseCursor(await context.cursor.get());
  let phase = saved?.phase ?? "history";
  let start = saved ? Date.parse(saved.start) : oldest;
  const main = saved ? toLane(saved.through, saved) : newLane(oldest);
  const settle = saved ? toLane(saved.settle.at, saved.settle) : newLane(oldest);
  let missed = toMark(saved?.missed);
  let unsettled = toMark(saved?.unsettled);
  let fresh = saved?.fresh ? Date.parse(saved.fresh) : null;
  const persist = () =>
    context.cursor.set(
      JSON.stringify({
        v: 1,
        phase,
        start: iso(start),
        through: iso(main.at),
        ...laneState(main),
        settle: { at: iso(settle.at), ...laneState(settle) },
        ...(missed ? { missed: fromMark(missed) } : {}),
        ...(unsettled ? { unsettled: fromMark(unsettled) } : {}),
        ...(fresh !== null ? { fresh: iso(fresh) } : {}),
      } satisfies DirectoryChangesCursor),
    );
  // One finding for a span that keeps growing: a new span that starts within the margin of the
  // last one's end continues it.
  const widen = (mark: SpanMark | null, from: number, to: number): SpanMark =>
    mark && from <= mark.to + margin
      ? { from: Math.min(mark.from, from), to: Math.max(mark.to, to) }
      : { from, to };

  // A position in the future means the clock ran ahead and came back. While it was ahead, the
  // lanes read windows that had not happened yet and passed real time, so changes made since
  // may lie behind them. The main lane comes back to now. The settle lane, which trails by the
  // margin, has passed real time only when the clock ran ahead by more than the margin; then it
  // reads everything Microsoft still keeps again, and repeats change nothing.
  if (main.at > now) {
    const ahead = main.at - now;
    main.at = now;
    main.trying = null;
    settle.trying = null;
    settle.at =
      ahead > margin ? Math.max(start, oldest) : Math.min(settle.at, Math.max(start, settledBy));
    fresh = Math.min(fresh ?? settle.at, settle.at);
  }
  settle.at = Math.min(settle.at, main.at);

  // The settle lane is behind what Microsoft still keeps (both lanes were held, or collection
  // was stopped). What it can no longer read again, the main lane read once; the part it read
  // while it was fresh could have gained records Microsoft published late, and is recorded
  // rather than skipped silently. Changes the main lane read when they were already old (the
  // history load, a catch-up) count as settled.
  const settleFloor = Math.min(oldest, main.at);
  if (settle.at < settleFloor) {
    if (fresh !== null && Math.max(settle.at, fresh) < settleFloor) {
      unsettled = widen(unsettled, Math.max(settle.at, fresh), settleFloor);
      await upsertSafely(context, gapFinding("unsettled", unsettled.from, unsettled.to));
      run.complete = false;
    }
    settle.at = settleFloor;
    fresh = fresh !== null && main.at > settleFloor ? Math.max(fresh, settleFloor) : null;
  }

  // The main lane is behind what Microsoft still keeps. When it had read nothing yet (the first
  // load never got going, say every read was refused), history simply starts later. Otherwise a
  // span went unread: one finding from where it began, widened while the lane stays behind.
  if (main.at < oldest) {
    if (main.at > start) {
      missed = widen(missed, main.at, oldest);
      await upsertSafely(context, gapFinding("missed", missed.from, missed.to));
      run.complete = false;
    } else start = oldest;
    main.at = oldest;
  }
  // Nothing older than Microsoft's retention can be read again, late or not.
  settle.at = Math.max(settle.at, Math.min(oldest, main.at));
  start = Math.min(start, settle.at);

  const advance = (lane: Lane, end: number) => {
    // The main lane passing what is younger than the margin: from here on, the settle lane
    // reads again what it reads.
    if (lane === main && end > settledBy) fresh ??= Math.max(lane.at, settledBy);
    lane.at = Math.max(lane.at, end);
    if (lane.strikes && Date.parse(lane.strikes.at) < end) lane.strikes = undefined;
    if (lane.hold !== null && end >= lane.hold) lane.hold = null;
    if (phase === "history" && main.at >= now) phase = "live";
  };

  const gaps = await openGaps(context);
  // A gap the settle lane later read whole, and that read returned records inside it, is
  // retitled to say nothing is missing. A read that came back with nothing proves nothing: Learn
  // says this API can answer empty under load, which is when windows fail.
  const heal = async (from: number, end: number, times: readonly number[]) => {
    for (const gap of [...gaps]) {
      if (gap.start < from || gap.end > end) continue;
      if (!times.some((at) => at >= gap.start && at <= gap.end)) continue;
      await upsertSafely(context, healedFinding(gap, now));
      gaps.splice(gaps.indexOf(gap), 1);
    }
  };
  const giveUp = async (from: number, end: number, failure: Failure) => {
    const finding = gapFinding("unanswered", from, end, { failure });
    await upsertSafely(context, finding);
    gaps.push({ dedupeKey: finding.dedupeKey, start: from, end });
    run.complete = false;
  };

  // A lane still marked as reading a window was stopped there by the last run: that window failed.
  for (const lane of [main, settle]) {
    if (lane.trying === null) continue;
    const from = lane.trying;
    const end = windowEnd(from, lane.span, now);
    lane.trying = null;
    if (from < now && windowFailed(lane, from, end, now) === "give_up") {
      await giveUp(from, end, "stopped");
      if (from <= lane.at) advance(lane, end);
    }
  }
  settle.at = Math.min(settle.at, main.at);
  await persist();
  // Where the main lane stood when this run began: earlier runs read everything before it.
  const savedThrough = main.at;

  // A window that fails is blamed on the window only when Microsoft answers other requests: for
  // making it smaller, anything answered in this run; for a strike against a window that cannot
  // be split, something answered since the last failure. Without that, a one-millisecond read of
  // now tells. If that fails too, Microsoft is failing rather than the window (an outage, or
  // throttling): the run ends with that failure and changes nothing, so it leaves no gap and
  // the next try waits for the cadence.
  let answeredAtFailure = 0;
  const blameWindow = async (splittable: boolean): Promise<boolean> => {
    const enough = splittable ? 0 : answeredAtFailure;
    try {
      if (run.answers > enough) return true;
      return (await readWindow(run, now, now)).status !== "stopped";
    } catch (error) {
      await persist();
      throw error;
    } finally {
      answeredAtFailure = run.answers;
    }
  };

  // Reads one lane forward to `until`, one window at a time, saving its place after each. A
  // window that failed is tried again at once, smaller, and given up after its strikes, while
  // the run has time.
  const walk = async (lane: Lane, until: number, deadline: number): Promise<WalkResult> => {
    run.deadline = deadline;
    while (lane.at < until) {
      const from = lane.at;
      const end = windowEnd(from, lane.span, until);
      if (lastTryWaits(lane, from, now)) return { done: false, waiting: true };
      // A window that had less than half the run's budget and still could not be finished
      // proves nothing about its size; only one that had the time counts as failed.
      const hadTime = deadline - clock() >= budget / 2;
      // Saved before reading, so a run stopped mid-window is visible to the next one.
      lane.trying = from;
      await persist();
      let outcome: WindowRead | { status: "given_up"; read: 0; times: [] };
      try {
        outcome = await readWindow(run, from, end);
      } catch (error) {
        lane.trying = null;
        const failure = windowFailure(error);
        if (!failure) {
          await persist();
          throw error;
        }
        // Out of time before Microsoft's health was known: nothing is counted.
        if (!(await blameWindow(end - from > MIN_SPAN_MS))) {
          await persist();
          return { done: false };
        }
        if (windowFailed(lane, from, end, now) === "retry") {
          await persist();
          continue;
        }
        await giveUp(from, end, failure);
        outcome = { status: "given_up", read: 0, times: [] };
      }
      lane.trying = null;

      if (outcome.status === "stopped") {
        // The time ran out inside this window. When the window had the time to itself, it
        // cannot be read in one run at this size, so the next run tries a smaller one.
        if (!hadTime || windowFailed(lane, from, end, now) === "retry") {
          await persist();
          return { done: false };
        }
        await giveUp(from, end, "stopped");
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
        if (lane === settle) await heal(from, end, outcome.times);
        // Past the busy stretch and quiet again: grow back toward full-size windows.
        if (lane.hold === null && outcome.read * 2 < config.windowRecordLimit)
          lane.span = Math.min(MAX_SPAN_MS, lane.span * 2);
      }
      advance(lane, end);
      await persist();
      await judgePending(context);
    }
    return { done: true };
  };

  // New records first. When the settle lane is behind, the main lane leaves it a third of the
  // run.
  const backlog = Math.min(main.at, settledBy) - settle.at;
  const mainDeadline =
    phase === "live" && backlog > 0 ? began + Math.floor((budget * 2) / 3) : began + budget;
  const mainWalk = await walk(main, now, mainDeadline);
  await persist();
  // While history loads, the settle lane waits: it reads all of history again afterwards.
  const behind = (walked: WalkResult) => !walked.done && !walked.waiting;
  if (phase === "history") return { complete: run.complete, more: behind(mainWalk) };

  const settleWalk = await walk(settle, Math.min(main.at, settledBy), began + budget);
  await persist();

  // Best effort, newest first: what earlier runs read that has not settled yet is read again
  // with the time left, so a record Microsoft published a few minutes late shows up now rather
  // than when it settles. It saves no position and heals nothing; the settle lane is what
  // guarantees every such record is read once it is the margin old.
  run.deadline = began + budget;
  const floor = Math.max(settle.at, oldest);
  for (let to = savedThrough; to > floor && clock() <= run.deadline; ) {
    const from = Math.max(floor, to - FRESH_WINDOW_MS);
    try {
      if ((await readWindow(run, from, to)).status === "stopped") break;
    } catch (error) {
      if (windowFailure(error)) break;
      throw error;
    }
    to = from;
  }
  await judgePending(context);
  return { complete: run.complete, more: behind(mainWalk) || behind(settleWalk) };
}

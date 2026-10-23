import {
  type AttentionItem,
  type AuditCatalog,
  type AuditRun,
  type AuditRunDetail,
  type CollectionStatus,
  type Finding,
  type FindingState,
  type ModuleList,
  type PersonList,
  type PersonRecord,
  type PersonSigninPage,
  type AttentionSettings,
  type AuditScheduleSettings,
  type AuditSettings,
  auditScheduleSchema,
  auditScopeSchema,
  attentionPreferencesSchema,
  auditCsv,
  findingStateSchema,
  queryAuditRows,
} from "@prism365/contracts";

// The portfolio demo has no server. `installReplay` swaps `window.fetch` for a shim that answers
// `/api/v1/...` from responses recorded off the real API (see scripts/demo/record.ts), so every
// screen shows what Prism365 produced for the invented company in the recording. Nothing here
// runs in the normal build: main.tsx loads this file only when VITE_DEMO is "true".
//
// Reads come from the recording. Writes are simulated in memory so the screens behave: the
// state lives only as long as the page and resets on reload.

export interface RecordedResponse {
  status: number;
  body: unknown;
}

export interface Fixtures {
  // When the recording finished. Every timestamp in `responses` is at or before it.
  recordedAt: string;
  responses: Record<string, RecordedResponse>;
}

export interface DemoResponse {
  status: number;
  contentType: string;
  // An object is sent as JSON; a string is sent as it is (the CSV export).
  body: unknown;
  headers?: Record<string, string>;
}

export interface DemoRequest {
  method: string;
  // Path and query, e.g. "/api/v1/people?days=30".
  url: string;
  body?: unknown;
}

const API = "/api/v1";
const DAY_MS = 86_400_000;
// A run started in the demo reads as finished this long after it started.
export const SIMULATED_RUN_MS = 2_000;
// Only used to parse relative URLs; nothing is ever requested from it.
const BASE = "http://demo.invalid";

export const unavailableBody = {
  error: "demo",
  message: "This is not available in the demo.",
} as const;

// ---------------------------------------------------------------------------------------------
// Keys
// ---------------------------------------------------------------------------------------------

// "GET /api/v1/people?days=30": the method, the path, and the query with its parameters sorted,
// so the order a screen happens to write them in never matters.
export function normalizeKey(method: string, url: string): string {
  const parsed = new URL(url, BASE);
  parsed.searchParams.sort();
  const query = parsed.searchParams.toString();
  return `${method.toUpperCase()} ${parsed.pathname}${query ? `?${query}` : ""}`;
}

// The exact key first, then the same path with no query. null when neither was recorded.
export function matchFixture(
  responses: Record<string, RecordedResponse>,
  method: string,
  url: string,
): RecordedResponse | null {
  const exact = responses[normalizeKey(method, url)];
  if (exact) return exact;
  const parsed = new URL(url, BASE);
  return responses[`${method.toUpperCase()} ${parsed.pathname}`] ?? null;
}

// ---------------------------------------------------------------------------------------------
// Time
// ---------------------------------------------------------------------------------------------

const ISO_DATETIME = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})$/;
const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

// Whole days from the recording to now, never negative. A whole number keeps the time of day, so
// a shifted timestamp is never later than now: the newest thing in the recording reads as at
// most a day old.
export function wholeDaysSince(recordedAt: string, now: Date): number {
  const recorded = Date.parse(recordedAt);
  if (Number.isNaN(recorded)) return 0;
  return Math.max(0, Math.floor((now.getTime() - recorded) / DAY_MS));
}

function addDays(date: string, days: number): string {
  const shifted = new Date(
    Date.UTC(
      Number(date.slice(0, 4)),
      Number(date.slice(5, 7)) - 1,
      Number(date.slice(8, 10)) + days,
    ),
  );
  return shifted.toISOString().slice(0, 10);
}

// Moves one date or timestamp by whole days and leaves its time and offset alone; any other
// string comes back unchanged.
export function shiftString(value: string, days: number): string {
  if (days === 0) return value;
  if (ISO_DATETIME.test(value)) return `${addDays(value.slice(0, 10), days)}${value.slice(10)}`;
  if (ISO_DATE.test(value)) return addDays(value, days);
  return value;
}

// A copy of a recorded value with every ISO-8601 timestamp and date string moved by `days`.
export function shiftTimestamps<T>(value: T, days: number): T {
  if (days === 0) return value;
  if (typeof value === "string") return shiftString(value, days) as T;
  if (Array.isArray(value)) return value.map((item) => shiftTimestamps(item, days)) as T;
  if (value !== null && typeof value === "object")
    return Object.fromEntries(
      Object.entries(value).map(([key, item]) => [key, shiftTimestamps(item, days)]),
    ) as T;
  return value;
}

// ---------------------------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------------------------

interface SimulatedRun {
  id: string;
  auditId: string;
  startedAt: number;
  // The recorded run whose result this one repeats.
  source: AuditRunDetail;
}

export interface DemoState {
  now: () => Date;
  // Already moved to today.
  responses: Record<string, RecordedResponse>;
  attention: AttentionItem[];
  findings: Finding[];
  settings: AttentionSettings | null;
  auditSettings: AuditSettings | null;
  auditSchedule: AuditScheduleSettings | null;
  collection: CollectionStatus | null;
  runs: SimulatedRun[];
}

type Json = Record<string, unknown>;

const rank = { critical: 4, high: 3, medium: 2, low: 1, info: 0 } as const;

function bodyOf<T>(responses: Record<string, RecordedResponse>, key: string): T | null {
  const hit = responses[key];
  return hit && hit.status === 200 ? (hit.body as T) : null;
}

export function createDemoState(fixtures: Fixtures, now: () => Date = () => new Date()): DemoState {
  const days = wholeDaysSince(fixtures.recordedAt, now());
  const responses = shiftTimestamps(fixtures.responses, days);
  const attention = new Map<string, AttentionItem>();
  const findings = new Map<string, Finding>();
  for (const [key, hit] of Object.entries(responses)) {
    if (hit.status !== 200) continue;
    if (key.startsWith("GET /api/v1/attention?"))
      for (const item of (hit.body as { items?: AttentionItem[] }).items ?? [])
        attention.set(item.id, item);
    if (key.startsWith("GET /api/v1/findings?"))
      for (const item of (hit.body as { items?: Finding[] }).items ?? [])
        findings.set(item.id, item);
    if (/^GET \/api\/v1\/people\/[^/?]+$/.test(key))
      for (const item of (hit.body as { findings?: Finding[] }).findings ?? [])
        findings.set(item.id, item);
  }
  const modules = bodyOf<ModuleList>(responses, "GET /api/v1/modules");
  return {
    now,
    responses,
    attention: [...attention.values()],
    findings: [...findings.values()],
    settings: bodyOf<AttentionSettings>(responses, "GET /api/v1/settings/attention"),
    auditSettings: bodyOf<AuditSettings>(responses, "GET /api/v1/settings/audits"),
    auditSchedule: bodyOf<AuditScheduleSettings>(responses, "GET /api/v1/settings/audit-schedule"),
    collection: modules ? { ...modules.collection } : null,
    runs: [],
  };
}

const json = (status: number, body: unknown): DemoResponse => ({
  status,
  contentType: "application/json",
  body,
});
const unavailable = (): DemoResponse => json(404, unavailableBody);
const clone = <T>(value: T): T => structuredClone(value);

// ---------------------------------------------------------------------------------------------
// Notifications
// ---------------------------------------------------------------------------------------------

// Newest and most serious first, unread before read: the order the server lists them in.
function attentionOrder(a: AttentionItem, b: AttentionItem): number {
  return (
    Number(b.unread) - Number(a.unread) ||
    rank[b.severity] - rank[a.severity] ||
    b.changedAt.localeCompare(a.changedAt) ||
    a.id.localeCompare(b.id)
  );
}

export function attentionCounts(items: readonly AttentionItem[]) {
  return {
    unread: items.filter((item) => item.unread && !item.muted).length,
    active: items.filter(
      (item) => item.provider.active && item.review.state !== "done" && !item.muted,
    ).length,
  };
}

function attentionFilter(filter: string, item: AttentionItem): boolean {
  if (filter === "unread") return item.unread && !item.muted;
  if (filter === "muted") return item.muted;
  if (filter === "open") return item.provider.active && item.review.state !== "done" && !item.muted;
  return true;
}

function listAttention(state: DemoState, url: URL): DemoResponse | null {
  const base = Object.entries(state.responses).find(
    ([key, hit]) => key.startsWith("GET /api/v1/attention?") && hit.status === 200,
  )?.[1].body as { system?: Json[]; sources?: Json[] } | undefined;
  if (!base) return null;
  const filter = url.searchParams.get("filter") ?? "open";
  if (!["open", "unread", "all", "muted"].includes(filter))
    return json(400, { error: "invalid_request", message: "Check the filter." });
  const requested = Number(url.searchParams.get("limit") ?? "100");
  const limit = Math.min(Math.max(Number.isFinite(requested) ? requested : 100, 1), 500);
  const items = state.attention
    .filter((item) => attentionFilter(filter, item))
    .sort(attentionOrder)
    .slice(0, limit);
  // Stopping collection raises the same alert the server raises for it.
  const stopped = state.collection?.state === "stopped";
  const system = (base.system ?? []).filter(
    (alert) => (alert as { id?: string }).id !== "collection.stopped",
  );
  if (stopped)
    system.unshift({
      id: "collection.stopped",
      severity: "medium",
      title: "Prism365 is not collecting",
      detail: "Nothing is being checked until an administrator starts collecting.",
      observedAt: null,
      action: { label: "Start collecting", href: "view:signins" },
    });
  return json(200, {
    mode: "live",
    items: clone(items),
    system,
    sources: base.sources ?? [],
    counts: attentionCounts(state.attention),
  });
}

const reviewStates = ["new", "reviewed", "done"] as const;

function changeAttention(
  state: DemoState,
  id: string,
  change: (item: AttentionItem, at: string) => void,
): DemoResponse | null {
  const item = state.attention.find((candidate) => candidate.id === id);
  if (!item) return json(404, { error: "not_found", message: "The item is not available." });
  change(item, state.now().toISOString());
  return null;
}

// ---------------------------------------------------------------------------------------------
// Findings
// ---------------------------------------------------------------------------------------------

function listFindings(state: DemoState, url: URL): DemoResponse {
  const module = url.searchParams.get("module");
  const wanted = url.searchParams.get("state");
  const requested = Number(url.searchParams.get("limit") ?? "50");
  const limit = Math.min(Math.max(Number.isFinite(requested) ? requested : 50, 1), 200);
  const items = state.findings
    .filter((item) => (!module || item.module === module) && (!wanted || item.state === wanted))
    .sort((a, b) => b.lastSeen.localeCompare(a.lastSeen) || a.id.localeCompare(b.id))
    .slice(0, limit);
  return json(200, { mode: "live", items: clone(items) });
}

// Findings keep the state an administrator gave them in this page, wherever they are shown.
function withFindingStates(state: DemoState, findings: Finding[]): Finding[] {
  return findings.map((item) => state.findings.find((known) => known.id === item.id) ?? item);
}

function openFindingCount(state: DemoState, userId: string): number {
  return state.findings.filter(
    (item) =>
      item.module === "signins.audit" &&
      item.subject.kind === "person" &&
      item.subject.id === userId &&
      item.state !== "resolved",
  ).length;
}

// ---------------------------------------------------------------------------------------------
// People
// ---------------------------------------------------------------------------------------------

// The search box asks for lists nobody recorded, so the recorded list for that range is
// searched here the way the server searches it: name, sign-in name, or the exact account id.
function listPeople(state: DemoState, url: URL): DemoResponse | null {
  const days = url.searchParams.get("days") ?? "30";
  const base = bodyOf<PersonList>(state.responses, `GET ${API}/people?days=${days}`);
  if (!base) return null;
  const search = (url.searchParams.get("search") ?? "").trim().toLowerCase().slice(0, 200);
  const items = base.items
    .filter(
      (item) =>
        !search ||
        (item.userPrincipalName ?? "").toLowerCase().includes(search) ||
        (item.displayName ?? "").toLowerCase().includes(search) ||
        item.userId.toLowerCase() === search,
    )
    .map((item) => ({ ...item, openFindings: openFindingCount(state, item.userId) }));
  return json(200, { ...base, items });
}

// The filters on a person's sign-ins ("Successful", "Not successful") are answered from the
// recording of all their sign-ins, which fits on one page. That is the same split the server
// makes: a sign-in is a success when its error code is zero, and failed when it is not.
function personSignins(state: DemoState, id: string, url: URL): DemoResponse | null {
  if (url.searchParams.has("before")) return null;
  const limit = url.searchParams.get("limit") ?? "100";
  const key = normalizeKey("GET", `${API}/people/${id}/signins?result=all&limit=${limit}`);
  const all = bodyOf<PersonSigninPage>(state.responses, key);
  if (!all || all.next !== null) return null;
  const result = url.searchParams.get("result") ?? "all";
  if (!["all", "success", "failed"].includes(result)) return null;
  const items = all.items.filter(
    (item) => result === "all" || (result === "success") === (item.result === "success"),
  );
  return json(200, { items, next: null });
}

function personRecord(state: DemoState, id: string): DemoResponse | null {
  const base = bodyOf<PersonRecord>(state.responses, `GET ${API}/people/${id}`);
  if (!base) return null;
  return json(200, { ...base, findings: clone(withFindingStates(state, base.findings)) });
}

// ---------------------------------------------------------------------------------------------
// Collection
// ---------------------------------------------------------------------------------------------

function listModules(state: DemoState): DemoResponse | null {
  const base = bodyOf<ModuleList>(state.responses, `GET ${API}/modules`);
  if (!base || !state.collection) return null;
  return json(200, { ...base, collection: { ...state.collection } });
}

function setCollection(state: DemoState, collecting: boolean): DemoResponse {
  if (!state.collection)
    return json(409, {
      error: "collection_unavailable",
      message: "Background collection needs live mode and a signed-in administrator.",
    });
  state.collection = collecting
    ? {
        ...state.collection,
        state: "collecting",
        credentialKind: "delegated",
        boundAt: state.now().toISOString(),
        resumes: false,
      }
    : {
        ...state.collection,
        state: "stopped",
        credentialKind: null,
        boundAt: null,
        lastStopReason: "operator_stopped",
        resumes: false,
      };
  return json(200, { ...state.collection });
}

// ---------------------------------------------------------------------------------------------
// Audits
// ---------------------------------------------------------------------------------------------

// A UUID, with a fallback for pages served over plain http, where crypto.randomUUID is missing.
function newId(): string {
  const random = globalThis.crypto?.randomUUID?.bind(globalThis.crypto);
  if (random) return random();
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (char) => {
    const digit = Math.floor(Math.random() * 16);
    return (char === "x" ? digit : (digit & 0x3) | 0x8).toString(16);
  });
}

function runSummary(detail: AuditRunDetail | AuditRun): AuditRun {
  const { rows: _rows, ...summary } = detail as AuditRunDetail;
  return summary;
}

function isRunning(state: DemoState, run: SimulatedRun): boolean {
  return state.now().getTime() < run.startedAt + SIMULATED_RUN_MS;
}

// A started run reads as running, then as the recorded result it repeats.
function simulated(state: DemoState, run: SimulatedRun): AuditRunDetail {
  const startedAt = new Date(run.startedAt).toISOString();
  if (isRunning(state, run))
    return {
      ...runSummary(run.source),
      id: run.id,
      status: "running",
      complete: false,
      startedAt,
      finishedAt: null,
      error: null,
      summary: [],
      notes: [],
      rowCount: 0,
      rows: [],
    };
  return {
    ...clone(run.source),
    id: run.id,
    startedAt,
    finishedAt: new Date(run.startedAt + SIMULATED_RUN_MS).toISOString(),
  };
}

function recordedRuns(state: DemoState, auditId: string): AuditRun[] {
  return (
    bodyOf<{ runs: AuditRun[] }>(state.responses, `GET ${API}/audits/${auditId}/runs`)?.runs ?? []
  );
}

function simulatedFor(state: DemoState, auditId: string): SimulatedRun[] {
  return state.runs
    .filter((run) => run.auditId === auditId)
    .sort((a, b) => b.startedAt - a.startedAt);
}

function listCatalog(state: DemoState): DemoResponse | null {
  const base = bodyOf<AuditCatalog>(state.responses, `GET ${API}/audits`);
  if (!base) return null;
  return json(200, {
    ...base,
    audits: base.audits.map((entry) => {
      const newest = simulatedFor(state, entry.definition.id)[0];
      if (!newest) return entry;
      const run = simulated(state, newest);
      return {
        ...entry,
        lastRun: runSummary(run),
        runningId: run.status === "running" ? run.id : null,
      };
    }),
  });
}

function listRuns(state: DemoState, auditId: string): DemoResponse | null {
  const base = bodyOf<{ runs: AuditRun[] }>(state.responses, `GET ${API}/audits/${auditId}/runs`);
  if (!base) return null;
  const started = simulatedFor(state, auditId).map((run) => runSummary(simulated(state, run)));
  return json(200, { runs: [...started, ...base.runs] });
}

function runDetail(state: DemoState, id: string): AuditRunDetail | null {
  const run = state.runs.find((candidate) => candidate.id === id);
  if (run) return simulated(state, run);
  return bodyOf<AuditRunDetail>(state.responses, `GET ${API}/audit-runs/${id}`);
}

function getRun(state: DemoState, id: string): DemoResponse {
  const detail = runDetail(state, id);
  return detail
    ? json(200, detail)
    : json(404, { error: "run_not_found", message: "That run was deleted." });
}

function startRun(state: DemoState, auditId: string): DemoResponse {
  const entry = bodyOf<AuditCatalog>(state.responses, `GET ${API}/audits`)?.audits.find(
    (candidate) => candidate.definition.id === auditId,
  );
  if (!entry) return json(404, { error: "unknown_audit", message: "There is no such audit." });
  // A run already in progress is joined, as the server does.
  const current = simulatedFor(state, auditId).find((run) => isRunning(state, run));
  if (current) return json(200, { run: runSummary(simulated(state, current)), started: false });
  const latest = recordedRuns(state, auditId).find((run) => run.status === "succeeded");
  const source = latest ? runDetail(state, latest.id) : null;
  if (!source) return json(409, { error: "run_unavailable", message: "Reload the page." });
  const run: SimulatedRun = {
    id: newId(),
    auditId,
    startedAt: state.now().getTime(),
    source,
  };
  state.runs.push(run);
  return json(202, { run: runSummary(simulated(state, run)), started: true });
}

// The CSV the server would send for the same query: what the page shows, built here from the
// rows and the audit's own definition.
function exportCsv(state: DemoState, id: string, url: URL): DemoResponse {
  const detail = runDetail(state, id);
  const definition = bodyOf<AuditCatalog>(state.responses, `GET ${API}/audits`)?.audits.find(
    (entry) => entry.definition.id === detail?.auditId,
  )?.definition;
  if (!detail || !definition || detail.status !== "succeeded")
    return json(404, { error: "run_not_found", message: "That run has no results." });
  const query = url.searchParams;
  const facets = Object.fromEntries(
    [...query.entries()]
      .filter(([key]) => key.startsWith("f."))
      .map(([key, value]) => [key.slice(2), value]),
  );
  const sort = query.get("sort");
  const rows = queryAuditRows(definition, detail.rows, {
    view: query.get("view"),
    facets,
    search: query.get("search") ?? "",
    sort: sort
      ? { column: sort, direction: query.get("direction") === "asc" ? "asc" : "desc" }
      : null,
  });
  return {
    status: 200,
    contentType: "text/csv; charset=utf-8",
    body: auditCsv(definition, rows),
    headers: {
      "Content-Disposition": `attachment; filename="prism365-${detail.auditId}-${detail.startedAt.slice(0, 10)}.csv"`,
    },
  };
}

// The saved schedule, saying collection is on or off as the page last left it.
function scheduleAnswer(state: DemoState): AuditScheduleSettings | null {
  if (!state.auditSchedule) return null;
  return {
    ...clone(state.auditSchedule),
    collecting: state.collection
      ? state.collection.state === "collecting"
      : state.auditSchedule.collecting,
  };
}

// ---------------------------------------------------------------------------------------------
// Requests
// ---------------------------------------------------------------------------------------------

function fromRecording(state: DemoState, request: DemoRequest): DemoResponse {
  const hit = matchFixture(state.responses, request.method, request.url);
  return hit ? json(hit.status, hit.body) : unavailable();
}

function handleGet(state: DemoState, path: string, url: URL, request: DemoRequest): DemoResponse {
  const segments = path.split("/").filter(Boolean);
  const [first, second, third] = segments;
  let answer: DemoResponse | null = null;
  if (path === "/attention") answer = listAttention(state, url);
  else if (path === "/findings") answer = listFindings(state, url);
  else if (path === "/modules") answer = listModules(state);
  else if (path === "/people") answer = listPeople(state, url);
  else if (first === "people" && second && second !== "lookup" && segments.length === 2)
    answer = personRecord(state, second);
  else if (first === "people" && second && third === "signins" && segments.length === 3) {
    // An exact recording wins; the other filters are worked out from the one for everything.
    const exact = state.responses[normalizeKey("GET", request.url)];
    answer = exact ? json(exact.status, exact.body) : personSignins(state, second, url);
  } else if (path === "/settings/attention" && state.settings)
    answer = json(200, clone(state.settings));
  else if (path === "/settings/audits" && state.auditSettings)
    answer = json(200, clone(state.auditSettings));
  else if (path === "/settings/audit-schedule" && state.auditSchedule)
    answer = json(200, scheduleAnswer(state));
  else if (path === "/audits") answer = listCatalog(state);
  else if (first === "audits" && second && third === "runs" && segments.length === 3)
    answer = listRuns(state, second);
  else if (first === "audit-runs" && second && segments.length === 2)
    answer = getRun(state, second);
  else if (first === "audit-runs" && second && third === "export.csv" && segments.length === 3)
    answer = exportCsv(state, second, url);
  return answer ?? fromRecording(state, request);
}

function handlePost(state: DemoState, path: string, body: unknown): DemoResponse {
  const segments = path.split("/").filter(Boolean);
  const [first, second, third] = segments;
  const payload = (body !== null && typeof body === "object" ? body : {}) as Json;

  if (first === "attention" && second === "read-all" && segments.length === 2) {
    let updated = 0;
    for (const item of state.attention)
      if (item.unread) {
        item.unread = false;
        updated += 1;
      }
    return json(200, { updated });
  }
  if (first === "attention" && second && third && segments.length === 3) {
    if (third === "review") {
      const next = reviewStates.find((value) => value === payload.state);
      if (!next)
        return json(400, { error: "invalid_request", message: "Choose new, reviewed or done." });
      const missing = changeAttention(state, second, (item, at) => {
        // Reviewing an item also reads it, and never changes what Microsoft said.
        item.review = { state: next, changedAt: at };
        if (next !== "new") item.unread = false;
      });
      return missing ?? json(200, clone(state.attention.find((item) => item.id === second)));
    }
    if (third === "mute") {
      if (typeof payload.muted !== "boolean")
        return json(400, { error: "invalid_request", message: "Choose muted or not." });
      const muted = payload.muted;
      const missing = changeAttention(state, second, (item) => {
        item.muted = muted;
        if (muted) item.unread = false;
      });
      return missing ?? json(200, clone(state.attention.find((item) => item.id === second)));
    }
    if (third === "read") {
      const item = state.attention.find((candidate) => candidate.id === second);
      const wasUnread = Boolean(item?.unread);
      if (item) item.unread = false;
      return json(200, { updated: wasUnread ? 1 : 0 });
    }
  }
  if (path === "/settings/attention" && state.settings) {
    const parsed = attentionPreferencesSchema.safeParse(payload);
    if (!parsed.success)
      return json(400, { error: "invalid_request", message: "Check the settings." });
    state.settings = { ...state.settings, preferences: parsed.data };
    return json(200, clone(state.settings));
  }
  if (path === "/settings/audits" && state.auditSettings) {
    // Saved for this page; the recorded results are not re-run with it.
    const parsed = auditScopeSchema.safeParse(payload);
    if (!parsed.success)
      return json(400, { error: "invalid_request", message: "Check the domains." });
    state.auditSettings = { ...state.auditSettings, scope: parsed.data };
    return json(200, clone(state.auditSettings));
  }
  if (path === "/settings/audit-schedule" && state.auditSchedule) {
    // Saved for this page; nothing runs on its own in the demo.
    const parsed = auditScheduleSchema.safeParse(payload);
    if (!parsed.success)
      return json(400, {
        error: "invalid_request",
        message: "Choose off, daily or weekly for each audit.",
      });
    // Only the audits named change, and audits the recording does not know are dropped, as the
    // server does. The demo's administrator may run every audit, so none is refused.
    const saved = parsed.data.audits;
    const current = state.auditSchedule.audits;
    state.auditSchedule = {
      ...state.auditSchedule,
      audits: Object.fromEntries(
        Object.keys(current).map((id) => [
          id,
          Object.hasOwn(saved, id) ? (saved[id] ?? "off") : (current[id] ?? "off"),
        ]),
      ),
    };
    return json(200, scheduleAnswer(state));
  }
  if (first === "findings" && second && third === "state" && segments.length === 3) {
    const next = findingStateSchema.safeParse(payload.state);
    if (!next.success)
      return json(400, {
        error: "invalid_request",
        message: "Choose open, acknowledged, or resolved.",
      });
    const finding = state.findings.find((item) => item.id === second);
    if (!finding)
      return json(404, { error: "finding_not_found", message: "The finding is not available." });
    finding.state = next.data as FindingState;
    finding.stateChangedAt = state.now().toISOString();
    return json(200, clone(finding));
  }
  if (path === "/collection/start") return setCollection(state, true);
  if (path === "/collection/stop") return setCollection(state, false);
  if (first === "audits" && second && third === "runs" && segments.length === 3)
    return startRun(state, second);
  if (path === "/auth/logout") return json(200, { signedOut: true });
  return unavailable();
}

export function handleRequest(state: DemoState, request: DemoRequest): DemoResponse {
  const url = new URL(request.url, BASE);
  if (url.pathname !== API && !url.pathname.startsWith(`${API}/`)) return unavailable();
  const path = url.pathname.slice(API.length);
  const method = request.method.toUpperCase();
  if (method === "GET") return handleGet(state, path, url, request);
  if (method === "POST") return handlePost(state, path, request.body);
  return unavailable();
}

// ---------------------------------------------------------------------------------------------
// Installing the shim
// ---------------------------------------------------------------------------------------------

export interface ReplayTarget {
  fetch: typeof fetch;
  location?: { origin: string };
  document?: Document;
}

export interface Replay {
  state: DemoState;
  // Puts the page's own fetch back.
  restore: () => void;
}

function toResponse(answer: DemoResponse): Response {
  const text = typeof answer.body === "string" ? answer.body : JSON.stringify(answer.body);
  return new Response(text, {
    status: answer.status,
    headers: { "Content-Type": answer.contentType, ...answer.headers },
  });
}

function readBody(init: RequestInit | undefined): unknown {
  if (typeof init?.body !== "string") return undefined;
  try {
    return JSON.parse(init.body);
  } catch {
    return undefined;
  }
}

// The export is a plain link with the `download` attribute, which no fetch replacement can
// reach, so the click is answered here: the CSV is built from the recording and saved.
function saveCsv(answer: DemoResponse, document: Document) {
  if (answer.status !== 200 || typeof answer.body !== "string") {
    const message = (answer.body as { message?: string }).message ?? unavailableBody.message;
    globalThis.alert?.(message);
    return;
  }
  const name = /filename="([^"]+)"/.exec(answer.headers?.["Content-Disposition"] ?? "")?.[1];
  const link = document.createElement("a");
  link.href = URL.createObjectURL(new Blob([answer.body], { type: answer.contentType }));
  link.download = name ?? "prism365-export.csv";
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(link.href), 1_000);
}

export function installReplay(
  fixtures: Fixtures,
  options: { target?: ReplayTarget; now?: () => Date } = {},
): Replay {
  const target = options.target ?? (globalThis as unknown as ReplayTarget);
  const state = createDemoState(fixtures, options.now);
  const real = target.fetch;
  const origin = target.location?.origin ?? BASE;

  target.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    const isRequest = typeof Request !== "undefined" && input instanceof Request;
    const url = new URL(isRequest ? input.url : String(input), origin);
    if (url.origin !== new URL(origin).origin || !url.pathname.startsWith(`${API}/`))
      return real.call(target, input, init);
    if (init?.signal?.aborted) throw new DOMException("The request was aborted.", "AbortError");
    const method = init?.method ?? (isRequest ? input.method : "GET");
    return toResponse(
      handleRequest(state, { method, url: `${url.pathname}${url.search}`, body: readBody(init) }),
    );
  }) as typeof fetch;

  const document = target.document;
  const onClick = (event: MouseEvent) => {
    const link = (event.target as Element | null)?.closest?.("a[href]");
    const href = link?.getAttribute("href") ?? "";
    if (!href.startsWith(`${API}/audit-runs/`) || !href.includes("/export.csv")) return;
    event.preventDefault();
    saveCsv(handleRequest(state, { method: "GET", url: href }), document as Document);
  };
  document?.addEventListener("click", onClick, true);

  return {
    state,
    restore: () => {
      target.fetch = real;
      document?.removeEventListener("click", onClick, true);
    },
  };
}

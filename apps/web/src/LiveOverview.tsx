import {
  type AttentionItem,
  type AttentionList,
  type AttentionSource,
  type AuditCatalog,
  auditCatalogSchema,
  type Finding,
  findingListSchema,
  type ModuleList,
  moduleListSchema,
  type PrismSession,
  type Severity,
  type SigninSummary,
  signinSummarySchema,
} from "@prism365/contracts";
import { type ReactNode, useCallback, useEffect, useState } from "react";
import type { WorkspaceView } from "./App.tsx";
import { useAttention } from "./Attention.tsx";
import { auditResultCounts } from "./Audits.tsx";
import { day, number, RelativeTime, readJson, SeverityDot, SkeletonRows } from "./common.tsx";

// The live home page: one glance at whether anything is wrong, then the open findings and what
// Microsoft is reporting, each opening the page that owns it. The tiles and lists are worked out
// from the data by the small pure functions below, so each rule has a test of its own.

const SIGNINS = "signins.audit";
const CHANGES = "directory.changes";
// How many open findings each module's request asks for, and so the most a count can be sure of.
const FINDING_LIMIT = 100;
// Rows in each of the two lists.
const ROW_LIMIT = 8;
// The rules whose findings say that part of what a collector reads could not be read.
const GAP_RULES: ReadonlySet<string> = new Set(["signins.collection_gap", "collection.gap"]);

// One request's answer. Both empty means it has not come back; a value with an error is the last
// good answer after a refresh failed; an error alone is a request that never worked.
export interface Source<T> {
  value: T | null;
  error: string | null;
}

export interface OverviewData {
  modules: Source<ModuleList>;
  summary: Source<SigninSummary>;
  signins: Source<Finding[]>;
  changes: Source<Finding[]>;
  audits: Source<AuditCatalog>;
  attention: Source<AttentionList>;
}

const pending = <T,>(): Source<T> => ({ value: null, error: null });

// Nothing has come back yet: every tile and list draws its loading rows.
export const emptyOverviewData: OverviewData = {
  modules: pending(),
  summary: pending(),
  signins: pending(),
  changes: pending(),
  audits: pending(),
  attention: pending(),
};

const isPending = (source: Source<unknown>) => source.value === null && source.error === null;
const isFailed = (source: Source<unknown>) => source.value === null && source.error !== null;

// The answer to keep after a request finished: a new value, or the old one with the reason the
// refresh failed, so one slow or broken request never blanks what already showed.
export function settle<T>(previous: Source<T>, result: PromiseSettledResult<T>): Source<T> {
  if (result.status === "fulfilled") return { value: result.value, error: null };
  const reason = result.reason;
  return {
    value: previous.value,
    error: reason instanceof Error && reason.message ? reason.message : "The request failed.",
  };
}

type SourceStatus = "loading" | "unavailable" | "ready";

// Unavailable as soon as one of the sources is, because a figure built from the rest would read as
// complete when it is not.
function statusOf(...sources: Source<unknown>[]): SourceStatus {
  if (sources.some(isFailed)) return "unavailable";
  if (sources.some(isPending)) return "loading";
  return "ready";
}

// ---------------------------------------------------------------------------------------------
// Open findings
// ---------------------------------------------------------------------------------------------

const severityRank: Record<Severity, number> = {
  critical: 4,
  high: 3,
  medium: 2,
  low: 1,
  info: 0,
};

export interface FindingCounts {
  total: number;
  critical: number;
  high: number;
  medium: number;
  // Low and info together: part of the total only.
  low: number;
  // A module returned as many as were asked for, so there may be more than counted.
  capped: boolean;
}

// The open findings of both modules, counted by severity. Only open ones count, whatever the
// lists hold.
export function findingCounts(
  lists: readonly (readonly Finding[])[],
  limit = FINDING_LIMIT,
): FindingCounts {
  const counts: FindingCounts = {
    total: 0,
    critical: 0,
    high: 0,
    medium: 0,
    low: 0,
    capped: false,
  };
  for (const list of lists) {
    if (list.length >= limit) counts.capped = true;
    for (const finding of list) {
      if (finding.state !== "open") continue;
      counts.total += 1;
      if (finding.severity === "critical") counts.critical += 1;
      else if (finding.severity === "high") counts.high += 1;
      else if (finding.severity === "medium") counts.medium += 1;
      else counts.low += 1;
    }
  }
  return counts;
}

// Open findings from both modules, the most severe first and the newest first among equals, cut
// to the rows the Overview has room for.
export function needsAttention(
  lists: readonly (readonly Finding[])[],
  limit = ROW_LIMIT,
): Finding[] {
  return lists
    .flat()
    .filter((finding) => finding.state === "open")
    .sort(
      (a, b) =>
        severityRank[b.severity] - severityRank[a.severity] ||
        Date.parse(b.lastSeen) - Date.parse(a.lastSeen) ||
        a.id.localeCompare(b.id),
    )
    .slice(0, limit);
}

function findingView(finding: Finding): "signins" | "changes" {
  return finding.module === CHANGES ? "changes" : "signins";
}

// Findings live on two pages, one per module, and there is no page for both. "View all" and the
// findings tile open the one with more of them, the sign-in audit when they are level.
export function findingsHome(
  signins: readonly Finding[],
  changes: readonly Finding[],
): "signins" | "changes" {
  const open = (list: readonly Finding[]) => list.filter((item) => item.state === "open").length;
  return open(changes) > open(signins) ? "changes" : "signins";
}

// ---------------------------------------------------------------------------------------------
// Collection
// ---------------------------------------------------------------------------------------------

export type CollectionName = "collecting" | "stopped" | "gap";

export interface CollectionSummary {
  state: CollectionName;
  // The start of the newest run of any collector; null when none has run.
  lastRun: string | null;
}

// Stopped when nobody is collecting; a gap when it is running but an open finding says a span of
// what a collector reads could not be read; otherwise collecting.
export function collectionSummary(
  modules: ModuleList,
  open: readonly Finding[],
): CollectionSummary {
  let newest: string | null = null;
  for (const module of modules.modules)
    for (const collector of module.collectors) {
      const started = collector.lastRun?.startedAt ?? null;
      if (started && (newest === null || Date.parse(started) > Date.parse(newest)))
        newest = started;
    }
  if (modules.collection.state === "stopped") return { state: "stopped", lastRun: newest };
  const gap = open.some((finding) => finding.state === "open" && GAP_RULES.has(finding.ruleId));
  return { state: gap ? "gap" : "collecting", lastRun: newest };
}

// ---------------------------------------------------------------------------------------------
// Microsoft
// ---------------------------------------------------------------------------------------------

const WATCHED_SOURCES = ["service.health", "service.messages"];

export interface MicrosoftSummary {
  incidents: AttentionItem[];
  actions: AttentionItem[];
  // At least one of the two sources is being read. Without it, no items says nothing.
  checked: boolean;
  // Each source on its own: a count of zero from a source that is not read is not a zero.
  incidentsChecked: boolean;
  postsChecked: boolean;
  // Plain words for what is not being read, when something is not.
  problem: string | null;
}

const sourceProblems: Record<AttentionSource["state"], string> = {
  available: "is available",
  not_enabled: "is not set up",
  missing_permission: "needs more access",
  not_collecting: "is not collecting",
};

function deadlineOf(item: AttentionItem): number {
  return item.provider.deadline ? Date.parse(item.provider.deadline) : Number.POSITIVE_INFINITY;
}

// What an empty list means, only for the sources that were read.
export function emptyMicrosoft(summary: MicrosoftSummary): string | null {
  if (summary.incidentsChecked && summary.postsChecked)
    return "Microsoft reports no incidents, and no Message center post needs action.";
  if (summary.incidentsChecked) return "Microsoft reports no incidents.";
  if (summary.postsChecked) return "No Message center post needs action.";
  return null;
}

// Open incidents and the Message center posts that ask for action, from the notification list. A
// source the list does not mention is taken to be working: the items speak for themselves.
export function microsoftSummary(list: AttentionList): MicrosoftSummary {
  const watched = WATCHED_SOURCES.map((id) => list.sources.find((source) => source.module === id));
  const broken = watched.filter(
    (source): source is AttentionSource => source !== undefined && source.state !== "available",
  );
  const incidents = list.items
    .filter((item) => item.class === "incident" && item.provider.active)
    .sort((a, b) => Date.parse(b.provider.updatedAt) - Date.parse(a.provider.updatedAt));
  const actions = list.items
    .filter((item) => item.class === "message.action" && item.provider.active)
    .sort((a, b) => {
      const left = deadlineOf(a);
      const right = deadlineOf(b);
      return left === right ? a.title.localeCompare(b.title) : left < right ? -1 : 1;
    });
  return {
    incidents,
    actions,
    checked: broken.length < watched.length,
    incidentsChecked: !broken.some((source) => source.module === "service.health"),
    postsChecked: !broken.some((source) => source.module === "service.messages"),
    problem:
      broken.length === 0
        ? null
        : `${broken.map((source) => `${source.title} ${sourceProblems[source.state]}`).join(". ")}.`,
  };
}

// ---------------------------------------------------------------------------------------------
// Tiles
// ---------------------------------------------------------------------------------------------

export type TileId = "findings" | "audits" | "signins" | "collection" | "microsoft";

export interface Tile {
  id: TileId;
  label: string;
  // The page the tile opens.
  view: WorkspaceView;
  status: SourceStatus;
  value: string;
  // A coloured dot before the value, for the one tile whose value is a state and not a number.
  dot: "good" | "warning" | null;
  detail: ReactNode;
}

const plural = (count: number, one: string, many: string) =>
  `${number.format(count)} ${count === 1 ? one : many}`;

const collectionLabels: Record<CollectionName, string> = {
  collecting: "Collecting",
  stopped: "Stopped",
  gap: "Gap",
};

function tile(id: TileId, label: string, view: WorkspaceView, status: SourceStatus): Tile {
  return { id, label, view, status, value: "", dot: null, detail: null };
}

export function buildTiles(data: OverviewData, now: number): Tile[] {
  const findings = tile(
    "findings",
    "Open findings",
    findingsHome(data.signins.value ?? [], data.changes.value ?? []),
    statusOf(data.signins, data.changes),
  );
  if (findings.status === "ready") {
    const counts = findingCounts([data.signins.value ?? [], data.changes.value ?? []]);
    findings.value = `${number.format(counts.total)}${counts.capped ? "+" : ""}`;
    const shown = (["critical", "high", "medium"] as const).filter((level) => counts[level] > 0);
    findings.detail =
      counts.total === 0 ? (
        "Nothing open"
      ) : shown.length === 0 ? (
        "Low severity only"
      ) : (
        <span className="overview-tile__counts">
          {shown.map((level) => (
            <SeverityDot key={level} severity={level} count={counts[level]} />
          ))}
        </span>
      );
  }

  const audits = tile("audits", "Critical audit results", "audits", statusOf(data.audits));
  if (audits.status === "ready" && data.audits.value) {
    const { critical, warning, ran } = auditResultCounts(data.audits.value);
    audits.value = number.format(critical);
    audits.detail =
      ran === 0
        ? "No audit has run yet"
        : `${warning > 0 ? `${number.format(warning)} more with warnings, ` : ""}of ${number.format(ran)} run`;
  }

  const signins = tile("signins", "Sign-ins, last 7 days", "signins", statusOf(data.summary));
  if (signins.status === "ready" && data.summary.value) {
    const { totals, coverage } = data.summary.value;
    signins.value = number.format(totals.successes);
    signins.detail =
      coverage.events === 0 ? (
        "Nothing collected yet"
      ) : (
        <span className={totals.failures > 0 ? "overview-tile__alert" : undefined}>
          {number.format(totals.failures)} failed
        </span>
      );
  }

  // A gap shows in the findings, so the tile waits for them as well as for the modules. If they
  // never load it reports what it can: collecting or stopped, without a gap.
  const collection = tile(
    "collection",
    "Collection",
    "signins",
    isFailed(data.modules)
      ? "unavailable"
      : [data.modules, data.signins, data.changes].some(isPending)
        ? "loading"
        : "ready",
  );
  if (collection.status === "ready" && data.modules.value) {
    const summary = collectionSummary(data.modules.value, [
      ...(data.signins.value ?? []),
      ...(data.changes.value ?? []),
    ]);
    collection.value = collectionLabels[summary.state];
    collection.dot = summary.state === "collecting" ? "good" : "warning";
    collection.detail = summary.lastRun ? (
      <>
        Last run <RelativeTime value={summary.lastRun} now={now} />
      </>
    ) : (
      "No run yet"
    );
  }

  const microsoft = tile("microsoft", "Microsoft", "notifications", statusOf(data.attention));
  if (microsoft.status === "ready" && data.attention.value) {
    const summary = microsoftSummary(data.attention.value);
    if (!summary.checked) {
      microsoft.value = "Not checked";
      microsoft.detail = summary.problem;
    } else {
      microsoft.value = number.format(summary.incidents.length + summary.actions.length);
      const incidents = summary.incidentsChecked
        ? plural(summary.incidents.length, "incident", "incidents")
        : "incidents not checked";
      const posts = summary.postsChecked
        ? plural(summary.actions.length, "post needs action", "posts need action")
        : "Message center not checked";
      microsoft.detail = `${incidents}, ${posts}`;
    }
  }

  return [findings, audits, signins, collection, microsoft];
}

// ---------------------------------------------------------------------------------------------
// The page
// ---------------------------------------------------------------------------------------------

const sourceNames: Record<keyof OverviewData, string> = {
  modules: "collection",
  summary: "sign-ins",
  signins: "sign-in findings",
  changes: "directory change findings",
  audits: "audits",
  attention: "Microsoft updates",
};

// A source that failed with nothing to show says so where it would have appeared. Only one that
// failed to refresh needs a word here, because its old figures are still on the page.
function Banner({ data }: { data: OverviewData }) {
  const stale = (Object.keys(sourceNames) as (keyof OverviewData)[]).filter(
    (key) => data[key].value !== null && data[key].error !== null,
  );
  if (stale.length === 0) return null;
  return (
    <p className="form-error form-error--banner" role="alert">
      Showing the last figures that loaded. Could not refresh{" "}
      {stale.map((key) => sourceNames[key]).join(", ")}.
    </p>
  );
}

function TileButton({ item, onOpen }: { item: Tile; onOpen: () => void }) {
  return (
    <button type="button" className="overview-tile" onClick={onOpen}>
      <span className="overview-tile__label">{item.label}</span>
      {item.status === "loading" ? (
        <SkeletonRows rows={2} label={`Loading ${item.label.toLowerCase()}`} />
      ) : item.status === "unavailable" ? (
        <>
          <span className="overview-tile__value overview-tile__value--muted">Unavailable</span>
          <span className="overview-tile__detail">Could not be loaded</span>
        </>
      ) : (
        <>
          <span className="overview-tile__value">
            {item.dot ? (
              <span className={`tile-dot tile-dot--${item.dot}`} aria-hidden="true" />
            ) : null}
            {item.value}
          </span>
          <span className="overview-tile__detail">{item.detail}</span>
        </>
      )}
    </button>
  );
}

function CardHeader({ id, title, action }: { id: string; title: string; action: ReactNode }) {
  return (
    <div className="overview-card__header">
      <h2 id={id}>{title}</h2>
      {action}
    </div>
  );
}

function NeedsAttention({
  data,
  now,
  onNavigate,
  onOpenPerson,
}: {
  data: OverviewData;
  now: number;
  onNavigate: (view: WorkspaceView) => void;
  onOpenPerson: (userId: string) => void;
}) {
  const status = statusOf(data.signins, data.changes);
  const rows = needsAttention([data.signins.value ?? [], data.changes.value ?? []]);
  const missing = [
    isFailed(data.signins) ? "Sign-in findings" : null,
    isFailed(data.changes) ? "Directory change findings" : null,
  ].filter(Boolean);
  return (
    <section className="panel overview-card" aria-labelledby="overview-attention">
      <CardHeader
        id="overview-attention"
        title="Needs attention"
        action={
          rows.length > 0 ? (
            <button
              className="link-button"
              type="button"
              onClick={() =>
                onNavigate(findingsHome(data.signins.value ?? [], data.changes.value ?? []))
              }
            >
              View all
            </button>
          ) : null
        }
      />
      {rows.length > 0 ? (
        <ul className="overview-rows">
          {rows.map((finding) => (
            <li key={finding.id}>
              <button
                type="button"
                className="overview-row"
                onClick={() =>
                  finding.subject.kind === "person"
                    ? onOpenPerson(finding.subject.id)
                    : onNavigate(findingView(finding))
                }
              >
                <span className="overview-row__head" title={finding.title}>
                  <SeverityDot severity={finding.severity} />{" "}
                  <span className="overview-row__title">{finding.title}</span>
                </span>
                <span className="overview-row__meta">
                  <span className="overview-row__subject">
                    {finding.subject.label ?? finding.subject.id}
                  </span>
                  <span className="overview-row__when">
                    <RelativeTime value={finding.lastSeen} now={now} />
                  </span>
                </span>
              </button>
            </li>
          ))}
        </ul>
      ) : status === "loading" ? (
        <SkeletonRows rows={5} label="Loading findings" />
      ) : status === "unavailable" ? (
        <p className="empty-data">Findings could not be loaded.</p>
      ) : (
        <p className="empty-data">No open findings. New ones appear here as collection runs.</p>
      )}
      {rows.length > 0 && missing.length > 0 ? (
        <p className="overview-card__note">{missing.join(" and ")} could not be loaded.</p>
      ) : null}
    </section>
  );
}

function dueText(item: AttentionItem): string {
  if (item.class === "incident") return item.provider.status ?? "Incident";
  return item.provider.deadline ? `Due ${day(item.provider.deadline)}` : "No deadline";
}

function FromMicrosoft({
  data,
  now,
  onNavigate,
}: {
  data: OverviewData;
  now: number;
  onNavigate: (view: WorkspaceView) => void;
}) {
  const status = statusOf(data.attention);
  const summary = data.attention.value ? microsoftSummary(data.attention.value) : null;
  const rows = summary ? [...summary.incidents, ...summary.actions].slice(0, ROW_LIMIT) : [];
  return (
    <section className="panel overview-card" aria-labelledby="overview-microsoft">
      <CardHeader
        id="overview-microsoft"
        title="From Microsoft"
        action={
          rows.length > 0 ? (
            <button
              className="link-button"
              type="button"
              onClick={() => onNavigate("notifications")}
            >
              View all
            </button>
          ) : null
        }
      />
      {status === "loading" ? (
        <SkeletonRows rows={4} label="Loading Microsoft updates" />
      ) : status === "unavailable" || !summary ? (
        <p className="empty-data">Microsoft updates could not be loaded.</p>
      ) : (
        <>
          {rows.length > 0 ? (
            <ul className="overview-rows">
              {rows.map((item) => {
                const late =
                  item.class !== "incident" &&
                  item.provider.deadline !== null &&
                  Date.parse(item.provider.deadline) < now;
                return (
                  <li key={item.id}>
                    <button
                      type="button"
                      className="overview-row overview-row--microsoft"
                      onClick={() => onNavigate("notifications")}
                    >
                      <span className="overview-row__head" title={item.title}>
                        <span className="overview-row__title">{item.title}</span>
                      </span>
                      <span className="overview-row__meta">
                        <span className="overview-row__subject">
                          {item.services.length > 0 ? item.services.join(", ") : "Microsoft 365"}
                        </span>
                        <span
                          className={`overview-row__due${late ? " overview-row__due--late" : ""}`}
                        >
                          {dueText(item)}
                        </span>
                      </span>
                    </button>
                  </li>
                );
              })}
            </ul>
          ) : emptyMicrosoft(summary) ? (
            <p className="empty-data">{emptyMicrosoft(summary)}</p>
          ) : null}
          {summary.problem ? (
            <p className="overview-card__note">
              {summary.problem}{" "}
              <button className="link-button" type="button" onClick={() => onNavigate("access")}>
                Microsoft access
              </button>
            </p>
          ) : null}
        </>
      )}
    </section>
  );
}

// The page without any fetching, so every state of it can be drawn from plain data.
export function OverviewView({
  data,
  now,
  onNavigate,
  onOpenPerson,
}: {
  data: OverviewData;
  now: number;
  onNavigate: (view: WorkspaceView) => void;
  onOpenPerson: (userId: string) => void;
}) {
  return (
    <div className="overview">
      <Banner data={data} />
      <ul className="overview-tiles" aria-label="Summary">
        {buildTiles(data, now).map((item) => (
          <li key={item.id}>
            <TileButton item={item} onOpen={() => onNavigate(item.view)} />
          </li>
        ))}
      </ul>
      <div className="overview-columns">
        <NeedsAttention data={data} now={now} onNavigate={onNavigate} onOpenPerson={onOpenPerson} />
        <FromMicrosoft data={data} now={now} onNavigate={onNavigate} />
      </div>
    </div>
  );
}

function findingsUrl(module: string): string {
  return `/api/v1/findings?limit=${FINDING_LIMIT}&module=${module}&state=open`;
}

// Every request the page makes besides the notification list, which has its own hook. The public
// demo answers only requests it recorded, so these are written in exactly the recorded forms.
export const OVERVIEW_REQUESTS = {
  modules: "/api/v1/modules",
  summary: `/api/v1/modules/${SIGNINS}/summary?days=7`,
  signins: findingsUrl(SIGNINS),
  changes: findingsUrl(CHANGES),
  audits: "/api/v1/audits",
} as const;

export function LiveOverview({
  session,
  onOpenPerson,
  onNavigate,
}: {
  session: PrismSession;
  onOpenPerson: (userId: string) => void;
  onNavigate: (view: WorkspaceView) => void;
}) {
  const [data, setData] = useState<Omit<OverviewData, "attention">>(emptyOverviewData);
  const { list, error } = useAttention(session, "open");

  const load = useCallback(
    async (signal?: AbortSignal) => {
      if (!session.authenticated) return;
      const init = signal ? { signal } : {};
      // Each request stands alone: one that fails leaves the others, and the page says which.
      const [modules, summary, signins, changes, audits] = await Promise.allSettled([
        readJson(OVERVIEW_REQUESTS.modules, (value) => moduleListSchema.parse(value), init),
        readJson(OVERVIEW_REQUESTS.summary, (value) => signinSummarySchema.parse(value), init),
        readJson(OVERVIEW_REQUESTS.signins, (value) => findingListSchema.parse(value).items, init),
        readJson(OVERVIEW_REQUESTS.changes, (value) => findingListSchema.parse(value).items, init),
        readJson(OVERVIEW_REQUESTS.audits, (value) => auditCatalogSchema.parse(value), init),
      ]);
      if (signal?.aborted) return;
      setData((previous) => ({
        modules: settle(previous.modules, modules),
        summary: settle(previous.summary, summary),
        signins: settle(previous.signins, signins),
        changes: settle(previous.changes, changes),
        audits: settle(previous.audits, audits),
      }));
    },
    [session.authenticated],
  );

  useEffect(() => {
    const controller = new AbortController();
    void load(controller.signal);
    const timer = setInterval(() => void load(), 60_000);
    return () => {
      controller.abort();
      clearInterval(timer);
    };
  }, [load]);

  return (
    <OverviewView
      data={{ ...data, attention: { value: list, error: error || null } }}
      now={Date.now()}
      onNavigate={onNavigate}
      onOpenPerson={onOpenPerson}
    />
  );
}

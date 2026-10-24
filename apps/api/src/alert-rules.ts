import { type StoredTrendPoint, trendDifferences } from "@prism365/audits";
import {
  type AlertRule,
  cleanText,
  type Digest,
  directoryAlertRuleIds,
  type Severity,
} from "@prism365/contracts";
import { type FindingNotification, type StopReason, severityAtLeast } from "@prism365/modules";
import type { AlertItem, AlertMessage } from "./delivery-payload.ts";

// What each kind of alert rule watches (docs/alerts.md). Pure functions: the delivery service
// (delivery.ts) feeds them findings from the scheduler's sink, trend points after an audit run is
// stored, and the collection state, and sends what they return. A rule never reads Microsoft.

export const DIRECTORY_MODULE = "directory.changes";
// Findings that say a module could not read a span of its source.
export const GAP_RULE_IDS: readonly string[] = ["collection.gap", "signins.collection_gap"];
export const COLLECTION_STOPPED_SUBJECT = "collection.stopped";

const severityRank: Record<Severity, number> = {
  info: 0,
  low: 1,
  medium: 2,
  high: 3,
  critical: 4,
};
const severityWords: Record<Severity, string> = {
  info: "info",
  low: "low",
  medium: "medium",
  high: "high",
  critical: "critical",
};

// Whether a new or escalated finding is one this rule alerts on.
export function findingMatches(rule: AlertRule, notification: FindingNotification): boolean {
  if (!rule.enabled) return false;
  if (rule.kind === "finding")
    return (
      severityAtLeast(notification.severity, rule.minimumSeverity) &&
      (rule.modules.length === 0 || rule.modules.includes(notification.module))
    );
  if (rule.kind === "directory")
    return (
      notification.module === DIRECTORY_MODULE &&
      (rule.ruleIds as readonly string[]).includes(notification.ruleId)
    );
  if (rule.kind === "collection") return rule.gaps && GAP_RULE_IDS.includes(notification.ruleId);
  return false;
}

// The page a finding is looked at on.
export function findingLink(module: string): string {
  if (module === "signins.audit") return "/#/signins";
  if (module === DIRECTORY_MODULE) return "/#/changes";
  return "/#/overview";
}

const changeNotes: Record<FindingNotification["change"], string> = {
  created: "new",
  reopened: "reopened",
  escalated: "escalated",
};

// One alert for the findings a rule matched in one batch, worst first. `details` holds the
// evidence fields of each finding by id, read only when a destination asked for them.
export function findingAlert(
  rule: AlertRule,
  notifications: readonly FindingNotification[],
  details: ReadonlyMap<string, Record<string, string | number | boolean | null>> = new Map(),
): AlertMessage {
  const sorted = [...notifications].sort(
    (a, b) =>
      severityRank[b.severity] - severityRank[a.severity] ||
      b.observedAt.localeCompare(a.observedAt),
  );
  const worst = sorted[0]?.severity ?? null;
  const count = sorted.length;
  const kind =
    rule.kind === "directory" ? "directory" : rule.kind === "collection" ? "collection" : "finding";
  const title =
    kind === "directory"
      ? count === 1
        ? "A directory change needs a look"
        : `${count} directory changes need a look`
      : kind === "collection"
        ? count === 1
          ? "A module could not read part of its source"
          : `${count} modules could not read part of their source`
        : count === 1
          ? `A new or escalated finding at ${severityWords[worst ?? "info"]} severity`
          : `${count} new or escalated findings, the worst at ${severityWords[worst ?? "info"]} severity`;
  const modules = new Set(sorted.map((notification) => notification.module));
  const items: AlertItem[] = sorted.slice(0, 10).map((notification) => ({
    title: notification.title,
    severity: notification.severity,
    subject: notification.subjectLabel,
    note: changeNotes[notification.change],
    details: {
      summary: notification.summary,
      ...(details.get(notification.findingId) ?? {}),
    },
  }));
  return {
    kind,
    title,
    severity: worst,
    count,
    items,
    lines: [`Rule: ${cleanText(rule.name)}`],
    link: modules.size === 1 ? findingLink([...modules][0] ?? "") : "/#/overview",
  };
}

// The newest point compared with the newest earlier point it can be compared with: same audit
// version, same scope, both complete (trendDifferences, as the trend chart decides). null when
// the tile is missing from the newest point or the newest point is partial.
export function auditMovement(
  points: readonly StoredTrendPoint[],
  tileId: string,
): { value: number; previous: number | null } | null {
  const newest = points.at(-1);
  if (!newest?.complete) return null;
  const tile = newest.tiles.find((candidate) => candidate.id === tileId);
  if (!tile) return null;
  for (let index = points.length - 2; index >= 0; index -= 1) {
    const point = points[index];
    if (!point || trendDifferences(point, newest).length > 0) continue;
    const before = point.tiles.find((candidate) => candidate.id === tileId);
    return { value: tile.value, previous: before ? before.value : null };
  }
  return { value: tile.value, previous: null };
}

// Whether an audit rule fires for this movement. above and below fire when the value crosses the
// threshold: past it now, and not past it at the previous comparable point (or there is none).
// change fires when the value moved by at least the threshold against that point.
export function auditFires(
  rule: Extract<AlertRule, { kind: "audit" }>,
  movement: { value: number; previous: number | null },
): boolean {
  if (!rule.enabled) return false;
  const { value, previous } = movement;
  if (rule.condition === "above")
    return value > rule.value && (previous === null || previous <= rule.value);
  if (rule.condition === "below")
    return value < rule.value && (previous === null || previous >= rule.value);
  return previous !== null && Math.abs(value - previous) >= rule.value;
}

export function auditAlert(
  rule: Extract<AlertRule, { kind: "audit" }>,
  audit: { title: string; tileLabel: string },
  movement: { value: number; previous: number | null },
): AlertMessage {
  const what = `${cleanText(audit.title)}: ${cleanText(audit.tileLabel)}`;
  const condition =
    rule.condition === "above"
      ? `It is now above ${rule.value}.`
      : rule.condition === "below"
        ? `It is now below ${rule.value}.`
        : `It moved by ${Math.abs(movement.value - (movement.previous ?? movement.value))}, at least the ${rule.value} this rule watches for.`;
  return {
    kind: "audit",
    title: `${what} is ${movement.value}`,
    severity: null,
    count: 1,
    items: [],
    lines: [
      condition,
      movement.previous === null
        ? "There is no earlier comparable run to compare with."
        : `It was ${movement.previous} at the previous comparable run.`,
      `Rule: ${cleanText(rule.name)}`,
    ],
    link: `/#/audits/${encodeURIComponent(rule.auditId)}`,
  };
}

export interface CollectionState {
  // An administrator chose to collect (the standing intent), so stopped collection is not on
  // purpose.
  intended: boolean;
  collecting: boolean;
  stoppedSince: Date | null;
  reason: StopReason | null;
}

export function collectionStoppedTooLong(
  rule: AlertRule,
  state: CollectionState,
  now: Date,
): boolean {
  if (!rule.enabled || rule.kind !== "collection" || rule.stoppedHours === null) return false;
  if (!state.intended || state.collecting || !state.stoppedSince) return false;
  return now.getTime() - state.stoppedSince.getTime() >= rule.stoppedHours * 3_600_000;
}

const reasons: Record<StopReason, string> = {
  operator_stopped: "An administrator stopped it.",
  signed_out: "The administrator it collects for signed out.",
  credential_unavailable:
    "Microsoft refused the sign-in it collects with. Signing in again starts it.",
  replaced: "Another administrator's sign-in took over and then ended.",
  shutdown: "Prism365 restarted and nobody has signed in since.",
};

export function collectionAlert(
  rule: Extract<AlertRule, { kind: "collection" }>,
  state: CollectionState,
  now: Date,
): AlertMessage {
  const hours = state.stoppedSince
    ? Math.floor((now.getTime() - state.stoppedSince.getTime()) / 3_600_000)
    : (rule.stoppedHours ?? 0);
  return {
    kind: "collection",
    title: "Background collection has stopped",
    severity: "high",
    count: 1,
    items: [],
    lines: [
      `Collection is meant to be on but has been stopped for ${hours} hour${hours === 1 ? "" : "s"}, so nothing new is being read.`,
      state.reason ? reasons[state.reason] : reasons.shutdown,
      `Rule: ${cleanText(rule.name)}`,
    ],
    link: "/#/overview",
  };
}

// The directory change ids a rule may name, for the page.
export const directoryRuleChoices: readonly string[] = directoryAlertRuleIds;

// Scheduled exports -------------------------------------------------------------------------------

export function exportAlert(input: {
  auditId: string;
  title: string;
  rows: number;
  complete: boolean;
  scoped: boolean;
  attached: boolean;
  // Why the file is not attached, when it is not.
  withheld: "webhook" | "details_off" | "too_large" | null;
}): AlertMessage {
  const withheld = {
    webhook: "Open Prism365 for the rows and the CSV file.",
    details_off:
      "The CSV is not attached because this destination does not include details. Open Prism365 for the file.",
    too_large: "The CSV is larger than 5 MB, so it is not attached. Open Prism365 for the file.",
  } as const;
  return {
    kind: "export",
    title: `Scheduled audit: ${cleanText(input.title)}`,
    severity: null,
    count: input.rows,
    items: [],
    lines: [
      `${input.rows} row${input.rows === 1 ? "" : "s"}${input.complete ? "" : ", from a partial result"}${input.scoped ? ", limited to some account domains" : ""}.`,
      input.attached ? "The CSV is attached." : withheld[input.withheld ?? "webhook"],
    ],
    link: `/#/audits/${encodeURIComponent(input.auditId)}`,
  };
}

// The weekly summary ------------------------------------------------------------------------------

interface ZonedParts {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  // 0 for Monday through 6 for Sunday.
  weekday: number;
}

const weekdays = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

function zonedParts(date: Date, timeZone: string): ZonedParts {
  const parts = Object.fromEntries(
    new Intl.DateTimeFormat("en-US", {
      timeZone,
      hourCycle: "h23",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      weekday: "short",
    })
      .formatToParts(date)
      .map((part) => [part.type, part.value]),
  );
  return {
    year: Number(parts.year),
    month: Number(parts.month),
    day: Number(parts.day),
    hour: Number(parts.hour),
    minute: Number(parts.minute),
    weekday: Math.max(0, weekdays.indexOf(String(parts.weekday))),
  };
}

// The instant a wall-clock time in the zone happens, found by correcting a first guess with the
// zone's offset at it, twice, which settles across a daylight saving change.
function zonedTime(
  date: { year: number; month: number; day: number },
  hour: number,
  timeZone: string,
): Date {
  const wanted = Date.UTC(date.year, date.month - 1, date.day, hour, 0);
  let guess = wanted;
  for (let round = 0; round < 3; round += 1) {
    const shown = zonedParts(new Date(guess), timeZone);
    guess += wanted - Date.UTC(shown.year, shown.month - 1, shown.day, shown.hour, shown.minute);
  }
  return new Date(guess);
}

// The latest Monday 07:00 in the time zone at or before `now`: when this week's summary is due.
export function weeklySummarySlot(now: Date, timeZone: string): Date {
  const local = zonedParts(now, timeZone);
  for (const back of [0, 7]) {
    const monday = new Date(
      Date.UTC(local.year, local.month - 1, local.day - local.weekday - back),
    );
    const slot = zonedTime(
      {
        year: monday.getUTCFullYear(),
        month: monday.getUTCMonth() + 1,
        day: monday.getUTCDate(),
      },
      7,
      timeZone,
    );
    if (slot.getTime() <= now.getTime()) return slot;
  }
  return new Date(now.getTime() - 7 * 86_400_000);
}

const total = (counts: Record<string, number>) =>
  Object.values(counts).reduce((sum, value) => sum + value, 0);

// The This week page as counts only: no names, addresses or titles.
export function weeklySummaryAlert(digest: Digest): AlertMessage {
  const lines: string[] = [];
  const { signInFindings: signins, directoryChanges: changes, quietAccounts: quiet } = digest;
  lines.push(
    signins.state === "unavailable"
      ? "Sign-in findings: not available."
      : `New sign-in findings: ${total(signins.opened)}, of which ${signins.opened.critical} critical and ${signins.opened.high} high. ${signins.openTotal} open, ${signins.resolved} resolved.`,
  );
  lines.push(
    changes.state === "unavailable"
      ? "Directory changes: not available."
      : `Directory changes: ${changes.total}, of which ${changes.flaggedTotal} flagged.`,
  );
  lines.push(
    quiet.state === "unavailable"
      ? "Accounts that went quiet: not available."
      : `Accounts that went quiet: ${quiet.total}.`,
  );
  const moved = digest.auditDrift.audits.reduce(
    (sum, audit) => sum + audit.tiles.filter((tile) => tile.change !== 0).length,
    0,
  );
  lines.push(
    `Audit numbers that moved: ${moved}, across ${digest.auditDrift.compared} audits compared with an earlier run.`,
  );
  const { posts, incidents } = digest.microsoft;
  lines.push(
    digest.microsoft.state === "unavailable"
      ? "Microsoft: not available."
      : `Microsoft: ${posts.total} Message center posts with a deadline, ${incidents.opened} service incidents opened.`,
  );
  const { collection } = digest;
  lines.push(
    `Collection is ${collection.collecting ? "on" : "off"}. Failed collector runs: ${collection.failedRuns}. Gaps: ${collection.gaps}.`,
  );
  return {
    kind: "summary",
    title: "This week in Prism365",
    severity: null,
    count: 0,
    items: [],
    lines,
    link: "/#/week",
  };
}

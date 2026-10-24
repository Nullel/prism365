import { createHash } from "node:crypto";
import { GraphBoundaryError } from "@prism365/graph";
import { type CollectorContext, type FindingInput, FindingPolicyError } from "@prism365/modules";
import type { SqlClient } from "@prism365/storage";
import { z } from "zod";
import { CANARY_SETTING, parseCanaryList } from "./canaries.ts";
import type { SigninsAuditConfig } from "./config.ts";
import { MIN_SPLIT_MS, parseCursor, type SigninsCursor, WINDOW_MS } from "./cursor.ts";
import {
  CANARY_RULE,
  type CanaryEntry,
  type CanaryIndex,
  canaryDedupeKey,
  canaryIndex,
  type FailurePattern,
  type FailureRecord,
  failurePatterns,
  matchCanary,
  type Reason,
  RULES_VERSION,
  type SuccessHistory,
  severityFor,
  sprayingAddresses,
  successReasons,
  totalScore,
} from "./rules.ts";
import {
  credentialFailureCodes,
  graphSignInSchema,
  maskIp,
  mfaFailureCodes,
  normalizeSignIn,
  type SigninEvent,
  signInContext,
} from "./source.ts";

export const MODULE_ID = "signins.audit";
export const UNFAMILIAR_RULE = "signins.unfamiliar_signin";
export const COLLECTION_GAP_RULE = "signins.collection_gap";

const PAGE_SIZE = 100;
const DETECTION_BATCH = 2_000;

// Thrown when a window cannot be finished in this run: the budget ran out, or the window holds
// more than one walk can page through. Nothing from that window is stored, so the next run reads
// it again from its start, smaller if it was the run's first window.
class RunBudgetSpentError extends Error {
  constructor() {
    super("The run's time budget is spent.");
    this.name = "RunBudgetSpentError";
  }
}

interface Budget {
  deadline: number;
  startedAt: number;
  // True while a pass reads its first window. If that window cannot be finished, the next run
  // tries a smaller one.
  first: boolean;
  // True while that first window has had the whole run's budget to itself. Only then does a
  // window that still cannot be finished count as a gap rather than a retry. A timed-out request
  // spends part of the budget, so it clears this.
  whole: boolean;
  // The deadline once the first window is done: the main lane's first window may use the whole
  // run, its later windows leave the settle lane its share.
  later?: number;
}

const MAX_WALK_ROUNDS = 50;
// Window size of the best-effort re-read of recent, not yet settled sign-ins.
const FRESH_WINDOW_MS = 10 * 60_000;
// A five-minute window whose requests keep timing out is recorded as a gap after this many runs.
const TIMEOUT_STRIKES = 3;

const iso = (value: Date) => value.toISOString();

// Finding keys must fit 200 characters. Directory values are not ours to bound, so a long key
// becomes a stable hash instead of a validation failure that would repeat on every run.
export function findingKey(key: string): string {
  return key.length <= 200 ? key : `sha256:${createHash("sha256").update(key).digest("hex")}`;
}
const min = (a: Date, b: Date) => (a.getTime() < b.getTime() ? a : b);
const max = (a: Date, b: Date) => (a.getTime() > b.getTime() ? a : b);

const sharedLimitations = [
  "Locations come from Microsoft's IP geolocation and are approximate; VPNs, mobile networks and cloud proxies move them.",
  "Sign-in records can arrive late or be missing; this reflects what Microsoft Graph returned when Prism365 collected.",
  "This is security evidence about an account, not a measure of the person's work.",
];

interface WindowResult {
  events: SigninEvent[];
  complete: boolean;
  skipped: number;
  // Smallest windows Microsoft returned more sign-ins for than could be read.
  gaps: { from: Date; to: Date }[];
  // Paging stopped at the deadline: what was read is kept, the rest of the window is not read.
  stopped?: boolean;
}

// Records are validated one at a time: an unreadable record is left out and counted, so it can
// never stall collection for the whole tenant on every later run. No request starts past the
// deadline, not even a later page of this window, so a run always keeps the time it needs to save
// its place before the scheduler's limit.
async function readWindow(
  context: CollectorContext,
  config: SigninsAuditConfig,
  from: Date,
  to: Date,
  options: { deadline: number; before?: boolean },
): Promise<WindowResult> {
  if (Date.now() > options.deadline) throw new RunBudgetSpentError();
  // `before`: up to but not including `to`, used to read under a millisecond already given up on.
  const query = new URLSearchParams({
    $filter: `createdDateTime ge ${iso(from)} and createdDateTime ${options.before ? "lt" : "le"} ${iso(to)}`,
    $top: String(PAGE_SIZE),
  });
  const page = await context.graph.getCollection(`auditLogs/signIns?${query}`, signInContext, {
    schema: z.unknown(),
    maxPages: 100,
    maxRecords: config.windowRecordLimit,
    deadline: options.deadline,
  });
  const events: SigninEvent[] = [];
  let skipped = 0;
  for (const item of page.items) {
    const parsed = graphSignInSchema.safeParse(item);
    if (parsed.success) events.push(normalizeSignIn(parsed.data));
    else skipped += 1;
  }
  return {
    events,
    complete: page.complete,
    skipped,
    gaps: [],
    ...(page.stopped === "deadline" ? { stopped: true } : {}),
  };
}

const oldest = (events: readonly SigninEvent[]) =>
  events.reduce<Date | null>(
    (current, event) => (!current || event.createdAt < current ? event.createdAt : current),
    null,
  );

// One Graph time window. Graph returns newest first and stops at the record cap, so a capped
// window is paged backwards: ask again for everything up to the oldest record received, keeping
// every record already read, until the window is exhausted. Records at the boundary come back
// twice and are de-duplicated on insert. A window Microsoft answers too slowly is split in half
// instead, down to the five-minute minimum.
//
// When a window cannot be finished (the budget ran out between or during requests, or it needs
// more than MAX_WALK_ROUNDS requests), the whole window is given up with RunBudgetSpentError, so
// the lane retries it next time, smaller if it was the pass's first window. Only a minimum-size
// window that had the whole run's budget keeps what it read and reports a gap; one that did not
// is owed the whole budget next run. More than the cap at one millisecond is recorded as a gap
// for that millisecond alone, and the walk continues below it.
async function fetchWindow(
  context: CollectorContext,
  config: SigninsAuditConfig,
  from: Date,
  to: Date,
  budget: Budget,
  depth = 0,
): Promise<WindowResult> {
  const minimum = to.getTime() - from.getTime() <= MIN_SPLIT_MS;
  let first: WindowResult;
  try {
    first = await readWindow(context, config, from, to, { deadline: budget.deadline });
  } catch (error) {
    if (
      !(error instanceof GraphBoundaryError && error.reason === "timeout") ||
      minimum ||
      depth >= 10
    )
      throw error;
    // The timed-out request used part of the budget: no half of this window has had all of it.
    budget.whole = false;
    const middle = new Date(from.getTime() + Math.floor((to.getTime() - from.getTime()) / 2));
    const older = await fetchWindow(context, config, from, middle, budget, depth + 1);
    const newer = await fetchWindow(context, config, middle, to, budget, depth + 1);
    return {
      events: [...older.events, ...newer.events],
      complete: older.complete && newer.complete,
      skipped: older.skipped + newer.skipped,
      gaps: [...older.gaps, ...newer.gaps],
    };
  }
  if (first.complete) return first;
  if (first.stopped) {
    if (minimum && budget.whole) return { ...first, stopped: false, gaps: [{ from, to }] };
    throw new RunBudgetSpentError();
  }
  const events = [...first.events];
  const gaps: { from: Date; to: Date }[] = [];
  let skipped = first.skipped;
  let floor = oldest(first.events);
  // The next request reads up to `upper`: one millisecond past the oldest record received
  // (Microsoft keeps finer timestamps than a JavaScript date, so a bound at that record's own
  // millisecond could cut off others in it), or strictly before a millisecond given up on.
  let upper = floor ? new Date(floor.getTime() + 1) : to;
  let before = false;
  for (let round = 0; floor; round += 1) {
    if (round >= MAX_WALK_ROUNDS || Date.now() > budget.deadline) {
      if (minimum && budget.whole)
        return { events, complete: false, skipped, gaps: [...gaps, { from, to }] };
      throw new RunBudgetSpentError();
    }
    const next = await readWindow(context, config, from, upper, {
      deadline: budget.deadline,
      before,
    });
    events.push(...next.events);
    skipped += next.skipped;
    if (next.complete) return { events, complete: gaps.length === 0, skipped, gaps };
    if (next.stopped) {
      if (minimum && budget.whole)
        return { events, complete: false, skipped, gaps: [...gaps, { from, to }] };
      throw new RunBudgetSpentError();
    }
    const reached = oldest(next.events);
    if (!reached) break;
    if (!before && reached.getTime() >= floor.getTime()) {
      // More than the cap share this millisecond: no smaller window can page past it. Keep what
      // was read of it, record that millisecond as a gap, and read on below it.
      gaps.push({ from: floor, to: new Date(floor.getTime() + 1) });
      upper = floor;
      before = true;
      continue;
    }
    floor = reached;
    upper = new Date(floor.getTime() + 1);
    before = false;
  }
  return { events, complete: false, skipped, gaps: [...gaps, { from, to }] };
}

// A window Microsoft returned more sign-ins for than Prism365 could read. The records that were
// read are kept; this finding says the rest are missing, so a dense burst (which can itself be
// an attack) is never silently dropped.
function gapFinding(from: Date, to: Date, read: number): FindingInput {
  const oneMillisecond = to.getTime() - from.getTime() <= 1;
  return {
    ruleId: COLLECTION_GAP_RULE,
    ruleVersion: RULES_VERSION,
    dedupeKey: findingKey(iso(from)),
    purpose: "operations",
    severity: "medium",
    score: 40,
    subject: { kind: "service", id: MODULE_ID, label: "Sign-in collection" },
    title: "Some sign-ins could not be read",
    summary: (oneMillisecond
      ? `More sign-ins share the millisecond ${iso(from)} than Microsoft returns in one request. ${read} were read and kept; the rest of that millisecond is missing from history and findings. Check the Entra sign-in log for that moment: a burst this dense can be an attack.`
      : `Microsoft returned more sign-ins between ${iso(from)} and ${iso(to)} than Prism365 could read in one run. ${read} were read and kept; the rest are missing from history and findings. Check the Entra sign-in log for that span: a burst this dense can be an attack.`
    ).slice(0, 1_000),
    evidence: [
      {
        source: "Prism365 sign-in collector",
        ref: `${iso(from)}/${iso(to)}`,
        observedAt: iso(to),
        fields: { read, windowStart: iso(from), windowEnd: iso(to) },
      },
    ],
    limitations: [
      "Prism365 reads one busy window at a time within a run's time budget; Microsoft keeps the full log for 30 days.",
    ],
    window: { start: from, end: to },
    observedAt: to,
  };
}

// A span Prism365 could not read at all: Microsoft did not answer in time on several runs in a
// row, or each run was stopped while reading it. Nothing was read, so this is not a burst.
function unansweredFinding(from: Date, to: Date): FindingInput {
  return {
    ruleId: COLLECTION_GAP_RULE,
    ruleVersion: RULES_VERSION,
    dedupeKey: findingKey(`unanswered:${iso(from)}`),
    purpose: "operations",
    severity: "medium",
    score: 40,
    subject: { kind: "service", id: MODULE_ID, label: "Sign-in collection" },
    title: "Some sign-ins could not be read",
    summary: `Prism365 could not read the sign-ins between ${iso(from)} and ${iso(to)}: Microsoft did not answer in time, or the run was stopped while reading, on ${TIMEOUT_STRIKES} runs in a row. Collection moved past that span, so any sign-ins in it are missing from history and findings. Check the Entra sign-in log for it.`,
    evidence: [
      {
        source: "Prism365 sign-in collector",
        ref: `${iso(from)}/${iso(to)}`,
        observedAt: iso(to),
        fields: { attempts: TIMEOUT_STRIKES, windowStart: iso(from), windowEnd: iso(to) },
      },
    ],
    limitations: [
      "The span may hold no sign-ins at all; Prism365 could not tell, because Microsoft never answered for it.",
    ],
    window: { start: from, end: to },
    observedAt: to,
  };
}

async function insertEvents(
  context: CollectorContext,
  events: readonly SigninEvent[],
  backfilled: boolean,
): Promise<number> {
  const table = context.table("event");
  return context.db(async (sql) => {
    let inserted = 0;
    for (const event of events) {
      const result = await sql.query(
        `INSERT INTO ${table} (tenant_id, id, created_at, user_id, upn, display_name, app_id, app_name,
           resource_name, client_app, ip, country, state, city, lat, lon, device_id, device_os,
           device_browser, is_compliant, is_managed, trust_type, error_code, failure_reason, ca_status,
           risk_level, risk_level_aggregated, risk_state, risk_event_types, is_interactive, backfilled,
           processed, collected_at)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,
           $24,$25,$26,$27,$28,$29::jsonb,$30,$31,false,$32)
         ON CONFLICT (tenant_id, id) DO NOTHING`,
        [
          context.tenantId,
          event.id,
          iso(event.createdAt),
          event.userId,
          event.upn,
          event.displayName,
          event.appId,
          event.appName,
          event.resourceName,
          event.clientApp,
          event.ip,
          event.country,
          event.state,
          event.city,
          event.lat,
          event.lon,
          event.deviceId,
          event.deviceOs,
          event.deviceBrowser,
          event.isCompliant,
          event.isManaged,
          event.trustType,
          event.errorCode,
          event.failureReason,
          event.caStatus,
          event.riskLevel,
          event.riskLevelAggregated,
          event.riskState,
          JSON.stringify(event.riskEventTypes),
          event.isInteractive,
          backfilled,
          iso(context.now),
        ],
      );
      inserted += result.rowCount;
    }
    return inserted;
  });
}

interface EventRow {
  id: string;
  created_at: Date | string;
  user_id: string;
  upn: string | null;
  display_name: string | null;
  app_id: string | null;
  app_name: string | null;
  resource_name: string | null;
  client_app: string | null;
  ip: string | null;
  country: string | null;
  state: string | null;
  city: string | null;
  lat: number | null;
  lon: number | null;
  device_id: string | null;
  device_os: string | null;
  device_browser: string | null;
  is_compliant: boolean | null;
  is_managed: boolean | null;
  trust_type: string | null;
  error_code: number;
  failure_reason: string | null;
  ca_status: string | null;
  risk_level: string | null;
  risk_level_aggregated: string | null;
  risk_state: string | null;
  risk_event_types: unknown;
  is_interactive: boolean | null;
  backfilled: boolean;
}

function fromRow(row: EventRow): SigninEvent & { backfilled: boolean } {
  const types =
    typeof row.risk_event_types === "string"
      ? JSON.parse(row.risk_event_types)
      : row.risk_event_types;
  return {
    id: row.id,
    createdAt: new Date(row.created_at),
    userId: row.user_id,
    upn: row.upn,
    displayName: row.display_name,
    appId: row.app_id,
    appName: row.app_name,
    resourceName: row.resource_name,
    clientApp: row.client_app,
    ip: row.ip,
    country: row.country,
    state: row.state,
    city: row.city,
    lat: row.lat === null ? null : Number(row.lat),
    lon: row.lon === null ? null : Number(row.lon),
    deviceId: row.device_id,
    deviceOs: row.device_os,
    deviceBrowser: row.device_browser,
    isCompliant: row.is_compliant,
    isManaged: row.is_managed,
    trustType: row.trust_type,
    errorCode: row.error_code,
    failureReason: row.failure_reason,
    caStatus: row.ca_status,
    riskLevel: row.risk_level,
    riskLevelAggregated: row.risk_level_aggregated,
    riskState: row.risk_state,
    riskEventTypes: Array.isArray(types)
      ? types.filter((item): item is string => typeof item === "string")
      : [],
    isInteractive: row.is_interactive,
    backfilled: row.backfilled,
  };
}

type SeenKind = "ip" | "country" | "state" | "device";
interface SeenEntry {
  trusted: boolean;
  pending: string | null;
}

function seenValues(event: SigninEvent): [SeenKind, string][] {
  const values: [SeenKind, string | null][] = [
    ["ip", event.ip],
    ["country", event.country],
    ["state", event.state],
    ["device", event.deviceId],
  ];
  return values.filter((entry): entry is [SeenKind, string] => Boolean(entry[1]));
}

const placeholders = (count: number, offset: number) =>
  Array.from({ length: count }, (_value, index) => `$${index + offset}`).join(",");

const failureCodeList = [...credentialFailureCodes, ...mfaFailureCodes].join(",");

function who(event: { displayName: string | null; upn: string | null }): string {
  return event.displayName ?? event.upn ?? "An account";
}

// Names come from the directory and are not ours to vet. If one trips the finding language
// policy (an account called "Productivity Team", say), record the finding without the name
// rather than failing the run on the same event forever.
async function upsertSafely(context: CollectorContext, finding: FindingInput): Promise<void> {
  const { label, ...anonymous } = finding.subject;
  const attempts: FindingInput[] = [
    finding,
    {
      ...finding,
      subject: anonymous,
      summary: label ? finding.summary.split(label).join("An account") : finding.summary,
    },
    {
      ...finding,
      subject: anonymous,
      summary:
        "A sign-in matched this rule. Directory names were left out because one conflicted with the finding language policy; open the evidence for details.",
      evidence: finding.evidence.map((item) => ({
        ...item,
        fields: Object.fromEntries(
          Object.entries(item.fields).filter(
            ([key, value]) =>
              typeof value !== "string" || ["network", "country", "state"].includes(key),
          ),
        ),
      })),
    },
  ];
  for (const [index, attempt] of attempts.entries()) {
    try {
      await context.findings.upsert(attempt);
      return;
    } catch (error) {
      if (!(error instanceof FindingPolicyError) || index === attempts.length - 1) throw error;
    }
  }
}

function unfamiliarFinding(
  event: SigninEvent,
  reasons: readonly Reason[],
  score: number,
  previous: SuccessHistory["previous"],
): FindingInput {
  const location = [event.city, event.state, event.country].filter(Boolean).join(", ");
  const signals = reasons.filter((reason) => reason.score > 0);
  const summary = `${who(event)} signed in successfully from ${location || "an unreported location"}${event.appName ? ` to ${event.appName}` : ""}. ${signals.map((reason) => reason.detail).join(" ")}`;
  return {
    ruleId: UNFAMILIAR_RULE,
    ruleVersion: RULES_VERSION,
    dedupeKey: findingKey(`${event.userId}|${event.country ?? "-"}|${event.state ?? "-"}`),
    purpose: "security",
    severity: severityFor(score),
    score,
    subject: {
      kind: "person",
      id: event.userId,
      ...(event.displayName || event.upn
        ? { label: (event.displayName ?? event.upn ?? "").slice(0, 200) }
        : {}),
    },
    title: "Sign-in that does not match this account's history",
    summary: summary.slice(0, 1_000),
    evidence: [
      {
        source: "Microsoft Entra sign-in logs",
        ref: event.id,
        observedAt: iso(event.createdAt),
        fields: {
          signals: signals
            .map((reason) => reason.rule)
            .join(", ")
            .slice(0, 500),
          score,
          network: maskIp(event.ip),
          country: event.country,
          state: event.state,
          city: event.city,
          app: event.appName,
          client: event.clientApp,
          device: [event.deviceOs, event.deviceBrowser].filter(Boolean).join(" / ") || null,
          compliantDevice: event.isCompliant,
          conditionalAccess: event.caStatus,
          microsoftRisk: event.riskLevel,
        },
      },
    ],
    limitations: [
      ...sharedLimitations,
      "Normal locations are learned from this account's own history in Prism365; a new office or a trip looks unfamiliar until this finding is resolved.",
    ],
    window: { start: previous?.createdAt ?? event.createdAt, end: event.createdAt },
    observedAt: event.createdAt,
  };
}

// A sign-in attempt by a canary account. Whatever its result, nobody should be using the account,
// so it is always critical. One finding per account and UTC day: later attempts that day update it
// and count up, and the evidence shown is the latest attempt.
export function canaryFinding(event: SigninEvent, canary: CanaryEntry): FindingInput {
  const name = (event.displayName ?? event.upn ?? canary.principal).slice(0, 200);
  const location = [event.city, event.state, event.country].filter(Boolean).join(", ");
  const result = event.errorCode === 0 ? "success" : event.errorCode === -1 ? "unknown" : "failure";
  const what =
    result === "success"
      ? "A successful sign-in"
      : result === "failure"
        ? `A failed sign-in attempt (error ${event.errorCode})`
        : "A sign-in attempt with no reported result";
  const summary = `${what} on ${name}, a canary account.${location ? ` It came from ${location}.` : ""}${event.appName ? ` The app was ${event.appName}.` : ""} Nobody should be using this account.`;
  return {
    ruleId: CANARY_RULE,
    ruleVersion: RULES_VERSION,
    dedupeKey: findingKey(canaryDedupeKey(canary, event.createdAt)),
    purpose: "security",
    severity: "critical",
    score: 100,
    subject: { kind: "person", id: canary.userId, label: name },
    title: "Sign-in attempt on a canary account",
    summary: summary.slice(0, 1_000),
    evidence: [
      {
        source: "Microsoft Entra sign-in logs",
        ref: event.id,
        observedAt: iso(event.createdAt),
        // The full address, unlike the other rules: this finding is about one account and the
        // address is what an administrator would act on. The same view is on the person's page.
        fields: {
          result,
          errorCode: event.errorCode,
          failureReason: event.failureReason?.slice(0, 300) ?? null,
          ip: event.ip,
          location: location || null,
          app: event.appName,
          client: event.clientApp,
        },
      },
    ],
    limitations: [
      "Canaries are accounts you marked; Prism365 cannot tell a mistaken sign-in from an attack.",
      ...sharedLimitations,
    ],
    window: { start: event.createdAt, end: event.createdAt },
    observedAt: event.createdAt,
  };
}

const failureCopy: Record<
  FailurePattern["rule"],
  {
    title: string;
    subject: "person" | "resource";
    summary: (pattern: FailurePattern, name: string, config: SigninsAuditConfig) => string;
  }
> = {
  failed_burst_ip: {
    title: "Repeated failed sign-ins from one address",
    subject: "resource",
    summary: (pattern, _name, config) =>
      `${pattern.count} failed sign-ins from ${pattern.key} within ${config.bruteForceMinutes} minutes, against ${pattern.accounts} account${pattern.accounts === 1 ? "" : "s"}.`,
  },
  password_spray: {
    title: "Password spray pattern from one address",
    subject: "resource",
    summary: (pattern, _name, config) =>
      `${pattern.key} failed to sign in to ${pattern.accounts} different accounts within ${config.sprayMinutes} minutes (${pattern.count} attempts).`,
  },
  failed_burst_account: {
    title: "Repeated failed sign-ins for one account",
    subject: "person",
    summary: (pattern, name, config) =>
      `${pattern.count} failed sign-ins for ${name} within ${config.bruteForceMinutes} minutes.`,
  },
  mfa_denials: {
    title: "Repeated failed MFA prompts for one account",
    subject: "person",
    summary: (pattern, name, config) =>
      `${pattern.count} failed MFA prompts for ${name} within ${config.mfaMinutes} minutes. Prompts the person did not start can mean someone else has the password.`,
  },
};

function failureFinding(
  pattern: FailurePattern,
  records: ReadonlyMap<string, FailureRecord>,
  names: ReadonlyMap<string, string>,
  config: SigninsAuditConfig,
): FindingInput {
  const copy = failureCopy[pattern.rule];
  const name = copy.subject === "person" ? (names.get(pattern.key) ?? "An account") : pattern.key;
  const limitations = [...sharedLimitations];
  if (pattern.rule === "mfa_denials")
    limitations.push(
      "Code 500121 is not in Microsoft's published error reference; its meaning comes from Microsoft Q&A answers. This rule is experimental.",
    );
  return {
    ruleId: `signins.${pattern.rule}`,
    ruleVersion: RULES_VERSION,
    dedupeKey: findingKey(pattern.key),
    purpose: "security",
    severity: severityFor(pattern.score),
    score: pattern.score,
    subject: {
      kind: copy.subject,
      id: pattern.key,
      label: (copy.subject === "person" ? name : pattern.key).slice(0, 200),
    },
    title: copy.title,
    summary: copy.summary(pattern, name, config).slice(0, 1_000),
    evidence: pattern.eventIds.slice(-10).map((id) => {
      const record = records.get(id);
      return {
        source: "Microsoft Entra sign-in logs",
        ref: id,
        observedAt: iso(record?.createdAt ?? pattern.end),
        fields: {
          network: maskIp(record?.ip ?? null),
          country: record?.country ?? null,
          errorCode: record?.errorCode ?? null,
        },
      };
    }),
    limitations,
    window: { start: pattern.start, end: pattern.end },
    observedAt: pattern.end,
  };
}

async function loadCanaries(context: CollectorContext): Promise<CanaryIndex> {
  return canaryIndex(parseCanaryList(await context.settings.get(CANARY_SETTING)).accounts);
}

// Values that took part in a flagged sign-in become normal once that finding is resolved.
async function promoteResolved(context: CollectorContext): Promise<void> {
  await context.db((sql) =>
    sql.query(
      `UPDATE ${context.table("seen")} AS s SET trusted = true, pending_finding = NULL
         FROM prism365_finding AS f
        WHERE s.tenant_id = $1 AND s.trusted = false AND s.pending_finding IS NOT NULL
          AND f.tenant_id = $1 AND f.module = $2 AND f.rule_id = $3
          AND f.dedupe_key = s.pending_finding AND f.state = 'resolved'`,
      [context.tenantId, MODULE_ID, UNFAMILIAR_RULE],
    ),
  );
}

async function loadBaselines(
  sql: SqlClient,
  context: CollectorContext,
  users: readonly string[],
): Promise<Map<string, Map<string, SeenEntry>>> {
  const baselines = new Map<string, Map<string, SeenEntry>>();
  for (let index = 0; index < users.length; index += 500) {
    const chunk = users.slice(index, index + 500);
    const result = await sql.query<{
      user_id: string;
      kind: SeenKind;
      value: string;
      trusted: boolean;
      pending_finding: string | null;
    }>(
      `SELECT user_id, kind, value, trusted, pending_finding FROM ${context.table("seen")}
        WHERE tenant_id = $1 AND user_id IN (${placeholders(chunk.length, 2)})`,
      [context.tenantId, ...chunk],
    );
    for (const row of result.rows) {
      const map = baselines.get(row.user_id) ?? new Map<string, SeenEntry>();
      map.set(`${row.kind}|${row.value}`, { trusted: row.trusted, pending: row.pending_finding });
      baselines.set(row.user_id, map);
    }
  }
  return baselines;
}

async function hourShare(
  sql: SqlClient,
  context: CollectorContext,
  config: SigninsAuditConfig,
  userId: string,
  at: Date,
  cache: Map<string, { first: number; byHour: Map<number, number>; total: number }>,
): Promise<number | null> {
  let profile = cache.get(userId);
  if (!profile) {
    const result = await sql.query<{ h: number; n: number; first: Date | string }>(
      `SELECT EXTRACT(HOUR FROM created_at AT TIME ZONE $3)::int AS h, count(*)::int AS n, min(created_at) AS first
         FROM ${context.table("event")}
        WHERE tenant_id = $1 AND user_id = $2 AND error_code = 0
        GROUP BY 1`,
      [context.tenantId, userId, config.timeZone],
    );
    const byHour = new Map(result.rows.map((row) => [row.h, row.n]));
    const firsts = result.rows.map((row) => new Date(row.first).getTime());
    profile = {
      first: firsts.length ? Math.min(...firsts) : Number.POSITIVE_INFINITY,
      byHour,
      total: result.rows.reduce((sum, row) => sum + row.n, 0),
    };
    cache.set(userId, profile);
  }
  if (profile.total === 0 || at.getTime() - profile.first < 30 * 86_400_000) return null;
  const hour = Number(
    new Intl.DateTimeFormat("en-US", {
      hour: "numeric",
      hourCycle: "h23",
      timeZone: config.timeZone,
    }).format(at),
  );
  return (profile.byHour.get(hour) ?? 0) / profile.total;
}

// Runs the rules over events that have not been processed yet, in time order, then records the
// baseline changes and marks the events processed in one transaction.
async function detect(
  context: CollectorContext,
  config: SigninsAuditConfig,
  canaries: CanaryIndex,
  through: Date,
): Promise<void> {
  const eventTable = context.table("event");
  const seenTable = context.table("seen");
  const lookbackMinutes =
    Math.max(config.sprayMinutes, config.bruteForceMinutes, config.mfaMinutes) +
    config.successAfterFailuresMinutes;
  for (;;) {
    const pending = await context.db((sql) =>
      sql.query<EventRow>(
        `SELECT * FROM ${eventTable}
          WHERE tenant_id = $1 AND processed = false AND created_at <= $2
          ORDER BY created_at, id LIMIT ${DETECTION_BATCH}`,
        [context.tenantId, iso(through)],
      ),
    );
    if (pending.rows.length === 0) return;
    const events = pending.rows.map(fromRow);
    const first = events[0]?.createdAt ?? through;
    const last = events.at(-1)?.createdAt ?? through;

    const prepared = await context.db(async (sql) => {
      const failures = await sql.query<{
        id: string;
        created_at: Date | string;
        user_id: string;
        ip: string | null;
        country: string | null;
        error_code: number;
      }>(
        `SELECT id, created_at, user_id, ip, country, error_code FROM ${eventTable}
          WHERE tenant_id = $1 AND error_code IN (${failureCodeList})
            AND created_at >= $2 AND created_at <= $3
          ORDER BY created_at, id LIMIT 50000`,
        [context.tenantId, iso(new Date(first.getTime() - lookbackMinutes * 60_000)), iso(last)],
      );
      const window: FailureRecord[] = failures.rows.map((row) => ({
        id: row.id,
        createdAt: new Date(row.created_at),
        userId: row.user_id,
        ip: row.ip,
        country: row.country,
        errorCode: row.error_code,
      }));
      const successes = events.filter((event) => event.errorCode === 0 && event.userId);
      const users = [...new Set(successes.map((event) => event.userId))];
      const baselines = await loadBaselines(sql, context, users);
      const previous = new Map<string, SuccessHistory["previous"]>();
      const hours = new Map<string, number | null>();
      const cache = new Map<
        string,
        { first: number; byHour: Map<number, number>; total: number }
      >();
      for (const event of successes) {
        const result = await sql.query<{
          created_at: Date | string;
          lat: number | null;
          lon: number | null;
        }>(
          `SELECT created_at, lat, lon FROM ${eventTable}
            WHERE tenant_id = $1 AND user_id = $2 AND error_code = 0 AND created_at < $3
            ORDER BY created_at DESC LIMIT 1`,
          [context.tenantId, event.userId, iso(event.createdAt)],
        );
        const row = result.rows[0];
        previous.set(
          event.id,
          row
            ? {
                createdAt: new Date(row.created_at),
                lat: row.lat === null ? null : Number(row.lat),
                lon: row.lon === null ? null : Number(row.lon),
              }
            : null,
        );
        hours.set(
          event.id,
          await hourShare(sql, context, config, event.userId, event.createdAt, cache),
        );
      }
      return { window, baselines, previous, hours };
    });

    // Canary accounts first: any attempt is a finding, so this needs no baseline, history or
    // score, and it applies to history being loaded as much as to live sign-ins.
    for (const event of events) {
      const canary = matchCanary(event, canaries);
      if (canary) await upsertSafely(context, canaryFinding(event, canary));
    }

    // Failure patterns that end at one of these events.
    const batchIds = new Set(events.map((event) => event.id));
    const freshFailures = prepared.window.filter((record) => batchIds.has(record.id));
    const records = new Map(prepared.window.map((record) => [record.id, record]));
    const names = new Map<string, string>();
    for (const event of events)
      if (event.userId && (event.displayName || event.upn)) names.set(event.userId, who(event));
    for (const pattern of failurePatterns(prepared.window, freshFailures, config))
      await upsertSafely(context, failureFinding(pattern, records, names, config));

    // Successful sign-ins, oldest first, against a baseline that grows as they are read.
    const spraying = sprayingAddresses(prepared.window, config);
    const writes: {
      userId: string;
      kind: SeenKind;
      value: string;
      at: Date;
      trusted: boolean;
      pending: string | null;
    }[] = [];
    for (const event of events) {
      if (event.errorCode !== 0 || !event.userId) continue;
      const baseline = prepared.baselines.get(event.userId) ?? new Map<string, SeenEntry>();
      prepared.baselines.set(event.userId, baseline);
      const trusted = (kind: SeenKind, value: string | null) =>
        value ? baseline.get(`${kind}|${value}`)?.trusted === true : false;
      const recent = prepared.window.filter(
        (record) =>
          record.userId === event.userId &&
          credentialFailureCodes.has(record.errorCode) &&
          record.createdAt.getTime() < event.createdAt.getTime() &&
          record.createdAt.getTime() >=
            event.createdAt.getTime() - config.successAfterFailuresMinutes * 60_000,
      );
      const history: SuccessHistory = {
        hasBaseline: [...baseline.values()].some((entry) => entry.trusted),
        trusted: {
          ip: trusted("ip", event.ip),
          country: trusted("country", event.country),
          state: trusted("state", event.state),
          device: trusted("device", event.deviceId),
        },
        previous: prepared.previous.get(event.id) ?? null,
        recentFailures: recent.length,
        recentSprayFailures: recent.filter((record) => record.ip && spraying.has(record.ip)).length,
        hourShare: prepared.hours.get(event.id) ?? null,
      };
      const learning = event.backfilled;
      const reasons = successReasons(event, history, config, learning);
      const score = totalScore(reasons);
      let flagged: string | null = null;
      if (score >= config.minimumScore) {
        const finding = unfamiliarFinding(event, reasons, score, history.previous);
        await upsertSafely(context, finding);
        flagged = finding.dedupeKey;
      }
      for (const [kind, value] of seenValues(event)) {
        const key = `${kind}|${value}`;
        const existing = baseline.get(key);
        // A value in a flagged sign-in stays untrusted; an untrusted value stays untrusted until
        // its finding is resolved, even if later sign-ins from it are quiet.
        const nextTrusted = existing ? existing.trusted : flagged === null;
        const nextPending = existing ? existing.pending : flagged;
        baseline.set(key, { trusted: nextTrusted, pending: nextTrusted ? null : nextPending });
        writes.push({
          userId: event.userId,
          kind,
          value,
          at: event.createdAt,
          trusted: nextTrusted,
          pending: nextTrusted ? null : nextPending,
        });
      }
    }

    await context.db(async (sql) => {
      for (const write of writes)
        await sql.query(
          `INSERT INTO ${seenTable} (tenant_id, user_id, kind, value, first_seen, last_seen, trusted, pending_finding)
           VALUES ($1,$2,$3,$4,$5,$5,$6,$7)
           ON CONFLICT (tenant_id, user_id, kind, value) DO UPDATE SET
             last_seen = GREATEST(${seenTable}.last_seen, EXCLUDED.last_seen),
             first_seen = LEAST(${seenTable}.first_seen, EXCLUDED.first_seen)`,
          [
            context.tenantId,
            write.userId,
            write.kind,
            write.value,
            iso(write.at),
            write.trusted,
            write.pending,
          ],
        );
      await sql.query(
        `UPDATE ${eventTable} SET processed = true
          WHERE tenant_id = $1 AND id IN (${placeholders(events.length, 2)})`,
        [context.tenantId, ...events.map((event) => event.id)],
      );
    });
    if (pending.rows.length < DETECTION_BATCH) return;
  }
}

async function collectRange(
  context: CollectorContext,
  config: SigninsAuditConfig,
  canaries: CanaryIndex,
  from: Date,
  to: Date,
  backfilled: boolean,
  budget: Budget,
  recordGaps = true,
): Promise<boolean> {
  const fetched = await fetchWindow(context, config, from, to, budget);
  context.count.fetched(fetched.events.length + fetched.skipped);
  context.count.skipped(fetched.skipped);
  context.count.inserted(await insertEvents(context, fetched.events, backfilled));
  await detect(context, config, canaries, to);
  for (const gap of recordGaps ? fetched.gaps : []) {
    const read = new Set(
      fetched.events
        .filter((event) => event.createdAt >= gap.from && event.createdAt <= gap.to)
        .map((event) => event.id),
    ).size;
    await upsertSafely(context, gapFinding(gap.from, gap.to, read));
  }
  return fetched.complete && fetched.skipped === 0;
}

export interface SigninsRunResult {
  // false when records were left out: unreadable ones, or a window too busy to read in full.
  complete: boolean;
  // true when the run stopped at its time budget with more to read now.
  more: boolean;
}

// One reading position and its window state (see cursor.ts).
interface Lane {
  place: Date;
  span: number;
  hold: Date | null;
  strikes: { at: string; count: number } | undefined;
  trying: string | undefined;
  // A minimum-size window of this lane could not be read without the whole run's budget: the
  // next run reads this lane first.
  owed: boolean;
}

type LaneState = Pick<SigninsCursor, "span" | "hold" | "strikes" | "trying" | "owed">;

function toLane(place: string, state: LaneState): Lane {
  return {
    place: new Date(place),
    span: state.span ?? WINDOW_MS,
    hold: state.hold ? new Date(state.hold) : null,
    strikes: state.strikes,
    trying: state.trying,
    owed: state.owed === true,
  };
}

function laneState(lane: Lane): LaneState {
  return {
    ...(lane.span < WINDOW_MS ? { span: lane.span } : {}),
    ...(lane.hold ? { hold: iso(lane.hold) } : {}),
    ...(lane.strikes ? { strikes: lane.strikes } : {}),
    ...(lane.trying ? { trying: lane.trying } : {}),
    ...(lane.owed ? { owed: true as const } : {}),
  };
}

// The next run tries half of a window that could not be finished, and keeps windows that small
// until the lane is past it.
function shrink(lane: Lane, from: Date, end: Date): void {
  lane.span = Math.max(MIN_SPLIT_MS, Math.floor((end.getTime() - from.getTime()) / 2));
  lane.hold = max(end, lane.hold ?? end);
}

// A window that timed out, or that a run was stopped in: a larger one is shrunk; a minimum-size
// one is retried a few runs, then given up, so one stretch cannot stall the lane.
function windowFailed(lane: Lane, from: Date, end: Date): "retry" | "give_up" {
  if (end.getTime() - from.getTime() > MIN_SPLIT_MS) {
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

// One collector run. The first run loads up to backfillDays of history, oldest first (learning
// baselines, and surfacing only signals that do not depend on a baseline). Later runs follow new
// sign-ins on the main lane, then re-read behind it on the settle lane, so records Microsoft
// publishes up to lateArrivalMinutes late are collected. Both lanes save their place after every
// window, so a slow or stopped run leaves a backlog for the next one, never a hole.
export async function collectSignins(
  context: CollectorContext,
  config: SigninsAuditConfig,
): Promise<SigninsRunResult> {
  const startedAt = Date.now();
  const totalDeadline = startedAt + config.runBudgetSeconds * 1_000;
  const now = context.now;
  const lateMs = config.lateArrivalMinutes * 60_000;
  // Everything before this has been published long enough to be read for the last time.
  const settledBy = new Date(now.getTime() - lateMs);
  const oldestAllowed = new Date(now.getTime() - 30 * 86_400_000);
  let complete = true;

  await promoteResolved(context);
  // Read once: marking an account as a canary takes effect from the next run.
  const canaries = await loadCanaries(context);
  let cursor = parseCursor(await context.cursor.get());
  if (!cursor) {
    const start = new Date(now.getTime() - config.backfillDays * 86_400_000);
    if (config.backfillDays <= 0) {
      await context.cursor.set(
        JSON.stringify({ v: 1, phase: "live", through: iso(now), settle: { at: iso(now) } }),
      );
      return { complete, more: false };
    }
    cursor = { v: 1, phase: "backfill", start: iso(start), next: iso(start) };
  }
  const backfill = cursor.phase === "backfill" ? cursor : null;
  const main = toLane(cursor.phase === "backfill" ? cursor.next : cursor.through, cursor);
  const settle =
    cursor.phase === "live"
      ? toLane(
          cursor.settle?.at ??
            iso(min(settledBy, new Date(new Date(cursor.through).getTime() - lateMs))),
          cursor.settle ?? {},
        )
      : null;
  // Where the main lane stood when this run began.
  const savedThrough = new Date(cursor.phase === "live" ? cursor.through : cursor.next);
  const persist = () =>
    context.cursor.set(
      JSON.stringify(
        backfill
          ? {
              v: 1,
              phase: "backfill",
              start: backfill.start,
              next: iso(main.place),
              ...laneState(main),
            }
          : {
              v: 1,
              phase: "live",
              through: iso(main.place),
              ...laneState(main),
              ...(settle ? { settle: { at: iso(settle.place), ...laneState(settle) } } : {}),
            },
      ),
    );

  // A lane still marked as reading a window was stopped there by the last run: count that window
  // as failed, so the next attempt is smaller or, after a few, passes it.
  const recoverStopped = async (lane: Lane, until: Date) => {
    const stopped = lane.trying === iso(lane.place) && lane.place.getTime() < until.getTime();
    lane.trying = undefined;
    if (!stopped) return;
    const end = min(new Date(lane.place.getTime() + lane.span), until);
    if (windowFailed(lane, lane.place, end) === "give_up") {
      await upsertSafely(context, unansweredFinding(lane.place, end));
      complete = false;
      lane.place = end;
    }
  };
  await recoverStopped(main, now);
  if (settle) await recoverStopped(settle, min(main.place, settledBy));
  await persist();

  // Reads one lane forward to `until`, one window at a time, saving its place after each.
  // Returns true when it got there, false when its budget ran out first.
  const walk = async (
    lane: Lane,
    until: Date,
    budget: Budget,
    backfilled: boolean,
    afterRead?: (from: Date, to: Date) => void,
  ): Promise<boolean> => {
    lane.place = max(lane.place, oldestAllowed);
    while (lane.place.getTime() < until.getTime()) {
      if (Date.now() > budget.deadline) return false;
      const from = lane.place;
      const end = min(new Date(from.getTime() + lane.span), until);
      // Saved before reading, so a run stopped mid-window is visible to the next one.
      lane.trying = iso(from);
      await persist();
      const began = Date.now();
      let read: boolean;
      try {
        read = await collectRange(context, config, canaries, from, end, backfilled, budget);
      } catch (error) {
        lane.trying = undefined;
        if (error instanceof RunBudgetSpentError) {
          if (budget.first) {
            // The pass's first window could not be finished: smaller next time. One already at
            // the minimum that lacked the whole budget gets it next run instead.
            shrink(lane, from, end);
            if (!budget.whole && end.getTime() - from.getTime() <= MIN_SPLIT_MS) lane.owed = true;
          }
          await persist();
          return false;
        }
        if (!(error instanceof GraphBoundaryError && error.reason === "timeout")) {
          await persist();
          throw error;
        }
        if (windowFailed(lane, from, end) === "retry") {
          await persist();
          throw error;
        }
        await upsertSafely(context, unansweredFinding(from, end));
        read = false;
      }
      lane.trying = undefined;
      if (!read) complete = false;
      budget.first = false;
      budget.whole = false;
      if (budget.later !== undefined) {
        budget.deadline = Math.min(budget.deadline, budget.later);
        delete budget.later;
      }
      lane.owed = false;
      lane.place = max(end, lane.place);
      if (read) afterRead?.(from, end);
      if (lane.hold && lane.place.getTime() >= lane.hold.getTime()) lane.hold = null;
      // Grow back toward full windows once past the busy stretch and windows are quick again.
      if (!lane.hold && Date.now() - began < (budget.deadline - budget.startedAt) / 4)
        lane.span = Math.min(WINDOW_MS, lane.span * 2);
      if (lane.strikes && lane.strikes.at !== iso(lane.place)) lane.strikes = undefined;
      await persist();
    }
    return true;
  };

  if (backfill) {
    const done = await walk(
      main,
      now,
      { deadline: totalDeadline, startedAt, first: true, whole: true },
      true,
    );
    if (!done) return { complete, more: true };
    // Settled: what this run read that is lateArrivalMinutes old, and what earlier runs read that
    // was that old when they read it. Earlier runs read up to where this one started, so anything
    // lateArrivalMinutes before that point qualifies; the rest is re-read on the settle lane.
    const settled = min(settledBy, new Date(savedThrough.getTime() - lateMs));
    await context.cursor.set(
      JSON.stringify({
        v: 1,
        phase: "live",
        through: iso(now),
        settle: { at: iso(max(settled, new Date(backfill.start))) },
      }),
    );
    return { complete, more: false };
  }
  if (!settle) return { complete, more: false };

  // The lane that reads first has the whole run's budget for its first window: normally the main
  // lane, or the settle lane when it is owed that (see walk). When the settle lane has a backlog,
  // the main lane's later windows stop at two thirds of the budget. A settle lane that is behind
  // and gets no time at all, because the main lane's window took the whole run, is owed the next
  // run, so heavy new traffic delays the re-read by a run at most, never indefinitely.
  const settleFrom = settle.place.getTime();
  const settleTo = () => min(main.place, settledBy);
  const backlog = settleTo().getTime() - settle.place.getTime();
  const share = startedAt + Math.floor(((totalDeadline - startedAt) * 2) / 3);
  const readSettle = async (whole: boolean) => {
    if (settle.place.getTime() >= settleTo().getTime()) return;
    if (Date.now() > totalDeadline) {
      settle.owed = true;
      await persist();
      return;
    }
    try {
      await walk(
        settle,
        settleTo(),
        { deadline: totalDeadline, startedAt: Date.now(), first: true, whole },
        false,
      );
    } catch (error) {
      // A slow re-read is already recorded on the lane (smaller window, or a strike); the new
      // sign-ins this run read are kept and the run stays a success.
      if (error instanceof GraphBoundaryError && error.reason === "timeout") return;
      // Any other failure must not let an owed settle lane go first again and block new
      // sign-ins every run: it waits its turn behind the main lane until it reads cleanly.
      settle.owed = false;
      await persist().catch(() => {});
      throw error;
    }
  };
  const settleFirst = settle.owed && !main.owed;
  if (settleFirst) await readSettle(true);
  const caughtUp = await walk(
    main,
    now,
    {
      deadline: totalDeadline,
      startedAt,
      first: true,
      whole: !settleFirst,
      ...(backlog > 0 && !settleFirst ? { later: share } : {}),
    },
    false,
    // A window read when it was already old enough is settled by that read: the settle lane does
    // not read it again if it is contiguous with what is already settled.
    (from, to) => {
      if (settle.place.getTime() >= from.getTime())
        settle.place = max(settle.place, min(to, settledBy));
    },
  );
  await persist();
  if (!settleFirst) await readSettle(false);

  // Best effort, newest first: what was read before this run and is not settled yet is read
  // again, so a record Microsoft published a few minutes late shows up now rather than when it
  // settles. It uses only time left over and saves nothing; the settle lane is what guarantees
  // every such record is read once it is lateArrivalMinutes old.
  const freshFloor = max(settle.place, oldestAllowed);
  let freshTo = min(savedThrough, main.place);
  const freshBudget: Budget = {
    deadline: totalDeadline,
    startedAt: Date.now(),
    first: false,
    whole: false,
  };
  while (freshTo.getTime() > freshFloor.getTime() && Date.now() <= totalDeadline) {
    const from = max(freshFloor, new Date(freshTo.getTime() - FRESH_WINDOW_MS));
    try {
      await collectRange(context, config, canaries, from, freshTo, false, freshBudget, false);
    } catch (error) {
      if (error instanceof RunBudgetSpentError) break;
      if (error instanceof GraphBoundaryError && error.reason === "timeout") break;
      throw error;
    }
    freshTo = from;
  }

  const settleBehind = settle.place.getTime() < settleTo().getTime();
  return {
    complete,
    more:
      !caughtUp ||
      main.owed ||
      settle.owed ||
      (settleBehind && settle.place.getTime() > settleFrom),
  };
}

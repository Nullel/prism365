import {
  actionLabel,
  type PersonSummary,
  type Severity,
  type SummaryAudit,
} from "@prism365/contracts";
import { number, relativeTime } from "./common.tsx";
import { INACTIVE_FROM_DAYS, joinWords } from "./personAccessSummary.ts";
import type { PersonTab } from "./PersonTabs.tsx";

// What the Overview's "Where they show up" panel says: one short sentence per source, as plain
// data. The panel is an index, not a report: each sentence answers "anything here?" and the row
// opens the detail. This file decides which sources get a row, in what order and column, how each
// reads when it has nothing to say, and where its row goes. Pure (the clock is a parameter), so
// every wording and every choice about what to leave out is a test.

// Where a row sends the administrator: a section of one of this page's tabs, or a page elsewhere
// in the app (a plain hash link).
export type Jump = { kind: "section"; tab: PersonTab; id: string } | { kind: "page"; href: string };

export type Tone = "good" | "warning" | "bad";

export interface SummaryItem {
  id: string;
  source: string;
  // The panel has two columns: risk and access on the left, the account itself on the right.
  column: "risk" | "account";
  // A rating shown with its name beside it, never as colour alone.
  severity: Severity | null;
  // Words to show instead of the rating's own name ("Privileged").
  severityLabel?: string;
  // The state the sentence itself describes, in the shared style.
  tone: Tone | null;
  sentence: string;
  // When the source was last run or exported; null when it has no run to date.
  ranAt: string | null;
  // The run could not check everything.
  incomplete: boolean;
  // Nothing to report about this person: the row is shown in the muted style.
  muted: boolean;
  jump: Jump;
  // Where the row goes, in words: "the Access tab".
  destination: string;
}

export interface SummaryLines {
  items: SummaryItem[];
  // Sources that have never run, folded into one last line.
  notRun: string[];
}

// The ids of the sections the rows scroll to. PersonAccess.tsx, People.tsx and PersonChanges.tsx
// put them on their sections, and a test holds the two lists together.
export const sectionIds = {
  combinedRisks: "person-section-combined-risks",
  findings: "person-section-findings",
  groups: "person-section-groups",
  owns: "person-section-owns",
  sharedMailboxes: "person-section-shared-mailboxes",
  methods: "person-section-sign-in-methods",
  devices: "person-section-devices",
  mailbox: "person-section-mailbox",
  password: "person-section-password",
  risk: "person-section-risk",
  activity: "person-section-activity",
  guest: "person-section-guest",
  apps: "person-section-m365-apps",
  changes: "person-section-changes",
  receipts: "person-section-receipts",
} as const;

const severityNames: Record<Severity, string> = {
  critical: "critical",
  high: "high",
  medium: "medium",
  low: "low",
  info: "info",
};

const plural = (count: number, one: string, many = `${one}s`) =>
  `${number.format(count)} ${count === 1 ? one : many}`;

type Row = Extract<SummaryAudit, { status: "ok" }>["rows"][number];
const cell = (row: Row, key: string) => row.cells.find((candidate) => candidate.key === key)?.value;
const textOf = (row: Row, key: string): string | null => {
  const value = cell(row, key);
  return typeof value === "string" && value !== "" ? value : null;
};
const numberOf = (row: Row, key: string): number | null => {
  const value = cell(row, key);
  return typeof value === "number" ? value : null;
};
const flagOf = (row: Row, key: string): boolean | null => {
  const value = cell(row, key);
  return typeof value === "boolean" ? value : null;
};
const listOf = (row: Row, key: string): string[] => {
  const value = cell(row, key);
  return Array.isArray(value)
    ? value.filter((item): item is string => typeof item === "string")
    : [];
};
const lowerFirst = (value: string) => value.charAt(0).toLowerCase() + value.slice(1);
const gigabytes = (value: number) => number.format(Math.round(value * 10) / 10);

interface Said {
  sentence: string;
  tone?: Tone;
  severity?: Severity;
  severityLabel?: string;
}
type Wording = (rows: readonly Row[], now: number) => Said | null;

const wordings: Record<string, Wording> = {
  "identity.risky-users": (rows) => {
    const row = rows[0];
    if (!row) return null;
    const level = textOf(row, "level");
    const state = textOf(row, "state");
    const lead = level ? `${level} risk` : "At risk";
    return {
      sentence: state ? `${lead}, ${lowerFirst(state)}` : lead,
      tone: level === "Low" ? "warning" : "bad",
    };
  },
  "directory.privileged-roles": (rows) => {
    const first = rows[0];
    if (!first) return null;
    const name = textOf(first, "role") ?? "A role";
    const privileged = rows.some((row) => flagOf(row, "privileged") === true);
    const more = rows.length - 1;
    return {
      sentence: more > 0 ? `${name} and ${plural(more, "more role")}` : name,
      ...(privileged ? { severity: "high" as const, severityLabel: "Privileged" } : {}),
    };
  },
  "mfa.registration": (rows) => {
    const row = rows[0];
    if (!row) return null;
    const registered = flagOf(row, "registered");
    const strength = textOf(row, "strength");
    if (registered === false) return { sentence: "No MFA registered", tone: "bad" };
    if (registered === true)
      return {
        sentence: strength ? `MFA registered, ${strength}` : "MFA registered",
        tone: "good",
      };
    return strength ? { sentence: `Strongest method ${strength}` } : null;
  },
  "accounts.passwords": (rows, now) => {
    const row = rows[0];
    if (!row) return null;
    const changed = textOf(row, "lastChanged");
    const policy = textOf(row, "policy");
    const status = textOf(row, "status");
    const never = flagOf(row, "neverExpires") === true;
    const weak = flagOf(row, "weak") === true;
    const parts = [
      changed ? `Changed ${relativeTime(changed, now)}` : "Password change date unknown",
      policy === "Managed on-premises"
        ? "managed on-premises"
        : policy === "Never"
          ? "never expires"
          : policy && policy !== "Unknown"
            ? `expires after ${policy}`
            : null,
      never && policy !== "Never" ? "never expires for this account" : null,
      weak ? "strong password not required" : null,
      status && status !== "OK" ? lowerFirst(status) : null,
    ].filter(Boolean);
    const concern = never || weak || (status !== null && status !== "OK");
    return { sentence: parts.join(", "), ...(concern ? { tone: "warning" as const } : {}) };
  },
  "guests.review": (rows, now) => {
    const row = rows[0];
    if (!row) return null;
    const invitation = textOf(row, "invitation");
    const invited = textOf(row, "invited");
    const last = textOf(row, "lastSuccess");
    const days = numberOf(row, "days");
    if (invitation === "Not accepted")
      return {
        sentence: invited
          ? `Invitation not accepted, sent ${relativeTime(invited, now)}`
          : "Invitation not accepted",
        tone: "warning",
      };
    if (days !== null && days >= INACTIVE_FROM_DAYS)
      return { sentence: `Guest, inactive for ${plural(days, "day")}`, tone: "warning" };
    return { sentence: last ? `Guest, signed in ${relativeTime(last, now)}` : "Guest" };
  },
  "devices.managed": (rows) => {
    if (rows.length === 0) return null;
    const total = rows.length;
    const notCompliant = rows.filter((row) => textOf(row, "compliance") !== "Compliant").length;
    const unencrypted = rows.filter((row) => flagOf(row, "encrypted") === false).length;
    const lead = plural(total, "device");
    if (notCompliant > 0)
      return { sentence: `${notCompliant} of ${total} not compliant`, tone: "warning" };
    if (unencrypted > 0)
      return { sentence: `${lead}, ${unencrypted} not encrypted`, tone: "warning" };
    const all = total === 1 ? "compliant" : "all compliant";
    const encrypted = rows.every((row) => flagOf(row, "encrypted") === true);
    return { sentence: `${lead}, ${all}${encrypted ? " and encrypted" : ""}` };
  },
  "mail.mailbox-storage": (rows) => {
    const row = rows[0];
    if (!row) return null;
    const used = numberOf(row, "used");
    const quota = numberOf(row, "sendQuota");
    const percent = numberOf(row, "percent");
    const status = textOf(row, "status");
    if (used === null) return null;
    const size =
      quota === null
        ? `${gigabytes(used)} GB used`
        : `${gigabytes(used)} of ${gigabytes(quota)} GB used`;
    const concern = status !== null && status !== "OK";
    return {
      sentence: concern ? `${size}, ${lowerFirst(status)}` : size,
      ...(concern || (percent !== null && percent >= 90) ? { tone: "warning" as const } : {}),
    };
  },
  "usage.m365-apps": (rows, now) => {
    const row = rows[0];
    if (!row) return null;
    const last = textOf(row, "lastActivity");
    const apps = listOf(row, "apps");
    if (!last) return { sentence: "No app activity in 90 days", tone: "warning" };
    return {
      sentence: `Active ${relativeTime(last, now)}${apps.length > 0 ? ` in ${joinWords(apps)}` : ""}`,
    };
  },
};

// Account activity folds in inactivity and licences to reclaim, as the Access tab does.
function activityWording(
  byId: ReadonlyMap<string, SummaryAudit>,
  now: number,
): { said: Said | null; spoke: boolean } {
  const rowsOf = (id: string): readonly Row[] => {
    const audit = byId.get(id);
    return audit?.status === "ok" ? audit.rows : [];
  };
  const row = rowsOf("accounts.activity")[0];
  const inactive = rowsOf("accounts.inactive")[0];
  const reclaim = rowsOf("licenses.reclaim")[0];
  const days = inactive ? numberOf(inactive, "days") : null;
  const stale = days !== null && days >= INACTIVE_FROM_DAYS;
  const parts: string[] = [];
  let warn = false;
  if (row) {
    const verdict = textOf(row, "verdict");
    const last = textOf(row, "lastSuccess");
    if (verdict === "Active")
      parts.push(last ? `Active, signed in ${relativeTime(last, now)}` : "Active");
    else if (verdict)
      parts.push(last ? `${verdict}; last sign-in ${relativeTime(last, now)}` : verdict);
    warn = verdict !== null && verdict !== "Active" && verdict !== "Not enough data";
  }
  if (stale) {
    parts.push(`inactive for ${plural(days, "day")}`);
    warn = true;
  }
  if (reclaim) {
    parts.push("license to reclaim");
    warn = true;
  }
  if (parts.length === 0) return { said: null, spoke: false };
  const sentence = parts.join(", ");
  return {
    said: {
      sentence: sentence.charAt(0).toUpperCase() + sentence.slice(1),
      ...(warn ? { tone: "warning" } : {}),
    },
    spoke: true,
  };
}

// The audit sources, in the panel's order, with the column and Access tab section of each and
// what to show when the run has nothing about this person. "calm" says so as good news, "muted"
// says so quietly, and "omit" leaves the source out, because most people are not in it.
const auditSources = [
  {
    id: "identity.risky-users",
    source: "Risky users",
    column: "risk",
    target: sectionIds.risk,
    absent: "calm",
    calm: "Not flagged",
  },
  {
    id: "directory.privileged-roles",
    source: "Privileged roles",
    column: "risk",
    target: sectionIds.groups,
    absent: "omit",
  },
  {
    id: "accounts.activity",
    source: "Account activity",
    column: "account",
    target: sectionIds.activity,
    absent: "muted",
  },
  {
    id: "mfa.registration",
    source: "MFA registration",
    column: "account",
    target: sectionIds.methods,
    absent: "muted",
  },
  {
    id: "accounts.passwords",
    source: "Passwords",
    column: "account",
    target: sectionIds.password,
    absent: "muted",
  },
  {
    id: "guests.review",
    source: "Guest accounts",
    column: "account",
    target: sectionIds.guest,
    absent: "omit",
  },
  {
    id: "devices.managed",
    source: "Managed devices",
    column: "account",
    target: sectionIds.devices,
    absent: "omit",
  },
  {
    id: "mail.mailbox-storage",
    source: "Mailbox storage",
    column: "account",
    target: sectionIds.mailbox,
    absent: "omit",
  },
  {
    id: "usage.m365-apps",
    source: "Microsoft 365 apps",
    column: "account",
    target: sectionIds.apps,
    absent: "omit",
  },
] as const;

const access = (id: string): Pick<SummaryItem, "jump" | "destination"> => ({
  jump: { kind: "section", tab: "access", id },
  destination: "the Access tab",
});

const made = (
  base: Pick<SummaryItem, "id" | "source" | "column" | "jump" | "destination">,
  over: Partial<SummaryItem> & { sentence: string },
): SummaryItem => ({
  severity: null,
  tone: null,
  ranAt: null,
  incomplete: false,
  muted: false,
  ...base,
  ...over,
});

const failed = (
  base: Pick<SummaryItem, "id" | "source" | "column" | "jump" | "destination">,
  what: string,
) => made(base, { sentence: `${what} could not be read just now`, muted: true });

function auditItem(
  source: (typeof auditSources)[number],
  audit: SummaryAudit,
  byId: ReadonlyMap<string, SummaryAudit>,
  now: number,
): SummaryItem | null {
  const base = {
    id: source.id,
    source: source.source,
    column: source.column,
    ...access(source.target),
  };
  if (audit.status === "no_run") return null;
  if (audit.status === "failed") return failed(base, source.source);
  if (audit.status === "no_name")
    return made(base, {
      sentence: "Could not be matched, no sign-in name stored for this account",
      muted: true,
      ranAt: audit.finishedAt,
      incomplete: !audit.complete,
    });
  const run = { ranAt: audit.finishedAt, incomplete: !audit.complete };
  const activity = source.id === "accounts.activity" ? activityWording(byId, now) : null;
  if (activity?.said) {
    const { sentence, tone } = activity.said;
    return made(base, { sentence, ...(tone ? { tone } : {}), ...run });
  }
  const said = audit.status === "ok" ? wordings[source.id]?.(audit.rows, now) : null;
  if (said) {
    const rowsMore = audit.status === "ok" && audit.truncated;
    return made(base, {
      sentence: rowsMore ? `${said.sentence}, and more` : said.sentence,
      tone: said.tone ?? null,
      severity: said.severity ?? null,
      ...(said.severityLabel ? { severityLabel: said.severityLabel } : {}),
      ...run,
    });
  }
  if (source.absent === "omit") return null;
  if (source.absent === "calm") return made(base, { sentence: source.calm, tone: "good", ...run });
  return made(base, { sentence: "Not in the latest run", muted: true, ...run });
}

export function summaryItems(summary: PersonSummary, now: number = Date.now()): SummaryLines {
  const items: SummaryItem[] = [];
  const notRun: string[] = [];
  const byId = new Map<string, SummaryAudit>(summary.audits.map((audit) => [audit.auditId, audit]));

  // Combined risks.
  const combinedBase = {
    id: "combined-risks",
    source: "Combined risks",
    column: "risk" as const,
    ...access(sectionIds.combinedRisks),
  };
  const combined = summary.combinedRisks;
  if (combined.status === "no_run") notRun.push("Combined risks");
  else if (combined.status === "failed") items.push(failed(combinedBase, "Combined risks"));
  else if (combined.severity === null)
    items.push(
      made(combinedBase, {
        sentence: "No combined risks",
        tone: "good",
        ranAt: combined.finishedAt,
        incomplete: !combined.complete,
      }),
    );
  else {
    const first = combined.rules[0]?.label ?? "Combined risk";
    const more = combined.rules.length - 1;
    items.push(
      made(combinedBase, {
        severity: combined.severity,
        sentence: more > 0 ? `${first} and ${plural(more, "more rule")}` : first,
        ranAt: combined.finishedAt,
        incomplete: !combined.complete,
      }),
    );
  }

  // Sign-in findings.
  const findingsBase = {
    id: "signin-findings",
    source: "Sign-in findings",
    column: "risk" as const,
    jump: { kind: "section", tab: "overview", id: sectionIds.findings } as Jump,
    destination: "the findings below",
  };
  const findings = summary.signinFindings;
  if (findings.status === "failed") items.push(failed(findingsBase, "Sign-in findings"));
  else if (findings.status === "ok")
    items.push(
      findings.open === 0
        ? made(findingsBase, { sentence: "No open findings", tone: "good" })
        : made(findingsBase, {
            severity: findings.highest,
            sentence: findings.more
              ? `${number.format(findings.open)} or more open findings`
              : plural(findings.open, "open finding"),
          }),
    );

  const place = (id: string, column: "risk" | "account") => {
    const source = auditSources.find((candidate) => candidate.id === id);
    const audit = byId.get(id);
    if (!source || !audit || source.column !== column) return;
    if (audit.status === "no_run") {
      // Guest accounts matter only to guests, so never having run it is not a gap worth naming.
      if (id !== "guests.review") notRun.push(source.source);
      return;
    }
    const built = auditItem(source, audit, byId, now);
    if (built) items.push(built);
  };
  place("identity.risky-users", "risk");
  place("directory.privileged-roles", "risk");

  // Groups they own.
  const ownsBase = {
    id: "owned-groups",
    source: "Groups they own",
    column: "risk" as const,
    ...access(sectionIds.owns),
  };
  const owned = summary.owned;
  if (owned.status === "no_run") notRun.push("Groups and Teams");
  else if (owned.status === "failed") items.push(failed(ownsBase, "Groups and Teams"));
  else if (owned.entries.length > 0) {
    const concerns = owned.entries.filter(
      (entry) => entry.status !== "OK" && entry.status !== "" && entry.status !== "No usage data",
    ).length;
    const lead = `Owns ${owned.more ? "at least " : ""}${plural(owned.entries.length, "group or team", "groups and teams")}`;
    items.push(
      made(ownsBase, {
        sentence:
          concerns > 0
            ? `${lead}, ${concerns} ${concerns === 1 ? "needs" : "need"} attention`
            : lead,
        ...(concerns > 0 ? { tone: "warning" as const } : {}),
        ranAt: owned.finishedAt,
        incomplete: !owned.complete,
      }),
    );
  }

  // Shared mailboxes.
  const mailboxBase = {
    id: "shared-mailboxes",
    source: "Shared mailboxes",
    column: "risk" as const,
    ...access(sectionIds.sharedMailboxes),
  };
  const shared = summary.sharedMailboxes;
  if (shared.status === "failed") items.push(failed(mailboxBase, "Shared mailbox access"));
  else if (shared.status === "no_import")
    items.push(
      made(
        {
          ...mailboxBase,
          jump: { kind: "page", href: "#/access" },
          destination: "Microsoft access",
        },
        { sentence: "No mailbox access export imported", muted: true },
      ),
    );
  else {
    const direct = shared.mailboxes;
    const sentence =
      direct > 0
        ? `Can use ${shared.mailboxesMore ? "at least " : ""}${plural(direct, "shared mailbox", "shared mailboxes")}`
        : shared.holders !== null
          ? `${plural(shared.holders, "account")} can use their mailbox`
          : "No direct shared mailbox rights";
    items.push(
      made(mailboxBase, {
        sentence,
        ranAt: shared.exportedAt,
        muted: direct === 0 && shared.holders === null,
      }),
    );
  }

  // The account itself.
  for (const id of [
    "accounts.activity",
    "mfa.registration",
    "accounts.passwords",
    "guests.review",
    "devices.managed",
    "mail.mailbox-storage",
    "usage.m365-apps",
  ])
    place(id, "account");

  const changesBase = {
    id: "directory-changes",
    source: "Directory changes",
    column: "account" as const,
    jump: { kind: "section", tab: "changes", id: sectionIds.changes } as Jump,
    destination: "the Changes tab",
  };
  const changes = summary.changes;
  if (changes.status === "failed") items.push(failed(changesBase, "Directory changes"));
  else if (changes.status === "ok")
    items.push(
      changes.count === 0
        ? made(changesBase, {
            sentence: `No changes in the last ${changes.days} days`,
            muted: true,
          })
        : made(changesBase, {
            sentence: `${changes.more ? `${number.format(changes.count)} or more changes` : plural(changes.count, "change")} in the last ${changes.days} days`,
          }),
    );

  const handoffBase = {
    id: "handoff",
    source: "Handoff",
    column: "account" as const,
    jump: { kind: "section", tab: "changes", id: sectionIds.receipts } as Jump,
    destination: "the Changes tab",
  };
  const handoff = summary.handoff;
  if (handoff.status === "failed") items.push(failed(handoffBase, "Handoffs"));
  else if (handoff.latest) {
    const { state, action, performedAt } = handoff.latest;
    items.push(
      made(handoffBase, {
        sentence: `${actionLabel(action)}, ${state} ${relativeTime(performedAt, now)}`,
        ...(state === "failed"
          ? { tone: "bad" as const }
          : state === "skipped"
            ? { tone: "warning" as const }
            : {}),
      }),
    );
  }
  return { items, notRun };
}

// What a screen reader hears for a row: the source, the rating, the sentence and where the row
// goes. The visible row is shortened with an ellipsis, so this is the full text.
export function accessibleName(entry: SummaryItem): string {
  const rating = entry.severity ? (entry.severityLabel ?? severityNames[entry.severity]) : null;
  const lead = rating ? `${entry.source}, ${rating}: ` : `${entry.source}: `;
  const run = entry.incomplete ? " That run could not check everything." : "";
  return `${lead}${entry.sentence}.${run} Opens ${entry.destination}.`;
}

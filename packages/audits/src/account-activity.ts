import type { AuditDefinitionInput, AuditRow, AuditSummaryItem } from "@prism365/contracts";
import { CSV_MAX_ROWS } from "@prism365/graph";
import {
  MAX_MEMBERSHIP_ACCOUNTS,
  type Membership,
  type MembershipRead,
  readMemberships,
} from "./activity-groups.ts";
import {
  type AccountFacts,
  type ReportId,
  type ReportOutcome,
  type ReportsRead,
  readReports,
} from "./activity-reports.ts";
import type { Audit, AuditContext } from "./definition.ts";
import {
  type Activity,
  activityOf,
  byName,
  coverageNotes,
  type DirectoryUser,
  displayDate,
  licenseNamesNote,
  licenseSourcesOf,
  licensesOf,
  personCell,
  readLicenseNames,
  readUsers,
  signInNote,
  statusOf,
  typeOf,
  unreadableNote,
} from "./directory.ts";
import { lastActivityOf, PERIOD_DAYS, readSetting, type Setting } from "./m365-app-usage.ts";
import {
  domainColumn,
  domainInventory,
  excludedNote,
  inAuditScope,
  principalDomain,
  scopeNote,
} from "./scope.ts";

// Whether an account is really unused and whether anyone depends on it: sign-ins, the four
// workloads' use, the account's OneDrive, mailbox size, licenses and direct group memberships,
// with one verdict per account. Prism365 reports what Microsoft recorded. It does not measure
// anyone's work: a sign-in or an opened file says an account was used, not how well or why.

export const VERDICTS = [
  "Active",
  "Signs in, uses nothing",
  "Quiet, others use their files",
  "Quiet everywhere",
  "Recently quiet",
  "Never signed in",
  "Not enough data",
] as const;
export type Verdict = (typeof VERDICTS)[number];

// A successful sign-in or workload activity inside this many days makes an account active.
export const ACTIVE_DAYS = 30;
// No sign-in for this many days is quiet. It is also the usage reports' period, so activity
// inside the period is activity in the last 90 days.
export const QUIET_DAYS = 90;
const GIB = 1_073_741_824;
export const LARGE_MAILBOX_BYTES = 10 * GIB;

const verdictTags: Record<Verdict, string> = {
  Active: "active",
  "Signs in, uses nothing": "uses-nothing",
  "Quiet, others use their files": "others-use-files",
  "Quiet everywhere": "quiet",
  "Recently quiet": "recently-quiet",
  "Never signed in": "never",
  "Not enough data": "no-data",
};

// What the reports that were read let the audit conclude.
export interface Availability {
  // Every workload (mail, Teams, OneDrive, SharePoint) has at least one report that was read.
  workloads: boolean;
  // The OneDrive usage report was read, so the account's own OneDrive could be checked.
  drive: boolean;
}

export function availabilityOf(read: ReadonlySet<ReportId>): Availability {
  const has = (...ids: ReportId[]) => ids.some((id) => read.has(id));
  return {
    workloads:
      has("email", "mailbox", "activeUsers") &&
      has("teams", "activeUsers") &&
      has("oneDriveActivity", "activeUsers") &&
      has("sharePoint", "activeUsers"),
    drive: has("oneDriveUsage"),
  };
}

export interface Evidence {
  // Days since the latest use of the account the reports date; null when they date none.
  latestDays: number | null;
  // Any use inside the period: a dated one, or a count above zero.
  inPeriod: boolean;
  // Whether the reports can say there was no use: every workload and the OneDrive account were
  // read, and no date on this account was unreadable.
  decidable: boolean;
  // The account's OneDrive shows file activity or active files in the period.
  sharedFiles: boolean;
}

// Use of the account by its owner, as the reports date and count it. Mail an account only
// receives is left out: someone else did that. A date that cannot be read is neither use nor
// its absence, so it makes the account undecidable rather than counting for either.
export function evidenceOf(
  facts: AccountFacts | undefined,
  availability: Availability,
  now: Date,
): Evidence {
  let latest: number | null = null;
  let inPeriod = false;
  let unreadable = false;
  const dates = facts
    ? [
        facts.activeExchange,
        facts.activeTeams,
        facts.activeOneDrive,
        facts.activeSharePoint,
        facts.mailLast,
        facts.teamsLast,
        facts.oneDriveLast,
        facts.sharePointLast,
        facts.mailboxLast,
      ]
    : [];
  for (const value of dates) {
    if (!value) continue;
    const last = lastActivityOf(value, now);
    if (last.state === "unreadable") unreadable = true;
    if (last.state !== "known" || last.days === null) continue;
    latest = latest === null ? last.days : Math.min(latest, last.days);
    if (last.days < PERIOD_DAYS) inPeriod = true;
  }
  const counts = facts
    ? [
        facts.mailSent,
        facts.mailRead,
        facts.teamsChats,
        facts.teamsCalls,
        facts.teamsMeetings,
        facts.oneDriveFiles,
        facts.sharePointFiles,
        facts.sharePointPages,
      ]
    : [];
  if (counts.some((count) => (count ?? 0) > 0) || facts?.teamsOther === true) inPeriod = true;
  const drive = lastActivityOf(facts?.driveLast ?? null, now);
  if (drive.state === "unreadable") unreadable = true;
  const sharedFiles =
    (drive.state === "known" && (drive.days ?? PERIOD_DAYS) < PERIOD_DAYS) ||
    (facts?.driveActive ?? 0) > 0;
  return {
    latestDays: latest,
    inPeriod,
    decidable: availability.workloads && availability.drive && !unreadable,
    sharedFiles,
  };
}

// The verdict, decided by the first rule that fits. Positive evidence (a recent sign-in or use)
// is enough on its own; every verdict that says something was not there needs the reports to
// have been read.
export function verdictOf(activity: Activity, evidence: Evidence): Verdict {
  const days = activity.days ?? 0;
  const signedIn = activity.basis === "success" && days < ACTIVE_DAYS;
  if (signedIn && evidence.decidable && !evidence.inPeriod) return "Signs in, uses nothing";
  if (signedIn || (evidence.latestDays !== null && evidence.latestDays < ACTIVE_DAYS))
    return "Active";
  const unused = evidence.decidable && !evidence.inPeriod;
  if (activity.basis === "none") {
    if (unused && evidence.sharedFiles) return "Quiet, others use their files";
    // Sign-in records alone decide this one. A new account has simply not been used yet.
    return days >= ACTIVE_DAYS ? "Never signed in" : "Not enough data";
  }
  if (activity.basis === "unreadable" || !evidence.decidable) return "Not enough data";
  // Attempts with no success on record, none of them older than the quiet period: the account is
  // not quiet, and not signing in either.
  if (activity.basis === "attempt" && days < QUIET_DAYS) return "Not enough data";
  if (days >= QUIET_DAYS && unused)
    return evidence.sharedFiles ? "Quiet, others use their files" : "Quiet everywhere";
  return "Recently quiet";
}

const gigabytes = (bytes: number | null | undefined) =>
  bytes === null || bytes === undefined ? null : Math.round((bytes / GIB) * 100) / 100;

// The later of the dates, as an ISO timestamp, or null when none is readable.
function latestDate(values: readonly (string | null | undefined)[], now: Date): string | null {
  let best: string | null = null;
  for (const value of values) {
    if (!value) continue;
    const last = lastActivityOf(value, now);
    if (last.state === "known" && last.at !== null && (best === null || last.at > best))
      best = last.at;
  }
  return best;
}

const plural = (count: number, one: string, many: string) =>
  `${count.toLocaleString("en-US")} ${count === 1 ? one : many}`;

function listOf(names: readonly string[]): string {
  if (names.length <= 1) return names.join("");
  return `${names.slice(0, -1).join(", ")} and ${names.at(-1)}`;
}

const NOT_READ = "Not read";
const NO_DATA = "No data";

type Column = AuditDefinitionInput["columns"][number];

const date = (key: string, label: string, help: string | null, hidden = false): Column => ({
  key,
  label,
  kind: "date",
  hidden,
  help,
  empty: NO_DATA,
});
const count = (key: string, label: string, help: string | null): Column => ({
  key,
  label,
  kind: "number",
  hidden: true,
  help,
  empty: NO_DATA,
});

const verdictHelp =
  "Active: signed in, or used mail, Teams, OneDrive or SharePoint, in the last 30 days. Signs in, uses nothing: signed in in the last 30 days, no use in 90. Quiet everywhere: no sign-in or use for 90 days. Quiet, others use their files: the same, yet its OneDrive shows file activity. Recently quiet: last seen 30 to 89 days ago. Never signed in: no sign-in on record, account over 30 days old. Not enough data: the reports or dates cannot settle it.";

const columns: Column[] = [
  {
    key: "person",
    label: "Account",
    kind: "person",
  },
  {
    key: "verdict",
    label: "Verdict",
    kind: "enum",
    facet: true,
    help: verdictHelp,
  },
  {
    key: "status",
    label: "Status",
    kind: "enum",
    facet: true,
  },
  { key: "type", label: "Type", kind: "enum", facet: true, hidden: true },
  domainColumn,
  {
    key: "lastSuccess",
    label: "Last successful sign-in",
    kind: "date",
    help: "Interactive or background. Microsoft has recorded this since 1 December 2023.",
    empty: "None recorded",
  },
  {
    key: "days",
    label: "Since sign-in",
    kind: "days",
    help: "Days since the last successful sign-in, counted as in the Inactive accounts audit: with no success on record, since the last attempt; with no sign-in at all, since the account was created.",
    empty: "Unknown",
  },
  {
    key: "lastInteractive",
    label: "Last interactive attempt",
    kind: "date",
    hidden: true,
    help: "The last time the person signed in themselves, successfully or not.",
    empty: "None recorded",
  },
  {
    key: "lastNonInteractive",
    label: "Last background attempt",
    kind: "date",
    hidden: true,
    help: "The last time an app signed in for the person, successfully or not.",
    empty: "None recorded",
  },
  date(
    "mailLast",
    "Mail last activity",
    "The later of the dates the active users report and the email report give. Microsoft gives a date, not a time. No data means the account has no row in the reports, or they were not read.",
  ),
  count("mailSent", "Mail sent", "Messages sent in the last 90 days."),
  count(
    "mailReceived",
    "Mail received",
    "Messages received in the last 90 days. Shown, but not counted as use of the account.",
  ),
  count("mailRead", "Mail read", "Messages read in the last 90 days."),
  date(
    "teamsLast",
    "Teams last activity",
    "The later of the dates the active users report and the Teams report give.",
  ),
  count(
    "teamsChats",
    "Teams chats",
    "Team chat and private chat messages sent in the last 90 days.",
  ),
  count("teamsCalls", "Teams calls", "Calls in the last 90 days."),
  count("teamsMeetings", "Teams meetings attended", "Meetings attended in the last 90 days."),
  date(
    "oneDriveLast",
    "OneDrive last activity",
    "The later of the dates the active users report and the OneDrive activity report give: what the account's owner did.",
  ),
  count(
    "oneDriveFiles",
    "OneDrive files viewed or edited",
    "Files the account's owner viewed or edited in the last 90 days.",
  ),
  date(
    "oneDriveAccountLast",
    "OneDrive account last file activity",
    "From the OneDrive usage report, which describes the account's OneDrive rather than what its owner did. Microsoft does not say whose activity it counts.",
    true,
  ),
  count(
    "oneDriveActiveFiles",
    "OneDrive active files",
    "Files in the account's OneDrive that were active in the last 90 days. Activity here on a quiet account means someone still opens its files; Microsoft does not say who.",
  ),
  count("oneDriveFileCount", "OneDrive files", "Files in the account's OneDrive."),
  count(
    "oneDriveStorage",
    "OneDrive storage used (GB)",
    "Gigabytes used. A gigabyte here is 1,073,741,824 bytes.",
  ),
  date(
    "sharePointLast",
    "SharePoint last activity",
    "The later of the dates the active users report and the SharePoint activity report give.",
  ),
  count(
    "sharePointFiles",
    "SharePoint files viewed or edited",
    "Files viewed or edited in the last 90 days.",
  ),
  count("sharePointPages", "SharePoint pages visited", "Pages visited in the last 90 days."),
  count(
    "mailboxStorage",
    "Mailbox storage used (GB)",
    "Gigabytes used. A gigabyte here is 1,073,741,824 bytes.",
  ),
  count("mailboxItems", "Mailbox items", "Items in the mailbox."),
  date(
    "mailboxLast",
    "Mailbox last activity",
    "From the mailbox usage report. It counts as use of the account.",
    true,
  ),
  {
    key: "mailboxArchive",
    label: "Mailbox has archive",
    kind: "boolean",
    hidden: true,
    help: "Some of Microsoft's reports leave this out.",
    empty: NO_DATA,
  },
  {
    key: "licenses",
    label: "Licenses",
    kind: "list",
    hidden: true,
  },
  {
    key: "licenseAssignment",
    label: "License assignment",
    kind: "enum",
    facet: true,
    help: "Direct, Through groups, Both or None. A license assigned through a group comes back if you remove it from the account: remove the account from the group instead.",
  },
  {
    key: "licenseGroups",
    label: "License groups",
    kind: "list",
    hidden: true,
    help: "The groups that assign a license the account holds.",
    empty: NOT_READ,
  },
  {
    key: "groups",
    label: "Groups",
    kind: "number",
    help: "Direct memberships only: nested groups are not followed. Directory roles and administrative units are counted apart.",
    empty: NOT_READ,
  },
  {
    key: "groupNames",
    label: "Group names",
    kind: "list",
    hidden: true,
    empty: NOT_READ,
  },
  {
    key: "roles",
    label: "Directory roles",
    kind: "number",
    hidden: true,
    help: "Directory roles the account holds directly, such as Global Administrator.",
    empty: NOT_READ,
  },
  {
    key: "roleNames",
    label: "Role names",
    kind: "list",
    hidden: true,
    empty: NOT_READ,
  },
  {
    key: "adminUnits",
    label: "Administrative units",
    kind: "number",
    hidden: true,
    empty: NOT_READ,
  },
];

// What the usage reports were, for the notes and the verdicts.
type Usage =
  | { state: "not_requested" }
  | { state: "concealed"; setting: Setting }
  | { state: "read"; setting: Setting; reports: ReportsRead };

async function readUsage(context: AuditContext, users: readonly DirectoryUser[]): Promise<Usage> {
  if (!context.granted.has("Reports.Read.All")) return { state: "not_requested" };
  const setting = await readSetting(context);
  if (setting.concealed === true) return { state: "concealed", setting };
  const reports = await readReports(context, users);
  return reports.concealed ? { state: "concealed", setting } : { state: "read", setting, reports };
}

const concealedAdvice =
  "A Global Administrator can turn concealment off in the Microsoft 365 admin center (Settings, Org settings, Services, Reports). Prism365 will not unmask anything itself.";

function usageNotes(usage: Usage): string[] {
  if (usage.state === "not_requested")
    return [
      "Usage reports were not read: this run did not have Reports.Read.All (the usage reports read pack is off, or the sign-in has not granted it). The workload columns show No data, and every verdict that needs them is Not enough data.",
    ];
  if (usage.state === "concealed")
    return [
      `${
        usage.setting.concealed === true
          ? "Microsoft's report privacy setting says names are concealed."
          : usage.setting.concealed === false
            ? "Microsoft's privacy setting says names are shown, but the reports still return hashed identifiers (a change can take a few minutes to reach them)."
            : "Prism365 could not read the report privacy setting, and the reports' identifiers look like hashes, so they are treated as concealed."
      } Prism365 did not join the reports to accounts: the sign-in, license and group columns are shown, the workload columns show No data, and every verdict that needs them is Not enough data.`,
      concealedAdvice,
    ];
  const { setting, reports } = usage;
  const notes: string[] = [];
  if (setting.concealed === null) {
    const reason =
      setting.why === "not_granted"
        ? "Prism365 did not read the report privacy setting (it needs ReportSettings.Read.All)"
        : setting.why === "refused"
          ? "Microsoft refused the report privacy setting"
          : "Microsoft's answer about the report privacy setting could not be read";
    notes.push(`${reason}. Names are shown because the reports returned readable sign-in names.`);
  }
  const failed = new Map<string, string[]>();
  for (const outcome of reports.outcomes) {
    if (outcome.state !== "failed" || !outcome.reason) continue;
    failed.set(outcome.reason, [...(failed.get(outcome.reason) ?? []), outcome.name]);
  }
  for (const [reason, names] of failed)
    notes.push(
      `The ${listOf(names)} ${names.length === 1 ? "report" : "reports"} could not be read: ${reason}. ${names.length === 1 ? "Its columns show" : "Their columns show"} No data, and a verdict that needs ${names.length === 1 ? "it" : "them"} is Not enough data.`,
    );
  if (reports.truncated.length > 0)
    notes.push(
      `The ${listOf(reports.truncated)} ${reports.truncated.length === 1 ? "report has" : "reports have"} more than ${CSV_MAX_ROWS.toLocaleString("en-US")} rows; the rest was not read, so an account may show less use than it had.`,
    );
  if (reports.skippedRows > 0)
    notes.push(
      `${plural(reports.skippedRows, "row in the reports", "rows in the reports")} could not be read and ${reports.skippedRows === 1 ? "is" : "are"} left out, so an account may show less use than it had.`,
    );
  if (reports.unreadableValues > 0)
    notes.push(
      `${plural(reports.unreadableValues, "value in the reports", "values in the reports")} could not be read as a number or a yes or no and ${reports.unreadableValues === 1 ? "shows" : "show"} as No data.`,
    );
  if (reports.unmatched > 0)
    notes.push(
      `${plural(reports.unmatched, "name in the reports matches", "names in the reports match")} no account in the directory (deleted or renamed accounts, for example) and ${reports.unmatched === 1 ? "is" : "are"} not shown.`,
    );
  return notes;
}

function membershipNotes(read: MembershipRead, asked: number): string[] {
  const notes: string[] = [];
  if (read.notAsked > 0)
    notes.push(
      `Group memberships were read for ${MAX_MEMBERSHIP_ACCOUNTS.toLocaleString("en-US")} of ${asked.toLocaleString("en-US")} accounts, quiet ones first. ${plural(read.notAsked, "account shows", "accounts show")} Not read for groups and directory roles.`,
    );
  if (read.stopped === "throttled")
    notes.push(
      `Microsoft asked Prism365 to slow down, so group memberships stopped after ${plural(read.byAccount.size, "account", "accounts")}. Run the audit again in a few minutes.`,
    );
  else if (read.stopped === "refused")
    notes.push(
      `Microsoft refused to list group memberships (HTTP ${read.status}), so ${read.byAccount.size === 0 ? "none were read" : `they were read for ${plural(read.byAccount.size, "account", "accounts")} only`}. Listing another user's memberships needs User.Read.All.`,
    );
  else if (read.unanswered > 0)
    notes.push(
      `${plural(read.unanswered, "account", "accounts")} had no answer for group memberships and ${read.unanswered === 1 ? "shows" : "show"} Not read.`,
    );
  if (read.partial > 0)
    notes.push(
      `${plural(read.partial, "account has", "accounts have")} memberships Prism365 could not read or more than it reads for one account, so the group and role counts are low.`,
    );
  return notes;
}

interface Entry {
  user: DirectoryUser;
  activity: Activity;
  facts: AccountFacts | undefined;
  verdict: Verdict;
  covered: boolean;
}

// The order group memberships are asked for in when there are more accounts than the limit
// allows: quiet accounts in scope, quiet accounts outside it (the critical tile counts them),
// then everyone else, least active first.
const askOrder: Verdict[] = [
  "Never signed in",
  "Recently quiet",
  "Not enough data",
  "Signs in, uses nothing",
  "Active",
];

function membershipOrder(entries: readonly Entry[]): DirectoryUser[] {
  const quiet = (entry: Entry) =>
    entry.verdict === "Quiet everywhere" || entry.verdict === "Quiet, others use their files";
  return [
    ...entries.filter((entry) => entry.covered && quiet(entry)),
    ...entries.filter((entry) => !entry.covered && entry.verdict === "Quiet everywhere"),
    ...askOrder.flatMap((verdict) =>
      entries.filter((entry) => entry.covered && entry.verdict === verdict),
    ),
  ].map((entry) => entry.user);
}

export const accountActivityAudit: Audit = {
  definition: {
    id: "accounts.activity",
    // Accounts an action tool could act on (docs/handoff.md).
    handoff: true,
    version: 1,
    title: "Account activity",
    description:
      "Is this account really unused, and does anyone depend on it? One row per account: sign-ins, mail, Teams, OneDrive and SharePoint use, mailbox size, licenses and groups, with a verdict. The list to check before offboarding or reclaiming licenses.",
    category: "Accounts",
    subject: { one: "account", many: "accounts" },
    columns,
    views: [
      { id: "quiet", label: "Quiet everywhere", tags: ["quiet"] },
      {
        id: "others-use-files",
        label: "Quiet, others use their files",
        tags: ["others-use-files"],
      },
      { id: "uses-nothing", label: "Signs in, uses nothing", tags: ["uses-nothing"] },
      { id: "recently-quiet", label: "Recently quiet", tags: ["recently-quiet"] },
      { id: "never", label: "Never signed in", tags: ["never"] },
      { id: "active", label: "Active", tags: ["active"] },
      { id: "no-data", label: "Not enough data", tags: ["no-data"] },
      { id: "group-only", label: "Licensed through groups only", tags: ["group-only"] },
      { id: "large-mailbox", label: "Large mailbox, quiet", tags: ["large-mailbox"] },
      { id: "roles", label: "Has directory roles", tags: ["roles"] },
      {
        id: "quiet-roles",
        label: "Quiet everywhere, with directory roles",
        tags: ["quiet", "roles"],
      },
      { id: "all", label: "All accounts", tags: [] },
    ],
    defaultView: "all",
    defaultSort: { column: "days", direction: "desc" },
    requirements: [
      "Microsoft Entra ID P1 or P2 in the tenant, which Microsoft requires for sign-in dates.",
      "A role that can read sign-in activity, such as Reports Reader, Security Reader or Global Reader.",
      "The workload columns need Reports Reader or a service administrator role (Exchange, SharePoint or Teams). Global Reader sees tenant totals only.",
      "Names show in the usage reports only when the tenant's report privacy setting shows them. A Global Administrator controls that setting.",
    ],
  },
  packs: ["directory.users.read", "audit.signins.read"],
  optionalPacks: ["reports.usage.read", "reports.settings.read", "tenant.licenses.read"],
  scoped: true,
  forbidden:
    "Microsoft refused to list accounts or their last sign-in dates. Listing accounts needs User.Read.All. Sign-in dates need AuditLog.Read.All, Microsoft Entra ID P1 or P2, and a role such as Reports Reader, Security Reader or Global Reader.",
  async run(context) {
    const read = await readUsers(context, { signInActivity: "required", groupAssignment: true });
    const names = await readLicenseNames(context);
    const usage = await readUsage(context, read.users);
    const reports = usage.state === "read" ? usage.reports : null;
    const availability = availabilityOf(reports?.readIds ?? new Set());

    // Every account gets a verdict, in scope or not: the tile for quiet accounts with directory
    // roles counts the ones outside the scope too.
    const entries: Entry[] = read.users.map((user) => {
      const activity = activityOf(user, context.now);
      const facts = reports?.facts.get(user.id);
      const evidence = evidenceOf(facts, availability, context.now);
      return {
        user,
        activity,
        facts,
        verdict: verdictOf(activity, evidence),
        covered: inAuditScope(context.scope, user.userPrincipalName),
      };
    });
    const order = membershipOrder(entries);
    const memberships = await readMemberships(context, order);

    const rows: AuditRow[] = [];
    let outsideQuietRoles = 0;
    for (const entry of entries) {
      const { user, activity, facts, verdict } = entry;
      const member: Membership | undefined = memberships.byAccount.get(user.id);
      if (!entry.covered) {
        if (verdict === "Quiet everywhere" && (member?.roles.length ?? 0) > 0)
          outsideQuietRoles += 1;
        continue;
      }
      const sources = licenseSourcesOf(user);
      const held = user.assignedLicenses.length > 0;
      const assignment = !held
        ? "None"
        : !sources.known || (!sources.direct && sources.groupIds.length === 0)
          ? "Unknown"
          : sources.direct && sources.groupIds.length > 0
            ? "Both"
            : sources.direct
              ? "Direct"
              : "Through groups";
      const groupNames = new Map(member?.groups.map((group) => [group.id, group.name]));
      const licenseGroups = member
        ? sources.groupIds.map((id) => groupNames.get(id) ?? "Another group")
        : null;
      const roles = member?.roles ?? [];
      const mailboxBytes = facts?.mailboxBytes ?? null;
      const largeMailbox =
        verdict === "Quiet everywhere" && (mailboxBytes ?? 0) > LARGE_MAILBOX_BYTES;
      const sign = user.signInActivity;
      rows.push({
        key: user.id,
        personId: user.id,
        tags: [
          verdictTags[verdict],
          ...(assignment === "Through groups" ? ["group-only"] : []),
          ...(largeMailbox ? ["large-mailbox"] : []),
          ...(roles.length > 0 ? ["roles"] : []),
        ],
        cells: {
          person: personCell(user),
          verdict,
          status: statusOf(user),
          type: typeOf(user.userType),
          domain: principalDomain(user.userPrincipalName),
          lastSuccess: displayDate(sign?.lastSuccessfulSignInDateTime ?? null),
          days: activity.days,
          lastInteractive: displayDate(sign?.lastSignInDateTime ?? null),
          lastNonInteractive: displayDate(sign?.lastNonInteractiveSignInDateTime ?? null),
          mailLast: latestDate([facts?.activeExchange, facts?.mailLast], context.now),
          mailSent: facts?.mailSent ?? null,
          mailReceived: facts?.mailReceived ?? null,
          mailRead: facts?.mailRead ?? null,
          teamsLast: latestDate([facts?.activeTeams, facts?.teamsLast], context.now),
          teamsChats: facts?.teamsChats ?? null,
          teamsCalls: facts?.teamsCalls ?? null,
          teamsMeetings: facts?.teamsMeetings ?? null,
          oneDriveLast: latestDate([facts?.activeOneDrive, facts?.oneDriveLast], context.now),
          oneDriveFiles: facts?.oneDriveFiles ?? null,
          oneDriveAccountLast: latestDate([facts?.driveLast], context.now),
          oneDriveActiveFiles: facts?.driveActive ?? null,
          oneDriveFileCount: facts?.driveFiles ?? null,
          oneDriveStorage: gigabytes(facts?.driveBytes),
          sharePointLast: latestDate([facts?.activeSharePoint, facts?.sharePointLast], context.now),
          sharePointFiles: facts?.sharePointFiles ?? null,
          sharePointPages: facts?.sharePointPages ?? null,
          mailboxStorage: gigabytes(mailboxBytes),
          mailboxItems: facts?.mailboxItems ?? null,
          mailboxLast: latestDate([facts?.mailboxLast], context.now),
          mailboxArchive: facts?.mailboxArchive ?? null,
          licenses: licensesOf(user, names),
          licenseAssignment: assignment,
          licenseGroups,
          groups: member ? member.groups.length : null,
          groupNames: member ? member.groups.map((group) => group.name) : null,
          roles: member ? roles.length : null,
          roleNames: member ? roles : null,
          adminUnits: member ? member.units : null,
        },
      });
    }
    rows.sort(byName);
    const tagged = (...tags: string[]) =>
      rows.filter((row) => tags.every((tag) => row.tags.includes(tag))).length;
    const unreadableDates = rows.filter((row) => row.cells.days === null).length;
    const quiet = tagged("quiet");
    const quietRoles = tagged("quiet", "roles");
    const othersUse = tagged("others-use-files");
    const usesNothing = tagged("uses-nothing");
    const groupOnly = tagged("group-only");

    const summary: AuditSummaryItem[] = [
      { label: "Accounts", value: rows.length, tone: "neutral", view: "all" },
      {
        label: "Quiet everywhere, with directory roles",
        value: quietRoles,
        // Never "good": a zero here also happens when the reports or the memberships were not read.
        tone: quietRoles > 0 ? "critical" : "neutral",
        view: "quiet-roles",
      },
      ...(outsideQuietRoles > 0
        ? [
            {
              label: "Quiet everywhere, with directory roles, outside the scope",
              value: outsideQuietRoles,
              tone: "critical" as const,
              view: null,
            },
          ]
        : []),
      {
        label: "Quiet everywhere",
        value: quiet,
        tone: quiet > 0 ? "warning" : "neutral",
        view: "quiet",
      },
      {
        label: "Quiet, others use their files",
        value: othersUse,
        tone: othersUse > 0 ? "warning" : "neutral",
        view: "others-use-files",
      },
      {
        label: "Signs in, uses nothing",
        value: usesNothing,
        tone: "neutral",
        view: "uses-nothing",
      },
      {
        label: "Licensed through groups only",
        value: groupOnly,
        tone: "neutral",
        view: "group-only",
      },
    ];

    const limited = scopeNote(context.scope);
    const left = excludedNote(read.users.length - rows.length);
    const refreshed = reports?.refreshed;
    // A result keeps 20 notes at most. The notes only some runs have come first, so when a run has
    // too many it is the standing ones at the end that go.
    const notes = [
      ...(limited ? [limited] : []),
      ...(left ? [left] : []),
      ...(outsideQuietRoles > 0
        ? [
            `${plural(outsideQuietRoles, "account", "accounts")} quiet everywhere ${outsideQuietRoles === 1 ? "holds a directory role" : "hold directory roles"} in domains the audit scope leaves out. Choose All domains and run again to see ${outsideQuietRoles === 1 ? "it" : "them"}.`,
          ]
        : []),
      ...usageNotes(usage),
      ...membershipNotes(memberships, order.length),
      ...(read.groupAssignment
        ? []
        : ["How licenses were assigned could not be read, so License assignment shows Unknown."]),
      ...signInNote(read),
      ...unreadableNote(unreadableDates),
      ...(names ? [] : [licenseNamesNote]),
      ...coverageNotes(read),
      `Microsoft's usage reports lag by 24 to 72 hours and sometimes by several days${refreshed ? `; the oldest report used was refreshed on ${refreshed}` : ""}. The last few days of use can be missing.`,
      "Per-user report detail needs Reports Reader or a service administrator role such as Exchange, SharePoint or Teams administrator. Global Reader sees tenant totals only.",
      "Use means the account sent or read mail, chatted, called, attended a meeting, or viewed or edited a file. Mail an account only receives is shown but not counted, and Microsoft does not say whether receiving mail moves its last activity date. None of this measures anyone's work.",
      "Quiet, others use their files: the account itself is quiet but its OneDrive shows file activity or active files. Microsoft does not say who opened them, so check before removing it.",
      "Group memberships are direct only: nested groups are not followed, and directory roles and administrative units are counted apart from groups. Prism365 does not read mailbox type (shared or user), forwarding or delegates (ADR-0009). Check them before acting; an action tool's receipt may report them.",
    ].slice(0, 20);

    const reportsClean =
      reports?.outcomes.every((outcome: ReportOutcome) => outcome.state === "read") === true &&
      reports.truncated.length === 0 &&
      reports.skippedRows === 0 &&
      reports.unreadableValues === 0;
    return {
      domains: domainInventory(read.users.map((user) => user.userPrincipalName)),
      rows,
      complete:
        read.complete &&
        read.skipped === 0 &&
        unreadableDates === 0 &&
        reportsClean &&
        memberships.notAsked === 0 &&
        memberships.unanswered === 0 &&
        memberships.partial === 0,
      summary,
      notes,
    };
  },
};

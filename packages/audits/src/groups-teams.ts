import { type AuditDefinitionInput, type AuditRow, cleanText } from "@prism365/contracts";
import {
  CSV_MAX_ROWS,
  type CsvTable,
  GraphBoundaryError,
  GraphRequestError,
  GraphRetryDeferredError,
} from "@prism365/graph";
import { z } from "zod";
import { countOf, failureOf, type Tally } from "./activity-reports.ts";
import type { Audit, AuditContext } from "./definition.ts";
import { readPages } from "./graph-pages.ts";
import { HASHED, lastActivityOf, readSetting, type Setting } from "./m365-app-usage.ts";

// Every group and team, who owns it, and whether anyone uses it: ownerless groups, teams that
// depend on one person or on an account that is disabled, teams nobody has used for 90 days,
// groups about to expire. Members are not read (the read pack cannot reach them); owners are read
// for the groups that need one. Microsoft Graph v1.0 documentation verified 2026-10-24:
// https://learn.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/group-list-owners?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/reportroot-getteamsteamactivitydetail?view=graph-rest-1.0

const groupsContext = { feature: "directory.groups.read", permission: "Group.Read.All" } as const;
const reportContext = { feature: "reports.usage.read", permission: "Reports.Read.All" } as const;

// A stored result holds 20,000 rows, so no more groups than that are read.
export const MAX_GROUPS = 20_000;
// Owner lists are one request each. 3,000 of them at four at a time stay well inside the audit's
// ten minutes; the rest of the groups say so.
export const MAX_OWNER_LISTS = 3_000;
export const OWNER_CONCURRENCY = 4;
// Microsoft documents $top up to 999 on the group list.
const PAGE_SIZE = 999;
const MAX_PAGES = 300;
const MAX_OWNERS_PER_GROUP = 200;
const OWNER_NAMES_SHOWN = 5;
const INACTIVE_DAYS = 90;
const EXPIRY_DAYS = 30;
const DAY = 86_400_000;
const REPORT_PATH = "reports/getTeamsTeamActivityDetail(period='D90')";

const GROUP_FIELDS = [
  "id",
  "displayName",
  "groupTypes",
  "mailEnabled",
  "securityEnabled",
  "resourceProvisioningOptions",
  "visibility",
  "expirationDateTime",
  "renewedDateTime",
  "isAssignableToRole",
  "onPremisesSyncEnabled",
  "createdDateTime",
].join(",");

export const GROUP_KINDS = [
  "Team",
  "Microsoft 365 group",
  "Dynamic security group",
  "Mail-enabled security group",
  "Security group",
  "Distribution list",
] as const;
export type GroupKind = (typeof GROUP_KINDS)[number];

// Worst first. A group gets the first that applies; the tags carry every condition it meets.
export const GROUP_STATUSES = [
  "No owners",
  "Only disabled owners",
  "One owner",
  "Inactive 90 days",
  "Expires within 30 days",
  "Owners not read",
  "No usage data",
  "OK",
] as const;
export type GroupStatus = (typeof GROUP_STATUSES)[number];

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const text = z
  .string()
  .transform((value) => cleanText(value.slice(0, 500)))
  .nullish()
  .catch(null)
  .transform((value) => value ?? null);
const flag = z
  .boolean()
  .nullish()
  .catch(null)
  .transform((value) => value ?? null);
// Labels are compared without regard to letter case, so Microsoft's spelling does not matter.
const labels = z
  .array(z.string().max(100))
  .max(50)
  .nullish()
  .catch(null)
  .transform((value) => (value ?? []).map((label) => label.toLowerCase()));
const isoDate = z.string().datetime({ offset: true });

const groupSchema = z.object({
  id: z.string().regex(GUID),
  displayName: text,
  groupTypes: labels,
  mailEnabled: flag,
  securityEnabled: flag,
  resourceProvisioningOptions: labels,
  visibility: text,
  // Read by hand: a date that cannot be read is counted, not taken for "none".
  expirationDateTime: z.unknown(),
  renewedDateTime: z.unknown(),
  createdDateTime: z.unknown(),
  isAssignableToRole: flag,
  onPremisesSyncEnabled: flag,
});

const ownerSchema = z.object({
  id: z.string().min(1).max(200),
  displayName: text,
  userPrincipalName: text,
  accountEnabled: flag,
});

interface Group {
  id: string;
  name: string | null;
  kind: GroupKind;
  visibility: string | null;
  created: string | null;
  expires: string | null;
  renewed: string | null;
  assignable: boolean | null;
  synced: boolean;
}

// First match: a team is a Microsoft 365 group that was given a team, so it is both and is called
// a team; dynamic membership belongs to a security group once the Microsoft 365 case is out of the
// way; what is neither mail-enabled nor a security group is left as a distribution list.
export function kindOf(group: {
  groupTypes: readonly string[];
  resourceProvisioningOptions: readonly string[];
  mailEnabled: boolean | null;
  securityEnabled: boolean | null;
}): GroupKind {
  if (group.resourceProvisioningOptions.includes("team")) return "Team";
  if (group.groupTypes.includes("unified")) return "Microsoft 365 group";
  if (group.groupTypes.includes("dynamicmembership")) return "Dynamic security group";
  if (group.mailEnabled === true && group.securityEnabled === true)
    return "Mail-enabled security group";
  if (group.securityEnabled === true) return "Security group";
  return "Distribution list";
}

const VISIBILITY: Record<string, string> = {
  public: "Public",
  private: "Private",
  hiddenmembership: "Hidden membership",
};

function groupOf(raw: unknown, unreadableDates: { count: number }): Group | null {
  const parsed = groupSchema.safeParse(raw);
  if (!parsed.success) return null;
  const group = parsed.data;
  const dateOf = (value: unknown): string | null => {
    if (value === null || value === undefined) return null;
    const date = isoDate.safeParse(value);
    if (date.success) return date.data;
    unreadableDates.count += 1;
    return null;
  };
  return {
    id: group.id.toLowerCase(),
    name: group.displayName?.trim() || null,
    kind: kindOf(group),
    visibility: group.visibility
      ? (VISIBILITY[group.visibility.toLowerCase()] ?? group.visibility.slice(0, 40))
      : null,
    created: dateOf(group.createdDateTime),
    expires: dateOf(group.expirationDateTime),
    renewed: dateOf(group.renewedDateTime),
    assignable: group.isAssignableToRole,
    synced: group.onPremisesSyncEnabled === true,
  };
}

// Owners are asked for where an ownerless group is a problem: a Microsoft 365 group or team has
// nobody to approve members or renew it, and a group that can be given a directory role has
// nobody answering for who may join it.
const wantsOwners = (group: Group) =>
  group.kind === "Team" || group.kind === "Microsoft 365 group" || group.assignable === true;

// Teams first, then the oldest groups (the likeliest to have lost their owner), then by ID so the
// order never depends on how Microsoft paged.
function askOrder(groups: readonly Group[]): Group[] {
  const time = (group: Group) => (group.created ? Date.parse(group.created) : Number.MAX_VALUE);
  return groups
    .filter(wantsOwners)
    .sort(
      (a, b) =>
        Number(b.kind === "Team") - Number(a.kind === "Team") ||
        time(a) - time(b) ||
        a.id.localeCompare(b.id),
    );
}

interface Owners {
  count: number;
  enabled: number;
  disabled: number;
  names: string[];
}

interface OwnerRead {
  byGroup: Map<string, Owners>;
  // Groups asked for that got no answer.
  unanswered: number;
  // Groups beyond the request limit, which were not asked.
  notAsked: number;
  // Owner lists that were cut short or held records that could not be read.
  partial: number;
  // Why reading stopped before every group was asked, if it did.
  stopped: "throttled" | "refused" | null;
  status: number | null;
}

function ownersOf(items: readonly unknown[]): { owners: Owners; unreadable: number } {
  const owners: Owners = { count: 0, enabled: 0, disabled: 0, names: [] };
  let unreadable = 0;
  for (const raw of items) {
    const parsed = ownerSchema.safeParse(raw);
    if (!parsed.success) {
      unreadable += 1;
      continue;
    }
    const owner = parsed.data;
    owners.count += 1;
    // A service principal or a record that does not carry the field is unknown, not enabled.
    if (owner.accountEnabled === true) owners.enabled += 1;
    else if (owner.accountEnabled === false) owners.disabled += 1;
    const name = (owner.displayName ?? owner.userPrincipalName ?? "").trim();
    if (name) owners.names.push(name.slice(0, 200));
  }
  owners.names.sort((a, b) => a.localeCompare(b));
  owners.names.length = Math.min(owners.names.length, OWNER_NAMES_SHOWN);
  return { owners, unreadable };
}

// Asks for each group's owners, a few at a time. A 404 means the group is gone since the list was
// read and costs only that group. Throttling, or a refusal for any other reason, ends the reading
// and keeps what was read: the same refusal would meet every group. A session that ended is not
// swallowed; the runner explains it.
async function readOwners(
  context: AuditContext,
  groups: readonly Group[],
  limit: number,
): Promise<OwnerRead & { unreadable: number }> {
  const ordered = askOrder(groups);
  const asked = ordered.slice(0, limit);
  const result: OwnerRead & { unreadable: number } = {
    byGroup: new Map(),
    unanswered: 0,
    notAsked: ordered.length - asked.length,
    partial: 0,
    stopped: null,
    status: null,
    unreadable: 0,
  };
  let next = 0;
  let halted = false;
  const worker = async () => {
    while (!halted) {
      context.signal.throwIfAborted();
      const group = asked[next];
      next += 1;
      if (!group) return;
      try {
        const read = await context.graph.getCollection(
          `groups/${encodeURIComponent(group.id)}/owners?$select=id,displayName,userPrincipalName,accountEnabled`,
          groupsContext,
          { schema: z.unknown(), maxPages: 5, maxRecords: MAX_OWNERS_PER_GROUP },
        );
        const { owners, unreadable } = ownersOf(read.items);
        if (!read.complete || unreadable > 0) result.partial += 1;
        result.unreadable += unreadable;
        result.byGroup.set(group.id, owners);
      } catch (error) {
        if (context.signal.aborted) throw error;
        if (error instanceof GraphRetryDeferredError) {
          result.stopped ??= "throttled";
          halted = true;
        } else if (error instanceof GraphRequestError && error.status === 404) {
          // Gone since the directory was read; nothing to say about it.
        } else if (error instanceof GraphRequestError && error.status !== 401) {
          result.stopped ??= "refused";
          result.status ??= error.status;
          halted = true;
        } else if (error instanceof GraphBoundaryError && error.reason !== "credential") {
          // One answer Prism365 could not read safely; the group is left without a result.
        } else {
          halted = true;
          throw error;
        }
      }
    }
  };
  await Promise.all(Array.from({ length: OWNER_CONCURRENCY }, worker));
  result.unanswered = asked.length - result.byGroup.size;
  return result;
}

// What the Teams activity report says about one team. Null where the report leaves it empty.
interface TeamUsage {
  last: string | null;
  // none: the report has no date. unreadable: it has one that cannot be real.
  lastState: "none" | "known" | "unreadable";
  days: number | null;
  activeUsers: number | null;
  channels: number | null;
  guests: number | null;
}

type TeamsReport =
  | { state: "not_requested" }
  | { state: "failed"; reason: string }
  | { state: "concealed"; setting: Setting }
  | {
      state: "read";
      byTeam: Map<string, TeamUsage>;
      refreshed: string | null;
      skippedRows: number;
      truncated: boolean;
      unreadableValues: number;
      unreadableDates: number;
    };

const REPORT_REQUIRED = ["Team Id", "Last Activity Date"] as const;

// The Teams activity report joined by Team Id, which is the group's object ID. Names are not
// used: whether concealment hides the Team Id is not documented, so a report whose Team Ids all
// look like hashes is treated as concealed and nothing is matched or unmasked.
async function readTeams(context: AuditContext): Promise<TeamsReport> {
  if (!context.granted.has("Reports.Read.All")) return { state: "not_requested" };
  let table: CsvTable;
  try {
    table = await context.graph.getReport(REPORT_PATH, reportContext);
  } catch (error) {
    if (context.signal.aborted) throw error;
    const reason = failureOf(error);
    if (reason === null) throw error;
    return { state: "failed", reason };
  }
  // Headers are matched by name and without regard to letter case, never by position.
  const index = new Map<string, number>();
  for (const [position, name] of table.header.entries()) {
    const key = name.trim().toLowerCase();
    if (!index.has(key)) index.set(key, position);
  }
  const missing = REPORT_REQUIRED.find((name) => !index.has(name.toLowerCase()));
  if (missing !== undefined) return { state: "failed", reason: `it has no "${missing}" column` };
  const cell = (row: readonly string[], name: string): string | undefined => {
    const at = index.get(name.toLowerCase());
    return at === undefined ? undefined : row[at];
  };
  const identifier = (row: readonly string[]) => (cell(row, "Team Id") ?? "").trim();
  if (table.rows.length > 0 && table.rows.every((row) => HASHED.test(identifier(row))))
    return { state: "concealed", setting: await readSetting(context) };

  const tally: Tally = { unreadable: 0 };
  const byTeam = new Map<string, TeamUsage>();
  let skippedRows = table.skippedRows;
  let unreadableDates = 0;
  for (const row of table.rows) {
    const id = identifier(row).toLowerCase();
    if (!GUID.test(id)) {
      skippedRows += 1;
      continue;
    }
    // One row for each team is enough; a second would only repeat or contradict it.
    if (byTeam.has(id)) continue;
    const last = lastActivityOf(cell(row, "Last Activity Date")?.trim() || null, context.now);
    if (last.state === "unreadable") unreadableDates += 1;
    byTeam.set(id, {
      last: last.at,
      lastState: last.state,
      days: last.days,
      activeUsers: countOf(cell(row, "Active Users"), tally) ?? null,
      channels: countOf(cell(row, "Active Channels"), tally) ?? null,
      guests: countOf(cell(row, "Guests"), tally) ?? null,
    });
  }
  const refreshAt = index.get("report refresh date");
  const refreshed = lastActivityOf(
    refreshAt === undefined ? null : (table.rows[0]?.[refreshAt]?.trim() ?? null),
    context.now,
  );
  return {
    state: "read",
    byTeam,
    refreshed: refreshed.state === "known" && refreshed.at ? refreshed.at.slice(0, 10) : null,
    skippedRows,
    truncated: table.truncated,
    unreadableValues: tally.unreadable,
    unreadableDates,
  };
}

const plural = (count: number, one: string, many: string) =>
  `${count.toLocaleString("en-US")} ${count === 1 ? one : many}`;

const concealedAdvice =
  "A Global Administrator can turn off concealment in the Microsoft 365 admin center (Settings, Org settings, Reports). Prism365 will not unmask anything itself.";

function reportNotes(report: TeamsReport): string[] {
  if (report.state === "not_requested")
    return [
      "The Teams activity report was not read: this run did not have Reports.Read.All (the usage reports read pack is off, or the sign-in has not granted it). Teams show No usage data, and none can be called inactive.",
    ];
  if (report.state === "failed")
    return [
      `The Teams activity report could not be read: ${report.reason}. Teams show No usage data, and none can be called inactive.`,
    ];
  if (report.state === "concealed")
    return [
      `${
        report.setting.concealed === true
          ? "Microsoft's report privacy setting says names are concealed, and the Teams report returns hashed team identifiers."
          : report.setting.concealed === false
            ? "Microsoft's privacy setting says names are shown, but the Teams report still returns hashed team identifiers (a change can take a few minutes to reach the reports)."
            : "The Teams report's team identifiers look like hashes, so they are treated as concealed."
      } Prism365 did not match them to groups: teams show No usage data, and none can be called inactive.`,
      concealedAdvice,
    ];
  const notes = [
    `Team usage comes from Microsoft's Teams activity report for the last 90 days${report.refreshed ? `, refreshed ${report.refreshed}` : ""}, matched to groups by team ID. The report lags by 24 to 72 hours and sometimes by several days. A team is inactive when the report has no activity date for it or one more than 90 days old.`,
  ];
  if (report.truncated)
    notes.push(
      `The Teams report has more than ${CSV_MAX_ROWS.toLocaleString("en-US")} rows; the rest was not read, so a team may show No usage data.`,
    );
  if (report.skippedRows > 0)
    notes.push(
      `${plural(report.skippedRows, "row in the Teams report", "rows in the Teams report")} could not be read and ${report.skippedRows === 1 ? "is" : "are"} left out.`,
    );
  if (report.unreadableValues > 0)
    notes.push(
      `${plural(report.unreadableValues, "value in the Teams report", "values in the Teams report")} could not be read as a number and ${report.unreadableValues === 1 ? "shows" : "show"} as empty.`,
    );
  if (report.unreadableDates > 0)
    notes.push(
      `${plural(report.unreadableDates, "team has", "teams have")} a last activity date Prism365 could not read. ${report.unreadableDates === 1 ? "It is" : "They are"} never called inactive.`,
    );
  return notes;
}

type Column = AuditDefinitionInput["columns"][number];

const usage = (key: string, label: string, help: string): Column => ({
  key,
  label,
  kind: key === "lastActivity" ? "date" : "number",
  empty: "No usage data",
  help,
});

const columns: Column[] = [
  { key: "group", label: "Group", kind: "text" },
  {
    key: "kind",
    label: "Kind",
    kind: "enum",
    facet: true,
    help: "The first that applies: Team, Microsoft 365 group, Dynamic security group, Mail-enabled security group, Security group, or Distribution list for the rest. A team is found by the group's provisioning options, which a few old unused teams lack; those show as Microsoft 365 groups.",
  },
  {
    key: "status",
    label: "Status",
    kind: "enum",
    facet: true,
    help: "The first that applies. No owners: owners were read and there are none. Only disabled owners: every owner's account is disabled. One owner: a team with a single owner. Inactive 90 days: a team whose report row has no activity date or one over 90 days old. Expires within 30 days: the expiry date is within 30 days or already past. Owners not read: owners were due and Prism365 did not get them. No usage data: a team the Teams report does not list. OK: none of these.",
  },
  {
    key: "visibility",
    label: "Visibility",
    kind: "enum",
    facet: true,
    empty: "Not set",
    help: "Public groups can be joined by anyone in the organisation. Security groups and distribution lists have no visibility.",
  },
  {
    key: "owners",
    label: "Owners",
    kind: "number",
    empty: "Not read",
    help: "Owners are read for Microsoft 365 groups, teams and groups that can be given a directory role. Other groups show Not read.",
  },
  {
    key: "enabledOwners",
    label: "Enabled owners",
    kind: "number",
    empty: "Not read",
    help: "Owners whose account is enabled. An owner Microsoft does not say is enabled, such as an app, is not counted.",
  },
  {
    key: "ownerNames",
    label: "Owner names",
    kind: "list",
    empty: "Not read",
    help: "The first five by name.",
  },
  { key: "created", label: "Created", kind: "date", empty: "Not recorded" },
  {
    key: "expires",
    label: "Expires",
    kind: "date",
    empty: "No expiry",
    help: "When the group's expiration policy deletes it unless someone renews it. Empty means no policy applies.",
  },
  { key: "renewed", label: "Renewed", kind: "date", hidden: true, empty: "Not recorded" },
  {
    key: "roleAssignable",
    label: "Role-assignable",
    kind: "boolean",
    empty: "Unknown",
    help: "A group that can be given a Microsoft Entra role. Whoever owns it controls who gets that role.",
  },
  {
    key: "synced",
    label: "Synced from on-premises",
    kind: "boolean",
    hidden: true,
    help: "Managed in an on-premises directory and copied to the cloud, so it is changed there.",
  },
  usage(
    "lastActivity",
    "Last activity",
    "From the Teams activity report. Empty means the team is not in the report, or the report has no date for it.",
  ),
  usage("activeUsers", "Active users", "People active in the team in the last 90 days."),
  usage("guests", "Guests", "Guests in the team, as the Teams activity report counts them."),
  usage(
    "activeChannels",
    "Active channels",
    "Channels with activity in the last 90 days, as the report counts them.",
  ),
];

interface Entry {
  group: Group;
  owners: Owners | undefined;
  usage: TeamUsage | undefined;
  inactive: boolean;
  expiring: boolean;
  status: GroupStatus;
}

export interface GroupsTeamsLimits {
  maxGroups?: number;
  maxOwnerLists?: number;
}

export function createGroupsTeamsAudit(limits: GroupsTeamsLimits = {}): Audit {
  const maxGroups = limits.maxGroups ?? MAX_GROUPS;
  const maxOwnerLists = limits.maxOwnerLists ?? MAX_OWNER_LISTS;
  return {
    definition: {
      id: "groups.teams",
      handoff: false,
      version: 1,
      title: "Groups and Teams",
      description:
        "Which groups and teams have nobody looking after them: no owners, only disabled owners, a single owner, no use for 90 days, or an expiry date close. Owners are named; members are not read.",
      category: "Groups",
      subject: { one: "group", many: "groups" },
      columns,
      views: [
        { id: "no-owners", label: "No owners", tags: ["no-owners"] },
        { id: "disabled-owners", label: "Only disabled owners", tags: ["disabled-owners"] },
        { id: "one-owner", label: "Teams with one owner", tags: ["one-owner"] },
        { id: "guests", label: "Teams with guests", tags: ["guests"] },
        { id: "inactive", label: "Inactive teams", tags: ["inactive"] },
        { id: "public-teams", label: "Public teams", tags: ["public-team"] },
        { id: "expiring", label: "Expiring within 30 days", tags: ["expiring"] },
        { id: "role-assignable", label: "Role-assignable groups", tags: ["role-assignable"] },
        { id: "owners-not-read", label: "Owners not read", tags: ["owners-not-read"] },
        { id: "no-usage", label: "No usage data", tags: ["no-usage"] },
        { id: "teams", label: "Teams", tags: ["team"] },
        { id: "all", label: "All groups", tags: [] },
      ],
      defaultView: "all",
      defaultSort: { column: "group", direction: "asc" },
      requirements: [
        "A role that can read groups and their owners, such as Global Reader or Directory Readers.",
        "The usage columns need the usage reports read pack and a role that can read them, such as Reports Reader or Teams Administrator.",
      ],
    },
    packs: ["directory.groups.read"],
    optionalPacks: ["reports.usage.read", "reports.settings.read"],
    scoped: false,
    forbidden:
      "Microsoft refused to read groups or their owners. It needs Group.Read.All and a role such as Global Reader or Directory Readers.",
    async run(context) {
      context.signal.throwIfAborted();
      const read = await readPages(
        context,
        groupsContext,
        `groups?$select=${GROUP_FIELDS}&$top=${PAGE_SIZE}`,
        {
          maxRecords: maxGroups,
          maxPages: MAX_PAGES,
        },
      );
      const unreadableDates = { count: 0 };
      const groups: Group[] = [];
      let skipped = 0;
      const seen = new Set<string>();
      for (const raw of read.items) {
        const group = groupOf(raw, unreadableDates);
        if (group === null) skipped += 1;
        else if (!seen.has(group.id)) {
          seen.add(group.id);
          groups.push(group);
        }
      }

      const owners = await readOwners(context, groups, maxOwnerLists);
      const report = await readTeams(context);
      const byTeam = report.state === "read" ? report.byTeam : null;

      const now = context.now.getTime();
      const entries: Entry[] = groups.map((group) => {
        const ownerList = owners.byGroup.get(group.id);
        const row = group.kind === "Team" ? byTeam?.get(group.id) : undefined;
        const inactive =
          row !== undefined &&
          (row.lastState === "none" ||
            (row.lastState === "known" && (row.days ?? 0) > INACTIVE_DAYS));
        const expiring =
          group.expires !== null && Date.parse(group.expires) <= now + EXPIRY_DAYS * DAY;
        const status = ((): GroupStatus => {
          if (ownerList && ownerList.count === 0) return "No owners";
          if (ownerList && ownerList.count > 0 && ownerList.disabled === ownerList.count)
            return "Only disabled owners";
          if (group.kind === "Team" && ownerList?.count === 1) return "One owner";
          if (inactive) return "Inactive 90 days";
          if (expiring) return "Expires within 30 days";
          if (wantsOwners(group) && !ownerList) return "Owners not read";
          if (group.kind === "Team" && row === undefined) return "No usage data";
          return "OK";
        })();
        return { group, owners: ownerList, usage: row, inactive, expiring, status };
      });
      entries.sort(
        (a, b) =>
          GROUP_STATUSES.indexOf(a.status) - GROUP_STATUSES.indexOf(b.status) ||
          (a.group.name ?? "").toLowerCase().localeCompare((b.group.name ?? "").toLowerCase()) ||
          a.group.id.localeCompare(b.group.id),
      );

      const rows: AuditRow[] = entries.map((entry) => {
        const { group, owners: ownerList, usage: row } = entry;
        const team = group.kind === "Team";
        return {
          key: group.id,
          // A group has no page of its own.
          personId: null,
          tags: [
            ...(team ? ["team"] : []),
            ...(ownerList?.count === 0 ? ["no-owners"] : []),
            ...(ownerList && ownerList.count > 0 && ownerList.disabled === ownerList.count
              ? ["disabled-owners"]
              : []),
            ...(team && ownerList?.count === 1 ? ["one-owner"] : []),
            ...((row?.guests ?? 0) > 0 ? ["guests"] : []),
            ...(entry.inactive ? ["inactive"] : []),
            ...(team && group.visibility === "Public" ? ["public-team"] : []),
            ...(entry.expiring ? ["expiring"] : []),
            ...(group.assignable === true ? ["role-assignable"] : []),
            ...(entry.status === "Owners not read" ? ["owners-not-read"] : []),
            ...(entry.status === "No usage data" ? ["no-usage"] : []),
          ],
          cells: {
            group: group.name ?? "Unnamed group",
            kind: group.kind,
            status: entry.status,
            visibility: group.visibility,
            owners: ownerList?.count ?? null,
            enabledOwners: ownerList?.enabled ?? null,
            ownerNames: ownerList?.names ?? null,
            created: group.created,
            expires: group.expires,
            renewed: group.renewed,
            roleAssignable: group.assignable,
            synced: group.synced,
            lastActivity: row?.last ?? null,
            activeUsers: row?.activeUsers ?? null,
            guests: row?.guests ?? null,
            activeChannels: row?.channels ?? null,
          },
        };
      });

      const tagged = (tag: string) => rows.filter((row) => row.tags.includes(tag)).length;
      const usageRead = report.state === "read";
      const tone = (value: number, level: "warning" | "neutral" = "warning") =>
        value > 0 ? level : "good";

      // Microsoft 365 groups that the Teams report lists are teams whose group lacks the
      // provisioning option; they stay Microsoft 365 groups here and say so.
      const unmarked = byTeam
        ? groups.filter((group) => group.kind === "Microsoft 365 group" && byTeam.has(group.id))
            .length
        : 0;
      const noOwnerAnswer = owners.unanswered > 0 || owners.notAsked > 0 || owners.stopped !== null;

      const notes = [
        "Members are not read, so this audit cannot say who is in a group or how many people it has.",
        "Owners are read for Microsoft 365 groups, teams and groups that can be given a directory role. Other groups show Not read. A group with no owner has nobody to approve members, answer for its access or renew it.",
        "Expiry dates come from a group's expiration policy. A Microsoft 365 group is deleted after its expiry date unless an owner renews it, and a date already past counts as expiring.",
        ...reportNotes(report),
        ...(unmarked > 0
          ? [
              `${plural(unmarked, "Microsoft 365 group appears", "Microsoft 365 groups appear")} in the Teams report, so ${unmarked === 1 ? "it is" : "they are"} really ${unmarked === 1 ? "a team" : "teams"} whose group lacks the provisioning option. ${unmarked === 1 ? "It is" : "They are"} listed as Microsoft 365 ${unmarked === 1 ? "group" : "groups"} and ${unmarked === 1 ? "its" : "their"} usage is not matched.`,
            ]
          : []),
        ...(owners.notAsked > 0
          ? [
              `Owners were read for ${maxOwnerLists.toLocaleString("en-US")} of ${(owners.notAsked + maxOwnerLists).toLocaleString("en-US")} groups, teams first and then the oldest groups. ${plural(owners.notAsked, "group shows", "groups show")} Owners not read.`,
            ]
          : []),
        ...(owners.stopped === "throttled"
          ? [
              `Microsoft asked Prism365 to slow down, so owners stopped after ${plural(owners.byGroup.size, "group", "groups")}. Run the audit again in a few minutes.`,
            ]
          : owners.stopped === "refused"
            ? [
                `Microsoft refused to list owners (HTTP ${owners.status}), so ${owners.byGroup.size === 0 ? "none were read" : `they were read for ${plural(owners.byGroup.size, "group", "groups")} only`}. Listing a group's owners needs Group.Read.All and a role such as Global Reader or Directory Readers.`,
              ]
            : owners.unanswered > 0
              ? [
                  `${plural(owners.unanswered, "group", "groups")} had no answer for owners and ${owners.unanswered === 1 ? "shows" : "show"} Owners not read.`,
                ]
              : []),
        ...(owners.partial > 0
          ? [
              `${plural(owners.partial, "owner list holds", "owner lists hold")} more than Prism365 reads or records it could not read, so ${owners.partial === 1 ? "its" : "their"} counts may be low.`,
            ]
          : []),
        ...(read.throttled
          ? [
              `Microsoft asked Prism365 to slow down after ${plural(read.items.length, "group", "groups")}. The rest were not read. Run the audit again in a few minutes.`,
            ]
          : read.complete
            ? []
            : [
                `Stopped at ${maxGroups.toLocaleString("en-US")} groups; the rest of the directory was not read.`,
              ]),
        ...(skipped > 0
          ? [
              `${plural(skipped, "group", "groups")} could not be read and ${skipped === 1 ? "is" : "are"} left out.`,
            ]
          : []),
        ...(unreadableDates.count > 0
          ? [
              `${plural(unreadableDates.count, "date", "dates")} on groups could not be read and ${unreadableDates.count === 1 ? "shows" : "show"} as empty. A group with an expiry date that cannot be read is not called expiring.`,
            ]
          : []),
      ];

      const usageLabel = (label: string) => (usageRead ? label : `${label} (usage not read)`);
      const ownerless = tagged("no-owners");
      const disabledOnly = tagged("disabled-owners");
      const oneOwner = tagged("one-owner");
      const expiring = tagged("expiring");
      const withGuests = usageRead ? tagged("guests") : 0;
      const inactiveTeams = usageRead ? tagged("inactive") : 0;
      return {
        rows,
        complete:
          read.complete &&
          skipped === 0 &&
          unreadableDates.count === 0 &&
          !noOwnerAnswer &&
          owners.partial === 0 &&
          (report.state === "not_requested" ||
            (report.state === "read" &&
              !report.truncated &&
              report.skippedRows === 0 &&
              report.unreadableValues === 0 &&
              report.unreadableDates === 0)),
        summary: [
          {
            label: "Groups and teams with no owners",
            value: ownerless,
            tone: tone(ownerless),
            view: "no-owners",
          },
          {
            label: "Groups with only disabled owners",
            value: disabledOnly,
            tone: tone(disabledOnly),
            view: "disabled-owners",
          },
          {
            label: "Teams with one owner",
            value: oneOwner,
            tone: tone(oneOwner),
            view: "one-owner",
          },
          {
            label: usageLabel("Teams with guests"),
            value: withGuests,
            tone: "neutral",
            view: usageRead ? "guests" : null,
          },
          {
            label: usageLabel("Inactive teams"),
            value: inactiveTeams,
            tone: tone(inactiveTeams),
            view: usageRead ? "inactive" : null,
          },
          {
            label: "Public teams",
            value: tagged("public-team"),
            tone: "neutral",
            view: "public-teams",
          },
          {
            label: "Groups expiring in 30 days",
            value: expiring,
            tone: tone(expiring),
            view: "expiring",
          },
          {
            label: "Role-assignable groups",
            value: tagged("role-assignable"),
            tone: "neutral",
            view: "role-assignable",
          },
        ],
        notes,
      };
    },
  };
}

export const groupsTeamsAudit: Audit = createGroupsTeamsAudit();

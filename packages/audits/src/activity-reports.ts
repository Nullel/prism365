import {
  type CsvTable,
  GraphBoundaryError,
  GraphReportDownloadError,
  GraphReportHostError,
  GraphRequestError,
  GraphRetryDeferredError,
} from "@prism365/graph";
import type { AuditContext } from "./definition.ts";
import type { DirectoryUser } from "./directory.ts";
import { HASHED, lastActivityOf } from "./m365-app-usage.ts";

// The seven per-user usage reports behind the Account activity audit, read as CSV through the
// Graph client's report reader (ADR-0010) and joined to directory accounts. Every report is read
// by column name, never by position, and every value is parsed to a date, a count or a flag:
// no report text is ever shown. Microsoft Graph v1.0 documentation verified 2026-10-23:
// https://learn.microsoft.com/en-us/graph/api/reportroot-getoffice365activeuserdetail?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/reportroot-getemailactivityuserdetail?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/reportroot-getteamsuseractivityuserdetail?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/reportroot-getonedriveactivityuserdetail?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/reportroot-getonedriveusageaccountdetail?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/reportroot-getsharepointactivityuserdetail?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/reportroot-getmailboxusagedetail?view=graph-rest-1.0

const reportContext = { feature: "reports.usage.read", permission: "Reports.Read.All" } as const;

export type ReportId =
  | "activeUsers"
  | "email"
  | "teams"
  | "oneDriveActivity"
  | "oneDriveUsage"
  | "sharePoint"
  | "mailbox";

// What the reports say about one account. A field is undefined when no report that was read
// carries it for the account (no row, or no such column), null when the report has the field and
// it is empty or could not be read. Dates stay as the report wrote them ("2026-10-20") until
// they are judged, so an unreadable one is never mistaken for none.
export interface AccountFacts {
  activeExchange?: string | null | undefined;
  activeTeams?: string | null | undefined;
  activeOneDrive?: string | null | undefined;
  activeSharePoint?: string | null | undefined;
  mailLast?: string | null | undefined;
  mailSent?: number | null | undefined;
  mailReceived?: number | null | undefined;
  mailRead?: number | null | undefined;
  teamsLast?: string | null | undefined;
  teamsChats?: number | null | undefined;
  teamsCalls?: number | null | undefined;
  teamsMeetings?: number | null | undefined;
  teamsOther?: boolean | null | undefined;
  oneDriveLast?: string | null | undefined;
  oneDriveFiles?: number | null | undefined;
  driveLast?: string | null | undefined;
  driveActive?: number | null | undefined;
  driveFiles?: number | null | undefined;
  driveBytes?: number | null | undefined;
  sharePointLast?: string | null | undefined;
  sharePointFiles?: number | null | undefined;
  sharePointPages?: number | null | undefined;
  mailboxLast?: string | null | undefined;
  mailboxItems?: number | null | undefined;
  mailboxBytes?: number | null | undefined;
  mailboxArchive?: boolean | null | undefined;
}

type Cell = (name: string) => string | undefined;
interface Tally {
  // Values that were neither empty nor the kind of value the column holds.
  unreadable: number;
}

const dateOf = (raw: string | undefined): string | null | undefined => {
  if (raw === undefined) return undefined;
  const value = raw.trim();
  return value === "" ? null : value.slice(0, 40);
};

function countOf(raw: string | undefined, tally: Tally): number | null | undefined {
  if (raw === undefined) return undefined;
  const value = raw.trim();
  if (value === "") return null;
  if (!/^\d{1,15}$/.test(value)) {
    tally.unreadable += 1;
    return null;
  }
  return Number(value);
}

function flagOf(raw: string | undefined, tally: Tally): boolean | null | undefined {
  if (raw === undefined) return undefined;
  const value = raw.trim().toLowerCase();
  if (value === "") return null;
  if (value === "true") return true;
  if (value === "false") return false;
  tally.unreadable += 1;
  return null;
}

// A total over parts a report may not carry: undefined when it carries none of them, null when it
// carries them and none could be read.
function sumOf(...parts: (number | null | undefined)[]): number | null | undefined {
  const present = parts.filter((part) => part !== undefined);
  if (present.length === 0) return undefined;
  const numbers = present.filter((part): part is number => part !== null);
  return numbers.length === 0 ? null : numbers.reduce((total, part) => total + part, 0);
}

interface ReportSpec {
  id: ReportId;
  // How a note names the report.
  name: string;
  path: string;
  // The column that names the account, and the Teams report's column for its object ID.
  principal: string;
  userId?: string;
  // Columns without which the report cannot say anything about activity.
  required: readonly string[];
  read(cell: Cell, tally: Tally): AccountFacts;
}

const PERIOD = "D90";
const report = (name: string) => `reports/get${name}(period='${PERIOD}')`;

const specs: readonly ReportSpec[] = [
  {
    id: "activeUsers",
    name: "active users",
    path: report("Office365ActiveUserDetail"),
    principal: "User Principal Name",
    required: [
      "User Principal Name",
      "Exchange Last Activity Date",
      "OneDrive Last Activity Date",
      "SharePoint Last Activity Date",
      "Teams Last Activity Date",
    ],
    read: (cell) => ({
      activeExchange: dateOf(cell("Exchange Last Activity Date")),
      activeTeams: dateOf(cell("Teams Last Activity Date")),
      activeOneDrive: dateOf(cell("OneDrive Last Activity Date")),
      activeSharePoint: dateOf(cell("SharePoint Last Activity Date")),
    }),
  },
  {
    id: "email",
    name: "email activity",
    path: report("EmailActivityUserDetail"),
    principal: "User Principal Name",
    required: ["User Principal Name", "Last Activity Date"],
    read: (cell, tally) => ({
      mailLast: dateOf(cell("Last Activity Date")),
      mailSent: countOf(cell("Send Count"), tally),
      mailReceived: countOf(cell("Receive Count"), tally),
      mailRead: countOf(cell("Read Count"), tally),
    }),
  },
  {
    id: "teams",
    name: "Teams activity",
    path: report("TeamsUserActivityUserDetail"),
    principal: "User Principal Name",
    userId: "User Id",
    required: ["User Principal Name", "Last Activity Date"],
    read: (cell, tally) => ({
      teamsLast: dateOf(cell("Last Activity Date")),
      teamsChats: sumOf(
        countOf(cell("Team Chat Message Count"), tally),
        countOf(cell("Private Chat Message Count"), tally),
      ),
      teamsCalls: countOf(cell("Call Count"), tally),
      teamsMeetings: countOf(cell("Meetings Attended Count"), tally),
      teamsOther: flagOf(cell("Has Other Action"), tally),
    }),
  },
  {
    id: "oneDriveActivity",
    name: "OneDrive activity",
    path: report("OneDriveActivityUserDetail"),
    principal: "User Principal Name",
    required: ["User Principal Name", "Last Activity Date"],
    read: (cell, tally) => ({
      oneDriveLast: dateOf(cell("Last Activity Date")),
      oneDriveFiles: countOf(cell("Viewed Or Edited File Count"), tally),
    }),
  },
  {
    id: "oneDriveUsage",
    name: "OneDrive usage",
    path: report("OneDriveUsageAccountDetail"),
    principal: "Owner Principal Name",
    required: ["Owner Principal Name", "Last Activity Date"],
    read: (cell, tally) => ({
      driveLast: dateOf(cell("Last Activity Date")),
      driveFiles: countOf(cell("File Count"), tally),
      driveActive: countOf(cell("Active File Count"), tally),
      driveBytes: countOf(cell("Storage Used (Byte)"), tally),
    }),
  },
  {
    id: "sharePoint",
    name: "SharePoint activity",
    path: report("SharePointActivityUserDetail"),
    principal: "User Principal Name",
    required: ["User Principal Name", "Last Activity Date"],
    read: (cell, tally) => ({
      sharePointLast: dateOf(cell("Last Activity Date")),
      sharePointFiles: countOf(cell("Viewed Or Edited File Count"), tally),
      sharePointPages: countOf(cell("Visited Page Count"), tally),
    }),
  },
  {
    id: "mailbox",
    name: "mailbox usage",
    path: report("MailboxUsageDetail"),
    principal: "User Principal Name",
    required: ["User Principal Name", "Last Activity Date"],
    // The archive column is missing from some of Microsoft's samples, so it is optional.
    read: (cell, tally) => ({
      mailboxLast: dateOf(cell("Last Activity Date")),
      mailboxItems: countOf(cell("Item Count"), tally),
      mailboxBytes: countOf(cell("Storage Used (Byte)"), tally),
      mailboxArchive: flagOf(cell("Has Archive"), tally),
    }),
  },
];

export interface ReportOutcome {
  id: ReportId;
  name: string;
  state: "read" | "failed";
  // Why a failed report could not be read, as a phrase that follows a colon. Never provider text.
  reason?: string;
}

export interface ReportsRead {
  outcomes: ReportOutcome[];
  // The reports that were read, whole enough to join.
  readIds: Set<ReportId>;
  facts: Map<string, AccountFacts>;
  // The identifiers look like hashes: Microsoft is concealing names, so nothing was joined.
  concealed: boolean;
  // Distinct identifiers in the reports that match no account in the directory.
  unmatched: number;
  // Values that could not be read as a number or a flag.
  unreadableValues: number;
  // Rows the CSV reader left out, over every report that was read.
  skippedRows: number;
  // The reports that had more rows than the reader keeps.
  truncated: string[];
  // The oldest refresh date over the reports that were read, as YYYY-MM-DD.
  refreshed: string | null;
}

// A name that two spellings of an account share: compatibility-normalised and in lower case.
export const joinKey = (value: string): string => value.normalize("NFKC").trim().toLowerCase();

// Why a report could not be read, as a fixed phrase, or null for an error the audit must not
// swallow (a session that ended, or anything that is a bug rather than an answer).
function failureOf(error: unknown): string | null {
  if (error instanceof GraphReportHostError)
    return error.host === null
      ? "Microsoft sent a download link Prism365 could not read"
      : error.knownHost
        ? `the download link on ${error.host} was not accepted (${error.problem})`
        : `Microsoft sent it from a host Prism365 does not recognise (${error.host})`;
  if (error instanceof GraphReportDownloadError) return "the download link failed";
  if (error instanceof GraphRetryDeferredError)
    return "Microsoft is limiting how often reports can be read, so run the audit again in about ten minutes";
  if (error instanceof GraphRequestError) {
    if (error.status === 401) return null;
    if (error.status === 403) return "Microsoft refused (HTTP 403)";
    if (error.status === 404) return "Microsoft has no such report (HTTP 404)";
    return `Microsoft answered HTTP ${error.status}`;
  }
  if (error instanceof GraphBoundaryError) {
    if (error.reason === "credential") return null;
    if (error.reason === "timeout") return "Microsoft took too long to answer";
    return `its file could not be read safely (${error.reason})`;
  }
  return null;
}

// The date part of a refresh date a report states, if it is a real one.
function refreshOf(raw: string | undefined, now: Date): string | null {
  if (!raw) return null;
  const last = lastActivityOf(raw.trim(), now);
  return last.state === "known" && last.at ? last.at.slice(0, 10) : null;
}

// Reads the seven reports one at a time, each once (Microsoft allows 14 requests per ten minutes
// for each report, so a run never throttles itself), and joins every row to an account by sign-in
// name, or for the Teams report by object ID first. A report that cannot be read is recorded and
// the others carry on. As soon as a report shows hashed identifiers the rest are not requested:
// concealment is tenant-wide, and Prism365 never tries to unmask a name.
export async function readReports(
  context: AuditContext,
  users: readonly DirectoryUser[],
): Promise<ReportsRead> {
  const byId = new Map<string, DirectoryUser>();
  const byPrincipal = new Map<string, DirectoryUser>();
  for (const user of users) {
    byId.set(joinKey(user.id), user);
    if (user.userPrincipalName && !byPrincipal.has(joinKey(user.userPrincipalName)))
      byPrincipal.set(joinKey(user.userPrincipalName), user);
  }
  const result: ReportsRead = {
    outcomes: [],
    readIds: new Set(),
    facts: new Map(),
    concealed: false,
    unmatched: 0,
    unreadableValues: 0,
    skippedRows: 0,
    truncated: [],
    refreshed: null,
  };
  const unmatched = new Set<string>();
  for (const spec of specs) {
    context.signal.throwIfAborted();
    let table: CsvTable;
    try {
      table = await context.graph.getReport(spec.path, reportContext);
    } catch (error) {
      if (context.signal.aborted) throw error;
      const reason = failureOf(error);
      if (reason === null) throw error;
      result.outcomes.push({ id: spec.id, name: spec.name, state: "failed", reason });
      continue;
    }
    const missing = spec.required.find((name) => !table.index.has(name));
    if (missing !== undefined) {
      result.outcomes.push({
        id: spec.id,
        name: spec.name,
        state: "failed",
        reason: `it has no "${missing}" column`,
      });
      continue;
    }
    const principalAt = table.index.get(spec.principal) ?? -1;
    if (
      table.rows.length > 0 &&
      table.rows.every((row) => HASHED.test((row[principalAt] ?? "").trim()))
    ) {
      result.concealed = true;
      result.outcomes.length = 0;
      result.readIds.clear();
      result.facts.clear();
      return result;
    }
    const idAt = spec.userId === undefined ? undefined : table.index.get(spec.userId);
    const deletedAt = table.index.get("Is Deleted");
    const tally: Tally = { unreadable: 0 };
    const seen = new Set<string>();
    for (const row of table.rows) {
      const principal = (row[principalAt] ?? "").trim();
      const objectId = idAt === undefined ? "" : (row[idAt] ?? "").trim();
      // A row Microsoft marks deleted is an account that is gone, even if a new one took its name.
      const deleted =
        deletedAt !== undefined && (row[deletedAt] ?? "").trim().toLowerCase() === "true";
      const user = deleted
        ? undefined
        : ((objectId ? byId.get(joinKey(objectId)) : undefined) ??
          (principal ? byPrincipal.get(joinKey(principal)) : undefined));
      if (!user) {
        unmatched.add(joinKey(principal || objectId));
        continue;
      }
      // One row for each account is enough; a second would only repeat or contradict it.
      if (seen.has(user.id)) continue;
      seen.add(user.id);
      const cell: Cell = (name) => {
        const at = table.index.get(name);
        return at === undefined ? undefined : row[at];
      };
      const facts = result.facts.get(user.id) ?? {};
      Object.assign(facts, spec.read(cell, tally));
      result.facts.set(user.id, facts);
    }
    const refreshAt = table.index.get("Report Refresh Date");
    const refreshed = refreshOf(
      refreshAt === undefined ? undefined : table.rows[0]?.[refreshAt],
      context.now,
    );
    if (refreshed && (result.refreshed === null || refreshed < result.refreshed))
      result.refreshed = refreshed;
    result.unreadableValues += tally.unreadable;
    result.skippedRows += table.skippedRows;
    if (table.truncated) result.truncated.push(spec.name);
    result.outcomes.push({ id: spec.id, name: spec.name, state: "read" });
    result.readIds.add(spec.id);
  }
  result.unmatched = unmatched.size;
  return result;
}

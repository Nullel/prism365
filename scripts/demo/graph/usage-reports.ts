import { createHash } from "node:crypto";
import { type DirectoryUser, guid, ORGANISATION, type SyntheticTenant } from "../tenant.ts";
import { LABORATORY, LICENSING, STAFF, TEAM_USAGE } from "./groups.ts";
import type { DemoGraphExtension } from "./types.ts";

// The Microsoft 365 apps usage report, the seven CSV usage reports behind the Account activity
// audit (active users, email, Teams, OneDrive activity, OneDrive usage, SharePoint and mailbox),
// the direct group memberships it reads, the per-team Teams activity report behind the Groups and
// Teams audit, and the report privacy setting for the synthetic tenant.
// Names are shown in the demo (a Global Administrator turned concealment off), so the report names
// people by sign-in name. Accounts that have not signed in for 90 days have no activity, which is
// what ties this report to the Inactive accounts audit.

const DAY = 86_400_000;
const PERIOD = 90;

const apps = ["outlook", "word", "excel", "powerPoint", "oneNote", "teams"] as const;
const platforms = [
  ["windows", "Windows"],
  ["mac", "Mac"],
  ["mobile", "Mobile"],
  ["web", "Web"],
] as const;
type App = (typeof apps)[number];
type Platform = (typeof platforms)[number][0];

// A number in [0, 1) that depends only on the label, so the same tenant always gets the same
// report.
function unit(label: string): number {
  return createHash("sha256").update(`usage:${label}`).digest().readUInt32LE(0) / 4_294_967_296;
}

const day = (ms: number) => new Date(ms).toISOString().slice(0, 10);

interface Use {
  apps: Set<App>;
  platforms: Set<Platform>;
}

// What a person with activity used. The first few of each pattern are planted so every tile of
// the audit has something in it.
function useOf(user: DirectoryUser, position: number): Use {
  const label = user.id;
  if (position % 11 === 3)
    return { apps: new Set<App>(["teams"]), platforms: new Set<Platform>(["windows"]) };
  if (position % 13 === 5)
    return {
      apps: new Set<App>(["outlook", "word", "teams"]),
      platforms: new Set<Platform>(["web"]),
    };
  if (position % 9 === 7)
    return {
      apps: new Set<App>(["outlook", "teams"]),
      platforms: new Set<Platform>(["mobile", "web"]),
    };
  const used = new Set<App>(["outlook"]);
  const chance: [App, number][] = [
    ["teams", 0.85],
    ["word", 0.7],
    ["excel", 0.55],
    ["powerPoint", 0.4],
    ["oneNote", 0.2],
  ];
  for (const [app, probability] of chance) if (unit(`${label}:${app}`) < probability) used.add(app);
  const where = new Set<Platform>([unit(`${label}:desktop`) < 0.75 ? "windows" : "mac"]);
  if (unit(`${label}:mobile`) < 0.45) where.add("mobile");
  if (unit(`${label}:web`) < 0.25) where.add("web");
  return { apps: used, platforms: where };
}

function flagsOf(use: Use) {
  const perApp: Record<string, boolean> = {};
  for (const app of apps)
    for (const [platform, label] of platforms)
      perApp[`${app}${label}`] = use.apps.has(app) && use.platforms.has(platform);
  return {
    reportPeriod: PERIOD,
    windows: use.platforms.has("windows"),
    mac: use.platforms.has("mac"),
    mobile: use.platforms.has("mobile"),
    web: use.platforms.has("web"),
    outlook: use.apps.has("outlook"),
    word: use.apps.has("word"),
    excel: use.apps.has("excel"),
    powerPoint: use.apps.has("powerPoint"),
    oneNote: use.apps.has("oneNote"),
    teams: use.apps.has("teams"),
    ...perApp,
  };
}

function reportOf(tenant: SyntheticTenant) {
  const now = tenant.now.getTime();
  const people = tenant.users.filter(
    (user) => user.userType === "Member" && user.accountEnabled && user.assignedLicenses.length > 0,
  );
  const refreshed = day(now - 2 * DAY);
  let active = 0;
  return people.map((user) => {
    const signedIn = user.signInActivity.lastSuccessfulSignInDateTime;
    const last = signedIn === null ? null : Date.parse(signedIn);
    const lastActivation = day(now - (120 + unit(`${user.id}:activated`) * 500) * DAY);
    const named = {
      reportRefreshDate: refreshed,
      userPrincipalName: user.userPrincipalName,
      lastActivationDate: lastActivation,
    };
    // No successful sign-in for 90 days, so no activity in the report.
    if (last === null || now - last >= PERIOD * DAY)
      return {
        ...named,
        lastActivityDate: "",
        details: [flagsOf({ apps: new Set(), platforms: new Set() })],
      };
    const use = useOf(user, active);
    active += 1;
    return {
      ...named,
      // The report is two days behind, so nobody's last activity is later than that.
      lastActivityDate: day(Math.min(last, now - 2 * DAY)),
      details: [flagsOf(use)],
    };
  });
}

// The Microsoft 365 active users report (getOffice365ActiveUserDetail), the first of the CSV
// reports that Microsoft answers with a redirect to a download (ADR-0010). Columns and their order
// are verbatim from the v1.0 page, checked 2026-10-23; the "+" between product names and the
// True/False spelling are what third-party samples show, not what Microsoft documents.
export const ACTIVE_USER_COLUMNS = [
  "Report Refresh Date",
  "User Principal Name",
  "Display Name",
  "Is Deleted",
  "Deleted Date",
  "Has Exchange License",
  "Has OneDrive License",
  "Has SharePoint License",
  "Has Skype For Business License",
  "Has Yammer License",
  "Has Teams License",
  "Exchange Last Activity Date",
  "OneDrive Last Activity Date",
  "SharePoint Last Activity Date",
  "Skype For Business Last Activity Date",
  "Yammer Last Activity Date",
  "Teams Last Activity Date",
  "Exchange License Assign Date",
  "OneDrive License Assign Date",
  "SharePoint License Assign Date",
  "Skype For Business License Assign Date",
  "Yammer License Assign Date",
  "Teams License Assign Date",
  "Assigned Products",
] as const;
type ActiveUserColumn = (typeof ACTIVE_USER_COLUMNS)[number];
const ACTIVE_USERS = /^reports\/getOffice365ActiveUserDetail\(period='D(?:7|30|90|180)'\)$/;

const productNames: Record<string, string> = {
  SPB: "MICROSOFT 365 BUSINESS PREMIUM",
  EXCHANGESTANDARD: "EXCHANGE ONLINE (PLAN 1)",
  POWER_BI_PRO: "POWER BI PRO",
  AAD_PREMIUM_P2: "MICROSOFT ENTRA ID P2",
  FLOW_FREE: "MICROSOFT POWER AUTOMATE FREE",
};

// One CSV field, quoted when it holds a comma, a quote or a line break.
const field = (value: string) =>
  /[",\r\n]/.test(value) ? `"${value.replaceAll('"', '""')}"` : value;

// A report file as Microsoft most likely writes it: a byte-order mark, then CRLF lines.
function csvOf(rows: readonly (readonly string[])[]): string {
  return `\uFEFF${rows.map((row) => row.map(field).join(",")).join("\r\n")}\r\n`;
}

// ---------------------------------------------------------------------------------------------
// What each account did, for the Account activity audit
// ---------------------------------------------------------------------------------------------

const MIB = 1_048_576;
const GIB = 1_073_741_824;
// Microsoft's reports run behind: nobody's last activity is later than this many days ago.
const REPORT_LAG_DAYS = 2;

const keyOf = (name: string) => guid(`user:${name}`);
// Accounts the audit's stories are about, by roster key (scripts/demo/tenant.ts).
const SIGN_IN_AND_USE_NOTHING = new Set([keyOf("bokafor"), keyOf("icalloway")]);
const FILES_STILL_OPENED = keyOf("dokoro");
const LARGE_QUIET_MAILBOX = keyOf("rpemberton");
const SHARED_MAILBOX = keyOf("labresults");
// An unlicensed resource account the mailbox report lists as a room.
const ROOM_MAILBOX = keyOf("kiosk");
const GUEST_ON_SHAREPOINT = keyOf("awhitcombe");

// Mailboxes the Mailbox storage audit has something to say about, by roster key: the size in
// gibibytes, set against the usual quotas of 49, 49.5 and 50 GiB so that each lands on a different
// status. None of them belongs to an account the Account activity audit calls quiet, so its large
// mailbox view still holds one mailbox.
const MAILBOX_STORIES = new Map<
  string,
  { gib: number; items: number; archive: boolean; quotas?: "blank" }
>([
  // Full: nothing more arrives.
  [SHARED_MAILBOX, { gib: 50.3, items: 212_400, archive: false }],
  // Cannot send.
  [keyOf("mbrandt"), { gib: 49.7, items: 96_300, archive: true }],
  // Over the warning quota.
  [keyOf("dholloway"), { gib: 49.2, items: 88_150, archive: true }],
  // Near full, with an archive to move old mail to.
  [keyOf("isolheim"), { gib: 46.5, items: 71_900, archive: true }],
  // Near full and nowhere to move it.
  [keyOf("korne"), { gib: 41, items: 54_600, archive: false }],
  // The report gives no quotas.
  [keyOf("waris"), { gib: 8.2, items: 31_200, archive: false, quotas: "blank" }],
  [ROOM_MAILBOX, { gib: 0.4, items: 410, archive: false }],
]);

interface Activity {
  // Days since each workload was last used by the account's owner; null for none in 90 days.
  mail: number | null;
  teams: number | null;
  oneDrive: number | null;
  sharePoint: number | null;
  sent: number;
  received: number;
  read: number;
  chats: number;
  calls: number;
  meetings: number;
  files: number;
  sharePointFiles: number;
  pages: number;
  // The account's own OneDrive, which others may still open; null without one.
  drive: { lastDays: number | null; active: number; files: number; bytes: number } | null;
  mailbox: { items: number; bytes: number; archive: boolean } | null;
  teamsLicensed: boolean;
}

// What the reports say each account did, from its sign-ins: people who signed in recently used
// their workloads, people who went quiet used them until they did. The stories are planted by key.
function activitiesOf(tenant: SyntheticTenant): Map<string, Activity> {
  const now = tenant.now.getTime();
  const partNumber = new Map(tenant.skus.map((sku) => [sku.skuId, sku.skuPartNumber]));
  const result = new Map<string, Activity>();
  for (const user of tenant.users) {
    const parts = user.assignedLicenses.map((license) => partNumber.get(license.skuId) ?? "");
    const member = user.userType === "Member";
    const suite = member && parts.includes("SPB");
    const shared = user.id === SHARED_MAILBOX;
    const hasMailbox =
      shared ||
      user.id === ROOM_MAILBOX ||
      (member && (suite || parts.includes("EXCHANGESTANDARD")));
    const signedIn = user.signInActivity.lastSuccessfulSignInDateTime;
    const since = signedIn === null ? null : Math.floor((now - Date.parse(signedIn)) / DAY);
    const nothing = SIGN_IN_AND_USE_NOTHING.has(user.id);
    const lastUse = (workload: string, licensed: boolean, chance: number) => {
      if (nothing || !licensed || since === null || since >= PERIOD) return null;
      if (unit(`${user.id}:${workload}`) >= chance) return null;
      const days = Math.max(
        REPORT_LAG_DAYS,
        since + Math.floor(unit(`${user.id}:${workload}:lag`) * 4),
      );
      return days < PERIOD ? days : null;
    };
    const amount = (label: string, base: number, spread: number) =>
      base + Math.floor(unit(`${user.id}:${label}`) * spread);
    const story = MAILBOX_STORIES.get(user.id);
    const mail = shared ? 1 : lastUse("mail", hasMailbox, 1);
    const teams = lastUse("teams", suite, 0.85);
    const oneDrive = lastUse("onedrive", suite, 0.7);
    const sharePoint = user.id === GUEST_ON_SHAREPOINT ? 5 : lastUse("sharepoint", suite, 0.6);
    // A mailbox goes on receiving after its owner stops using it; only sending and reading count.
    const received =
      !hasMailbox || (!user.accountEnabled && !shared) ? 0 : amount("received", 20, 180);
    const files = oneDrive === null ? 0 : amount("files", 2, 60);
    const driveFiles = suite ? amount("drive-files", 20, 500) : 0;
    result.set(user.id, {
      mail,
      teams,
      oneDrive,
      sharePoint,
      sent: shared ? 0 : mail === null ? 0 : amount("sent", 5, 60),
      received: shared ? 140 : received,
      read: shared ? 22 : mail === null ? 0 : amount("read", 20, 120),
      chats: teams === null ? 0 : amount("chats", 3, 80),
      calls: teams === null ? 0 : amount("calls", 0, 12),
      meetings: teams === null ? 0 : amount("meetings", 0, 20),
      files,
      sharePointFiles:
        sharePoint === null ? 0 : user.id === GUEST_ON_SHAREPOINT ? 8 : amount("sp-files", 2, 40),
      pages: sharePoint === null ? 0 : user.id === GUEST_ON_SHAREPOINT ? 3 : amount("pages", 1, 30),
      drive: !suite
        ? null
        : user.id === FILES_STILL_OPENED
          ? { lastDays: 6, active: 14, files: 212, bytes: 212 * 9 * MIB }
          : {
              lastDays: oneDrive,
              active:
                oneDrive === null
                  ? 0
                  : Math.min(driveFiles, 1 + Math.floor(unit(`${user.id}:active`) * 30)),
              files: driveFiles,
              bytes: driveFiles * Math.round((2 + unit(`${user.id}:size`) * 38) * MIB),
            },
      mailbox: !hasMailbox
        ? null
        : user.id === LARGE_QUIET_MAILBOX
          ? { items: 61_240, bytes: Math.round(12.4 * GIB), archive: true }
          : story
            ? { items: story.items, bytes: Math.round(story.gib * GIB), archive: story.archive }
            : {
                items: amount("items", 200, 8_000),
                bytes: Math.round((0.1 + unit(`${user.id}:mailbox`) * 2.4) * GIB),
                archive: unit(`${user.id}:archive`) < 0.2,
              },
      teamsLicensed: suite,
    });
  }
  return result;
}

const profiles = new WeakMap<SyntheticTenant, Map<string, Activity>>();
function profilesOf(tenant: SyntheticTenant): Map<string, Activity> {
  let found = profiles.get(tenant);
  if (!found) {
    found = activitiesOf(tenant);
    profiles.set(tenant, found);
  }
  return found;
}

const daysAgo = (tenant: SyntheticTenant, days: number | null) =>
  days === null ? "" : day(tenant.now.getTime() - days * DAY);

// An account that was deleted since: the reports still name it and the directory does not.
const DEPARTED = "r.hollis@legacy.larkspur.example";

// A report from its columns and rows by column name; a column a row does not name is empty.
function fileOf(columns: readonly string[], rows: readonly Record<string, string>[]): string {
  return csvOf([columns, ...rows.map((row) => columns.map((name) => row[name] ?? ""))]);
}

// The columns of the other six reports, in the order of the v1.0 pages (checked 2026-10-23).
// Booleans are spelled True and False and the empty date is empty, as third-party samples show;
// Microsoft documents neither.
const columnsOf = (list: string) => list.split(",");
const EMAIL_COLUMNS = columnsOf(
  "Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Last Activity Date,Send Count,Receive Count,Read Count,Meeting Created Count,Meeting Interacted Count,Assigned Products,Report Period",
);
const TEAMS_COLUMNS = columnsOf(
  "Report Refresh Date,Tenant Display Name,Shared Channel Tenant Display Names,User Id,User Principal Name,Last Activity Date,Is Deleted,Deleted Date,Assigned Products,Team Chat Message Count,Private Chat Message Count,Call Count,Meeting Count,Post Messages,Reply Messages,Urgent Messages,Meetings Organized Count,Meetings Attended Count,Ad Hoc Meetings Organized Count,Ad Hoc Meetings Attended Count,Scheduled One-time Meetings Organized Count,Scheduled One-time Meetings Attended Count,Scheduled Recurring Meetings Organized Count,Scheduled Recurring Meetings Attended Count,Audio Duration,Video Duration,Screen Share Duration,Audio Duration In Seconds,Video Duration In Seconds,Screen Share Duration In Seconds,Has Other Action,Is Licensed,Report Period",
);
const ONEDRIVE_ACTIVITY_COLUMNS = columnsOf(
  "Report Refresh Date,User Principal Name,Is Deleted,Deleted Date,Last Activity Date,Viewed Or Edited File Count,Synced File Count,Shared Internally File Count,Shared Externally File Count,Assigned Products,Report Period",
);
const ONEDRIVE_USAGE_COLUMNS = columnsOf(
  "Report Refresh Date,Site URL,Owner Display Name,Is Deleted,Last Activity Date,File Count,Active File Count,Storage Used (Byte),Storage Allocated (Byte),Owner Principal Name,Report Period",
);
const SHAREPOINT_COLUMNS = columnsOf(
  "Report Refresh Date,User Principal Name,Is Deleted,Deleted Date,Last Activity Date,Viewed Or Edited File Count,Synced File Count,Shared Internally File Count,Shared Externally File Count,Visited Page Count,Assigned Products,Report Period",
);
const MAILBOX_COLUMNS = columnsOf(
  "Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Created Date,Last Activity Date,Item Count,Storage Used (Byte),Issue Warning Quota (Byte),Prohibit Send Quota (Byte),Prohibit Send/Receive Quota (Byte),Deleted Item Count,Deleted Item Size (Byte),Deleted Item Quota (Byte),Has Archive,Recipient Type,Report Period",
);
// The per-team report, in the order of the v1.0 page (checked 2026-10-24).
const TEAMS_TEAM_COLUMNS = columnsOf(
  "Report Refresh Date,Team Name,Team Id,Team Type,Last Activity Date,Report Period,Active Users,Active Channels,Guests,Reactions,Meetings Organized,Post Messages,Reply Messages,Channel Messages,Urgent Messages,Mentions,Active Shared Channels,Active External Users",
);
const EMAIL = /^reports\/getEmailActivityUserDetail\(period='D(?:7|30|90|180)'\)$/;
const TEAMS = /^reports\/getTeamsUserActivityUserDetail\(period='D(?:7|30|90|180)'\)$/;
const ONEDRIVE_ACTIVITY = /^reports\/getOneDriveActivityUserDetail\(period='D(?:7|30|90|180)'\)$/;
const ONEDRIVE_USAGE = /^reports\/getOneDriveUsageAccountDetail\(period='D(?:7|30|90|180)'\)$/;
const SHAREPOINT = /^reports\/getSharePointActivityUserDetail\(period='D(?:7|30|90|180)'\)$/;
const MAILBOX = /^reports\/getMailboxUsageDetail\(period='D(?:7|30|90|180)'\)$/;
const TEAMS_TEAM = /^reports\/getTeamsTeamActivityDetail\(period='D(?:7|30|90|180)'\)$/;

const PERIOD_COLUMN = String(PERIOD);
const refreshOf = (tenant: SyntheticTenant) => day(tenant.now.getTime() - REPORT_LAG_DAYS * DAY);
const productsOf = (tenant: SyntheticTenant, user: DirectoryUser) => {
  const partNumber = new Map(tenant.skus.map((sku) => [sku.skuId, sku.skuPartNumber]));
  return user.assignedLicenses
    .map((license) => productNames[partNumber.get(license.skuId) ?? ""] ?? "")
    .filter(Boolean)
    .join("+");
};
const count = (value: number) => String(value);

// The people each report lists: those with the service. A guest is listed only where the story
// puts one.
function emailOf(tenant: SyntheticTenant): string {
  const profile = profilesOf(tenant);
  const rows = tenant.users.flatMap((user) => {
    const use = profile.get(user.id);
    if (!use?.mailbox) return [];
    return [
      {
        "Report Refresh Date": refreshOf(tenant),
        "User Principal Name": user.userPrincipalName,
        "Display Name": user.displayName,
        "Is Deleted": "False",
        "Last Activity Date": daysAgo(tenant, use.mail),
        "Send Count": count(use.sent),
        "Receive Count": count(use.received),
        "Read Count": count(use.read),
        "Meeting Created Count": count(use.mail === null ? 0 : Math.floor(use.sent / 12)),
        "Meeting Interacted Count": count(use.mail === null ? 0 : Math.floor(use.read / 6)),
        "Assigned Products": productsOf(tenant, user),
        "Report Period": PERIOD_COLUMN,
      },
    ];
  });
  rows.push({
    "Report Refresh Date": refreshOf(tenant),
    "User Principal Name": DEPARTED,
    "Display Name": "R. Hollis",
    "Is Deleted": "True",
    "Last Activity Date": "",
    "Send Count": "0",
    "Receive Count": "0",
    "Read Count": "0",
    "Meeting Created Count": "0",
    "Meeting Interacted Count": "0",
    "Assigned Products": "",
    "Report Period": PERIOD_COLUMN,
  });
  return fileOf(EMAIL_COLUMNS, rows);
}

function teamsOf(tenant: SyntheticTenant): string {
  const profile = profilesOf(tenant);
  const rows = tenant.users.flatMap((user) => {
    const use = profile.get(user.id);
    if (!use?.teamsLicensed) return [];
    const team = use.teams === null ? 0 : Math.floor(use.chats * 0.35);
    return [
      {
        "Report Refresh Date": refreshOf(tenant),
        "Tenant Display Name": ORGANISATION,
        "User Id": user.id,
        "User Principal Name": user.userPrincipalName,
        "Last Activity Date": daysAgo(tenant, use.teams),
        "Is Deleted": "False",
        "Assigned Products": productsOf(tenant, user),
        "Team Chat Message Count": count(team),
        "Private Chat Message Count": count(use.chats - team),
        "Call Count": count(use.calls),
        "Meeting Count": count(use.meetings + Math.floor(use.meetings / 4)),
        "Meetings Attended Count": count(use.meetings),
        "Has Other Action":
          use.teams !== null && use.chats + use.calls + use.meetings === 0 ? "True" : "False",
        "Is Licensed": "Yes",
        "Report Period": PERIOD_COLUMN,
      },
    ];
  });
  return fileOf(TEAMS_COLUMNS, rows);
}

function oneDriveActivityOf(tenant: SyntheticTenant): string {
  const profile = profilesOf(tenant);
  const rows = tenant.users.flatMap((user) => {
    const use = profile.get(user.id);
    if (!use?.teamsLicensed) return [];
    return [
      {
        "Report Refresh Date": refreshOf(tenant),
        "User Principal Name": user.userPrincipalName,
        "Is Deleted": "False",
        "Last Activity Date": daysAgo(tenant, use.oneDrive),
        "Viewed Or Edited File Count": count(use.files),
        "Synced File Count": count(use.oneDrive === null ? 0 : Math.floor(use.files / 2)),
        "Shared Internally File Count": count(
          use.oneDrive === null ? 0 : Math.floor(use.files / 5),
        ),
        "Shared Externally File Count": "0",
        "Assigned Products": productsOf(tenant, user),
        "Report Period": PERIOD_COLUMN,
      },
    ];
  });
  return fileOf(ONEDRIVE_ACTIVITY_COLUMNS, rows);
}

function oneDriveUsageOf(tenant: SyntheticTenant): string {
  const profile = profilesOf(tenant);
  const rows = tenant.users.flatMap((user) => {
    const drive = profile.get(user.id)?.drive;
    if (!drive) return [];
    const owner = user.userPrincipalName.split("@")[0] ?? "";
    return [
      {
        "Report Refresh Date": refreshOf(tenant),
        "Site URL": `https://larkspur-my.sharepoint.example/personal/${owner.replace(/[^a-z0-9]+/gi, "_")}`,
        "Owner Display Name": user.displayName,
        "Is Deleted": "False",
        "Last Activity Date": daysAgo(tenant, drive.lastDays),
        "File Count": count(drive.files),
        "Active File Count": count(drive.active),
        "Storage Used (Byte)": count(drive.bytes),
        "Storage Allocated (Byte)": count(1_099_511_627_776),
        "Owner Principal Name": user.userPrincipalName,
        "Report Period": PERIOD_COLUMN,
      },
    ];
  });
  return fileOf(ONEDRIVE_USAGE_COLUMNS, rows);
}

function sharePointOf(tenant: SyntheticTenant): string {
  const profile = profilesOf(tenant);
  const rows = tenant.users.flatMap((user) => {
    const use = profile.get(user.id);
    if (!use?.teamsLicensed && user.id !== GUEST_ON_SHAREPOINT) return [];
    return [
      {
        "Report Refresh Date": refreshOf(tenant),
        "User Principal Name": user.userPrincipalName,
        "Is Deleted": "False",
        "Last Activity Date": daysAgo(tenant, use?.sharePoint ?? null),
        "Viewed Or Edited File Count": count(use?.sharePointFiles ?? 0),
        "Synced File Count": "0",
        "Shared Internally File Count": count(Math.floor((use?.sharePointFiles ?? 0) / 8)),
        "Shared Externally File Count": "0",
        "Visited Page Count": count(use?.pages ?? 0),
        "Assigned Products": productsOf(tenant, user),
        "Report Period": PERIOD_COLUMN,
      },
    ];
  });
  return fileOf(SHAREPOINT_COLUMNS, rows);
}

function mailboxOf(tenant: SyntheticTenant): string {
  const profile = profilesOf(tenant);
  const rows = tenant.users.flatMap((user) => {
    const use = profile.get(user.id);
    if (!use?.mailbox) return [];
    const blank = MAILBOX_STORIES.get(user.id)?.quotas === "blank";
    return [
      {
        "Report Refresh Date": refreshOf(tenant),
        "User Principal Name": user.userPrincipalName,
        "Display Name": user.displayName,
        "Is Deleted": "False",
        "Created Date": day(Date.parse(user.createdDateTime)),
        "Last Activity Date": daysAgo(tenant, use.mail),
        "Item Count": count(use.mailbox.items),
        "Storage Used (Byte)": count(use.mailbox.bytes),
        "Issue Warning Quota (Byte)": blank ? "" : count(49 * GIB),
        "Prohibit Send Quota (Byte)": blank ? "" : count(49.5 * GIB),
        "Prohibit Send/Receive Quota (Byte)": blank ? "" : count(50 * GIB),
        "Deleted Item Count": count(Math.floor(use.mailbox.items / 40)),
        "Deleted Item Size (Byte)": count(Math.floor(use.mailbox.bytes / 60)),
        "Deleted Item Quota (Byte)": count(30 * GIB),
        "Has Archive": use.mailbox.archive ? "True" : "False",
        "Recipient Type": recipientOf(user),
        "Report Period": PERIOD_COLUMN,
      },
    ];
  });
  return fileOf(MAILBOX_COLUMNS, rows);
}

// How the mailbox report names the kind of mailbox. The wording is a guess: Microsoft's page names
// the column and does not list its values.
const recipientOf = (user: DirectoryUser) =>
  user.id === SHARED_MAILBOX
    ? "SharedMailbox"
    : user.id === ROOM_MAILBOX
      ? "RoomMailbox"
      : "UserMailbox";

// The per-team Teams activity report behind the Groups and Teams audit: one row for each team in
// groups.ts that has used Teams, with the counts that report carries.
function teamsTeamOf(tenant: SyntheticTenant): string {
  const rows = TEAM_USAGE.map((team) => ({
    "Report Refresh Date": refreshOf(tenant),
    "Team Name": team.name,
    "Team Id": team.id,
    "Team Type": team.type,
    "Last Activity Date": daysAgo(tenant, team.lastDaysAgo),
    "Report Period": PERIOD_COLUMN,
    "Active Users": count(team.users),
    "Active Channels": count(team.channels),
    Guests: count(team.guests),
    Reactions: count(team.users * 6),
    "Meetings Organized": count(Math.floor(team.users / 3)),
    "Post Messages": count(team.users * 4),
    "Reply Messages": count(team.users * 7),
    "Channel Messages": count(team.users * 11),
    "Urgent Messages": count(Math.floor(team.users / 8)),
    Mentions: count(team.users * 2),
    "Active Shared Channels": "0",
    "Active External Users": count(team.guests),
  }));
  return fileOf(TEAMS_TEAM_COLUMNS, rows);
}

// The active users report (getOffice365ActiveUserDetail), the first of the CSV reports that
// Microsoft answers with a redirect to a download (ADR-0010), dated from what each account did.
function activeUsersOf(tenant: SyntheticTenant): string {
  const profile = profilesOf(tenant);
  const partNumber = new Map(tenant.skus.map((sku) => [sku.skuId, sku.skuPartNumber]));
  const flag = (value: boolean) => (value ? "True" : "False");
  const people = tenant.users.filter(
    (user) => user.userType === "Member" && user.assignedLicenses.length > 0,
  );
  const rows = people.map((user) => {
    const parts = user.assignedLicenses.map((license) => partNumber.get(license.skuId) ?? "");
    const suite = parts.includes("SPB");
    const exchange = suite || parts.includes("EXCHANGESTANDARD");
    const use = profile.get(user.id);
    // A little Yammer for some of the people who use the rest.
    const yammer =
      suite && use?.mail != null && unit(`${user.id}:yammer`) < 0.1 ? use.mail + 1 : null;
    const assigned = (licensed: boolean) => (licensed ? day(Date.parse(user.createdDateTime)) : "");
    const row: Record<ActiveUserColumn, string> = {
      "Report Refresh Date": refreshOf(tenant),
      "User Principal Name": user.userPrincipalName,
      "Display Name": user.displayName,
      "Is Deleted": "False",
      "Deleted Date": "",
      "Has Exchange License": flag(exchange),
      "Has OneDrive License": flag(suite),
      "Has SharePoint License": flag(suite),
      "Has Skype For Business License": "False",
      "Has Yammer License": flag(suite),
      "Has Teams License": flag(suite),
      "Exchange Last Activity Date": exchange ? daysAgo(tenant, use?.mail ?? null) : "",
      "OneDrive Last Activity Date": suite ? daysAgo(tenant, use?.oneDrive ?? null) : "",
      "SharePoint Last Activity Date": suite ? daysAgo(tenant, use?.sharePoint ?? null) : "",
      "Skype For Business Last Activity Date": "",
      "Yammer Last Activity Date": daysAgo(tenant, yammer),
      "Teams Last Activity Date": suite ? daysAgo(tenant, use?.teams ?? null) : "",
      "Exchange License Assign Date": assigned(exchange),
      "OneDrive License Assign Date": assigned(suite),
      "SharePoint License Assign Date": assigned(suite),
      "Skype For Business License Assign Date": "",
      "Yammer License Assign Date": assigned(suite),
      "Teams License Assign Date": assigned(suite),
      "Assigned Products": parts.map((part) => productNames[part] ?? part).join("+"),
    };
    return ACTIVE_USER_COLUMNS.map((name) => row[name]);
  });
  return csvOf([ACTIVE_USER_COLUMNS, ...rows]);
}

const csvReports: readonly [RegExp, (tenant: SyntheticTenant) => string][] = [
  [ACTIVE_USERS, activeUsersOf],
  [EMAIL, emailOf],
  [TEAMS, teamsOf],
  [ONEDRIVE_ACTIVITY, oneDriveActivityOf],
  [ONEDRIVE_USAGE, oneDriveUsageOf],
  [SHAREPOINT, sharePointOf],
  [MAILBOX, mailboxOf],
  [TEAMS_TEAM, teamsTeamOf],
];

// ---------------------------------------------------------------------------------------------
// Direct group memberships
// ---------------------------------------------------------------------------------------------

// Three of the groups of the synthetic tenant (groups.ts has them all). The licensing group is the
// one tenant.ts names as the source of the Business Premium license of accounts that get it
// through a group.
const ROLES = new Map<string, string[]>([
  [keyOf("admin"), ["Global Administrator"]],
  [keyOf("tvance"), ["Helpdesk Administrator"]],
  [keyOf("agray"), ["Exchange Administrator"]],
]);

// What GET /users/{id}/memberOf returns: groups and directory roles, each with the type Microsoft
// puts on every directory object. Direct memberships only.
function membershipsOf(user: DirectoryUser) {
  const groups = [
    ...(user.userType === "Member" && user.accountEnabled ? [STAFF] : []),
    ...(user.userType === "Member" && user.department === "Laboratory" ? [LABORATORY] : []),
    ...(user.licenseAssignmentStates.some((state) => state.assignedByGroup === LICENSING.id)
      ? [LICENSING]
      : []),
  ];
  return [
    ...groups.map((group) => ({
      "@odata.type": "#microsoft.graph.group",
      id: group.id,
      displayName: group.name,
    })),
    ...(ROLES.get(user.id) ?? []).map((role) => ({
      "@odata.type": "#microsoft.graph.directoryRole",
      id: guid(`role:${role}`),
      displayName: role,
    })),
  ];
}

const MEMBER_OF = /^users(?:\/([^/]+)|\('((?:[^']|'')*)'\))\/memberOf$/;

const built = new WeakMap<SyntheticTenant, object[]>();

// Microsoft 365 usage reports, the direct group memberships and the report privacy setting.
export const usageReports: DemoGraphExtension = (request, tools) => {
  if (request.path === "admin/reportSettings") return tools.json({ displayConcealedNames: false });
  for (const [pattern, build] of csvReports)
    if (pattern.test(request.path)) return tools.download(build(request.tenant));
  const owner = MEMBER_OF.exec(request.path);
  if (owner) {
    const principal = (owner[2] ?? "").replaceAll("''", "'").toLowerCase();
    const user = request.tenant.users.find((candidate) =>
      owner[1] !== undefined
        ? candidate.id === owner[1]
        : candidate.userPrincipalName.toLowerCase() === principal,
    );
    return user ? tools.page(request.url, membershipsOf(user), null) : tools.notFound();
  }
  if (request.path !== "reports/getM365AppUserDetail(period='D90')") return null;
  let report = built.get(request.tenant);
  if (!report) {
    report = reportOf(request.tenant);
    built.set(request.tenant, report);
  }
  return tools.page(request.url, report, null);
};

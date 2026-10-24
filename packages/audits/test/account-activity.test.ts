import { auditRowSchema, queryAuditRows } from "@prism365/contracts";
import { describe, expect, it } from "vitest";
import { activityOf, type DirectoryUser } from "../src/directory.ts";
import {
  ACTIVE_DAYS,
  type Availability,
  accountActivityAudit,
  auditDefinition,
  availabilityOf,
  evidenceOf,
  MAX_MEMBERSHIP_ACCOUNTS,
  MEMBERSHIP_CONCURRENCY,
  QUIET_DAYS,
  VERDICTS,
  verdictOf,
} from "../src/index.ts";
import {
  daysAgo,
  expectValidTable,
  fake,
  find,
  now,
  paged,
  type Route,
} from "./apps-devices-usage-support.ts";

// Synthetic accounts on a.example and b.example, with invented groups, roles and report files.
// Every report path answers 302 with a link on the download host, as Microsoft does, and the
// download host serves the CSV.

const audit = accountActivityAudit;
const REPORT_HOST = "reports.office.com";
const granted = [
  "User.Read.All",
  "AuditLog.Read.All",
  "Reports.Read.All",
  "ReportSettings.Read.All",
  "LicenseAssignment.Read.All",
];
const E3 = "11111111-0000-4000-8000-0000000000e3";
const BASIC = "11111111-0000-4000-8000-0000000000b1";
const LICENSING_GROUP = "22222222-0000-4000-8000-0000000000a1";
const BASIC_GROUP = "22222222-0000-4000-8000-0000000000a2";
const GONE_GROUP = "22222222-0000-4000-8000-0000000000a3";

// The headers of Microsoft's v1.0 pages, verbatim (checked 2026-10-23).
const columnsOf = (list: string) => list.split(",");
const headers = {
  activeUsers: columnsOf(
    "Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Has Exchange License,Has OneDrive License,Has SharePoint License,Has Skype For Business License,Has Yammer License,Has Teams License,Exchange Last Activity Date,OneDrive Last Activity Date,SharePoint Last Activity Date,Skype For Business Last Activity Date,Yammer Last Activity Date,Teams Last Activity Date,Exchange License Assign Date,OneDrive License Assign Date,SharePoint License Assign Date,Skype For Business License Assign Date,Yammer License Assign Date,Teams License Assign Date,Assigned Products",
  ),
  email: columnsOf(
    "Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Last Activity Date,Send Count,Receive Count,Read Count,Meeting Created Count,Meeting Interacted Count,Assigned Products,Report Period",
  ),
  teams: columnsOf(
    "Report Refresh Date,Tenant Display Name,Shared Channel Tenant Display Names,User Id,User Principal Name,Last Activity Date,Is Deleted,Deleted Date,Assigned Products,Team Chat Message Count,Private Chat Message Count,Call Count,Meeting Count,Post Messages,Reply Messages,Urgent Messages,Meetings Organized Count,Meetings Attended Count,Has Other Action,Is Licensed,Report Period",
  ),
  oneDriveActivity: columnsOf(
    "Report Refresh Date,User Principal Name,Is Deleted,Deleted Date,Last Activity Date,Viewed Or Edited File Count,Synced File Count,Shared Internally File Count,Shared Externally File Count,Assigned Products,Report Period",
  ),
  oneDriveUsage: columnsOf(
    "Report Refresh Date,Site URL,Owner Display Name,Is Deleted,Last Activity Date,File Count,Active File Count,Storage Used (Byte),Storage Allocated (Byte),Owner Principal Name,Report Period",
  ),
  sharePoint: columnsOf(
    "Report Refresh Date,User Principal Name,Is Deleted,Deleted Date,Last Activity Date,Viewed Or Edited File Count,Synced File Count,Shared Internally File Count,Shared Externally File Count,Visited Page Count,Assigned Products,Report Period",
  ),
  mailbox: columnsOf(
    "Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Created Date,Last Activity Date,Item Count,Storage Used (Byte),Issue Warning Quota (Byte),Prohibit Send Quota (Byte),Prohibit Send/Receive Quota (Byte),Deleted Item Count,Deleted Item Size (Byte),Deleted Item Quota (Byte),Has Archive,Report Period",
  ),
};
type ReportKey = keyof typeof headers;

const paths: Record<ReportKey, string> = {
  activeUsers: "/v1.0/reports/getOffice365ActiveUserDetail(period='D90')",
  email: "/v1.0/reports/getEmailActivityUserDetail(period='D90')",
  teams: "/v1.0/reports/getTeamsUserActivityUserDetail(period='D90')",
  oneDriveActivity: "/v1.0/reports/getOneDriveActivityUserDetail(period='D90')",
  oneDriveUsage: "/v1.0/reports/getOneDriveUsageAccountDetail(period='D90')",
  sharePoint: "/v1.0/reports/getSharePointActivityUserDetail(period='D90')",
  mailbox: "/v1.0/reports/getMailboxUsageDetail(period='D90')",
};
const reportKeys = Object.keys(paths) as ReportKey[];

const GIB = 1_073_741_824;
const dayOf = (days: number) =>
  new Date(now.getTime() - days * 86_400_000).toISOString().slice(0, 10);

const BYTE_ORDER_MARK = "\uFEFF";
const field = (value: string) =>
  /[",\r\n]/.test(value) ? `"${value.replaceAll('"', '""')}"` : value;
// A report file as Microsoft most likely writes it: a byte-order mark, then CRLF lines. A column
// a row does not name is empty.
function csv(
  header: readonly string[],
  rows: readonly Record<string, string>[],
  refreshed = dayOf(2),
): string {
  const lines = [
    header.join(","),
    ...rows.map((row) =>
      header
        .map((name) => field(name === "Report Refresh Date" ? refreshed : (row[name] ?? "")))
        .join(","),
    ),
  ];
  return `${BYTE_ORDER_MARK}${lines.join("\r\n")}\r\n`;
}

const sign = (days: number | null, extra: Record<string, string | null> = {}) =>
  days === null
    ? null
    : {
        lastSuccessfulSignInDateTime: daysAgo(days),
        lastSignInDateTime: daysAgo(days),
        lastNonInteractiveSignInDateTime: daysAgo(days),
        ...extra,
      };

interface Account {
  id: string;
  upn: string;
  name: string;
  signedIn: number | null;
  created?: number;
  licenses?: { sku: string; group?: string }[];
  states?: boolean;
  type?: "Member" | "Guest";
  enabled?: boolean;
  signInActivity?: unknown;
}
const account = (value: Account) => ({
  id: value.id,
  displayName: value.name,
  userPrincipalName: value.upn,
  userType: value.type ?? "Member",
  accountEnabled: value.enabled ?? true,
  createdDateTime: daysAgo(value.created ?? 900),
  assignedLicenses: (value.licenses ?? []).map((license) => ({ skuId: license.sku })),
  ...(value.states === false
    ? {}
    : {
        licenseAssignmentStates: (value.licenses ?? []).map((license) => ({
          skuId: license.sku,
          assignedByGroup: license.group ?? null,
        })),
      }),
  signInActivity: "signInActivity" in value ? value.signInActivity : sign(value.signedIn),
});

const accounts = [
  account({
    id: "u-active",
    upn: "ada@a.example",
    name: "Ada Active",
    signedIn: 5,
    licenses: [{ sku: E3 }],
  }),
  account({
    id: "u-uses-nothing",
    upn: "ben@a.example",
    name: "Ben Browses",
    signedIn: 29,
    licenses: [{ sku: E3, group: LICENSING_GROUP }],
  }),
  account({
    id: "U-Active-Work",
    upn: "cy@a.example",
    name: "Cy Chats",
    signedIn: 100,
    licenses: [{ sku: E3 }, { sku: BASIC, group: BASIC_GROUP }],
  }),
  account({ id: "u-recent-use", upn: "dee@a.example", name: "Dee Docs", signedIn: 10 }),
  account({
    id: "u-others",
    upn: "eli@a.example",
    name: "Eli Elsewhere",
    signedIn: 120,
    licenses: [{ sku: BASIC, group: GONE_GROUP }],
  }),
  account({
    id: "u-quiet",
    upn: "fay@a.example",
    name: "Fay Faded",
    signedIn: 120,
    licenses: [{ sku: E3 }],
  }),
  account({
    id: "u-quiet-90",
    upn: "gus@a.example",
    name: "Gus Gone",
    signedIn: 90,
    states: false,
    licenses: [{ sku: BASIC }],
  }),
  account({ id: "u-89", upn: "hal@a.example", name: "Hal Halfway", signedIn: 89 }),
  account({ id: "u-31", upn: "ivy@a.example", name: "Ivy Idle", signedIn: 31 }),
  account({ id: "u-work-31", upn: "jo@a.example", name: "Jo Jotting", signedIn: 100 }),
  account({ id: "u-never", upn: "kit@a.example", name: "Kit Keen", signedIn: null, created: 60 }),
  account({ id: "u-new", upn: "lee@a.example", name: "Lee Latest", signedIn: null, created: 20 }),
  account({
    id: "u-never-others",
    upn: "mia@a.example",
    name: "Mia Meadow",
    signedIn: null,
    created: 200,
  }),
  account({
    id: "u-attempt",
    upn: "ned@a.example",
    name: "Ned Nearly",
    signedIn: null,
    signInActivity: {
      lastSuccessfulSignInDateTime: null,
      lastSignInDateTime: daysAgo(20),
      lastNonInteractiveSignInDateTime: null,
    },
  }),
  account({ id: "u-own-drive", upn: "pat@a.example", name: "Pat Parcel", signedIn: 100 }),
  account({ id: "u-empty-drive", upn: "quin@a.example", name: "Quin Quiet", signedIn: 100 }),
  account({ id: "u-quiet-b", upn: "ruth@b.example", name: "Ruth Remote", signedIn: 150 }),
];
// Dates Microsoft could not have recorded make an account's sign-in unreadable, and the run partial.
// It is kept out of the shared list so that what else makes a run partial shows by itself.
const unreadable = account({
  id: "u-unreadable",
  upn: "opal@a.example",
  name: "Opal Odd",
  signedIn: null,
  signInActivity: {
    lastSuccessfulSignInDateTime: "0001-01-01T00:00:00Z",
    lastSignInDateTime: null,
    lastNonInteractiveSignInDateTime: null,
  },
});
const userIds = new Map(accounts.map((item) => [item.userPrincipalName, item.id]));
const idOf = (upn: string) => userIds.get(upn) ?? "";

// What the reports say. Names differ in case or compatibility form from the directory's on
// purpose; Teams names one account only by its object ID.
const files = {
  activeUsers: csv(headers.activeUsers, [
    {
      "User Principal Name": "ADA@A.EXAMPLE",
      "Is Deleted": "False",
      "Exchange Last Activity Date": dayOf(3),
      "Teams Last Activity Date": dayOf(4),
    },
    {
      "User Principal Name": "dee@a.example",
      "Is Deleted": "False",
      "SharePoint Last Activity Date": dayOf(45),
    },
    { "User Principal Name": "ghost@a.example", "Is Deleted": "False" },
  ]),
  email: csv(headers.email, [
    {
      "User Principal Name": "ａｄａ@a.example",
      "Last Activity Date": dayOf(3),
      "Send Count": "12",
      "Receive Count": "40",
      "Read Count": "30",
    },
    {
      "User Principal Name": "fay@a.example",
      "Last Activity Date": "",
      "Send Count": "0",
      "Receive Count": "300",
      "Read Count": "0",
    },
  ]),
  teams: csv(headers.teams, [
    {
      "User Id": idOf("cy@a.example").toLowerCase(),
      "User Principal Name": "cy.old@a.example",
      "Last Activity Date": dayOf(10),
      "Team Chat Message Count": "4",
      "Private Chat Message Count": "2",
      "Call Count": "1",
      "Meetings Attended Count": "3",
      "Has Other Action": "False",
    },
    { "User Principal Name": "jo@a.example", "Last Activity Date": dayOf(31) },
    { "User Principal Name": "ghost@a.example", "Last Activity Date": dayOf(5) },
  ]),
  oneDriveActivity: csv(headers.oneDriveActivity, [
    {
      "User Principal Name": "pat@a.example",
      "Last Activity Date": dayOf(40),
      "Viewed Or Edited File Count": "3",
    },
  ]),
  oneDriveUsage: csv(headers.oneDriveUsage, [
    {
      "Owner Principal Name": "ELI@a.example",
      "Last Activity Date": dayOf(12),
      "File Count": "120",
      "Active File Count": "7",
      "Storage Used (Byte)": String(2 * GIB),
    },
    {
      "Owner Principal Name": "mia@a.example",
      "Last Activity Date": dayOf(8),
      "File Count": "9",
      "Active File Count": "3",
    },
    {
      "Owner Principal Name": "pat@a.example",
      "Last Activity Date": dayOf(30),
      "File Count": "40",
      "Active File Count": "5",
    },
    {
      "Owner Principal Name": "quin@a.example",
      "Last Activity Date": "",
      "File Count": "10",
      "Active File Count": "0",
      "Storage Used (Byte)": String(Math.round(0.5 * GIB)),
    },
    { "Owner Principal Name": "ghost@a.example", "Active File Count": "9" },
  ]),
  sharePoint: csv(headers.sharePoint, [
    {
      "User Principal Name": "dee@a.example",
      "Last Activity Date": dayOf(45),
      "Viewed Or Edited File Count": "3",
      "Visited Page Count": "1",
    },
  ]),
  mailbox: csv(headers.mailbox, [
    {
      "User Principal Name": "ada@a.example",
      "Last Activity Date": dayOf(3),
      "Item Count": "1200",
      "Storage Used (Byte)": String(3 * GIB),
      "Has Archive": "True",
    },
    {
      "User Principal Name": "fay@a.example",
      "Last Activity Date": "",
      "Item Count": "90000",
      "Storage Used (Byte)": String(12 * GIB),
      "Has Archive": "False",
    },
  ]),
} satisfies Record<ReportKey, string>;

const member = (type: string, id: string, displayName: string) => ({
  "@odata.type": `#microsoft.graph.${type}`,
  id,
  displayName,
});
const memberships: Record<string, unknown[]> = {
  "u-active": [
    member("group", "g-1", "Staff"),
    member("group", "g-2", "Laboratory"),
    member("directoryRole", "r-1", "Helpdesk Administrator"),
    member("administrativeUnit", "au-1", "Head office"),
  ],
  "u-uses-nothing": [member("group", LICENSING_GROUP, "Licensing, Business Premium")],
  "U-Active-Work": [
    member("group", BASIC_GROUP, "Basic licensing"),
    member("group", "g-1", "Staff"),
  ],
  "u-quiet": [
    member("group", "g-1", "Staff"),
    member("directoryRole", "r-2", "Global Administrator"),
  ],
  "u-quiet-b": [member("directoryRole", "r-3", "User Administrator")],
};

// A response made fresh for each request: a Response cloned for a second request leaves the
// first one's body to be cancelled, and the client waits for that.
type Answer = () => Response;
const status =
  (code: number, headers: Record<string, string> = {}): Answer =>
  () =>
    new Response("{}", { status: code, headers });

interface World {
  users?: unknown[];
  // A report file, or the response for the report's path.
  reports?: Partial<Record<ReportKey, string | Answer>>;
  settings?: Answer;
  // Responses for the download host, by report.
  downloads?: Partial<Record<ReportKey, Answer>>;
  memberships?: Record<string, unknown[] | Answer>;
}

// A tenant: the directory, subscriptions, the report privacy setting, the seven reports behind
// redirects, and each account's memberships.
function tenant(world: World = {}): Route {
  const served = new Map<string, Answer>();
  let issued = 0;
  const reports = { ...files, ...world.reports };
  const memberOf = { ...memberships, ...world.memberships };
  return (url) => {
    if (url.hostname === REPORT_HOST) return served.get(url.pathname)?.();
    if (url.pathname === "/v1.0/users") return paged(url, world.users ?? accounts, 500);
    if (url.pathname === "/v1.0/subscribedSkus")
      return Response.json({
        value: [
          { skuId: E3, skuPartNumber: "SPE_E3" },
          { skuId: BASIC, skuPartNumber: "O365_BUSINESS_ESSENTIALS" },
        ],
      });
    if (url.pathname === "/v1.0/admin/reportSettings")
      return world.settings?.() ?? Response.json({ displayConcealedNames: false });
    for (const key of reportKeys) {
      if (url.pathname !== paths[key]) continue;
      const body = reports[key];
      if (typeof body === "function") return body();
      issued += 1;
      const link = `/data/download/${key}-${issued}`;
      served.set(
        link,
        world.downloads?.[key] ??
          (() => new Response(body, { headers: { "Content-Type": "application/octet-stream" } })),
      );
      return new Response(null, {
        status: 302,
        headers: { Location: `https://${REPORT_HOST}${link}` },
      });
    }
    const owner = /^\/v1\.0\/users\/([^/]+)\/memberOf$/.exec(url.pathname)?.[1];
    if (owner !== undefined) {
      const answer = memberOf[decodeURIComponent(owner)] ?? [];
      return typeof answer === "function" ? answer() : Response.json({ value: answer });
    }
    return undefined;
  };
}

const run = (route: Route = tenant(), options: { granted?: string[]; scope?: string[] } = {}) => {
  const world = fake(audit, route, {
    granted: options.granted ?? granted,
    scope: options.scope ? new Set(options.scope) : null,
  });
  return audit.run(world.context).then((result) => ({ ...result, requests: world.requests }));
};
const note = (result: { notes: string[] }) => result.notes.join("\n");
const verdictOfRow = (rows: { key: string; cells: Record<string, unknown> }[], key: string) =>
  find(rows, key).cells.verdict;
const reportRequests = (requests: URL[]) =>
  requests.filter((url) => url.pathname.startsWith("/v1.0/reports/"));

describe("Account activity", () => {
  it("publishes a valid table", async () => {
    const result = await expectValidTable(audit, tenant(), { granted });
    expect(result.rows).toHaveLength(accounts.length);
  });

  it("declares the sign-in packs as required and the usage packs as optional, with a scope and a handoff", () => {
    const definition = auditDefinition(audit);
    expect(definition).toMatchObject({
      id: "accounts.activity",
      title: "Account activity",
      category: "Accounts",
      permissions: ["User.Read.All", "AuditLog.Read.All"],
      optionalPermissions: [
        "Reports.Read.All",
        "ReportSettings.Read.All",
        "LicenseAssignment.Read.All",
      ],
      scoped: true,
      handoff: true,
    });
    expect(audit.packs).toEqual(["directory.users.read", "audit.signins.read"]);
    expect(audit.optionalPacks).toEqual([
      "reports.usage.read",
      "reports.settings.read",
      "tenant.licenses.read",
    ]);
  });

  it("keeps every help text and the description whole", () => {
    const definition = auditDefinition(audit);
    for (const column of definition.columns) {
      expect(
        audit.definition.columns.find((item) => item.key === column.key)?.help ?? null,
        column.key,
      ).toBe(column.help);
    }
    expect(definition.description.length).toBeLessThanOrEqual(500);
    expect(definition.columns.length).toBeLessThanOrEqual(40);
  });

  it("shows the verdict, the sign-in and the four workloads by default, and hides the detail", () => {
    const definition = auditDefinition(audit);
    const shown = definition.columns.filter((column) => !column.hidden).map((column) => column.key);
    expect(shown).toEqual([
      "person",
      "verdict",
      "status",
      "lastSuccess",
      "days",
      "mailLast",
      "teamsLast",
      "oneDriveLast",
      "sharePointLast",
      "licenseAssignment",
      "groups",
    ]);
  });

  describe("the seven reports", () => {
    it("reads each report once, from Microsoft's redirect, and each account's memberships once", async () => {
      const result = await run();
      const requested = reportRequests(result.requests).map(
        (url) => `${url.pathname}${url.search}`,
      );
      expect(requested.sort()).toEqual(Object.values(paths).sort());
      expect(result.requests.filter((url) => url.hostname === REPORT_HOST)).toHaveLength(7);
      const memberOf = result.requests.filter((url) => url.pathname.endsWith("/memberOf"));
      expect(memberOf).toHaveLength(accounts.length);
      for (const url of memberOf) expect(url.searchParams.get("$select")).toBe("id,displayName");
      expect(new Set(memberOf.map((url) => url.pathname)).size).toBe(accounts.length);
    });

    it("joins by sign-in name in any letter case and compatibility form, and Teams by object ID", async () => {
      const result = await run();
      // The active users report names her in capitals, the email report in full-width letters.
      expect(find(result.rows, "u-active").cells).toMatchObject({
        mailLast: `${dayOf(3)}T00:00:00.000Z`,
        mailSent: 12,
        mailReceived: 40,
        mailRead: 30,
        teamsLast: `${dayOf(4)}T00:00:00.000Z`,
        mailboxItems: 1200,
        mailboxStorage: 3,
        mailboxArchive: true,
      });
      // The Teams report knows this account by its object ID under another name.
      expect(find(result.rows, "U-Active-Work").cells).toMatchObject({
        teamsLast: `${dayOf(10)}T00:00:00.000Z`,
        teamsChats: 6,
        teamsCalls: 1,
        teamsMeetings: 3,
      });
      expect(find(result.rows, "U-Active-Work").cells.verdict).toBe("Active");
    });

    it("joins OneDrive usage by owner principal name, and the others by user principal name", async () => {
      const result = await run();
      expect(find(result.rows, "u-others").cells).toMatchObject({
        oneDriveAccountLast: `${dayOf(12)}T00:00:00.000Z`,
        oneDriveActiveFiles: 7,
        oneDriveFileCount: 120,
        oneDriveStorage: 2,
        oneDriveLast: null,
      });
      expect(find(result.rows, "u-own-drive").cells).toMatchObject({
        oneDriveLast: `${dayOf(40)}T00:00:00.000Z`,
        oneDriveFiles: 3,
      });
      expect(find(result.rows, "u-recent-use").cells).toMatchObject({
        sharePointLast: `${dayOf(45)}T00:00:00.000Z`,
        sharePointFiles: 3,
        sharePointPages: 1,
      });
    });

    it("shows the later of the two dates that cover each workload", async () => {
      const row = (name: string, extra: Record<string, string>) => ({
        "User Principal Name": "ada@a.example",
        ...extra,
        ...(name ? {} : {}),
      });
      const result = await run(
        tenant({
          reports: {
            activeUsers: csv(headers.activeUsers, [
              row("", {
                "Exchange Last Activity Date": dayOf(9),
                "Teams Last Activity Date": dayOf(2),
                "OneDrive Last Activity Date": dayOf(8),
                "SharePoint Last Activity Date": dayOf(6),
              }),
            ]),
            email: csv(headers.email, [row("", { "Last Activity Date": dayOf(3) })]),
            teams: csv(headers.teams, [row("", { "Last Activity Date": dayOf(5) })]),
            oneDriveActivity: csv(headers.oneDriveActivity, [
              row("", { "Last Activity Date": dayOf(4) }),
            ]),
            sharePoint: csv(headers.sharePoint, [row("", { "Last Activity Date": dayOf(10) })]),
          },
        }),
      );
      const at = (days: number) => `${dayOf(days)}T00:00:00.000Z`;
      expect(find(result.rows, "u-active").cells).toMatchObject({
        mailLast: at(3),
        teamsLast: at(2),
        oneDriveLast: at(4),
        sharePointLast: at(6),
      });
    });

    it("counts report rows that match no account once each, and shows none of them", async () => {
      const result = await run();
      // ghost@a.example is in four reports and counts once.
      expect(note(result)).toContain("1 name in the reports matches no account in the directory");
      expect(JSON.stringify(result.rows)).not.toContain("ghost");
      expect(result.rows).toHaveLength(accounts.length);
    });

    it("says how old the oldest report is", async () => {
      const result = await run(
        tenant({
          reports: {
            email: csv(headers.email, [{ "User Principal Name": "ada@a.example" }], dayOf(4)),
            teams: csv(headers.teams, [{ "User Principal Name": "jo@a.example" }], dayOf(3)),
          },
        }),
      );
      expect(note(result)).toContain(`the oldest report used was refreshed on ${dayOf(4)}`);
    });

    it("skips a row Microsoft marks deleted and keeps the first of two rows for one account", async () => {
      const result = await run(
        tenant({
          reports: {
            email: csv(headers.email, [
              {
                "User Principal Name": "ada@a.example",
                "Is Deleted": "True",
                "Last Activity Date": dayOf(1),
                "Send Count": "99",
              },
              {
                "User Principal Name": "ada@a.example",
                "Last Activity Date": dayOf(3),
                "Send Count": "12",
              },
              {
                "User Principal Name": "ada@a.example",
                "Last Activity Date": dayOf(2),
                "Send Count": "7",
              },
            ]),
          },
        }),
      );
      expect(find(result.rows, "u-active").cells.mailSent).toBe(12);
      // The deleted row and ghost@a.example match no live account.
      expect(note(result)).toContain("2 names in the reports match no account");
    });

    it("reads a column by its name, whatever its position", async () => {
      const shuffled = [...headers.email].reverse();
      const result = await run(
        tenant({
          reports: {
            email: csv(shuffled, [
              {
                "User Principal Name": "ada@a.example",
                "Last Activity Date": dayOf(3),
                "Send Count": "12",
              },
            ]),
          },
        }),
      );
      expect(find(result.rows, "u-active").cells).toMatchObject({
        mailSent: 12,
        mailLast: `${dayOf(3)}T00:00:00.000Z`,
      });
    });

    it("keeps a value it cannot read apart from none, and says so", async () => {
      const result = await run(
        tenant({
          reports: {
            email: csv(headers.email, [
              {
                "User Principal Name": "ada@a.example",
                "Last Activity Date": dayOf(3),
                "Send Count": "many",
                "Read Count": "30",
              },
            ]),
          },
        }),
      );
      expect(find(result.rows, "u-active").cells).toMatchObject({ mailSent: null, mailRead: 30 });
      expect(note(result)).toContain("1 value in the reports could not be read as a number");
      expect(result.complete).toBe(false);
    });
  });

  describe("verdicts", () => {
    it("gives every account the verdict its sign-ins and reports earn", async () => {
      const result = await run();
      const verdicts = Object.fromEntries(result.rows.map((row) => [row.key, row.cells.verdict]));
      expect(verdicts).toEqual({
        "u-active": "Active",
        "u-uses-nothing": "Signs in, uses nothing",
        "U-Active-Work": "Active",
        "u-recent-use": "Active",
        "u-others": "Quiet, others use their files",
        "u-quiet": "Quiet everywhere",
        "u-quiet-90": "Quiet everywhere",
        "u-89": "Recently quiet",
        "u-31": "Recently quiet",
        "u-work-31": "Recently quiet",
        "u-never": "Never signed in",
        "u-new": "Not enough data",
        "u-never-others": "Quiet, others use their files",
        "u-attempt": "Not enough data",
        "u-own-drive": "Recently quiet",
        "u-empty-drive": "Quiet everywhere",
        "u-quiet-b": "Quiet everywhere",
      });
      expect(new Set(Object.values(verdicts))).toEqual(new Set(VERDICTS));
    });

    it("tags each row with its verdict and opens each view on its accounts", async () => {
      const result = await run();
      const definition = auditDefinition(audit);
      const keys = (view: string) =>
        queryAuditRows(definition, result.rows, { view })
          .map((row) => row.key)
          .sort();
      expect(keys("quiet")).toEqual(["u-empty-drive", "u-quiet", "u-quiet-90", "u-quiet-b"]);
      expect(keys("others-use-files")).toEqual(["u-never-others", "u-others"]);
      expect(keys("uses-nothing")).toEqual(["u-uses-nothing"]);
      expect(keys("recently-quiet")).toEqual(["u-31", "u-89", "u-own-drive", "u-work-31"]);
      expect(keys("never")).toEqual(["u-never"]);
      expect(keys("active")).toEqual(["U-Active-Work", "u-active", "u-recent-use"]);
      expect(keys("no-data")).toEqual(["u-attempt", "u-new"]);
      expect(keys("all")).toHaveLength(accounts.length);
    });

    it("draws the line at 30 days for a sign-in and for use", async () => {
      // 29 days is active, 30 is not; the account that signed in at 29 days and used nothing is the
      // near miss for Signs in, uses nothing, and the one at 31 days for Recently quiet.
      const result = await run();
      expect(verdictOfRow(result.rows, "u-uses-nothing")).toBe("Signs in, uses nothing");
      expect(verdictOfRow(result.rows, "u-31")).toBe("Recently quiet");
      // Teams use 31 days ago does not make a 100-day-old sign-in active; 10 days ago does.
      expect(verdictOfRow(result.rows, "u-work-31")).toBe("Recently quiet");
      expect(verdictOfRow(result.rows, "U-Active-Work")).toBe("Active");
    });

    it("draws the line at 90 days for being quiet", async () => {
      const result = await run();
      expect(verdictOfRow(result.rows, "u-quiet-90")).toBe("Quiet everywhere");
      expect(verdictOfRow(result.rows, "u-89")).toBe("Recently quiet");
    });

    it("does not call an account quiet when its own OneDrive shows its owner's use", async () => {
      const result = await run();
      // pat has OneDrive files someone opens, and also used OneDrive 40 days ago herself.
      expect(verdictOfRow(result.rows, "u-own-drive")).toBe("Recently quiet");
      // quin's OneDrive has a row with no activity at all.
      expect(verdictOfRow(result.rows, "u-empty-drive")).toBe("Quiet everywhere");
    });

    it("counts a sign-in with use in the last 90 days as active, not as using nothing", async () => {
      const result = await run();
      expect(verdictOfRow(result.rows, "u-recent-use")).toBe("Active");
    });

    it("calls an account never signed in only after 30 days, and lets use or shared files outrank it", async () => {
      const result = await run();
      expect(verdictOfRow(result.rows, "u-never")).toBe("Never signed in");
      expect(verdictOfRow(result.rows, "u-new")).toBe("Not enough data");
      expect(verdictOfRow(result.rows, "u-never-others")).toBe("Quiet, others use their files");
    });

    it("is not enough data for attempts with no success", async () => {
      const result = await run();
      expect(verdictOfRow(result.rows, "u-attempt")).toBe("Not enough data");
      expect(find(result.rows, "u-attempt").cells.days).toBe(20);
      expect(result.complete).toBe(true);
    });

    it("is not enough data for dates it cannot read, which makes the run partial", async () => {
      const result = await run(tenant({ users: [...accounts, unreadable] }));
      expect(find(result.rows, "u-unreadable").cells).toMatchObject({
        verdict: "Not enough data",
        days: null,
      });
      expect(note(result)).toContain("1 account has sign-in dates Prism365 could not read");
      expect(result.complete).toBe(false);
    });

    describe("as a rule", () => {
      const sinceSignIn = (days: number | null): DirectoryUser =>
        ({
          createdDateTime: daysAgo(900),
          signInActivity:
            days === null
              ? null
              : {
                  lastSuccessfulSignInDateTime: daysAgo(days),
                  lastSignInDateTime: daysAgo(days),
                  lastNonInteractiveSignInDateTime: null,
                },
        }) as DirectoryUser;
      const everything: Availability = { workloads: true, drive: true };
      const verdict = (
        days: number | null,
        facts: Parameters<typeof evidenceOf>[0],
        availability: Availability = everything,
      ) => verdictOf(activityOf(sinceSignIn(days), now), evidenceOf(facts, availability, now));

      it("uses the same thresholds the documentation gives", () => {
        expect([ACTIVE_DAYS, QUIET_DAYS]).toEqual([30, 90]);
      });

      it("decides on each side of every boundary", () => {
        expect(verdict(29, undefined)).toBe("Signs in, uses nothing");
        expect(verdict(30, undefined)).toBe("Recently quiet");
        expect(verdict(89, undefined)).toBe("Recently quiet");
        expect(verdict(90, undefined)).toBe("Quiet everywhere");
        expect(verdict(100, { mailLast: dayOf(29) })).toBe("Active");
        expect(verdict(100, { mailLast: dayOf(30) })).toBe("Recently quiet");
        expect(verdict(100, { mailLast: dayOf(90) })).toBe("Quiet everywhere");
        expect(verdict(100, { mailLast: dayOf(89) })).toBe("Recently quiet");
      });

      it("treats a count above zero as use, and received mail as none", () => {
        expect(verdict(120, { mailSent: 1 })).toBe("Recently quiet");
        expect(verdict(120, { mailRead: 1 })).toBe("Recently quiet");
        expect(verdict(120, { teamsChats: 1 })).toBe("Recently quiet");
        expect(verdict(120, { teamsCalls: 1 })).toBe("Recently quiet");
        expect(verdict(120, { teamsMeetings: 1 })).toBe("Recently quiet");
        expect(verdict(120, { teamsOther: true })).toBe("Recently quiet");
        expect(verdict(120, { oneDriveFiles: 1 })).toBe("Recently quiet");
        expect(verdict(120, { sharePointFiles: 1 })).toBe("Recently quiet");
        expect(verdict(120, { sharePointPages: 1 })).toBe("Recently quiet");
        expect(verdict(120, { mailReceived: 500 })).toBe("Quiet everywhere");
        expect(verdict(120, { mailSent: 0, mailRead: 0, teamsOther: false })).toBe(
          "Quiet everywhere",
        );
      });

      it("takes any dated source as use", () => {
        for (const key of [
          "activeExchange",
          "activeTeams",
          "activeOneDrive",
          "activeSharePoint",
          "mailLast",
          "teamsLast",
          "oneDriveLast",
          "sharePointLast",
          "mailboxLast",
        ] as const)
          expect(verdict(120, { [key]: dayOf(10) }), key).toBe("Active");
      });

      it("takes shared files from the account's own OneDrive, by activity or by active files", () => {
        expect(verdict(120, { driveLast: dayOf(10) })).toBe("Quiet, others use their files");
        expect(verdict(120, { driveActive: 1 })).toBe("Quiet, others use their files");
        expect(verdict(120, { driveLast: dayOf(90) })).toBe("Quiet everywhere");
        expect(verdict(120, { driveActive: 0, driveFiles: 400 })).toBe("Quiet everywhere");
      });

      it("never says an account is quiet when a report that could show use was not read", () => {
        // The active users report still dates OneDrive use, so the workload has a source.
        const withoutOneDriveActivity = availabilityOf(
          new Set(["activeUsers", "email", "teams", "sharePoint", "mailbox", "oneDriveUsage"]),
        );
        expect(withoutOneDriveActivity.workloads).toBe(true);
        expect(verdict(120, undefined, withoutOneDriveActivity)).toBe("Quiet everywhere");
        const noWorkloadSource = availabilityOf(
          new Set(["email", "mailbox", "teams", "sharePoint", "oneDriveUsage"]),
        );
        expect(noWorkloadSource.workloads).toBe(false);
        expect(verdict(120, undefined, noWorkloadSource)).toBe("Not enough data");
        expect(verdict(120, undefined, { workloads: true, drive: false })).toBe("Not enough data");
        expect(verdict(45, undefined, { workloads: false, drive: true })).toBe("Not enough data");
        expect(verdict(10, undefined, { workloads: false, drive: false })).toBe("Active");
        expect(verdict(10, { mailSent: 1 }, { workloads: true, drive: true })).toBe("Active");
      });

      it("needs every workload to have a source", () => {
        const sources = {
          mail: ["email", "mailbox", "activeUsers"],
          teams: ["teams", "activeUsers"],
          oneDrive: ["oneDriveActivity", "activeUsers"],
          sharePoint: ["sharePoint", "activeUsers"],
        } as const;
        const all = new Set([
          "activeUsers",
          "email",
          "teams",
          "oneDriveActivity",
          "oneDriveUsage",
          "sharePoint",
          "mailbox",
        ] as const);
        expect(availabilityOf(all)).toEqual({ workloads: true, drive: true });
        for (const ids of Object.values(sources)) {
          const without = new Set(all);
          for (const id of ids) without.delete(id);
          expect(availabilityOf(without).workloads, ids.join()).toBe(false);
        }
        expect(availabilityOf(new Set(["activeUsers"]))).toEqual({ workloads: true, drive: false });
        expect(availabilityOf(new Set())).toEqual({ workloads: false, drive: false });
      });

      it("makes an unreadable report date undecidable, not none", () => {
        expect(verdict(120, { mailLast: "yesterday" })).toBe("Not enough data");
        expect(verdict(120, { driveLast: "soon" })).toBe("Not enough data");
        expect(verdict(10, { mailLast: "yesterday" })).toBe("Active");
        expect(verdict(120, { mailLast: "yesterday", teamsLast: dayOf(5) })).toBe("Active");
      });
    });
  });

  describe("licenses and groups", () => {
    it("tells direct, group, both and none apart, with the groups' names", async () => {
      const result = await run();
      expect(find(result.rows, "u-active").cells).toMatchObject({
        licenses: ["Microsoft 365 E3"],
        licenseAssignment: "Direct",
        licenseGroups: [],
      });
      expect(find(result.rows, "u-uses-nothing").cells).toMatchObject({
        licenseAssignment: "Through groups",
        licenseGroups: ["Licensing, Business Premium"],
      });
      expect(find(result.rows, "U-Active-Work").cells).toMatchObject({
        licenseAssignment: "Both",
        licenseGroups: ["Basic licensing"],
      });
      expect(find(result.rows, "u-recent-use").cells).toMatchObject({
        licenses: [],
        licenseAssignment: "None",
      });
    });

    it("names a group that is not among the account's direct memberships only as another group", async () => {
      const result = await run();
      expect(find(result.rows, "u-others").cells).toMatchObject({
        licenseAssignment: "Through groups",
        licenseGroups: ["Another group"],
      });
    });

    it("says Unknown when Microsoft did not give the assignment states", async () => {
      const result = await run();
      expect(find(result.rows, "u-quiet-90").cells.licenseAssignment).toBe("Unknown");
    });

    it("lists the accounts whose licenses all come through groups", async () => {
      const result = await run();
      const definition = auditDefinition(audit);
      expect(
        queryAuditRows(definition, result.rows, { view: "group-only" })
          .map((row) => row.key)
          .sort(),
      ).toEqual(["u-others", "u-uses-nothing"]);
      expect(result.summary.find((tile) => tile.label === "Licensed through groups only")).toEqual({
        label: "Licensed through groups only",
        value: 2,
        tone: "neutral",
        view: "group-only",
      });
    });

    it("counts groups, directory roles and administrative units apart", async () => {
      const result = await run();
      expect(find(result.rows, "u-active").cells).toMatchObject({
        groups: 2,
        groupNames: ["Staff", "Laboratory"],
        roles: 1,
        roleNames: ["Helpdesk Administrator"],
        adminUnits: 1,
      });
      expect(find(result.rows, "u-quiet").cells).toMatchObject({
        groups: 1,
        roles: 1,
        adminUnits: 0,
      });
      expect(find(result.rows, "u-31").cells).toMatchObject({ groups: 0, roles: 0, adminUnits: 0 });
    });

    it("follows no nested groups: it asks for memberOf and never transitiveMemberOf", async () => {
      const result = await run();
      expect(result.requests.some((url) => url.pathname.includes("transitive"))).toBe(false);
      expect(note(result)).toContain("Group memberships are direct only");
    });

    it("leaves the group columns Not read for an account that is gone, and says so", async () => {
      const result = await run(tenant({ memberships: { "u-31": status(404) } }));
      expect(find(result.rows, "u-31").cells).toMatchObject({
        groups: null,
        groupNames: null,
        roles: null,
        licenseGroups: null,
      });
      expect(find(result.rows, "u-89").cells.groups).toBe(0);
      expect(note(result)).toContain("1 account had no answer for group memberships");
      expect(result.complete).toBe(false);
    });

    it("keeps what it read when Microsoft asks it to slow down, and says so", async () => {
      // The quiet accounts are asked for first, so these are the ones that get through.
      const first = new Set([
        "u-others",
        "u-quiet",
        "u-quiet-90",
        "u-never-others",
        "u-empty-drive",
        "u-quiet-b",
      ]);
      const throttled: Record<string, Answer> = {};
      for (const item of accounts)
        if (!first.has(item.id)) throttled[item.id] = status(429, { "Retry-After": "600" });
      const result = await run(tenant({ memberships: throttled }));
      expect(note(result)).toContain("Microsoft asked Prism365 to slow down");
      expect(result.complete).toBe(false);
      expect(find(result.rows, "u-quiet").cells.groups).toBe(1);
      expect(find(result.rows, "u-active").cells.groups).toBeNull();
      expect(result.rows).toHaveLength(accounts.length);
      // Once Microsoft says to slow down, only the requests already in flight are left to finish.
      const asked = result.requests.filter((url) => url.pathname.endsWith("/memberOf"));
      expect(asked.length).toBeLessThanOrEqual(first.size + MEMBERSHIP_CONCURRENCY);
      expect(asked.length).toBeLessThan(accounts.length);
    });

    it("keeps the rest of the result when Microsoft refuses memberships, naming the status", async () => {
      const refused: Record<string, Answer> = {};
      for (const item of accounts) refused[item.id] = status(403);
      const result = await run(tenant({ memberships: refused }));
      expect(note(result)).toContain(
        "Microsoft refused to list group memberships (HTTP 403), so none were read.",
      );
      expect(result.rows).toHaveLength(accounts.length);
      expect(verdictOfRow(result.rows, "u-quiet")).toBe("Quiet everywhere");
      expect(find(result.rows, "u-quiet").cells.groups).toBeNull();
      expect(result.complete).toBe(false);
      const asked = result.requests.filter((url) => url.pathname.endsWith("/memberOf"));
      expect(asked.length).toBeLessThanOrEqual(MEMBERSHIP_CONCURRENCY);
    });

    it("counts a membership it cannot read, and a type it does not know, as unread", async () => {
      const result = await run(
        tenant({
          memberships: {
            "u-31": [
              member("group", "g-1", "Staff"),
              member("device", "d-1", "Laptop"),
              { displayName: "no id" },
            ],
          },
        }),
      );
      expect(find(result.rows, "u-31").cells).toMatchObject({ groups: 1, roles: 0 });
      expect(note(result)).toContain("1 account has memberships Prism365 could not read");
      expect(result.complete).toBe(false);
    });

    it("cleans names and bounds their length", async () => {
      const long = "x".repeat(900);
      const result = await run(
        tenant({
          memberships: {
            "u-31": [
              member("group", "g-1", `Staff\u0000${long}`),
              member("directoryRole", "r-1", ""),
            ],
          },
        }),
      );
      const cells = find(result.rows, "u-31").cells;
      const groupNames = cells.groupNames as string[];
      expect(groupNames[0]?.includes("\u0000")).toBe(false);
      expect(groupNames[0]?.length).toBeLessThanOrEqual(200);
      expect(cells.roleNames).toEqual(["Unnamed role"]);
    });
  });

  describe("directory roles", () => {
    it("raises the critical tile for a quiet account that holds a role, and links to its view", async () => {
      const result = await run();
      const definition = auditDefinition(audit);
      expect(result.summary).toContainEqual({
        label: "Quiet everywhere, with directory roles",
        value: 2,
        tone: "critical",
        view: "quiet-roles",
      });
      expect(
        queryAuditRows(definition, result.rows, { view: "quiet-roles" })
          .map((row) => row.key)
          .sort(),
      ).toEqual(["u-quiet", "u-quiet-b"]);
      expect(
        queryAuditRows(definition, result.rows, { view: "roles" })
          .map((row) => row.key)
          .sort(),
      ).toEqual(["u-active", "u-quiet", "u-quiet-b"]);
    });

    it("never calls the tile good when it is zero", async () => {
      const result = await run(tenant({ memberships: { "u-quiet": [], "u-quiet-b": [] } }));
      const tile = result.summary.find((item) => item.label.startsWith("Quiet everywhere, with"));
      expect(tile).toMatchObject({ value: 0, tone: "neutral" });
    });
  });

  describe("the tiles", () => {
    it("count the accounts in each view", async () => {
      const result = await run();
      expect(Object.fromEntries(result.summary.map((tile) => [tile.label, tile.value]))).toEqual({
        Accounts: accounts.length,
        "Quiet everywhere, with directory roles": 2,
        "Quiet everywhere": 4,
        "Quiet, others use their files": 2,
        "Signs in, uses nothing": 1,
        "Licensed through groups only": 2,
      });
      const tones = Object.fromEntries(result.summary.map((tile) => [tile.label, tile.tone]));
      expect(tones).toMatchObject({
        "Quiet everywhere": "warning",
        "Quiet, others use their files": "warning",
        "Signs in, uses nothing": "neutral",
        "Licensed through groups only": "neutral",
      });
    });

    it("does not call a zero a warning", async () => {
      const result = await run(tenant({ users: [accounts[0]] }));
      const tones = Object.fromEntries(result.summary.map((tile) => [tile.label, tile.tone]));
      expect(tones["Quiet everywhere"]).toBe("neutral");
      expect(tones["Quiet, others use their files"]).toBe("neutral");
    });

    it("lists a large mailbox only on an account that is quiet everywhere", async () => {
      const result = await run();
      const definition = auditDefinition(audit);
      expect(
        queryAuditRows(definition, result.rows, { view: "large-mailbox" }).map((row) => row.key),
      ).toEqual(["u-quiet"]);
      expect(find(result.rows, "u-quiet").cells.mailboxStorage).toBe(12);
      // Ada's mailbox is small and she is active; a big mailbox on an active account is not listed.
      const big = await run(
        tenant({
          reports: {
            mailbox: csv(headers.mailbox, [
              {
                "User Principal Name": "ada@a.example",
                "Last Activity Date": dayOf(3),
                "Storage Used (Byte)": String(40 * GIB),
              },
              { "User Principal Name": "fay@a.example", "Storage Used (Byte)": String(10 * GIB) },
              {
                "User Principal Name": "gus@a.example",
                "Storage Used (Byte)": String(10 * GIB + 1),
              },
            ]),
          },
        }),
      );
      expect(
        queryAuditRows(definition, big.rows, { view: "large-mailbox" }).map((row) => row.key),
      ).toEqual(["u-quiet-90"]);
    });
  });

  describe("when a report cannot be read", () => {
    it("leaves its columns without data, marks the run partial and names the report and why", async () => {
      const result = await run(tenant({ reports: { email: status(403) } }));
      expect(result.complete).toBe(false);
      expect(note(result)).toContain(
        "The email activity report could not be read: Microsoft refused (HTTP 403).",
      );
      for (const row of result.rows)
        expect(row.cells, row.key).toMatchObject({
          mailSent: null,
          mailReceived: null,
          mailRead: null,
        });
      // The active users report still dates mail, so Ada's mail date survives.
      expect(find(result.rows, "u-active").cells.mailLast).toBe(`${dayOf(3)}T00:00:00.000Z`);
      // Everything else was read, and the run did not fail.
      expect(find(result.rows, "u-active").cells.teamsLast).toBe(`${dayOf(4)}T00:00:00.000Z`);
      expect(result.rows).toHaveLength(accounts.length);
      expect(verdictOfRow(result.rows, "u-quiet")).toBe("Quiet everywhere");
    });

    it("turns every verdict that says nothing was there into Not enough data when no source is left", async () => {
      const result = await run(
        tenant({
          reports: {
            activeUsers: status(403),
            oneDriveActivity: status(403),
          },
        }),
      );
      expect(verdictOfRow(result.rows, "u-quiet")).toBe("Not enough data");
      expect(verdictOfRow(result.rows, "u-uses-nothing")).toBe("Active");
      expect(verdictOfRow(result.rows, "u-31")).toBe("Not enough data");
      expect(verdictOfRow(result.rows, "u-others")).toBe("Not enough data");
      expect(verdictOfRow(result.rows, "u-never")).toBe("Never signed in");
      expect(verdictOfRow(result.rows, "U-Active-Work")).toBe("Active");
      // Both refused for the same reason, so one note names both.
      expect(note(result)).toContain(
        "The active users and OneDrive activity reports could not be read: Microsoft refused (HTTP 403).",
      );
      expect(result.summary.find((tile) => tile.label === "Quiet everywhere")?.value).toBe(0);
    });

    it("cannot call an account quiet without the OneDrive usage report, which shows who depends on it", async () => {
      const result = await run(tenant({ reports: { oneDriveUsage: status(403) } }));
      expect(verdictOfRow(result.rows, "u-quiet")).toBe("Not enough data");
      expect(verdictOfRow(result.rows, "u-others")).toBe("Not enough data");
      expect(verdictOfRow(result.rows, "u-uses-nothing")).toBe("Active");
      expect(find(result.rows, "u-others").cells.oneDriveActiveFiles).toBeNull();
    });

    it("explains a throttled report without failing the audit", async () => {
      const result = await run(
        tenant({
          reports: {
            teams: status(429, { "Retry-After": "600" }),
          },
        }),
      );
      expect(note(result)).toContain(
        "The Teams activity report could not be read: Microsoft is limiting how often reports can be read",
      );
      expect(result.complete).toBe(false);
      expect(result.rows).toHaveLength(accounts.length);
    });

    it("explains a download link on a host it does not know, naming the host", async () => {
      const result = await run(
        tenant({
          reports: {
            mailbox: () =>
              new Response(null, {
                status: 302,
                headers: { Location: "https://downloads.invalid/data/download/abc?token=secret" },
              }),
          },
        }),
      );
      const text = note(result);
      expect(text).toContain("The mailbox usage report could not be read");
      expect(text).toContain("a host Prism365 does not recognise (downloads.invalid)");
      expect(text).not.toContain("secret");
      expect(result.complete).toBe(false);
    });

    it("explains a download that failed, without the link", async () => {
      const result = await run(
        tenant({ downloads: { sharePoint: () => new Response("gone", { status: 410 }) } }),
      );
      const text = note(result);
      expect(text).toContain(
        "The SharePoint activity report could not be read: the download link failed.",
      );
      expect(text).not.toContain("/data/download/");
      expect(find(result.rows, "u-recent-use").cells.sharePointLast).toBe(
        `${dayOf(45)}T00:00:00.000Z`,
      );
    });

    it("explains a file it cannot parse, and a report without the columns it needs", async () => {
      const result = await run(
        tenant({
          reports: {
            oneDriveActivity: 'User Principal Name,Last Activity Date\r\n"unterminated',
            sharePoint: csv(["Report Refresh Date", "User Principal Name"], []),
          },
        }),
      );
      const text = note(result);
      expect(text).toContain(
        "The OneDrive activity report could not be read: its file could not be read safely (csv).",
      );
      expect(text).toContain(
        'The SharePoint activity report could not be read: it has no "Last Activity Date" column.',
      );
    });

    it("counts rows the reader skipped, and says an account may show less than it had", async () => {
      const broken = `${csv(headers.email, [
        { "User Principal Name": "ada@a.example", "Last Activity Date": dayOf(3) },
      ])}only,three,fields\r\n`;
      const result = await run(tenant({ reports: { email: broken } }));
      expect(note(result)).toContain("1 row in the reports could not be read and is left out");
      expect(find(result.rows, "u-active").cells.mailLast).toBe(`${dayOf(3)}T00:00:00.000Z`);
      expect(result.complete).toBe(false);
    });

    it("lets a session that ended through, for the runner to explain", async () => {
      await expect(run(tenant({ reports: { email: status(401) } }))).rejects.toMatchObject({
        status: 401,
      });
    });
  });

  describe("when names are concealed", () => {
    it("reads no report and joins nothing when the setting says so, and tells a Global Administrator where to turn it off", async () => {
      const result = await run(
        tenant({ settings: () => Response.json({ displayConcealedNames: true }) }),
      );
      expect(reportRequests(result.requests)).toHaveLength(0);
      expect(result.requests.some((url) => url.hostname === REPORT_HOST)).toBe(false);
      const text = note(result);
      expect(text).toContain("Microsoft's report privacy setting says names are concealed");
      expect(text).toContain("(Settings, Org settings, Services, Reports)");
      expect(text).toContain("will not unmask anything itself");
      expect(result.complete).toBe(false);
      for (const row of result.rows) {
        expect(row.cells, row.key).toMatchObject({
          mailLast: null,
          mailSent: null,
          teamsLast: null,
          oneDriveLast: null,
          oneDriveActiveFiles: null,
          sharePointLast: null,
          mailboxStorage: null,
        });
      }
    });

    it("still shows sign-in, license and group columns, and decides on sign-ins alone where it can", async () => {
      const result = await run(
        tenant({ settings: () => Response.json({ displayConcealedNames: true }) }),
      );
      expect(result.rows).toHaveLength(accounts.length);
      expect(find(result.rows, "u-uses-nothing").cells).toMatchObject({
        licenseAssignment: "Through groups",
        days: 29,
        groups: 1,
      });
      expect(verdictOfRow(result.rows, "u-active")).toBe("Active");
      expect(verdictOfRow(result.rows, "u-never")).toBe("Never signed in");
      expect(verdictOfRow(result.rows, "u-quiet")).toBe("Not enough data");
      expect(verdictOfRow(result.rows, "u-89")).toBe("Not enough data");
      expect(verdictOfRow(result.rows, "u-others")).toBe("Not enough data");
      expect(result.summary.find((tile) => tile.label === "Quiet everywhere")?.value).toBe(0);
    });

    it("treats identifiers that look like hashes as concealed when the setting cannot be read, and stops after one report", async () => {
      const hash = (index: number) => index.toString(16).padStart(32, "0");
      const result = await run(
        tenant({
          settings: status(403),
          reports: {
            activeUsers: csv(headers.activeUsers, [
              { "User Principal Name": hash(1), "Exchange Last Activity Date": dayOf(1) },
              { "User Principal Name": hash(2), "Exchange Last Activity Date": dayOf(2) },
            ]),
          },
        }),
      );
      expect(reportRequests(result.requests)).toHaveLength(1);
      const text = note(result);
      expect(text).toContain(
        "could not read the report privacy setting, and the reports' identifiers look like hashes",
      );
      expect(text).toContain("Services, Reports");
      expect(verdictOfRow(result.rows, "u-quiet")).toBe("Not enough data");
      expect(find(result.rows, "u-active").cells.mailLast).toBeNull();
    });

    it("says so when the setting shows names but the reports still carry hashes", async () => {
      const hash = (index: number) => index.toString(16).padStart(32, "0");
      const result = await run(
        tenant({
          reports: {
            activeUsers: csv(headers.activeUsers, [{ "User Principal Name": hash(1) }]),
          },
        }),
      );
      expect(note(result)).toContain(
        "privacy setting says names are shown, but the reports still return hashed identifiers",
      );
      expect(find(result.rows, "u-active").cells.mailLast).toBeNull();
    });

    it("does not mistake readable names for hashes when only some look like one", async () => {
      const result = await run(
        tenant({
          reports: {
            activeUsers: csv(headers.activeUsers, [
              { "User Principal Name": "a".repeat(32) },
              { "User Principal Name": "ada@a.example", "Exchange Last Activity Date": dayOf(3) },
            ]),
          },
        }),
      );
      expect(note(result)).not.toContain("concealed");
      expect(find(result.rows, "u-active").cells.mailLast).toBe(`${dayOf(3)}T00:00:00.000Z`);
    });

    it("applies the audit scope to the directory even though it cannot to the reports", async () => {
      const result = await run(
        tenant({ settings: () => Response.json({ displayConcealedNames: true }) }),
        { scope: ["b.example"] },
      );
      expect(result.rows.map((row) => row.key)).toEqual(["u-quiet-b"]);
    });
  });

  describe("without the usage reports", () => {
    it("reads no report when the session has not granted Reports.Read.All, and says what that costs", async () => {
      const result = await run(tenant(), {
        granted: ["User.Read.All", "AuditLog.Read.All", "LicenseAssignment.Read.All"],
      });
      expect(reportRequests(result.requests)).toHaveLength(0);
      expect(result.requests.some((url) => url.pathname === "/v1.0/admin/reportSettings")).toBe(
        false,
      );
      expect(note(result)).toContain("Usage reports were not read");
      expect(result.complete).toBe(false);
      expect(verdictOfRow(result.rows, "u-active")).toBe("Active");
      expect(verdictOfRow(result.rows, "u-quiet")).toBe("Not enough data");
    });

    it("says the privacy setting was not read when only that permission is missing, and still reads names", async () => {
      const result = await run(tenant(), {
        granted: ["User.Read.All", "AuditLog.Read.All", "Reports.Read.All"],
      });
      expect(note(result)).toContain(
        "did not read the report privacy setting (it needs ReportSettings.Read.All)",
      );
      expect(find(result.rows, "u-active").cells.mailSent).toBe(12);
      expect(find(result.rows, "u-active").cells.licenses).toEqual([
        expect.stringMatching(/^1111/),
      ]);
    });
  });

  describe("the request limit", () => {
    // Quiet accounts first: with more accounts than the limit, the ones that matter are asked.
    const many = (total: number) => {
      const list: unknown[] = [];
      for (let index = 0; index < total; index += 1)
        list.push(
          account({
            id: `u-${index}`,
            upn: `person${index}@a.example`,
            name: `Person ${index}`,
            signedIn: index === total - 1 ? 200 : 1,
          }),
        );
      return list;
    };

    it("asks for the first 2,000 accounts, quiet ones first, and says the rest are not read", async () => {
      const total = MAX_MEMBERSHIP_ACCOUNTS + 100;
      const result = await run(tenant({ users: many(total), memberships: {} }));
      const asked = result.requests.filter((url) => url.pathname.endsWith("/memberOf"));
      expect(asked).toHaveLength(MAX_MEMBERSHIP_ACCOUNTS);
      expect(asked.some((url) => url.pathname.includes(`/u-${total - 1}/`))).toBe(true);
      expect(result.complete).toBe(false);
      expect(note(result)).toContain(
        "Group memberships were read for 2,000 of 2,100 accounts, quiet ones first. 100 accounts show Not read",
      );
      const unread = result.rows.filter((row) => row.cells.groups === null);
      expect(unread).toHaveLength(100);
      expect(find(result.rows, `u-${total - 1}`).cells.groups).toBe(0);
    });

    it("stays at the limit exactly without complaining", async () => {
      const result = await run(tenant({ users: many(MAX_MEMBERSHIP_ACCOUNTS), memberships: {} }));
      expect(result.requests.filter((url) => url.pathname.endsWith("/memberOf"))).toHaveLength(
        MAX_MEMBERSHIP_ACCOUNTS,
      );
      expect(note(result)).not.toContain("Group memberships were read for");
    });
  });

  describe("the audit scope", () => {
    it("keeps only the chosen domains in the rows, the tiles and the notes", async () => {
      const result = await run(tenant(), { scope: ["a.example"] });
      expect(result.rows.map((row) => row.cells.domain)).not.toContain("b.example");
      expect(result.rows).toHaveLength(accounts.length - 1);
      expect(result.notes[0]).toBe("Limited to accounts in: a.example.");
      expect(note(result)).toContain(
        "1 account in another domain was left out by the audit scope.",
      );
      expect(result.summary.find((tile) => tile.label === "Accounts")?.value).toBe(
        accounts.length - 1,
      );
      expect(result.summary.find((tile) => tile.label === "Quiet everywhere")?.value).toBe(3);
      expect(result.domains).toEqual([
        { domain: "a.example", accounts: accounts.length - 1 },
        { domain: "b.example", accounts: 1 },
      ]);
    });

    it("shows quiet accounts with directory roles in domains the scope leaves out, in their own critical tile", async () => {
      const result = await run(tenant(), { scope: ["a.example"] });
      expect(result.summary).toContainEqual({
        label: "Quiet everywhere, with directory roles, outside the scope",
        value: 1,
        tone: "critical",
        view: null,
      });
      expect(result.summary).toContainEqual({
        label: "Quiet everywhere, with directory roles",
        value: 1,
        tone: "critical",
        view: "quiet-roles",
      });
      expect(note(result)).toContain(
        "1 account quiet everywhere holds a directory role in domains the audit scope leaves out. Choose All domains and run again to see it.",
      );
    });

    it("has no outside tile when nothing outside the scope qualifies", async () => {
      const result = await run(tenant({ memberships: { "u-quiet-b": [] } }), {
        scope: ["a.example"],
      });
      expect(result.summary.map((tile) => tile.label)).not.toContain(
        "Quiet everywhere, with directory roles, outside the scope",
      );
      const all = await run();
      expect(all.summary.map((tile) => tile.label)).not.toContain(
        "Quiet everywhere, with directory roles, outside the scope",
      );
    });

    it("asks for memberships of quiet accounts outside the scope, which no row shows", async () => {
      const result = await run(tenant(), { scope: ["a.example"] });
      expect(result.requests.some((url) => url.pathname.includes("/u-quiet-b/memberOf"))).toBe(
        true,
      );
    });

    it("gives a domain column to filter on without changing the scope", async () => {
      const result = await run();
      expect(find(result.rows, "u-quiet-b").cells.domain).toBe("b.example");
      expect(find(result.rows, "u-active").cells.domain).toBe("a.example");
    });
  });

  describe("the handoff", () => {
    it("carries the directory id of every account", async () => {
      const result = await run();
      expect(auditDefinition(audit).handoff).toBe(true);
      for (const row of result.rows) {
        expect(row.personId, row.key).toBe(row.key);
        expect(auditRowSchema.parse(row).personId).toBe(row.key);
      }
    });
  });

  describe("the notes", () => {
    it("fit the result, whole, and read like a colleague wrote them", async () => {
      const results = [
        await run(),
        await run(tenant({ reports: { email: status(403) } })),
        await run(tenant({ settings: () => Response.json({ displayConcealedNames: true }) })),
        await run(tenant(), { scope: ["a.example"] }),
      ];
      for (const result of results) {
        expect(result.notes.length).toBeLessThanOrEqual(20);
        for (const item of result.notes) {
          expect(item.length, item).toBeLessThanOrEqual(500);
          expect(item).not.toMatch(/[—·!]/);
        }
        for (const tile of result.summary) expect(tile.label).not.toMatch(/[—·!]/);
      }
    });

    it("keep to 20 when a run has more to say, dropping the standing ones last", async () => {
      const result = await run(
        tenant({
          users: [...accounts, unreadable, { displayName: "no id" }],
          settings: status(403),
          reports: {
            // Read, with a row the reader cannot keep and the departed account every file names.
            activeUsers: `${files.activeUsers}only,three,fields\r\n`,
            // Read, with a count that is not a number.
            oneDriveUsage: csv(headers.oneDriveUsage, [
              { "Owner Principal Name": "pat@a.example", "Active File Count": "many" },
            ]),
            // The other five cannot be read, each for its own reason.
            email: status(404),
            teams: status(500),
            oneDriveActivity: status(429, { "Retry-After": "600" }),
            sharePoint: csv(["Report Refresh Date", "User Principal Name"], []),
          },
          downloads: { mailbox: () => new Response("gone", { status: 410 }) },
          memberships: { "u-31": [member("device", "d-1", "Laptop")] },
        }),
        {
          scope: ["a.example"],
          granted: ["User.Read.All", "AuditLog.Read.All", "Reports.Read.All"],
        },
      );
      expect(result.notes).toHaveLength(20);
      expect(result.notes[0]).toBe("Limited to accounts in: a.example.");
      const text = note(result);
      for (const expected of [
        "left out by the audit scope",
        "holds a directory role in domains the audit scope leaves out",
        "did not read the report privacy setting",
        "The email activity report could not be read",
        "The Teams activity report could not be read",
        "The OneDrive activity report could not be read",
        "The mailbox usage report could not be read",
        "The SharePoint activity report could not be read",
        "1 row in the reports could not be read",
        "1 value in the reports could not be read",
        "1 name in the reports matches no account",
        "1 account has memberships Prism365 could not read",
        "1 account has sign-in dates Prism365 could not read",
        "License names could not be read",
        "1 account could not be read",
        "lag by 24 to 72 hours",
        "None of this measures anyone's work",
        "Microsoft does not say who opened them",
      ])
        expect(text, expected).toContain(expected);
      // The standing note that comes last is the one that goes.
      expect(text).not.toContain("does not read mailbox type");
    });

    it("give the report lag, the roles that unlock per-user detail, the direct-only groups and the mailbox limit", async () => {
      const text = note(await run());
      expect(text).toContain("lag by 24 to 72 hours");
      expect(text).toContain("Reports Reader or a service administrator role");
      expect(text).toContain("Group memberships are direct only");
      expect(text).toContain("does not read mailbox type");
      expect(text).toContain("ADR-0009");
      expect(text).toContain("None of this measures anyone's work");
    });

    it("say why the OneDrive flag asks for a check", async () => {
      expect(note(await run())).toContain("Microsoft does not say who opened them");
    });
  });

  describe("the result", () => {
    it("is complete when everything was read", async () => {
      const result = await run();
      expect(result.complete).toBe(true);
    });

    it("is partial when accounts could not be read", async () => {
      const result = await run(tenant({ users: [...accounts, { displayName: "no id" }] }));
      expect(result.complete).toBe(false);
      expect(note(result)).toContain("1 account could not be read");
    });

    it("lists every domain it read with its number of accounts", async () => {
      const result = await run();
      expect(result.domains).toEqual([
        { domain: "a.example", accounts: accounts.length - 1 },
        { domain: "b.example", accounts: 1 },
      ]);
    });

    it("fails when Microsoft refuses sign-in dates, so the runner can name the permission", async () => {
      await expect(run(tenant(), { granted: ["User.Read.All"] })).rejects.toMatchObject({
        status: 403,
      });
    });

    it("stops reading when the run is stopped", async () => {
      const controller = new AbortController();
      const world = fake(audit, tenant(), { granted, signal: controller.signal });
      controller.abort(new Error("stopped"));
      await expect(audit.run(world.context)).rejects.toThrow("stopped");
      expect(world.requests).toHaveLength(0);
    });

    it("reads nothing but with GET, and no beta endpoint", async () => {
      const result = await run();
      for (const url of result.requests) expect(url.pathname).not.toContain("/beta/");
    });
  });
});

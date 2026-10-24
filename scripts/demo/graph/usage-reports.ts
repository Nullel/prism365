import { createHash } from "node:crypto";
import type { DirectoryUser, SyntheticTenant } from "../tenant.ts";
import type { DemoGraphExtension } from "./types.ts";

// The Microsoft 365 apps usage report, the active users report and the report privacy setting for
// the synthetic tenant.
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

function activeUsersOf(tenant: SyntheticTenant): string {
  const now = tenant.now.getTime();
  const partNumber = new Map(tenant.skus.map((sku) => [sku.skuId, sku.skuPartNumber]));
  const flag = (value: boolean) => (value ? "True" : "False");
  const people = tenant.users.filter(
    (user) => user.userType === "Member" && user.assignedLicenses.length > 0,
  );
  const rows = people.map((user) => {
    const parts = user.assignedLicenses.map((license) => partNumber.get(license.skuId) ?? "");
    const suite = parts.includes("SPB");
    const exchange = suite || parts.includes("EXCHANGESTANDARD");
    const signedIn = user.signInActivity.lastSuccessfulSignInDateTime;
    // The report is two days behind, so nobody's last activity is later than that.
    const last = signedIn === null ? null : Math.min(Date.parse(signedIn), now - 2 * DAY);
    const activity = (workload: string, licensed: boolean, chance: number) =>
      licensed && last !== null && unit(`${user.id}:${workload}`) < chance
        ? day(last - Math.floor(unit(`${user.id}:${workload}:lag`) * 4) * DAY)
        : "";
    const assigned = (licensed: boolean) => (licensed ? day(Date.parse(user.createdDateTime)) : "");
    const row: Record<ActiveUserColumn, string> = {
      "Report Refresh Date": day(now - 2 * DAY),
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
      "Exchange Last Activity Date": activity("exchange", exchange, 1),
      "OneDrive Last Activity Date": activity("onedrive", suite, 0.7),
      "SharePoint Last Activity Date": activity("sharepoint", suite, 0.6),
      "Skype For Business Last Activity Date": "",
      "Yammer Last Activity Date": activity("yammer", suite, 0.1),
      "Teams Last Activity Date": activity("teams", suite, 0.85),
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

const built = new WeakMap<SyntheticTenant, object[]>();

// Microsoft 365 usage reports and the report privacy setting.
export const usageReports: DemoGraphExtension = (request, tools) => {
  if (request.path === "admin/reportSettings") return tools.json({ displayConcealedNames: false });
  if (ACTIVE_USERS.test(request.path)) return tools.download(activeUsersOf(request.tenant));
  if (request.path !== "reports/getM365AppUserDetail(period='D90')") return null;
  let report = built.get(request.tenant);
  if (!report) {
    report = reportOf(request.tenant);
    built.set(request.tenant, report);
  }
  return tools.page(request.url, report, null);
};

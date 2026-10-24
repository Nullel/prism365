import { createHash } from "node:crypto";
import type { DirectoryUser, SyntheticTenant } from "../tenant.ts";
import type { DemoGraphExtension } from "./types.ts";

// The Microsoft 365 apps usage report and the report privacy setting for the synthetic tenant.
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

const built = new WeakMap<SyntheticTenant, object[]>();

// Microsoft 365 usage reports and the report privacy setting.
export const usageReports: DemoGraphExtension = (request, tools) => {
  if (request.path === "admin/reportSettings") return tools.json({ displayConcealedNames: false });
  if (request.path !== "reports/getM365AppUserDetail(period='D90')") return null;
  let report = built.get(request.tenant);
  if (!report) {
    report = reportOf(request.tenant);
    built.set(request.tenant, report);
  }
  return tools.page(request.url, report, null);
};

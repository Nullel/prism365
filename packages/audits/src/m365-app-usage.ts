import { type AuditRow, cleanText } from "@prism365/contracts";
import { GraphRequestError } from "@prism365/graph";
import { z } from "zod";
import type { Audit, AuditContext } from "./definition.ts";
import { DAY_MS } from "./directory.ts";
import { firstAccepted, readPages } from "./graph-pages.ts";
import {
  domainColumn,
  domainInventory,
  excludedNote,
  inAuditScope,
  principalDomain,
  scopeNote,
} from "./scope.ts";

// Who has used Microsoft 365 apps in the last 90 days, and on which platforms, from Microsoft's
// Microsoft 365 apps usage report (the one usage report v1.0 serves as JSON). Microsoft Graph v1.0
// documentation verified 2026-10-22:
// https://learn.microsoft.com/en-us/graph/api/reportroot-getm365appuserdetail?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/adminreportsettings-get?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/reportroot-authorization

const reportContext = { feature: "reports.usage.read", permission: "Reports.Read.All" } as const;
const settingsContext = {
  feature: "reports.settings.read",
  permission: "ReportSettings.Read.All",
} as const;

export const PERIOD_DAYS = 90;
export const MAX_USAGE_ACCOUNTS = 10_000;
// Microsoft documents $top for this report without a maximum, so a larger page is asked for first
// (the report allows few requests every ten minutes) and the default of 200 is the fallback.
const REPORT = "reports/getM365AppUserDetail(period='D90')";
const forms = [
  `${REPORT}?$format=application/json&$top=1000`,
  `${REPORT}?$format=application/json`,
];
const EARLIEST = Date.parse("2000-01-01T00:00:00Z");
// Microsoft hides names behind MD5 hashes, which are 32 hexadecimal characters.
export const HASHED = /^[0-9a-f]{32}$/i;

const apps = [
  { key: "outlook", label: "Outlook" },
  { key: "word", label: "Word" },
  { key: "excel", label: "Excel" },
  { key: "powerPoint", label: "PowerPoint" },
  { key: "oneNote", label: "OneNote" },
  { key: "teams", label: "Teams" },
] as const;
const platforms = [
  { key: "windows", label: "Windows" },
  { key: "mac", label: "Mac" },
  { key: "mobile", label: "Mobile" },
  { key: "web", label: "Web" },
] as const;

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
// A row key is at most 200 characters, and no sign-in name is longer than that.
const principal = z
  .string()
  .max(200)
  .transform((value) => cleanText(value))
  .nullish()
  .catch(null)
  .transform((value) => value ?? null);

// Microsoft's sample nests the app and platform flags under details[], while the flat shape of
// its CSV would put them on the record itself. Both are read.
const usedFlags = z.object({
  outlook: flag,
  word: flag,
  excel: flag,
  powerPoint: flag,
  oneNote: flag,
  teams: flag,
  windows: flag,
  mac: flag,
  mobile: flag,
  web: flag,
});
const recordSchema = usedFlags.extend({
  userPrincipalName: principal,
  displayName: text,
  userDisplayName: text,
  lastActivityDate: text,
  lastActivationDate: text,
  details: z
    .union([z.array(usedFlags).max(20), usedFlags])
    .nullish()
    .catch(null)
    .transform((value) => value ?? null),
});
type UsageRecord = z.output<typeof recordSchema>;
type Flag = keyof z.output<typeof usedFlags>;

function used(record: UsageRecord, key: Flag): boolean {
  if (record[key] === true) return true;
  const details = record.details === null ? [] : [record.details].flat();
  return details.some((detail) => detail[key] === true);
}

export interface Usage {
  apps: string[];
  platforms: string[];
  teamsOnly: boolean;
  desktop: boolean;
  webOnly: boolean;
  // Used on the web or a phone, and on no desktop.
  webMobile: boolean;
}

export function usageOf(record: UsageRecord): Usage {
  const appsUsed = apps.filter((app) => used(record, app.key));
  const platformsUsed = platforms.filter((platform) => used(record, platform.key));
  const on = (key: Flag) => used(record, key);
  const desktop = on("windows") || on("mac");
  return {
    apps: appsUsed.map((app) => app.label),
    platforms: platformsUsed.map((platform) => platform.label),
    teamsOnly: on("teams") && appsUsed.length === 1,
    desktop,
    webOnly: on("web") && !desktop && !on("mobile"),
    webMobile: (on("web") || on("mobile")) && !desktop,
  };
}

export interface LastActivity {
  at: string | null;
  days: number | null;
  // none: the report has no activity for the account. unreadable: it has a date that cannot be real.
  state: "none" | "known" | "unreadable";
}

// The report gives dates without a time ("2026-10-20").
export function lastActivityOf(value: string | null, now: Date): LastActivity {
  if (!value) return { at: null, days: null, state: "none" };
  const date = /^\d{4}-\d{2}-\d{2}/.exec(value)?.[0];
  const at = date ? Date.parse(`${date}T00:00:00Z`) : Number.NaN;
  if (Number.isNaN(at) || at < EARLIEST || at > now.getTime() + DAY_MS)
    return { at: null, days: null, state: "unreadable" };
  return {
    at: new Date(at).toISOString(),
    days: Math.max(0, Math.floor((now.getTime() - at) / DAY_MS)),
    state: "known",
  };
}

// An account with no activity in the report, or none inside the period.
const isIdle = (last: LastActivity) =>
  last.state === "none" || (last.state === "known" && (last.days ?? 0) >= PERIOD_DAYS);

const settingsSchema = z.object({
  displayConcealedNames: flag,
  // Microsoft's sample nests the setting under value.
  value: z
    .object({ displayConcealedNames: flag })
    .nullish()
    .catch(null)
    .transform((value) => value ?? null),
});

export interface Setting {
  // true when Microsoft says names are concealed, false when it says they are shown, null when
  // the setting was not read.
  concealed: boolean | null;
  why: "read" | "not_granted" | "refused";
}

export async function readSetting(context: AuditContext): Promise<Setting> {
  if (!context.granted.has("ReportSettings.Read.All"))
    return { concealed: null, why: "not_granted" };
  try {
    const parsed = settingsSchema.safeParse(
      await context.graph.get("admin/reportSettings", settingsContext, z.unknown()),
    );
    const concealed = parsed.success
      ? (parsed.data.displayConcealedNames ?? parsed.data.value?.displayConcealedNames ?? null)
      : null;
    return { concealed, why: "read" };
  } catch (error) {
    // The setting only decides how to show the report: a refusal does not fail the run.
    if (error instanceof GraphRequestError && (error.status === 403 || error.status === 404))
      return { concealed: null, why: "refused" };
    throw error;
  }
}

const concealedAdvice =
  "A Global Administrator can turn off concealment in the Microsoft 365 admin center (Settings, Org settings, Reports). Prism365 will not unmask anything itself.";

// What the administrator is told about names: why there are none, or how Prism365 knew it could
// show them.
function privacyNotes(setting: Setting, concealed: boolean): string[] {
  if (concealed)
    return [
      setting.concealed === true
        ? "Names are concealed in this tenant's Microsoft usage reports, so there are no rows to show, only the totals above."
        : setting.concealed === false
          ? "Microsoft's privacy setting says names are shown, but the report still returns hashed identifiers (a change can take a few minutes to reach the reports). Only totals are shown for now."
          : "Prism365 could not read the report privacy setting, and the report's identifiers look like hashes, so they are treated as concealed. Only totals are shown.",
      concealedAdvice,
    ];
  if (setting.concealed !== null) return [];
  const reason =
    setting.why === "not_granted"
      ? "Prism365 did not read the report privacy setting (it needs ReportSettings.Read.All)."
      : setting.why === "refused"
        ? "Microsoft refused the report privacy setting."
        : "Microsoft's answer about the report privacy setting could not be read.";
  return [`${reason} Names are shown because Microsoft returned readable sign-in names.`];
}

const plural = (count: number, one: string, many: string) =>
  `${count.toLocaleString("en-US")} ${count === 1 ? one : many}`;

interface UsageAuditOptions {
  // A pause between two pages of the report, in milliseconds. Tests set it to zero.
  pageDelayMilliseconds?: number;
}

export function createM365AppUsageAudit(options: UsageAuditOptions = {}): Audit {
  const pageDelayMilliseconds = options.pageDelayMilliseconds ?? 0;
  return {
    definition: {
      id: "usage.m365-apps",
      // No handoff: the report names accounts only by sign-in name, so no row carries the
      // directory id an action tool acts on (docs/handoff.md).
      handoff: false,
      version: 1,
      title: "Microsoft 365 app usage",
      description:
        "Which accounts used Outlook, Word, Excel, PowerPoint, OneNote or Teams in the last 90 days, on which platforms, and which used none: reclaim candidates the license audits cannot see. For mail, Teams, OneDrive and SharePoint activity, see Account activity.",
      category: "Licenses",
      subject: { one: "account", many: "accounts" },
      columns: [
        { key: "person", label: "Account", kind: "person" },
        {
          key: "lastActivity",
          label: "Last activity",
          kind: "date",
          empty: "None in 90 days",
          help: "The latest day the report shows the account using any of the six apps. Microsoft gives a date, not a time.",
        },
        {
          key: "days",
          label: "Since activity",
          kind: "days",
          empty: "No activity",
          help: "Days since the last activity. Empty when the report has none in the period.",
        },
        {
          key: "apps",
          label: "Apps used",
          kind: "list",
          help: "Used at least once in the period.",
        },
        {
          key: "platforms",
          label: "Platforms",
          kind: "list",
          help: "Where any of the apps was used in the period.",
        },
        ...apps.map((app) => ({
          key: app.key,
          label: app.label,
          kind: "boolean" as const,
          hidden: true,
        })),
        domainColumn,
        {
          key: "activated",
          label: "Last activation",
          kind: "date",
          hidden: true,
          empty: "Unknown",
        },
      ],
      views: [
        { id: "idle", label: "No app activity in 90 days", tags: ["idle"] },
        { id: "teams-only", label: "Teams only", tags: ["teams-only"] },
        { id: "web-only", label: "Web only", tags: ["web-only"] },
        { id: "desktop", label: "Desktop apps used", tags: ["desktop"] },
        { id: "web-mobile", label: "Web and mobile only", tags: ["web-mobile"] },
        { id: "all", label: "All", tags: [] },
      ],
      defaultView: "idle",
      defaultSort: { column: "person", direction: "asc" },
      requirements: [
        "Reports Reader or a service administrator role (Exchange, SharePoint or Teams). Global Reader and Usage Summary Reports Reader see tenant totals only, not this per-user report.",
        "Names show only when the tenant's report privacy setting shows them. A Global Administrator controls that setting.",
      ],
    },
    packs: ["reports.usage.read"],
    optionalPacks: ["reports.settings.read"],
    scoped: true,
    forbidden:
      "Microsoft refused the Microsoft 365 apps usage report. It needs Reports.Read.All and a role that can see per-user usage, such as Reports Reader or an Exchange, SharePoint or Teams administrator. Global Reader sees tenant totals only.",
    async run(context) {
      context.signal.throwIfAborted();
      const setting = await readSetting(context);
      const read = await firstAccepted(forms, (path) =>
        readPages(context, reportContext, path, {
          maxRecords: MAX_USAGE_ACCOUNTS,
          pageDelayMilliseconds,
        }),
      );
      let skipped = 0;
      const records: (UsageRecord & { userPrincipalName: string })[] = [];
      for (const raw of read.items) {
        const parsed = recordSchema.safeParse(raw);
        if (!parsed.success || !parsed.data.userPrincipalName) {
          skipped += 1;
          continue;
        }
        records.push({ ...parsed.data, userPrincipalName: parsed.data.userPrincipalName });
      }
      const hashed = records.length > 0 && records.every((r) => HASHED.test(r.userPrincipalName));
      const concealed = setting.concealed === true || hashed;

      const analysed = records.map((record) => ({
        record,
        usage: usageOf(record),
        last: lastActivityOf(record.lastActivityDate, context.now),
      }));
      // Names are never matched, shown or stored when Microsoft conceals them: the audit then
      // reports totals and no rows.
      const covered = concealed
        ? analysed
        : analysed.filter((item) => inAuditScope(context.scope, item.record.userPrincipalName));
      const rows: AuditRow[] = concealed
        ? []
        : covered.map(({ record, usage, last }) => ({
            key: record.userPrincipalName,
            // The report names accounts by sign-in name only, so there is nothing to link to.
            personId: null,
            tags: [
              ...(isIdle(last) ? ["idle"] : []),
              ...(usage.teamsOnly ? ["teams-only"] : []),
              ...(usage.webOnly ? ["web-only"] : []),
              ...(usage.desktop ? ["desktop"] : []),
              ...(usage.webMobile ? ["web-mobile"] : []),
            ],
            cells: {
              person: {
                name: record.displayName ?? record.userDisplayName,
                principal: record.userPrincipalName,
              },
              lastActivity: last.at,
              days: last.days,
              apps: usage.apps,
              platforms: usage.platforms,
              ...Object.fromEntries(apps.map((app) => [app.key, used(record, app.key)])),
              domain: principalDomain(record.userPrincipalName),
              activated: lastActivityOf(record.lastActivationDate, context.now).at,
            },
          }));
      rows.sort((a, b) => {
        const name = (row: AuditRow) => {
          const person = row.cells.person as { name: string | null; principal: string | null };
          return (person.name ?? person.principal ?? "").toLowerCase();
        };
        return name(a).localeCompare(name(b));
      });

      const unreadableDates = covered.filter((item) => item.last.state === "unreadable").length;
      const idle = covered.filter((item) => isIdle(item.last)).length;
      const desktop = covered.filter((item) => item.usage.desktop).length;
      const webMobile = covered.filter((item) => item.usage.webMobile).length;
      const view = (id: string) => (concealed ? null : id);

      const limited = concealed ? null : scopeNote(context.scope);
      const left = concealed ? null : excludedNote(analysed.length - covered.length);
      const notes = [
        ...(limited ? [limited] : []),
        ...(left ? [left] : []),
        ...privacyNotes(setting, concealed),
        ...(concealed && context.scope && context.scope.size > 0
          ? ["The audit scope could not be applied: the report gives no names to sort by domain."]
          : []),
        "Microsoft's usage reports lag by 24 to 72 hours and sometimes by several days, so the most recent activity can be missing.",
        `The report covers the last ${PERIOD_DAYS} days and has columns for Outlook, Word, Excel, PowerPoint, OneNote and Teams only.`,
        "Per-user detail needs a role such as Reports Reader or a service administrator. Global Reader sees tenant totals only.",
        "Used means the app was opened on that platform at least once in the period. It says nothing about how long or how well, and it does not measure anyone's work.",
        ...(concealed
          ? []
          : [
              "Rows do not link to a person page: the report names accounts by sign-in name, not by object ID.",
            ]),
        ...(read.throttled
          ? [
              `Microsoft limits how often this report can be read, and asked Prism365 to slow down after ${plural(read.items.length, "account", "accounts")}. The rest were not read. Run the audit again in about ten minutes.`,
            ]
          : read.complete
            ? []
            : [
                `Stopped at ${MAX_USAGE_ACCOUNTS.toLocaleString("en-US")} accounts; the rest of the report was not read.`,
              ]),
        ...(unreadableDates > 0
          ? [
              `${plural(unreadableDates, "account has", "accounts have")} a last activity date Prism365 could not read. ${unreadableDates === 1 ? "It is" : "They are"} listed but never counted as idle.`,
            ]
          : []),
        ...(skipped > 0
          ? [
              `${plural(skipped, "account", "accounts")} in the report could not be read and ${skipped === 1 ? "is" : "are"} left out.`,
            ]
          : []),
      ];

      return {
        rows,
        complete: read.complete && skipped === 0 && unreadableDates === 0,
        ...(concealed
          ? {}
          : { domains: domainInventory(analysed.map((i) => i.record.userPrincipalName)) }),
        summary: [
          {
            label: `No Microsoft 365 app activity in ${PERIOD_DAYS} days`,
            value: idle,
            tone: idle > 0 ? "warning" : "good",
            view: view("idle"),
          },
          { label: "Desktop apps used", value: desktop, tone: "neutral", view: view("desktop") },
          {
            label: "Web and mobile only",
            value: webMobile,
            tone: "neutral",
            view: view("web-mobile"),
          },
        ],
        notes,
      };
    },
  };
}

export const m365AppUsageAudit = createM365AppUsageAudit({ pageDelayMilliseconds: 500 });

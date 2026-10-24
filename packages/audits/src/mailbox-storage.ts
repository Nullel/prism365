import { type AuditDefinitionInput, type AuditRow, cleanText } from "@prism365/contracts";
import { CSV_MAX_ROWS } from "@prism365/graph";
import { countOf, flagOf, joinKey, type Tally } from "./activity-reports.ts";
import { type Audit, AuditError } from "./definition.ts";
import { HASHED, lastActivityOf, readSetting, type Setting } from "./m365-app-usage.ts";
import {
  domainColumn,
  domainInventory,
  excludedNote,
  inAuditScope,
  principalDomain,
  scopeNote,
} from "./scope.ts";

// How full each mailbox is against its quotas, from Microsoft's mailbox usage report (D90). The
// report names mailboxes by sign-in name and carries no object ID, so nothing is joined to the
// directory: the audit needs only the usage reports read pack. Microsoft Graph v1.0 documentation
// verified 2026-10-24:
// https://learn.microsoft.com/en-us/graph/api/reportroot-getmailboxusagedetail?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/exchange/mailbox-quotas

const reportContext = { feature: "reports.usage.read", permission: "Reports.Read.All" } as const;
const REPORT_PATH = "reports/getMailboxUsageDetail(period='D90')";
const GIB = 1_073_741_824;
// The columns without which the report says nothing about size. Every other column is optional:
// Microsoft adds columns over time, and a missing one reads as unknown.
const REQUIRED = ["User Principal Name", "Storage Used (Byte)"] as const;

export const MAILBOX_STATUSES = [
  "Cannot send or receive",
  "Cannot send",
  "Over warning",
  "Near full",
  "Near full, no archive",
  "Quota unknown",
  "OK",
] as const;
export type MailboxStatus = (typeof MAILBOX_STATUSES)[number];

// A mailbox with nothing to say has no tag, so no view lists it among the problems.
const statusTags: Record<MailboxStatus, string[]> = {
  "Cannot send or receive": ["cannot-send-receive"],
  "Cannot send": ["cannot-send"],
  "Over warning": ["over-warning"],
  "Near full": ["near-full"],
  "Near full, no archive": ["near-full-no-archive"],
  "Quota unknown": ["quota-unknown"],
  OK: [],
};

// A quota in bytes, or null when the report leaves it blank or gives 0, which is no quota at all.
export interface Quotas {
  warning: number | null;
  send: number | null;
  sendReceive: number | null;
}

// The first status that matches, worst first. Reaching a quota counts as being over it, and the
// percentages are compared as whole numbers so a mailbox exactly at 90 or 80 percent is not lost
// to rounding. A quota the report does not give is skipped; without the warning quota the two
// near-full checks cannot be made, so what is left is Quota unknown, not OK.
export function mailboxStatus(
  used: number,
  quotas: Quotas,
  archive: boolean | null,
): MailboxStatus {
  const { warning, send, sendReceive } = quotas;
  if (sendReceive !== null && used >= sendReceive) return "Cannot send or receive";
  if (send !== null && used >= send) return "Cannot send";
  if (warning === null) return "Quota unknown";
  if (used >= warning) return "Over warning";
  if (used * 10 >= warning * 9) return "Near full";
  // Only a mailbox known to have no archive: an unknown archive is not called missing.
  if (archive === false && used * 5 >= warning * 4) return "Near full, no archive";
  return "OK";
}

const gigabytes = (bytes: number | null): number | null =>
  bytes === null ? null : Math.round((bytes / GIB) * 10) / 10;

const plural = (count: number, one: string, many: string) =>
  `${count.toLocaleString("en-US")} ${count === 1 ? one : many}`;

type Column = AuditDefinitionInput["columns"][number];

const size = (key: string, label: string, help: string | null, hidden = false): Column => ({
  key,
  label,
  kind: "number",
  hidden,
  help,
  empty: "Not in the report",
});

const columns: Column[] = [
  { key: "mailbox", label: "Mailbox", kind: "text" },
  { key: "displayName", label: "Display name", kind: "text", empty: "No name" },
  {
    key: "recipientType",
    label: "Recipient type",
    kind: "enum",
    facet: true,
    empty: "Unknown",
    help: "As Microsoft's report states it (a user, shared or room mailbox, for example). Older copies of the report have no such column, and then it is unknown.",
  },
  {
    key: "status",
    label: "Status",
    kind: "enum",
    facet: true,
    help: "The first that applies. Cannot send or receive: at or over the prohibit send and receive quota. Cannot send: at or over the prohibit send quota. Over warning: at or over the warning quota. Near full: 90 percent of the warning quota or more. Near full, no archive: 80 percent or more, and the mailbox has no archive. Quota unknown: the report gives no warning quota. Reaching a quota counts as being over it.",
  },
  size(
    "used",
    "Used (GB)",
    "Gigabytes in the mailbox, to one decimal. A gigabyte here is 1,073,741,824 bytes.",
  ),
  size("warningQuota", "Warning quota (GB)", "The size at which Microsoft warns the owner."),
  size(
    "sendQuota",
    "Prohibit send quota (GB)",
    "The size at which the mailbox can no longer send mail.",
  ),
  size(
    "sendReceiveQuota",
    "Prohibit send and receive quota (GB)",
    "The size at which the mailbox can neither send nor receive mail.",
  ),
  {
    key: "percent",
    label: "Percent of prohibit send quota",
    kind: "number",
    empty: "Not in the report",
    help: "Used divided by the prohibit send quota. Over 100 means the mailbox is past it.",
  },
  {
    key: "items",
    label: "Items",
    kind: "number",
    hidden: true,
    empty: "Not in the report",
    help: "Items in the mailbox, as the report counts them.",
  },
  size(
    "deletedSize",
    "Deleted items (GB)",
    "Gigabytes held in the Deleted Items and recoverable items folders, to one decimal.",
    true,
  ),
  {
    key: "archive",
    label: "Has archive",
    kind: "boolean",
    empty: "Unknown",
    help: "Whether the mailbox has an archive mailbox. A full mailbox with an archive can often make room by moving old mail there.",
  },
  {
    key: "lastActivity",
    label: "Last activity",
    kind: "date",
    empty: "No activity recorded",
    help: "From the mailbox usage report.",
  },
  domainColumn,
];

interface Mailbox {
  principal: string;
  name: string | null;
  type: string | null;
  used: number;
  quotas: Quotas;
  items: number | null;
  deletedBytes: number | null;
  archive: boolean | null;
  last: string | null;
  status: MailboxStatus;
}

const concealedAdvice =
  "A Global Administrator can turn off concealment in the Microsoft 365 admin center (Settings, Org settings, Reports). Prism365 will not unmask anything itself.";

// Why there are no rows, or how Prism365 knew it could show names.
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
      ? "Prism365 did not read the report privacy setting (it needs ReportSettings.Read.All)"
      : setting.why === "refused"
        ? "Microsoft refused the report privacy setting"
        : "Microsoft's answer about the report privacy setting could not be read";
  return [`${reason}. Names are shown because the report returned readable sign-in names.`];
}

export const mailboxStorageAudit: Audit = {
  definition: {
    id: "mail.mailbox-storage",
    handoff: false,
    version: 1,
    title: "Mailbox storage",
    description:
      "Which mailboxes are full or close to it: size against the warning, send and send-and-receive quotas, with the archive, items and deleted items. The list to check before mail starts bouncing.",
    category: "Mail",
    subject: { one: "mailbox", many: "mailboxes" },
    columns,
    views: [
      { id: "cannot-send-receive", label: "Cannot send or receive", tags: ["cannot-send-receive"] },
      { id: "cannot-send", label: "Cannot send", tags: ["cannot-send"] },
      { id: "over-warning", label: "Over warning", tags: ["over-warning"] },
      { id: "near-full", label: "Near full", tags: ["near-full"] },
      {
        id: "near-full-no-archive",
        label: "Near full, no archive",
        tags: ["near-full-no-archive"],
      },
      { id: "quota-unknown", label: "Quota unknown", tags: ["quota-unknown"] },
      { id: "all", label: "All mailboxes", tags: [] },
    ],
    defaultView: "all",
    defaultSort: { column: "used", direction: "desc" },
    requirements: [
      "A role that can read the usage reports, such as Reports Reader or Exchange Administrator.",
      "Names show in the usage reports only when the tenant's report privacy setting shows them. A Global Administrator controls that setting.",
    ],
  },
  packs: ["reports.usage.read"],
  optionalPacks: ["reports.settings.read"],
  scoped: true,
  forbidden:
    "Microsoft refused the mailbox usage report. It needs Reports.Read.All and a role that can read the usage reports, such as Reports Reader or Exchange Administrator.",
  async run(context) {
    context.signal.throwIfAborted();
    const setting = await readSetting(context);
    // The one report is the audit, so a report that cannot be read fails the run, and the runner
    // explains why (refused, throttled, a download link it does not accept).
    const table = await context.graph.getReport(REPORT_PATH, reportContext);
    // Headers are matched by name and without regard to letter case, never by position.
    const index = new Map<string, number>();
    for (const [position, name] of table.header.entries()) {
      const key = name.trim().toLowerCase();
      if (!index.has(key)) index.set(key, position);
    }
    for (const name of REQUIRED)
      if (!index.has(name.toLowerCase()))
        throw new AuditError(
          "report_columns",
          `Microsoft's mailbox usage report has no "${name}" column, so Prism365 cannot read it.`,
        );
    const tally: Tally = { unreadable: 0 };
    const cell = (row: readonly string[], name: string): string | undefined => {
      const at = index.get(name.toLowerCase());
      return at === undefined ? undefined : row[at];
    };
    // A quota of 0 is no quota: read as one, every mailbox would be over it.
    const quota = (row: readonly string[], name: string): number | null => {
      const value = countOf(cell(row, name), tally);
      return value ? value : null;
    };
    const text = (raw: string | undefined, length: number): string | null => {
      const value = cleanText((raw ?? "").trim().slice(0, length)).trim();
      return value === "" ? null : value;
    };

    const mailboxes: Mailbox[] = [];
    const seen = new Set<string>();
    let deleted = 0;
    let unreadable = 0;
    for (const row of table.rows) {
      // A row Microsoft marks deleted is a mailbox that is gone.
      if (flagOf(cell(row, "Is Deleted"), tally) === true) {
        deleted += 1;
        continue;
      }
      const principal = text(cell(row, "User Principal Name"), 200);
      const used = countOf(cell(row, "Storage Used (Byte)"), tally);
      if (principal === null || used === null || used === undefined) {
        unreadable += 1;
        continue;
      }
      // One row for each mailbox is enough; a second would only repeat or contradict it.
      const key = joinKey(principal);
      if (seen.has(key)) continue;
      seen.add(key);
      const quotas: Quotas = {
        warning: quota(row, "Issue Warning Quota (Byte)"),
        send: quota(row, "Prohibit Send Quota (Byte)"),
        sendReceive: quota(row, "Prohibit Send/Receive Quota (Byte)"),
      };
      const archive = flagOf(cell(row, "Has Archive"), tally) ?? null;
      const last = lastActivityOf(cell(row, "Last Activity Date")?.trim() || null, context.now);
      if (last.state === "unreadable") tally.unreadable += 1;
      mailboxes.push({
        principal,
        name: text(cell(row, "Display Name"), 200),
        type: text(cell(row, "Recipient Type"), 40),
        used,
        quotas,
        items: countOf(cell(row, "Item Count"), tally) ?? null,
        deletedBytes: countOf(cell(row, "Deleted Item Size (Byte)"), tally) ?? null,
        archive,
        last: last.at,
        status: mailboxStatus(used, quotas, archive),
      });
    }

    const concealed =
      setting.concealed === true ||
      (mailboxes.length > 0 && mailboxes.every((mailbox) => HASHED.test(mailbox.principal)));
    // Names are never matched, shown or stored when Microsoft conceals them: the audit then
    // reports totals and no rows.
    const covered = concealed
      ? mailboxes
      : mailboxes.filter((mailbox) => inAuditScope(context.scope, mailbox.principal));
    covered.sort((a, b) => b.used - a.used || a.principal.localeCompare(b.principal));

    const rows: AuditRow[] = concealed
      ? []
      : covered.map((mailbox) => ({
          key: joinKey(mailbox.principal),
          // The report names mailboxes by sign-in name only, so there is nothing to link to.
          personId: null,
          tags: statusTags[mailbox.status],
          cells: {
            mailbox: mailbox.principal,
            displayName: mailbox.name,
            recipientType: mailbox.type,
            status: mailbox.status,
            used: gigabytes(mailbox.used),
            warningQuota: gigabytes(mailbox.quotas.warning),
            sendQuota: gigabytes(mailbox.quotas.send),
            sendReceiveQuota: gigabytes(mailbox.quotas.sendReceive),
            percent:
              mailbox.quotas.send === null
                ? null
                : Math.round((mailbox.used / mailbox.quotas.send) * 1_000) / 10,
            items: mailbox.items,
            deletedSize: gigabytes(mailbox.deletedBytes),
            archive: mailbox.archive,
            lastActivity: mailbox.last,
            domain: principalDomain(mailbox.principal),
          },
        }));

    const count = (status: MailboxStatus) =>
      covered.filter((mailbox) => mailbox.status === status).length;
    const total = covered.reduce((sum, mailbox) => sum + mailbox.used, 0);
    const view = (id: string) => (concealed ? null : id);
    const unknown = count("Quota unknown");
    const noArchiveColumn = !index.has("has archive");
    const noTypeColumn = !index.has("recipient type");
    const refreshAt = index.get("report refresh date");
    const refreshed = lastActivityOf(
      refreshAt === undefined ? null : (table.rows[0]?.[refreshAt]?.trim() ?? null),
      context.now,
    );

    const limited = concealed ? null : scopeNote(context.scope);
    const left = concealed ? null : excludedNote(mailboxes.length - covered.length);
    const notes = [
      ...(limited ? [limited] : []),
      ...(left ? [left] : []),
      ...privacyNotes(setting, concealed),
      ...(concealed && context.scope && context.scope.size > 0
        ? ["The audit scope could not be applied: the report gives no names to sort by domain."]
        : []),
      `Sizes come from Microsoft's mailbox usage report for the last 90 days${refreshed.state === "known" && refreshed.at ? `, refreshed ${refreshed.at.slice(0, 10)}` : ""}. The report lags by 24 to 72 hours and sometimes by several days, so a mailbox can be fuller than it shows. A gigabyte here is 1,073,741,824 bytes.`,
      "A mailbox at a quota counts as over it. Near full is 90 percent of the warning quota or more, and near full without an archive is 80 percent or more for a mailbox known to have no archive.",
      ...(deleted > 0
        ? [
            `${plural(deleted, "deleted mailbox", "deleted mailboxes")} in the report ${deleted === 1 ? "is" : "are"} not listed.`,
          ]
        : []),
      ...(unknown > 0
        ? [
            `${plural(unknown, "mailbox has", "mailboxes have")} no warning quota in the report and ${unknown === 1 ? "shows" : "show"} Quota unknown.`,
          ]
        : []),
      ...(noTypeColumn
        ? ["The report has no Recipient Type column, so the type of every mailbox is unknown."]
        : []),
      ...(noArchiveColumn
        ? [
            "The report has no archive column, so no mailbox is called near full without an archive.",
          ]
        : []),
      ...(concealed
        ? []
        : [
            "Rows do not link to a person page: the report names mailboxes by sign-in name, not by object ID.",
          ]),
      ...(table.truncated
        ? [
            `The report has more than ${CSV_MAX_ROWS.toLocaleString("en-US")} rows; the rest was not read.`,
          ]
        : []),
      ...(table.skippedRows > 0
        ? [
            `${plural(table.skippedRows, "row in the report", "rows in the report")} could not be read and ${table.skippedRows === 1 ? "is" : "are"} left out.`,
          ]
        : []),
      ...(unreadable > 0
        ? [
            `${plural(unreadable, "mailbox", "mailboxes")} in the report ${unreadable === 1 ? "has" : "have"} no readable name or size and ${unreadable === 1 ? "is" : "are"} left out.`,
          ]
        : []),
      ...(tally.unreadable > 0
        ? [
            `${plural(tally.unreadable, "value in the report", "values in the report")} could not be read as a number, a date or a yes or no and ${tally.unreadable === 1 ? "shows" : "show"} as blank. A quota that cannot be read counts as missing.`,
          ]
        : []),
    ];

    return {
      rows,
      complete:
        !table.truncated && table.skippedRows === 0 && unreadable === 0 && tally.unreadable === 0,
      ...(concealed
        ? {}
        : { domains: domainInventory(mailboxes.map((mailbox) => mailbox.principal)) }),
      summary: [
        {
          label: "Cannot send or receive",
          value: count("Cannot send or receive"),
          tone: count("Cannot send or receive") > 0 ? "critical" : "good",
          view: view("cannot-send-receive"),
        },
        {
          label: "Cannot send",
          value: count("Cannot send"),
          tone: count("Cannot send") > 0 ? "critical" : "good",
          view: view("cannot-send"),
        },
        {
          label: "Over warning",
          value: count("Over warning"),
          tone: count("Over warning") > 0 ? "warning" : "good",
          view: view("over-warning"),
        },
        {
          label: "Near full",
          value: count("Near full"),
          tone: count("Near full") > 0 ? "warning" : "good",
          view: view("near-full"),
        },
        {
          label: "Near full, no archive",
          value: count("Near full, no archive"),
          tone: count("Near full, no archive") > 0 ? "warning" : "good",
          view: view("near-full-no-archive"),
        },
        {
          label: "Total storage used (GB)",
          value: Math.round(total / GIB),
          tone: "neutral",
          view: null,
        },
      ],
      notes,
    };
  },
};

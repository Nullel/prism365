import type { AccessAuditResult } from "@prism365/contracts";
import { number, relativeTime } from "./common.tsx";

// What the Access tab says about an account from each audit's stored rows: a few fields in a line,
// not the audit's table. Which columns arrive is ACCESS_AUDIT_FIELDS in the contract; this decides
// how they read. Pure, so every wording is a test.

export type Tone = "good" | "warning" | "bad";

// One piece of a line: a label or sentence, or a date shown as "3 h ago", with a colour for a
// good or bad state. The colour is always accompanied by words.
export interface SummaryPart {
  prefix?: string;
  text?: string;
  date?: string;
  tone?: Tone;
}
export interface SummaryLine {
  // The thing the line is about (a role, a device), when an audit has several.
  title: string | null;
  // A role that can change the tenant, marked in the shared severity style.
  privileged?: boolean;
  parts: SummaryPart[];
}

type Row = Extract<AccessAuditResult, { status: "ok" }>["rows"][number];
type Value = Row["cells"][number]["value"];

const cellValue = (row: Row, key: string): Value =>
  row.cells.find((cell) => cell.key === key)?.value ?? null;
const textOf = (row: Row, key: string): string | null => {
  const value = cellValue(row, key);
  return typeof value === "string" && value !== "" ? value : null;
};
const numberOf = (row: Row, key: string): number | null => {
  const value = cellValue(row, key);
  return typeof value === "number" ? value : null;
};
const flagOf = (row: Row, key: string): boolean | null => {
  const value = cellValue(row, key);
  return typeof value === "boolean" ? value : null;
};
const listOf = (row: Row, key: string): string[] => {
  const value = cellValue(row, key);
  return Array.isArray(value) ? value : [];
};
const compact = (parts: (SummaryPart | null)[]): SummaryPart[] =>
  parts.filter((part): part is SummaryPart => part !== null);
const gigabytes = (value: number) => number.format(Math.round(value * 10) / 10);
const plural = (count: number, one: string) =>
  `${number.format(count)} ${one}${count === 1 ? "" : "s"}`;

// Roles worth listing apart from the ones held now: the rest of the page already shows those.
const NOT_HELD_NOW = new Set(["Time-bound", "Activated", "Eligible", "Eligible time-bound"]);

function roleLines(rows: readonly Row[]): SummaryLine[] {
  return rows
    .filter((row) => NOT_HELD_NOW.has(textOf(row, "held") ?? ""))
    .map((row) => {
      const ends = textOf(row, "ends");
      return {
        title: textOf(row, "role") ?? "A role",
        privileged: flagOf(row, "privileged") === true,
        parts: compact([
          textOf(row, "held") ? { text: textOf(row, "held") as string } : null,
          ends ? { prefix: "ends", date: ends } : null,
        ]),
      };
    });
}

function methodLines(rows: readonly Row[]): SummaryLine[] {
  const row = rows[0];
  if (!row) return [];
  const registered = flagOf(row, "registered");
  const methods = listOf(row, "methods");
  return [
    {
      title: null,
      parts: compact([
        registered === null
          ? null
          : registered
            ? { text: "MFA registered", tone: "good" }
            : { text: "No MFA registered", tone: "bad" },
        textOf(row, "strength")
          ? { prefix: "Strongest method", text: textOf(row, "strength") as string }
          : null,
        methods.length > 0 ? { prefix: "Methods", text: methods.join(", ") } : null,
      ]),
    },
  ];
}

function complianceTone(state: string): Tone {
  if (state === "Compliant") return "good";
  return state === "Not compliant" || state === "Error" ? "bad" : "warning";
}

function deviceLines(rows: readonly Row[]): SummaryLine[] {
  return rows.map((row) => {
    const system = [textOf(row, "os"), textOf(row, "osVersion")].filter(Boolean).join(" ");
    const compliance = textOf(row, "compliance");
    const checkIn = textOf(row, "lastCheckIn");
    const encrypted = flagOf(row, "encrypted");
    return {
      title: textOf(row, "device") ?? "A device",
      parts: compact([
        system ? { text: system } : null,
        compliance ? { text: compliance, tone: complianceTone(compliance) } : null,
        checkIn ? { prefix: "last check-in", date: checkIn } : null,
        encrypted === null
          ? null
          : encrypted
            ? { text: "Encrypted" }
            : { text: "Not encrypted", tone: "warning" },
        textOf(row, "jailbroken") === "Yes" ? { text: "Jailbroken or rooted", tone: "bad" } : null,
      ]),
    };
  });
}

function mailboxLines(rows: readonly Row[]): SummaryLine[] {
  const row = rows[0];
  if (!row) return [];
  const used = numberOf(row, "used");
  const quota = numberOf(row, "sendQuota");
  const percent = numberOf(row, "percent");
  const archive = flagOf(row, "archive");
  const status = textOf(row, "status");
  const size =
    used === null
      ? null
      : quota === null
        ? `Used ${gigabytes(used)} GB`
        : `Used ${gigabytes(used)} of ${gigabytes(quota)} GB${percent === null ? "" : ` (${gigabytes(percent)}%)`}`;
  return [
    {
      title: null,
      parts: compact([
        size ? { text: size } : null,
        archive === null ? null : { text: archive ? "Has an archive" : "No archive" },
        textOf(row, "recipientType") ? { text: textOf(row, "recipientType") as string } : null,
        status && status !== "OK" ? { text: status, tone: "warning" } : null,
      ]),
    },
  ];
}

function passwordLines(rows: readonly Row[]): SummaryLine[] {
  const row = rows[0];
  if (!row) return [];
  const changed = textOf(row, "lastChanged");
  const age = numberOf(row, "days");
  const policy = textOf(row, "policy");
  const status = textOf(row, "status");
  return [
    {
      title: null,
      parts: compact([
        changed ? { prefix: "Last changed", date: changed } : null,
        age === null ? null : { text: `${plural(age, "day")} old` },
        policy ? { prefix: "Expiry policy", text: policy } : null,
        flagOf(row, "neverExpires") === true
          ? { text: "Never expires for this account", tone: "warning" }
          : null,
        flagOf(row, "weak") === true
          ? { text: "Strong password not required", tone: "warning" }
          : null,
        status && status !== "OK" ? { text: status, tone: "warning" } : null,
      ]),
    },
  ];
}

function riskLines(rows: readonly Row[]): SummaryLine[] {
  const row = rows[0];
  if (!row) return [];
  const level = textOf(row, "level");
  const updated = textOf(row, "updated");
  return [
    {
      title: null,
      parts: compact([
        level ? { text: `${level} risk`, tone: level === "Low" ? "warning" : "bad" } : null,
        textOf(row, "state") ? { prefix: "State", text: textOf(row, "state") as string } : null,
        textOf(row, "detail") ? { text: textOf(row, "detail") as string } : null,
        updated ? { prefix: "updated", date: updated } : null,
      ]),
    },
  ];
}

// Days of inactivity are said only from a month on; a shorter gap is ordinary.
export const INACTIVE_FROM_DAYS = 30;

type Rows = readonly Row[];
function activityLines(activity: Rows, inactive: Rows, reclaim: Rows): SummaryLine[] {
  const lines: SummaryLine[] = [];
  const row = activity[0];
  const days = inactive[0] ? numberOf(inactive[0], "days") : null;
  if (row) {
    const last = textOf(row, "lastSuccess");
    const service = (prefix: string, key: string): SummaryPart => {
      const at = textOf(row, key);
      return at ? { prefix, date: at } : { prefix, text: "none" };
    };
    lines.push({
      title: null,
      parts: compact([
        textOf(row, "verdict") ? { text: textOf(row, "verdict") as string } : null,
        last ? { prefix: "Last successful sign-in", date: last } : null,
        days !== null && days >= INACTIVE_FROM_DAYS
          ? { text: `Inactive for ${plural(days, "day")}`, tone: "warning" }
          : null,
      ]),
    });
    lines.push({
      title: "Last activity",
      parts: [
        service("Mail", "mailLast"),
        service("Teams", "teamsLast"),
        service("OneDrive", "oneDriveLast"),
        service("SharePoint", "sharePointLast"),
      ],
    });
  } else if (days !== null && days >= INACTIVE_FROM_DAYS) {
    lines.push({
      title: null,
      parts: [{ text: `Inactive for ${plural(days, "day")}`, tone: "warning" }],
    });
  }
  const candidate = reclaim[0];
  if (candidate) {
    const licenses = listOf(candidate, "licenses");
    lines.push({
      title: "License to reclaim",
      parts: compact([
        textOf(candidate, "status")
          ? { text: textOf(candidate, "status") as string, tone: "warning" }
          : null,
        licenses.length > 0 ? { text: licenses.join(", ") } : null,
      ]),
    });
  }
  return lines;
}

function guestLines(rows: Rows): SummaryLine[] {
  const row = rows[0];
  if (!row) return [];
  const invitation = textOf(row, "invitation");
  const invited = textOf(row, "invited");
  const days = numberOf(row, "days");
  const last = textOf(row, "lastSuccess");
  const status = textOf(row, "status");
  return [
    {
      title: null,
      parts: compact([
        invitation
          ? invitation === "Not accepted"
            ? { text: invitation, tone: "warning" }
            : { text: invitation }
          : null,
        invited ? { prefix: "Invited", date: invited } : null,
        last ? { prefix: "Last successful sign-in", date: last } : null,
        days !== null && days >= INACTIVE_FROM_DAYS
          ? { text: `Inactive for ${plural(days, "day")}`, tone: "warning" }
          : null,
        status && status !== "OK" ? { text: status, tone: "warning" } : null,
      ]),
    },
  ];
}

// "Outlook, Word and Teams".
export function joinWords(words: readonly string[]): string {
  if (words.length <= 1) return words.join("");
  return `${words.slice(0, -1).join(", ")} and ${words.at(-1)}`;
}

// One sentence: "Last active 2 days ago in Outlook, Word and Teams, on Windows and Web". The time
// is worded from the clock when the page draws, so a section drawn later reads later.
function appUsageLines(rows: Rows): SummaryLine[] {
  const row = rows[0];
  if (!row) return [];
  const last = textOf(row, "lastActivity");
  const apps = listOf(row, "apps");
  const platforms = listOf(row, "platforms");
  if (!last)
    return [{ title: null, parts: [{ text: "No app activity in 90 days", tone: "warning" }] }];
  const where = platforms.length > 0 ? `, on ${joinWords(platforms)}` : "";
  return [
    {
      title: null,
      parts: [
        {
          text: `Last active ${relativeTime(last)}${apps.length > 0 ? ` in ${joinWords(apps)}` : ""}${where}`,
        },
      ],
    },
  ];
}

// Every role the audit lists for the account, held now or not, for the summary panel. The Access
// tab lists the roles held now with the groups and keeps only the others here (roleLines).
export function heldRoleLines(results: ReadonlyMap<string, AccessAuditResult>): SummaryLine[] {
  const result = results.get("directory.privileged-roles");
  const rows = result?.status === "ok" ? result.rows : [];
  return rows.map((row) => {
    const ends = textOf(row, "ends");
    return {
      title: textOf(row, "role") ?? "A role",
      privileged: flagOf(row, "privileged") === true,
      parts: compact([
        textOf(row, "held") ? { text: textOf(row, "held") as string } : null,
        ends ? { prefix: "ends", date: ends } : null,
      ]),
    };
  });
}

const builders: Record<string, (rows: Rows) => SummaryLine[]> = {
  "guests.review": guestLines,
  "usage.m365-apps": appUsageLines,
  "directory.privileged-roles": roleLines,
  "mfa.registration": methodLines,
  "devices.managed": deviceLines,
  "mail.mailbox-storage": mailboxLines,
  "accounts.passwords": passwordLines,
  "identity.risky-users": riskLines,
};

export function summaryLines(
  results: ReadonlyMap<string, AccessAuditResult>,
  auditId: string,
): SummaryLine[] {
  const rowsOf = (id: string): Rows => {
    const result = results.get(id);
    return result?.status === "ok" ? result.rows : [];
  };
  if (auditId === "accounts.activity")
    return activityLines(
      rowsOf("accounts.activity"),
      rowsOf("accounts.inactive"),
      rowsOf("licenses.reclaim"),
    );
  return builders[auditId]?.(rowsOf(auditId)) ?? [];
}

// A stable key for a part or a line, from what it says.
export const partKey = (part: SummaryPart) =>
  `${part.prefix ?? ""}|${part.text ?? ""}|${part.date ?? ""}`;
export const lineKey = (line: SummaryLine) =>
  `${line.title ?? ""}|${line.parts.map(partKey).join("/")}`;

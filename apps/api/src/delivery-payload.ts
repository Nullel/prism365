import { isIPv6 } from "node:net";
import { cleanText, type DestinationKind, type Severity } from "@prism365/contracts";
import { ipv6Groups } from "./delivery-address.ts";

// What an alert says, and how each destination receives it (docs/alerts.md, ADR-0011).
//
// By default an alert carries only its title, severity, counts, the subject's label and a link
// back to Prism365. A destination's "include details" switch adds evidence fields. A full IP
// address never leaves unless the destination is email (whose recipients are always in
// ALERT_EMAIL_DOMAINS) AND its details are on: everywhere else, every address in every string,
// the subject label and the summary included, is cut to its network (/24 or /48).
//
// Provider text is untrusted. Teams and Slack would render some of it as formatting, links or
// mentions, so it is neutralised for them; the generic JSON and email carry plain text.

export const MAX_ALERT_ITEMS = 10;
const MAX_DETAIL_FIELDS = 12;

export interface AlertItem {
  title: string;
  severity: Severity | null;
  subject: string | null;
  // Fixed words such as "New" or "Escalated", never provider text.
  note: string | null;
  details: Record<string, string | number | boolean | null>;
}

export interface AlertMessage {
  kind: "finding" | "audit" | "directory" | "collection" | "test" | "export" | "summary";
  // Prism365's own headline, such as "3 new findings at high or above".
  title: string;
  severity: Severity | null;
  // How many things the alert covers; items lists at most MAX_ALERT_ITEMS of them.
  count: number;
  items: AlertItem[];
  // Lines of Prism365's own text, for alerts that are not a list.
  lines: string[];
  // Path and hash on PUBLIC_ORIGIN, such as "/#/signins".
  link: string;
}

export interface DestinationView {
  kind: DestinationKind;
  includeDetails: boolean;
}

const severityWords: Record<Severity, string> = {
  info: "Info",
  low: "Low",
  medium: "Medium",
  high: "High",
  critical: "Critical",
};

// One line of plain text: provider text cleaned, every line break and control character turned
// into a space, and bounded.
export function line(value: string, max = 300): string {
  const flat = cleanText(value)
    .replace(/[\p{Cc}\p{Zl}\p{Zp}]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim();
  return flat.length > max ? `${flat.slice(0, max - 1)}…` : flat;
}

function maskV4(address: string): string {
  const parts = address.split(".");
  return `${Number(parts[0])}.${Number(parts[1])}.${Number(parts[2])}.0/24`;
}

function maskV6(address: string): string | null {
  const groups = ipv6Groups(address.replace(/%.*$/, ""));
  if (!groups) return null;
  // An IPv4-mapped address is cut as the IPv4 address it carries.
  if (groups.slice(0, 5).every((group) => group === 0) && groups[5] === 0xffff) {
    const high = groups[6] ?? 0;
    return `::ffff:${high >> 8}.${high & 0xff}.${(groups[7] ?? 0) >> 8}.0/24`;
  }
  return `${groups
    .slice(0, 3)
    .map((group) => group.toString(16))
    .join(":")}::/48`;
}

// Runs of digits and dots that could hold an IPv4 address, glued to anything or not.
const v4Pattern = /[0-9][0-9.]*[0-9]/g;
const quad = /^([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{1,3})/;

const octets = (parts: readonly string[]) => parts.every((part) => Number(part) <= 255);

// The first dotted quad of numbers up to 255 inside a run that is not one itself: with a digit
// glued on before ("9" and 67.178.80.106) or after (72.183.241.74 and "9"), or a dotted part on
// either side (1. and an address, an address and .7). null when there is none.
function innerQuad(run: string): string | null {
  for (let start = 0; start < run.length; start += 1) {
    const match = run.slice(start).match(quad);
    if (!match) continue;
    const [, a = "", b = "", c = "", d = ""] = match;
    // The last part may run on into more digits: try it shorter.
    for (let length = d.length; length > 0; length -= 1)
      if (octets([a, b, c, d.slice(0, length)])) return `${a}.${b}.${c}.${d.slice(0, length)}`;
  }
  return null;
}
// Anything that could be an IPv6 address: hex digits, colons and dots, with at least two colons.
const v6Pattern = /[0-9A-Fa-f:.]*:[0-9A-Fa-f:.]*:[0-9A-Fa-f:.]*(?:%[0-9A-Za-z]+)?/g;

// The address inside a run of hex digits, colons and dots, with what surrounds it: the run as it
// is first (so "::1" and "::ffff:c000:201" are found), then without the punctuation at its ends.
function addressIn(run: string): { before: string; address: string; after: string } | null {
  const trimmed = run.match(/^([.:]*)(.*?)([.:]*)$/);
  const candidates: [string, string, string][] = [["", run, ""]];
  const dotless = run.match(/^(.*?)(\.*)$/);
  if (dotless) candidates.push(["", dotless[1] ?? "", dotless[2] ?? ""]);
  if (trimmed) candidates.push([trimmed[1] ?? "", trimmed[2] ?? "", trimmed[3] ?? ""]);
  for (const [before, address, after] of candidates)
    if (isIPv6(address.replace(/%.*$/, ""))) return { before, address, after };
  return null;
}

// A run no single address could be found in, but which could still hold one: eight or more hex
// groups (a full address with a port or a prefix glued on, "2001:db8:1:2:3:4:5:6:443"), or a "::"
// with at least two groups. Times ("12:30:45") and MAC addresses have neither.
function addressLike(run: string): boolean {
  const groups = run.split(/[:.]/).filter((group) => /^[0-9A-Fa-f]{1,4}$/.test(group));
  return groups.length >= 8 || (run.includes("::") && groups.length >= 2);
}

// The whole run cut to its first three groups, so no address inside it survives.
function maskRun(run: string): string {
  const groups = run
    .split(":")
    .filter((group) => /^[0-9A-Fa-f]{1,4}$/.test(group))
    .slice(0, 3)
    .map((group) => Number.parseInt(group, 16).toString(16));
  return `${groups.join(":")}::/48`;
}

// Every IPv4 and IPv6 address in the text, cut to its network. A dotted quad of numbers up to 255
// counts as an address even with leading zeros, which some parsers accept.
export function maskAddresses(text: string): string {
  // Percent-encoded colons and dots, as in a query string, are read as what they stand for.
  const decoded = text.replace(/%3a/gi, ":").replace(/%2e/gi, ".");
  // IPv6 first, so an IPv4-mapped address is cut once, as a whole.
  const v6Masked = decoded.replace(v6Pattern, (run) => {
    const found = addressIn(run);
    if (found) {
      const masked = maskV6(found.address);
      if (masked) return `${found.before}${masked}${found.after}`;
    }
    return addressLike(run) ? maskRun(run) : run;
  });
  return v6Masked.replace(v4Pattern, (match, offset: number, whole: string) => {
    const parts = match.split(".");
    if (parts.length === 4 && parts.every((part) => /^[0-9]{1,3}$/.test(part)) && octets(parts)) {
      // A /24 network is already cut.
      if (match.endsWith(".0") && whole.startsWith("/24", offset + match.length)) return match;
      return maskV4(match);
    }
    // Not one address, but it may hold one: the whole run is cut to that address's network.
    const inner = innerQuad(match);
    return inner ? maskV4(inner) : match;
  });
}

// The alert as this destination may receive it: details only when it asked for them, and full
// addresses only for email with details on.
export function minimise(message: AlertMessage, destination: DestinationView): AlertMessage {
  const fullAddresses = destination.kind === "email" && destination.includeDetails;
  const text = (value: string) => (fullAddresses ? value : maskAddresses(value));
  return {
    ...message,
    title: text(message.title),
    lines: message.lines.map(text),
    items: message.items.slice(0, MAX_ALERT_ITEMS).map((item) => ({
      title: text(item.title),
      severity: item.severity,
      subject: item.subject === null ? null : text(item.subject),
      note: item.note,
      details: destination.includeDetails
        ? Object.fromEntries(
            Object.entries(item.details)
              .slice(0, MAX_DETAIL_FIELDS)
              .map(([key, value]) => [key, typeof value === "string" ? text(value) : value]),
          )
        : {},
    })),
  };
}

function itemLine(item: AlertItem): string {
  const severity = item.severity ? `${severityWords[item.severity]}: ` : "";
  const subject = item.subject ? ` (${line(item.subject, 200)})` : "";
  const note = item.note ? `, ${item.note}` : "";
  return `${severity}${line(item.title, 160)}${subject}${note}`;
}

function detailLines(item: AlertItem): string[] {
  return Object.entries(item.details)
    .filter(([, value]) => value !== null && value !== "")
    .map(([key, value]) => `${key}: ${line(String(value), 200)}`);
}

// The alert as plain-text lines, without the link back, shared by every format but the generic
// JSON. Each line is one line: provider text never carries a line break into it.
export function textLines(message: AlertMessage): string[] {
  const lines = [line(message.title, 200)];
  if (message.severity) lines.push(`Severity: ${severityWords[message.severity]}`);
  for (const text of message.lines) lines.push(line(text, 500));
  for (const item of message.items) {
    lines.push(itemLine(item));
    for (const detail of detailLines(item)) lines.push(`  ${detail}`);
  }
  if (message.count > message.items.length && message.items.length > 0)
    lines.push(`And ${message.count - message.items.length} more in Prism365.`);
  return lines;
}

// Chat apps turn some characters into links, formatting or mentions. Square brackets lose their
// link meaning, URL schemes are broken so nothing in provider text is a one-click link, and the
// characters that start Slack's control sequences are escaped as Slack documents.
function defang(value: string): string {
  return value
    .replace(/\[/g, "(")
    .replace(/\]/g, ")")
    .replace(/\b([a-z][a-z0-9+.-]{1,20}):\/\//gi, "$1[:]//");
}

export function slackText(value: string): string {
  return defang(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/@(channel|here|everyone)\b/gi, "@\u200b$1");
}

// Teams Adaptive Card text supports a little Markdown, of which only links matter here: they are
// broken as for Slack. Bold or a list marker in provider text changes how a line looks, nothing
// more.
export function teamsText(value: string): string {
  return defang(value);
}

export function renderWebhook(
  kind: Exclude<DestinationKind, "email">,
  message: AlertMessage,
  origin: string,
  meta: { deliveryId: string; sentAt: string },
): string {
  const link = `${origin}${message.link}`;
  if (kind === "slack") {
    const lines = textLines(message).map(slackText);
    // The link back is Prism365's own, so it is the one link in the message.
    lines.push(`<${link}|Open Prism365>`);
    return JSON.stringify({
      text: lines.join("\n"),
      unfurl_links: false,
      unfurl_media: false,
    });
  }
  if (kind === "teams") {
    const body = textLines(message).map((text, index) => ({
      type: "TextBlock",
      text: teamsText(text),
      wrap: true,
      ...(index === 0 ? { weight: "Bolder", size: "Medium" } : {}),
    }));
    return JSON.stringify({
      type: "message",
      attachments: [
        {
          contentType: "application/vnd.microsoft.card.adaptive",
          contentUrl: null,
          content: {
            $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
            type: "AdaptiveCard",
            version: "1.4",
            body,
            actions: [{ type: "Action.OpenUrl", title: "Open Prism365", url: link }],
          },
        },
      ],
    });
  }
  // The generic format, documented in docs/alerts.md. Version 1.
  return JSON.stringify({
    version: 1,
    type: "prism365.alert",
    deliveryId: meta.deliveryId,
    sentAt: meta.sentAt,
    kind: message.kind,
    title: line(message.title, 200),
    severity: message.severity,
    count: message.count,
    lines: message.lines.map((text) => line(text, 500)),
    items: message.items.map((item) => ({
      title: line(item.title, 160),
      severity: item.severity,
      subject: item.subject === null ? null : line(item.subject, 200),
      note: item.note,
      details: Object.fromEntries(
        Object.entries(item.details).map(([key, value]) => [
          key,
          typeof value === "string" ? line(value, 200) : value,
        ]),
      ),
    })),
    link,
  });
}

export function renderEmail(
  message: AlertMessage,
  origin: string,
): { subject: string; text: string } {
  return {
    subject: `Prism365: ${line(message.title, 150)}`,
    text: `${[...textLines(message), `Open Prism365: ${origin}${message.link}`].join("\n")}\n\nThis message was sent by Prism365 to the addresses this alert destination lists.\n`,
  };
}

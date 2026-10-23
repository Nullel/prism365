import { z } from "zod";

// Microsoft Graph service communications, read as Microsoft returns them. Field names and the
// example payloads were checked against Microsoft Learn on 2026-10-22:
// https://learn.microsoft.com/en-us/graph/api/serviceannouncement-list-issues?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/serviceannouncement-list-messages?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/serviceannouncement-list-healthoverviews?view=graph-rest-1.0
// Microsoft's documentation shows enum values in more than one casing ("Incident", "incident"),
// so every comparison here ignores case. Long provider text is cut, never a reason to reject.

export const healthContext = {
  feature: "service.health.read",
  permission: "ServiceHealth.Read.All",
} as const;
export const messageContext = {
  feature: "service.messages.read",
  permission: "ServiceMessage.Read.All",
} as const;

// PostgreSQL text cannot hold U+0000, so it is removed from every provider string.
export const clean = (value: string) => value.replaceAll("\u0000", "");
const text = (limit: number) =>
  z
    .string()
    .nullable()
    .optional()
    .catch(null)
    .transform((value) => (value ? clean(value).slice(0, limit) : null));
const time = z
  .string()
  .datetime({ offset: true })
  .nullable()
  .optional()
  .catch(null)
  .transform((value) => value ?? null);
const words = z
  .array(z.string())
  .nullable()
  .optional()
  .catch(null)
  .transform((value) => (value ?? []).map((item) => item.slice(0, 120)).slice(0, 50));

export const overviewSchema = z.object({
  id: z.string().min(1).max(200).transform(clean),
  service: z
    .string()
    .min(1)
    .transform((value) => clean(value).slice(0, 120)),
  status: z.string().max(80).catch("unknown"),
});
export type ServiceOverview = z.infer<typeof overviewSchema>;

const postSchema = z.object({
  createdDateTime: time,
  postType: text(40),
  description: z
    .object({ content: text(20_000), contentType: text(20) })
    .nullable()
    .optional()
    .catch(null),
});

export const issueSchema = z.object({
  id: z.string().min(1).max(200).transform(clean),
  title: z
    .string()
    .min(1)
    .transform((value) => clean(value).trim().slice(0, 500) || "Untitled"),
  service: z
    .string()
    .min(1)
    .transform((value) => clean(value).trim().slice(0, 120) || "Unknown service"),
  feature: text(200),
  featureGroup: text(200),
  impactDescription: text(2_000),
  classification: z.string().max(40).catch("unknown"),
  status: z.string().max(80).catch("unknown"),
  origin: text(40),
  isResolved: z.boolean().catch(false),
  startDateTime: time,
  endDateTime: time,
  lastModifiedDateTime: z.string().datetime({ offset: true }),
  posts: z.array(postSchema).nullable().optional().catch(null),
});
export type ServiceIssue = z.infer<typeof issueSchema>;

export const messageSchema = z.object({
  id: z.string().min(1).max(200).transform(clean),
  title: z
    .string()
    .min(1)
    .transform((value) => clean(value).trim().slice(0, 500) || "Untitled"),
  category: text(60),
  severity: text(40),
  tags: words,
  services: words,
  isMajorChange: z.boolean().nullable().optional().catch(null),
  actionRequiredByDateTime: time,
  startDateTime: time,
  endDateTime: time,
  lastModifiedDateTime: z.string().datetime({ offset: true }),
  body: z
    .object({ content: text(40_000), contentType: text(20) })
    .nullable()
    .optional()
    .catch(null),
});
export type ServiceMessage = z.infer<typeof messageSchema>;

// "serviceRestored" or "ServiceRestored" -> "Service restored".
export function humanize(value: string | null | undefined): string {
  if (!value) return "Unknown";
  const spaced = value
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .replace(/[_-]+/g, " ")
    .trim()
    .toLowerCase();
  return spaced ? spaced[0]?.toUpperCase() + spaced.slice(1) : "Unknown";
}

export const lower = (value: string | null | undefined) => (value ?? "").trim().toLowerCase();

const entities: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
  nbsp: " ",
};

// Microsoft's post and message bodies are HTML. Prism365 shows text only: tags become line breaks
// or nothing, and nothing from the body is ever rendered as markup.
export function plainText(html: string | null | undefined, limit = 4_000): string {
  if (!html) return "";
  // Only the start of a long body can survive the limit, so only that much is processed.
  return clean(html.slice(0, limit * 6 + 2_000))
    .replace(/<(script|style)[^>]*>[\s\S]*?<\/\1>/gi, " ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|li|h[1-6]|tr)>/gi, "\n")
    .replace(/<li[^>]*>/gi, "• ")
    .replace(/<[^>]+>/g, "")
    .replace(/&(#\d+|#x[0-9a-f]+|[a-z]+);/gi, (match, code: string) => {
      const point = code.startsWith("#x")
        ? Number.parseInt(code.slice(2), 16)
        : code.startsWith("#")
          ? Number.parseInt(code.slice(1), 10)
          : Number.NaN;
      if (code.startsWith("#"))
        return Number.isInteger(point) && point > 0 && point <= 0x10ffff
          ? String.fromCodePoint(point)
          : " ";
      return entities[code.toLowerCase()] ?? match;
    })
    .replace(/[ \t\f\v]+/g, " ")
    .replace(/ *\n */g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim()
    .slice(0, limit);
}

const sectionLabels = [
  "title",
  "user impact",
  "more info",
  "current status",
  "scope of impact",
  "workaround",
  "start time",
  "end time",
  "estimated time to resolve",
  "root cause",
  "preliminary root cause",
  "next update by",
  "final status",
];

// Microsoft's incident posts are written as labeled sections ("User impact: ...", "Workaround:
// ..."). Returns the sections found, keyed by lower-case label. Text before any label is "".
export function postSections(content: string): Map<string, string> {
  const sections = new Map<string, string>();
  let current = "";
  const pattern = new RegExp(`^(${sectionLabels.join("|")})\\s*:\\s*(.*)$`, "i");
  for (const line of content.split("\n")) {
    const match = pattern.exec(line.trim());
    if (match?.[1]) {
      current = match[1].toLowerCase();
      sections.set(current, (match[2] ?? "").trim());
    } else if (line.trim()) {
      sections.set(current, [sections.get(current), line.trim()].filter(Boolean).join("\n"));
    }
  }
  return sections;
}

// The newest post, as plain text.
export function latestPost(issue: ServiceIssue): { at: string | null; text: string } | null {
  const posts = [...(issue.posts ?? [])].sort((a, b) =>
    (b.createdDateTime ?? "").localeCompare(a.createdDateTime ?? ""),
  );
  const post = posts[0];
  if (!post) return null;
  return { at: post.createdDateTime, text: plainText(post.description?.content ?? "", 8_000) };
}

// Admin center deep links. The pattern is community-documented, not a published Microsoft
// specification, and it opens only for an administrator signed in to this tenant.
export const incidentLink = (id: string) =>
  `https://admin.microsoft.com/#/servicehealth/:/alerts/${encodeURIComponent(id)}`;
export const messageLink = (id: string) =>
  `https://admin.microsoft.com/#/MessageCenter/:/messages/${encodeURIComponent(id)}`;

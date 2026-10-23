import type { Severity } from "@prism365/contracts";
import { GraphRequestError } from "@prism365/graph";
import { type AttentionInput, AttentionInputError, type CollectorContext } from "@prism365/modules";
import { z } from "zod";
import {
  healthContext,
  humanize,
  incidentLink,
  issueSchema,
  latestPost,
  lower,
  messageContext,
  messageLink,
  messageSchema,
  overviewSchema,
  plainText,
  postSections,
  type ServiceIssue,
  type ServiceMessage,
} from "./source.ts";

export const HEALTH_MODULE = "service.health";
export const MESSAGES_MODULE = "service.messages";
export const INCIDENT_KIND = "service.incident";
export const MESSAGE_KIND = "service.message";

const DAY = 86_400_000;
// Resolved incidents are read back this far, so a resolution is seen even if it happened while
// collection was stopped.
const RESOLVED_LOOKBACK_DAYS = 7;

function parseAll<T>(context: CollectorContext, items: unknown[], schema: z.ZodType<T>): T[] {
  const parsed: T[] = [];
  let skipped = 0;
  for (const item of items) {
    const result = schema.safeParse(item);
    if (result.success) parsed.push(result.data);
    else skipped += 1;
  }
  context.count.fetched(items.length);
  context.count.skipped(skipped);
  return parsed;
}

// One item a connector built badly (or that Microsoft sent in a shape no cleaning can save) is
// counted as skipped; it never stops the rest of the run from being stored.
async function report(context: CollectorContext, item: AttentionInput): Promise<boolean> {
  try {
    await context.attention.upsert(item);
    return true;
  } catch (error) {
    if (!(error instanceof AttentionInputError)) throw error;
    context.count.skipped(1);
    return false;
  }
}

async function saveSuccess(context: CollectorContext) {
  await context.cursor.set(JSON.stringify({ v: 1, lastSuccess: context.now.toISOString() }));
}

export function incidentAttention(issue: ServiceIssue, observedAt: Date): AttentionInput {
  const active = !issue.isResolved;
  const incident = lower(issue.classification) === "incident";
  const classification = incident ? "Incident" : "Advisory";
  const status = humanize(issue.status);
  const post = latestPost(issue);
  const sections = post ? postSections(post.text) : new Map<string, string>();
  const scope = sections.get("scope of impact") || null;
  const impact =
    [issue.impactDescription, scope ? `Scope of impact: ${scope}` : null]
      .filter(Boolean)
      .join("\n\n") || null;
  const severity: Severity = !active
    ? "info"
    : incident
      ? lower(issue.status).includes("interruption")
        ? "critical"
        : "high"
      : "medium";
  return {
    kind: INCIDENT_KIND,
    ref: issue.id,
    class: incident ? "incident" : "advisory",
    title: issue.title,
    summary: (
      sections.get("user impact") ||
      issue.impactDescription ||
      post?.text ||
      issue.title
    ).slice(0, 4_000),
    services: [issue.service],
    provider: {
      status,
      labels: [classification, ...(issue.feature ? [issue.feature.slice(0, 80)] : [])],
      deadline: null,
      updatedAt: new Date(issue.lastModifiedDateTime),
      active,
      impact: impact ? impact.slice(0, 4_000) : null,
      workaround: (sections.get("workaround") || null)?.slice(0, 4_000) ?? null,
    },
    localLabels: [],
    severity,
    link: incidentLink(issue.id),
    meaningful: {
      Status: status,
      Classification: classification,
      Title: issue.title,
      Impact: issue.impactDescription ?? null,
    },
    observedAt,
  };
}

// Reads the services Microsoft reports for the tenant and every open incident or advisory, plus
// those resolved in the last week.
export async function collectHealth(context: CollectorContext): Promise<{ complete: boolean }> {
  const overview = await context.graph.getCollection(
    "admin/serviceAnnouncement/healthOverviews",
    healthContext,
    { schema: z.unknown(), maxPages: 10, maxRecords: 1_000 },
  );
  const services = parseAll(context, overview.items, overviewSchema);

  const since = new Date(context.now.getTime() - RESOLVED_LOOKBACK_DAYS * DAY);
  const read = (filter: string | null) =>
    context.graph.getCollection(
      filter
        ? `admin/serviceAnnouncement/issues?${new URLSearchParams({ $filter: filter })}`
        : "admin/serviceAnnouncement/issues",
      healthContext,
      { schema: z.unknown(), maxPages: 20, maxRecords: 2_000 },
    );
  let raw: unknown[];
  let complete = overview.complete;
  // Whether this run saw every issue Microsoft lists as open, which is what lets it close items
  // Microsoft no longer lists as open (for example ones resolved while collection was stopped).
  let openComplete: boolean;
  try {
    const open = await read("isResolved eq false");
    const recent = await read(`lastModifiedDateTime ge ${since.toISOString()}`);
    raw = [...open.items, ...recent.items];
    complete = complete && open.complete && recent.complete;
    openComplete = open.complete;
  } catch (error) {
    // A tenant or cloud that rejects the filters gets the whole list, filtered here instead.
    if (!(error instanceof GraphRequestError && error.status === 400)) throw error;
    const all = await read(null);
    raw = all.items;
    complete = complete && all.complete;
    openComplete = all.complete;
  }
  const parsedIssues = parseAll(context, raw, issueSchema);
  if (parsedIssues.length < raw.length) openComplete = false;
  const byId = new Map<string, ServiceIssue>();
  for (const issue of parsedIssues)
    if (!issue.isResolved || new Date(issue.lastModifiedDateTime) >= since)
      byId.set(issue.id, issue);
  const issues = [...byId.values()];

  const now = context.now.toISOString();
  await context.db(async (sql) => {
    const serviceTable = context.table("service");
    for (const service of services)
      await sql.query(
        `INSERT INTO ${serviceTable} (tenant_id, id, name, status, collected_at) VALUES ($1,$2,$3,$4,$5)
         ON CONFLICT (tenant_id, id) DO UPDATE SET name=EXCLUDED.name, status=EXCLUDED.status,
           collected_at=EXCLUDED.collected_at`,
        [context.tenantId, service.id, service.service, service.status, now],
      );
    if (overview.complete)
      await sql.query(`DELETE FROM ${serviceTable} WHERE tenant_id=$1 AND collected_at < $2`, [
        context.tenantId,
        now,
      ]);
    const issueTable = context.table("issue");
    let inserted = 0;
    for (const issue of issues) {
      const post = latestPost(issue);
      const sections = post ? postSections(post.text) : new Map<string, string>();
      const result = await sql.query(
        `INSERT INTO ${issueTable} (tenant_id, id, service, feature, feature_group, title,
           impact_description, classification, status, origin, is_resolved, start_at, end_at,
           last_modified_at, latest_post_at, latest_post, scope_of_impact, workaround, collected_at)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19)
         ON CONFLICT (tenant_id, id) DO UPDATE SET service=EXCLUDED.service, feature=EXCLUDED.feature,
           feature_group=EXCLUDED.feature_group, title=EXCLUDED.title,
           impact_description=EXCLUDED.impact_description, classification=EXCLUDED.classification,
           status=EXCLUDED.status, origin=EXCLUDED.origin, is_resolved=EXCLUDED.is_resolved,
           start_at=EXCLUDED.start_at, end_at=EXCLUDED.end_at,
           last_modified_at=EXCLUDED.last_modified_at, latest_post_at=EXCLUDED.latest_post_at,
           latest_post=EXCLUDED.latest_post, scope_of_impact=EXCLUDED.scope_of_impact,
           workaround=EXCLUDED.workaround, collected_at=EXCLUDED.collected_at`,
        [
          context.tenantId,
          issue.id,
          issue.service,
          issue.feature,
          issue.featureGroup,
          issue.title,
          issue.impactDescription,
          issue.classification,
          issue.status,
          issue.origin,
          issue.isResolved,
          issue.startDateTime,
          issue.endDateTime,
          issue.lastModifiedDateTime,
          post?.at ?? null,
          post?.text.slice(0, 8_000) ?? null,
          sections.get("scope of impact")?.slice(0, 2_000) ?? null,
          sections.get("workaround")?.slice(0, 2_000) ?? null,
          now,
        ],
      );
      inserted += result.rowCount;
    }
    context.count.inserted(inserted);
  });

  let reported = true;
  for (const issue of issues)
    if (!(await report(context, incidentAttention(issue, context.now)))) reported = false;
  if (openComplete && reported)
    await context.attention.retire(
      INCIDENT_KIND,
      issues.filter((issue) => !issue.isResolved).map((issue) => issue.id),
      "Microsoft no longer lists this as open.",
    );
  await saveSuccess(context);
  return { complete: complete && reported };
}

const retirementTag = (tags: readonly string[]) => tags.some((tag) => /retire/i.test(tag));
const retirementTitle = /\b(retir(e|es|ing|ement)|deprecat(e|ed|ion)|end of support)\b/i;

// Whether a Message Center post belongs in Admin attention, by Microsoft's own fields only: a
// deadline, a retirement tag, a major planned change or a high severity, and recent enough to
// matter. Returns null for posts that are informational only.
export function messageAttention(message: ServiceMessage, observedAt: Date): AttentionInput | null {
  const now = observedAt.getTime();
  const deadline = message.actionRequiredByDateTime
    ? new Date(message.actionRequiredByDateTime)
    : null;
  const modified = new Date(message.lastModifiedDateTime).getTime();
  const recent = modified >= now - 90 * DAY;
  const retirement = retirementTag(message.tags);
  const major = message.isMajorChange === true;
  const planned = lower(message.category) === "planforchange";
  const severe = ["high", "critical"].includes(lower(message.severity));
  const actionable = deadline !== null && deadline.getTime() >= now - 7 * DAY;
  if (!actionable && !(recent && (retirement || (major && planned) || severe))) return null;

  const itemClass = actionable
    ? "message.action"
    : retirement
      ? "message.retirement"
      : "message.change";
  const soon =
    deadline !== null && deadline.getTime() >= now && deadline.getTime() - now <= 14 * DAY;
  let severity: Severity = actionable ? (soon ? "high" : "medium") : retirement ? "medium" : "low";
  if (severe && severity === "low") severity = "medium";
  if (lower(message.severity) === "critical" && severity === "medium") severity = "high";
  const labels = [
    ...(actionable ? ["Action required"] : []),
    ...(message.category ? [humanize(message.category)] : []),
    ...(major ? ["Major change"] : []),
    ...(severe ? [`${humanize(message.severity)} severity`] : []),
    ...message.tags,
  ];
  return {
    kind: MESSAGE_KIND,
    ref: message.id,
    class: itemClass,
    title: message.title,
    summary: plainText(message.body?.content, 1_200),
    services: message.services.slice(0, 50),
    provider: {
      status: null,
      labels: [...new Set(labels)].map((label) => label.slice(0, 80)).slice(0, 20),
      deadline,
      updatedAt: new Date(message.lastModifiedDateTime),
      active: true,
      impact: null,
      workaround: null,
    },
    // Inferred by Prism365, not Microsoft's label: kept apart so it is never mistaken for one.
    localLabels:
      !retirement && retirementTitle.test(message.title)
        ? ["Possible retirement (from the title)"]
        : [],
    severity,
    link: messageLink(message.id),
    meaningful: {
      Title: message.title,
      "Action required by": deadline ? deadline.toISOString().slice(0, 10) : null,
      "Major change": major,
      Tags: [...message.tags].sort().join(", ") || null,
    },
    observedAt,
  };
}

// Reads every Message Center post, keeps the provider record, and turns the ones that need an
// administrator into notification-center items. Posts Microsoft stops listing become inactive.
export async function collectMessages(context: CollectorContext): Promise<{ complete: boolean }> {
  const page = await context.graph.getCollection(
    "admin/serviceAnnouncement/messages",
    messageContext,
    { schema: z.unknown(), maxPages: 60, maxRecords: 6_000 },
  );
  const messages = parseAll(context, page.items, messageSchema);
  // Closing posts or deleting stored ones needs a read that saw everything: complete, with no
  // unreadable records, and not empty (an empty answer is far likelier a fault than a tenant
  // with no posts at all).
  const whole = page.complete && messages.length === page.items.length && messages.length > 0;
  const now = context.now.toISOString();
  await context.db(async (sql) => {
    const table = context.table("post");
    let inserted = 0;
    for (const message of messages) {
      const result = await sql.query(
        `INSERT INTO ${table} (tenant_id, id, title, category, severity, tags, services,
           is_major_change, action_required_by, start_at, end_at, last_modified_at, body_text,
           collected_at)
         VALUES ($1,$2,$3,$4,$5,$6::jsonb,$7::jsonb,$8,$9,$10,$11,$12,$13,$14)
         ON CONFLICT (tenant_id, id) DO UPDATE SET title=EXCLUDED.title, category=EXCLUDED.category,
           severity=EXCLUDED.severity, tags=EXCLUDED.tags, services=EXCLUDED.services,
           is_major_change=EXCLUDED.is_major_change, action_required_by=EXCLUDED.action_required_by,
           start_at=EXCLUDED.start_at, end_at=EXCLUDED.end_at,
           last_modified_at=EXCLUDED.last_modified_at, body_text=EXCLUDED.body_text,
           collected_at=EXCLUDED.collected_at`,
        [
          context.tenantId,
          message.id,
          message.title,
          message.category,
          message.severity,
          JSON.stringify(message.tags),
          JSON.stringify(message.services),
          message.isMajorChange === true,
          message.actionRequiredByDateTime,
          message.startDateTime,
          message.endDateTime,
          message.lastModifiedDateTime,
          plainText(message.body?.content, 4_000) || null,
          now,
        ],
      );
      inserted += result.rowCount;
    }
    if (whole)
      await sql.query(`DELETE FROM ${table} WHERE tenant_id=$1 AND collected_at < $2`, [
        context.tenantId,
        now,
      ]);
    context.count.inserted(inserted);
  });

  const keep: string[] = [];
  let reported = true;
  for (const message of messages) {
    const item = messageAttention(message, context.now);
    if (!item) continue;
    keep.push(item.ref);
    if (!(await report(context, item))) reported = false;
  }
  if (whole && reported) {
    // Two different reasons, said as what they are: Microsoft stopped listing the post, or
    // Prism365's own selection rule no longer picks it.
    await context.attention.retire(
      MESSAGE_KIND,
      messages.map((message) => message.id),
      "No longer listed by Microsoft.",
    );
    await context.attention.retire(
      MESSAGE_KIND,
      keep,
      "No longer selected by Prism365: its deadline passed more than 7 days ago, or it is older than 90 days.",
    );
  }
  await saveSuccess(context);
  return { complete: page.complete && messages.length === page.items.length && reported };
}

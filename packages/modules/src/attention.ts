import { randomUUID } from "node:crypto";
import {
  type AttentionClass,
  type AttentionItem,
  type AttentionPreferences,
  type AttentionReviewState,
  attentionClassSchema,
  attentionPreferencesSchema,
  type Severity,
  severitySchema,
} from "@prism365/contracts";
import type { SqlClient } from "@prism365/storage";
import { z } from "zod";

// The notification center's store. Connectors report items with upsertAttention on every refresh;
// the store decides, deterministically, whether anything meaningful changed and whether that is
// worth a notification. Refreshing the same facts never creates a second item or a new unread.

export class AttentionInputError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "AttentionInputError";
  }
}

const meaningfulValue = z.union([
  z.string().max(4_000),
  z.number().finite(),
  z.boolean(),
  z.null(),
]);

const attentionInputSchema = z.object({
  kind: z.string().regex(/^[a-z][a-z0-9_.-]{1,60}$/),
  ref: z.string().min(1).max(200),
  class: attentionClassSchema,
  title: z.string().trim().min(1).max(500),
  summary: z.string().max(4_000),
  services: z.array(z.string().min(1).max(120)).max(50),
  provider: z.object({
    status: z.string().max(120).nullable(),
    labels: z.array(z.string().min(1).max(80)).max(20),
    deadline: z.date().nullable(),
    updatedAt: z.date(),
    active: z.boolean(),
    impact: z.string().max(4_000).nullable(),
    workaround: z.string().max(4_000).nullable(),
  }),
  localLabels: z.array(z.string().min(1).max(80)).max(10),
  severity: severitySchema,
  link: z
    .string()
    .url()
    .max(500)
    .refine((value) => value.startsWith("https://"), "Links must use HTTPS.")
    .nullable(),
  meaningful: z.record(z.string().min(1).max(60), meaningfulValue),
  observedAt: z.date(),
});

export type AttentionInput = z.input<typeof attentionInputSchema>;

// Provider text is cleaned and cut to size, never a reason to reject the item: one long or odd
// Microsoft record must not fail a run, which would stop every later record from being stored.
// Only structural mistakes in a connector (a bad kind, link or date) still throw.
const clip = (value: string, max: number) => value.replaceAll("\u0000", "").slice(0, max);
const list = (items: readonly string[], max: number, count: number) =>
  items
    .map((item) => clip(String(item), max).trim())
    .filter(Boolean)
    .slice(0, count);
const optional = (value: string | null, max: number) => (value === null ? null : clip(value, max));

function sanitize(raw: AttentionInput): AttentionInput {
  return {
    ...raw,
    ref: clip(raw.ref, 200),
    title: clip(raw.title, 500).trim() || clip(raw.ref, 200),
    summary: clip(raw.summary, 4_000),
    services: list(raw.services, 120, 50),
    provider: {
      ...raw.provider,
      status: optional(raw.provider.status, 120),
      labels: list(raw.provider.labels, 80, 20),
      impact: optional(raw.provider.impact, 4_000),
      workaround: optional(raw.provider.workaround, 4_000),
    },
    localLabels: list(raw.localLabels, 80, 10),
    meaningful: Object.fromEntries(
      Object.entries(raw.meaningful)
        .slice(0, 30)
        .map(([key, value]) => [
          clip(key, 60),
          typeof value === "string" ? clip(value, 4_000) : value,
        ]),
    ),
  };
}
export type AttentionChange = "created" | "changed" | "resolved" | "reopened" | "unchanged";

export const defaultAttentionPreferences: AttentionPreferences = attentionPreferencesSchema.parse(
  {},
);

const PREFERENCE_KEY = "attention";

export async function readAttentionPreferences(
  sql: SqlClient,
  tenantId: string,
): Promise<AttentionPreferences> {
  const result = await sql.query<{ value: unknown }>(
    "SELECT value FROM prism365_preference WHERE tenant_id=$1 AND key=$2",
    [tenantId, PREFERENCE_KEY],
  );
  const raw = result.rows[0]?.value;
  const parsed = attentionPreferencesSchema.safeParse(
    typeof raw === "string" ? JSON.parse(raw) : (raw ?? {}),
  );
  return parsed.success ? parsed.data : defaultAttentionPreferences;
}

export async function writeAttentionPreferences(
  sql: SqlClient,
  tenantId: string,
  preferences: AttentionPreferences,
  actorId: string,
): Promise<AttentionPreferences> {
  const value = attentionPreferencesSchema.parse(preferences);
  await sql.query(
    `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
     VALUES ($1,$2,$3::jsonb,now(),$4)
     ON CONFLICT (tenant_id, key) DO UPDATE SET value=EXCLUDED.value, updated_at=EXCLUDED.updated_at,
       updated_by=EXCLUDED.updated_by`,
    [tenantId, PREFERENCE_KEY, JSON.stringify(value), actorId],
  );
  return value;
}

const same = (a: string, b: string) => a.trim().toLowerCase() === b.trim().toLowerCase();

export function importantService(
  preferences: AttentionPreferences,
  services: readonly string[],
): boolean {
  return (
    preferences.services.length > 0 &&
    services.some((service) => preferences.services.some((wanted) => same(wanted, service)))
  );
}

// Whether an item of this class, for these services, should mark itself unread.
export function worthNotifying(
  preferences: AttentionPreferences,
  itemClass: AttentionClass,
  services: readonly string[],
): boolean {
  // A class this build has no switch for notifies, rather than staying silent.
  if (preferences.notify[itemClass] === false) return false;
  if (preferences.services.length === 0 || services.length === 0) return true;
  return importantService(preferences, services);
}

const shown = (value: unknown) =>
  value === null || value === undefined || value === ""
    ? "none"
    : typeof value === "boolean"
      ? value
        ? "yes"
        : "no"
      : String(value).slice(0, 80);

const full = (value: unknown) =>
  value === null || value === undefined || value === "" ? "" : String(value);

// "Status: Investigating to Service restored" for each meaningful field that changed. Values are
// compared whole; only the note shortens them.
function describeChanges(
  before: Record<string, unknown>,
  after: Record<string, unknown>,
): string[] {
  const keys = [...new Set([...Object.keys(before), ...Object.keys(after)])];
  return keys
    .filter((key) => full(before[key]) !== full(after[key]))
    .map((key) =>
      String(after[key] ?? "").length > 80 || String(before[key] ?? "").length > 80
        ? `${key} changed`
        : `${key}: ${shown(before[key])} to ${shown(after[key])}`,
    );
}

interface ExistingRow {
  id: string;
  meaningful: unknown;
  active: boolean;
  muted: boolean;
  review_state: AttentionReviewState;
  revision: number;
}

export async function upsertAttention(
  sql: SqlClient,
  tenantId: string,
  moduleId: string,
  raw: AttentionInput,
  preferences: AttentionPreferences,
): Promise<{ id: string; change: AttentionChange }> {
  const parsed = attentionInputSchema.safeParse(sanitize(raw));
  if (!parsed.success)
    throw new AttentionInputError("The attention item does not match its contract.");
  const input = parsed.data;
  const local = [
    ...input.localLabels,
    ...(importantService(preferences, input.services) ? ["Important service"] : []),
  ].slice(0, 10);
  const notify = worthNotifying(preferences, input.class, input.services);
  const seen = input.observedAt.toISOString();
  const common = [
    input.title,
    input.summary,
    JSON.stringify(input.services),
    input.provider.status,
    JSON.stringify(input.provider.labels),
    input.provider.deadline?.toISOString() ?? null,
    input.provider.updatedAt.toISOString(),
    input.provider.impact,
    input.provider.workaround,
    input.provider.active,
    JSON.stringify(local),
    input.severity,
    input.link,
    JSON.stringify(input.meaningful),
    seen,
  ];

  const inserted = await sql.query<{ id: string }>(
    `INSERT INTO prism365_attention (id, tenant_id, module, kind, ref, class, title, summary, services,
       provider_status, provider_labels, provider_deadline, provider_updated_at, provider_impact,
       provider_workaround, active, local_labels, severity, link, meaningful, revision, change_note,
       first_seen, last_seen, changed_at, unread, review_state, muted)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9::jsonb,$10,$11::jsonb,$12,$13,$14,$15,$16,$17::jsonb,$18,$19,
       $20::jsonb,1,NULL,$21,$21,$21,$22,'new',false)
     ON CONFLICT (tenant_id, module, kind, ref) DO NOTHING
     RETURNING id`,
    [randomUUID(), tenantId, moduleId, input.kind, input.ref, input.class, ...common, notify],
  );
  const created = inserted.rows[0];
  if (created) return { id: created.id, change: "created" };

  const current = await sql.query<ExistingRow>(
    `SELECT id, meaningful, active, muted, review_state, revision FROM prism365_attention
      WHERE tenant_id=$1 AND module=$2 AND kind=$3 AND ref=$4 FOR UPDATE`,
    [tenantId, moduleId, input.kind, input.ref],
  );
  const existing = current.rows[0];
  if (!existing) throw new Error("The attention item disappeared during its update.");
  const before =
    typeof existing.meaningful === "string"
      ? (JSON.parse(existing.meaningful) as Record<string, unknown>)
      : ((existing.meaningful ?? {}) as Record<string, unknown>);
  const diffs = describeChanges(before, input.meaningful);
  const resolved = existing.active && !input.provider.active;
  const reopened = !existing.active && input.provider.active;
  const change: AttentionChange = resolved
    ? "resolved"
    : reopened
      ? "reopened"
      : diffs.length > 0
        ? "changed"
        : "unchanged";

  if (change === "unchanged") {
    // Same facts: refresh what Microsoft wrote, but no new revision and no new notification.
    await sql.query(
      `UPDATE prism365_attention SET title=$3, summary=$4, services=$5::jsonb, provider_status=$6,
         provider_labels=$7::jsonb, provider_deadline=$8, provider_updated_at=$9, provider_impact=$10,
         provider_workaround=$11, active=$12, local_labels=$13::jsonb, severity=$14, link=$15,
         meaningful=$16::jsonb, last_seen=GREATEST(last_seen,$17), class=$18
       WHERE tenant_id=$1 AND id=$2`,
      [tenantId, existing.id, ...common, input.class],
    );
    return { id: existing.id, change };
  }

  // Only the provider reopening an incident puts it back to "not reviewed". A post that drops out
  // of the feed and comes back keeps the administrator's review.
  const resetReview = reopened && (input.class === "incident" || input.class === "advisory");
  const note = [resolved ? "Resolved by Microsoft." : reopened ? "Active again." : null, ...diffs]
    .filter(Boolean)
    .join(" ")
    .slice(0, 500);
  await sql.query(
    `UPDATE prism365_attention SET title=$3, summary=$4, services=$5::jsonb, provider_status=$6,
       provider_labels=$7::jsonb, provider_deadline=$8, provider_updated_at=$9, provider_impact=$10,
       provider_workaround=$11, active=$12, local_labels=$13::jsonb, severity=$14, link=$15,
       meaningful=$16::jsonb, last_seen=GREATEST(last_seen,$17), changed_at=$17,
       revision=revision+1, change_note=$18, class=$21,
       unread=CASE WHEN muted THEN unread ELSE (unread OR $19) END,
       review_state=CASE WHEN $20 THEN 'new' ELSE review_state END,
       review_changed_at=CASE WHEN $20 THEN NULL ELSE review_changed_at END,
       review_changed_by=CASE WHEN $20 THEN NULL ELSE review_changed_by END
     WHERE tenant_id=$1 AND id=$2`,
    [tenantId, existing.id, ...common, note, notify, resetReview, input.class],
  );
  return { id: existing.id, change };
}

// Items of one kind the provider stopped listing (an expired Message Center post) are no longer
// active. Only call this after a complete read of that kind.
export async function retireAttention(
  sql: SqlClient,
  tenantId: string,
  moduleId: string,
  kind: string,
  keep: readonly string[],
  observedAt: Date,
  // Says who decided: "No longer listed by Microsoft." only when Microsoft stopped listing it.
  note: string,
): Promise<number> {
  const result = await sql.query(
    `UPDATE prism365_attention SET active=false, change_note=$6,
       revision=revision+1, changed_at=$5
     WHERE tenant_id=$1 AND module=$2 AND kind=$3 AND active=true
       AND NOT (ref = ANY($4::text[]))`,
    [tenantId, moduleId, kind, [...keep], observedAt.toISOString(), note.slice(0, 500)],
  );
  return result.rowCount;
}

interface AttentionRow {
  id: string;
  module: string;
  kind: string;
  class: AttentionClass;
  ref: string;
  title: string;
  summary: string;
  services: unknown;
  provider_status: string | null;
  provider_labels: unknown;
  provider_deadline: Date | string | null;
  provider_updated_at: Date | string;
  provider_impact: string | null;
  provider_workaround: string | null;
  active: boolean;
  local_labels: unknown;
  severity: Severity;
  link: string | null;
  revision: number;
  change_note: string | null;
  first_seen: Date | string;
  last_seen: Date | string;
  changed_at: Date | string;
  unread: boolean;
  review_state: AttentionReviewState;
  review_changed_at: Date | string | null;
  muted: boolean;
}

const columns = `id, module, kind, class, ref, title, summary, services, provider_status,
  provider_labels, provider_deadline, provider_updated_at, provider_impact, provider_workaround,
  active, local_labels, severity, link, revision, change_note, first_seen, last_seen, changed_at,
  unread, review_state, review_changed_at, muted`;

const iso = (value: Date | string) => new Date(value).toISOString();
const json = (value: unknown): string[] => {
  const parsed = typeof value === "string" ? (JSON.parse(value) as unknown) : value;
  return Array.isArray(parsed)
    ? parsed.filter((item): item is string => typeof item === "string")
    : [];
};

function toItem(row: AttentionRow): AttentionItem {
  return {
    id: row.id,
    module: row.module,
    kind: row.kind,
    class: row.class,
    ref: row.ref,
    title: row.title,
    summary: row.summary,
    services: json(row.services),
    provider: {
      status: row.provider_status,
      labels: json(row.provider_labels),
      deadline: row.provider_deadline ? iso(row.provider_deadline) : null,
      updatedAt: iso(row.provider_updated_at),
      active: row.active,
      impact: row.provider_impact,
      workaround: row.provider_workaround,
    },
    localLabels: json(row.local_labels),
    severity: row.severity,
    link: row.link,
    revision: row.revision,
    changeNote: row.change_note,
    firstSeen: iso(row.first_seen),
    lastSeen: iso(row.last_seen),
    changedAt: iso(row.changed_at),
    unread: row.unread,
    review: {
      state: row.review_state,
      changedAt: row.review_changed_at ? iso(row.review_changed_at) : null,
    },
    muted: row.muted,
  };
}

export type AttentionFilter = "open" | "unread" | "all" | "muted";

// open: active, not done, not muted. unread: notifications not yet read. all: everything kept.
export async function listAttention(
  sql: SqlClient,
  tenantId: string,
  options: { filter?: AttentionFilter; kinds?: readonly string[]; limit?: number } = {},
): Promise<AttentionItem[]> {
  const limit = Math.min(Math.max(options.limit ?? 100, 1), 500);
  const filter = options.filter ?? "open";
  const where =
    filter === "unread"
      ? "unread = true AND muted = false"
      : filter === "muted"
        ? "muted = true"
        : filter === "open"
          ? "active = true AND review_state <> 'done' AND muted = false"
          : "true";
  const result = await sql.query<AttentionRow>(
    `SELECT ${columns} FROM prism365_attention
      WHERE tenant_id=$1 AND ${where} AND ($2::text[] IS NULL OR kind = ANY($2::text[]))
      ORDER BY unread DESC,
        CASE severity WHEN 'critical' THEN 4 WHEN 'high' THEN 3 WHEN 'medium' THEN 2 WHEN 'low' THEN 1 ELSE 0 END DESC,
        changed_at DESC, id
      LIMIT $3`,
    [tenantId, options.kinds ? [...options.kinds] : null, limit],
  );
  return result.rows.map(toItem);
}

// One module's items of one kind that are active or changed since a time, most recent first. The
// filter is applied before the limit, so read or reviewed items can never push an active one out.
export async function listAttentionWindow(
  sql: SqlClient,
  tenantId: string,
  options: { module: string; kind: string; since: Date; limit?: number },
): Promise<AttentionItem[]> {
  const limit = Math.min(Math.max(options.limit ?? 300, 1), 1_000);
  const result = await sql.query<AttentionRow>(
    `SELECT ${columns} FROM prism365_attention
      WHERE tenant_id=$1 AND module=$2 AND kind=$3 AND (active OR changed_at >= $4)
      ORDER BY active DESC, provider_updated_at DESC, id
      LIMIT $5`,
    [tenantId, options.module, options.kind, options.since.toISOString(), limit],
  );
  return result.rows.map(toItem);
}

export async function attentionCounts(
  sql: SqlClient,
  tenantId: string,
): Promise<{ unread: number; active: number }> {
  const result = await sql.query<{ unread: number; active: number }>(
    `SELECT count(*) FILTER (WHERE unread AND NOT muted)::int AS unread,
            count(*) FILTER (WHERE active AND review_state <> 'done' AND NOT muted)::int AS active
       FROM prism365_attention WHERE tenant_id=$1`,
    [tenantId],
  );
  return result.rows[0] ?? { unread: 0, active: 0 };
}

export async function getAttention(
  sql: SqlClient,
  tenantId: string,
  id: string,
): Promise<AttentionItem | null> {
  const result = await sql.query<AttentionRow>(
    `SELECT ${columns} FROM prism365_attention WHERE tenant_id=$1 AND id=$2`,
    [tenantId, id],
  );
  const row = result.rows[0];
  return row ? toItem(row) : null;
}

// Reviewing an item also reads it. Microsoft's status is untouched.
export async function setAttentionReview(
  sql: SqlClient,
  tenantId: string,
  id: string,
  state: AttentionReviewState,
  actorId: string,
): Promise<AttentionItem | null> {
  await sql.query(
    `UPDATE prism365_attention SET review_state=$3, review_changed_at=now(), review_changed_by=$4,
       unread=CASE WHEN $3 = 'new' THEN unread ELSE false END
     WHERE tenant_id=$1 AND id=$2`,
    [tenantId, id, state, actorId],
  );
  return getAttention(sql, tenantId, id);
}

export async function setAttentionMuted(
  sql: SqlClient,
  tenantId: string,
  id: string,
  muted: boolean,
): Promise<AttentionItem | null> {
  await sql.query(
    "UPDATE prism365_attention SET muted=$3, unread=CASE WHEN $3 THEN false ELSE unread END WHERE tenant_id=$1 AND id=$2",
    [tenantId, id, muted],
  );
  return getAttention(sql, tenantId, id);
}

export async function markAttentionRead(
  sql: SqlClient,
  tenantId: string,
  id: string | null,
): Promise<number> {
  const result = await sql.query(
    "UPDATE prism365_attention SET unread=false WHERE tenant_id=$1 AND unread=true AND ($2::uuid IS NULL OR id=$2)",
    [tenantId, id],
  );
  return result.rowCount;
}

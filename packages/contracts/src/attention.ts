import { z } from "zod";

// The notification center ("attention"). Every item comes from a connector (a module) and keeps
// two things apart: what the provider says (status, labels, deadline) and what the administrator
// did (reviewed, done, muted). Labels Prism365 inferred itself are kept in their own list.

const text = z.string().transform((value) => value.slice(0, 500));
const longText = z.string().transform((value) => value.slice(0, 4_000));
const time = z.string().datetime({ offset: true });

export const attentionSeveritySchema = z.enum(["info", "low", "medium", "high", "critical"]);
export const attentionReviewStateSchema = z.enum(["new", "reviewed", "done"]);
export type AttentionReviewState = z.infer<typeof attentionReviewStateSchema>;

// What kind of thing an item is, which decides the default notification preference.
export const attentionClassSchema = z.enum([
  "incident",
  "advisory",
  "message.action",
  "message.retirement",
  "message.change",
]);
export type AttentionClass = z.infer<typeof attentionClassSchema>;

export const attentionItemSchema = z.object({
  id: z.uuid(),
  module: z.string().min(1).max(80),
  kind: z.string().min(1).max(80),
  // Read tolerantly: a row written by a newer build with a class this build does not know must
  // not break the inbox. Connectors still write only known classes.
  class: z.string().min(1).max(60),
  ref: text,
  title: text,
  summary: longText,
  services: z.array(text),
  provider: z.object({
    status: text.nullable(),
    labels: z.array(text),
    deadline: time.nullable(),
    updatedAt: time,
    active: z.boolean(),
    // Microsoft's words about who is affected, verbatim, when it gives them.
    impact: longText.nullable(),
    workaround: longText.nullable(),
  }),
  localLabels: z.array(text),
  severity: attentionSeveritySchema,
  link: z.string().url().nullable(),
  revision: z.number().int().positive(),
  changeNote: text.nullable(),
  firstSeen: time,
  lastSeen: time,
  changedAt: time,
  unread: z.boolean(),
  review: z.object({
    state: attentionReviewStateSchema,
    changedAt: time.nullable(),
  }),
  muted: z.boolean(),
});
export type AttentionItem = z.infer<typeof attentionItemSchema>;

// Problems that stop Prism365 from checking something. Computed from collection state when asked,
// never stored, so they cannot go stale.
export const systemAlertSchema = z.object({
  id: z.string().min(1).max(120),
  severity: attentionSeveritySchema,
  title: text,
  detail: longText,
  // When the problem was last observed (for example the failed run's start), shown by the page.
  observedAt: time.nullable(),
  action: z.object({ label: text, href: z.string().max(500) }).nullable(),
});
export type SystemAlert = z.infer<typeof systemAlertSchema>;

export const attentionSourceSchema = z.object({
  module: z.string().min(1).max(80),
  title: text,
  state: z.enum(["available", "missing_permission", "not_enabled", "not_collecting"]),
  missing: z.array(text),
});
export type AttentionSource = z.infer<typeof attentionSourceSchema>;

export const attentionListSchema = z.object({
  mode: z.enum(["demo", "live"]),
  items: z.array(attentionItemSchema),
  system: z.array(systemAlertSchema),
  sources: z.array(attentionSourceSchema),
  counts: z.object({
    unread: z.number().int().nonnegative(),
    active: z.number().int().nonnegative(),
  }),
});
export type AttentionList = z.infer<typeof attentionListSchema>;

export const attentionReviewRequestSchema = z.object({ state: attentionReviewStateSchema });
export const attentionMuteRequestSchema = z.object({ muted: z.boolean() });

export const attentionPreferencesSchema = z.object({
  // Services to notify about. Empty means every service the tenant has.
  services: z.array(z.string().min(1).max(120)).max(200).default([]),
  // Which classes of item mark themselves unread when they appear or change.
  notify: z
    .object({
      incident: z.boolean().default(true),
      advisory: z.boolean().default(true),
      "message.action": z.boolean().default(true),
      "message.retirement": z.boolean().default(true),
      "message.change": z.boolean().default(false),
    })
    // Switches for classes a newer build added are kept, so saving here never drops them.
    .catchall(z.boolean())
    .default({
      incident: true,
      advisory: true,
      "message.action": true,
      "message.retirement": true,
      "message.change": false,
    }),
});
export type AttentionPreferences = z.infer<typeof attentionPreferencesSchema>;

export const attentionSettingsSchema = z.object({
  preferences: attentionPreferencesSchema,
  // Services Microsoft reports for this tenant, to choose from.
  services: z.array(text),
});
export type AttentionSettings = z.infer<typeof attentionSettingsSchema>;

// Microsoft service health, as the Overview and Microsoft updates show it. `state` says whether the check itself worked,
// separately from what it found.
export const serviceHealthStatusSchema = z.object({
  state: z.enum(["ok", "partial", "stale", "failed", "never", "missing_permission", "not_enabled"]),
  detail: text.nullable(),
  lastSuccess: time.nullable(),
  lastAttempt: time.nullable(),
  services: z.array(z.object({ name: text, status: text })),
  incidents: z.array(attentionItemSchema),
  messages: z.object({
    state: z.enum([
      "ok",
      "partial",
      "stale",
      "failed",
      "never",
      "missing_permission",
      "not_enabled",
    ]),
    lastSuccess: time.nullable(),
  }),
});
export type ServiceHealthStatus = z.infer<typeof serviceHealthStatusSchema>;

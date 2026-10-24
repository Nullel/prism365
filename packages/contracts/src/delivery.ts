import { z } from "zod";
import { attentionSeveritySchema } from "./attention.ts";
import { auditIdSchema } from "./audits.ts";
import { auditTileIdSchema } from "./insights.ts";
import { cleanText } from "./text.ts";

// Alerts (docs/alerts.md, ADR-0011): rules that watch findings, audit results and collection,
// and the destinations they send to. Prism365 never sends through Microsoft: email goes through
// the operator's own mail relay, configured by environment only, and webhooks go to Teams
// Workflows, Slack or any HTTPS endpoint the administrator names. Rules are one setting per
// tenant in prism365_preference, key "alerts.rules"; destinations and the delivery log are core
// tables (007-delivery.sql).

export const MAX_ALERT_RULES = 50;
export const MAX_DESTINATIONS = 20;
export const MAX_EMAIL_RECIPIENTS = 10;
export const MAX_RULE_DESTINATIONS = 5;
// A webhook address longer than this is refused. Teams Workflows addresses run to about 400.
export const MAX_WEBHOOK_URL_LENGTH = 2_048;

export const destinationKindSchema = z.enum(["email", "teams", "slack", "generic"]);
export type DestinationKind = z.infer<typeof destinationKindSchema>;

// A name the administrator gives a destination or rule: one line of plain text.
const nameSchema = z
  .string()
  .transform((value) =>
    cleanText(value)
      .replace(/[\p{Cc}\p{Zl}\p{Zp}]+/gu, " ")
      .trim(),
  )
  .pipe(z.string().min(1).max(80));

// The part of an address before the @: letters, digits and . _ + ' - only, with no dot at either
// end and no two dots together. Characters a relay may read as a route to another host (% and !,
// which Postfix honours by default, quotes, a second @) are refused, so an address can only ever
// mean a mailbox in the domain after its @.
export const EMAIL_LOCAL_PART = /^[A-Za-z0-9_+'-]+(?:\.[A-Za-z0-9_+'-]+)*$/;
const domainLabel = "(?!-)[A-Za-z0-9-]{1,63}(?<!-)";
const plainAddress = new RegExp(
  `^(?=[^@]{1,64}@)${EMAIL_LOCAL_PART.source.slice(1, -1)}@${domainLabel}(?:\\.${domainLabel})+$`,
);

// One mailbox, written plainly: no display name, comment, quoted local part or second address.
// ASCII only, so a look-alike domain can never pass the allow-list check by its letters.
export const emailAddressSchema = z
  .string()
  .trim()
  .max(254)
  .regex(plainAddress)
  .transform((value) => value.toLowerCase());

// Fixed categories, never a provider's or relay's message: they are stored and shown as is.
export const deliveryErrorSchema = z.enum([
  "invalid_url",
  "port_not_allowed",
  "dns_failed",
  "blocked_address",
  "connection_failed",
  "tls_failed",
  "timeout",
  "redirect",
  "rejected",
  "rate_limited",
  "server_error",
  "secret_unreadable",
  "email_not_configured",
  "recipient_not_allowed",
  "smtp_auth_failed",
  "smtp_rejected",
  "too_large",
  "destination_deleted",
  "ledger_failed",
  "queue_full",
  // More matching findings arrived at once than the buffer holds; the lowest went first.
  "dropped",
  // The relay stopped answering after the message was handed over, so it may have arrived.
  // Never tried again, so it is never sent twice.
  "outcome_unknown",
  "interrupted",
  "error",
]);
export type DeliveryError = z.infer<typeof deliveryErrorSchema>;

export const destinationSchema = z.object({
  id: z.uuid(),
  kind: destinationKindSchema,
  label: z.string().min(1).max(80),
  // The webhook's host name, the only part of its address shown after it is saved.
  host: z.string().max(253).nullable(),
  recipients: z.array(z.string().max(254)).max(MAX_EMAIL_RECIPIENTS),
  includeDetails: z.boolean(),
  // False when the stored address was encrypted with another key, so it cannot be used.
  readable: z.boolean(),
  // False when a recipient is refused now (outside ALERT_EMAIL_DOMAINS, or an address this
  // version no longer accepts); nothing is sent to it until it is added again.
  recipientsAllowed: z.boolean(),
  createdAt: z.string().datetime({ offset: true }),
});
export type Destination = z.infer<typeof destinationSchema>;

export const destinationCreateSchema = z.discriminatedUnion("kind", [
  z.object({
    kind: z.literal("email"),
    label: nameSchema,
    recipients: z
      .array(emailAddressSchema)
      .min(1)
      .max(MAX_EMAIL_RECIPIENTS)
      .refine((list) => new Set(list).size === list.length, "List each recipient once."),
    includeDetails: z.boolean().default(false),
  }),
  z.object({
    kind: z.enum(["teams", "slack", "generic"]),
    label: nameSchema,
    url: z.string().trim().min(1).max(MAX_WEBHOOK_URL_LENGTH),
    includeDetails: z.boolean().default(false),
  }),
]);
export type DestinationCreate = z.input<typeof destinationCreateSchema>;

// The answer to a new destination. The signing secret is shown this once and never again.
export const destinationCreatedSchema = z.object({
  destination: destinationSchema,
  signingSecret: z.string().nullable(),
});
export type DestinationCreated = z.infer<typeof destinationCreatedSchema>;

export const deliverySourceSchema = z.enum(["rule", "test", "export", "summary"]);
export type DeliverySource = z.infer<typeof deliverySourceSchema>;

export const deliveryRecordSchema = z.object({
  id: z.uuid(),
  destinationId: z.uuid(),
  source: deliverySourceSchema,
  ruleId: z.uuid().nullable(),
  items: z.number().int().nonnegative(),
  status: z.enum(["pending", "sent", "failed"]),
  attempts: z.number().int().min(0).max(3),
  lastError: deliveryErrorSchema.nullable(),
  createdAt: z.string().datetime({ offset: true }),
  updatedAt: z.string().datetime({ offset: true }),
});
export type DeliveryRecord = z.infer<typeof deliveryRecordSchema>;

// Rule ids a directory rule may name, as the directory changes module reports them.
export const directoryAlertRuleIds = [
  "role.privileged-assigned",
  "app.credential-added",
  "app.consent-granted",
  "policy.conditional-access-changed",
  "domain.federation-changed",
  "auth-method.removed-by-admin",
] as const;
export const directoryAlertRuleIdSchema = z.enum(directoryAlertRuleIds);

const moduleIdSchema = z.string().regex(/^[a-z][a-z0-9-]*(?:\.[a-z][a-z0-9-]*)*$/);

const ruleFields = {
  name: nameSchema,
  destinations: z
    .array(z.uuid())
    .min(1)
    .max(MAX_RULE_DESTINATIONS)
    .refine((list) => new Set(list).size === list.length, "Name each destination once."),
  // At most one delivery per rule and subject in this many minutes.
  quietMinutes: z
    .number()
    .int()
    .min(5)
    .max(7 * 24 * 60),
};

const findingRule = z.object({
  kind: z.literal("finding"),
  ...ruleFields,
  minimumSeverity: attentionSeveritySchema,
  // Empty means every module.
  modules: z.array(moduleIdSchema).max(10).default([]),
});
const auditRule = z
  .object({
    kind: z.literal("audit"),
    ...ruleFields,
    auditId: auditIdSchema,
    tileId: auditTileIdSchema,
    // above and below: the tile crosses the value. change: it moves by at least the value
    // against its previous comparable point.
    condition: z.enum(["above", "below", "change"]),
    value: z.number().int().min(0).max(10_000_000),
  })
  .refine((rule) => rule.condition !== "change" || rule.value >= 1, {
    message: "A change rule needs a change of at least 1.",
  });
const directoryRule = z.object({
  kind: z.literal("directory"),
  ...ruleFields,
  ruleIds: z
    .array(directoryAlertRuleIdSchema)
    .min(1)
    .max(directoryAlertRuleIds.length)
    .refine((list) => new Set(list).size === list.length, "Name each change once."),
});
const collectionRule = z
  .object({
    kind: z.literal("collection"),
    ...ruleFields,
    // Collection that should be on but has been stopped for this long. null: not watched.
    stoppedHours: z.number().int().min(1).max(168).nullable(),
    // A module reported a span it could not read.
    gaps: z.boolean(),
  })
  .refine((rule) => rule.stoppedHours !== null || rule.gaps, {
    message: "Watch for stopped collection, gaps, or both.",
  });

export const alertRuleInputSchema = z.discriminatedUnion("kind", [
  findingRule,
  auditRule,
  directoryRule,
  collectionRule,
]);
export type AlertRuleInput = z.input<typeof alertRuleInputSchema>;

const storedFields = z.object({
  id: z.uuid(),
  enabled: z.boolean(),
  createdAt: z.string().datetime({ offset: true }),
  // The administrator's object id. Deliveries for the rule are recorded in the ledger under it.
  createdBy: z.uuid(),
});

export const alertRuleSchema = z.intersection(alertRuleInputSchema, storedFields);
export type AlertRule = z.infer<typeof alertRuleSchema>;

export const alertRuleListSchema = z.object({
  rules: z.array(alertRuleSchema).max(MAX_ALERT_RULES),
});
export type AlertRuleList = z.infer<typeof alertRuleListSchema>;

export const alertRuleUpdateSchema = z.object({ enabled: z.boolean() });

// What the page needs to know about this installation's delivery settings. Never the relay's
// address, its user or its password.
export const alertSettingsSchema = z.object({
  email: z.object({
    configured: z.boolean(),
    domains: z.array(z.string().max(253)).max(50),
  }),
  webhookPorts: z.array(z.number().int().min(1).max(65_535)).max(20),
  // The encryption key lives only in memory, so saved webhooks stop working after a restart.
  temporaryKey: z.boolean(),
});
export type AlertSettings = z.infer<typeof alertSettingsSchema>;

// What a rule can name: the installed modules, and each audit with the tiles its newest run had.
export const alertChoicesSchema = z.object({
  modules: z.array(z.object({ id: moduleIdSchema, title: z.string().max(200) })).max(50),
  audits: z
    .array(
      z.object({
        id: auditIdSchema,
        title: z.string().max(200),
        tiles: z.array(z.object({ id: auditTileIdSchema, label: z.string().max(500) })).max(20),
      }),
    )
    .max(50),
});
export type AlertChoices = z.infer<typeof alertChoicesSchema>;

// Scheduled exports: after each scheduled run of an audit, its result goes to a destination.
// Stored in prism365_preference, key "alerts.exports".
export const alertExportSchema = z.object({ auditId: auditIdSchema, destinationId: z.uuid() });
export type AlertExport = z.infer<typeof alertExportSchema>;
export const alertExportChangeSchema = z.object({
  auditId: auditIdSchema,
  destinationId: z.uuid().nullable(),
});
// The This week summary, every Monday at 07:00 in SIGNINS_TIME_ZONE. Key "alerts.weekly".
export const alertWeeklyChangeSchema = z.object({ destinationId: z.uuid().nullable() });

// The largest CSV sent as an attachment; a larger one is sent as a link only.
export const MAX_EXPORT_ATTACHMENT_BYTES = 5 * 1024 * 1024;

export const alertOverviewSchema = z.object({
  settings: alertSettingsSchema,
  choices: alertChoicesSchema,
  exports: z.array(alertExportSchema).max(50),
  weekly: z.object({ destinationId: z.uuid().nullable(), timeZone: z.string().max(64) }),
  destinations: z.array(destinationSchema).max(MAX_DESTINATIONS),
  rules: z.array(alertRuleSchema).max(MAX_ALERT_RULES),
  // Stored rules this version keeps but does not run: made by a newer version (with fields this
  // one does not know), or unreadable.
  otherRules: z
    .array(
      z.object({
        id: z.string().max(100).nullable(),
        name: z.string().max(80).nullable(),
        newer: z.boolean(),
      }),
    )
    .max(MAX_ALERT_RULES),
  deliveries: z.array(deliveryRecordSchema).max(100),
});
export type AlertOverview = z.infer<typeof alertOverviewSchema>;

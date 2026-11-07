import { z } from "zod";
import { auditCellSchema, auditColumnKindSchema, auditIdSchema } from "./audits.ts";

// A person's access, in one answer (docs/people.md, "Access tab"): groups and roles, what they own,
// the Conditional Access policies that reach them, and what each stored audit last said about
// them. Every section stands alone. One that cannot be read says why and the others still come
// back, so a missing permission never blanks the page.

const text = z.string().max(500);
const time = z.string().datetime({ offset: true });
const count = z.number().int().nonnegative();

export const MAX_ACCESS_MEMBERSHIPS = 1_000;
export const MAX_ACCESS_OWNED = 200;
export const MAX_ACCESS_AUDIT_ROWS = 25;

// Why a section is not shown. A section the person's session cannot read names the pack and the
// Microsoft permission it needs, and whether the installation has the pack on at all.
export const accessNotGrantedSchema = z.object({
  status: z.literal("not_granted"),
  pack: z.string().max(100),
  permission: z.string().max(100),
  enabled: z.boolean(),
});
export const accessFailedSchema = z.object({ status: z.literal("failed"), message: text });

// The kinds a membership can be, in the order the page lists them.
export const accessMembershipKinds = [
  "microsoft365",
  "security",
  "mail-security",
  "distribution",
  "dynamic",
  "role",
  "administrative-unit",
  "hidden",
] as const;
export const accessMembershipKindSchema = z.enum(accessMembershipKinds);
export type AccessMembershipKind = z.infer<typeof accessMembershipKindSchema>;

export const accessMembershipSchema = z.object({
  id: z.string().max(200),
  kind: accessMembershipKindSchema,
  // null for an object Microsoft returned without a name (kind "hidden").
  name: text.nullable(),
  relation: z.enum(["direct", "nested"]),
  // A Microsoft 365 group has a shared mailbox of its own.
  hasGroupMailbox: z.boolean(),
  // A group whose members come from a rule. Also true for a dynamic Microsoft 365 group.
  dynamic: z.boolean(),
  // Role template id, which Conditional Access names roles by. Roles only.
  roleTemplateId: z.string().max(100).nullable(),
});
export type AccessMembership = z.infer<typeof accessMembershipSchema>;

export const accessMembershipsSchema = z.discriminatedUnion("status", [
  z.object({
    status: z.literal("ok"),
    readAt: time,
    entries: z.array(accessMembershipSchema).max(MAX_ACCESS_MEMBERSHIPS),
    // More memberships exist than were read.
    truncated: z.boolean(),
  }),
  accessNotGrantedSchema,
  accessFailedSchema,
]);

export const accessOwnedKinds = ["group", "team", "application", "service-principal"] as const;
export const accessOwnedSchema = z.discriminatedUnion("status", [
  z.object({
    status: z.literal("ok"),
    readAt: time,
    entries: z
      .array(
        z.object({
          id: z.string().max(200),
          kind: z.enum(accessOwnedKinds),
          name: text.nullable(),
        }),
      )
      .max(MAX_ACCESS_OWNED),
    truncated: z.boolean(),
  }),
  accessNotGrantedSchema,
  accessFailedSchema,
]);

// Why a policy includes or leaves out a person. Names come from the memberships section.
export const policyReasonSchema = z.object({
  kind: z.enum(["all", "person", "group", "role", "guests"]),
  id: z.string().max(200).nullable(),
  name: text.nullable(),
});
export type PolicyReason = z.infer<typeof policyReasonSchema>;

export const policyStateSchema = z.enum(["on", "report-only", "off", "unknown"]);

export const accessPolicySchema = z.object({
  id: z.string().max(200),
  name: text.nullable(),
  state: policyStateSchema,
  // What a sign-in must satisfy, in plain words.
  controls: text,
  effect: z.enum(["applies", "excluded"]),
  // Whether the policy's inclusions reach the person at all. An exclusion can name someone the
  // policy would not have reached.
  targeted: z.boolean(),
  includedBy: z.array(policyReasonSchema).max(20),
  excludedBy: z.array(policyReasonSchema).max(20),
});
export type AccessPolicy = z.infer<typeof accessPolicySchema>;

export const accessConditionalAccessSchema = z.discriminatedUnion("status", [
  z.object({
    status: z.literal("ok"),
    readAt: time,
    policies: z.array(accessPolicySchema).max(1_000),
    // Policies that do not reach the person, counted and not listed.
    unaffected: count,
    // Policies Microsoft returned in a shape Prism365 could not read.
    unreadable: count,
    // More policies exist than were read.
    truncated: z.boolean(),
    // false when group and role memberships were not read, so only policies that name everyone,
    // the person or guests could be matched.
    membershipsKnown: z.boolean(),
  }),
  accessNotGrantedSchema,
  accessFailedSchema,
]);

// The columns of each audit this page shows, in the order they are sent. Everything else the audit
// stores about the account stays on the audit's own page. The route sends only these, and the page
// builds its summary lines from them, so this is the one place that says what the tab shows.
export const ACCESS_AUDIT_FIELDS: Readonly<Record<string, readonly string[]>> = {
  "directory.privileged-roles": ["role", "privileged", "held", "ends"],
  "mfa.registration": ["registered", "strength", "methods"],
  "devices.managed": [
    "device",
    "os",
    "osVersion",
    "compliance",
    "lastCheckIn",
    "encrypted",
    "jailbroken",
  ],
  "mail.mailbox-storage": ["used", "sendQuota", "percent", "archive", "recipientType", "status"],
  "accounts.passwords": ["lastChanged", "days", "policy", "neverExpires", "weak", "status"],
  "identity.risky-users": ["level", "state", "detail", "updated"],
  "accounts.activity": [
    "verdict",
    "lastSuccess",
    "mailLast",
    "teamsLast",
    "oneDriveLast",
    "sharePointLast",
  ],
  "accounts.inactive": ["days"],
  "licenses.reclaim": ["status", "licenses"],
};

export const accessAuditCellSchema = z.object({
  key: z.string().max(40),
  label: text,
  kind: auditColumnKindSchema,
  value: auditCellSchema,
});
export const accessAuditRowSchema = z.object({
  key: z.string().max(200),
  cells: z.array(accessAuditCellSchema).max(40),
});

export const accessAuditResultSchema = z.discriminatedUnion("status", [
  z.object({
    status: z.literal("ok"),
    auditId: auditIdSchema,
    name: text,
    finishedAt: time.nullable(),
    // The account domains the run was limited to; null when it covered every domain.
    scope: z.array(z.string().max(253)).max(100).nullable(),
    // How the rows were found: by the account's id, or by its sign-in name (mailbox storage).
    matchedBy: z.enum(["account", "sign-in name"]),
    rows: z.array(accessAuditRowSchema).max(MAX_ACCESS_AUDIT_ROWS),
    // The run held more rows for this person than were sent.
    truncated: z.boolean(),
  }),
  // The audit has never finished a run.
  z.object({ status: z.literal("no_run"), auditId: auditIdSchema, name: text }),
  // The stored result could not be matched, with the reason in plain words.
  z.object({ status: z.literal("failed"), auditId: auditIdSchema, name: text, message: text }),
]);
export type AccessAuditResult = z.infer<typeof accessAuditResultSchema>;

export const accessAuditsSchema = z.discriminatedUnion("status", [
  z.object({ status: z.literal("ok"), results: z.array(accessAuditResultSchema).max(20) }),
  accessFailedSchema,
]);

export const personAccessSchema = z.object({
  memberships: accessMembershipsSchema,
  owned: accessOwnedSchema,
  conditionalAccess: accessConditionalAccessSchema,
  audits: accessAuditsSchema,
});
export type PersonAccess = z.infer<typeof personAccessSchema>;

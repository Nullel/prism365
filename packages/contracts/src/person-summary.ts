import { z } from "zod";
import { auditIdSchema } from "./audits.ts";
import { receiptStatusSchema } from "./handoff.ts";
import {
  accessAuditRowSchema,
  accessFailedSchema,
  combinedRisksSchema,
  MAX_ACCESS_AUDIT_ROWS,
  MAX_ACCESS_OWNED,
} from "./person-access.ts";

// Where Prism365 knows one person from, in one answer (docs/people.md, "Where they show up"): what
// each stored audit last said about them, the groups and teams they own, their open sign-in
// findings, the shared mailboxes they can use, recent directory changes and their latest handoff.
// All of it is read from what Prism365 already stored. Nothing here asks Microsoft anything, and
// every section stands alone: one that cannot be read says so and the others still come back.

const text = z.string().max(500);
const time = z.string().datetime({ offset: true });
const count = z.number().int().nonnegative();

// The audits that describe one account, in the order the page lists them.
export const SUMMARY_AUDIT_IDS = [
  "identity.risky-users",
  "directory.privileged-roles",
  "accounts.activity",
  "accounts.inactive",
  "licenses.reclaim",
  "mfa.registration",
  "accounts.passwords",
  "guests.review",
  "devices.managed",
  "mail.mailbox-storage",
  "usage.m365-apps",
] as const;

// Directory changes are counted over this many days.
export const SUMMARY_CHANGE_DAYS = 30;

const auditBase = { auditId: auditIdSchema, name: text };
const runFacts = {
  finishedAt: time.nullable(),
  // False when the run could not check everything.
  complete: z.boolean(),
  // The account domains the run was limited to; null when it covered every domain.
  scope: z.array(z.string().max(253)).max(100).nullable(),
  // How the rows were found: by the account's id, or by its sign-in name.
  matchedBy: z.enum(["account", "sign-in name"]),
};

// One person-level audit. "ok" holds the person's rows (the fields ACCESS_AUDIT_FIELDS names);
// "not_in_run" is a run that says nothing about them; "no_name" is an audit that finds accounts by
// sign-in name when Prism365 has none for this person.
export const summaryAuditSchema = z.discriminatedUnion("status", [
  z.object({
    status: z.literal("ok"),
    ...auditBase,
    ...runFacts,
    rows: z.array(accessAuditRowSchema).max(MAX_ACCESS_AUDIT_ROWS),
    // The run held more rows for this person than were sent.
    truncated: z.boolean(),
  }),
  z.object({ status: z.literal("not_in_run"), ...auditBase, ...runFacts }),
  z.object({ status: z.literal("no_run"), ...auditBase }),
  z.object({
    status: z.literal("no_name"),
    ...auditBase,
    finishedAt: time.nullable(),
    complete: z.boolean(),
  }),
  z.object({ status: z.literal("failed"), ...auditBase, message: text }),
]);
export type SummaryAudit = z.infer<typeof summaryAuditSchema>;

const findingSeveritySchema = z.enum(["info", "low", "medium", "high", "critical"]);

export const MAX_SUMMARY_FINDINGS = 200;

export const summarySigninFindingsSchema = z.discriminatedUnion("status", [
  z.object({
    status: z.literal("ok"),
    // Findings not yet resolved, open or acknowledged.
    open: count,
    // The highest severity among them; null when there are none.
    highest: findingSeveritySchema.nullable(),
    // The page reads at most MAX_SUMMARY_FINDINGS findings, so a count at the limit is "or more".
    more: z.boolean(),
  }),
  // The sign-in audit is not part of this installation.
  z.object({ status: z.literal("unavailable") }),
  accessFailedSchema,
]);

export const summaryOwnedGroupSchema = z.object({
  id: z.string().max(200),
  name: text,
  // Team, Microsoft 365 group, Security group and so on, as Groups and Teams names them.
  kind: text,
  // The group's Groups and Teams status: No owners, One owner, OK and so on.
  status: text,
  // The audit's tags: inactive, expiring, guests, one-owner, public-team and the rest.
  tags: z.array(z.string().max(40)).max(20),
});

export const summaryOwnedSchema = z.discriminatedUnion("status", [
  z.object({
    status: z.literal("ok"),
    finishedAt: time.nullable(),
    complete: z.boolean(),
    entries: z.array(summaryOwnedGroupSchema).max(MAX_ACCESS_OWNED),
    // More groups than were sent have this person as an owner.
    more: z.boolean(),
  }),
  // Groups and Teams has never finished a run.
  z.object({ status: z.literal("no_run") }),
  accessFailedSchema,
]);

export const summarySharedMailboxesSchema = z.discriminatedUnion("status", [
  z.object({
    status: z.literal("ok"),
    exportedAt: time,
    importedAt: time,
    // The export covers user mailboxes as well as shared ones.
    includesUserMailboxes: z.boolean(),
    // Mailboxes the export gives this person a right on, directly. Rights granted to a group are
    // not expanded here.
    mailboxes: count,
    // The list was cut at its limit.
    mailboxesMore: z.boolean(),
    // Who holds rights on this person's own mailbox, when they are a mailbox in the export; null
    // when they are not.
    holders: count.nullable(),
    holdersMore: z.boolean(),
  }),
  // No mailbox access export is imported.
  z.object({ status: z.literal("no_import") }),
  accessFailedSchema,
]);

export const summaryChangesSchema = z.discriminatedUnion("status", [
  z.object({
    status: z.literal("ok"),
    days: z.number().int().positive(),
    count,
    // The count was cut at its limit.
    more: z.boolean(),
  }),
  // The directory change feed is not part of this installation.
  z.object({ status: z.literal("unavailable") }),
  accessFailedSchema,
]);

export const summaryHandoffSchema = z.discriminatedUnion("status", [
  z.object({
    status: z.literal("ok"),
    // The newest thing an action tool reported doing to this account; null when none did.
    latest: z
      .object({
        state: receiptStatusSchema,
        action: text,
        tool: text,
        performedAt: time,
      })
      .nullable(),
  }),
  accessFailedSchema,
]);

export const personSummarySchema = z.object({
  combinedRisks: combinedRisksSchema,
  signinFindings: summarySigninFindingsSchema,
  audits: z.array(summaryAuditSchema).max(20),
  owned: summaryOwnedSchema,
  sharedMailboxes: summarySharedMailboxesSchema,
  changes: summaryChangesSchema,
  handoff: summaryHandoffSchema,
});
export type PersonSummary = z.infer<typeof personSummarySchema>;

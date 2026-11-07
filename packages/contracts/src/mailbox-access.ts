import { z } from "zod";
import { cleanReceiptText } from "./handoff.ts";
import { accessFailedSchema, accessNotGrantedSchema } from "./person-access.ts";

// Shared mailbox access from a read-only Exchange export (ADR-0012). Microsoft offers no read-only
// API for Full Access, Send As or Send on Behalf, so an administrator runs
// Export-MailboxAccess.ps1 in their own Exchange Online session and imports the file it writes.
// Prism365 never holds an Exchange credential and never calls Exchange.
//
// The file is untrusted input: anyone who can edit it can make it say anything. The format is
// strict (an unknown field anywhere is refused rather than half read), every string is bounded
// on the text as sent and cleaned, and the file has to agree with itself. What is stored is shown
// as the export's claim, with its date, and nothing ever acts on it.

export const MAILBOX_ACCESS_FORMAT = "prism365.mailbox-access";
// The import route's body limit, on that route only.
export const MAX_MAILBOX_ACCESS_BYTES = 25 * 1024 * 1024;
export const MAX_EXPORT_MAILBOXES = 20_000;
export const MAX_EXPORT_GRANTS = 200_000;
// An export older than this is refused: access that old is a guess, not a record.
export const MAILBOX_ACCESS_MAX_AGE_DAYS = 90;
// The pages warn when the stored export is older than this.
export const MAILBOX_ACCESS_STALE_DAYS = 30;
// How far ahead of the server's clock an export may be dated: the computer that ran the script
// and the server never agree exactly.
export const MAILBOX_ACCESS_CLOCK_SKEW_MS = 5 * 60_000;
// The most rows the person route answers in each list.
export const MAX_PERSON_MAILBOX_GRANTS = 500;
export const MAX_MAILBOX_HOLDERS = 500;

export const MAX_MAILBOX_NAME = 256;
export const MAX_TRUSTEE_NAME = 512;
export const MAX_ADDRESS = 320;

const DAY = 86_400_000;
const time = z.string().datetime({ offset: true });
const count = z.number().int().nonnegative();

// Line breaks, tabs and the other C0 and C1 controls, and the Unicode line and paragraph
// separators. A name or an address never needs one, and in a name they could make one line on the
// page read as two.
const CONTROLS = /[\p{Cc}\p{Zl}\p{Zp}]/gu;
// Code points that draw nothing (zero-width spaces, the Hangul filler, tag characters). They are
// kept as sent, but a name made only of them is no name.
const INVISIBLE = /\p{Default_Ignorable_Code_Point}/gu;

// Text from the export as it is kept: NUL removed, broken UTF-16 and direction controls dealt with
// (cleanReceiptText), other controls and line separators as spaces, trimmed.
export function cleanExportText(value: string): string {
  return cleanReceiptText(value).replace(CONTROLS, " ").trim();
}

// How many characters of the text show: without invisible code points and outer white space.
export function visibleLength(value: string): number {
  return value.replace(INVISIBLE, "").trim().length;
}

// The bound is checked on the text as sent, never cut, so nothing in the file is silently
// shortened. The minimum counts only characters that show.
const exportText = (min: number, max: number) =>
  z
    .string()
    .max(max)
    .transform(cleanExportText)
    .refine((value) => visibleLength(value) >= min, {
      message:
        min === 1 ? "is empty or shows nothing" : `is shorter than ${min} visible characters`,
    });

// Ids are directory object ids (Exchange's ExternalDirectoryObjectId): GUIDs, kept in lower case
// as Graph writes them, so one id written in either case is one object.
const guid = z.guid().transform((value) => value.toLowerCase());

// An address as Exchange reports PrimarySmtpAddress: one @ with something on each side and no
// white space. It is shown, never used to match anyone, so this is a sanity check, not RFC 5321.
const address = exportText(3, MAX_ADDRESS).refine((value) => /^[^\s@]+@[^\s@]+$/u.test(value), {
  message: "is not an email address",
});

// When the script ran, kept as UTC. PostgreSQL refuses time zone offsets beyond 15:59 and dates
// outside its range, so they are refused here rather than failing once stored.
const exportedAt = time
  .refine(
    (value) => {
      const offset = /([+-])(\d{2}):(\d{2})$/.exec(value);
      return !offset || Number(offset[2]) <= 15;
    },
    { message: "has a time zone offset beyond 15:59, which no time zone uses" },
  )
  .refine(
    (value) => {
      const at = Date.parse(value);
      return at >= Date.UTC(2000, 0, 1) && at < Date.UTC(2100, 0, 1);
    },
    { message: "is not between the years 2000 and 2099" },
  )
  .transform((value) => new Date(value).toISOString());

export const mailboxRights = ["FullAccess", "SendAs", "SendOnBehalf"] as const;
export const mailboxRightSchema = z.enum(mailboxRights);
export type MailboxRight = z.infer<typeof mailboxRightSchema>;

export const mailboxTypeSchema = z.enum(["SharedMailbox", "UserMailbox"]);
export type MailboxType = z.infer<typeof mailboxTypeSchema>;

// What the script found a trustee to be. "unresolved" is a value Exchange returned that did not
// resolve to exactly one recipient (a deleted account's SID, an ambiguous name): kept as the raw
// text, never guessed.
export const trusteeKindSchema = z.enum(["user", "group", "other", "unresolved"]);
export type TrusteeKind = z.infer<typeof trusteeKindSchema>;

export const exportMailboxSchema = z.strictObject({
  id: guid,
  address,
  displayName: exportText(0, MAX_MAILBOX_NAME),
  type: mailboxTypeSchema,
});

export const exportTrusteeSchema = z
  .strictObject({
    id: guid.nullable(),
    address: address.nullable(),
    name: exportText(1, MAX_TRUSTEE_NAME),
    kind: trusteeKindSchema,
  })
  // An unresolved trustee is only its raw text: an id or an address on it would be a guess.
  .refine((trustee) => trustee.kind !== "unresolved" || trustee.id === null, {
    message: "has an id, but the trustee is unresolved",
    path: ["id"],
  })
  .refine((trustee) => trustee.kind !== "unresolved" || trustee.address === null, {
    message: "has an address, but the trustee is unresolved",
    path: ["address"],
  })
  // People and groups are matched by id, so one without an id could never be found.
  .refine((trustee) => !(trustee.kind === "user" || trustee.kind === "group") || trustee.id, {
    message: "is missing: a user or group trustee needs its directory id",
    path: ["id"],
  });
export type ExportTrustee = z.infer<typeof exportTrusteeSchema>;

export const exportGrantSchema = z.strictObject({
  mailboxId: guid,
  right: mailboxRightSchema,
  trustee: exportTrusteeSchema,
});
export type ExportGrant = z.infer<typeof exportGrantSchema>;

// The key that makes two grants the same: one right on one mailbox to one trustee, by id, or by
// raw text when the trustee has no id.
const grantKey = (grant: ExportGrant) =>
  JSON.stringify([
    grant.mailboxId,
    grant.right,
    grant.trustee.id ?? null,
    grant.trustee.id ? null : grant.trustee.name,
  ]);

export const mailboxAccessExportSchema = z
  .strictObject({
    format: z.literal(MAILBOX_ACCESS_FORMAT),
    version: z.literal(1),
    tenantId: guid,
    exportedAt,
    includeUserMailboxes: z.boolean(),
    mailboxes: z.array(exportMailboxSchema).max(MAX_EXPORT_MAILBOXES),
    grants: z.array(exportGrantSchema).max(MAX_EXPORT_GRANTS),
    unresolved: count.max(MAX_EXPORT_GRANTS),
  })
  // The file has to agree with itself. Anything the script cannot have written means the file
  // was edited or broken, and a file like that is refused whole rather than half believed.
  .superRefine((file, context) => {
    const mailboxes = new Map<string, number>();
    for (const [index, mailbox] of file.mailboxes.entries()) {
      const first = mailboxes.get(mailbox.id);
      if (first !== undefined) {
        context.addIssue({
          code: "custom",
          path: ["mailboxes", index, "id"],
          message: `repeats mailboxes[${first}]`,
        });
        return;
      }
      mailboxes.set(mailbox.id, index);
      if (mailbox.type === "UserMailbox" && !file.includeUserMailboxes) {
        context.addIssue({
          code: "custom",
          path: ["mailboxes", index, "type"],
          message: "is UserMailbox, but includeUserMailboxes is false",
        });
        return;
      }
    }
    const grants = new Map<string, number>();
    let unresolved = 0;
    for (const [index, grant] of file.grants.entries()) {
      if (!mailboxes.has(grant.mailboxId)) {
        context.addIssue({
          code: "custom",
          path: ["grants", index, "mailboxId"],
          message: "is not a mailbox in this file",
        });
        return;
      }
      const key = grantKey(grant);
      const first = grants.get(key);
      if (first !== undefined) {
        context.addIssue({
          code: "custom",
          path: ["grants", index],
          message: `repeats grants[${first}]`,
        });
        return;
      }
      grants.set(key, index);
      if (grant.trustee.kind === "unresolved") unresolved += 1;
    }
    if (unresolved !== file.unresolved)
      context.addIssue({
        code: "custom",
        path: ["unresolved"],
        message: `says ${file.unresolved}, but ${unresolved} ${unresolved === 1 ? "grant has" : "grants have"} an unresolved trustee`,
      });
  });
export type MailboxAccessExport = z.infer<typeof mailboxAccessExportSchema>;

function pathText(path: readonly PropertyKey[]): string {
  let text = "";
  for (const part of path)
    text += typeof part === "number" ? `[${part}]` : `${text ? "." : ""}${String(part)}`;
  return text;
}

function valueAt(input: unknown, path: readonly PropertyKey[]): unknown {
  let current = input;
  for (const part of path) {
    if (current === null || typeof current !== "object") return undefined;
    current = (current as Record<PropertyKey, unknown>)[part];
  }
  return current;
}

const nullableFields = new Set(["id", "address"]);
const choices: Record<string, string> = {
  right: "FullAccess, SendAs or SendOnBehalf",
  kind: "user, group, other or unresolved",
  type: "SharedMailbox or UserMailbox",
};

// The first problem with an export, in plain words, naming where it is ("grants[3].right"). The
// file's own text is echoed only short and cleaned, so a message never grows with what was sent.
export function mailboxAccessProblem(input: unknown, error: z.ZodError): string {
  const issue = error.issues[0];
  if (!issue) return "The export is not valid.";
  const path = issue.path;
  const where = path.length > 0 ? pathText(path) : "The export";
  const field = String(path.at(-1) ?? "");
  const value = valueAt(input, path);
  if (path.length === 0 && issue.code === "invalid_type")
    return "The file is not a JSON object. Choose the file Export-MailboxAccess.ps1 wrote.";
  if (issue.code === "invalid_type" && value === undefined) return `${where} is missing.`;
  if (issue.code === "invalid_type")
    return `${where} should be a ${issue.expected}${nullableFields.has(field) ? " or null" : ""}.`;
  if (issue.code === "unrecognized_keys")
    return `${where} has a field Prism365 does not read: "${cleanExportText((issue.keys[0] ?? "").slice(0, 80))}". Check that the file came from Export-MailboxAccess.ps1 and was not edited.`;
  if (issue.code === "too_big")
    return issue.origin === "array"
      ? `${where} has more than ${Number(issue.maximum).toLocaleString("en-US")} entries.`
      : issue.origin === "number"
        ? `${where} is larger than ${Number(issue.maximum).toLocaleString("en-US")}.`
        : `${where} is longer than ${issue.maximum} characters.`;
  if (issue.code === "too_small")
    return issue.origin === "number" ? `${where} is below zero.` : `${where} is empty.`;
  if (issue.code === "invalid_value") {
    if (field === "format" && path.length === 1)
      return `This is not a Prism365 mailbox access export: format should be "${MAILBOX_ACCESS_FORMAT}".`;
    if (field === "version" && path.length === 1)
      return typeof value === "number" && Number.isFinite(value)
        ? `This export is version ${value}. This Prism365 reads version 1.`
        : "This export has no version Prism365 reads. This Prism365 reads version 1.";
    const choice = choices[field];
    if (choice) return `${where} should be ${choice}.`;
    return `${where} has a value Prism365 does not read.`;
  }
  if (issue.code === "invalid_format") {
    if (issue.format === "guid") return `${where} is not a GUID.`;
    if (issue.format === "datetime")
      return `${where} is not a date and time with a time zone, such as 2026-11-06T14:30:00Z.`;
    return `${where} is not in the expected form.`;
  }
  if (issue.code === "invalid_union") return `${where} is not valid.`;
  // The schema's own checks are written to follow the field's name.
  if (issue.code === "custom") return `${where} ${issue.message}.`;
  return `${where} is not valid.`;
}

// Why an export's date is refused, or null when it is acceptable: not ahead of the server's
// clock (beyond a few minutes), and not older than MAILBOX_ACCESS_MAX_AGE_DAYS.
export function exportedAtProblem(exportedAtIso: string, now: number): string | null {
  const at = Date.parse(exportedAtIso);
  if (Number.isNaN(at)) return "The export has no date Prism365 can read.";
  if (at > now + MAILBOX_ACCESS_CLOCK_SKEW_MS)
    return `The export is dated ${exportedAtIso}, which is in the future. Check the clock of the computer that ran the script, and export again.`;
  if (at < now - MAILBOX_ACCESS_MAX_AGE_DAYS * DAY)
    return `The export is from ${exportedAtIso.slice(0, 10)}, more than ${MAILBOX_ACCESS_MAX_AGE_DAYS} days ago. Run the script again for a current export.`;
  return null;
}

// True when the export is old enough that the pages should say so.
export function isStaleExport(exportedAtIso: string, now: number = Date.now()): boolean {
  return now - Date.parse(exportedAtIso) > MAILBOX_ACCESS_STALE_DAYS * DAY;
}

// ---------------------------------------------------------------------------------------------
// API responses
// ---------------------------------------------------------------------------------------------

// The stored import: when the export was taken and imported, by whom, and how much it holds.
export const mailboxAccessImportSchema = z.object({
  id: z.uuid(),
  exportedAt: time,
  importedAt: time,
  importedBy: z.object({ id: z.uuid(), name: z.string().max(500).nullable() }),
  includeUserMailboxes: z.boolean(),
  mailboxes: count,
  grants: count,
  unresolved: count,
});
export type MailboxAccessImport = z.infer<typeof mailboxAccessImportSchema>;

export const mailboxAccessStatusSchema = z.object({
  import: mailboxAccessImportSchema.nullable(),
});
export type MailboxAccessStatus = z.infer<typeof mailboxAccessStatusSchema>;

// The answer to an import. replaced says an earlier import was there and is gone.
export const mailboxAccessImportResultSchema = z.object({
  import: mailboxAccessImportSchema,
  replaced: z.boolean(),
});
export type MailboxAccessImportResult = z.infer<typeof mailboxAccessImportResultSchema>;

export const mailboxAccessClearSchema = z.object({ cleared: z.boolean() });

const storedMailboxSchema = z.object({
  id: z.guid(),
  address: z.string().max(MAX_ADDRESS),
  name: z.string().max(MAX_MAILBOX_NAME),
  type: mailboxTypeSchema,
});

export const storedTrusteeSchema = z.object({
  id: z.guid().nullable(),
  address: z.string().max(MAX_ADDRESS).nullable(),
  name: z.string().max(MAX_TRUSTEE_NAME),
  kind: trusteeKindSchema,
});
export type StoredTrustee = z.infer<typeof storedTrusteeSchema>;

// One right the person holds on a mailbox, and how: granted to the account itself, or to a group
// the person is in (directly or through nesting), named as the export names it.
export const personMailboxGrantSchema = z.object({
  mailbox: storedMailboxSchema,
  right: mailboxRightSchema,
  via: z.discriminatedUnion("kind", [
    z.object({ kind: z.literal("direct") }),
    z.object({ kind: z.literal("group"), id: z.guid(), name: z.string().max(MAX_TRUSTEE_NAME) }),
  ]),
});
export type PersonMailboxGrant = z.infer<typeof personMailboxGrantSchema>;

// Who holds a right on the person's own mailbox, when the person is a mailbox in the export.
export const mailboxHolderSchema = z.object({
  right: mailboxRightSchema,
  trustee: storedTrusteeSchema,
});
export type MailboxHolder = z.infer<typeof mailboxHolderSchema>;

export const personMailboxAccessSchema = z.object({
  // null when nothing has been imported; every list is then empty.
  import: mailboxAccessImportSchema.nullable(),
  // How the person's groups were read, for grants made to groups. Without them only grants to the
  // account itself are found. "not_read" when there is no import to match them against.
  groups: z.discriminatedUnion("status", [
    z.object({ status: z.literal("ok"), readAt: time, count, truncated: z.boolean() }),
    z.object({ status: z.literal("not_read") }),
    accessNotGrantedSchema,
    accessFailedSchema,
  ]),
  grants: z.array(personMailboxGrantSchema).max(MAX_PERSON_MAILBOX_GRANTS),
  grantsTruncated: z.boolean(),
  holders: z.array(mailboxHolderSchema).max(MAX_MAILBOX_HOLDERS),
  holdersTruncated: z.boolean(),
});
export type PersonMailboxAccess = z.infer<typeof personMailboxAccessSchema>;

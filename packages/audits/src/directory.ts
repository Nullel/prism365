import { cleanText, licenseName } from "@prism365/contracts";
import { GraphRequestError } from "@prism365/graph";
import { z } from "zod";
import type { AuditContext } from "./definition.ts";

// Reads shared by the account audits: every user in the directory with the fields the audits
// use, and the tenant's subscriptions for license names. Provider values are cut to size or
// dropped, never a reason to fail the run; a record that still cannot be read is counted and
// left out, and the run says so.
// Microsoft Graph v1.0 documentation verified 2026-10-22:
// https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/resources/signinactivity?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/subscribedsku-list?view=graph-rest-1.0

export { cleanText };

export const DAY_MS = 86_400_000;
// Microsoft keeps interactive sign-in dates from April 2020 and has recorded the last
// successful sign-in only since 1 December 2023, without backfilling it.
export const SIGNIN_RECORDS_SINCE = Date.parse("2020-04-01T00:00:00Z");
export const SUCCESS_RECORDED_SINCE = Date.parse("2023-12-01T00:00:00Z");
// Graph caps a page at 500 users when signInActivity is selected, 999 otherwise. The client
// reads at most 10,000 records per collection.
export const MAX_ACCOUNTS = 10_000;
export const INACTIVITY_THRESHOLDS = [7, 30, 60, 90, 120, 180, 365] as const;

const directoryContext = { feature: "directory.users.read", permission: "User.Read.All" } as const;
const skuContext = {
  feature: "tenant.licenses.read",
  permission: "LicenseAssignment.Read.All",
} as const;

const text = z
  .string()
  .transform((value) => cleanText(value.slice(0, 500)))
  .nullish()
  .catch(null)
  .transform((value) => value ?? null);
const flag = z
  .boolean()
  .nullish()
  .catch(null)
  .transform((value) => value ?? null);
const time = z
  .string()
  .datetime({ offset: true })
  .nullish()
  .catch(null)
  .transform((value) => value ?? null);
// Dates that decide how long an account counts as inactive are kept as Microsoft sent them, with
// a value that is not a string marked unreadable rather than silently becoming "never".
// activityOf decides what is readable.
const UNREADABLE = "unreadable";
const rawDate = z
  .string()
  .max(64)
  .nullish()
  .catch(UNREADABLE)
  .transform((value) => value ?? null);
const count = z
  .number()
  .int()
  .nonnegative()
  .nullish()
  .catch(null)
  .transform((value) => value ?? 0);

const userSchema = z.object({
  id: z
    .string()
    .min(1)
    .max(200)
    .regex(/^[A-Za-z0-9._@:-]+$/),
  displayName: text,
  userPrincipalName: text,
  mail: text,
  userType: text,
  accountEnabled: flag,
  createdDateTime: rawDate,
  onPremisesSyncEnabled: flag,
  department: text,
  jobTitle: text,
  externalUserState: text,
  externalUserStateChangeDateTime: time,
  assignedLicenses: z
    .array(z.object({ skuId: z.string().max(100).nullish() }))
    .max(500)
    .nullish()
    .catch(null)
    .transform((value) =>
      (value ?? []).map((item) => item.skuId).filter((id): id is string => Boolean(id)),
    ),
  // null when not read; how each license reached the account.
  licenseAssignmentStates: z
    .array(
      z.object({
        skuId: z.string().max(100).nullish(),
        assignedByGroup: z.string().max(200).nullish(),
      }),
    )
    .max(1_000)
    .nullish()
    .catch(null)
    .transform((value) => value ?? null),
  signInActivity: z
    .object({
      lastSignInDateTime: rawDate,
      lastNonInteractiveSignInDateTime: rawDate,
      lastSuccessfulSignInDateTime: rawDate,
    })
    .nullish()
    .catch({
      lastSignInDateTime: UNREADABLE,
      lastNonInteractiveSignInDateTime: UNREADABLE,
      lastSuccessfulSignInDateTime: UNREADABLE,
    })
    .transform((value) => value ?? null),
});
export type DirectoryUser = z.output<typeof userSchema>;

const baseFields = [
  "id",
  "displayName",
  "userPrincipalName",
  "mail",
  "userType",
  "accountEnabled",
  "createdDateTime",
  "onPremisesSyncEnabled",
  "department",
  "jobTitle",
  "externalUserState",
  "externalUserStateChangeDateTime",
  "assignedLicenses",
];

export interface UserRead {
  users: DirectoryUser[];
  skipped: number;
  complete: boolean;
  // Whether last sign-in dates were read, and if not, why.
  signInActivity: "read" | "not_granted" | "refused";
  groupAssignment: boolean;
}

async function readUserPages(context: AuditContext, fields: string[], top: number) {
  context.signal.throwIfAborted();
  const result = await context.graph.getCollection(
    `users?$select=${fields.join(",")}&$top=${top}`,
    directoryContext,
    { schema: z.unknown(), maxPages: 100, maxRecords: MAX_ACCOUNTS },
  );
  const users: DirectoryUser[] = [];
  let skipped = 0;
  for (const raw of result.items) {
    const parsed = userSchema.safeParse(raw);
    if (parsed.success) users.push(parsed.data);
    else skipped += 1;
  }
  return { users, skipped, complete: result.complete };
}

const status = (error: unknown) => (error instanceof GraphRequestError ? error.status : 0);

// Every account in the directory. "required" sign-in activity fails the run when Microsoft
// refuses it (no Entra ID P1/P2, no reports role); "optional" reads without it instead.
export async function readUsers(
  context: AuditContext,
  options: { signInActivity: "required" | "optional"; groupAssignment?: boolean },
): Promise<UserRead> {
  let signIn: UserRead["signInActivity"] = context.granted.has("AuditLog.Read.All")
    ? "read"
    : "not_granted";
  if (signIn === "not_granted" && options.signInActivity === "required")
    throw new GraphRequestError(403);
  let groups = Boolean(options.groupAssignment);
  for (;;) {
    const fields = [
      ...baseFields,
      ...(signIn === "read" ? ["signInActivity"] : []),
      ...(groups ? ["licenseAssignmentStates"] : []),
    ];
    try {
      const read = await readUserPages(context, fields, signIn === "read" ? 500 : 999);
      return { ...read, signInActivity: signIn, groupAssignment: groups };
    } catch (error) {
      if (status(error) === 403 && signIn === "read" && options.signInActivity === "optional") {
        signIn = "refused";
        continue;
      }
      // A tenant that rejects the extra field still gets the rest of the audit.
      if (status(error) === 400 && groups) {
        groups = false;
        continue;
      }
      throw error;
    }
  }
}

const skuSchema = z.object({
  skuId: z.string().min(1).max(100),
  skuPartNumber: text,
  consumedUnits: count,
  prepaidUnits: z
    .object({ enabled: count, suspended: count, warning: count, lockedOut: count })
    .nullish()
    .catch(null)
    .transform((value) => value ?? { enabled: 0, suspended: 0, warning: 0, lockedOut: 0 }),
  capabilityStatus: text,
  appliesTo: text,
});
export type SubscribedSku = z.output<typeof skuSchema>;

export async function readSkus(
  context: AuditContext,
): Promise<{ skus: SubscribedSku[]; skipped: number; complete: boolean }> {
  context.signal.throwIfAborted();
  const result = await context.graph.getCollection("subscribedSkus", skuContext, {
    schema: z.unknown(),
    maxPages: 10,
    maxRecords: 1_000,
  });
  const skus: SubscribedSku[] = [];
  let skipped = 0;
  for (const raw of result.items) {
    const parsed = skuSchema.safeParse(raw);
    if (parsed.success) skus.push(parsed.data);
    else skipped += 1;
  }
  return { skus, skipped, complete: result.complete };
}

// License names by SKU ID, when the session may read subscriptions. null otherwise: the audit
// still runs and shows Microsoft's product IDs.
export async function readLicenseNames(context: AuditContext): Promise<Map<string, string> | null> {
  if (!context.granted.has("LicenseAssignment.Read.All")) return null;
  try {
    const { skus } = await readSkus(context);
    return new Map(
      skus.map((sku) => [sku.skuId, licenseName(sku.skuPartNumber ?? sku.skuId)] as const),
    );
  } catch (error) {
    if (context.signal.aborted) throw error;
    return null;
  }
}

export function licensesOf(user: DirectoryUser, names: Map<string, string> | null): string[] {
  return user.assignedLicenses
    .map((skuId) => names?.get(skuId) ?? skuId)
    .sort((a, b) => a.localeCompare(b));
}

export const licenseNamesNote =
  "License names could not be read (LicenseAssignment.Read.All), so licenses show Microsoft's product IDs.";

export type ActivityBasis = "success" | "attempt" | "none" | "unreadable";

export interface Activity {
  basis: ActivityBasis;
  // What the days are counted from; null when a date could not be read.
  from: number | null;
  days: number | null;
  // Sign-in attempts after Microsoft began recording successes, and no success recorded.
  noSuccess: boolean;
}

// A date Microsoft could have recorded: parseable, not before its records start, and not in the
// future (a day of clock difference allowed). Anything else, such as 0001-01-01, is unreadable:
// counting from it would make an active account look years inactive.
type Checked = { readable: true; at: number | null } | { readable: false };
function checkedDate(value: string | null, floor: number, now: Date): Checked {
  if (value === null) return { readable: true, at: null };
  if (!z.string().datetime({ offset: true }).safeParse(value).success) return { readable: false };
  const at = Date.parse(value);
  if (Number.isNaN(at) || at < floor || at > now.getTime() + DAY_MS) return { readable: false };
  return { readable: true, at };
}
// Successes are recorded from 1 December 2023; a month earlier is still accepted as real.
const SUCCESS_FLOOR = SUCCESS_RECORDED_SINCE - 30 * DAY_MS;
const CREATION_FLOOR = Date.parse("2000-01-01T00:00:00Z");

// How long an account has gone without signing in. Counted from the last successful sign-in
// (interactive or background). Without one, from the last attempt: that can only make the
// account look more recent than it is, never older, so nothing is flagged on a guess. With no
// sign-in on record at all, from creation or the start of Microsoft's records, whichever is
// later.
//
// Any date involved that cannot be read makes the whole account "unreadable": it gets no day
// count and no range, and the run is marked partial, so nothing is flagged on a broken value.
export function activityOf(user: DirectoryUser, now: Date): Activity {
  const activity = user.signInActivity;
  const unreadable: Activity = { basis: "unreadable", from: null, days: null, noSuccess: false };
  const success = checkedDate(activity?.lastSuccessfulSignInDateTime ?? null, SUCCESS_FLOOR, now);
  const interactive = checkedDate(activity?.lastSignInDateTime ?? null, SIGNIN_RECORDS_SINCE, now);
  const background = checkedDate(
    activity?.lastNonInteractiveSignInDateTime ?? null,
    SIGNIN_RECORDS_SINCE,
    now,
  );
  if (!success.readable || !interactive.readable || !background.readable) return unreadable;
  const since = (from: number) => ({
    from,
    days: Math.max(0, Math.floor((now.getTime() - from) / DAY_MS)),
  });
  if (success.at !== null) return { basis: "success", noSuccess: false, ...since(success.at) };
  const attempts = [interactive.at, background.at].filter((at): at is number => at !== null);
  if (attempts.length > 0) {
    const attempt = Math.max(...attempts);
    return { basis: "attempt", noSuccess: attempt >= SUCCESS_RECORDED_SINCE, ...since(attempt) };
  }
  const created = checkedDate(user.createdDateTime, CREATION_FLOOR, now);
  if (!created.readable) return unreadable;
  return {
    basis: "none",
    noSuccess: false,
    ...since(Math.max(created.at ?? SIGNIN_RECORDS_SINCE, SIGNIN_RECORDS_SINCE)),
  };
}

// A date for display: the value when it is readable, otherwise empty.
export function displayDate(value: string | null): string | null {
  if (value === null || !z.string().datetime({ offset: true }).safeParse(value).success)
    return null;
  return value;
}

export function basisLabel(activity: Activity): string {
  if (activity.basis === "success") return "Successful sign-in";
  if (activity.basis === "attempt") return "Sign-in attempt";
  if (activity.basis === "unreadable") return "Unreadable dates";
  return (activity.from ?? 0) > SIGNIN_RECORDS_SINCE
    ? "Account created"
    : "Start of records (April 2020)";
}

export function inactivityTags(activity: Activity): string[] {
  const days = activity.days;
  if (activity.basis === "unreadable" || days === null) return ["unreadable"];
  return [
    ...INACTIVITY_THRESHOLDS.filter((threshold) => days >= threshold).map(
      (threshold) => `d${threshold}`,
    ),
    ...(activity.basis === "none" ? ["never"] : []),
    ...(activity.noSuccess ? ["no-success"] : []),
  ];
}

export const iso = (value: number | null) =>
  value === null ? null : new Date(value).toISOString();

export function personCell(user: DirectoryUser) {
  return { name: user.displayName, principal: user.userPrincipalName };
}

export function statusOf(user: DirectoryUser): "Enabled" | "Disabled" | "Unknown" {
  return user.accountEnabled === null ? "Unknown" : user.accountEnabled ? "Enabled" : "Disabled";
}

export function typeOf(userType: string | null): string {
  if (!userType) return "Unknown";
  return `${userType.charAt(0).toUpperCase()}${userType.slice(1).toLowerCase()}`;
}

// A provider value as a row tag. Tags are short lowercase slugs; anything else is dropped
// rather than failing the run. The first replace leaves no two dashes in a row, so each end
// has at most one to trim, and no pattern has a run to backtrack over.
export function tagOf(value: string | null): string[] {
  const slug = (value ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 40)
    .replace(/-$/, "");
  return slug ? [slug] : [];
}

export function isGuest(user: DirectoryUser): boolean {
  return user.userType?.toLowerCase() === "guest";
}

// A guest's own organisation: from its email, or from the principal Microsoft builds for it
// ("alex_contoso.com#EXT#@tenant.onmicrosoft.com").
export function guestDomain(user: DirectoryUser): string | null {
  const mail = user.mail?.split("@")[1];
  if (mail) return mail.toLowerCase();
  const external = user.userPrincipalName?.split("#EXT#")[0];
  if (!external || external === user.userPrincipalName) return null;
  const at = external.lastIndexOf("_");
  return at > 0 ? external.slice(at + 1).toLowerCase() || null : null;
}

export function byName(
  a: { cells: Record<string, unknown> },
  b: { cells: Record<string, unknown> },
) {
  const name = (row: { cells: Record<string, unknown> }) => {
    const person = row.cells.person as
      | { name?: string | null; principal?: string | null }
      | undefined;
    return (person?.name ?? person?.principal ?? "").toLowerCase();
  };
  return name(a).localeCompare(name(b));
}

// The notes every account audit adds about how much of the directory it read.
export function coverageNotes(read: { skipped: number; complete: boolean }): string[] {
  return [
    ...(read.complete
      ? []
      : [
          `Stopped at ${MAX_ACCOUNTS.toLocaleString("en-US")} accounts; the rest of the directory was not read.`,
        ]),
    ...(read.skipped > 0
      ? [
          `${read.skipped} ${read.skipped === 1 ? "account" : "accounts"} could not be read and are left out.`,
        ]
      : []),
  ];
}

// Accounts whose sign-in dates could not be read are listed but never counted as inactive.
export function unreadableNote(count: number): string[] {
  return count > 0
    ? [
        `${count} ${count === 1 ? "account has" : "accounts have"} sign-in dates Prism365 could not read (malformed, or outside what Microsoft records). They are listed with "Unreadable dates", have no day count, and are left out of every range.`,
      ]
    : [];
}

export function signInNote(read: UserRead): string[] {
  if (read.signInActivity === "not_granted")
    return [
      "Last sign-in dates were not read: the session has not granted AuditLog.Read.All. Only disabled accounts and invitation states can be judged.",
    ];
  if (read.signInActivity === "refused")
    return [
      "Microsoft refused last sign-in dates. They need Microsoft Entra ID P1 or P2 and a role such as Reports Reader, Security Reader or Global Reader.",
    ];
  return [];
}

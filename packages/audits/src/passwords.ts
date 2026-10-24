import type { AuditRow } from "@prism365/contracts";
import { GraphRequestError } from "@prism365/graph";
import { z } from "zod";
import type { Audit, AuditContext } from "./definition.ts";
import {
  byName,
  coverageNotes,
  DAY_MS,
  type DirectoryUser,
  displayDate,
  isGuest,
  personCell,
  readUsers,
  statusOf,
  tagOf,
} from "./directory.ts";
import { firstAccepted } from "./graph-pages.ts";
import {
  domainColumn,
  domainInventory,
  excludedNote,
  inAuditScope,
  principalDomain,
  scopeNote,
} from "./scope.ts";

// How old each member account's password is, whether it can expire, and when. From the user's
// lastPasswordChangeDateTime and passwordPolicies (returned only when selected) and, when the
// domain read pack is on, each domain's password policy. Microsoft Graph v1.0 documentation
// verified 2026-10-24:
// https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/domain-list?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/resources/domain?view=graph-rest-1.0

const domainContext = { feature: "directory.domains.read", permission: "Domain.Read.All" } as const;

// What Microsoft applies when a domain leaves a setting unset.
const DEFAULT_VALIDITY_DAYS = 90;
const DEFAULT_NOTIFICATION_DAYS = 14;
// How Microsoft stores "passwords never expire" on a domain.
const NEVER_EXPIRES_DAYS = 2_147_483_647;
const MAX_DOMAINS = 1_000;
// The oldest password date taken at its word (Microsoft 365 did not exist before this), and the
// latest date a Date can show as a four-digit year.
const FIRST_PASSWORD_DATE = Date.parse("2000-01-01T00:00:00Z");
const LAST_DISPLAY_DATE = Date.parse("9999-12-31T23:59:59Z");
// A password date a day ahead of the clock is a clock difference, not a change in the future.
const CLOCK_MARGIN = DAY_MS;

const AGE_THRESHOLDS = [90, 180, 365] as const;
const thresholdLabel = (days: number) => (days === 365 ? "1 year" : `${days} days`);

const STATUS = {
  weak: "Strong password not required",
  expired: "Expired",
  expiring: "Expires soon",
  unreadable: "Unreadable date",
  never: "Never changed",
  ok: "OK",
} as const;
const olderThan = (days: number) => `Older than ${thresholdLabel(days)}`;

// A domain's policy: the days a password lasts (null for never) and how many days before the end
// Microsoft starts telling people to change it. A setting left unset takes Microsoft's default.
interface DomainPolicy {
  days: number | null;
  notificationDays: number;
}

const setting = (schema: z.ZodNumber) => schema.nullish().transform((value) => value ?? null);
const domainSchema = z.object({
  id: z.string().min(1).max(253),
  passwordValidityPeriodInDays: setting(z.number().int().positive()),
  passwordNotificationWindowInDays: setting(z.number().int().nonnegative()),
});

interface DomainRead {
  policies: Map<string, DomainPolicy>;
  skipped: number;
  complete: boolean;
}

async function readDomains(context: AuditContext): Promise<DomainRead> {
  // Microsoft documents a problem with $top on this list, so the request never sends it.
  const read = (path: string) => {
    context.signal.throwIfAborted();
    return context.graph.getCollection(path, domainContext, {
      schema: z.unknown(),
      maxPages: 20,
      maxRecords: MAX_DOMAINS,
    });
  };
  // $select is not confirmed for this list, so a tenant that rejects it still gets the audit.
  const result = await firstAccepted(
    ["domains?$select=id,passwordValidityPeriodInDays,passwordNotificationWindowInDays", "domains"],
    read,
  );
  const policies = new Map<string, DomainPolicy>();
  let skipped = 0;
  for (const raw of result.items) {
    const parsed = domainSchema.safeParse(raw);
    if (!parsed.success) {
      skipped += 1;
      continue;
    }
    const {
      id,
      passwordValidityPeriodInDays: validity,
      passwordNotificationWindowInDays,
    } = parsed.data;
    policies.set(id.toLowerCase(), {
      days:
        validity === null
          ? DEFAULT_VALIDITY_DAYS
          : validity >= NEVER_EXPIRES_DAYS
            ? null
            : validity,
      notificationDays: passwordNotificationWindowInDays ?? DEFAULT_NOTIFICATION_DAYS,
    });
  }
  return { policies, skipped, complete: result.complete };
}

type Changed =
  | { state: "none" }
  | { state: "unreadable" }
  | { state: "date"; at: number; days: number };

// When the password last changed. A missing date means it never did; a date that is malformed,
// older than Microsoft 365 or ahead of the clock is unreadable, never "never".
function changedOf(value: string | null, now: Date): Changed {
  if (value === null) return { state: "none" };
  const shown = displayDate(value);
  const at = shown === null ? Number.NaN : Date.parse(shown);
  if (Number.isNaN(at) || at < FIRST_PASSWORD_DATE || at > now.getTime() + CLOCK_MARGIN)
    return { state: "unreadable" };
  return { state: "date", at, days: Math.max(0, Math.floor((now.getTime() - at) / DAY_MS)) };
}

function policyFlags(user: DirectoryUser) {
  const flags = new Set(
    (user.passwordPolicies ?? "").split(",").map((flag) => flag.trim().toLowerCase()),
  );
  return {
    neverExpires: flags.has("disablepasswordexpiration"),
    weak: flags.has("disablestrongpassword"),
  };
}

const daysText = (days: number) => (days === 1 ? "1 day" : `${days} days`);

function assess(user: DirectoryUser, now: Date, domains: Map<string, DomainPolicy> | null) {
  const changed = changedOf(user.lastPasswordChangeDateTime, now);
  const flags = policyFlags(user);
  const synced = user.onPremisesSyncEnabled === true;
  const policy = domains?.get(principalDomain(user.userPrincipalName) ?? "") ?? null;

  let policyLabel = "Unknown";
  if (synced) policyLabel = "Managed on-premises";
  else if (policy) policyLabel = policy.days === null ? "Never" : daysText(policy.days);

  // No date for an account the cloud policy does not govern, or one the domain never expires.
  let expiresAt: number | null = null;
  if (
    changed.state === "date" &&
    !synced &&
    !flags.neverExpires &&
    policy !== null &&
    policy.days !== null
  ) {
    const end = changed.at + policy.days * DAY_MS;
    if (end <= LAST_DISPLAY_DATE) expiresAt = end;
  }

  let status: string = STATUS.ok;
  const tags: string[] = [];
  if (flags.weak) tags.push("weak");
  if (flags.neverExpires) tags.push("never-expires");
  if (changed.state === "none") tags.push("never-changed");
  if (changed.state === "unreadable") tags.push("unreadable");
  if (changed.state === "date")
    for (const threshold of AGE_THRESHOLDS)
      if (changed.days >= threshold) tags.push(`d${threshold}`);
  let expired = false;
  let expiring = false;
  if (expiresAt !== null) {
    expired = expiresAt <= now.getTime();
    expiring = !expired && expiresAt - now.getTime() <= (policy?.notificationDays ?? 0) * DAY_MS;
    if (expired) tags.push("expired");
    if (expiring) tags.push("expiring");
  }

  // The first match wins. An old password is named by the largest age it passes, so the choice of
  // age in the views never changes what a row says.
  const passed = AGE_THRESHOLDS.filter(
    (threshold) => changed.state === "date" && changed.days >= threshold,
  );
  const oldest = passed[passed.length - 1];
  if (flags.weak) status = STATUS.weak;
  else if (expired) status = STATUS.expired;
  else if (expiring) status = STATUS.expiring;
  else if (changed.state === "unreadable") status = STATUS.unreadable;
  else if (changed.state === "none") status = STATUS.never;
  else if (oldest !== undefined) status = olderThan(oldest);

  return {
    tags,
    status,
    synced,
    policyLabel,
    flags,
    lastChanged: changed.state === "date" ? new Date(changed.at).toISOString() : null,
    days: changed.state === "date" ? changed.days : null,
    expires: expiresAt === null ? null : new Date(expiresAt).toISOString(),
  };
}

const agedViews = [...AGE_THRESHOLDS].reverse().map((days) => ({
  id: `d${days}`,
  label: `Older than ${thresholdLabel(days)}`,
  tags: [`d${days}`],
}));

// Member accounts: how old each password is, which can expire, and which are set up weakly.
export const passwordsAudit: Audit = {
  definition: {
    id: "accounts.passwords",
    // Accounts an action tool could act on (docs/handoff.md).
    handoff: true,
    version: 1,
    title: "Passwords",
    description:
      "How old each member account's password is, which accounts are set never to expire or to skip the strong password rule, and when passwords expire.",
    category: "Accounts",
    subject: { one: "account", many: "accounts" },
    columns: [
      { key: "person", label: "Account", kind: "person" },
      { key: "enabled", label: "Enabled", kind: "enum", facet: true },
      { key: "source", label: "Source", kind: "enum", facet: true },
      {
        key: "lastChanged",
        label: "Last changed",
        kind: "date",
        empty: "Never changed",
        help: "When the password was last changed or created, whichever is later. A reset by an administrator counts.",
      },
      {
        key: "days",
        label: "Password age",
        kind: "days",
        empty: "Unknown",
        help: "Days since the password last changed.",
      },
      {
        key: "policy",
        label: "Expiry policy",
        kind: "enum",
        facet: true,
        help: "The password lifetime of the account's domain. Unknown when domains were not read; accounts synced from on-premises follow the on-premises policy.",
      },
      {
        key: "expires",
        label: "Expires on",
        kind: "date",
        empty: "No date",
        help: "The last change plus the domain's password lifetime. Worked out by Prism365, not returned by Microsoft. Empty when the password cannot expire or the policy is unknown.",
      },
      {
        key: "neverExpires",
        label: "Never expires for this account",
        kind: "boolean",
        help: "The account has the DisablePasswordExpiration setting, whatever its domain says.",
      },
      {
        key: "weak",
        label: "Strong password not required",
        kind: "boolean",
        help: "The account has the DisableStrongPassword setting, so it may keep a password Microsoft would otherwise refuse.",
      },
      { key: "status", label: "Status", kind: "enum", facet: true },
      domainColumn,
      { key: "created", label: "Created", kind: "date", hidden: true, empty: "Unknown" },
    ],
    views: [
      ...agedViews,
      { id: "all", label: "All member accounts", tags: [] },
      { id: "weak", label: "Strong password not required", tags: ["weak"] },
      { id: "never-expires", label: "Never expires", tags: ["never-expires"] },
      { id: "expired", label: "Expired", tags: ["expired"] },
      { id: "expiring", label: "Expires soon", tags: ["expiring"] },
      { id: "never-changed", label: "Enabled, never changed", tags: ["never-changed", "enabled"] },
    ],
    defaultView: "d365",
    defaultSort: { column: "days", direction: "desc" },
    requirements: [
      "Expiry dates need the domains read pack and a role such as Global Reader or Directory Readers. Without it the audit shows password age and the account settings only.",
    ],
  },
  packs: ["directory.users.read"],
  optionalPacks: ["directory.domains.read"],
  scoped: true,
  forbidden: "Microsoft refused to list accounts. Listing them needs User.Read.All.",
  async run(context) {
    const read = await readUsers(context, { signInActivity: "none", passwords: true });
    const guests = read.users.filter(isGuest).length;
    const members = read.users.filter((user) => !isGuest(user));
    const covered = members.filter((user) => inAuditScope(context.scope, user.userPrincipalName));

    // Domains are optional: a refusal leaves the expiry columns unknown and the run partial.
    let domains: DomainRead | null = null;
    let domainsRefused = false;
    if (context.granted.has("Domain.Read.All")) {
      try {
        domains = await readDomains(context);
      } catch (error) {
        if (!(error instanceof GraphRequestError) || error.status === 401) throw error;
        domainsRefused = true;
      }
    }
    const policies = domains?.policies ?? null;

    const rows: AuditRow[] = covered.map((user) => {
      const found = assess(user, context.now, policies);
      const status = statusOf(user);
      return {
        key: user.id,
        personId: user.id,
        tags: [...found.tags, ...tagOf(status), ...(found.synced ? ["synced"] : [])],
        cells: {
          person: personCell(user),
          enabled: status,
          source: found.synced ? "On-premises" : "Cloud",
          lastChanged: found.lastChanged,
          days: found.days,
          policy: found.policyLabel,
          expires: found.expires,
          neverExpires: found.flags.neverExpires,
          weak: found.flags.weak,
          status: found.status,
          domain: principalDomain(user.userPrincipalName),
          created: displayDate(user.createdDateTime),
        },
      };
    });
    rows.sort(byName);
    const count = (...tags: string[]) =>
      rows.filter((row) => tags.every((tag) => row.tags.includes(tag))).length;
    const unreadable = count("unreadable");
    const limited = scopeNote(context.scope);
    const left = excludedNote(members.length - covered.length);
    const skippedDomains = domains?.skipped ?? 0;
    return {
      domains: domainInventory(members.map((user) => user.userPrincipalName)),
      rows,
      complete:
        read.complete &&
        read.skipped === 0 &&
        unreadable === 0 &&
        !domainsRefused &&
        (domains === null || (domains.complete && skippedDomains === 0)),
      summary: [
        { label: "Member accounts", value: rows.length, tone: "neutral", view: "all" },
        {
          label: "Strong password not required",
          value: count("weak"),
          tone: count("weak") > 0 ? "critical" : "neutral",
          view: "weak",
        },
        {
          label: "Never expires",
          value: count("never-expires"),
          tone: count("never-expires") > 0 ? "warning" : "neutral",
          view: "never-expires",
        },
        {
          label: `Older than ${thresholdLabel(365)}`,
          value: count("d365"),
          tone: count("d365") > 0 ? "warning" : "neutral",
          view: "d365",
        },
        {
          label: "Expired",
          value: count("expired"),
          tone: count("expired") > 0 ? "warning" : "neutral",
          view: "expired",
        },
        ...(policies
          ? [
              {
                label: "Expires soon",
                value: count("expiring"),
                tone: "neutral" as const,
                view: "expiring",
              },
            ]
          : []),
        {
          label: "Enabled, never changed",
          value: count("never-changed", "enabled"),
          tone: count("never-changed", "enabled") > 0 ? "warning" : "neutral",
          view: "never-changed",
        },
      ],
      notes: [
        ...(limited ? [limited] : []),
        ...(left ? [left] : []),
        ...(guests > 0
          ? [
              `${guests} ${guests === 1 ? "guest account is" : "guest accounts are"} not listed. A guest signs in with a password from its own organisation.`,
            ]
          : []),
        "Password age counts from the last change or creation, whichever is later. A reset by an administrator counts as a change.",
        ...(policies
          ? [
              "Expiry dates are the last change plus the password lifetime of the account's domain, worked out by Prism365. A domain that leaves the lifetime unset gets Microsoft's default of 90 days, and a notification window left unset gets 14 days.",
            ]
          : domainsRefused
            ? [
                "Microsoft refused to list domains, so expiry dates are unknown. Listing them needs Domain.Read.All and a role such as Global Reader or Directory Readers.",
              ]
            : [
                "Expiry dates were not worked out because domains were not read. Turn on the domains read pack (Domain.Read.All) and run again.",
              ]),
        ...(domains && (skippedDomains > 0 || !domains.complete)
          ? ["Some domains could not be read. Their accounts show an unknown expiry policy."]
          : []),
        "Accounts synced from on-premises follow the on-premises password policy, which Prism365 cannot see, so they have no expiry date here. Microsoft's EnforceCloudPasswordPolicyForPasswordSyncedUsers setting can apply the cloud policy to them as well; Prism365 does not read that setting.",
        "An account that signs in without a password, with a passkey or Windows Hello for example, can have an old password date that means little.",
        ...(unreadable > 0
          ? [
              `${unreadable} ${unreadable === 1 ? "account has" : "accounts have"} a password date Prism365 could not read. ${unreadable === 1 ? "It is" : "They are"} listed with "${STATUS.unreadable}" and left out of every age range.`,
            ]
          : []),
        ...coverageNotes(read),
      ],
    };
  },
};

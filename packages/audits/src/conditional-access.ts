import { type AuditRow, cleanText } from "@prism365/contracts";
import { GraphRequestError } from "@prism365/graph";
import { z } from "zod";
import type { Audit, AuditContext } from "./definition.ts";
import { createPacer, type Pacer, type Pause, readPaced } from "./paced-read.ts";

// Conditional Access policies: who each one applies to, what it leaves out, and whether the tenant
// blocks legacy authentication and asks administrators for MFA. From Microsoft Graph v1.0, checked
// 2026-10-22:
// https://learn.microsoft.com/en-us/graph/api/conditionalaccessroot-list-policies?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/conditionalaccessroot-list-namedlocations?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/resources/conditionalaccesspolicy?view=graph-rest-1.0

const policyContext = {
  feature: "identity.conditional-access.read",
  permission: "Policy.Read.All",
} as const;
const usersContext = { feature: "directory.users.read", permission: "User.Read.All" } as const;
const skuContext = {
  feature: "tenant.licenses.read",
  permission: "LicenseAssignment.Read.All",
} as const;

// Microsoft does not document a page size for either endpoint, so neither request sets $top and
// both follow the continuation link. A tenant has tens of policies, not thousands.
const MAX_POLICIES = 1_000;
const MAX_LOCATIONS = 1_000;
const MAX_PAGES = 20;
// Excluded users looked up by name. Each is one request to the user endpoint, which has its own
// limit, so this bounds the run rather than the throttle.
const MAX_NAMED_USERS = 50;
const NAME_BATCH = 8;
// Names listed in one cell before "and N more".
const SHOWN_NAMES = 5;

// Role template IDs, which is how policies name roles. Only the roles a policy for administrators
// usually lists; the IDs are from Microsoft Entra's built-in roles reference, checked 2026-10-22:
// https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/permissions-reference
const GLOBAL_ADMINISTRATOR = "62e90394-69f5-4237-9190-012177145e10";
const roleNames = new Map<string, string>([
  [GLOBAL_ADMINISTRATOR, "Global Administrator"],
  ["e8611ab8-c189-46e8-94e1-60213ab1f814", "Privileged Role Administrator"],
  ["194ae4cb-b126-40b2-bd5b-6091b380977d", "Security Administrator"],
  ["29232cdf-9323-42fd-ade2-1d097af3e4de", "Exchange Administrator"],
  ["f28a1f50-f6e7-4571-818b-6a12f2af6b6c", "SharePoint Administrator"],
  ["fe930be7-5e62-47db-91af-98c3a49a38b1", "User Administrator"],
  ["b1be1c3e-b65d-4f19-8427-f6fa0d97feb9", "Conditional Access Administrator"],
  ["9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3", "Application Administrator"],
  ["158c047a-c907-4556-b7ef-446551a6b5f7", "Cloud Application Administrator"],
  ["729827e3-9c14-49f7-bb1b-9608f156bbb8", "Helpdesk Administrator"],
  ["c4e39bd9-1100-46d3-8c65-fb160da0071f", "Authentication Administrator"],
  ["7be44c8a-adaf-4e2a-84d6-ab2649e08a13", "Privileged Authentication Administrator"],
  ["966707d0-3269-4727-9be2-8c3a10f19b9d", "Password Administrator"],
  ["b0f54661-2d74-4c50-afa3-1ec803f12efe", "Billing Administrator"],
  ["3a2c62db-5318-420d-8d74-23affee5d9d5", "Intune Administrator"],
  ["69091246-20e8-4a56-aa4d-066075b2a7a8", "Teams Administrator"],
]);

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// Provider values. A value of the wrong type reads as missing instead of failing the policy; a
// policy is unreadable only when its shape is wrong at a level the checks depend on.
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
const whole = z
  .number()
  .int()
  .nullish()
  .catch(null)
  .transform((value) => value ?? null);
// A list of strings: other entries are dropped and a very long list is cut, so a count is exact
// up to the cut.
const strings = z
  .array(z.unknown())
  .nullish()
  .catch(null)
  .transform((value) =>
    (value ?? [])
      .filter((item): item is string => typeof item === "string")
      .slice(0, 5_000)
      .map((item) => cleanText(item.slice(0, 100))),
  );

// An object Microsoft may leave out entirely, read as the same object with nothing set.
function orEmpty<S extends z.ZodObject>(schema: S) {
  return schema.nullish().transform((value) => value ?? schema.parse({}));
}

// The wire format of guest and external user types is a comma-separated string; the list endpoint
// may return it as one or as an array (the documentation does not say), so both are read.
const guestSelection = z
  .object({ guestOrExternalUserTypes: z.union([z.string(), z.array(z.unknown())]).nullish() })
  .nullish()
  .catch(null)
  .transform((value) => {
    if (!value) return null;
    const raw = value.guestOrExternalUserTypes;
    const parts = Array.isArray(raw)
      ? raw.filter((item): item is string => typeof item === "string")
      : (raw ?? "").split(",");
    return {
      types: parts
        .map((part) => cleanText(part.trim().slice(0, 60)))
        .filter(Boolean)
        .slice(0, 20),
    };
  });

const usersSchema = z.object({
  includeUsers: strings,
  excludeUsers: strings,
  includeGroups: strings,
  excludeGroups: strings,
  includeRoles: strings,
  excludeRoles: strings,
  includeGuestsOrExternalUsers: guestSelection,
  excludeGuestsOrExternalUsers: guestSelection,
});
const applicationsSchema = z.object({
  includeApplications: strings,
  excludeApplications: strings,
  includeUserActions: strings,
  includeAuthenticationContextClassReferences: strings,
  applicationFilter: z
    .object({})
    .nullish()
    .catch(null)
    .transform((value) => value !== null && value !== undefined),
});
const platformsSchema = z.object({ includePlatforms: strings, excludePlatforms: strings });
const locationsSchema = z.object({ includeLocations: strings, excludeLocations: strings });
const workloadsSchema = z.object({ includeServicePrincipals: strings });

const conditionsSchema = z.object({
  users: orEmpty(usersSchema),
  applications: orEmpty(applicationsSchema),
  clientApplications: orEmpty(workloadsSchema),
  clientAppTypes: strings,
  platforms: orEmpty(platformsSchema),
  locations: orEmpty(locationsSchema),
  signInRiskLevels: strings,
  userRiskLevels: strings,
  servicePrincipalRiskLevels: strings,
});

const grantSchema = z.object({
  operator: text,
  builtInControls: strings,
  customAuthenticationFactors: strings,
  termsOfUse: strings,
  // Documented as a relationship but returned inline, null or populated.
  authenticationStrength: z
    .object({ displayName: text })
    .nullish()
    .catch(null)
    .transform((value) => value ?? null),
});

const sessionSchema = z.object({
  signInFrequency: z
    .object({
      isEnabled: flag,
      type: text,
      value: whole,
      frequencyInterval: text,
    })
    .nullish()
    .catch(null)
    .transform((value) => value ?? null),
  persistentBrowser: z
    .object({ isEnabled: flag, mode: text })
    .nullish()
    .catch(null)
    .transform((value) => value ?? null),
  applicationEnforcedRestrictions: z
    .object({ isEnabled: flag })
    .nullish()
    .catch(null)
    .transform((value) => value ?? null),
  cloudAppSecurity: z
    .object({ isEnabled: flag })
    .nullish()
    .catch(null)
    .transform((value) => value ?? null),
  disableResilienceDefaults: flag,
});

const policySchema = z.object({
  id: z
    .string()
    .min(1)
    .max(200)
    .regex(/^[A-Za-z0-9._@:-]+$/),
  displayName: text,
  state: z.string().max(100),
  createdDateTime: time,
  modifiedDateTime: time,
  conditions: orEmpty(conditionsSchema),
  grantControls: orEmpty(grantSchema),
  sessionControls: orEmpty(sessionSchema),
});
type Policy = z.output<typeof policySchema>;

const locationSchema = z.object({
  id: z.string().min(1).max(200),
  displayName: text,
});

const userNameSchema = z.object({ displayName: text, userPrincipalName: text });

const planSchema = z.object({ servicePlanName: text, provisioningStatus: text });
const entitlementSchema = z.object({
  capabilityStatus: text,
  servicePlans: z
    .array(z.unknown())
    .nullish()
    .catch(null)
    .transform((value) =>
      (value ?? []).flatMap((item) => {
        const parsed = planSchema.safeParse(item);
        return parsed.success ? [parsed.data] : [];
      }),
    ),
});

const lower = (value: string) => value.toLowerCase();
const has = (list: readonly string[], literal: string) =>
  list.some((item) => lower(item) === lower(literal));
const plural = (count: number, one: string, many: string) => `${count} ${count === 1 ? one : many}`;

function joinNames(names: readonly string[]): string {
  if (names.length <= 1) return names.join("");
  return `${names.slice(0, -1).join(", ")} and ${names[names.length - 1]}`;
}

// Names with a cut: "a, b, c and 4 more".
function namesWithMore(names: readonly string[]): string {
  const shown = names.slice(0, SHOWN_NAMES);
  const rest = names.length - shown.length;
  return rest > 0 ? `${shown.join(", ")} and ${rest} more` : shown.join(", ");
}

function humanize(value: string): string {
  const spaced = value.replace(/([a-z0-9])([A-Z])/g, "$1 $2").toLowerCase();
  return spaced.trim() || "unknown";
}

type PolicyState = "On" | "Report only" | "Off" | "Unknown";
function stateOf(policy: Policy): PolicyState {
  const state = lower(policy.state);
  if (state === "enabled") return "On";
  if (state === "enabledforreportingbutnotenforced") return "Report only";
  if (state === "disabled") return "Off";
  return "Unknown";
}

function userIds(list: readonly string[]): string[] {
  return [...new Set(list.filter((item) => GUID.test(item)).map(lower))];
}

function requiresMfa(policy: Policy): boolean {
  const grant = policy.grantControls;
  return has(grant.builtInControls, "mfa") || grant.authenticationStrength !== null;
}

// An enabled policy that blocks both Exchange ActiveSync and the other legacy clients for everyone.
// A block on "all client apps" does not count: it is not aimed at legacy clients, and in practice
// another condition (a location, a device, a risk level) narrows it. The check does not weigh
// apps, locations or exclusions, and the notes say so.
function blocksLegacyAuthentication(policy: Policy): boolean {
  if (stateOf(policy) !== "On") return false;
  const { users, clientAppTypes } = policy.conditions;
  const legacy = has(clientAppTypes, "exchangeActiveSync") && has(clientAppTypes, "other");
  return (
    has(users.includeUsers, "All") && legacy && has(policy.grantControls.builtInControls, "block")
  );
}

// An enabled policy that asks for MFA from all users or from Global Administrators, as long as it
// does not leave Global Administrators out.
function requiresMfaForAdministrators(policy: Policy): boolean {
  if (stateOf(policy) !== "On" || !requiresMfa(policy)) return false;
  const { users } = policy.conditions;
  const covered = has(users.includeUsers, "All") || has(users.includeRoles, GLOBAL_ADMINISTRATOR);
  return covered && !has(users.excludeRoles, GLOBAL_ADMINISTRATOR);
}

function rolesText(ids: readonly string[]): string {
  const known = ids.map((id) => roleNames.get(lower(id))).filter((name): name is string => !!name);
  const total = plural(ids.length, "role", "roles");
  if (known.length === 0) return total;
  const shown = known.slice(0, 3);
  const rest = ids.length - shown.length;
  return `${total} (${shown.join(", ")}${rest > 0 ? ` and ${rest} ${rest === 1 ? "other" : "others"}` : ""})`;
}

const guestTypeLabels: Record<string, string> = {
  internalguest: "internal guests",
  b2bcollaborationguest: "B2B collaboration guests",
  b2bcollaborationmember: "B2B collaboration members",
  b2bdirectconnectuser: "B2B direct connect users",
  otherexternaluser: "other external users",
  serviceprovider: "service providers",
};

function guestsText(selection: { types: string[] } | null, literal: boolean): string | null {
  if (!selection && !literal) return null;
  const types = (selection?.types ?? [])
    .filter((type) => lower(type) !== "none")
    .map((type) => guestTypeLabels[lower(type)] ?? humanize(type));
  return types.length > 0
    ? `Guests and external users (${joinNames(types)})`
    : "Guests and external users";
}

function appliesTo(policy: Policy): string {
  const { users, clientApplications } = policy.conditions;
  if (has(users.includeUsers, "All")) return "All users";
  const parts: string[] = [];
  const ids = userIds(users.includeUsers);
  if (ids.length > 0) parts.push(plural(ids.length, "user", "users"));
  if (users.includeGroups.length > 0)
    parts.push(plural(users.includeGroups.length, "group", "groups"));
  if (users.includeRoles.length > 0) parts.push(rolesText(users.includeRoles));
  const guests = guestsText(
    users.includeGuestsOrExternalUsers,
    has(users.includeUsers, "GuestsOrExternalUsers"),
  );
  if (guests) parts.push(guests);
  if (parts.length === 0 && clientApplications.includeServicePrincipals.length > 0)
    return "Workload identities";
  return parts.length > 0 ? parts.join(", ") : "No users selected";
}

function excludesOf(policy: Policy, names: ReadonlyMap<string, string | null> | null): string[] {
  const { users } = policy.conditions;
  const items: string[] = [];
  const ids = userIds(users.excludeUsers);
  if (ids.length > 0) {
    const total = plural(ids.length, "user", "users");
    // A name for each account Microsoft knows, then one count for those it no longer has.
    const known = names
      ? ids.flatMap((id) => {
          const name = names.get(id);
          return name ? [name] : [];
        })
      : [];
    const gone = names ? ids.filter((id) => names.get(id) === null).length : 0;
    const named = [
      ...(known.length > 0 ? [namesWithMore(known)] : []),
      ...(gone > 0
        ? [plural(gone, "deleted or unknown account", "deleted or unknown accounts")]
        : []),
    ];
    items.push(named.length > 0 ? `${total}: ${named.join(", ")}` : total);
  }
  if (users.excludeGroups.length > 0)
    items.push(plural(users.excludeGroups.length, "group", "groups"));
  if (users.excludeRoles.length > 0) items.push(rolesText(users.excludeRoles));
  const guests = guestsText(
    users.excludeGuestsOrExternalUsers,
    has(users.excludeUsers, "GuestsOrExternalUsers"),
  );
  if (guests) items.push(guests);
  return items;
}

function appsText(policy: Policy): string {
  const { applications } = policy.conditions;
  const parts: string[] = [];
  const specific = applications.includeApplications.filter(
    (item) => !["none", "all", "office365", "microsoftadminportals"].includes(lower(item)),
  );
  if (has(applications.includeApplications, "All")) parts.push("All apps");
  if (has(applications.includeApplications, "Office365")) parts.push("Office 365");
  if (has(applications.includeApplications, "MicrosoftAdminPortals"))
    parts.push("Microsoft admin portals");
  if (specific.length > 0) parts.push(plural(specific.length, "specific app", "specific apps"));
  if (applications.applicationFilter) parts.push("Apps matching a filter");
  for (const action of applications.includeUserActions) {
    if (lower(action) === "urn:user:registersecurityinfo") parts.push("Registering security info");
    else if (lower(action) === "urn:user:registerdevice") parts.push("Registering devices");
    else parts.push("A user action");
  }
  if (applications.includeAuthenticationContextClassReferences.length > 0)
    parts.push(
      plural(
        applications.includeAuthenticationContextClassReferences.length,
        "authentication context",
        "authentication contexts",
      ),
    );
  if (parts.length === 0) return "None selected";
  const left = applications.excludeApplications.filter((item) => lower(item) !== "none");
  return left.length > 0
    ? `${parts.join(", ")}, except ${plural(left.length, "app", "apps")}`
    : parts.join(", ");
}

const clientAppLabels: Record<string, string> = {
  browser: "Browser",
  mobileappsanddesktopclients: "Mobile apps and desktop clients",
  exchangeactivesync: "Exchange ActiveSync",
  eassupported: "Exchange ActiveSync on supported platforms",
  other: "Other clients",
};
const platformLabels: Record<string, string> = {
  android: "Android",
  ios: "iOS",
  windows: "Windows",
  windowsphone: "Windows Phone",
  macos: "macOS",
  linux: "Linux",
};

const label = (value: string, labels: Record<string, string>) =>
  labels[lower(value)] ?? humanize(value);
const sentenceCase = (value: string) => `${value.charAt(0).toUpperCase()}${value.slice(1)}`;

interface LocationNames {
  // Named locations by lower-case ID; null when the read failed.
  names: ReadonlyMap<string, string> | null;
}

function locationList(ids: readonly string[], known: LocationNames): string | null {
  const names: string[] = [];
  let unnamed = 0;
  for (const id of ids) {
    if (lower(id) === "all") names.push("any location");
    else if (lower(id) === "alltrusted") names.push("trusted locations");
    else if (lower(id) === "none") continue;
    else {
      const name = known.names?.get(lower(id));
      if (name) names.push(name);
      else unnamed += 1;
    }
  }
  if (unnamed > 0)
    names.push(`${plural(unnamed, "location", "locations")} that could not be named`);
  return names.length > 0 ? joinNames(names) : null;
}

function conditionsOf(policy: Policy, locations: LocationNames): string[] {
  const { conditions } = policy;
  const items: string[] = [];
  const apps = conditions.clientAppTypes.filter((type) => lower(type) !== "all");
  if (apps.length > 0)
    items.push(`Client apps: ${apps.map((type) => label(type, clientAppLabels)).join(", ")}`);
  const platforms = conditions.platforms;
  const includedPlatforms = platforms.includePlatforms.filter((item) => lower(item) !== "all");
  if (includedPlatforms.length > 0 || platforms.excludePlatforms.length > 0) {
    const included =
      includedPlatforms.length > 0
        ? includedPlatforms.map((item) => label(item, platformLabels)).join(", ")
        : "all platforms";
    const left = platforms.excludePlatforms.map((item) => label(item, platformLabels));
    items.push(`Platforms: ${included}${left.length > 0 ? `, except ${left.join(", ")}` : ""}`);
  }
  const included = locationList(conditions.locations.includeLocations, locations);
  const left = locationList(conditions.locations.excludeLocations, locations);
  if (included || left)
    items.push(`Locations: ${included ?? "any location"}${left ? `, except ${left}` : ""}`);
  for (const [title, levels] of [
    ["Sign-in risk", conditions.signInRiskLevels],
    ["User risk", conditions.userRiskLevels],
    ["Service principal risk", conditions.servicePrincipalRiskLevels],
  ] as const) {
    const shown = levels.filter((level) => lower(level) !== "unknownfuturevalue");
    if (shown.length > 0)
      items.push(`${title}: ${shown.map((level) => sentenceCase(lower(level))).join(", ")}`);
  }
  return items;
}

const controlLabels: Record<string, string> = {
  mfa: "MFA",
  compliantdevice: "a compliant device",
  domainjoineddevice: "a hybrid joined device",
  approvedapplication: "an approved client app",
  compliantapplication: "an app protection policy",
  passwordchange: "a password change",
  riskremediation: "risk remediation",
};

function grantText(policy: Policy): string {
  const grant = policy.grantControls;
  if (has(grant.builtInControls, "block")) return "Block access";
  const items = grant.builtInControls.map((control) => label(control, controlLabels));
  if (grant.authenticationStrength)
    items.push(
      grant.authenticationStrength.displayName
        ? `authentication strength ${grant.authenticationStrength.displayName}`
        : "an authentication strength",
    );
  if (grant.customAuthenticationFactors.length > 0) items.push("a custom control");
  if (grant.termsOfUse.length > 0) items.push("accepting terms of use");
  if (items.length === 0) return "No grant control";
  if (items.length === 1) return `Require ${items[0]}`;
  return lower(grant.operator ?? "") === "or"
    ? `Require one of: ${items.join(", ")}`
    : `Require ${joinNames(items)}`;
}

function sessionText(policy: Policy): string | null {
  const session = policy.sessionControls;
  const parts: string[] = [];
  const frequency = session.signInFrequency;
  if (frequency && frequency.isEnabled !== false) {
    if (lower(frequency.frequencyInterval ?? "") === "everytime") parts.push("Sign in every time");
    else if (frequency.value !== null && frequency.type) {
      const unit = lower(frequency.type) === "days" ? "day" : "hour";
      parts.push(
        frequency.value === 1
          ? `Sign in every ${unit}`
          : `Sign in every ${frequency.value} ${unit}s`,
      );
    }
  }
  const browser = session.persistentBrowser;
  if (browser && browser.isEnabled !== false && browser.mode)
    parts.push(`Persistent browser: ${lower(browser.mode)}`);
  const restrictions = session.applicationEnforcedRestrictions;
  if (restrictions && restrictions.isEnabled !== false) parts.push("App enforced restrictions");
  const appControl = session.cloudAppSecurity;
  if (appControl && appControl.isEnabled !== false) parts.push("Conditional Access App Control");
  if (session.disableResilienceDefaults === true) parts.push("Resilience defaults off");
  return parts.length > 0 ? parts.join("; ") : null;
}

type Entitlement = "yes" | "no" | "unknown";

// Whether a subscription that is in good standing carries Microsoft Entra ID P1 (the plan P2
// includes). "no" only when every subscription was read and none has it. Subscriptions list their
// service plans: https://learn.microsoft.com/en-us/graph/api/resources/serviceplaninfo?view=graph-rest-1.0
async function entraP1(context: AuditContext): Promise<Entitlement> {
  if (!context.granted.has("LicenseAssignment.Read.All")) return "unknown";
  try {
    context.signal.throwIfAborted();
    const result = await context.graph.getCollection("subscribedSkus", skuContext, {
      schema: z.unknown(),
      maxPages: 10,
      maxRecords: 1_000,
    });
    let unreadable = 0;
    for (const raw of result.items) {
      const parsed = entitlementSchema.safeParse(raw);
      if (!parsed.success) {
        unreadable += 1;
        continue;
      }
      const standing = ["enabled", "warning"].includes(lower(parsed.data.capabilityStatus ?? ""));
      const included = parsed.data.servicePlans.some(
        (plan) =>
          ["aad_premium", "aad_premium_p2"].includes(lower(plan.servicePlanName ?? "")) &&
          (plan.provisioningStatus === null || lower(plan.provisioningStatus) === "success"),
      );
      if (standing && included) return "yes";
    }
    return result.complete && unreadable === 0 ? "no" : "unknown";
  } catch (error) {
    if (context.signal.aborted) throw error;
    return "unknown";
  }
}

interface NamedUsers {
  // Display name by lower-case user ID; null for an account Microsoft no longer has.
  names: Map<string, string | null>;
  failed: boolean;
  capped: boolean;
}

async function nameUsers(context: AuditContext, ids: readonly string[]): Promise<NamedUsers> {
  const names = new Map<string, string | null>();
  let failed = false;
  const wanted = ids.slice(0, MAX_NAMED_USERS);
  for (let start = 0; start < wanted.length; start += NAME_BATCH) {
    context.signal.throwIfAborted();
    await Promise.all(
      wanted.slice(start, start + NAME_BATCH).map(async (id) => {
        try {
          const user = await context.graph.get(
            `users/${id}?$select=id,displayName,userPrincipalName`,
            usersContext,
            userNameSchema,
          );
          names.set(id, user.displayName ?? user.userPrincipalName ?? null);
        } catch (error) {
          if (context.signal.aborted) throw error;
          if (error instanceof GraphRequestError && error.status === 404) names.set(id, null);
          else failed = true;
        }
      }),
    );
  }
  return { names, failed, capped: ids.length > MAX_NAMED_USERS };
}

const REPORT_ONLY_NOTE =
  "Report-only policies are evaluated and logged but not enforced, so the checks above leave them out.";
const CHECKS_NOTE =
  "Blocks legacy authentication means an enabled policy for all users that blocks Exchange ActiveSync and other legacy clients. Requires MFA means the grant includes MFA or an authentication strength; when the operator is OR, another listed control can satisfy it instead. Neither check looks at apps, locations, risk levels or exclusions.";
const ROLES_NOTE =
  "Policies refer to roles by role template ID. Prism365 names the common administrator roles and counts the rest.";
const BREAK_GLASS_NOTE =
  "Users excluded from MFA policies often include break-glass (emergency access) accounts, which are excluded on purpose. Check that every other account on the list is meant to be there.";
const SECURITY_DEFAULTS_NOTE =
  "Security defaults are not Conditional Access policies and are not read here. When they are on, Microsoft blocks legacy authentication and requires MFA for administrators, so a missing policy above may not leave the tenant open.";

function emptyNote(entitlement: Entitlement): string {
  if (entitlement === "yes")
    return "Microsoft returned no policies. The tenant has Microsoft Entra ID P1 or P2, so Conditional Access is available and no policy has been created.";
  if (entitlement === "no")
    return "Microsoft returned no policies, and none of the tenant's subscriptions include Microsoft Entra ID P1 or P2, which Conditional Access needs.";
  return "Microsoft returned no policies. Whether the tenant has Microsoft Entra ID P1, which Conditional Access needs, could not be checked: reading subscriptions needs LicenseAssignment.Read.All.";
}

export interface ConditionalAccessOptions {
  // How the audit waits between requests to the throttled endpoints. Tests replace it.
  pause?: Pause;
}

export function createConditionalAccessAudit(options: ConditionalAccessOptions = {}): Audit {
  return {
    definition: {
      id: "identity.conditional-access",
      version: 1,
      title: "Conditional Access policies",
      description:
        "Every Conditional Access policy with who it applies to, who it leaves out and what it requires, and whether the tenant blocks legacy authentication and asks administrators for MFA.",
      category: "Identity",
      subject: { one: "policy", many: "policies" },
      columns: [
        { key: "policy", label: "Policy", kind: "text" },
        {
          key: "state",
          label: "State",
          kind: "enum",
          facet: true,
          help: "On enforces the policy. Report only evaluates and logs it without enforcing it.",
        },
        { key: "appliesTo", label: "Applies to", kind: "text" },
        {
          key: "excludes",
          label: "Leaves out",
          kind: "list",
          help: "Users, groups and roles the policy skips. User names appear when the session has granted User.Read.All.",
        },
        { key: "apps", label: "Apps", kind: "text" },
        {
          key: "conditions",
          label: "Conditions",
          kind: "list",
          help: "Client apps, platforms, locations and risk levels the policy narrows to. Locations are named from the tenant's named locations.",
        },
        {
          key: "grant",
          label: "Grant",
          kind: "text",
          help: "What a sign-in must satisfy. With one of, any listed control is enough.",
        },
        { key: "session", label: "Session", kind: "text", empty: "None" },
        { key: "modified", label: "Modified", kind: "date", empty: "Never" },
        { key: "created", label: "Created", kind: "date", hidden: true, empty: "Unknown" },
        { key: "excludedUsers", label: "Excluded users", kind: "number", hidden: true },
        { key: "excludedGroups", label: "Excluded groups", kind: "number", hidden: true },
        { key: "excludedRoles", label: "Excluded roles", kind: "number", hidden: true },
      ],
      views: [
        { id: "all", label: "All policies", tags: [] },
        { id: "enabled", label: "Enabled", tags: ["enabled"] },
        { id: "report-only", label: "Report only", tags: ["report-only"] },
        { id: "off", label: "Off", tags: ["off"] },
        { id: "legacy-block", label: "Blocks legacy authentication", tags: ["legacy-block"] },
        { id: "mfa", label: "Requires MFA", tags: ["mfa"] },
        { id: "exclusions", label: "Has exclusions", tags: ["exclusions"] },
      ],
      defaultView: "all",
      defaultSort: { column: "policy", direction: "asc" },
      requirements: [
        "A role such as Global Reader, Security Reader, Security Administrator or Conditional Access Administrator.",
        "Microsoft Entra ID P1 or P2.",
      ],
    },
    packs: ["identity.conditional-access.read"],
    optionalPacks: ["directory.users.read", "tenant.licenses.read"],
    scoped: false,
    forbidden:
      "Microsoft refused to list Conditional Access policies. It needs Policy.Read.All, a role such as Global Reader, Security Reader, Security Administrator or Conditional Access Administrator, and Microsoft Entra ID P1.",
    async run(context) {
      context.signal.throwIfAborted();
      const pacer: Pacer = createPacer(options.pause);
      const read = await readPaced(
        context.graph,
        "identity/conditionalAccess/policies",
        policyContext,
        {
          pacer,
          signal: context.signal,
          maxPages: MAX_PAGES,
          maxRecords: MAX_POLICIES,
        },
      );
      const policies: Policy[] = [];
      const seen = new Set<string>();
      let skipped = 0;
      for (const raw of read.items) {
        const parsed = policySchema.safeParse(raw);
        if (!parsed.success || seen.has(parsed.data.id)) {
          skipped += 1;
          continue;
        }
        seen.add(parsed.data.id);
        policies.push(parsed.data);
      }

      // Named locations are read only when a policy points at one, which saves a request that
      // Microsoft would make us wait a second for.
      const usesNamedLocations = policies.some((policy) =>
        [
          ...policy.conditions.locations.includeLocations,
          ...policy.conditions.locations.excludeLocations,
        ].some((id) => GUID.test(id)),
      );
      let locationNames: Map<string, string> | null = new Map();
      let locationsComplete = true;
      if (usesNamedLocations) {
        try {
          const locations = await readPaced(
            context.graph,
            "identity/conditionalAccess/namedLocations",
            policyContext,
            { pacer, signal: context.signal, maxPages: MAX_PAGES, maxRecords: MAX_LOCATIONS },
          );
          locationsComplete = locations.complete;
          for (const raw of locations.items) {
            const parsed = locationSchema.safeParse(raw);
            if (parsed.success)
              locationNames.set(
                lower(parsed.data.id),
                parsed.data.displayName ?? "Unnamed location",
              );
            else locationsComplete = false;
          }
        } catch (error) {
          if (context.signal.aborted) throw error;
          locationNames = null;
          locationsComplete = false;
        }
      }
      const locations: LocationNames = { names: locationNames };

      const excludedIds = [
        ...new Set(policies.flatMap((policy) => userIds(policy.conditions.users.excludeUsers))),
      ];
      const canName = context.granted.has("User.Read.All");
      const named =
        canName && excludedIds.length > 0 ? await nameUsers(context, excludedIds) : null;

      const rows: AuditRow[] = [];
      for (const policy of policies) {
        const state = stateOf(policy);
        const users = policy.conditions.users;
        const excludes = excludesOf(policy, named?.names ?? null);
        rows.push({
          key: policy.id,
          personId: null,
          tags: [
            ...(state === "On" ? ["enabled"] : []),
            ...(state === "Report only" ? ["report-only"] : []),
            ...(state === "Off" ? ["off"] : []),
            ...(blocksLegacyAuthentication(policy) ? ["legacy-block"] : []),
            ...(state === "On" && requiresMfa(policy) ? ["mfa"] : []),
            ...(excludes.length > 0 ? ["exclusions"] : []),
          ],
          cells: {
            policy: policy.displayName ?? "Unnamed policy",
            state,
            appliesTo: appliesTo(policy),
            excludes,
            apps: appsText(policy),
            conditions: conditionsOf(policy, locations),
            grant: grantText(policy),
            session: sessionText(policy),
            modified: policy.modifiedDateTime,
            created: policy.createdDateTime,
            excludedUsers: userIds(users.excludeUsers).length,
            excludedGroups: users.excludeGroups.length,
            excludedRoles: users.excludeRoles.length,
          },
        });
      }
      rows.sort((a, b) => String(a.cells.policy).localeCompare(String(b.cells.policy)));

      const empty = read.items.length === 0;
      const entitlement = empty ? await entraP1(context) : null;
      const count = (tag: string) => rows.filter((row) => row.tags.includes(tag)).length;
      const unknownState = rows.filter((row) => row.cells.state === "Unknown").length;
      // With no policies and no proof that Conditional Access is available, a missing policy is
      // not a finding: the tenant may not be able to have one.
      const assessable = !empty || entitlement === "yes";
      const noLegacyBlock = assessable && count("legacy-block") === 0;
      const noAdminMfa = assessable && !policies.some(requiresMfaForAdministrators);
      const mfaExcluded = new Set(
        policies
          .filter((policy) => stateOf(policy) === "On" && requiresMfa(policy))
          .flatMap((policy) => userIds(policy.conditions.users.excludeUsers)),
      );
      const hasExcludedUsers = excludedIds.length > 0;

      return {
        rows,
        complete:
          read.complete && skipped === 0 && locationsComplete && (named === null || !named.failed),
        summary: [
          { label: "Policies enabled", value: count("enabled"), tone: "neutral", view: "enabled" },
          {
            label: "Policies in report-only",
            value: count("report-only"),
            tone: "neutral",
            view: "report-only",
          },
          { label: "Policies off", value: count("off"), tone: "neutral", view: "off" },
          ...(noLegacyBlock
            ? [
                {
                  label: "No enabled policy blocks legacy authentication",
                  value: 1,
                  tone: "critical" as const,
                  view: "all",
                },
              ]
            : []),
          ...(noAdminMfa
            ? [
                {
                  label: "No enabled policy requires MFA for administrators",
                  value: 1,
                  tone: "warning" as const,
                  view: "all",
                },
              ]
            : []),
          {
            label: "Users excluded from MFA policies",
            value: mfaExcluded.size,
            tone: "neutral",
            view: "mfa",
          },
        ],
        notes: [
          ...(empty ? [emptyNote(entitlement ?? "unknown")] : []),
          ...(unknownState > 0
            ? [
                `${plural(unknownState, "policy has", "policies have")} a state Prism365 does not recognise, so ${unknownState === 1 ? "it is" : "they are"} counted as neither on, report-only nor off.`,
              ]
            : []),
          REPORT_ONLY_NOTE,
          CHECKS_NOTE,
          ...(noLegacyBlock || noAdminMfa ? [SECURITY_DEFAULTS_NOTE] : []),
          ROLES_NOTE,
          ...(mfaExcluded.size > 0 ? [BREAK_GLASS_NOTE] : []),
          ...(usesNamedLocations && locationNames === null
            ? ["Named locations could not be read, so locations are listed without names."]
            : []),
          ...(hasExcludedUsers && !canName
            ? [
                "Excluded users are shown as counts. Their names need User.Read.All, which the session has not granted.",
              ]
            : []),
          ...(named?.failed ? ["Some excluded users could not be named."] : []),
          ...(named?.capped
            ? [`Names were looked up for the first ${MAX_NAMED_USERS} excluded users only.`]
            : []),
          ...(skipped > 0
            ? [
                `${plural(skipped, "policy", "policies")} could not be read and ${skipped === 1 ? "is" : "are"} left out, so a check above may be missing a policy that would satisfy it.`,
              ]
            : []),
          ...(read.complete
            ? []
            : ["Stopped before the end of the policy list; the rest was not read."]),
        ],
      };
    },
  };
}

export const conditionalAccessAudit: Audit = createConditionalAccessAudit();

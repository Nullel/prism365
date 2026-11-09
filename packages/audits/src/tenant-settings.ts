import { type AuditRow, type AuditSummaryItem, cleanText } from "@prism365/contracts";
import { GraphBoundaryError, GraphRequestError, GraphRetryDeferredError } from "@prism365/graph";
import { z } from "zod";
import { FRESH_DAYS, freshnessOf } from "./combined-risks-inputs.ts";
import { ALL_USERS } from "./conditional-access.ts";
import type { Audit, AuditContext } from "./definition.ts";

// The tenant-wide settings an administrator checks first on any tenant review: whether security
// defaults are on, which sign-in methods the tenant allows, who can invite guests, who can
// register applications and who can consent to them. Five single reads under Policy.Read.All, one
// request each. A setting differs from the expected value is "Review", never "Fails": a tenant can
// have a reason, and Prism365 cannot see it.
// Microsoft Graph v1.0 documentation:
// https://learn.microsoft.com/en-us/graph/api/identitysecuritydefaultsenforcementpolicy-get?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/authenticationmethodspolicy-get?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/authorizationpolicy-get?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/adminconsentrequestpolicy-get?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/permissiongrantpolicy-list?view=graph-rest-1.0

export const TENANT_SETTINGS_AUDIT_ID = "tenant.settings";
// The row Combined risks reads, and the value it shows when security defaults are on themselves
// (not when a Conditional Access policy stands in for them).
export const SECURITY_DEFAULTS_KEY = "security-defaults";
export const SECURITY_DEFAULTS_ON = "On";
const CONDITIONAL_ACCESS_AUDIT_ID = "identity.conditional-access";

const policyContext = { feature: "identity.policies.read", permission: "Policy.Read.All" } as const;
const MAX_GRANT_POLICIES = 200;
const MAX_VALUE = 400;

export const AREAS = [
  "Security defaults",
  "Authentication methods",
  "Guests",
  "Applications",
  "Consent",
  "Users",
] as const;
type Area = (typeof AREAS)[number];

const areaTag = (area: Area) => area.toLowerCase().replaceAll(" ", "-");

type Status = "Meets" | "Review" | "Not read";
const statusTag: Record<Status, string> = {
  Meets: "meets",
  Review: "review",
  "Not read": "not-read",
};

// The restricted, limited and same-as-members guest roles Microsoft documents by ID.
const GUEST_ROLES = new Map([
  ["2af84b1e-32c8-42b7-82bc-daa82404023b", "Restricted access"],
  ["10dae51f-b6af-4016-8d66-8c2a99b929b3", "Limited access"],
  ["a0b1b346-4d3e-4e8b-98f8-753987be4970", "Same access as members"],
]);
const RESTRICTED_GUEST_ROLE = "2af84b1e-32c8-42b7-82bc-daa82404023b";

const GRANT_PREFIX = /^[^.]*\./;
const LOW_IMPACT = "microsoft-user-default-low";
// What Microsoft's two long-standing consent policies mean, for when its policy list does not
// give a display name.
const KNOWN_GRANT_POLICIES = new Map([
  [LOW_IMPACT, "Verified publishers, low impact"],
  ["microsoft-user-default-legacy", "Any app, any permission a user can consent to (legacy)"],
]);

// Reading ---------------------------------------------------------------------------------------

// Why a policy could not be read, as a fixed phrase that follows a colon, or null for an error the
// audit must not swallow (a session that ended, or anything that is a bug rather than an answer).
// Never provider text.
export function failureOf(error: unknown): string | null {
  if (error instanceof GraphRetryDeferredError)
    return "Microsoft is limiting requests, so run the audit again in a few minutes";
  if (error instanceof GraphRequestError) {
    if (error.status === 401) return null;
    if (error.status === 403) return "Microsoft refused (HTTP 403)";
    if (error.status === 404) return "Microsoft has no such policy (HTTP 404)";
    return `Microsoft answered HTTP ${error.status}`;
  }
  if (error instanceof GraphBoundaryError) {
    if (error.reason === "credential") return null;
    if (error.reason === "timeout") return "Microsoft took too long to answer";
    return `its answer could not be read safely (${error.reason})`;
  }
  return null;
}

type Record_ = Record<string, unknown>;
const isRecord = (value: unknown): value is Record_ =>
  typeof value === "object" && value !== null && !Array.isArray(value);

// Most of these policies are one object. Microsoft answers the authorization policy as a list of
// one in some versions, so a list is read by its first entry.
function single(body: unknown): Record_ | null {
  if (isRecord(body) && Array.isArray(body.value)) {
    const first = body.value[0];
    return isRecord(first) ? first : null;
  }
  return isRecord(body) ? body : null;
}

type Policy = { state: "read"; data: Record_ } | { state: "failed"; reason: string };

async function readPolicy(context: AuditContext, path: string): Promise<Policy> {
  context.signal.throwIfAborted();
  try {
    const data = single(await context.graph.get(path, policyContext, z.unknown()));
    return data ? { state: "read", data } : { state: "failed", reason: "its answer was empty" };
  } catch (error) {
    if (context.signal.aborted) throw error;
    const reason = failureOf(error);
    if (reason === null) throw error;
    return { state: "failed", reason };
  }
}

interface GrantNames {
  names: Map<string, string>;
  // Why the list could not be read in full, or null.
  problem: string | null;
}

// The permission grant policies only supply display names for the consent row. Reading them can
// fail without costing any row its value.
async function readGrantNames(context: AuditContext): Promise<GrantNames> {
  const names = new Map<string, string>();
  context.signal.throwIfAborted();
  try {
    const read = await context.graph.getCollection(
      "policies/permissionGrantPolicies",
      policyContext,
      { schema: z.unknown(), maxPages: 3, maxRecords: MAX_GRANT_POLICIES },
    );
    for (const item of read.items) {
      if (!isRecord(item) || typeof item.id !== "string") continue;
      const name = typeof item.displayName === "string" ? item.displayName.trim() : "";
      if (name) names.set(item.id.toLowerCase(), cleanText(name.slice(0, 120)));
    }
    return { names, problem: read.complete ? null : "the list was cut short" };
  } catch (error) {
    if (context.signal.aborted) throw error;
    const reason = failureOf(error);
    if (reason === null) throw error;
    return { names, problem: reason };
  }
}

// What a row says --------------------------------------------------------------------------------

interface Reading {
  status: Status;
  // Microsoft's value, in words.
  value: string;
  tags?: string[];
}

interface Spec {
  key: string;
  setting: string;
  area: Area;
  expected: string;
  why: string;
  // Where the value comes from: the policy and the field.
  path: string;
}

const CHOICE = "Your choice";

const meets = (value: string, tags: string[] = []): Reading => ({ status: "Meets", value, tags });
const review = (value: string): Reading => ({ status: "Review", value });
const missing = (): Reading => ({
  status: "Not read",
  value: "Microsoft did not return this setting",
});
const unread = (reason: string): Reading => ({
  status: "Not read",
  value: `${reason.charAt(0).toUpperCase()}${reason.slice(1)}`,
});

const bool = (value: unknown): boolean | null => (typeof value === "boolean" ? value : null);
const word = (value: unknown, max = 80): string | null =>
  typeof value === "string" && value.trim() !== "" ? cleanText(value.trim().slice(0, max)) : null;

// A reading of one boolean field: Meets when it is the expected value, Review when it is not,
// Not read when Microsoft left it out.
function flag(
  value: unknown,
  words: { true: string; false: string },
  expected: boolean | null,
): Reading {
  const found = bool(value);
  if (found === null) return missing();
  const shown = found ? words.true : words.false;
  return expected === null || found === expected ? meets(shown) : review(shown);
}

// Security defaults ------------------------------------------------------------------------------

interface DefaultsReading {
  reading: Reading;
  note: string | null;
}

async function securityDefaults(context: AuditContext, data: Record_): Promise<DefaultsReading> {
  const enabled = bool(data.isEnabled);
  if (enabled === null) return { reading: missing(), note: null };
  if (enabled) return { reading: meets(SECURITY_DEFAULTS_ON), note: null };

  const run = context.stored
    ? (context.signal.throwIfAborted(), await context.stored.latestRun(CONDITIONAL_ACCESS_AUDIT_ID))
    : null;
  const fresh = freshnessOf(run, context.now);
  if (fresh.state !== "fresh") {
    const when =
      fresh.state === "never"
        ? "has never run"
        : `last ran ${fresh.finishedAt.slice(0, 10)}, more than ${FRESH_DAYS} days ago`;
    return {
      reading: review("Off. Run Conditional Access policies first"),
      note: `Security defaults are off, so Prism365 looked for a Conditional Access policy that requires MFA for all users. The Conditional Access policies audit ${when}. Run it, then run this audit again.`,
    };
  }
  const day = fresh.finishedAt.slice(0, 10);
  const policy = fresh.run.rows.find(
    (row) =>
      row.tags.includes("enabled") && row.tags.includes("mfa") && row.cells.appliesTo === ALL_USERS,
  );
  if (policy) {
    const name = word(policy.cells.policy, 120) ?? "an unnamed policy";
    const exclusions = policy.tags.includes("exclusions")
      ? " It lists exclusions, which this audit does not weigh."
      : "";
    const unweighed =
      " Apps, conditions and grants that accept another control in place of MFA are not weighed either.";
    return {
      reading: meets("Off, with a Conditional Access policy that requires MFA for all users"),
      note: `Security defaults are off. The Conditional Access policies run of ${day} has an enabled policy that requires MFA for all users: ${name}.${exclusions}${unweighed}`,
    };
  }
  const partial = fresh.run.complete
    ? ""
    : " That run returned a partial result, so a policy may be missing from it.";
  return {
    reading: review("Off. No enabled Conditional Access policy requires MFA for all users"),
    note: `Security defaults are off, and the Conditional Access policies run of ${day} has no enabled policy that requires MFA for all users.${partial}`,
  };
}

// Authentication methods -------------------------------------------------------------------------

interface MethodSpec extends Spec {
  // The configuration's id in Microsoft's answer.
  id: string;
  want: "enabled" | "disabled" | null;
}

const methodPath = "policies/authenticationMethodsPolicy: authenticationMethodConfigurations";
const method = (
  id: string,
  key: string,
  setting: string,
  want: MethodSpec["want"],
  why: string,
): MethodSpec => ({
  id,
  key: `method-${key}`,
  setting,
  area: "Authentication methods",
  want,
  expected: want === null ? CHOICE : want === "enabled" ? "Enabled" : "Disabled",
  why,
  path: `${methodPath} (${id})`,
});

const METHODS: readonly MethodSpec[] = [
  method(
    "Sms",
    "sms",
    "Text message sign-in",
    "disabled",
    "Text messages can be intercepted or redirected with a SIM swap, so app-based methods are safer.",
  ),
  method(
    "Voice",
    "voice",
    "Voice call sign-in",
    "disabled",
    "Phone calls share the weaknesses of text messages.",
  ),
  method(
    "Email",
    "email",
    "Email one-time passcode",
    null,
    "Lets guests without a Microsoft account sign in with a code sent to their email address.",
  ),
  method(
    "Fido2",
    "fido2",
    "Security keys and passkeys",
    "enabled",
    "Security keys and passkeys resist phishing, which makes them the strongest method for administrators.",
  ),
  method(
    "MicrosoftAuthenticator",
    "authenticator",
    "Microsoft Authenticator",
    "enabled",
    "The app is the usual strong second step for most people and supports sign-in without a password.",
  ),
  method(
    "TemporaryAccessPass",
    "tap",
    "Temporary Access Pass",
    "enabled",
    "A short-lived passcode that lets someone register a first strong method or recover an account without a password.",
  ),
  method(
    "SoftwareOath",
    "oath",
    "Authenticator apps with codes",
    null,
    "Allows time-based codes from authenticator apps other than Microsoft's.",
  ),
  method(
    "X509Certificate",
    "x509",
    "Certificate-based sign-in",
    null,
    "Allows sign-in with a certificate, which suits tenants that issue smart cards.",
  ),
];

function configurationsOf(policy: Record_): Map<string, Record_> {
  const found = new Map<string, Record_>();
  const list = policy.authenticationMethodConfigurations;
  if (!Array.isArray(list)) return found;
  for (const item of list) {
    if (isRecord(item) && typeof item.id === "string") found.set(item.id.toLowerCase(), item);
  }
  return found;
}

function methodReading(spec: MethodSpec, configurations: Map<string, Record_>): Reading {
  const state = word(configurations.get(spec.id.toLowerCase())?.state)?.toLowerCase();
  if (state !== "enabled" && state !== "disabled") return missing();
  const shown = state === "enabled" ? "Enabled" : "Disabled";
  return spec.want === null || spec.want === state ? meets(shown) : review(shown);
}

const MIGRATION_STATES = new Map([
  ["premigration", "Not started"],
  ["migrationinprogress", "In progress"],
  ["migrationcomplete", "Complete"],
]);

const CAMPAIGN_STATES = new Map([
  ["enabled", "Enabled"],
  ["disabled", "Disabled"],
  ["default", "Managed by Microsoft"],
]);

function campaignReading(policy: Record_): Reading {
  const enforcement = policy.registrationEnforcement;
  const campaign = isRecord(enforcement)
    ? enforcement.authenticationMethodsRegistrationCampaign
    : null;
  const state = isRecord(campaign) ? word(campaign.state)?.toLowerCase() : undefined;
  if (!state) return missing();
  const shown = CAMPAIGN_STATES.get(state) ?? cleanText(state);
  return state === "enabled" ? meets(shown) : review(shown);
}

function migrationReading(policy: Record_): Reading {
  const state = word(policy.policyMigrationState);
  if (!state) return missing();
  const shown = MIGRATION_STATES.get(state.toLowerCase()) ?? state;
  return state.toLowerCase() === "migrationcomplete" ? meets(shown) : review(shown);
}

// The authorization policy -----------------------------------------------------------------------

const INVITES = new Map([
  ["none", "No one"],
  ["adminsandguestinviters", "Only administrators and guest inviters"],
  ["adminsguestinvitersandallmembers", "Administrators, guest inviters and all members"],
  ["everyone", "Everyone, including guests"],
]);

function invitesReading(policy: Record_): Reading {
  const found = word(policy.allowInvitesFrom);
  if (!found) return missing();
  const shown = INVITES.get(found.toLowerCase()) ?? found;
  const key = found.toLowerCase();
  return key === "none" || key === "adminsandguestinviters" ? meets(shown) : review(shown);
}

function guestRoleReading(policy: Record_): Reading {
  const found = word(policy.guestUserRoleId, 64);
  if (!found) return missing();
  const id = found.toLowerCase();
  const shown = GUEST_ROLES.get(id) ?? "A role Prism365 does not recognise";
  return id === RESTRICTED_GUEST_ROLE ? meets(shown) : review(shown);
}

function defaultRole(policy: Record_, field: string): unknown {
  return isRecord(policy.defaultUserRolePermissions)
    ? policy.defaultUserRolePermissions[field]
    : undefined;
}

type ConsentState = "off" | "limited" | "open";

// Microsoft keeps two kinds of entry in one list. ForSelf entries say what users may consent to
// for themselves. ForOwnedResource entries let group owners (and Teams chat and team owners)
// consent for the one group or team they own, which is not user consent to apps tenant-wide.
const OWNED_RESOURCE = /^ManagePermissionGrantsForOwnedResource\./i;

interface Grants {
  self: string[];
  owned: string[];
}

// null when the list is missing or holds anything but text: a malformed list is not read as "off".
function assignedGrants(policy: Record_): Grants | null {
  const assigned = defaultRole(policy, "permissionGrantPoliciesAssigned");
  if (!Array.isArray(assigned) || assigned.some((entry) => typeof entry !== "string")) return null;
  const entries = (assigned as string[]).map((entry) => entry.slice(0, 200));
  return {
    self: entries.filter((entry) => !OWNED_RESOURCE.test(entry)),
    owned: entries.filter((entry) => OWNED_RESOURCE.test(entry)),
  };
}

// Judged on the ForSelf entries only.
function consentStateOf(grants: Grants): ConsentState {
  if (grants.self.length === 0) return "off";
  const only =
    grants.self.length === 1 ? grants.self[0]?.replace(GRANT_PREFIX, "").toLowerCase() : null;
  return only === LOW_IMPACT ? "limited" : "open";
}

function grantName(entry: string, names: Map<string, string>): string {
  const id = entry.replace(GRANT_PREFIX, "").toLowerCase();
  return names.get(id) ?? KNOWN_GRANT_POLICIES.get(id) ?? cleanText(entry.slice(0, 120));
}

function consentReading(grants: Grants | null, names: Map<string, string>): Reading {
  if (grants === null) return missing();
  const state = consentStateOf(grants);
  const nameList = (entries: readonly string[]) =>
    entries.map((entry) => grantName(entry, names)).join(", ");
  const self =
    state === "off" ? "Off. Users cannot consent to apps" : `On: ${nameList(grants.self)}`;
  // Shown for reference; it never changes the status.
  const owned =
    grants.owned.length > 0
      ? `. Group owners can consent for their groups and teams: ${nameList(grants.owned)}`
      : "";
  const shown = `${self}${owned}`.slice(0, MAX_VALUE);
  return state === "open" ? review(shown) : meets(shown);
}

// What the admin consent workflow is judged against: whether users can already consent on their
// own. null when that could not be read.
function workflowReading(policy: Record_, consent: ConsentState | null): Reading {
  const enabled = bool(policy.isEnabled);
  if (enabled === null) return missing();
  const shown = enabled ? "On" : "Off";
  if (enabled || consent === "open") return meets(shown);
  return review(shown);
}

const WORKFLOW_EXPECTED = "On, unless users can consent to any app";

// The rows ---------------------------------------------------------------------------------------

const AUTHORIZATION = "policies/authorizationPolicy";
const authorization = (
  key: string,
  setting: string,
  area: Area,
  field: string,
  expected: string,
  why: string,
): Spec => ({ key, setting, area, expected, why, path: `${AUTHORIZATION}: ${field}` });

const SPECS = {
  defaults: {
    key: SECURITY_DEFAULTS_KEY,
    setting: "Security defaults",
    area: "Security defaults",
    expected:
      "On, or off with an enabled Conditional Access policy that requires MFA for all users",
    why: "Security defaults turn on a baseline of MFA and sign-in protections for every account.",
    path: "policies/identitySecurityDefaultsEnforcementPolicy: isEnabled",
  } satisfies Spec,
  campaign: {
    key: "registration-campaign",
    setting: "Registration campaign",
    area: "Authentication methods",
    expected: "Enabled",
    why: "Prompts people at sign-in to set up the Authenticator app, so more accounts have a strong method.",
    path: "policies/authenticationMethodsPolicy: registrationEnforcement",
  } satisfies Spec,
  migration: {
    key: "policy-migration",
    setting: "Legacy MFA and password reset policies",
    area: "Authentication methods",
    expected: "Complete (the legacy policies are retired)",
    why: "Until the old per-user MFA and password reset policies are retired, sign-in methods are managed in two places.",
    path: "policies/authenticationMethodsPolicy: policyMigrationState",
  } satisfies Spec,
  invites: authorization(
    "guest-invites",
    "Who can invite guests",
    "Guests",
    "allowInvitesFrom",
    "Only administrators and guest inviters, or no one",
    "Controls who can bring outside people into the directory.",
  ),
  guestRole: authorization(
    "guest-role",
    "What guests can see",
    "Guests",
    "guestUserRoleId",
    "Restricted access",
    "Controls how much of the directory a guest can read, such as other people and groups.",
  ),
  emailJoin: authorization(
    "email-verified-join",
    "Join by email verification",
    "Users",
    "allowEmailVerifiedUsersToJoinOrganization",
    "Not allowed",
    "Lets people join the tenant on their own by proving they control an email address.",
  ),
  subscriptions: authorization(
    "email-subscriptions",
    "Email-based subscription sign-up",
    "Users",
    "allowedToSignUpEmailBasedSubscriptions",
    "Not allowed",
    "Lets people sign up for email-based subscriptions, which can create accounts in the tenant.",
  ),
  createApps: authorization(
    "create-apps",
    "Users can register applications",
    "Applications",
    "defaultUserRolePermissions.allowedToCreateApps",
    "Not allowed",
    "Lets any user register an application, which can then ask others for access.",
  ),
  createGroups: authorization(
    "create-security-groups",
    "Users can create security groups",
    "Users",
    "defaultUserRolePermissions.allowedToCreateSecurityGroups",
    CHOICE,
    "Lets any user create security groups.",
  ),
  readUsers: authorization(
    "read-other-users",
    "Users can read other users",
    "Users",
    "defaultUserRolePermissions.allowedToReadOtherUsers",
    CHOICE,
    "Lets any user look up other people in the directory.",
  ),
  consent: authorization(
    "user-consent",
    "User consent to apps",
    "Consent",
    "defaultUserRolePermissions.permissionGrantPoliciesAssigned",
    "Off, or only apps from verified publishers with low-impact permissions",
    "Controls whether users can approve an app's access to company data without an administrator.",
  ),
  sspr: authorization(
    "self-service-password-reset",
    "Users can use self-service password reset",
    "Users",
    "allowedToUseSSPR",
    CHOICE,
    "Lets users reset their own password when self-service password reset is set up.",
  ),
  msol: authorization(
    "msol-powershell",
    "Retired MSOnline PowerShell module",
    "Users",
    "blockMsolPowerShell",
    CHOICE,
    "Stops the retired MSOnline PowerShell module from reading or changing the directory.",
  ),
  workflow: {
    key: "admin-consent-workflow",
    setting: "Admin consent workflow",
    area: "Consent",
    expected: WORKFLOW_EXPECTED,
    why: "Without the workflow, users have no way to ask for an app to be approved.",
    path: "policies/adminConsentRequestPolicy: isEnabled",
  } satisfies Spec,
};

function rowOf(spec: Spec, reading: Reading): AuditRow {
  return {
    key: spec.key,
    personId: null,
    tags: [statusTag[reading.status], areaTag(spec.area), ...(reading.tags ?? [])],
    cells: {
      setting: spec.setting,
      area: spec.area,
      value: reading.value.slice(0, MAX_VALUE),
      expected: spec.expected,
      status: reading.status,
      why: spec.why,
      path: spec.path,
    },
  };
}

const plural = (count: number, one: string, many: string) => `${count} ${count === 1 ? one : many}`;
const statusOrder: Record<Status, number> = { Review: 0, "Not read": 1, Meets: 2 };

const POLICY_NAMES = {
  defaults: "security defaults policy",
  methods: "authentication methods policy",
  authorization: "authorization policy",
  workflow: "admin consent request policy",
} as const;

export const tenantSettingsAudit: Audit = {
  definition: {
    id: TENANT_SETTINGS_AUDIT_ID,
    version: 1,
    title: "Tenant settings",
    description:
      "The settings that apply to the whole tenant: security defaults, the sign-in methods it allows, who can invite guests, who can register applications and who can consent to them.",
    category: "Identity",
    subject: { one: "setting", many: "settings" },
    columns: [
      { key: "setting", label: "Setting", kind: "text" },
      { key: "area", label: "Area", kind: "enum", facet: true },
      {
        key: "value",
        label: "Value",
        kind: "text",
        help: "What Microsoft reports for this tenant, in words.",
      },
      {
        key: "expected",
        label: "Expected",
        kind: "text",
        help: "The value Prism365 expects. Your choice means the setting is shown for reference and never needs review.",
      },
      {
        key: "status",
        label: "Status",
        kind: "enum",
        facet: true,
        help: "Meets when the value is the expected one. Review when it differs, which a tenant can have a reason for. Not read when Microsoft did not return it.",
      },
      { key: "why", label: "What it controls", kind: "text" },
      {
        key: "path",
        label: "Read from",
        kind: "text",
        hidden: true,
        help: "The Microsoft policy and field the value comes from.",
      },
    ],
    views: [
      { id: "all", label: "All settings", tags: [] },
      { id: "review", label: "Review", tags: ["review"] },
      { id: "meets", label: "Meets", tags: ["meets"] },
      { id: "not-read", label: "Not read", tags: ["not-read"] },
      ...AREAS.map((area) => ({ id: areaTag(area), label: area, tags: [areaTag(area)] })),
    ],
    defaultView: "all",
    // Status labels sort Review, Not read, Meets when descending, and the rows are stored in area
    // order, which a stable sort keeps within each status.
    defaultSort: { column: "status", direction: "desc" },
    requirements: [
      "A role that can read policies, such as Global Reader or Security Reader.",
      "Run Conditional Access policies first. Without a run from the last seven days, security defaults that are off cannot be checked against a policy that requires MFA.",
    ],
  },
  packs: ["identity.policies.read"],
  optionalPacks: [],
  scoped: false,
  usesStoredResults: true,
  forbidden:
    "Microsoft refused to read the tenant's policies. It needs Policy.Read.All and a role such as Global Reader or Security Reader.",
  async run(context) {
    const defaultsPolicy = await readPolicy(
      context,
      "policies/identitySecurityDefaultsEnforcementPolicy",
    );
    const methodsPolicy = await readPolicy(context, "policies/authenticationMethodsPolicy");
    const authorizationPolicy = await readPolicy(context, AUTHORIZATION);
    const workflowPolicy = await readPolicy(context, "policies/adminConsentRequestPolicy");
    const grants = await readGrantNames(context);

    const readings: [Spec, Reading][] = [];
    const notes: string[] = [];
    const failed: string[] = [];
    const fail = (name: string, policy: Policy, rows: string) => {
      if (policy.state !== "failed") return;
      failed.push(name);
      notes.push(
        `The ${name} could not be read: ${policy.reason}. ${rows} show Not read.`.slice(0, 500),
      );
    };

    // Security defaults
    if (defaultsPolicy.state === "read") {
      const found = await securityDefaults(context, defaultsPolicy.data);
      readings.push([SPECS.defaults, found.reading]);
      if (found.note) notes.push(found.note);
    } else {
      readings.push([SPECS.defaults, unread(defaultsPolicy.reason)]);
      fail(POLICY_NAMES.defaults, defaultsPolicy, "Its row shows");
    }

    // Authentication methods
    if (methodsPolicy.state === "read") {
      const configurations = configurationsOf(methodsPolicy.data);
      for (const spec of METHODS) readings.push([spec, methodReading(spec, configurations)]);
      readings.push([SPECS.campaign, campaignReading(methodsPolicy.data)]);
      readings.push([SPECS.migration, migrationReading(methodsPolicy.data)]);
    } else {
      for (const spec of [...METHODS, SPECS.campaign, SPECS.migration])
        readings.push([spec, unread(methodsPolicy.reason)]);
      fail(POLICY_NAMES.methods, methodsPolicy, "Its rows show");
    }

    // The authorization policy and the consent it carries
    let consent: ConsentState | null = null;
    if (authorizationPolicy.state === "read") {
      const data = authorizationPolicy.data;
      const assigned = assignedGrants(data);
      consent = assigned === null ? null : consentStateOf(assigned);
      const role = (field: string) => defaultRole(data, field);
      const allowed = { true: "Allowed", false: "Not allowed" };
      readings.push(
        [SPECS.invites, invitesReading(data)],
        [SPECS.guestRole, guestRoleReading(data)],
        [SPECS.emailJoin, flag(data.allowEmailVerifiedUsersToJoinOrganization, allowed, false)],
        [SPECS.subscriptions, flag(data.allowedToSignUpEmailBasedSubscriptions, allowed, false)],
        [SPECS.createApps, flag(role("allowedToCreateApps"), allowed, false)],
        [SPECS.createGroups, flag(role("allowedToCreateSecurityGroups"), allowed, null)],
        [SPECS.readUsers, flag(role("allowedToReadOtherUsers"), allowed, null)],
        [SPECS.consent, consentReading(assigned, grants.names)],
        [SPECS.sspr, flag(data.allowedToUseSSPR ?? role("allowedToUseSSPR"), allowed, null)],
        [
          SPECS.msol,
          flag(data.blockMsolPowerShell, { true: "Blocked", false: "Not blocked" }, null),
        ],
      );
    } else {
      for (const spec of [
        SPECS.invites,
        SPECS.guestRole,
        SPECS.emailJoin,
        SPECS.subscriptions,
        SPECS.createApps,
        SPECS.createGroups,
        SPECS.readUsers,
        SPECS.consent,
        SPECS.sspr,
        SPECS.msol,
      ])
        readings.push([spec, unread(authorizationPolicy.reason)]);
      fail(POLICY_NAMES.authorization, authorizationPolicy, "Its rows show");
    }

    // The admin consent workflow
    if (workflowPolicy.state === "read") {
      const reading = workflowReading(workflowPolicy.data, consent);
      const spec: Spec = {
        ...SPECS.workflow,
        expected: consent === "open" ? CHOICE : consent === null ? WORKFLOW_EXPECTED : "On",
        ...(consent === "open"
          ? {
              why: "Lets users ask an administrator to approve an app they cannot consent to themselves.",
            }
          : {}),
      };
      readings.push([spec, reading]);
    } else {
      readings.push([SPECS.workflow, unread(workflowPolicy.reason)]);
      fail(POLICY_NAMES.workflow, workflowPolicy, "Its row shows");
    }

    const rows = readings
      .map(([spec, reading], index) => ({
        row: rowOf(spec, reading),
        index,
        status: reading.status,
      }))
      .sort((a, b) => statusOrder[a.status] - statusOrder[b.status] || a.index - b.index)
      .map((item) => item.row);

    const count = (tag: string) => rows.filter((row) => row.tags.includes(tag)).length;
    const unreadable = count("not-read");
    const summary: AuditSummaryItem[] = [
      { label: "Review", value: count("review"), tone: "warning", view: "review" },
      { label: "Meets", value: count("meets"), tone: "good", view: "meets" },
      { label: "Not read", value: unreadable, tone: "neutral", view: "not-read" },
    ];

    const notesOut = [
      "Review means a setting differs from the value Prism365 expects. A tenant can have a reason for that, so it is a question to ask, not a verdict. Settings marked Your choice are shown for reference.",
      ...notes,
    ];
    const missingRows = rows.filter(
      (row) => row.tags.includes("not-read") && row.cells.value === missing().value,
    ).length;
    if (missingRows > 0)
      notesOut.push(
        `${plural(missingRows, "setting was", "settings were")} missing from Microsoft's answer and ${missingRows === 1 ? "shows" : "show"} Not read.`,
      );
    if (grants.problem)
      notesOut.push(
        `The permission grant policy list could not be read in full: ${grants.problem}. User consent shows Microsoft's policy ids where a name is missing.`,
      );
    notesOut.push(
      "Prism365 does not see per-user authentication method settings, self-service password reset configuration beyond whether users may use it, or Conditional Access coverage beyond a policy that requires MFA for all users.",
    );

    return {
      rows,
      summary,
      notes: notesOut.slice(0, 20),
      complete: failed.length === 0 && grants.problem === null,
    };
  },
};

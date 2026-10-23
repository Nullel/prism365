import { type AuditRow, cleanText } from "@prism365/contracts";
import { z } from "zod";
import type { Audit } from "./definition.ts";
import { byName, coverageNotes, MAX_ACCOUNTS, tagOf, typeOf } from "./directory.ts";
import {
  domainColumn,
  domainInventory,
  excludedNote,
  inAuditScope,
  principalDomain,
  scopeNote,
} from "./scope.ts";

// Who has registered a strong sign-in method, from Microsoft's authentication methods report.
// Microsoft Graph v1.0 documentation verified 2026-10-22:
// https://learn.microsoft.com/en-us/graph/api/authenticationmethodsroot-list-userregistrationdetails?view=graph-rest-1.0

const reportContext = {
  feature: "reports.authentication-methods.read",
  permission: "AuditLog.Read.All",
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
const names = z
  .array(z.string().max(100))
  .max(100)
  .nullish()
  .catch(null)
  .transform((value) => value ?? []);

const registrationSchema = z.object({
  id: z
    .string()
    .min(1)
    .max(200)
    .regex(/^[A-Za-z0-9._@:-]+$/),
  userPrincipalName: text,
  userDisplayName: text,
  userType: text,
  isAdmin: flag,
  isMfaRegistered: flag,
  isMfaCapable: flag,
  isPasswordlessCapable: flag,
  isSsprRegistered: flag,
  methodsRegistered: names,
  isSystemPreferredAuthenticationMethodEnabled: flag,
  systemPreferredAuthenticationMethods: names,
  userPreferredMethodForSecondaryAuthentication: text,
  lastUpdatedDateTime: z
    .string()
    .datetime({ offset: true })
    .nullish()
    .catch(null)
    .transform((value) => value ?? null),
});

const methodLabels: Record<string, string> = {
  mobilephone: "Mobile phone (text or call)",
  alternatemobilephone: "Alternate mobile phone",
  officephone: "Office phone",
  email: "Email (password reset only)",
  securityquestion: "Security questions",
  microsoftauthenticatorpush: "Authenticator app",
  microsoftauthenticatorpasswordless: "Authenticator passwordless",
  softwareonetimepasscode: "Authenticator code",
  hardwareonetimepasscode: "Hardware token",
  fido2securitykey: "Security key (FIDO2)",
  windowshelloforbusiness: "Windows Hello for Business",
  temporaryaccesspass: "Temporary Access Pass",
  passkeydevicebound: "Passkey",
  passkeydeviceboundauthenticator: "Passkey in Authenticator",
  passkeydeviceboundwindowshello: "Passkey in Windows Hello",
  passkeysynced: "Synced passkey",
  macossecureenclavekey: "Platform credential for macOS",
  externalauthenticationmethod: "External method",
};

const preferredLabels: Record<string, string> = {
  push: "Authenticator app",
  oath: "Authenticator or token code",
  voicemobile: "Call to mobile",
  voicealternatemobile: "Call to alternate mobile",
  voiceoffice: "Call to office phone",
  sms: "Text message",
  fido2: "Security key (FIDO2)",
  passkey: "Passkey",
  temporaryaccesspass: "Temporary Access Pass",
  email: "Email",
};

// Method families, strongest first. Email, security questions and Temporary Access Pass are not
// counted: they are password-reset or one-off methods, not a lasting second factor.
const passwordless = new Set([
  "microsoftauthenticatorpasswordless",
  "fido2securitykey",
  "windowshelloforbusiness",
  "passkeydevicebound",
  "passkeydeviceboundauthenticator",
  "passkeydeviceboundwindowshello",
  "passkeysynced",
  "macossecureenclavekey",
]);
const appOrToken = new Set([
  "microsoftauthenticatorpush",
  "softwareonetimepasscode",
  "hardwareonetimepasscode",
  "externalauthenticationmethod",
]);
const phone = new Set(["mobilephone", "alternatemobilephone", "officephone"]);

export type MethodStrength = "Passwordless" | "App or token" | "Phone or text" | "None";

export function strongestMethod(methods: readonly string[], passwordlessCapable: boolean | null) {
  const keys = methods.map((method) => method.toLowerCase());
  if (passwordlessCapable || keys.some((key) => passwordless.has(key))) return "Passwordless";
  if (keys.some((key) => appOrToken.has(key))) return "App or token";
  if (keys.some((key) => phone.has(key))) return "Phone or text";
  return "None";
}

function humanize(value: string): string {
  const spaced = value.replace(/([a-z0-9])([A-Z])/g, "$1 $2").toLowerCase();
  return `${spaced.charAt(0).toUpperCase()}${spaced.slice(1)}`;
}

export function methodLabel(method: string): string {
  return methodLabels[method.toLowerCase()] ?? humanize(method);
}

function preferred(row: z.output<typeof registrationSchema>): string | null {
  const system = row.systemPreferredAuthenticationMethods[0];
  if (row.isSystemPreferredAuthenticationMethodEnabled && system)
    return `${preferredLabels[system.toLowerCase()] ?? humanize(system)} (chosen by Microsoft)`;
  const own = row.userPreferredMethodForSecondaryAuthentication;
  if (!own || ["none", "unknownfuturevalue"].includes(own.toLowerCase())) return null;
  return preferredLabels[own.toLowerCase()] ?? humanize(own);
}

export const mfaRegistrationAudit: Audit = {
  definition: {
    id: "mfa.registration",
    version: 3,
    title: "MFA registration",
    description:
      "Who has registered a strong sign-in method and who has not, with the admins who rely on phone or text codes. From Microsoft's authentication methods report.",
    category: "Security",
    subject: { one: "account", many: "accounts" },
    columns: [
      { key: "person", label: "Account", kind: "person" },
      { key: "admin", label: "Admin", kind: "boolean" },
      {
        key: "registered",
        label: "MFA registered",
        kind: "boolean",
        help: "Microsoft's definition: at least one strong method your authentication methods policy allows is registered.",
      },
      { key: "strength", label: "Strongest method", kind: "enum", facet: true },
      { key: "methods", label: "Methods", kind: "list" },
      { key: "preferred", label: "Default method", kind: "text", empty: "Not set" },
      { key: "type", label: "Type", kind: "enum", facet: true },
      domainColumn,
      { key: "capable", label: "MFA capable", kind: "boolean", hidden: true },
      { key: "passwordless", label: "Passwordless capable", kind: "boolean", hidden: true },
      { key: "sspr", label: "Password reset registered", kind: "boolean", hidden: true },
      { key: "updated", label: "Report updated", kind: "date", hidden: true, empty: "Unknown" },
    ],
    views: [
      { id: "not-registered", label: "Not registered", tags: ["no-mfa"] },
      { id: "admins-no-mfa", label: "Admins without MFA", tags: ["admin", "no-mfa"] },
      { id: "admins", label: "Admins", tags: ["admin"] },
      { id: "phone-only", label: "Phone or text only", tags: ["phone-only"] },
      { id: "passwordless", label: "Passwordless", tags: ["passwordless"] },
      { id: "all", label: "Everyone", tags: [] },
    ],
    defaultView: "not-registered",
    defaultSort: { column: "admin", direction: "desc" },
    requirements: [
      "A role such as Reports Reader, Security Reader, Security Administrator or Global Reader.",
    ],
  },
  packs: ["reports.authentication-methods.read"],
  optionalPacks: [],
  scoped: true,
  forbidden:
    "Microsoft refused the authentication methods report. It needs AuditLog.Read.All and a role such as Reports Reader, Security Reader, Security Administrator or Global Reader.",
  async run(context) {
    context.signal.throwIfAborted();
    const result = await context.graph.getCollection(
      "reports/authenticationMethods/userRegistrationDetails",
      reportContext,
      { schema: z.unknown(), maxPages: 100, maxRecords: MAX_ACCOUNTS },
    );
    let skipped = 0;
    let outside = 0;
    let outsideAdmins = 0;
    const principals: (string | null)[] = [];
    const rows: AuditRow[] = [];
    for (const raw of result.items) {
      const parsed = registrationSchema.safeParse(raw);
      if (!parsed.success) {
        skipped += 1;
        continue;
      }
      const row = parsed.data;
      principals.push(row.userPrincipalName);
      if (!inAuditScope(context.scope, row.userPrincipalName)) {
        outside += 1;
        // An admin without MFA matters whatever its domain (an emergency-access account often
        // sits on the tenant's onmicrosoft.com domain), so the scope never hides one silently.
        if (row.isAdmin === true && row.isMfaRegistered !== true) outsideAdmins += 1;
        continue;
      }
      const strength = strongestMethod(row.methodsRegistered, row.isPasswordlessCapable);
      const registered = row.isMfaRegistered === true;
      const admin = row.isAdmin === true;
      const type = typeOf(row.userType);
      rows.push({
        key: row.id,
        personId: row.id,
        tags: [
          ...(registered ? [] : ["no-mfa"]),
          ...(admin ? ["admin"] : []),
          ...(registered && strength === "Phone or text" ? ["phone-only"] : []),
          ...(strength === "Passwordless" ? ["passwordless"] : []),
          ...tagOf(type),
        ],
        cells: {
          person: { name: row.userDisplayName, principal: row.userPrincipalName },
          admin,
          registered,
          strength,
          methods: row.methodsRegistered.map(methodLabel),
          preferred: preferred(row),
          type,
          domain: principalDomain(row.userPrincipalName),
          capable: row.isMfaCapable,
          passwordless: row.isPasswordlessCapable,
          sspr: row.isSsprRegistered,
          updated: row.lastUpdatedDateTime,
        },
      });
    }
    rows.sort(byName);
    const count = (...tags: string[]) =>
      rows.filter((row) => tags.every((tag) => row.tags.includes(tag))).length;
    const adminsWithout = count("admin", "no-mfa");
    const limited = scopeNote(context.scope);
    const left = excludedNote(outside);
    return {
      domains: domainInventory(principals),
      rows,
      complete: result.complete && skipped === 0,
      summary: [
        { label: "Accounts", value: rows.length, tone: "neutral", view: "all" },
        {
          label: "Not registered",
          value: count("no-mfa"),
          tone: "warning",
          view: "not-registered",
        },
        {
          label: "Admins without MFA",
          value: adminsWithout,
          // Never "good" under a scope: a clean count there says nothing about the other domains.
          tone: adminsWithout > 0 ? "critical" : context.scope ? "neutral" : "good",
          view: "admins-no-mfa",
        },
        ...(outsideAdmins > 0
          ? [
              {
                label: "Admins without MFA outside the scope",
                value: outsideAdmins,
                tone: "critical" as const,
                view: null,
              },
            ]
          : []),
        {
          label: "Phone or text only",
          value: count("phone-only"),
          tone: "warning",
          view: "phone-only",
        },
        { label: "Passwordless", value: count("passwordless"), tone: "good", view: "passwordless" },
      ],
      notes: [
        ...(limited ? [limited] : []),
        ...(left ? [left] : []),
        ...(outsideAdmins > 0
          ? [
              `${outsideAdmins === 1 ? "1 admin" : `${outsideAdmins} admins`} without MFA ${outsideAdmins === 1 ? "is" : "are"} in domains the audit scope leaves out. Choose All domains and run again to see ${outsideAdmins === 1 ? "it" : "them"}.`,
            ]
          : []),
        "Microsoft's report leaves out disabled accounts.",
        "Registered is not the same as required. Conditional Access or security defaults decide whether a sign-in asks for MFA.",
        "Phone or text only: the strongest method registered is a text message or call, which phishing and SIM swaps can defeat. Move these accounts, admins first, to the Authenticator app or a passkey.",
        "Microsoft refreshes this report on its own schedule. Report updated shows when each row was refreshed.",
        ...coverageNotes({ skipped, complete: result.complete }),
      ],
    };
  },
};

import { relative } from "node:path";
import {
  type CertificateFiles,
  certificateKey,
  type LocalSecrets,
  restrictsToOwner,
} from "./certificate.ts";
import {
  adminObjectIds,
  type EnvChange,
  type ManagedValues,
  mergeEnv,
  readEnv,
} from "./env-file.ts";
import { type GraphApplication, type GraphServicePrincipal, terminalText } from "./graph.ts";
import { hasCertificate, unreadableKeys } from "./key-credentials.ts";
import type { SetupOptions } from "./options.ts";
import {
  GRAPH_RESOURCE_APP_ID,
  graphResourceAccess,
  mergeRequiredResourceAccess,
  type RequestedPermission,
  type RequiredResourceAccess,
  requestedPermissions,
} from "./permissions.ts";

// The fixed tag that marks the registration this script made, so a re-run finds it again.
export const SETUP_TAG = "prism365-setup";
const DAY = 86_400_000;

// Everything the setup read before deciding, gathered by setup.ts.
export interface SetupState {
  options: SetupOptions;
  tenantId: string;
  admin: { id: string; userPrincipalName: string | null };
  application: GraphApplication | null;
  foundBy: "tag" | "name" | "appId" | null;
  // How many registrations matched the search; more than one stops the setup.
  matches: number;
  servicePrincipal: boolean;
  graph: GraphServicePrincipal;
  // The existing registration's certificates, read with their keys; empty for a new one.
  keyCredentials: unknown[];
  files: CertificateFiles;
  secrets: LocalSecrets;
  env: { exists: boolean; text: string; source: "env" | "example" | "empty" };
  cwd: string;
  now: Date;
}

export type CertificateStep =
  | { action: "blocked" }
  | { action: "create" }
  | { action: "rotate" }
  | { action: "certify" }
  | { action: "upload"; key: string; notAfter: Date }
  | { action: "keep"; notAfter: Date };

export type ApplicationStep =
  | { action: "create"; body: Record<string, unknown> }
  | {
      action: "update";
      id: string;
      appId: string;
      patch: Record<string, unknown>;
      addedRedirectUris: string[];
      addedPermissions: string[];
      removedPermissions: string[];
      addTag: boolean;
    }
  | { action: "keep"; id: string; appId: string };

export interface SetupPlan {
  tenantId: string;
  admin: SetupState["admin"];
  displayName: string;
  foundBy: SetupState["foundBy"];
  application: ApplicationStep;
  redirectUris: string[];
  permissions: RequestedPermission[];
  servicePrincipal: "create" | "keep";
  certificate: CertificateStep;
  restrictDirectory: boolean;
  restrictPrivateKey: boolean;
  env: { source: SetupState["env"]["source"]; changes: EnvChange[]; write: boolean };
  blockers: string[];
  warnings: string[];
}

export function consentLink(tenantId: string, appId: string): string {
  return `https://login.microsoftonline.com/${tenantId}/adminconsent?client_id=${appId}`;
}

const day = (date: Date) => date.toISOString().slice(0, 10);

// The .env values for a registration. `appId` is null before a new registration exists.
export function managedValues(state: SetupState, appId: string | null): ManagedValues {
  const current = readEnv(state.env.text);
  return {
    DATA_MODE: "live",
    ENTRA_TENANT_ID: state.tenantId,
    ENTRA_CLIENT_ID: appId ?? "pending",
    ENTRA_ADMIN_OBJECT_IDS: adminObjectIds(current.get("ENTRA_ADMIN_OBJECT_IDS"), state.admin.id),
    ENABLED_READ_FEATURES: state.options.features.join(","),
    PUBLIC_ORIGIN: state.options.origin,
    ENTRA_CERTIFICATE_PATH: state.files.certificate,
    ENTRA_PRIVATE_KEY_PATH: state.files.privateKey,
  };
}

function permissionName(
  graph: GraphServicePrincipal,
  entry: { resourceAppId: string; access: { id: string; type: string } },
): string {
  if (entry.resourceAppId.toLowerCase() !== GRAPH_RESOURCE_APP_ID)
    return `a permission of another API (${terminalText(entry.resourceAppId, 60)})`;
  const list = entry.access.type === "Role" ? (graph.appRoles ?? []) : graph.oauth2PermissionScopes;
  const found = list.find((item) => item.id.toLowerCase() === entry.access.id.toLowerCase());
  const name = found?.value ? terminalText(found.value, 100) : `permission ${entry.access.id}`;
  return entry.access.type === "Role" ? `${name} (application permission)` : name;
}

function applicationStep(state: SetupState, wanted: RequiredResourceAccess): ApplicationStep {
  const { options } = state;
  if (!state.application)
    return {
      action: "create",
      body: {
        displayName: options.name,
        signInAudience: "AzureADMyOrg",
        tags: [SETUP_TAG],
        web: {
          redirectUris: options.redirectUris,
          implicitGrantSettings: { enableIdTokenIssuance: false, enableAccessTokenIssuance: false },
        },
        requiredResourceAccess: [wanted],
      },
    };
  const application = state.application;
  const redirectUris = application.web?.redirectUris ?? [];
  const addedRedirectUris = options.redirectUris.filter((uri) => !redirectUris.includes(uri));
  const permissions = mergeRequiredResourceAccess(
    application.requiredResourceAccess ?? [],
    wanted,
    options.prunePermissions,
  );
  const tags = application.tags ?? [];
  const addTag = !tags.includes(SETUP_TAG);
  const patch: Record<string, unknown> = {};
  if (addedRedirectUris.length > 0) {
    // The rest of `web` goes back as it was. redirectUriSettings is left out: it pairs settings
    // with redirect URIs by position (SAML only), which adding a URI would shift.
    const { redirectUriSettings: _positional, ...web } = application.web ?? {};
    patch.web = { ...web, redirectUris: [...redirectUris, ...addedRedirectUris] };
  }
  if (permissions.added.length > 0 || permissions.removed.length > 0)
    patch.requiredResourceAccess = permissions.value;
  if (addTag) patch.tags = [...tags, SETUP_TAG];
  if (Object.keys(patch).length === 0)
    return { action: "keep", id: application.id, appId: application.appId };
  const scopes = state.graph.oauth2PermissionScopes;
  return {
    action: "update",
    id: application.id,
    appId: application.appId,
    patch,
    addedRedirectUris,
    addedPermissions: permissions.added.map(
      (access) => scopes.find((scope) => scope.id === access.id)?.value ?? access.id,
    ),
    removedPermissions: permissions.removed.map((entry) => permissionName(state.graph, entry)),
    addTag,
  };
}

function certificateStep(
  state: SetupState,
  blockers: string[],
  warnings: string[],
): CertificateStep {
  const { credential } = state.secrets;
  const shown = (path: string) => relative(state.cwd, path) || path;
  if (state.options.rotateCertificate)
    return credential.state === "none" ? { action: "create" } : { action: "rotate" };
  switch (credential.state) {
    case "none":
      return { action: "create" };
    case "keyOnly":
      return { action: "certify" };
    case "pair": {
      const notAfter = credential.certificate.validToDate;
      if (notAfter.getTime() <= state.now.getTime()) {
        blockers.push(
          `The certificate in ${shown(state.files.certificate)} expired on ${day(notAfter)}. Run the setup again with --rotate-certificate to replace it.`,
        );
        return { action: "blocked" };
      }
      if (notAfter.getTime() - state.now.getTime() < 30 * DAY)
        warnings.push(
          `The certificate in ${shown(state.files.certificate)} expires on ${day(notAfter)}. Run the setup with --rotate-certificate before then.`,
        );
      const key = certificateKey(credential.certificate);
      return hasCertificate(state.keyCredentials, key)
        ? { action: "keep", notAfter }
        : { action: "upload", key, notAfter };
    }
    case "certificateOnly":
      blockers.push(
        `${shown(state.files.certificate)} is there but its private key ${shown(state.files.privateKey)} is not, so Prism365 could not sign in with it. Restore the key, or run the setup again with --rotate-certificate to make a new pair.`,
      );
      return { action: "blocked" };
    case "mismatch":
      blockers.push(
        `${shown(state.files.privateKey)} and ${shown(state.files.certificate)} do not belong together. Restore the matching pair, or run the setup again with --rotate-certificate to replace both.`,
      );
      return { action: "blocked" };
    case "unreadable":
      blockers.push(
        `${shown(credential.file === "privateKey" ? state.files.privateKey : state.files.certificate)} is not a PEM ${credential.file === "privateKey" ? "private key" : "certificate"} the setup can read. Restore it, or run the setup again with --rotate-certificate to replace the pair.`,
      );
      return { action: "blocked" };
  }
}

export function buildPlan(state: SetupState): SetupPlan {
  const { options } = state;
  const blockers: string[] = [];
  const warnings: string[] = [];
  if (state.matches > 1)
    blockers.push(
      `${state.matches} app registrations match (by the ${SETUP_TAG} tag or the name "${options.name}"). Run the setup again with --app-id and the application ID of the one to use.`,
    );
  if (options.appId && !state.application)
    blockers.push(
      `No app registration with application ID ${options.appId} was found in this tenant.`,
    );
  const permissions = requestedPermissions(options.features);
  const wanted = graphResourceAccess(
    permissions.map((permission) => permission.name),
    state.graph.oauth2PermissionScopes,
  );
  const application = applicationStep(state, wanted);
  const certificate = certificateStep(state, blockers, warnings);
  if (certificate.action !== "keep" && certificate.action !== "blocked") {
    const unreadable = unreadableKeys(state.keyCredentials);
    if (unreadable.length > 0)
      blockers.push(
        `Microsoft did not return the full key of ${unreadable.join(", ")} on this registration, and adding a certificate means sending every existing one back. The setup stops rather than risk deleting one. Add the certificate in the Entra admin center instead (docs/SETUP.md).`,
      );
  }
  const credential = state.secrets.credential;
  const restrictPrivateKey =
    (credential.state === "pair" || credential.state === "keyOnly") &&
    certificate.action !== "rotate" &&
    !restrictsToOwner(credential.privateKeyMode);
  const appId = application.action === "create" ? null : application.appId;
  const merged = mergeEnv(state.env.text, managedValues(state, appId));
  const env = {
    source: state.env.source,
    changes: merged.changes,
    write: !state.env.exists || merged.text !== state.env.text,
  };

  if (state.application && state.application.signInAudience !== "AzureADMyOrg")
    warnings.push(
      `The registration accepts sign-ins from ${terminalText(state.application.signInAudience ?? "an unknown audience", 60)} accounts. Prism365 checks the tenant itself, but a single-tenant registration (AzureADMyOrg) is safer; change it under Authentication in the Entra admin center.`,
    );
  const implicit = state.application?.web?.implicitGrantSettings;
  if (implicit?.enableAccessTokenIssuance || implicit?.enableIdTokenIssuance)
    warnings.push(
      "The registration has implicit grant turned on. Prism365 does not use it; turn it off under Authentication in the Entra admin center.",
    );
  if (!options.features.includes("directory.users.read"))
    warnings.push(
      "The feature list leaves out directory.users.read, which the Sign in with Microsoft button on the Microsoft access page uses.",
    );
  const current = readEnv(state.env.text);
  const dropped = (current.get("ENABLED_READ_FEATURES") ?? "")
    .split(",")
    .map((item) => item.trim())
    .filter((item) => item && !(options.features as string[]).includes(item));
  if (state.env.exists && dropped.length > 0)
    warnings.push(
      `.env enables ${dropped.join(", ")} now and the new feature list leaves ${dropped.length === 1 ? "it" : "them"} out. Pass --features with the full list to keep ${dropped.length === 1 ? "it" : "them"}.`,
    );
  if (current.get("ENTRA_CLIENT_SECRET"))
    warnings.push(
      "ENTRA_CLIENT_SECRET is set in .env. The certificate takes its place; once sign-in works, remove the secret from .env and from the registration.",
    );

  return {
    tenantId: state.tenantId,
    admin: state.admin,
    displayName:
      application.action === "create"
        ? options.name
        : terminalText(state.application?.displayName ?? options.name, 120),
    foundBy: state.foundBy,
    application,
    redirectUris:
      application.action === "update" && application.patch.web
        ? ((application.patch.web as { redirectUris: string[] }).redirectUris ?? [])
        : application.action === "create"
          ? options.redirectUris
          : (state.application?.web?.redirectUris ?? []),
    permissions,
    servicePrincipal:
      application.action === "create" || !state.servicePrincipal ? "create" : "keep",
    certificate,
    restrictDirectory:
      state.secrets.directoryMode !== null && !restrictsToOwner(state.secrets.directoryMode),
    restrictPrivateKey,
    env,
    blockers,
    warnings,
  };
}

function list(items: readonly string[]): string {
  if (items.length <= 1) return items.join("");
  return `${items.slice(0, -1).join(", ")} and ${items.at(-1)}`;
}

// The plan in plain sentences, printed before anything changes.
export function planText(plan: SetupPlan, state: SetupState): string {
  const shown = (path: string) => relative(state.cwd, path) || path;
  const lines: string[] = [];
  const tenant =
    state.options.tenant === plan.tenantId
      ? plan.tenantId
      : `${state.options.tenant} (${plan.tenantId})`;
  lines.push(`Setup plan for tenant ${tenant}.`, "");
  lines.push(
    `Signed in as ${plan.admin.userPrincipalName ? terminalText(plan.admin.userPrincipalName, 200) : "your account"} (object ID ${plan.admin.id}).`,
  );
  lines.push("");

  const app = plan.application;
  const how = { tag: `its ${SETUP_TAG} tag`, name: "its name", appId: "the --app-id option" };
  if (app.action === "create")
    lines.push(
      `App registration: create "${plan.displayName}", a single-tenant web app with implicit grant off, tagged ${SETUP_TAG}.`,
    );
  else if (app.action === "keep")
    lines.push(
      `App registration: keep "${plan.displayName}" (application ID ${app.appId}, found by ${how[plan.foundBy ?? "tag"]}). It already has the redirect URIs and permissions below.`,
    );
  else {
    const changes = [
      ...(app.addedRedirectUris.length > 0
        ? [
            `add the redirect ${app.addedRedirectUris.length === 1 ? "URI" : "URIs"} ${list(app.addedRedirectUris)}`,
          ]
        : []),
      ...(app.addedPermissions.length > 0
        ? [
            `add the ${app.addedPermissions.length === 1 ? "permission" : "permissions"} ${list(app.addedPermissions)}`,
          ]
        : []),
      ...(app.removedPermissions.length > 0
        ? [`remove ${list(app.removedPermissions)} (--prune-permissions)`]
        : []),
      ...(app.addTag ? [`tag it ${SETUP_TAG} so the next run finds it`] : []),
    ];
    lines.push(
      `App registration: update "${plan.displayName}" (application ID ${app.appId}, found by ${how[plan.foundBy ?? "tag"]}): ${list(changes)}. Nothing else on it changes.`,
    );
  }
  lines.push(`Redirect URIs: ${list(plan.redirectUris.map((uri) => terminalText(uri, 300)))}.`);
  lines.push(
    "Delegated Microsoft Graph permissions it asks for (an administrator grants them with one consent click afterwards):",
  );
  const signIn = plan.permissions.filter((permission) => permission.features.length === 0);
  lines.push(`  ${list(signIn.map((permission) => permission.name))} for sign-in`);
  for (const permission of plan.permissions.filter((item) => item.features.length > 0))
    lines.push(`  ${permission.name} for ${list(permission.features)}`);
  lines.push(
    plan.servicePrincipal === "create"
      ? "Service principal: create the registration's service principal (its enterprise application)."
      : "Service principal: the registration already has one.",
  );
  lines.push("");

  const certificate = plan.certificate;
  const where = `${shown(state.files.privateKey)} (owner only) and ${shown(state.files.certificate)}`;
  const until = new Date(state.now.getTime() + 365 * DAY);
  if (certificate.action === "create")
    lines.push(
      `Certificate: create an RSA 2048 key and a self-signed certificate valid until ${day(until)}, write ${where}, and add the certificate to the registration. Certificates already on the registration stay.`,
    );
  else if (certificate.action === "rotate")
    lines.push(
      `Certificate: replace ${where} with a new key and certificate valid until ${day(until)}, and add it to the registration. The old certificate stays on the registration until you remove it in the Entra admin center.`,
    );
  else if (certificate.action === "certify")
    lines.push(
      `Certificate: ${shown(state.files.privateKey)} has no certificate yet. Make a self-signed certificate for that key, valid until ${day(until)}, write ${shown(state.files.certificate)} and add it to the registration. The key is not changed.`,
    );
  else if (certificate.action === "blocked")
    lines.push("Certificate: the setup cannot use the files in the secrets directory (see below).");
  else if (certificate.action === "upload")
    lines.push(
      `Certificate: add the existing ${shown(state.files.certificate)} (valid until ${day(certificate.notAfter)}) to the registration. Certificates already on it stay.`,
    );
  else
    lines.push(
      `Certificate: keep ${shown(state.files.certificate)} (valid until ${day(certificate.notAfter)}), which is already on the registration.`,
    );
  if (plan.restrictDirectory)
    lines.push(`Restrict ${shown(state.files.directory)} so only you can open it (mode 700).`);
  if (plan.restrictPrivateKey)
    lines.push(`Restrict ${shown(state.files.privateKey)} so only you can read it (mode 600).`);
  lines.push("");

  const envFile = shown(state.options.envFile);
  const describe = (change: EnvChange) =>
    change.key === "ENTRA_CLIENT_ID" && app.action === "create"
      ? "ENTRA_CLIENT_ID to the new registration's application ID"
      : change.before === null
        ? `${change.key} to ${change.after}`
        : `${change.key} from ${change.before} to ${change.after}`;
  const settings = plan.env.changes.map((change) => `  ${describe(change)}`);
  if (!plan.env.write) lines.push(`${envFile}: already has these settings.`);
  else if (plan.env.source !== "env")
    lines.push(
      `${envFile}: create it${plan.env.source === "example" ? " from .env.example" : ""}, readable only by you, and set:`,
      ...settings,
    );
  else if (plan.env.changes.length === 0)
    lines.push(`${envFile}: tidy duplicate lines of settings it manages. Every other line stays.`);
  else
    lines.push(
      `${envFile}: change these settings, keep every other line, and make the file readable only by you:`,
      ...settings,
    );

  for (const warning of plan.warnings) lines.push("", `Note: ${warning}`);
  if (plan.blockers.length > 0) {
    lines.push("", "The setup cannot continue:");
    for (const blocker of plan.blockers) lines.push(`  ${blocker}`);
  }
  lines.push(
    "",
    "Nothing here grants consent: that stays a click in your browser afterwards. Prism365 itself only reads.",
  );
  return lines.join("\n");
}

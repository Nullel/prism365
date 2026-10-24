import { type KeyObject, X509Certificate } from "node:crypto";
import { relative } from "node:path";
import { supportedReadFeatures } from "../../apps/api/src/config.ts";
import { type AccessList, modeDecision } from "./acl.ts";
import {
  type CertificateFiles,
  certificateKey,
  interruptedRotation,
  keyProblem,
  type LocalSecrets,
  ownerModes,
  type SetupRecord,
  thumbprintOf,
  type WriteModes,
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
  SIGN_IN_PERMISSIONS,
} from "./permissions.ts";

// A label on registrations this setup creates. It is only a label: the setup never looks a
// registration up by it, because anyone allowed to register apps can put it on theirs. The
// registration to work on is the one named by --app-id, or the one this install's .env uses.
export const SETUP_TAG = "prism365-setup";
// A second label on a create, unique to the run, so an answer Microsoft loses can be followed up
// by looking for exactly that registration. It is removed when the run finishes.
export const RUN_TAG_PREFIX = "prism365-setup-run-";
const DAY = 86_400_000;

export interface Owner {
  id: string;
  name: string | null;
}

// What the setup learned about an existing registration beyond its own properties.
export interface RegistrationFacts {
  // Its owners, and whether the list was read to the end.
  owners: Owner[];
  ownersComplete: boolean;
  // Federated identity credentials: another identity provider's tokens accepted as the app.
  federatedCredentials: number;
  // Its service principal (enterprise application), or null when it has none yet.
  servicePrincipal: {
    id: string;
    owners: Owner[];
    ownersComplete: boolean;
    passwordCredentials: number;
    keyCredentials: unknown[];
    // Application permissions granted to it.
    appRoleAssignments: { resourceDisplayName?: string | null | undefined }[];
    // Delegated consent already granted to it (oauth2PermissionGrants).
    grants: {
      consentType?: string | null | undefined;
      resourceId: string;
      scope?: string | null | undefined;
    }[];
  } | null;
}

// Everything the setup read before deciding, gathered by setup.ts.
export interface SetupState {
  options: SetupOptions;
  tenantId: string;
  admin: { id: string; userPrincipalName: string | null };
  // The registration to work on: named by --app-id, or by ENTRA_CLIENT_ID in .env. Null means
  // a new one is created, unless a named one is missing or invalid (missingTarget).
  application: GraphApplication | null;
  foundBy: "appId" | "env" | null;
  missingTarget: { appId: string; source: "appId" | "env" } | null;
  invalidEnvClientId: string | null;
  registration: RegistrationFacts | null;
  servicePrincipal: boolean;
  graph: GraphServicePrincipal;
  // The existing registration's certificates, read with their keys; empty for a new one.
  keyCredentials: unknown[];
  files: CertificateFiles;
  secrets: LocalSecrets;
  record: SetupRecord;
  // Access control lists, where they can be read (Linux with getfacl); null elsewhere.
  acl: {
    directory: AccessList | null;
    privateKey: AccessList | null;
    certificate: AccessList | null;
  };
  env: { exists: boolean; text: string; source: "env" | "example" | "empty" };
  cwd: string;
  now: Date;
  platform: NodeJS.Platform;
}

export type CertificateStep =
  | { action: "blocked" }
  | { action: "create" }
  | { action: "rotate" }
  | { action: "promote" }
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
      // Run labels a run that stopped halfway left on the registration.
      removedLabels: string[];
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
  // What the setup verified about an existing registration, for the plan to state.
  verified: string[];
  servicePrincipal: "create" | "keep";
  certificate: CertificateStep;
  writeModes: WriteModes;
  // New permission bits for the secrets directory and the key, or null to leave them.
  restrictDirectory: number | null;
  restrictPrivateKey: number | null;
  // ACL entries a rotation could not give the files last time, to apply now.
  reapplyAcl: { privateKey: string[]; certificate: string[] } | null;
  env: { source: SetupState["env"]["source"]; changes: EnvChange[]; write: boolean };
  blockers: string[];
  warnings: string[];
}

export function consentLink(tenantId: string, appId: string): string {
  return `https://login.microsoftonline.com/${tenantId}/adminconsent?client_id=${appId}`;
}

const day = (date: Date) => date.toISOString().slice(0, 10);
const sameId = (left: string, right: string) => left.toLowerCase() === right.toLowerCase();

function list(items: readonly string[]): string {
  if (items.length <= 1) return items.join("");
  return `${items.slice(0, -1).join(", ")} and ${items.at(-1)}`;
}

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
  if (!sameId(entry.resourceAppId, GRAPH_RESOURCE_APP_ID))
    return `a ${entry.access.type === "Role" ? "application" : "delegated"} permission of another API (${terminalText(entry.resourceAppId, 60)})`;
  const published =
    entry.access.type === "Role" ? (graph.appRoles ?? []) : graph.oauth2PermissionScopes;
  const found = published.find((item) => sameId(item.id, entry.access.id));
  const name = found?.value ? terminalText(found.value, 100) : `permission ${entry.access.id}`;
  return entry.access.type === "Role" ? `${name} (application permission)` : name;
}

// Whether a certificate on the registration is one this setup manages: one for the working key or
// a rotation's new key (whichever certificate file that key has now), or the certificate the last
// rotation replaced. Any other certificate was put there by someone else.
function managedBy(secrets: LocalSecrets): (certificate: X509Certificate) => boolean {
  const { credential } = secrets;
  const keys: KeyObject[] = [];
  const certificates: Uint8Array[] = [];
  if (
    credential.state === "pair" ||
    credential.state === "keyOnly" ||
    credential.state === "mismatch"
  )
    keys.push(credential.privateKey);
  if (credential.state === "pair" || credential.state === "mismatch")
    certificates.push(credential.certificate.raw);
  if (secrets.pending) {
    certificates.push(secrets.pending.certificate.raw);
    if (secrets.pending.privateKey) keys.push(secrets.pending.privateKey);
  }
  if (secrets.previous) certificates.push(secrets.previous.raw);
  return (certificate) =>
    keys.some((key) => certificate.checkPrivateKey(key)) ||
    certificates.some((raw) => Buffer.compare(raw, certificate.raw) === 0);
}

interface ForeignCertificate {
  label: string;
  // The certificate the setup's last rotation replaced, once its .previous files are gone.
  replaced: boolean;
}

function foreignCertificates(
  state: SetupState,
  credentials: readonly unknown[],
): ForeignCertificate[] {
  const managed = managedBy(state.secrets);
  return credentials
    .map((item) => item as { keyId?: unknown; key?: unknown; displayName?: unknown })
    .flatMap((credential) => {
      let certificate: X509Certificate | null = null;
      if (typeof credential.key === "string" && credential.key.length > 0)
        try {
          certificate = new X509Certificate(Buffer.from(credential.key, "base64"));
        } catch {
          certificate = null;
        }
      if (certificate && managed(certificate)) return [];
      const label =
        typeof credential.displayName === "string" && credential.displayName
          ? `${terminalText(credential.displayName, 90)} (${terminalText(String(credential.keyId), 40)})`
          : terminalText(String(credential.keyId), 40);
      return [
        {
          label,
          replaced:
            certificate !== null &&
            state.record.replacedThumbprint !== null &&
            thumbprintOf(certificate) === state.record.replacedThumbprint,
        },
      ];
    });
}

const ownerLabel = (owner: Owner) =>
  owner.name ? `${terminalText(owner.name, 80)} (${owner.id})` : owner.id;

// Every delegated permission Prism365 could ask for: sign-in, and every pack's permission.
const prism365Permissions = new Set<string>([
  ...SIGN_IN_PERMISSIONS,
  ...Object.values(supportedReadFeatures),
]);

// Everything that makes an existing registration unsafe to adopt, as blockers (the setup stops
// before any write) and, for what the administrator chose to keep, warnings. A registration is
// adopted only when the signed-in administrator is its only owner (or chose to keep the others)
// and it carries nothing that lets someone else sign in as Prism365, or that the consent link, or
// consent already given, grants beyond Prism365's own permissions. These checks are a snapshot:
// its owners can change it afterwards, and the plan says so.
function adoptionProblems(
  state: SetupState,
  wanted: RequiredResourceAccess,
  blockers: string[],
  warnings: string[],
): string[] {
  const application = state.application;
  if (!application) return [];
  const { options } = state;
  const name = terminalText(application.displayName ?? "(no name)", 120);
  const which = `"${name}" (application ID ${application.appId})`;
  const verified: string[] = [];
  const facts = state.registration;
  const principal = facts?.servicePrincipal ?? null;

  const owners = facts?.owners ?? [];
  const isMine = (owner: Owner) => sameId(owner.id, state.admin.id);
  if (!owners.some(isMine) || !facts?.ownersComplete)
    blockers.push(
      `You are not an owner of ${which}. The setup only works on a registration you own. Ask one of its owners to add you under Owners in the Entra admin center, or run the setup with --app-id for another registration, or with neither --app-id nor ENTRA_CLIENT_ID in .env to create a new one.`,
    );
  const others = [
    ...owners
      .filter((owner) => !isMine(owner))
      .map((owner) => `${ownerLabel(owner)} of the registration`),
    ...(principal?.owners ?? [])
      .filter((owner) => !isMine(owner))
      .map((owner) => `${ownerLabel(owner)} of its service principal`),
  ];
  if (principal && !principal.ownersComplete)
    others.push("more owners of its service principal than the setup could read");
  if (others.length > 0 && options.keepOtherOwners)
    warnings.push(
      `${which} has other owners: ${list(others)}. You chose to keep them. They can change the registration at any time, before or after this setup: add a secret, a redirect URI or permissions. The checks in this plan only describe it as it is now.`,
    );
  else if (others.length > 0)
    blockers.push(
      `${which} has owners besides you: ${list(others)}. Any owner can change the registration at any time, for example add a secret or a redirect URI that sends sign-ins elsewhere. Remove them under Owners in the Entra admin center (on the app registration and on the enterprise application), or run the setup again with --keep-other-owners if you trust them.`,
    );
  else if (owners.some(isMine)) verified.push("you are its only owner");

  const secrets =
    (application.passwordCredentials?.length ?? 0) + (principal?.passwordCredentials ?? 0);
  const federated = facts?.federatedCredentials ?? 0;
  if (secrets > 0)
    blockers.push(
      `${which} or its service principal has ${secrets === 1 ? "a client secret" : `${secrets} client secrets`}. Prism365 signs in only with its certificate, and a secret would let whoever holds it sign in as Prism365. Delete ${secrets === 1 ? "it" : "them"} under Certificates and secrets in the Entra admin center (and on the enterprise application), then run the setup again.`,
    );
  if (federated > 0)
    blockers.push(
      `${which} has ${federated === 1 ? "a federated credential" : `${federated} federated credentials`}: tokens from another identity provider that Microsoft accepts as Prism365. Delete ${federated === 1 ? "it" : "them"} under Certificates and secrets, Federated credentials, then run the setup again.`,
    );
  if (secrets === 0 && federated === 0)
    verified.push("it and its service principal have no client secrets or federated credentials");

  const publicUris = application.publicClient?.redirectUris ?? [];
  const spaUris = application.spa?.redirectUris ?? [];
  if (publicUris.length > 0 || spaUris.length > 0 || application.isFallbackPublicClient)
    blockers.push(
      `${which} also allows sign-in from ${list([
        ...(publicUris.length > 0 ? ["mobile and desktop applications"] : []),
        ...(spaUris.length > 0 ? ["single-page applications"] : []),
        ...(application.isFallbackPublicClient ? ["public client flows"] : []),
      ])}. Prism365 signs in only through its own server. Remove those platforms and turn off Allow public client flows under Authentication in the Entra admin center, then run the setup again.`,
    );

  const requested = application.requiredResourceAccess ?? [];
  const roles = requested.flatMap((entry) =>
    entry.resourceAccess
      .filter((access) => access.type === "Role")
      .map((access) => permissionName(state.graph, { resourceAppId: entry.resourceAppId, access })),
  );
  if (roles.length > 0)
    blockers.push(
      `${which} asks for application permissions: ${list(roles)}. Prism365 never uses application permissions, and the admin consent link would grant them. Remove them under API permissions in the Entra admin center, then run the setup again.`,
    );
  const assigned = principal?.appRoleAssignments ?? [];
  if (assigned.length > 0) {
    const resources = [
      ...new Set(assigned.map((item) => terminalText(item.resourceDisplayName ?? "an API", 80))),
    ];
    blockers.push(
      `${which} has already been granted ${assigned.length === 1 ? "an application permission" : `${assigned.length} application permissions`} (on ${list(resources)}). Prism365 never uses them. Revoke them under Enterprise applications, ${name}, Permissions, Admin consent in the Entra admin center, then run the setup again.`,
    );
  }
  if (roles.length === 0 && assigned.length === 0)
    verified.push("it has no application permissions");

  // Consent already given stays whatever the registration asks for now, so it is checked too.
  const granted = new Set<string>();
  for (const grant of principal?.grants ?? []) {
    if (!sameId(grant.resourceId, state.graph.id)) {
      granted.add("delegated permissions on another API");
      continue;
    }
    for (const scope of (grant.scope ?? "").split(/\s+/).filter(Boolean))
      if (!prism365Permissions.has(scope)) granted.add(terminalText(scope, 100));
  }
  if (granted.size > 0)
    blockers.push(
      `${which} has already been granted consent for ${list([...granted])}, which Prism365 does not use. Removing a permission from the registration does not take back consent already given. Revoke it under Enterprise applications, ${name}, Permissions in the Entra admin center: on the Admin consent tab choose it and Revoke permission; consent one person gave is on the User consent tab, which needs Microsoft Graph (docs/SETUP.md, Revoke the setup's access, shows the request). Then run the setup again.`,
    );
  else verified.push("no consent has been granted beyond Prism365's permissions");

  const wantedIds = wanted.resourceAccess;
  const extra = requested.flatMap((entry) =>
    entry.resourceAccess
      .filter(
        (access) =>
          access.type !== "Role" &&
          !(
            sameId(entry.resourceAppId, GRAPH_RESOURCE_APP_ID) &&
            wantedIds.some((item) => sameId(item.id, access.id) && item.type === access.type)
          ),
      )
      .map((access) => permissionName(state.graph, { resourceAppId: entry.resourceAppId, access })),
  );
  if (extra.length > 0 && options.keepOtherPermissions)
    warnings.push(
      `${which} also asks for ${list(extra)}, which Prism365 does not use. You chose to keep ${extra.length === 1 ? "it" : "them"}, so the admin consent link grants ${extra.length === 1 ? "it" : "them"} too.`,
    );
  else if (extra.length > 0 && !options.prunePermissions)
    blockers.push(
      `${which} asks for delegated permissions Prism365 does not use: ${list(extra)}. The admin consent link would grant them too. Run the setup again with --prune-permissions to remove them, or with --keep-other-permissions to keep them.`,
    );

  const origins = new Set(options.redirectUris.map((uri) => new URL(uri).origin));
  const outside = (application.web?.redirectUris ?? []).filter((uri) => {
    try {
      return !origins.has(new URL(uri).origin);
    } catch {
      return true;
    }
  });
  if (outside.length > 0 && options.keepOtherRedirectUris)
    warnings.push(
      `${which} also sends sign-ins to ${list(outside.map((uri) => terminalText(uri, 200)))}. You chose to keep ${outside.length === 1 ? "it" : "them"}.`,
    );
  else if (outside.length > 0)
    blockers.push(
      `${which} sends sign-ins to addresses outside ${list([...origins])}: ${list(outside.map((uri) => terminalText(uri, 200)))}. Whoever runs those addresses could receive sign-ins meant for Prism365. Remove them under Authentication in the Entra admin center, or run the setup again with --keep-other-redirect-uris if they are yours.`,
    );
  else verified.push("no redirect URIs outside its own address");

  const foreign = [
    ...foreignCertificates(state, state.keyCredentials),
    ...foreignCertificates(state, principal?.keyCredentials ?? []).map((item) => ({
      ...item,
      label: `${item.label} on its service principal`,
    })),
  ];
  const replaced = foreign.filter((item) => item.replaced);
  const unknown = foreign.filter((item) => !item.replaced);
  if (foreign.length > 0 && options.keepOtherCertificates)
    warnings.push(
      `${which} also has certificates this setup did not make: ${list(foreign.map((item) => item.label))}. You chose to keep them; whoever holds their private keys can sign in as Prism365.`,
    );
  else {
    if (replaced.length > 0)
      blockers.push(
        `${list(replaced.map((item) => item.label))} on ${which} is the certificate the setup replaced in its last rotation. Remove it from the registration under Certificates and secrets in the Entra admin center, then run the setup again.`,
      );
    if (unknown.length > 0)
      blockers.push(
        `${which} has certificates this setup did not make: ${list(unknown.map((item) => item.label))}. Whoever holds their private keys can sign in as Prism365. Remove them under Certificates and secrets in the Entra admin center, or run the setup again with --keep-other-certificates if they are yours.`,
      );
  }
  if (foreign.length === 0)
    verified.push(
      "it and its service principal have no certificates other than the ones this setup made",
    );
  return verified;
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
  // Run labels left by a run that stopped halfway are the setup's own; it takes them off.
  const tags = application.tags ?? [];
  const removedLabels = tags.filter((tag) => tag.startsWith(RUN_TAG_PREFIX));
  const patch: Record<string, unknown> = {};
  if (addedRedirectUris.length > 0) {
    // The rest of `web` goes back as it was. redirectUriSettings is left out: it pairs settings
    // with redirect URIs by position (SAML only), which adding a URI would shift.
    const { redirectUriSettings: _positional, ...web } = application.web ?? {};
    patch.web = { ...web, redirectUris: [...redirectUris, ...addedRedirectUris] };
  }
  if (permissions.added.length > 0 || permissions.removed.length > 0)
    patch.requiredResourceAccess = permissions.value;
  if (removedLabels.length > 0) patch.tags = tags.filter((tag) => !tag.startsWith(RUN_TAG_PREFIX));
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
      (access) => scopes.find((scope) => sameId(scope.id, access.id))?.value ?? access.id,
    ),
    removedPermissions: permissions.removed.map((entry) => permissionName(state.graph, entry)),
    removedLabels,
  };
}

function certificateStep(
  state: SetupState,
  blockers: string[],
  warnings: string[],
): CertificateStep {
  const { secrets } = state;
  const { credential } = secrets;
  const shown = (path: string) => relative(state.cwd, path) || path;
  const pendingOnRegistration =
    secrets.pending !== null &&
    hasCertificate(state.keyCredentials, certificateKey(secrets.pending.certificate));
  // A rotation that reached the registration but not the files is finished first, whatever else
  // was asked: its certificate is already the one Prism365 should move to.
  if (interruptedRotation(secrets) || pendingOnRegistration) {
    if (!pendingOnRegistration) {
      blockers.push(
        `A rotation stopped while moving ${shown(state.files.pendingKey)} into place, but its certificate is not on the registration. Restore the working pair from ${shown(state.files.previousKey)} and ${shown(state.files.previousCertificate)} by hand, then run the setup again.`,
      );
      return { action: "blocked" };
    }
    if (state.options.rotateCertificate)
      blockers.push(
        "An earlier rotation is not finished yet. Run the setup without --rotate-certificate first; it finishes that rotation.",
      );
    return { action: "promote" };
  }
  if (secrets.pendingFiles && !state.options.rotateCertificate)
    warnings.push(
      `${shown(state.files.pendingKey)} and ${shown(state.files.pendingCertificate)} are left from a rotation that did not reach the registration. They are not used; --rotate-certificate replaces them, or delete them.`,
    );
  if (state.options.rotateCertificate) {
    if (credential.state === "none") return { action: "create" };
    if (secrets.previousFiles) {
      blockers.push(
        `${shown(state.files.previousKey)} is still there from the last rotation. Once Prism365 signs in with the current certificate, remove the old one from the registration under Certificates and secrets in the Entra admin center, delete the .previous files, and rotate again.`,
      );
      return { action: "blocked" };
    }
    return { action: "rotate" };
  }
  const unusableKey = (key: Parameters<typeof keyProblem>[0]) => {
    const problem = keyProblem(key);
    if (problem)
      blockers.push(
        `${shown(state.files.privateKey)} ${problem}. Run the setup again with --rotate-certificate to replace the pair with an RSA 2048 one.`,
      );
    return problem !== null;
  };
  switch (credential.state) {
    case "none":
      return { action: "create" };
    case "keyOnly":
      return unusableKey(credential.privateKey) ? { action: "blocked" } : { action: "certify" };
    case "pair": {
      if (unusableKey(credential.privateKey)) return { action: "blocked" };
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

// The new permission bits for an existing file or directory (null to leave it), and what to say
// about bits it keeps.
function modeStep(
  state: SetupState,
  path: string,
  mode: number | null,
  acl: AccessList | null,
  notes: string[],
): number | null {
  const decision = modeDecision(mode, acl, state.platform);
  const shown = relative(state.cwd, path) || path;
  if (decision.kept === "acl")
    notes.push(
      `${shown} has an access control list (ACL), so the setup keeps its group bits (the list's mask)${decision.change !== null ? " and only removes access for everyone else" : ""}.`,
    );
  if (decision.kept === "group")
    notes.push(
      `${shown} can be read by its group (mode ${mode?.toString(8)}), which on Linux may be an access control list's mask, so the setup leaves the group bits${decision.change !== null ? " and only removes access for everyone else" : ""}. If no ACL is meant to be there, run chmod ${path === state.files.directory ? "700" : "600"} ${shown}.`,
    );
  return decision.change;
}

export function buildPlan(state: SetupState): SetupPlan {
  const { options } = state;
  const blockers: string[] = [];
  const warnings: string[] = [];
  if (state.invalidEnvClientId !== null)
    blockers.push(
      `ENTRA_CLIENT_ID in ${relative(state.cwd, options.envFile) || options.envFile} is "${terminalText(state.invalidEnvClientId, 80)}", which is not an application ID. Fix or remove it, or pass --app-id.`,
    );
  if (state.missingTarget?.source === "appId")
    blockers.push(
      `No app registration with application ID ${state.missingTarget.appId} was found in this tenant.`,
    );
  if (state.missingTarget?.source === "env")
    blockers.push(
      `ENTRA_CLIENT_ID in ${relative(state.cwd, options.envFile) || options.envFile} names application ${state.missingTarget.appId}, which this tenant does not have (or you cannot see). Pass --app-id with the registration to use, or remove ENTRA_CLIENT_ID from the file to create a new one.`,
    );
  const permissions = requestedPermissions(options.features);
  const wanted = graphResourceAccess(
    permissions.map((permission) => permission.name),
    state.graph.oauth2PermissionScopes,
  );
  const verified = adoptionProblems(state, wanted, blockers, warnings);
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
    certificate.action !== "rotate"
      ? modeStep(
          state,
          state.files.privateKey,
          credential.privateKeyMode,
          state.acl.privateKey,
          warnings,
        )
      : null;
  const restrictDirectory = modeStep(
    state,
    state.files.directory,
    state.secrets.directoryMode,
    state.acl.directory,
    warnings,
  );
  // ACL entries a rotation could not copy last time, where the files still lack them.
  const pending = state.record.pendingAcl;
  const missing = (wanted: readonly string[], current: AccessList | null) =>
    wanted.filter((entry) => !entry.startsWith("mask::") && !current?.entries.includes(entry));
  const reapplyAcl =
    pending && state.platform !== "win32"
      ? {
          privateKey: missing(pending.privateKey, state.acl.privateKey),
          certificate: missing(pending.certificate, state.acl.certificate),
        }
      : null;
  // With a default ACL on the directory, new files take their named entries from it, and the
  // group bits of their mode become the ACL's mask: 0600 would switch those entries off.
  const writeModes =
    state.platform !== "win32" && (state.acl.directory?.defaults.length ?? 0) > 0
      ? { privateKey: 0o640, certificate: 0o644 }
      : ownerModes;
  if (writeModes !== ownerModes)
    warnings.push(
      `${relative(state.cwd, state.files.directory) || state.files.directory} has a default access control list, so new key files are written with mode 640 and the list decides who reads them.`,
    );
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
    verified,
    servicePrincipal:
      application.action === "create" || !state.servicePrincipal ? "create" : "keep",
    certificate,
    writeModes,
    restrictDirectory,
    restrictPrivateKey,
    reapplyAcl:
      reapplyAcl && reapplyAcl.privateKey.length + reapplyAcl.certificate.length > 0
        ? reapplyAcl
        : null,
    env,
    blockers,
    warnings,
  };
}

// The plan in plain sentences, printed before anything changes.
export function planText(plan: SetupPlan, state: SetupState): string {
  const shown = (path: string) => relative(state.cwd, path) || path;
  const windows = state.platform === "win32";
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
  const how = {
    appId: "named by --app-id",
    env: `the one ${shown(state.options.envFile)} uses`,
  };
  const checked =
    plan.verified.length > 0
      ? ` The setup checked that ${list(plan.verified)}. That is how it is now: its owners can change it at any time.`
      : "";
  if (app.action === "create")
    lines.push(
      `App registration: create "${plan.displayName}", a single-tenant web app with implicit grant off, labelled ${SETUP_TAG}. You become its owner.`,
    );
  else if (app.action === "keep")
    lines.push(
      `App registration: keep "${plan.displayName}" (application ID ${app.appId}, ${how[plan.foundBy ?? "appId"]}). It already has the redirect URIs and permissions below.${checked}`,
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
        ? [
            `remove the ${app.removedPermissions.length === 1 ? "request" : "requests"} for ${list(app.removedPermissions)} (--prune-permissions; this removes the request only, and consent already granted stays until it is revoked in the Entra admin center)`,
          ]
        : []),
      ...(app.removedLabels.length > 0
        ? [
            `take off the run ${app.removedLabels.length === 1 ? "label" : "labels"} ${list(app.removedLabels)} an earlier run left`,
          ]
        : []),
    ];
    lines.push(
      `App registration: update "${plan.displayName}" (application ID ${app.appId}, ${how[plan.foundBy ?? "appId"]}): ${list(changes)}. The update sends only ${list(Object.keys(app.patch))}.${checked}`,
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
  const access = windows
    ? "(on Windows the folder's access control list decides who can read them; see docs/SETUP.md)"
    : plan.writeModes.privateKey === 0o600
      ? "(the key readable only by you)"
      : "(the key's readers set by the folder's access control list)";
  const where = `${shown(state.files.privateKey)} and ${shown(state.files.certificate)} ${access}`;
  const until = new Date(state.now.getTime() + 365 * DAY);
  if (certificate.action === "create")
    lines.push(
      `Certificate: create an RSA 2048 key and a self-signed certificate valid until ${day(until)}, write ${where}, and add the certificate to the registration. Certificates already on the registration stay.`,
    );
  else if (certificate.action === "rotate")
    lines.push(
      `Certificate: make a new key and certificate valid until ${day(until)} as ${shown(state.files.pendingKey)} and ${shown(state.files.pendingCertificate)}, add the certificate to the registration, and only once it is there move the pair into place, keeping the old one as ${shown(state.files.previousKey)} and ${shown(state.files.previousCertificate)}. If anything fails first, the working pair is untouched. The old certificate stays on the registration until you remove it in the Entra admin center.`,
    );
  else if (certificate.action === "promote")
    lines.push(
      `Certificate: finish the earlier rotation. Its certificate is already on the registration, so move ${shown(state.files.pendingKey)} and ${shown(state.files.pendingCertificate)} into place, keeping the old pair under .previous names.`,
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
  const bits = (mode: number) => `mode ${mode.toString(8)}`;
  if (plan.restrictDirectory !== null)
    lines.push(
      `Change ${shown(state.files.directory)} to ${bits(plan.restrictDirectory)} so others cannot open it.`,
    );
  if (plan.restrictPrivateKey !== null)
    lines.push(
      `Change ${shown(state.files.privateKey)} to ${bits(plan.restrictPrivateKey)} so others cannot read it.`,
    );
  if (plan.reapplyAcl)
    lines.push(
      "Give the key and certificate the access list entries the last rotation could not copy to them.",
    );
  lines.push("");

  const envFile = shown(state.options.envFile);
  const owner = windows
    ? "(on Windows its folder's access control list decides who can read it)"
    : "readable only by you";
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
      `${envFile}: create it${plan.env.source === "example" ? " from .env.example" : ""}, ${owner}, and set:`,
      ...settings,
    );
  else if (plan.env.changes.length === 0)
    lines.push(
      `${envFile}: rewrite the lines of settings it manages as KEY=value (one per key, as scripts/docker-live.sh reads them). Every other line stays.`,
    );
  else
    lines.push(
      `${envFile}: change these settings, keep every other line, and leave the file ${owner}:`,
      ...settings,
    );

  for (const warning of plan.warnings) lines.push("", `Note: ${warning}`);
  if (plan.blockers.length > 0) {
    lines.push("", "The setup cannot continue, and it has changed nothing:");
    for (const blocker of plan.blockers) lines.push(`  ${blocker}`);
  }
  lines.push(
    "",
    "Nothing here grants consent: that stays a click in your browser afterwards. Prism365 itself only reads.",
  );
  return lines.join("\n");
}

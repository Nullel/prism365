import { type SupportedReadFeature, supportedReadFeatures } from "../../apps/api/src/config.ts";

// Microsoft Graph's own application ID, the same in every tenant.
export const GRAPH_RESOURCE_APP_ID = "00000003-0000-0000-c000-000000000000";

// Sign-in itself: MSAL always asks for openid, profile and offline_access (the refresh token
// that keeps a session and collection going), and Prism365 reads the administrator's own
// profile with User.Read. They are listed on the registration so the one admin consent covers
// them too.
export const SIGN_IN_PERMISSIONS = ["openid", "profile", "offline_access", "User.Read"] as const;

export interface RequestedPermission {
  name: string;
  // The read packs this permission serves; empty for the sign-in permissions.
  features: SupportedReadFeature[];
}

// Every delegated permission the registration asks for, each once, with the packs it serves.
export function requestedPermissions(
  features: readonly SupportedReadFeature[],
): RequestedPermission[] {
  const permissions: RequestedPermission[] = SIGN_IN_PERMISSIONS.map((name) => ({
    name,
    features: [],
  }));
  for (const feature of features) {
    const name = supportedReadFeatures[feature];
    const existing = permissions.find((permission) => permission.name === name);
    if (existing) existing.features.push(feature);
    else permissions.push({ name, features: [feature] });
  }
  return permissions;
}

export interface PublishedPermission {
  id: string;
  value: string | null;
  isEnabled?: boolean | null | undefined;
}

export interface ResourceAccess {
  id: string;
  type: string;
}

export interface RequiredResourceAccess {
  resourceAppId: string;
  resourceAccess: ResourceAccess[];
}

export class UnknownPermissionError extends Error {
  readonly names: string[];
  constructor(names: string[]) {
    super(
      `Microsoft Graph does not publish an enabled delegated permission named ${names.join(", ")}. Check the feature list and try again.`,
    );
    this.name = "UnknownPermissionError";
    this.names = names;
  }
}

const sameId = (left: string, right: string) => left.toLowerCase() === right.toLowerCase();

// The Microsoft Graph entry for requiredResourceAccess: each permission's scope ID, looked up in
// Graph's own service principal so no ID is hard-coded. Unknown or disabled names fail together.
export function graphResourceAccess(
  permissions: readonly string[],
  published: readonly PublishedPermission[],
): RequiredResourceAccess {
  const unknown: string[] = [];
  const resourceAccess: ResourceAccess[] = [];
  for (const name of permissions) {
    const scope = published.find((item) => item.value === name && item.isEnabled !== false);
    if (!scope) unknown.push(name);
    else if (!resourceAccess.some((access) => sameId(access.id, scope.id)))
      resourceAccess.push({ id: scope.id, type: "Scope" });
  }
  if (unknown.length > 0) throw new UnknownPermissionError(unknown);
  return { resourceAppId: GRAPH_RESOURCE_APP_ID, resourceAccess };
}

export interface PermissionMerge {
  value: RequiredResourceAccess[];
  added: ResourceAccess[];
  removed: { resourceAppId: string; access: ResourceAccess }[];
}

// What the registration should ask for. Without `prune`, every permission already on it stays,
// whoever added it; the wanted Graph scopes are added beside them. With `prune`, the list
// becomes exactly the wanted scopes.
export function mergeRequiredResourceAccess(
  existing: readonly RequiredResourceAccess[],
  wanted: RequiredResourceAccess,
  prune: boolean,
): PermissionMerge {
  const has = (list: readonly ResourceAccess[], access: ResourceAccess) =>
    list.some((item) => sameId(item.id, access.id) && item.type === access.type);
  if (prune) {
    const removed = existing.flatMap((entry) =>
      entry.resourceAccess
        .filter(
          (access) =>
            !sameId(entry.resourceAppId, wanted.resourceAppId) ||
            !has(wanted.resourceAccess, access),
        )
        .map((access) => ({ resourceAppId: entry.resourceAppId, access })),
    );
    const current = existing
      .filter((entry) => sameId(entry.resourceAppId, wanted.resourceAppId))
      .flatMap((entry) => entry.resourceAccess);
    return {
      value: [{ resourceAppId: wanted.resourceAppId, resourceAccess: [...wanted.resourceAccess] }],
      added: wanted.resourceAccess.filter((access) => !has(current, access)),
      removed,
    };
  }
  const value = existing.map((entry) => ({
    resourceAppId: entry.resourceAppId,
    resourceAccess: [...entry.resourceAccess],
  }));
  const graph = value.find((entry) => sameId(entry.resourceAppId, wanted.resourceAppId));
  if (!graph) {
    value.push({ resourceAppId: wanted.resourceAppId, resourceAccess: [...wanted.resourceAccess] });
    return { value, added: [...wanted.resourceAccess], removed: [] };
  }
  const added = wanted.resourceAccess.filter((access) => !has(graph.resourceAccess, access));
  graph.resourceAccess.push(...added);
  return { value, added, removed: [] };
}

import { type AuditRow, cleanText } from "@prism365/contracts";
import { GraphRequestError, GraphRetryDeferredError } from "@prism365/graph";
import { z } from "zod";
import type { Audit, AuditContext } from "./definition.ts";
import { type PagedRead, readPages } from "./graph-pages.ts";

// What each application in the tenant is allowed to do: the application permissions granted to it
// (it acts with no one signed in), the delegated permissions an administrator granted for every
// user, and how many users consented for themselves. Microsoft Graph v1.0 documentation verified
// 2026-10-22:
// https://learn.microsoft.com/en-us/graph/api/serviceprincipal-list?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/serviceprincipal-list-approleassignments?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/oauth2permissiongrant-list?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/resources/approle?view=graph-rest-1.0

const applicationContext = {
  feature: "directory.applications.read",
  permission: "Application.Read.All",
} as const;
const grantContext = {
  feature: "directory.permission-grants.read",
  permission: "Directory.Read.All",
} as const;

export const MAX_SERVICE_PRINCIPALS = 2_000;
export const MAX_REQUESTS = 600;
// Requests kept back for turning permission IDs into names, so a tenant that uses the whole
// budget on applications still gets readable permissions.
const NAME_LOOKUP_RESERVE = 40;
const MAX_GRANT_REQUESTS = 100;
const MAX_GRANTS = 10_000;
const MAX_LIST_ENTRIES = 200;
const NO_SPECIFIC_ROLE = "00000000-0000-0000-0000-000000000000";

// The tenants Microsoft's own applications belong to. The first owns the apps that run Microsoft
// 365 services; the second is Microsoft's corporate tenant.
// https://learn.microsoft.com/en-us/troubleshoot/azure/active-directory/verify-first-party-apps-sign-in
export const MICROSOFT_TENANTS: ReadonlySet<string> = new Set([
  "f8cdef31-a31e-4b4a-93e4-5f571e91255a",
  "72f988bf-86f1-41af-91ab-2d7cd011db47",
]);

// Permissions that reach mail, files, chats, the directory or administrator roles. Prism365's own
// short list, not Microsoft's: a prompt to look at an app, not a verdict on it. The same names are
// used for application permissions and for delegated ones.
export const HIGH_RISK_PERMISSIONS = [
  "Mail.ReadWrite",
  "Mail.Read",
  "Mail.Send",
  "Mail.ReadBasic.All",
  "Files.ReadWrite.All",
  "Sites.ReadWrite.All",
  "Directory.ReadWrite.All",
  "RoleManagement.ReadWrite.Directory",
  "AppRoleAssignment.ReadWrite.All",
  "Application.ReadWrite.All",
  "User.ReadWrite.All",
  "full_access_as_app",
  "EWS.AccessAsUser.All",
  "Chat.Read.All",
] as const;
const highRiskByName = new Map<string, string>(
  HIGH_RISK_PERMISSIONS.map((name) => [name.toLowerCase(), name]),
);

// The high-risk permissions among these names, as the list spells them, sorted.
export function highRiskIn(names: Iterable<string>): string[] {
  const found = new Set<string>();
  for (const name of names) {
    const known = highRiskByName.get(name.toLowerCase());
    if (known) found.add(known);
  }
  return [...found].sort((a, b) => a.localeCompare(b));
}

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
// IDs go into request paths, so only the characters of a GUID get through.
const objectId = z
  .string()
  .min(1)
  .max(100)
  .regex(/^[A-Za-z0-9-]+$/);

const principalSchema = z.object({
  id: objectId,
  appId: text,
  displayName: text,
  accountEnabled: flag,
  servicePrincipalType: text,
  appOwnerOrganizationId: text,
  verifiedPublisher: z
    .object({ displayName: text })
    .nullish()
    .catch(null)
    .transform((value) => value ?? null),
});
type Principal = z.output<typeof principalSchema>;

const assignmentSchema = z.object({ resourceId: text, appRoleId: text });

const roleSchema = z.object({ id: text, value: text });
const resourceSchema = z.object({
  appRoles: z
    .array(z.unknown())
    .max(5_000)
    .nullish()
    .catch(null)
    .transform((roles) =>
      (roles ?? []).flatMap((item) => {
        const role = roleSchema.safeParse(item);
        return role.success && role.data.id && role.data.value
          ? [{ id: role.data.id.toLowerCase(), value: role.data.value }]
          : [];
      }),
    ),
});

const grantSchema = z.object({
  clientId: objectId,
  consentType: text,
  principalId: text,
  // Microsoft allows 3,850 characters of space-separated scope names.
  scope: z
    .string()
    .transform((value) => cleanText(value.slice(0, 4_000)))
    .nullish()
    .catch(null)
    .transform((value) => value ?? ""),
});

interface ClientGrants {
  // Delegated permissions an administrator granted for every user.
  all: Set<string>;
  // Users who consented for themselves, and what they consented to.
  users: Set<string>;
  userScopes: Set<string>;
}

// One shared count of Graph requests for the whole run.
class RequestBudget {
  used = 0;
  readonly #limit: number;
  constructor(limit: number) {
    this.#limit = limit;
  }
  // Takes one request unless that would leave fewer than `keep` for later.
  take(keep = 0): boolean {
    if (this.#limit - this.used <= keep) return false;
    this.used += 1;
    return true;
  }
}

export type Owner = "This tenant" | "Microsoft" | "Another organisation" | "Unknown";

// Who registered the app. Without the tenant ID there is no telling this tenant's own apps from
// other organisations', so those stay Unknown.
export function ownerOf(organizationId: string | null, tenantId: string | undefined): Owner {
  if (!organizationId) return "Unknown";
  const id = organizationId.toLowerCase();
  if (MICROSOFT_TENANTS.has(id)) return "Microsoft";
  if (!tenantId) return "Unknown";
  return id === tenantId.toLowerCase() ? "This tenant" : "Another organisation";
}

function capped(names: readonly string[]): string[] {
  if (names.length <= MAX_LIST_ENTRIES) return [...names];
  return [
    ...names.slice(0, MAX_LIST_ENTRIES - 1),
    `and ${names.length - (MAX_LIST_ENTRIES - 1)} more`,
  ];
}

const sorted = (names: Iterable<string>) => [...names].sort((a, b) => a.localeCompare(b));

const plural = (count: number, one: string, many: string) =>
  `${count.toLocaleString("en-US")} ${count === 1 ? one : many}`;

interface Assignment {
  resourceId: string;
  roleId: string;
}

interface Grants {
  byClient: Map<string, ClientGrants>;
  // Whether delegated grants were read at all, and why not when they were not.
  state: "read" | "not_granted" | "refused";
  complete: boolean;
  skipped: number;
}

async function readGrants(context: AuditContext, budget: RequestBudget): Promise<Grants> {
  const byClient = new Map<string, ClientGrants>();
  if (!context.granted.has("Directory.Read.All"))
    return { byClient, state: "not_granted", complete: false, skipped: 0 };
  let requests = 0;
  let read: PagedRead;
  try {
    read = await readPages(context, grantContext, "oauth2PermissionGrants", {
      maxRecords: MAX_GRANTS,
      allow: () => {
        if (requests >= MAX_GRANT_REQUESTS || !budget.take()) return false;
        requests += 1;
        return true;
      },
    });
  } catch (error) {
    // The grants only add detail: a refusal costs the audit its delegated columns, not the run.
    if (error instanceof GraphRequestError && error.status === 403)
      return { byClient, state: "refused", complete: false, skipped: 0 };
    throw error;
  }
  let skipped = 0;
  for (const raw of read.items) {
    const parsed = grantSchema.safeParse(raw);
    if (!parsed.success) {
      skipped += 1;
      continue;
    }
    const { clientId, consentType, principalId, scope } = parsed.data;
    const scopes = scope
      .split(/\s+/)
      .filter(Boolean)
      .map((name) => name.slice(0, 200));
    if (scopes.length === 0) continue;
    const entry = byClient.get(clientId) ?? {
      all: new Set<string>(),
      users: new Set<string>(),
      userScopes: new Set<string>(),
    };
    byClient.set(clientId, entry);
    const kind = consentType?.toLowerCase();
    if (kind === "allprincipals") for (const name of scopes) entry.all.add(name);
    else if (kind === "principal" && principalId) {
      entry.users.add(principalId);
      for (const name of scopes) entry.userScopes.add(name);
    } else skipped += 1;
  }
  return { byClient, state: "read", complete: read.complete, skipped };
}

export const appConsentsAudit: Audit = {
  definition: {
    id: "apps.consents",
    version: 1,
    title: "App consents",
    description:
      "Every application that holds permissions in this tenant: what it can do without anyone signed in, what an administrator granted for all users, how many users consented for themselves, and which permissions are high-risk.",
    category: "Applications",
    subject: { one: "application", many: "applications" },
    columns: [
      { key: "app", label: "Application", kind: "text", empty: "Unnamed application" },
      {
        key: "publisher",
        label: "Publisher",
        kind: "text",
        help: "The publisher Microsoft has verified for the app. Only apps owned by another organisation count as unverified: Microsoft's own apps and apps registered in this tenant do not.",
      },
      {
        key: "owner",
        label: "Owner",
        kind: "enum",
        facet: true,
        help: "Who registered the app: this tenant, Microsoft, or another organisation. Unknown when Microsoft gave no owner.",
      },
      { key: "enabled", label: "Enabled", kind: "boolean", empty: "Unknown" },
      {
        key: "highRisk",
        label: "High-risk permissions",
        kind: "list",
        help: "Permissions on Prism365's own short list: mail, files, chats, the directory and administrator roles. A prompt to review the app, not a verdict.",
      },
      {
        key: "appPermissions",
        label: "Application permissions",
        kind: "list",
        help: "Granted to the app itself. It can use them with no one signed in.",
      },
      {
        key: "delegated",
        label: "Delegated, all users",
        kind: "list",
        empty: "Not read",
        help: "Granted by an administrator for every user. The app uses them only while a person is using it. Not read without Directory.Read.All.",
      },
      {
        key: "consentedUsers",
        label: "Users who consented",
        kind: "number",
        empty: "Not read",
        help: "Users who agreed to the app's permissions for themselves, with no administrator involved.",
      },
      {
        key: "userScopes",
        label: "Permissions users consented to",
        kind: "list",
        hidden: true,
        empty: "Not read",
      },
      { key: "appId", label: "Application ID", kind: "text", hidden: true },
    ],
    views: [
      { id: "high-risk", label: "High-risk permissions", tags: ["high-risk"] },
      { id: "app-high-risk", label: "High-risk application permissions", tags: ["app-high-risk"] },
      {
        id: "unverified-high-risk",
        label: "Unverified with high-risk permissions",
        tags: ["unverified", "high-risk"],
      },
      { id: "app-permissions", label: "Application permissions", tags: ["app-permissions"] },
      { id: "admin-consent", label: "Admin consent for everyone", tags: ["admin-consent"] },
      { id: "user-consent", label: "Consented by users", tags: ["user-consent"] },
      { id: "unverified", label: "Unverified publishers", tags: ["unverified"] },
      { id: "other-org", label: "Another organisation", tags: ["other-org"] },
      { id: "all", label: "All", tags: [] },
    ],
    defaultView: "high-risk",
    defaultSort: { column: "highRisk", direction: "desc" },
    requirements: [
      "A role that can read applications and their app role assignments, such as Directory Readers, Application Administrator or Cloud Application Administrator. Microsoft does not list Global Reader for reading app role assignments.",
      "Delegated permissions (optional) need Directory.Read.All and a role such as Global Reader or Directory Readers.",
    ],
  },
  packs: ["directory.applications.read"],
  optionalPacks: ["directory.permission-grants.read"],
  scoped: false,
  forbidden:
    "Microsoft refused to read applications or their permissions. It needs Application.Read.All and a role that can read app role assignments: Directory Readers, Application Administrator or Cloud Application Administrator. Microsoft does not list Global Reader for that call.",
  async run(context) {
    const budget = new RequestBudget(MAX_REQUESTS);
    let complete = true;
    const notes: string[] = [];

    // Every service principal, a page of 100 at a time, with only the fields the audit uses (a
    // principal's own role and scope lists can run to hundreds of entries).
    const list = await readPages(
      context,
      applicationContext,
      "servicePrincipals?$select=id,appId,displayName,accountEnabled,servicePrincipalType,appOwnerOrganizationId,verifiedPublisher&$top=100",
      { maxRecords: MAX_SERVICE_PRINCIPALS, allow: () => budget.take() },
    );
    let skipped = 0;
    const applications: Principal[] = [];
    for (const raw of list.items) {
      const parsed = principalSchema.safeParse(raw);
      if (!parsed.success) {
        skipped += 1;
        continue;
      }
      // Managed identities and the other types are not applications, and neither is a record that
      // names no type: such a record is not unreadable, so it does not make the result partial.
      if (parsed.data.servicePrincipalType?.toLowerCase() === "application")
        applications.push(parsed.data);
    }
    if (!list.complete) {
      complete = false;
      notes.push(
        list.throttled
          ? `Microsoft asked Prism365 to slow down after ${plural(list.items.length, "service principal", "service principals")}, so the rest were not read. Run the audit again in a few minutes.`
          : `Stopped after ${plural(list.items.length, "service principal", "service principals")}; the rest of the directory was not read.`,
      );
    }

    const grants = await readGrants(context, budget);
    if (grants.state === "not_granted")
      notes.push(
        "Delegated permissions are not shown: the session has not granted Directory.Read.All, which lists what administrators and users consented to. Only application permissions are listed.",
      );
    else if (grants.state === "refused")
      notes.push(
        "Microsoft refused to list delegated permission grants. They need Directory.Read.All and a role such as Global Reader or Directory Readers. Only application permissions are listed.",
      );
    else if (!grants.complete) {
      notes.push(
        "Not every delegated permission grant was read (Prism365 stops at a fixed number of requests), so some consents may be missing.",
      );
    }
    const grantsRead = grants.state === "read";
    if (!grantsRead || !grants.complete) complete = false;

    // The application permissions of each app, third-party apps first so that a tenant too large
    // for the request budget loses Microsoft's own apps before anything else.
    const microsoftOwned = (app: Principal) =>
      ownerOf(app.appOwnerOrganizationId, context.tenantId) === "Microsoft";
    const queue = [
      ...applications.filter((app) => !microsoftOwned(app)),
      ...applications.filter(microsoftOwned),
    ];
    const assignments = new Map<string, Assignment[]>();
    let examined = 0;
    let readStopped: "budget" | "throttled" | null = null;
    let unreadable = 0;
    for (const app of queue) {
      context.signal.throwIfAborted();
      let read: PagedRead;
      try {
        read = await readPages(
          context,
          applicationContext,
          `servicePrincipals/${encodeURIComponent(app.id)}/appRoleAssignments`,
          { maxRecords: 1_000, allow: () => budget.take(NAME_LOOKUP_RESERVE) },
        );
      } catch (error) {
        if (error instanceof GraphRetryDeferredError) {
          readStopped = "throttled";
          break;
        }
        // The app was deleted after it was listed.
        if (error instanceof GraphRequestError && error.status === 404) {
          examined += 1;
          continue;
        }
        throw error;
      }
      const found: Assignment[] = [];
      for (const raw of read.items) {
        const parsed = assignmentSchema.safeParse(raw);
        if (!parsed.success || !parsed.data.resourceId || !parsed.data.appRoleId) {
          unreadable += 1;
          continue;
        }
        if (parsed.data.appRoleId === NO_SPECIFIC_ROLE) continue;
        found.push({ resourceId: parsed.data.resourceId, roleId: parsed.data.appRoleId });
      }
      assignments.set(app.id, found);
      if (!read.complete) {
        readStopped = read.throttled ? "throttled" : "budget";
        break;
      }
      examined += 1;
    }
    if (readStopped) {
      complete = false;
      notes.push(
        readStopped === "throttled"
          ? `Microsoft asked Prism365 to slow down after the application permissions of ${examined.toLocaleString("en-US")} of ${plural(queue.length, "application", "applications")} were read. Run the audit again in a few minutes.`
          : `Read the application permissions of ${examined.toLocaleString("en-US")} of ${plural(queue.length, "application", "applications")} before reaching the limit of ${MAX_REQUESTS} requests per run. Microsoft's own apps are read last, so they are the ones left out.`,
      );
    }

    // Permission IDs to names, once per resource the apps hold permissions on (Microsoft Graph,
    // Exchange Online, SharePoint and so on).
    const resourceIds = new Map<string, string>();
    for (const found of assignments.values())
      for (const { resourceId } of found) resourceIds.set(resourceId.toLowerCase(), resourceId);
    const names = new Map<string, Map<string, string>>();
    let lookupsStopped = false;
    for (const [key, resourceId] of resourceIds) {
      context.signal.throwIfAborted();
      if (!objectId.safeParse(resourceId).success) continue;
      if (!budget.take()) {
        lookupsStopped = true;
        break;
      }
      try {
        const resource = await context.graph.get(
          `servicePrincipals/${encodeURIComponent(resourceId)}?$select=id,displayName,appRoles`,
          applicationContext,
          resourceSchema,
        );
        names.set(key, new Map(resource.appRoles.map((role) => [role.id, role.value])));
      } catch (error) {
        if (error instanceof GraphRetryDeferredError) {
          lookupsStopped = true;
          break;
        }
        if (error instanceof GraphRequestError && error.status === 404) continue;
        throw error;
      }
    }

    let unresolved = 0;
    const permissionName = (assignment: Assignment) => {
      const name = names
        .get(assignment.resourceId.toLowerCase())
        ?.get(assignment.roleId.toLowerCase());
      if (name) return name;
      unresolved += 1;
      return assignment.roleId;
    };

    const rows: AuditRow[] = [];
    for (const app of applications) {
      const found = assignments.get(app.id) ?? [];
      const applicationNames = new Set(found.map(permissionName));
      const delegated = grants.byClient.get(app.id);
      const adminGranted = delegated?.all ?? new Set<string>();
      const userIds = delegated?.users ?? new Set<string>();
      const userScopes = delegated?.userScopes ?? new Set<string>();
      if (applicationNames.size === 0 && adminGranted.size === 0 && userIds.size === 0) continue;
      const owner = ownerOf(app.appOwnerOrganizationId, context.tenantId);
      const verified = app.verifiedPublisher?.displayName ?? null;
      const unverified = !verified && (owner === "Another organisation" || owner === "Unknown");
      const applicationRisk = highRiskIn(applicationNames);
      const highRisk = highRiskIn([...applicationNames, ...adminGranted, ...userScopes]);
      rows.push({
        key: app.id,
        personId: null,
        tags: [
          ...(highRisk.length > 0 ? ["high-risk"] : []),
          ...(applicationRisk.length > 0 ? ["app-high-risk"] : []),
          ...(applicationNames.size > 0 ? ["app-permissions"] : []),
          ...(adminGranted.size > 0 ? ["admin-consent"] : []),
          ...(userIds.size > 0 ? ["user-consent"] : []),
          ...(unverified ? ["unverified"] : []),
          ...(owner === "Another organisation" ? ["other-org"] : []),
        ],
        cells: {
          app: app.displayName,
          publisher: verified ?? (owner === "Microsoft" ? "Microsoft" : "Unverified"),
          owner,
          enabled: app.accountEnabled,
          highRisk,
          appPermissions: capped(sorted(applicationNames)),
          delegated: grantsRead ? capped(sorted(adminGranted)) : null,
          consentedUsers: grantsRead ? userIds.size : null,
          userScopes: grantsRead ? capped(sorted(userScopes)) : null,
          appId: app.appId,
        },
      });
    }
    rows.sort((a, b) =>
      String(a.cells.app ?? "")
        .toLowerCase()
        .localeCompare(String(b.cells.app ?? "").toLowerCase()),
    );

    const count = (...tags: string[]) =>
      rows.filter((row) => tags.every((tag) => row.tags.includes(tag))).length;
    const tone = (value: number, level: "warning" | "critical") => (value > 0 ? level : "good");
    // Without delegated grants these counts are only what the application permissions show.
    const grantTone = (value: number) => (grantsRead ? tone(value, "warning") : "neutral");

    if (unresolved > 0) {
      complete = false;
      notes.push(
        `${plural(unresolved, "application permission", "application permissions")} could not be matched to a name, because Microsoft did not return the app that defines ${unresolved === 1 ? "it" : "them"}${lookupsStopped ? " or the request limit was reached" : ""}. ${unresolved === 1 ? "It is" : "They are"} shown as ${unresolved === 1 ? "an ID" : "IDs"}, and the high-risk check cannot see ${unresolved === 1 ? "it" : "them"}.`,
      );
    }
    if (skipped > 0) {
      complete = false;
      notes.push(
        `${plural(skipped, "service principal", "service principals")} could not be read and ${skipped === 1 ? "is" : "are"} left out.`,
      );
    }
    if (unreadable > 0 || grants.skipped > 0) {
      complete = false;
      notes.push(
        `${plural(unreadable + grants.skipped, "permission record", "permission records")} could not be read and ${unreadable + grants.skipped === 1 ? "is" : "are"} left out.`,
      );
    }

    return {
      rows,
      complete,
      summary: [
        {
          label: "Unverified apps with high-risk permissions",
          value: count("unverified", "high-risk"),
          tone: tone(count("unverified", "high-risk"), "critical"),
          view: "unverified-high-risk",
        },
        {
          label: "Apps with high-risk application permissions",
          value: count("app-high-risk"),
          tone: tone(count("app-high-risk"), "warning"),
          view: "app-high-risk",
        },
        {
          label: "Apps with admin consent for everyone",
          value: count("admin-consent"),
          tone: grantTone(count("admin-consent")),
          view: "admin-consent",
        },
        {
          label: "Apps consented by individual users",
          value: count("user-consent"),
          tone: "neutral",
          view: "user-consent",
        },
      ],
      notes: [
        "Microsoft's own apps are left out unless they hold permissions granted in this tenant.",
        "Permissions granted or removed in the last few minutes may not show yet, because Microsoft's directory takes time to replicate.",
        "Application permissions let an app act with no one signed in. Delegated permissions apply only while a person is using the app.",
        `High-risk means the permission is on Prism365's short list: ${HIGH_RISK_PERMISSIONS.join(", ")}. A prompt to review an app, not a verdict.`,
        "Unverified means Microsoft has not verified the app's publisher. Only apps owned by another organisation count: Microsoft's own apps and apps registered in this tenant do not.",
        ...notes,
      ],
    };
  },
};

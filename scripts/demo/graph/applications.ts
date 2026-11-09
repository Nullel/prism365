import type { SyntheticTenant } from "../tenant.ts";
import { guid, idFilter, SYNTHETIC_TENANT_ID } from "../tenant.ts";
import type { DemoGraphExtension } from "./types.ts";

// Enterprise applications for the synthetic tenant, the application permissions granted to them
// and the delegated permissions people and administrators consented to. Every application and
// publisher is invented; Microsoft's own service principals carry only the names the audit has to
// resolve, and the same fixed ID Microsoft documents for the tenant that owns its apps. Six
// application registrations, each with the credentials and owners one App credentials finding needs.

const MICROSOFT_TENANT = "f8cdef31-a31e-4b4a-93e4-5f571e91255a";

interface AppRole {
  id: string;
  value: string;
  displayName: string;
  description: string;
  isEnabled: boolean;
  allowedMemberTypes: string[];
}

function roles(resource: string, names: readonly [string, string][]): AppRole[] {
  return names.map(([value, displayName]) => ({
    id: guid(`app-role:${resource}:${value}`),
    value,
    displayName,
    description: displayName,
    isEnabled: true,
    allowedMemberTypes: ["Application"],
  }));
}

const graphRoles = roles("microsoft-graph", [
  ["Mail.Read", "Read mail in all mailboxes"],
  ["Mail.ReadWrite", "Read and write mail in all mailboxes"],
  ["Mail.Send", "Send mail as any user"],
  ["Mail.ReadBasic.All", "Read basic mail in all mailboxes"],
  ["Calendars.Read", "Read calendars in all mailboxes"],
  ["Chat.Read.All", "Read all chat messages"],
  ["User.Read.All", "Read all users' full profiles"],
  ["Group.Read.All", "Read all groups"],
  ["Directory.Read.All", "Read directory data"],
  ["AuditLog.Read.All", "Read all audit log data"],
  ["Reports.Read.All", "Read all usage reports"],
  ["Files.Read.All", "Read files in all site collections"],
  ["Files.ReadWrite.All", "Read and write files in all site collections"],
  ["Sites.Read.All", "Read items in all site collections"],
  ["Sites.ReadWrite.All", "Read and write items in all site collections"],
  ["Application.ReadWrite.All", "Read and write all applications"],
  ["RoleManagement.ReadWrite.Directory", "Read and write all directory RBAC settings"],
]);
const exchangeRoles = roles("exchange-online", [
  ["full_access_as_app", "Use Exchange Web Services with full access to all mailboxes"],
  ["Exchange.ManageAsApp", "Manage Exchange As Application"],
]);

const PEOPLE = {
  orrery: [3, 4, 5, 6],
  pinwheel: [10, 11, 12, 13, 14, 15],
  copperline: [20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31],
};

interface Principal {
  id: string;
  appId: string;
  displayName: string;
  accountEnabled: boolean;
  servicePrincipalType: string;
  appOwnerOrganizationId: string;
  verifiedPublisher: {
    displayName: string;
    verifiedPublisherId: string;
    addedDateTime: string;
  } | null;
  appRoles: AppRole[];
  tags: string[];
}

interface World {
  principals: Principal[];
  assignments: Map<string, object[]>;
  grants: object[];
  registrations: object[];
  // The owners of each registration, by its object ID.
  owners: Map<string, object[]>;
}

const DAY = 86_400_000;

function build(tenant: SyntheticTenant): World {
  const added = new Date(tenant.now.getTime() - 400 * 86_400_000).toISOString();
  const members = tenant.users.filter((user) => user.userType === "Member" && user.accountEnabled);
  const verified = (name: string, label: string) => ({
    displayName: name,
    verifiedPublisherId: String(
      2_000_000 + (Number.parseInt(guid(label).slice(0, 6), 16) % 900_000),
    ),
    addedDateTime: added,
  });
  const principal = (
    key: string,
    displayName: string,
    options: Partial<Pick<Principal, "accountEnabled" | "servicePrincipalType" | "appRoles">> & {
      owner?: string;
      publisher?: string;
    } = {},
  ): Principal => ({
    id: guid(`sp:${key}`),
    appId: guid(`app:${key}`),
    displayName,
    accountEnabled: options.accountEnabled ?? true,
    servicePrincipalType: options.servicePrincipalType ?? "Application",
    appOwnerOrganizationId: options.owner ?? guid(`org:${key}`),
    verifiedPublisher: options.publisher ? verified(options.publisher, `publisher:${key}`) : null,
    appRoles: options.appRoles ?? [],
    tags: [],
  });

  const microsoftGraph = principal("microsoft-graph", "Microsoft Graph", {
    owner: MICROSOFT_TENANT,
    appRoles: graphRoles,
  });
  const exchange = principal("exchange-online", "Office 365 Exchange Online", {
    owner: MICROSOFT_TENANT,
    appRoles: exchangeRoles,
  });
  const commandLine = principal("graph-command-line", "Graph Command Line Tools", {
    owner: MICROSOFT_TENANT,
  });
  const portal = principal("office-portal", "Office 365 Portal", { owner: MICROSOFT_TENANT });
  const hybrid = principal("exchange-hybrid", "Exchange Hybrid Agent", { owner: MICROSOFT_TENANT });
  const quillnet = principal("quillnet", "Quillnet Mail Archiver");
  const orrery = principal("orrery", "Orrery Analytics", { publisher: "Orrery Software Ltd" });
  const tessellate = principal("tessellate", "Tessellate Planner", {
    publisher: "Tessellate Labs",
  });
  const pinwheel = principal("pinwheel", "Pinwheel Forms");
  const copperline = principal("copperline", "Copperline Helpdesk", {
    publisher: "Copperline Systems",
  });
  const marrow = principal("marrow", "Marrow Backup", { publisher: "Marrow Systems" });
  const fernwood = principal("fernwood", "Fernwood Timesheets", {
    accountEnabled: false,
    publisher: "Fernwood Ltd",
  });
  const driftwood = principal("driftwood", "Driftwood Notes Sync");
  const portalApp = principal("lab-portal", "Larkspur Lab Portal", { owner: SYNTHETIC_TENANT_ID });
  const gateway = principal("instrument-gateway", "Larkspur Instrument Gateway", {
    owner: SYNTHETIC_TENANT_ID,
  });
  const identity = principal("pipeline-identity", "lab-pipeline-identity", {
    owner: SYNTHETIC_TENANT_ID,
    servicePrincipalType: "ManagedIdentity",
  });

  const principals = [
    microsoftGraph,
    exchange,
    commandLine,
    portal,
    hybrid,
    quillnet,
    orrery,
    tessellate,
    pinwheel,
    copperline,
    marrow,
    fernwood,
    driftwood,
    portalApp,
    gateway,
    identity,
  ];

  const role = (resource: Principal, value: string) => {
    const found = resource.appRoles.find((item) => item.value === value);
    if (!found) throw new Error(`The synthetic ${resource.displayName} has no role ${value}.`);
    return { resourceId: resource.id, appRoleId: found.id };
  };
  const held: [Principal, Principal, string][] = [
    [quillnet, microsoftGraph, "Mail.Read"],
    [marrow, microsoftGraph, "Mail.ReadWrite"],
    [marrow, microsoftGraph, "Files.Read.All"],
    [portalApp, microsoftGraph, "User.Read.All"],
    [portalApp, microsoftGraph, "Group.Read.All"],
    [gateway, microsoftGraph, "Sites.ReadWrite.All"],
    [gateway, microsoftGraph, "Files.ReadWrite.All"],
    [hybrid, exchange, "full_access_as_app"],
    [identity, microsoftGraph, "Reports.Read.All"],
  ];
  const assignments = new Map<string, object[]>();
  for (const [client, resource, value] of held) {
    const list = assignments.get(client.id) ?? [];
    list.push({
      id: guid(`assignment:${client.id}:${resource.id}:${value}`),
      ...role(resource, value),
      createdDateTime: added,
      deletedDateTime: null,
      principalId: client.id,
      principalDisplayName: client.displayName,
      principalType: "ServicePrincipal",
      resourceDisplayName: resource.displayName,
    });
    assignments.set(client.id, list);
  }

  const grants: object[] = [];
  const grant = (
    client: Principal,
    consentType: "AllPrincipals" | "Principal",
    scope: string,
    user?: string,
  ) =>
    grants.push({
      id: guid(`grant:${client.id}:${consentType}:${user ?? "all"}`),
      clientId: client.id,
      consentType,
      principalId: user ?? null,
      resourceId: microsoftGraph.id,
      scope,
    });
  grant(orrery, "AllPrincipals", "User.Read Calendars.Read");
  grant(tessellate, "AllPrincipals", "Files.ReadWrite.All offline_access");
  grant(fernwood, "AllPrincipals", "User.Read");
  grant(commandLine, "AllPrincipals", "User.Read.All Directory.Read.All");
  const consenting = (positions: number[]) =>
    positions.flatMap((position) => {
      const user = members[position % members.length];
      return user ? [user.id] : [];
    });
  for (const user of consenting(PEOPLE.orrery)) grant(orrery, "Principal", "User.Read", user);
  for (const user of consenting(PEOPLE.pinwheel))
    grant(pinwheel, "Principal", "User.Read Mail.Send", user);
  for (const user of consenting(PEOPLE.copperline))
    grant(copperline, "Principal", "User.Read Sites.Read.All", user);

  const { registrations, owners } = buildRegistrations(tenant, members, principals, principal);
  return { principals, assignments, grants, registrations, owners };
}

// One registration per App credentials finding. Each has an enterprise application with the same
// application ID, so its panel opens, but none of them holds a permission, so App consents skips
// them.
function buildRegistrations(
  tenant: SyntheticTenant,
  members: SyntheticTenant["users"],
  principals: Principal[],
  principal: (key: string, displayName: string, options?: { owner?: string }) => Principal,
): { registrations: object[]; owners: Map<string, object[]> } {
  const at = (days: number) => new Date(tenant.now.getTime() + days * DAY).toISOString();
  const credential = (label: string, startedDaysAgo: number, endsInDays: number | null) => ({
    keyId: guid(`credential:${label}`),
    displayName: label,
    startDateTime: at(-startedDaysAgo),
    endDateTime: endsInDays === null ? null : at(endsInDays),
  });
  const registrations: object[] = [];
  const owners = new Map<string, object[]>();
  const register = (
    key: string,
    displayName: string,
    options: {
      audience: string;
      createdDaysAgo: number;
      secrets?: ReturnType<typeof credential>[];
      certificates?: ReturnType<typeof credential>[];
      redirects: string[];
      owners: number[];
    },
  ) => {
    const enterprise = principal(key, displayName, { owner: SYNTHETIC_TENANT_ID });
    principals.push(enterprise);
    const id = guid(`registration:${key}`);
    registrations.push({
      id,
      appId: enterprise.appId,
      displayName,
      createdDateTime: at(-options.createdDaysAgo),
      signInAudience: options.audience,
      passwordCredentials: options.secrets ?? [],
      keyCredentials: options.certificates ?? [],
      web: { redirectUris: options.redirects },
      spa: { redirectUris: [] },
      publicClient: { redirectUris: [] },
    });
    owners.set(
      id,
      options.owners.flatMap((position) => {
        const user = members[position % members.length];
        return user
          ? [
              {
                "@odata.type": "#microsoft.graph.user",
                id: user.id,
                displayName: user.displayName,
                userPrincipalName: user.userPrincipalName,
              },
            ]
          : [];
      }),
    );
  };
  register("payroll-export", "Larkspur Payroll Export", {
    audience: "AzureADMyOrg",
    createdDaysAgo: 540,
    secrets: [credential("payroll-export secret", 540, -40)],
    redirects: ["https://payroll.larkspur.example/signin-callback"],
    owners: [40],
  });
  register("badge-sync", "Larkspur Badge Sync", {
    audience: "AzureADMyOrg",
    createdDaysAgo: 353,
    secrets: [credential("badge-sync secret", 353, 12)],
    redirects: ["https://badges.larkspur.example/auth"],
    owners: [41, 42],
  });
  register("reporting-bridge", "Larkspur Reporting Bridge", {
    audience: "AzureADMyOrg",
    createdDaysAgo: 305,
    secrets: [credential("reporting-bridge secret", 305, 60)],
    redirects: ["https://reports.larkspur.example/auth"],
    owners: [43],
  });
  register("legacy-importer", "Larkspur Legacy Importer", {
    audience: "AzureADMyOrg",
    createdDaysAgo: 1_100,
    secrets: [credential("legacy-importer secret", 1_095, null)],
    redirects: ["https://import.larkspur.example/auth"],
    owners: [44],
  });
  register("visitor-kiosk", "Larkspur Visitor Kiosk", {
    audience: "AzureADMyOrg",
    createdDaysAgo: 220,
    secrets: [credential("visitor-kiosk secret", 220, 145)],
    redirects: ["http://kiosk.larkspur.example/signin"],
    owners: [],
  });
  register("partner-portal", "Larkspur Partner Portal", {
    audience: "AzureADMultipleOrgs",
    createdDaysAgo: 150,
    secrets: [credential("partner-portal secret", 150, 215)],
    certificates: [credential("partner-portal certificate", 150, 580)],
    redirects: ["https://partners.larkspur.example/auth"],
    owners: [45, 46],
  });
  return { registrations, owners };
}

const built = new WeakMap<SyntheticTenant, World>();

// Enterprise applications, their application permissions and delegated grants.
export const applications: DemoGraphExtension = (request, tools) => {
  const { path } = request;
  if (
    path !== "servicePrincipals" &&
    path !== "oauth2PermissionGrants" &&
    path !== "applications" &&
    !path.startsWith("servicePrincipals/") &&
    !path.startsWith("applications/")
  )
    return null;
  let world = built.get(request.tenant);
  if (!world) {
    world = build(request.tenant);
    built.set(request.tenant, world);
  }
  if (path === "servicePrincipals") {
    const only = idFilter(request.url);
    const listed = only ? world.principals.filter((item) => only.has(item.id)) : world.principals;
    return tools.page(request.url, listed, request.select);
  }
  if (path === "applications") return tools.page(request.url, world.registrations, request.select);
  const registered = /^applications\/([^/]+)\/owners$/.exec(path);
  if (registered) {
    const found = world.owners.get(registered[1] ?? "");
    return found ? tools.page(request.url, found, request.select) : tools.notFound();
  }
  if (path === "oauth2PermissionGrants") return tools.page(request.url, world.grants, null);
  const assigned = /^servicePrincipals\/([^/]+)\/appRoleAssignments$/.exec(path);
  if (assigned)
    return tools.page(request.url, world.assignments.get(assigned[1] ?? "") ?? [], null);
  const one = /^servicePrincipals\/([^/]+)$/.exec(path);
  const found = world.principals.find((item) => item.id === one?.[1]);
  if (!found) return tools.notFound();
  return tools.json({ id: found.id, displayName: found.displayName, appRoles: found.appRoles });
};

import { guid, type SyntheticTenant } from "../tenant.ts";
import type { DemoGraphExtension } from "./types.ts";

// The groups and teams of Larkspur Analytical, for the Groups and Teams audit: fourteen groups
// that between them have every kind and every status. Owners are accounts from the roster.
//
// An ownerless team, a team owned only by a disabled account, a team with one owner and guests,
// a public team nobody has used for months, and a team the Teams report does not list. Among the
// other groups: a Microsoft 365 group about to expire, one with no owners, a group that can hold
// a directory role, a dynamic group, a mail-enabled security group, a distribution list and one
// synced from on-premises. The Teams activity report that goes with the teams is served by
// usage-reports.ts, from TEAM_USAGE below.

const DAY = 86_400_000;

// The groups the memberships in usage-reports.ts name, so a person's group and the group list
// agree.
export const STAFF = { id: guid("group:staff"), name: "Larkspur staff" };
export const LICENSING = {
  id: guid("group:licensing"),
  name: "Licensing, Microsoft 365 Business Premium",
};
export const LABORATORY = { id: guid("group:laboratory"), name: "Laboratory team" };
// The group roles.ts gives Helpdesk Administrator.
const IT_OPERATIONS = { id: guid("group:it-operations-admins"), name: "IT operations admins" };

type Visibility = "Public" | "Private";

interface DemoGroup {
  id: string;
  name: string;
  // What Microsoft puts in groupTypes, and the three flags that tell the kinds apart.
  types: string[];
  mailEnabled: boolean;
  securityEnabled: boolean;
  team?: boolean;
  visibility?: Visibility;
  createdDaysAgo: number;
  // Only a Microsoft 365 group under an expiration policy has these.
  expiresInDays?: number;
  renewedDaysAgo?: number;
  assignable?: boolean;
  synced?: boolean;
  // Owners by roster key; omitted for the kinds whose owners are not read.
  owners?: string[];
}

const groups: readonly DemoGroup[] = [
  {
    ...LABORATORY,
    types: ["Unified"],
    mailEnabled: true,
    securityEnabled: false,
    team: true,
    visibility: "Private",
    createdDaysAgo: 640,
    expiresInDays: 301,
    renewedDaysAgo: 64,
    owners: ["isolheim", "dholloway"],
  },
  {
    id: guid("group:quality-reviews"),
    name: "Quality system reviews",
    types: ["Unified"],
    mailEnabled: true,
    securityEnabled: false,
    team: true,
    visibility: "Private",
    createdDaysAgo: 410,
    expiresInDays: 190,
    renewedDaysAgo: 175,
    // The one owner is the only person who can approve a guest.
    owners: ["korne"],
  },
  {
    id: guid("group:site-move"),
    name: "Madison site move 2024",
    types: ["Unified"],
    mailEnabled: true,
    securityEnabled: false,
    team: true,
    visibility: "Public",
    createdDaysAgo: 780,
    expiresInDays: 230,
    renewedDaysAgo: 135,
    owners: ["flowe", "mbrandt"],
  },
  {
    id: guid("group:sales-kickoff"),
    name: "Sales kickoff",
    types: ["Unified"],
    mailEnabled: true,
    securityEnabled: false,
    team: true,
    visibility: "Private",
    createdDaysAgo: 330,
    expiresInDays: 150,
    renewedDaysAgo: 215,
    // Its owner left and nobody took over.
    owners: [],
  },
  {
    id: guid("group:method-transfer"),
    name: "Chromatography method transfer",
    types: ["Unified"],
    mailEnabled: true,
    securityEnabled: false,
    team: true,
    visibility: "Private",
    createdDaysAgo: 520,
    expiresInDays: 85,
    renewedDaysAgo: 280,
    owners: ["hmoss"],
  },
  {
    id: guid("group:customer-pilots"),
    name: "Customer pilots",
    types: ["Unified"],
    mailEnabled: true,
    securityEnabled: false,
    team: true,
    visibility: "Public",
    createdDaysAgo: 24,
    expiresInDays: 341,
    renewedDaysAgo: 24,
    owners: ["yhalvorsen", "mcobb"],
  },
  {
    ...STAFF,
    types: ["Unified"],
    mailEnabled: true,
    securityEnabled: false,
    visibility: "Public",
    createdDaysAgo: 1_240,
    expiresInDays: 220,
    renewedDaysAgo: 145,
    owners: ["tvance", "rteague", "mbrandt"],
  },
  {
    id: guid("group:spring-conference"),
    name: "Spring conference planning",
    types: ["Unified"],
    mailEnabled: true,
    securityEnabled: false,
    visibility: "Private",
    createdDaysAgo: 712,
    // Renewed nearly a year ago: it expires unless an owner renews it now.
    expiresInDays: 12,
    renewedDaysAgo: 353,
    owners: ["rteague"],
  },
  {
    id: guid("group:research-collaboration"),
    name: "Research collaboration",
    types: ["Unified"],
    mailEnabled: true,
    securityEnabled: false,
    visibility: "Private",
    createdDaysAgo: 910,
    expiresInDays: 140,
    renewedDaysAgo: 225,
    owners: [],
  },
  {
    ...LICENSING,
    types: ["DynamicMembership"],
    mailEnabled: false,
    securityEnabled: true,
    createdDaysAgo: 1_100,
  },
  {
    id: guid("group:lab-alerts"),
    name: "Lab operations alerts",
    types: [],
    mailEnabled: true,
    securityEnabled: true,
    createdDaysAgo: 860,
  },
  {
    ...IT_OPERATIONS,
    types: [],
    mailEnabled: false,
    securityEnabled: true,
    assignable: true,
    createdDaysAgo: 266,
    owners: ["admin", "tvance"],
  },
  {
    id: guid("group:file-share-readers"),
    name: "Madison file share readers",
    types: [],
    mailEnabled: false,
    securityEnabled: true,
    synced: true,
    createdDaysAgo: 1_480,
  },
  {
    id: guid("group:announcements"),
    name: "Larkspur announcements",
    types: [],
    mailEnabled: true,
    securityEnabled: false,
    createdDaysAgo: 1_390,
  },
];

// What the Teams activity report says about each team: days since the last activity (null for
// none), and the counts for the last 90 days. The report does not list "Customer pilots", which
// was created last month, and it still lists one team that has been deleted since.
export interface TeamUsage {
  id: string;
  name: string;
  type: Visibility;
  lastDaysAgo: number | null;
  users: number;
  channels: number;
  guests: number;
}
const RETIRED_PILOT = guid("group:retired-pilot");
export const TEAM_USAGE: readonly TeamUsage[] = [
  {
    id: LABORATORY.id,
    name: LABORATORY.name,
    type: "Private",
    lastDaysAgo: 2,
    users: 14,
    channels: 5,
    guests: 2,
  },
  {
    id: guid("group:quality-reviews"),
    name: "Quality system reviews",
    type: "Private",
    lastDaysAgo: 6,
    users: 9,
    channels: 3,
    guests: 3,
  },
  {
    id: guid("group:site-move"),
    name: "Madison site move 2024",
    type: "Public",
    lastDaysAgo: 140,
    users: 0,
    channels: 0,
    guests: 0,
  },
  {
    id: guid("group:sales-kickoff"),
    name: "Sales kickoff",
    type: "Private",
    lastDaysAgo: 11,
    users: 22,
    channels: 4,
    guests: 0,
  },
  {
    id: guid("group:method-transfer"),
    name: "Chromatography method transfer",
    type: "Private",
    lastDaysAgo: 25,
    users: 4,
    channels: 2,
    guests: 1,
  },
  {
    id: RETIRED_PILOT,
    name: "Retired pilot team",
    type: "Private",
    lastDaysAgo: null,
    users: 0,
    channels: 0,
    guests: 0,
  },
];

function groupRecords(tenant: SyntheticTenant) {
  const now = tenant.now.getTime();
  const iso = (days: number) => new Date(now + days * DAY).toISOString();
  return groups.map((group) => ({
    id: group.id,
    displayName: group.name,
    groupTypes: group.types,
    mailEnabled: group.mailEnabled,
    securityEnabled: group.securityEnabled,
    resourceProvisioningOptions: group.team ? ["Team"] : [],
    visibility: group.visibility ?? null,
    expirationDateTime: group.expiresInDays === undefined ? null : iso(group.expiresInDays),
    renewedDateTime: group.renewedDaysAgo === undefined ? null : iso(-group.renewedDaysAgo),
    isAssignableToRole: group.assignable ?? false,
    onPremisesSyncEnabled: group.synced ? true : null,
    createdDateTime: iso(-group.createdDaysAgo),
  }));
}

function ownersOf(tenant: SyntheticTenant, group: DemoGroup) {
  return (group.owners ?? []).map((key) => {
    const user = tenant.users.find((candidate) => candidate.id === guid(`user:${key}`));
    if (!user) throw new Error(`The synthetic tenant has no user ${key}.`);
    return {
      "@odata.type": "#microsoft.graph.user",
      id: user.id,
      displayName: user.displayName,
      userPrincipalName: user.userPrincipalName,
      accountEnabled: user.accountEnabled,
    };
  });
}

const OWNERS = /^groups\/([^/]+)\/owners$/;

// The group list and each group's owners; members are not served, as the read pack cannot reach
// them either.
export const demoGroups: DemoGraphExtension = (request, tools) => {
  const { path, select, url, tenant } = request;
  if (path === "groups") return tools.page(url, groupRecords(tenant), select);
  const owned = OWNERS.exec(path);
  if (!owned) return null;
  const group = groups.find((candidate) => candidate.id === owned[1]);
  return group ? tools.page(url, ownersOf(tenant, group), select) : tools.notFound();
};

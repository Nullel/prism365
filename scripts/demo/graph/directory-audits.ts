import { createHash } from "node:crypto";
import { type DirectoryUser, guid, type SyntheticTenant } from "../tenant.ts";
import type { DemoGraphExtension } from "./types.ts";

// The directory audit log for Larkspur Analytical: four weeks of ordinary administration (new
// accounts, profile updates from the HR sync, password resets, group membership, licences) and
// the planted stories the rules should flag: one person added to Global Administrator, a new
// secret on the payroll connector, an update to the tenant-wide MFA policy, and admin consent to
// a survey app. Everything is built from the synthetic tenant's own users, deterministically
// from its clock, with activity names from Microsoft's audit activity reference.

const MINUTE = 60_000;
const HOUR = 60 * MINUTE;
const DAY = 24 * HOUR;
const WEEKS = 4;
const GLOBAL_ADMINISTRATOR = "62e90394-69f5-4237-9190-012177145e10";

type Rand = () => number;

function seeded(label: string): Rand {
  let state = createHash("sha256").update(`audit:${label}`).digest().readUInt32LE(0);
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4_294_967_296;
  };
}

const pick = <T>(rand: Rand, items: readonly T[]): T =>
  items[Math.floor(rand() * items.length)] as T;
const quoted = (value: string) => JSON.stringify(value);

interface Property {
  displayName: string;
  oldValue: string | null;
  newValue: string | null;
}

interface Target {
  id: string;
  displayName: string | null;
  type: string;
  userPrincipalName: string | null;
  modifiedProperties: Property[];
}

export interface DirectoryAuditRecord {
  id: string;
  category: string;
  correlationId: string;
  result: "success" | "failure";
  resultReason: string;
  activityDisplayName: string;
  activityDateTime: string;
  loggedByService: string;
  operationType: string;
  initiatedBy: {
    user: {
      id: string;
      displayName: string;
      userPrincipalName: string;
      ipAddress: string;
    } | null;
    app: {
      appId: string;
      displayName: string;
      servicePrincipalId: string;
      servicePrincipalName: string;
    } | null;
  };
  targetResources: Target[];
  additionalDetails: { key: string; value: string }[];
}

type Actor = DirectoryAuditRecord["initiatedBy"];

function asUser(user: DirectoryUser, ip: string): Actor {
  return {
    user: {
      id: user.id,
      displayName: user.displayName,
      userPrincipalName: user.userPrincipalName,
      ipAddress: ip,
    },
    app: null,
  };
}

// The HR system's sync app, which keeps job titles and departments in step.
const hrSync = (): Actor => ({
  user: null,
  app: {
    appId: guid("app:hr-sync"),
    displayName: "Larkspur HR Sync",
    servicePrincipalId: guid("sp:hr-sync"),
    servicePrincipalName: "Larkspur HR Sync",
  },
});

const userTarget = (user: DirectoryUser, properties: Property[] = []): Target => ({
  id: user.id,
  displayName: user.displayName,
  type: "User",
  userPrincipalName: user.userPrincipalName,
  modifiedProperties: properties,
});

const groupTarget = (key: string, name: string, properties: Property[] = []): Target => ({
  id: guid(`group:${key}`),
  displayName: name,
  type: "Group",
  userPrincipalName: null,
  modifiedProperties: properties,
});

function buildAudits(tenant: SyntheticTenant): DirectoryAuditRecord[] {
  const now = tenant.now.getTime();
  const rand = seeded(tenant.now.toISOString().slice(0, 10));
  const byName = (name: string) => {
    const user = tenant.users.find((candidate) => candidate.displayName === name);
    if (!user) throw new Error(`The synthetic tenant has no user named ${name}.`);
    return user;
  };
  const admin = byName("Demo administrator");
  const technician = byName("Tobias Vance");
  const adminActor = asUser(admin, "192.0.2.24");
  const technicianActor = asUser(technician, "192.0.2.31");
  const members = tenant.users.filter(
    (user) => user.userType === "Member" && user.accountEnabled && user.id !== admin.id,
  );
  const records: DirectoryAuditRecord[] = [];
  let sequence = 0;
  const add = (
    at: number,
    record: Omit<
      DirectoryAuditRecord,
      "id" | "activityDateTime" | "correlationId" | "result" | "resultReason"
    > &
      Partial<Pick<DirectoryAuditRecord, "result" | "resultReason">>,
  ) => {
    sequence += 1;
    if (at > now - 2 * MINUTE || at < now - WEEKS * 7 * DAY) return;
    records.push({
      id: `Directory_${guid(`audit:${sequence}`)}_LARK_${sequence}`,
      activityDateTime: new Date(at).toISOString(),
      correlationId: guid(`correlation:${sequence}`),
      result: "success",
      resultReason: "",
      ...record,
    });
  };
  // Office hours in Madison (UTC-5) on a given day back from now.
  const workday = (daysAgo: number, hour: number) =>
    Math.floor((now - daysAgo * DAY) / DAY) * DAY +
    (hour + 5) * HOUR +
    Math.floor(rand() * 50) * MINUTE;

  // New accounts: created, licensed, added to the staff group.
  for (const user of tenant.users) {
    const created = Date.parse(user.createdDateTime);
    if (now - created > WEEKS * 7 * DAY) continue;
    const guest = user.userType === "Guest";
    add(created, {
      category: "UserManagement",
      activityDisplayName: guest ? "Invite external user" : "Add user",
      loggedByService: guest ? "Invited Users" : "Core Directory",
      operationType: "Add",
      initiatedBy: adminActor,
      targetResources: [
        userTarget(user, [
          { displayName: "AccountEnabled", oldValue: null, newValue: "[true]" },
          { displayName: "DisplayName", oldValue: null, newValue: quoted(user.displayName) },
          {
            displayName: "UserPrincipalName",
            oldValue: null,
            newValue: quoted(user.userPrincipalName),
          },
        ]),
      ],
      additionalDetails: [{ key: "UserType", value: user.userType }],
    });
    if (guest) continue;
    add(created + 4 * MINUTE, {
      category: "UserManagement",
      activityDisplayName: "Change user license",
      loggedByService: "Core Directory",
      operationType: "Update",
      initiatedBy: adminActor,
      targetResources: [
        userTarget(user, [
          {
            displayName: "AssignedLicense",
            oldValue: "[]",
            newValue: quoted("[SkuName=SPB, AccountId=larkspur, SkuId=synthetic]"),
          },
        ]),
      ],
      additionalDetails: [],
    });
    add(created + 6 * MINUTE, {
      category: "GroupManagement",
      activityDisplayName: "Add member to group",
      loggedByService: "Core Directory",
      operationType: "Assign",
      initiatedBy: adminActor,
      targetResources: [
        userTarget(user),
        groupTarget("all-staff", "All staff", [
          { displayName: "Group.DisplayName", oldValue: null, newValue: quoted("All staff") },
        ]),
      ],
      additionalDetails: [],
    });
  }

  // Profile updates from the HR system, a few each week.
  const titles = ["Senior Analyst", "Laboratory Technician II", "Team Lead", "Project Coordinator"];
  const departments = ["Laboratory", "Quality", "Operations", "Finance"];
  for (let index = 0; index < 18; index += 1) {
    const user = pick(rand, members);
    const field = rand() < 0.5 ? "JobTitle" : "Department";
    const before = field === "JobTitle" ? user.jobTitle : user.department;
    const after = pick(rand, field === "JobTitle" ? titles : departments);
    add(workday(1 + Math.floor(rand() * 27), 9), {
      category: "UserManagement",
      activityDisplayName: "Update user",
      loggedByService: "Core Directory",
      operationType: "Update",
      initiatedBy: hrSync(),
      targetResources: [
        userTarget(user, [
          { displayName: field, oldValue: `[${quoted(before)}]`, newValue: `[${quoted(after)}]` },
          { displayName: "Included Updated Properties", oldValue: null, newValue: quoted(field) },
        ]),
      ],
      additionalDetails: [{ key: "UserType", value: "Member" }],
    });
  }

  // Password resets by the support technician, one of them refused.
  for (let index = 0; index < 7; index += 1) {
    const user = pick(rand, members);
    const failed = index === 3;
    add(workday(1 + Math.floor(rand() * 27), 10), {
      category: "UserManagement",
      activityDisplayName: "Reset password (by admin)",
      loggedByService: "Self-service Password Management",
      operationType: "Update",
      initiatedBy: technicianActor,
      targetResources: [userTarget(user)],
      additionalDetails: [],
      ...(failed
        ? {
            result: "failure" as const,
            resultReason: "The password does not meet the tenant's complexity requirements.",
          }
        : {}),
    });
  }

  // Field sampling staff come and go with the season.
  for (let index = 0; index < 8; index += 1) {
    const user = pick(rand, members);
    const adding = index % 3 !== 2;
    add(workday(1 + Math.floor(rand() * 27), 14), {
      category: "GroupManagement",
      activityDisplayName: adding ? "Add member to group" : "Remove member from group",
      loggedByService: "Core Directory",
      operationType: adding ? "Assign" : "Unassign",
      initiatedBy: technicianActor,
      targetResources: [
        userTarget(user),
        groupTarget("field-sampling-staff", "Field sampling staff", [
          {
            displayName: "Group.DisplayName",
            oldValue: adding ? null : quoted("Field sampling staff"),
            newValue: adding ? quoted("Field sampling staff") : null,
          },
        ]),
      ],
      additionalDetails: [],
    });
  }

  // The stories.
  add(workday(6, 16), {
    category: "RoleManagement",
    activityDisplayName: "Add member to role",
    loggedByService: "Core Directory",
    operationType: "Assign",
    initiatedBy: adminActor,
    targetResources: [
      userTarget(technician, [
        { displayName: "Role.ObjectID", oldValue: null, newValue: quoted(guid("role:global")) },
        {
          displayName: "Role.DisplayName",
          oldValue: null,
          newValue: quoted("Global Administrator"),
        },
        { displayName: "Role.TemplateId", oldValue: null, newValue: quoted(GLOBAL_ADMINISTRATOR) },
        {
          displayName: "Role.WellKnownObjectName",
          oldValue: null,
          newValue: quoted("TenantAdmins"),
        },
      ]),
    ],
    additionalDetails: [],
  });
  const key = (label: string) =>
    `[KeyIdentifier=${guid(`key:${label}`)}, KeyType=Password, KeyUsage=Verify, DisplayName=${label}]`;
  add(workday(4, 11), {
    category: "ApplicationManagement",
    activityDisplayName: "Update application – Certificates and secrets management ",
    loggedByService: "Core Directory",
    operationType: "Update",
    initiatedBy: adminActor,
    targetResources: [
      {
        id: guid("app:payroll"),
        displayName: "Larkspur Payroll Connector",
        type: "Application",
        userPrincipalName: null,
        modifiedProperties: [
          {
            displayName: "KeyDescription",
            oldValue: JSON.stringify([key("payroll-2025")]),
            newValue: JSON.stringify([key("payroll-2025"), key("payroll-october")]),
          },
          {
            displayName: "Included Updated Properties",
            oldValue: null,
            newValue: quoted("KeyDescription"),
          },
        ],
      },
    ],
    additionalDetails: [{ key: "AppId", value: guid("appid:payroll") }],
  });
  const policy = (state: string, excluded: string[]) =>
    JSON.stringify({
      id: guid("policy:mfa-all-users"),
      displayName: "Require MFA for all users",
      state,
      conditions: {
        users: { includeUsers: ["All"], excludeUsers: excluded },
        applications: { includeApplications: ["All"] },
      },
      grantControls: { operator: "OR", builtInControls: ["mfa"] },
    });
  add(workday(9, 15), {
    category: "Policy",
    activityDisplayName: "Update Conditional Access policy",
    loggedByService: "Conditional Access",
    operationType: "Update",
    initiatedBy: adminActor,
    targetResources: [
      {
        id: guid("policy:mfa-all-users"),
        displayName: "Require MFA for all users",
        type: "Policy",
        userPrincipalName: null,
        modifiedProperties: [
          {
            displayName: "ConditionalAccessPolicy",
            oldValue: policy("enabled", [admin.id]),
            newValue: policy("enabled", [admin.id, technician.id]),
          },
        ],
      },
    ],
    additionalDetails: [],
  });
  add(workday(12, 13), {
    category: "ApplicationManagement",
    activityDisplayName: "Consent to application",
    loggedByService: "Core Directory",
    operationType: "Assign",
    initiatedBy: technicianActor,
    targetResources: [
      {
        id: guid("sp:survey"),
        displayName: "Fieldnote Surveys",
        type: "ServicePrincipal",
        userPrincipalName: null,
        modifiedProperties: [
          {
            displayName: "ConsentContext.IsAdminConsent",
            oldValue: null,
            newValue: quoted("True"),
          },
          { displayName: "ConsentContext.IsAppOnly", oldValue: null, newValue: quoted("False") },
          { displayName: "ConsentContext.OnBehalfOfAll", oldValue: null, newValue: quoted("True") },
          {
            displayName: "ConsentAction.Permissions",
            oldValue: null,
            newValue: quoted("Scope: User.Read Files.Read.All offline_access"),
          },
        ],
      },
    ],
    additionalDetails: [{ key: "User-Agent", value: "Synthetic browser" }],
  });
  return records.sort(
    (a, b) => b.activityDateTime.localeCompare(a.activityDateTime) || b.id.localeCompare(a.id),
  );
}

const built = new WeakMap<SyntheticTenant, DirectoryAuditRecord[]>();

export function syntheticDirectoryAudits(tenant: SyntheticTenant): DirectoryAuditRecord[] {
  const cached = built.get(tenant);
  if (cached) return cached;
  const records = buildAudits(tenant);
  built.set(tenant, records);
  return records;
}

// Graph's documented window: `activityDateTime ge X and activityDateTime le Y`, both inclusive.
function windowFilter(filter: string) {
  const from = /activityDateTime ge ([0-9T:.\-Z+]+)/.exec(filter)?.[1];
  const to = /activityDateTime le ([0-9T:.\-Z+]+)/.exec(filter)?.[1];
  const lower = from ? Date.parse(from) : Number.NEGATIVE_INFINITY;
  const upper = to ? Date.parse(to) : Number.POSITIVE_INFINITY;
  return (record: DirectoryAuditRecord) => {
    const at = Date.parse(record.activityDateTime);
    return at >= lower && at <= upper;
  };
}

export const directoryAudits: DemoGraphExtension = ({ url, path, tenant }, { page }) => {
  if (path.toLowerCase() !== "auditlogs/directoryaudits") return null;
  const matches = windowFilter(url.searchParams.get("$filter") ?? "");
  return page(url, syntheticDirectoryAudits(tenant).filter(matches), null);
};

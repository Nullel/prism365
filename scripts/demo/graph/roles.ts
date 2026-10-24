import { privilegedRoles } from "@prism365/graph";
import type { SyntheticTenant } from "../tenant.ts";
import { DOMAIN, guid, RESEARCH_DOMAIN } from "../tenant.ts";
import type { DemoGraphExtension } from "./types.ts";

// Directory roles and domain password policies for Larkspur Analytical.
//
// Roles: three accounts hold Global Administrator now. The demo administrator holds it for good
// and so does the general manager, which is one permanent assignment too many; the IT support
// technician is eligible and has it switched on for the working day. A disabled laboratory account
// still holds Exchange Administrator, a guest holds Application Administrator, and the Exchange
// hybrid agent (an app from applications.ts) holds Exchange Administrator. One role is time-bound,
// one is only eligible, and a group holds Helpdesk Administrator, which nothing here names.
//
// Domains: the company domain keeps passwords 180 days and warns 14 days ahead, the research
// subsidiary's domain never expires them, and the initial domain leaves both settings unset.

const DAY = 86_400_000;
const HOUR = 3_600_000;

const GLOBAL_READER = "f2ef992c-3afb-46b9-b7cf-a126ee74c451";
const DIRECTORY_READERS = "88d8e3e3-8f55-4a1e-953a-9b9898b8876b";
const HELPDESK_ADMINISTRATOR = "729827e3-9c14-49f7-bb1b-9608f156bbb8";
const BILLING_ADMINISTRATOR = "b0f54661-2d74-4c50-afa3-1ec803f12efe";

const roleId = (name: string) => {
  const found = [...privilegedRoles].find(([, label]) => label === name);
  if (!found) throw new Error(`The shared privileged roles have no ${name}.`);
  return found[0];
};

function definitions() {
  const others: [string, string][] = [
    [GLOBAL_READER, "Global Reader"],
    [DIRECTORY_READERS, "Directory Readers"],
    [HELPDESK_ADMINISTRATOR, "Helpdesk Administrator"],
    [BILLING_ADMINISTRATOR, "Billing Administrator"],
  ];
  return [...privilegedRoles, ...others].map(([id, displayName]) => ({
    id,
    displayName,
    isBuiltIn: true,
    isEnabled: true,
  }));
}

interface Holding {
  holder: string;
  role: string;
  how: "permanent" | "time-bound" | "activated" | "eligible";
  // Days since the assignment started, and days until it ends when it ends.
  startedDaysAgo: number;
  endsInDays?: number;
  // The Activated instance starts and ends within hours.
  startedHoursAgo?: number;
  endsInHours?: number;
}

function world(tenant: SyntheticTenant) {
  const now = tenant.now.getTime();
  const person = (name: string) => {
    const found = tenant.users.find((candidate) => candidate.displayName === name);
    if (!found) throw new Error(`The synthetic tenant has no user named ${name}.`);
    return found.id;
  };
  const admin = person("Demo administrator");
  const technician = person("Tobias Vance");
  const manager = person("Mateus Brandt");
  const operations = person("Ferdinand Lowe");
  const people = person("Rosalind Teague");
  const departed = person("Harriet Moss");
  const guest = person("Alder Whitcombe");
  const quality = person("Kasimir Orne");
  const hybridAgent = guid("sp:exchange-hybrid");
  const itGroup = guid("group:it-operations-admins");

  const holdings: Holding[] = [
    { holder: admin, role: roleId("Global Administrator"), how: "permanent", startedDaysAgo: 612 },
    {
      holder: manager,
      role: roleId("Global Administrator"),
      how: "permanent",
      startedDaysAgo: 388,
    },
    {
      holder: technician,
      role: roleId("Global Administrator"),
      how: "activated",
      startedDaysAgo: 0,
      startedHoursAgo: 3,
      endsInHours: 5,
    },
    {
      holder: technician,
      role: roleId("Global Administrator"),
      how: "eligible",
      startedDaysAgo: 140,
    },
    { holder: technician, role: HELPDESK_ADMINISTRATOR, how: "permanent", startedDaysAgo: 301 },
    { holder: people, role: roleId("User Administrator"), how: "eligible", startedDaysAgo: 90 },
    {
      holder: operations,
      role: roleId("SharePoint Administrator"),
      how: "time-bound",
      startedDaysAgo: 9,
      endsInDays: 21,
    },
    {
      holder: departed,
      role: roleId("Exchange Administrator"),
      how: "permanent",
      startedDaysAgo: 506,
    },
    {
      holder: guest,
      role: roleId("Application Administrator"),
      how: "permanent",
      startedDaysAgo: 207,
    },
    {
      holder: hybridAgent,
      role: roleId("Exchange Administrator"),
      how: "permanent",
      startedDaysAgo: 801,
    },
    { holder: itGroup, role: HELPDESK_ADMINISTRATOR, how: "permanent", startedDaysAgo: 266 },
    { holder: quality, role: GLOBAL_READER, how: "permanent", startedDaysAgo: 150 },
  ];

  const assignments: object[] = [];
  const active: object[] = [];
  const eligible: object[] = [];
  for (const [index, holding] of holdings.entries()) {
    const label = `${holding.holder}:${holding.role}:${holding.how}:${index}`;
    const starts = holding.startedHoursAgo
      ? now - holding.startedHoursAgo * HOUR
      : now - holding.startedDaysAgo * DAY;
    const ends =
      holding.endsInHours !== undefined
        ? now + holding.endsInHours * HOUR
        : holding.endsInDays !== undefined
          ? now + holding.endsInDays * DAY
          : null;
    const common = {
      principalId: holding.holder,
      roleDefinitionId: holding.role,
      directoryScopeId: "/",
      memberType: "Direct",
      startDateTime: new Date(starts).toISOString(),
      endDateTime: ends === null ? null : new Date(ends).toISOString(),
    };
    if (holding.how === "eligible") {
      eligible.push({ id: guid(`eligibility:${label}`), ...common });
      continue;
    }
    const origin = guid(`assignment:${label}`);
    assignments.push({
      id: origin,
      principalId: holding.holder,
      roleDefinitionId: holding.role,
      directoryScopeId: "/",
    });
    active.push({
      id: guid(`instance:${label}`),
      ...common,
      assignmentType: holding.how === "activated" ? "Activated" : "Assigned",
      roleAssignmentOriginId: origin,
    });
  }

  const domain = (
    id: string,
    over: {
      validity?: number | null;
      window?: number | null;
      isDefault?: boolean;
      isInitial?: boolean;
    },
  ) => ({
    id,
    authenticationType: "Managed",
    availabilityStatus: null,
    isAdminManaged: true,
    isDefault: over.isDefault ?? false,
    isInitial: over.isInitial ?? false,
    isRoot: true,
    isVerified: true,
    passwordNotificationWindowInDays: over.window ?? null,
    passwordValidityPeriodInDays: over.validity ?? null,
    supportedServices: ["Email", "OfficeCommunicationsOnline"],
  });
  const domains = [
    domain(DOMAIN, { validity: 180, window: 14, isDefault: true }),
    // Microsoft stores "passwords never expire" as the largest 32-bit integer.
    domain(RESEARCH_DOMAIN, { validity: 2_147_483_647 }),
    domain("larkspur.onmicrosoft.com", { isInitial: true }),
  ];
  return { definitions: definitions(), assignments, active, eligible, domains };
}

const built = new WeakMap<SyntheticTenant, ReturnType<typeof world>>();

export const directoryRoles: DemoGraphExtension = (request, tools) => {
  const { path, select, url } = request;
  const prefix = "roleManagement/directory/";
  if (path !== "domains" && !path.startsWith(prefix)) return null;
  let found = built.get(request.tenant);
  if (!found) {
    found = world(request.tenant);
    built.set(request.tenant, found);
  }
  if (path === "domains") return tools.page(url, found.domains, select);
  switch (path.slice(prefix.length)) {
    case "roleDefinitions":
      return tools.page(url, found.definitions, select);
    case "roleAssignments":
      return tools.page(url, found.assignments, select);
    case "roleAssignmentScheduleInstances":
      return tools.page(url, found.active, select);
    case "roleEligibilityScheduleInstances":
      return tools.page(url, found.eligible, select);
    default:
      return null;
  }
};

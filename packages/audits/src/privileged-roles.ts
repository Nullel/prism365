import { type AuditRow, cleanText } from "@prism365/contracts";
import { GLOBAL_ADMINISTRATOR_ROLE_ID, GraphRequestError, privilegedRoles } from "@prism365/graph";
import { z } from "zod";
import type { Audit, AuditContext } from "./definition.ts";
import { MAX_ACCOUNTS, tagOf } from "./directory.ts";

// Who holds Microsoft Entra administrator roles, and how: permanently, for a time, after an
// activation, or only as eligible through Privileged Identity Management (PIM). Microsoft Graph
// v1.0 documentation verified 2026-10-24:
// https://learn.microsoft.com/en-us/graph/api/rbacapplication-list-roledefinitions?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/rbacapplication-list-roleassignments?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/rbacapplication-list-roleassignmentscheduleinstances?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/rbacapplication-list-roleeligibilityscheduleinstances?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/serviceprincipal-list?view=graph-rest-1.0

const rolesContext = {
  feature: "directory.roles.read",
  permission: "RoleManagement.Read.Directory",
} as const;
const usersContext = { feature: "directory.users.read", permission: "User.Read.All" } as const;
const applicationsContext = {
  feature: "directory.applications.read",
  permission: "Application.Read.All",
} as const;
const reportContext = {
  feature: "reports.authentication-methods.read",
  permission: "AuditLog.Read.All",
} as const;

// Microsoft documents no page size for the four role lists, so none is sent and the client
// follows the continuation links. A tenant has hundreds of role assignments, not tens of thousands;
// three lists of this size stay inside the 20,000 rows a stored result may hold.
const MAX_RECORDS = 5_000;
// Principals asked about in one request, in a `$filter=id in (...)` list.
const BATCH = 15;
// Principals looked up in all, so a very large tenant cannot turn one run into hundreds of calls.
const MAX_RESOLVED = 1_000;

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const guid = z.string().regex(GUID);
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
// Dates decide how a role is held, so a malformed one makes the record unreadable instead of
// silently reading as "no end".
const time = z
  .string()
  .datetime({ offset: true })
  .nullish()
  .transform((value) => value ?? null);

// A scope is "/" for the whole tenant or a directory object path; cut so a row key stays short.
const scope = z
  .string()
  .min(1)
  .max(200)
  .transform((value) => cleanText(value).slice(0, 100));

const definitionSchema = z.object({ id: guid, displayName: text });
const assignmentSchema = z.object({
  id: z.string().min(1).max(200),
  principalId: guid,
  roleDefinitionId: guid,
  directoryScopeId: scope,
});
const instanceSchema = z.object({
  principalId: guid,
  roleDefinitionId: guid,
  directoryScopeId: scope,
  startDateTime: time,
  endDateTime: time,
  memberType: text,
  assignmentType: text,
  roleAssignmentOriginId: z
    .string()
    .max(200)
    .nullish()
    .catch(null)
    .transform((value) => value ?? null),
});
const userSchema = z.object({
  id: guid,
  displayName: text,
  userPrincipalName: text,
  userType: text,
  accountEnabled: flag,
});
const principalSchema = z.object({ id: guid, displayName: text, accountEnabled: flag });
const registrationSchema = z.object({ id: guid, isMfaRegistered: flag });

type Held =
  | "Permanent"
  | "Time-bound"
  | "Activated"
  | "Active"
  | "Eligible"
  | "Eligible time-bound";
type Through = "Direct" | "Group" | "Inherited" | "Unknown";
type HolderType = "User" | "Guest" | "Service principal" | "Group or other";

// How a grant is held, in the order the table lists them.
const HELD_ORDER: readonly Held[] = [
  "Permanent",
  "Time-bound",
  "Activated",
  "Active",
  "Eligible",
  "Eligible time-bound",
];
const ELIGIBLE: readonly Held[] = ["Eligible", "Eligible time-bound"];
const THROUGH_ORDER: readonly Through[] = ["Direct", "Group", "Inherited", "Unknown"];

interface Grant {
  principalId: string;
  roleId: string;
  scope: string;
  held: Held;
  through: Through;
  starts: string | null;
  ends: string | null;
}

interface Holder {
  type: HolderType;
  name: string | null;
  principal: string | null;
  enabled: boolean | null;
  // Only an account has a page of its own.
  personId: string | null;
}

interface ListRead {
  items: unknown[];
  complete: boolean;
}

function readList(context: AuditContext, path: string): Promise<ListRead> {
  context.signal.throwIfAborted();
  return context.graph.getCollection(path, rolesContext, {
    schema: z.unknown(),
    maxPages: 100,
    maxRecords: MAX_RECORDS,
  });
}

// The two schedule lists need Privileged Identity Management and a role that can read it. Microsoft
// answers 403 without them and may answer 400 on a tenant that has no PIM, and neither is a failure:
// the run goes on with the active assignments alone.
async function readSchedule(context: AuditContext, path: string): Promise<ListRead | null> {
  try {
    return await readList(context, path);
  } catch (error) {
    if (error instanceof GraphRequestError && (error.status === 403 || error.status === 400))
      return null;
    throw error;
  }
}

const throughOf = (value: string | null): Through => {
  const kind = value?.toLowerCase();
  if (kind === "direct") return "Direct";
  if (kind === "group") return "Group";
  if (kind === "inherited") return "Inherited";
  return "Unknown";
};

// A grant is one principal holding one role at one scope in one way, so the same assignment read
// from two lists is one row. When it is read twice with different routes, the direct one wins.
function addGrant(grants: Map<string, Grant>, grant: Grant) {
  const key = [grant.held, grant.principalId, grant.roleId, grant.scope].join("|").toLowerCase();
  const existing = grants.get(key);
  if (!existing) grants.set(key, grant);
  else if (THROUGH_ORDER.indexOf(grant.through) < THROUGH_ORDER.indexOf(existing.through))
    existing.through = grant.through;
}

interface Joined {
  grants: Grant[];
  skipped: number;
  // Active assignments no schedule instance named, which cannot be called permanent or time-bound.
  unmatched: number;
}

// Active assignments come from roleAssignments. When the schedule lists could be read, each active
// instance names the assignment it belongs to (roleAssignmentOriginId), so it replaces that
// assignment and carries its dates; an assignment no instance names stays as "Active".
function join(
  assignments: readonly unknown[],
  active: readonly unknown[] | null,
  eligible: readonly unknown[] | null,
): Joined {
  const grants = new Map<string, Grant>();
  const origins = new Set<string>();
  let skipped = 0;
  for (const raw of active ?? []) {
    const parsed = instanceSchema.safeParse(raw);
    if (!parsed.success) {
      skipped += 1;
      continue;
    }
    const instance = parsed.data;
    if (instance.roleAssignmentOriginId) origins.add(instance.roleAssignmentOriginId.toLowerCase());
    const activated = instance.assignmentType?.toLowerCase() === "activated";
    addGrant(grants, {
      principalId: instance.principalId.toLowerCase(),
      roleId: instance.roleDefinitionId.toLowerCase(),
      scope: instance.directoryScopeId,
      held: activated ? "Activated" : instance.endDateTime ? "Time-bound" : "Permanent",
      through: throughOf(instance.memberType),
      starts: instance.startDateTime,
      ends: instance.endDateTime,
    });
  }
  let unmatched = 0;
  for (const raw of assignments) {
    const parsed = assignmentSchema.safeParse(raw);
    if (!parsed.success) {
      skipped += 1;
      continue;
    }
    const assignment = parsed.data;
    if (origins.has(assignment.id.toLowerCase())) continue;
    if (active !== null) unmatched += 1;
    // A role assignment names the principal it was given to, so it is held directly.
    addGrant(grants, {
      principalId: assignment.principalId.toLowerCase(),
      roleId: assignment.roleDefinitionId.toLowerCase(),
      scope: assignment.directoryScopeId,
      held: "Active",
      through: "Direct",
      starts: null,
      ends: null,
    });
  }
  for (const raw of eligible ?? []) {
    const parsed = instanceSchema.safeParse(raw);
    if (!parsed.success) {
      skipped += 1;
      continue;
    }
    const instance = parsed.data;
    addGrant(grants, {
      principalId: instance.principalId.toLowerCase(),
      roleId: instance.roleDefinitionId.toLowerCase(),
      scope: instance.directoryScopeId,
      held: instance.endDateTime ? "Eligible time-bound" : "Eligible",
      through: throughOf(instance.memberType),
      starts: instance.startDateTime,
      ends: instance.endDateTime,
    });
  }
  return { grants: [...grants.values()], skipped, unmatched };
}

const inBatches = <T>(items: readonly T[], size: number): T[][] => {
  const batches: T[][] = [];
  for (let index = 0; index < items.length; index += size)
    batches.push(items.slice(index, index + size));
  return batches;
};

interface Resolution {
  holders: Map<string, Holder>;
  // Whether a lookup that was tried was refused or failed, which leaves some holders unnamed.
  failed: boolean;
  usersAsked: boolean;
  applicationsAsked: boolean;
  capped: boolean;
}

// Names the principals in batches of 15 with `$filter=id in (...)`: accounts first, then service
// principals for the ones that were not accounts. Whatever stays unresolved (a group, among others)
// is a "Group or other" shown by its ID. A refused lookup never fails the run.
async function resolve(context: AuditContext, wanted: readonly string[]): Promise<Resolution> {
  const ids = wanted.slice(0, MAX_RESOLVED);
  const resolution: Resolution = {
    holders: new Map(),
    failed: false,
    usersAsked: context.granted.has("User.Read.All"),
    applicationsAsked: context.granted.has("Application.Read.All"),
    capped: wanted.length > ids.length,
  };
  const idFilter = (batch: readonly string[]) =>
    encodeURIComponent(`id in (${batch.map((id) => `'${id}'`).join(",")})`);

  if (resolution.usersAsked) {
    try {
      for (const batch of inBatches(ids, BATCH)) {
        context.signal.throwIfAborted();
        const result = await context.graph.getCollection(
          `users?$filter=${idFilter(batch)}&$select=id,displayName,userPrincipalName,userType,accountEnabled`,
          usersContext,
          { schema: z.unknown(), maxPages: 5, maxRecords: BATCH },
        );
        for (const raw of result.items) {
          const parsed = userSchema.safeParse(raw);
          if (!parsed.success) continue;
          const user = parsed.data;
          resolution.holders.set(user.id.toLowerCase(), {
            type: user.userType?.toLowerCase() === "guest" ? "Guest" : "User",
            name: user.displayName,
            principal: user.userPrincipalName,
            enabled: user.accountEnabled,
            personId: user.id,
          });
        }
      }
    } catch (error) {
      if (!(error instanceof GraphRequestError) || error.status === 401) throw error;
      resolution.failed = true;
    }
  }

  if (resolution.applicationsAsked) {
    const rest = ids.filter((id) => !resolution.holders.has(id));
    try {
      for (const batch of inBatches(rest, BATCH)) {
        context.signal.throwIfAborted();
        const result = await context.graph.getCollection(
          `servicePrincipals?$filter=${idFilter(batch)}&$select=id,displayName,accountEnabled`,
          applicationsContext,
          { schema: z.unknown(), maxPages: 5, maxRecords: BATCH },
        );
        for (const raw of result.items) {
          const parsed = principalSchema.safeParse(raw);
          if (!parsed.success) continue;
          const principal = parsed.data;
          resolution.holders.set(principal.id.toLowerCase(), {
            type: "Service principal",
            name: principal.displayName,
            principal: null,
            enabled: principal.accountEnabled,
            personId: null,
          });
        }
      }
    } catch (error) {
      if (!(error instanceof GraphRequestError) || error.status === 401) throw error;
      resolution.failed = true;
    }
  }
  return resolution;
}

// Which accounts registered a strong sign-in method, from the same report as the MFA audit. null
// when the report could not be read.
async function readMfa(context: AuditContext): Promise<Map<string, boolean | null> | null> {
  try {
    context.signal.throwIfAborted();
    const result = await context.graph.getCollection(
      "reports/authenticationMethods/userRegistrationDetails",
      reportContext,
      { schema: z.unknown(), maxPages: 100, maxRecords: MAX_ACCOUNTS },
    );
    const registered = new Map<string, boolean | null>();
    for (const raw of result.items) {
      const parsed = registrationSchema.safeParse(raw);
      if (parsed.success) registered.set(parsed.data.id.toLowerCase(), parsed.data.isMfaRegistered);
    }
    return registered;
  } catch (error) {
    if (!(error instanceof GraphRequestError) || error.status === 401) throw error;
    return null;
  }
}

interface Entry {
  grant: Grant;
  holder: Holder;
  role: string;
  privileged: boolean;
  mfa: boolean | null;
}

const unresolved = (id: string): Holder => ({
  type: "Group or other",
  name: id,
  principal: null,
  enabled: null,
  personId: null,
});

const scopeLabel = (value: string) => (value === "/" ? "Tenant" : value);
const isEligible = (held: Held) => ELIGIBLE.includes(held);
const plural = (count: number, one: string, many: string) => `${count} ${count === 1 ? one : many}`;

const GUIDANCE_LINKS =
  "https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/best-practices and https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/security-emergency-access";

export const privilegedRolesAudit: Audit = {
  definition: {
    id: "directory.privileged-roles",
    handoff: false,
    version: 1,
    title: "Privileged roles",
    description:
      "Who holds Microsoft Entra administrator roles: permanent, time-bound, activated and eligible assignments, with the disabled accounts, guests and apps among them.",
    category: "Identity",
    subject: { one: "assignment", many: "assignments" },
    columns: [
      { key: "holder", label: "Holder", kind: "person" },
      { key: "holderType", label: "Holder type", kind: "enum", facet: true },
      { key: "role", label: "Role", kind: "enum", facet: true },
      {
        key: "privileged",
        label: "Privileged",
        kind: "boolean",
        help: "One of the administrator roles Prism365 treats as privileged. Custom roles and the other built-in roles are listed but not flagged.",
      },
      {
        key: "held",
        label: "How held",
        kind: "enum",
        facet: true,
        help: "Permanent has no end date. Time-bound has one. Activated is an eligible role switched on for a while. Eligible can be activated but is not in effect. Active means Microsoft lists the assignment as in effect but Privileged Identity Management could not say more.",
      },
      {
        key: "through",
        label: "Through",
        kind: "enum",
        facet: true,
        help: "Direct is assigned to the holder itself. Group is a role the holder gets from a group it belongs to. Inherited comes from a wider scope.",
      },
      { key: "scope", label: "Scope", kind: "text" },
      { key: "starts", label: "Starts", kind: "date", empty: "Not recorded" },
      {
        key: "ends",
        label: "Ends",
        kind: "date",
        empty: "No end date",
        help: "Empty is read as no end date. Microsoft's documentation does not say so in as many words.",
      },
      {
        key: "enabled",
        label: "Account enabled",
        kind: "boolean",
        empty: "Unknown",
      },
      {
        key: "mfa",
        label: "MFA registered",
        kind: "boolean",
        empty: "Not read",
        help: "From Microsoft's authentication methods report, when that read pack is on. Accounts only.",
      },
    ],
    views: [
      { id: "privileged", label: "Privileged roles", tags: ["privileged"] },
      { id: "all", label: "All roles", tags: [] },
      { id: "global-admins", label: "Global Administrators", tags: ["global-admin", "in-effect"] },
      { id: "permanent", label: "Permanent", tags: ["privileged", "permanent"] },
      { id: "disabled", label: "Disabled accounts", tags: ["privileged", "disabled"] },
      { id: "guests", label: "Guests", tags: ["privileged", "guest"] },
      {
        id: "service-principals",
        label: "Service principals",
        tags: ["privileged", "service-principal"],
      },
    ],
    defaultView: "privileged",
    defaultSort: { column: "role", direction: "asc" },
    requirements: [
      "A role that can read directory roles, such as Global Reader, Directory Readers or Privileged Role Administrator. Directory Readers cannot read Privileged Identity Management schedules.",
      "Microsoft Entra ID P2 or ID Governance for eligible and time-bound assignments. Without it only active assignments are listed.",
    ],
  },
  packs: ["directory.roles.read"],
  optionalPacks: [
    "directory.users.read",
    "directory.applications.read",
    "reports.authentication-methods.read",
  ],
  scoped: false,
  forbidden:
    "Microsoft refused to read directory roles. It needs RoleManagement.Read.Directory and a role such as Global Reader, Directory Readers or Privileged Role Administrator.",
  async run(context) {
    const base = "roleManagement/directory/";
    const assignmentFields = "id,principalId,roleDefinitionId,directoryScopeId";
    const instanceFields =
      "principalId,roleDefinitionId,directoryScopeId,startDateTime,endDateTime,memberType";
    const definitions = await readList(context, `${base}roleDefinitions?$select=id,displayName`);
    const assignments = await readList(
      context,
      `${base}roleAssignments?$select=${assignmentFields}`,
    );
    const active = await readSchedule(
      context,
      `${base}roleAssignmentScheduleInstances?$select=${instanceFields},assignmentType,roleAssignmentOriginId`,
    );
    const eligible = await readSchedule(
      context,
      `${base}roleEligibilityScheduleInstances?$select=${instanceFields}`,
    );

    const names = new Map<string, string>();
    let skipped = 0;
    for (const raw of definitions.items) {
      const parsed = definitionSchema.safeParse(raw);
      if (!parsed.success) skipped += 1;
      else if (parsed.data.displayName)
        names.set(parsed.data.id.toLowerCase(), parsed.data.displayName);
    }
    const joined = join(assignments.items, active?.items ?? null, eligible?.items ?? null);
    skipped += joined.skipped;
    const grants = joined.grants;

    const wanted = [...new Set(grants.map((grant) => grant.principalId))].sort();
    const resolution = await resolve(context, wanted);
    const holderOf = (id: string) => resolution.holders.get(id) ?? unresolved(id);

    // Registration details are only worth reading when an account holds a role.
    const accounts = wanted.filter((id) => {
      const type = resolution.holders.get(id)?.type;
      return type === "User" || type === "Guest";
    });
    const mfaAsked = context.granted.has("AuditLog.Read.All") && accounts.length > 0;
    const registered = mfaAsked ? await readMfa(context) : null;

    const entries: Entry[] = grants.map((grant) => {
      const holder = holderOf(grant.principalId);
      const role = names.get(grant.roleId) ?? privilegedRoles.get(grant.roleId) ?? grant.roleId;
      return {
        grant,
        holder,
        role,
        privileged: privilegedRoles.has(grant.roleId),
        mfa: registered && holder.personId ? (registered.get(grant.principalId) ?? null) : null,
      };
    });
    entries.sort(
      (a, b) =>
        Number(b.privileged) - Number(a.privileged) ||
        a.role.localeCompare(b.role) ||
        HELD_ORDER.indexOf(a.grant.held) - HELD_ORDER.indexOf(b.grant.held) ||
        (a.holder.name ?? "").localeCompare(b.holder.name ?? "") ||
        a.grant.principalId.localeCompare(b.grant.principalId),
    );

    const rows: AuditRow[] = entries.map(({ grant, holder, role, privileged, mfa }) => ({
      key: [grant.held, grant.principalId, grant.roleId, grant.scope].join("|").toLowerCase(),
      personId: holder.personId,
      tags: [
        ...(privileged ? ["privileged"] : []),
        ...(grant.roleId === GLOBAL_ADMINISTRATOR_ROLE_ID ? ["global-admin"] : []),
        ...(isEligible(grant.held) ? [] : ["in-effect"]),
        ...tagOf(grant.held),
        ...(holder.enabled === false ? ["disabled"] : []),
        ...(holder.type === "Guest" ? ["guest"] : []),
        ...(holder.type === "Service principal" ? ["service-principal"] : []),
      ],
      cells: {
        holder: { name: holder.name, principal: holder.principal },
        holderType: holder.type,
        role,
        privileged,
        held: grant.held,
        through: grant.through,
        scope: scopeLabel(grant.scope),
        starts: grant.starts,
        ends: grant.ends,
        enabled: holder.enabled,
        mfa,
      },
    }));

    // Tiles count holders, not rows: one account with three roles is one holder. Only the
    // permanent tile counts assignments, as its label says.
    const holders = (keep: (entry: Entry) => boolean) =>
      new Set(entries.filter(keep).map((entry) => entry.grant.principalId));
    const isGlobalAdmin = (entry: Entry) => entry.grant.roleId === GLOBAL_ADMINISTRATOR_ROLE_ID;
    const administrators = holders(
      (entry) => isGlobalAdmin(entry) && !isEligible(entry.grant.held),
    );
    // Accounts that can activate the role but do not hold it now.
    const eligibleOnly = [
      ...holders((entry) => isGlobalAdmin(entry) && isEligible(entry.grant.held)),
    ].filter((id) => !administrators.has(id)).length;
    const privilegedHolders = (keep: (entry: Entry) => boolean) =>
      holders((entry) => entry.privileged && keep(entry)).size;
    const pimRead = active !== null;
    const permanent = entries.filter(
      (entry) => entry.privileged && entry.grant.held === "Permanent",
    ).length;
    const disabled = privilegedHolders((entry) => entry.holder.enabled === false);
    const guests = privilegedHolders((entry) => entry.holder.type === "Guest");
    const applications = privilegedHolders((entry) => entry.holder.type === "Service principal");

    const scheduleMissing = [
      ...(active === null ? ["active assignment schedules"] : []),
      ...(eligible === null ? ["eligible assignments"] : []),
    ];
    const unnamed = resolution.holders.size < wanted.length;
    const notes = [
      `Global Administrators: ${plural(administrators.size, "account holds", "accounts hold")} the role now${eligibleOnly > 0 ? `, and ${eligibleOnly} more can activate it` : ""}. Microsoft advises fewer than five Global Administrators and at least two emergency access accounts: ${GUIDANCE_LINKS}. The tile warns above 4 and below 2.`,
      ...(scheduleMissing.length > 0
        ? [
            `Microsoft did not return ${scheduleMissing.join(" or ")} from Privileged Identity Management, which needs Microsoft Entra ID P2 or ID Governance and a role that can read it, such as Global Reader or Privileged Role Administrator. Without them this list shows active assignments only and cannot tell permanent from time-bound or find eligible roles.`,
          ]
        : []),
      ...(pimRead && joined.unmatched > 0
        ? [
            `${plural(joined.unmatched, "active assignment has", "active assignments have")} no matching Privileged Identity Management record and ${joined.unmatched === 1 ? "is" : "are"} shown as Active.`,
          ]
        : []),
      `Privileged means one of ${privilegedRoles.size} built-in administrator roles that Prism365 lists, found by role template ID. Custom roles and other built-in roles are shown but not flagged.`,
      "A role given to a group is one row for the group, shown as Group or other. Its members appear only when Privileged Identity Management lists them with a Group route.",
      "Microsoft's documentation does not say that a missing end date means a permanent assignment. Prism365 reads it that way.",
      ...(!resolution.usersAsked
        ? [
            "Holders are shown by ID because the session has not granted User.Read.All, so accounts could not be told from groups.",
          ]
        : []),
      ...(resolution.usersAsked && !resolution.applicationsAsked && unnamed
        ? [
            "Some holders are shown by ID as Group or other. If any of them is an app, naming it needs the applications read pack (Application.Read.All).",
          ]
        : []),
      ...(resolution.failed
        ? ["Microsoft refused to name some holders. They are shown by ID as Group or other."]
        : []),
      ...(resolution.capped
        ? [`Only the first ${MAX_RESOLVED.toLocaleString("en-US")} holders were named.`]
        : []),
      ...(mfaAsked && registered === null
        ? [
            "Microsoft refused the authentication methods report, so MFA registration is empty. It needs AuditLog.Read.All and a role such as Reports Reader, Security Reader or Global Reader.",
          ]
        : []),
      ...(!definitions.complete ||
      !assignments.complete ||
      active?.complete === false ||
      eligible?.complete === false
        ? [
            `Stopped at ${MAX_RECORDS.toLocaleString("en-US")} records in a list; the rest of it was not read.`,
          ]
        : []),
      ...(skipped > 0
        ? [
            `${plural(skipped, "record", "records")} could not be read and ${skipped === 1 ? "is" : "are"} left out.`,
          ]
        : []),
    ];

    return {
      rows,
      complete:
        definitions.complete &&
        assignments.complete &&
        active?.complete !== false &&
        eligible?.complete !== false &&
        scheduleMissing.length === 0 &&
        skipped === 0 &&
        !resolution.failed &&
        !(mfaAsked && registered === null),
      summary: [
        {
          label: "Global Administrators",
          value: administrators.size,
          tone: administrators.size > 4 || administrators.size < 2 ? "warning" : "good",
          view: "global-admins",
        },
        {
          label: "Privileged role holders",
          value: privilegedHolders(() => true),
          tone: "neutral",
          view: "privileged",
        },
        pimRead
          ? {
              label: "Permanent privileged assignments",
              value: permanent,
              tone: permanent > 0 ? "warning" : "good",
              view: "permanent",
            }
          : {
              label: "Permanent privileged assignments (could not be read)",
              value: 0,
              tone: "neutral",
              view: null,
            },
        {
          label: "Disabled accounts with a privileged role",
          value: disabled,
          tone: disabled > 0 ? "critical" : "neutral",
          view: "disabled",
        },
        {
          label: "Guests with a privileged role",
          value: guests,
          tone: guests > 0 ? "warning" : "neutral",
          view: "guests",
        },
        {
          label: "Service principals with a privileged role",
          value: applications,
          tone: "neutral",
          view: "service-principals",
        },
      ],
      notes,
    };
  },
};

import { auditRowSchema, queryAuditRows } from "@prism365/contracts";
import { GraphHttpClient, GraphRequestError, privilegedRoles } from "@prism365/graph";
import { scopeGraph } from "@prism365/modules";
import { describe, expect, it } from "vitest";
import {
  type AuditContext,
  type AuditResult,
  auditDefinition,
  builtInAudits,
  privilegedRolesAudit,
} from "../src/index.ts";

// Synthetic tenant data only: invented people and apps, example.invalid domains, made-up IDs.
const now = new Date("2026-10-22T12:00:00Z");
const inDays = (days: number) => new Date(now.getTime() + days * 86_400_000).toISOString();
const id = (n: number) => `00000000-0000-4000-8000-${n.toString(16).padStart(12, "0")}`;

const GLOBAL_ADMIN = "62e90394-69f5-4237-9190-012177145e10";
const EXCHANGE_ADMIN = "29232cdf-9323-42fd-ade2-1d097af3e4de";
const USER_ADMIN = "fe930be7-5e62-47db-91af-98c3a49a38b1";
const HELPDESK = "729827e3-9c14-49f7-bb1b-9608f156bbb8";
const CUSTOM = "0c000000-0000-4000-8000-000000000001";
const UNIT = "/administrativeUnits/0a000000-0000-4000-8000-000000000001";

const definitions = [
  { id: GLOBAL_ADMIN, displayName: "Global Administrator" },
  { id: EXCHANGE_ADMIN, displayName: "Exchange Administrator" },
  { id: USER_ADMIN, displayName: "User Administrator" },
  { id: HELPDESK, displayName: "Helpdesk Administrator" },
  { id: CUSTOM, displayName: "Lab Reports Reader" },
];

type Json = Record<string, unknown>;
type Route = (url: URL) => Response | undefined;

const assignment = (n: number, principal: string, role: string, scope = "/") => ({
  id: `a-${n}`,
  principalId: principal,
  roleDefinitionId: role,
  directoryScopeId: scope,
});
const instance = (principal: string, role: string, over: Json = {}) => ({
  id: `i-${principal}-${role}`,
  principalId: principal,
  roleDefinitionId: role,
  directoryScopeId: "/",
  startDateTime: "2026-01-05T09:00:00Z",
  endDateTime: null,
  memberType: "Direct",
  ...over,
});
const assigned = (principal: string, role: string, origin: string, over: Json = {}) =>
  instance(principal, role, {
    assignmentType: "Assigned",
    roleAssignmentOriginId: origin,
    ...over,
  });

const person = (n: number, over: Json = {}) => ({
  id: id(n),
  displayName: `Person ${n}`,
  userPrincipalName: `person${n}@example.invalid`,
  userType: "Member",
  accountEnabled: true,
  ...over,
});

interface Tenant {
  assignments: Json[];
  active?: Json[];
  eligible?: Json[];
  users?: Json[];
  applications?: Json[];
  registrations?: Json[];
  definitions?: Json[];
}

// Microsoft's `id in (...)` filter, enough of it for these tests.
const filtered = (url: URL, records: readonly Json[]) => {
  const filter = url.searchParams.get("$filter") ?? "";
  const wanted = new Set([...filter.matchAll(/'([^']+)'/g)].map((match) => match[1]));
  return records.filter((record) => wanted.has(String(record.id)));
};

function setup(
  tenant: Tenant,
  options: {
    granted?: string[];
    override?: Route;
    optional?: (
      | "directory.users.read"
      | "directory.applications.read"
      | "reports.authentication-methods.read"
    )[];
  } = {},
) {
  const requests: URL[] = [];
  const methods: string[] = [];
  const route: Route = (url) => {
    const replaced = options.override?.(url);
    if (replaced) return replaced;
    const path = url.pathname.replace("/v1.0/", "");
    if (path === "roleManagement/directory/roleDefinitions")
      return Response.json({ value: tenant.definitions ?? definitions });
    if (path === "roleManagement/directory/roleAssignments")
      return Response.json({ value: tenant.assignments });
    if (path === "roleManagement/directory/roleAssignmentScheduleInstances")
      return Response.json({ value: tenant.active ?? [] });
    if (path === "roleManagement/directory/roleEligibilityScheduleInstances")
      return Response.json({ value: tenant.eligible ?? [] });
    if (path === "users") return Response.json({ value: filtered(url, tenant.users ?? []) });
    if (path === "servicePrincipals")
      return Response.json({ value: filtered(url, tenant.applications ?? []) });
    if (path === "reports/authenticationMethods/userRegistrationDetails")
      return Response.json({ value: tenant.registrations ?? [] });
    return undefined;
  };
  const client = new GraphHttpClient({
    getAccessToken: async () => "synthetic",
    maxResponseBytes: 8_388_608,
    maxRetries: 0,
    fetchImplementation: async (input, init) => {
      const url = new URL(String(input));
      requests.push(url);
      methods.push(init?.method ?? "GET");
      return route(url) ?? new Response("{}", { status: 404 });
    },
  });
  const context: AuditContext = {
    graph: scopeGraph(client, [
      ...privilegedRolesAudit.packs,
      ...privilegedRolesAudit.optionalPacks,
    ]),
    now,
    granted: new Set(
      options.granted ?? ["RoleManagement.Read.Directory", "User.Read.All", "Application.Read.All"],
    ),
    signal: new AbortController().signal,
    scope: null,
  };
  return { context, requests, methods };
}

async function run(tenant: Tenant, options: Parameters<typeof setup>[1] = {}) {
  const { context, requests, methods } = setup(tenant, options);
  const result = await privilegedRolesAudit.run(context);
  const rowsOf = (principal: string) =>
    result.rows.filter((row) => row.key.includes(`|${principal}|`));
  return { result, requests, methods, rowsOf };
}

const tile = (result: AuditResult, label: string) =>
  result.summary.find((item) => item.label === label);
const held = (rows: { cells: Json }[]) => rows.map((row) => row.cells.held).sort();
const refused =
  (status: number): Route =>
  () =>
    new Response("{}", { status });
const only =
  (suffix: string, status: number): Route =>
  (url) =>
    url.pathname.endsWith(suffix) ? new Response("{}", { status }) : undefined;

describe("how a role is held", () => {
  const alice = id(1);
  const bob = id(2);
  const carol = id(3);
  const dave = id(4);
  const erin = id(5);
  const frank = id(6);
  const tenant: Tenant = {
    assignments: [
      assignment(1, alice, GLOBAL_ADMIN),
      assignment(2, bob, EXCHANGE_ADMIN),
      assignment(3, carol, GLOBAL_ADMIN),
      assignment(4, frank, USER_ADMIN),
    ],
    active: [
      assigned(alice, GLOBAL_ADMIN, "a-1"),
      assigned(bob, EXCHANGE_ADMIN, "a-2", { endDateTime: inDays(30) }),
      instance(carol, GLOBAL_ADMIN, {
        assignmentType: "Activated",
        roleAssignmentOriginId: "a-3",
        startDateTime: "2026-10-22T08:00:00Z",
        endDateTime: inDays(0.3),
      }),
    ],
    eligible: [
      instance(carol, GLOBAL_ADMIN),
      instance(dave, USER_ADMIN, { endDateTime: inDays(60) }),
      instance(erin, EXCHANGE_ADMIN),
    ],
    users: [alice, bob, carol, dave, erin, frank].map((p) =>
      person(Number.parseInt(p.slice(-12), 16)),
    ),
  };

  it("gives each assignment one row, whichever lists name it", async () => {
    const { result } = await run(tenant);
    // Four assignments, three of them also named by an active instance, plus three eligible: 7 rows.
    expect(result.rows).toHaveLength(7);
    expect(new Set(result.rows.map((row) => row.key)).size).toBe(7);
  });

  it("tells permanent, time-bound, activated, eligible, eligible time-bound and active apart", async () => {
    const { rowsOf } = await run(tenant);
    expect(held(rowsOf(alice))).toEqual(["Permanent"]);
    expect(held(rowsOf(bob))).toEqual(["Time-bound"]);
    expect(held(rowsOf(carol))).toEqual(["Activated", "Eligible"]);
    expect(held(rowsOf(dave))).toEqual(["Eligible time-bound"]);
    expect(held(rowsOf(erin))).toEqual(["Eligible"]);
    // Named by no instance although the schedules were read.
    expect(held(rowsOf(frank))).toEqual(["Active"]);
  });

  it("carries the dates of the instance, and none for an assignment alone", async () => {
    const { rowsOf } = await run(tenant);
    expect(rowsOf(bob)[0]?.cells).toMatchObject({
      starts: "2026-01-05T09:00:00Z",
      ends: inDays(30),
    });
    expect(rowsOf(alice)[0]?.cells).toMatchObject({ starts: "2026-01-05T09:00:00Z", ends: null });
    expect(rowsOf(frank)[0]?.cells).toMatchObject({ starts: null, ends: null });
  });

  it("says when an active assignment has no schedule record", async () => {
    const { result } = await run(tenant);
    expect(result.notes.join("\n")).toContain(
      "1 active assignment has no matching Privileged Identity Management record",
    );
  });

  it("matches an instance to its assignment without regard to letter case", async () => {
    const { result } = await run({
      assignments: [assignment(1, alice, GLOBAL_ADMIN)],
      active: [assigned(alice, GLOBAL_ADMIN, "A-1")],
    });
    expect(result.rows).toHaveLength(1);
    expect(result.rows[0]?.cells.held).toBe("Permanent");
  });

  it("keeps a permanent and a time-bound assignment of the same role as two rows", async () => {
    const { result } = await run({
      assignments: [assignment(1, alice, EXCHANGE_ADMIN), assignment(2, alice, EXCHANGE_ADMIN)],
      active: [
        assigned(alice, EXCHANGE_ADMIN, "a-1"),
        assigned(alice, EXCHANGE_ADMIN, "a-2", { endDateTime: inDays(10) }),
      ],
    });
    expect(held(result.rows)).toEqual(["Permanent", "Time-bound"]);
  });

  it("merges the same grant read twice, keeping the direct route", async () => {
    const { result } = await run({
      assignments: [],
      active: [
        assigned(alice, EXCHANGE_ADMIN, "a-1", { memberType: "Group" }),
        assigned(alice, EXCHANGE_ADMIN, "a-9", { memberType: "Direct" }),
        assigned(alice, EXCHANGE_ADMIN, "a-8", { memberType: "Inherited" }),
      ],
    });
    expect(result.rows).toHaveLength(1);
    expect(result.rows[0]?.cells.through).toBe("Direct");
  });

  it("shows how a role reached the holder and at what scope", async () => {
    const { result } = await run({
      assignments: [assignment(1, alice, USER_ADMIN, UNIT)],
      active: [
        assigned(alice, USER_ADMIN, "a-1", { memberType: "Inherited", directoryScopeId: UNIT }),
      ],
      eligible: [
        instance(bob, USER_ADMIN, { memberType: "Group" }),
        instance(carol, USER_ADMIN, { memberType: "Something new" }),
      ],
    });
    expect(result.rows.map((row) => [row.cells.scope, row.cells.through]).sort()).toEqual([
      [UNIT, "Inherited"],
      ["Tenant", "Group"],
      ["Tenant", "Unknown"],
    ]);
  });

  it("names the scope Tenant for / and keeps any other scope as it is", async () => {
    const { result } = await run({
      assignments: [assignment(1, alice, USER_ADMIN), assignment(2, alice, USER_ADMIN, UNIT)],
    });
    expect(result.rows.map((row) => row.cells.scope).sort()).toEqual([UNIT, "Tenant"]);
  });

  it("reads more than one page", async () => {
    const { result, requests } = await run(
      { assignments: [] },
      {
        override: (url) => {
          if (!url.pathname.endsWith("/roleAssignments")) return undefined;
          if (url.searchParams.has("$skiptoken"))
            return Response.json({ value: [assignment(2, bob, USER_ADMIN)] });
          const next = new URL(url);
          next.searchParams.set("$skiptoken", "2");
          return Response.json({
            value: [assignment(1, alice, USER_ADMIN)],
            "@odata.nextLink": next.href,
          });
        },
      },
    );
    expect(result.rows).toHaveLength(2);
    expect(requests.filter((url) => url.pathname.endsWith("/roleAssignments"))).toHaveLength(2);
  });
});

describe("when Privileged Identity Management cannot be read", () => {
  const alice = id(1);
  const tenant: Tenant = {
    assignments: [assignment(1, alice, GLOBAL_ADMIN), assignment(2, id(2), EXCHANGE_ADMIN)],
    active: [assigned(alice, GLOBAL_ADMIN, "a-1")],
    eligible: [instance(id(3), GLOBAL_ADMIN)],
    users: [person(1), person(2), person(3)],
  };

  for (const status of [403, 400]) {
    it(`carries on with the active assignments when both schedule lists answer ${status}`, async () => {
      const { result } = await run(tenant, {
        override: (url) =>
          url.pathname.endsWith("ScheduleInstances") ? new Response("{}", { status }) : undefined,
      });
      expect(result.complete).toBe(false);
      expect(result.rows).toHaveLength(2);
      expect(held(result.rows)).toEqual(["Active", "Active"]);
      const notes = result.notes.join("\n");
      expect(notes).toContain("active assignment schedules or eligible assignments");
      expect(notes).toContain("Microsoft Entra ID P2");
      expect(notes).toContain("cannot tell permanent from time-bound");
    });
  }

  it("says the permanent assignments could not be read instead of counting none", async () => {
    const { result } = await run(tenant, {
      override: (url) =>
        url.pathname.endsWith("ScheduleInstances")
          ? new Response("{}", { status: 403 })
          : undefined,
    });
    const names = result.summary.map((item) => item.label);
    expect(names).toContain("Permanent privileged assignments (could not be read)");
    expect(names).not.toContain("Permanent privileged assignments");
    expect(tile(result, "Permanent privileged assignments (could not be read)")).toMatchObject({
      value: 0,
      tone: "neutral",
      view: null,
    });
  });

  it("names only the list that was refused and still counts permanent assignments", async () => {
    const { result } = await run(tenant, {
      override: only("roleEligibilityScheduleInstances", 403),
    });
    expect(result.complete).toBe(false);
    expect(result.notes.join("\n")).toContain("did not return eligible assignments from");
    expect(tile(result, "Permanent privileged assignments")?.value).toBe(1);
  });

  it("does not hide a throttled or failing schedule list", async () => {
    await expect(
      run(tenant, { override: only("roleAssignmentScheduleInstances", 500) }),
    ).rejects.toBeInstanceOf(GraphRequestError);
  });

  it("fails when the definitions or assignments are refused", async () => {
    await expect(run(tenant, { override: only("/roleAssignments", 403) })).rejects.toMatchObject({
      status: 403,
    });
    await expect(run(tenant, { override: only("/roleDefinitions", 403) })).rejects.toMatchObject({
      status: 403,
    });
  });

  it("is complete when everything is read", async () => {
    const { result } = await run({ ...tenant, eligible: [] });
    expect(result.complete).toBe(true);
  });
});

describe("naming the holders", () => {
  it("asks in batches of 15 and falls back to service principals for what accounts do not name", async () => {
    const accounts = Array.from({ length: 31 }, (_, index) => id(100 + index));
    const app = id(500);
    const group = id(600);
    const { result, requests, rowsOf } = await run({
      assignments: [...accounts, app, group].map((holder, index) =>
        assignment(index, holder, USER_ADMIN),
      ),
      users: accounts.map((_, index) => person(100 + index)),
      applications: [{ id: app, displayName: "Larkspur Backup", accountEnabled: true }],
    });
    const batches = (path: string) =>
      requests
        .filter((url) => url.pathname === path)
        .map((url) => (url.searchParams.get("$filter")?.match(/'/g)?.length ?? 0) / 2);
    expect(batches("/v1.0/users")).toEqual([15, 15, 3]);
    // The app and the group are what accounts did not name: one batch of two.
    expect(batches("/v1.0/servicePrincipals")).toEqual([2]);
    expect(rowsOf(app)[0]?.cells).toMatchObject({
      holderType: "Service principal",
      holder: { name: "Larkspur Backup", principal: null },
    });
    expect(rowsOf(group)[0]).toMatchObject({
      personId: null,
      cells: { holderType: "Group or other", holder: { name: group, principal: null } },
    });
    expect(result.complete).toBe(true);
  });

  it("filters by id in the form Microsoft documents, with only IDs from the role lists", async () => {
    const { requests } = await run({
      assignments: [assignment(1, id(1), USER_ADMIN)],
      users: [person(1)],
    });
    const lookup = requests.find((url) => url.pathname === "/v1.0/users");
    expect(lookup?.searchParams.get("$filter")).toBe(`id in ('${id(1)}')`);
    expect(lookup?.searchParams.get("$select")).toBe(
      "id,displayName,userPrincipalName,userType,accountEnabled",
    );
  });

  it("types an account as a user or a guest and links it to its page", async () => {
    const { rowsOf } = await run({
      assignments: [assignment(1, id(1), USER_ADMIN), assignment(2, id(2), USER_ADMIN)],
      users: [person(1), person(2, { userType: "Guest" })],
    });
    expect(rowsOf(id(1))[0]).toMatchObject({
      personId: id(1),
      cells: {
        holderType: "User",
        holder: { name: "Person 1", principal: "person1@example.invalid" },
        enabled: true,
      },
    });
    expect(rowsOf(id(2))[0]).toMatchObject({ cells: { holderType: "Guest" } });
  });

  it("shows every holder by ID, as Group or other, when Microsoft refuses the lookups", async () => {
    const { result } = await run(
      { assignments: [assignment(1, id(1), USER_ADMIN)], users: [person(1)] },
      {
        override: (url) =>
          url.pathname === "/v1.0/users" ? new Response("{}", { status: 403 }) : undefined,
      },
    );
    expect(result.rows[0]?.cells).toMatchObject({ holderType: "Group or other" });
    expect(result.complete).toBe(false);
    expect(result.notes.join("\n")).toContain("Microsoft refused to name some holders");
  });

  it("does not look anything up without the optional permissions, and says why", async () => {
    const { result, requests } = await run(
      { assignments: [assignment(1, id(1), USER_ADMIN)], users: [person(1)] },
      { granted: ["RoleManagement.Read.Directory"] },
    );
    expect(
      requests.some((url) => ["/v1.0/users", "/v1.0/servicePrincipals"].includes(url.pathname)),
    ).toBe(false);
    expect(result.rows[0]?.cells.holderType).toBe("Group or other");
    expect(result.complete).toBe(true);
    expect(result.notes.join("\n")).toContain("User.Read.All");
  });

  it("asks about a holder once however many roles it has", async () => {
    const { requests } = await run({
      assignments: [
        assignment(1, id(1), USER_ADMIN),
        assignment(2, id(1), EXCHANGE_ADMIN),
        assignment(3, id(1), GLOBAL_ADMIN),
      ],
      users: [person(1)],
    });
    const lookup = requests.find((url) => url.pathname === "/v1.0/users");
    expect(lookup?.searchParams.get("$filter")?.match(/'/g)).toHaveLength(2);
  });

  it("cleans a holder's name", async () => {
    const { result } = await run({
      assignments: [assignment(1, id(1), USER_ADMIN)],
      users: [person(1, { displayName: "Lone \uD800 surrogate\u0000" })],
    });
    expect(result.rows[0]?.cells.holder).toMatchObject({ name: "Lone � surrogate" });
  });
});

describe("which roles are privileged", () => {
  it("follows the shared role map by template ID, not by name", async () => {
    const { result } = await run({
      assignments: [
        assignment(1, id(1), EXCHANGE_ADMIN),
        assignment(2, id(1), HELPDESK),
        assignment(3, id(1), CUSTOM),
      ],
      // Microsoft renamed one built-in role; its ID still decides.
      definitions: [
        { id: EXCHANGE_ADMIN, displayName: "Mail Operator" },
        { id: HELPDESK, displayName: "Helpdesk Administrator" },
        { id: CUSTOM, displayName: "Lab Reports Reader" },
      ],
    });
    const byRole = Object.fromEntries(
      result.rows.map((row) => [row.cells.role, row.cells.privileged]),
    );
    expect(byRole).toEqual({
      "Mail Operator": true,
      "Helpdesk Administrator": false,
      "Lab Reports Reader": false,
    });
    expect(privilegedRoles.has(EXCHANGE_ADMIN)).toBe(true);
  });

  it("flags every role on the shared list and nothing else", async () => {
    const roles = [...privilegedRoles.keys(), HELPDESK, CUSTOM];
    const { result } = await run({
      assignments: roles.map((role, index) => assignment(index, id(1), role)),
    });
    expect(result.rows.filter((row) => row.cells.privileged)).toHaveLength(privilegedRoles.size);
    expect(result.rows).toHaveLength(roles.length);
  });

  it("names a role with no definition by its shared name, then by its ID", async () => {
    const unknown = "0d000000-0000-4000-8000-000000000009";
    const { result } = await run({
      assignments: [assignment(1, id(1), GLOBAL_ADMIN), assignment(2, id(1), unknown)],
      definitions: [],
    });
    expect(result.rows.map((row) => row.cells.role).sort()).toEqual([
      unknown,
      "Global Administrator",
    ]);
  });

  it("shows privileged roles first and hides the others in the default view", async () => {
    const definition = auditDefinition(privilegedRolesAudit);
    const { result } = await run({
      assignments: [assignment(1, id(1), HELPDESK), assignment(2, id(1), GLOBAL_ADMIN)],
    });
    expect(result.rows.map((row) => row.cells.privileged)).toEqual([true, false]);
    expect(definition.defaultView).toBe("privileged");
    expect(queryAuditRows(definition, result.rows)).toHaveLength(1);
    expect(queryAuditRows(definition, result.rows, { view: "all" })).toHaveLength(2);
  });
});

describe("the Global Administrator tile", () => {
  const withAdministrators = async (count: number, extra: Partial<Tenant> = {}) => {
    const holders = Array.from({ length: count }, (_, index) => id(200 + index));
    const { result } = await run({
      assignments: holders.map((holder, index) => assignment(index, holder, GLOBAL_ADMIN)),
      ...extra,
    });
    return tile(result, "Global Administrators");
  };

  it("warns above 4 and below 2, and not at 2, 3 or 4", async () => {
    const tones = await Promise.all(
      [0, 1, 2, 3, 4, 5, 6].map(async (count) => (await withAdministrators(count))?.tone),
    );
    expect(tones).toEqual(["warning", "warning", "good", "good", "good", "warning", "warning"]);
  });

  it("counts each account once, however it holds the role", async () => {
    const alice = id(1);
    const bob = id(2);
    const { result } = await run({
      assignments: [assignment(1, alice, GLOBAL_ADMIN), assignment(2, bob, GLOBAL_ADMIN)],
      active: [
        assigned(alice, GLOBAL_ADMIN, "a-1"),
        assigned(bob, GLOBAL_ADMIN, "a-2"),
        // Alice also activated the role through an eligibility, and has an inherited copy.
        instance(alice, GLOBAL_ADMIN, {
          assignmentType: "Activated",
          roleAssignmentOriginId: "a-9",
        }),
        assigned(alice, GLOBAL_ADMIN, "a-8", { memberType: "Group" }),
      ],
      eligible: [instance(alice, GLOBAL_ADMIN)],
    });
    expect(tile(result, "Global Administrators")?.value).toBe(2);
  });

  it("leaves eligible administrators out of the count and names them in the note", async () => {
    const { result } = await run({
      assignments: [assignment(1, id(1), GLOBAL_ADMIN), assignment(2, id(2), GLOBAL_ADMIN)],
      active: [assigned(id(1), GLOBAL_ADMIN, "a-1"), assigned(id(2), GLOBAL_ADMIN, "a-2")],
      eligible: [
        instance(id(1), GLOBAL_ADMIN),
        instance(id(3), GLOBAL_ADMIN),
        instance(id(4), GLOBAL_ADMIN),
      ],
    });
    expect(tile(result, "Global Administrators")).toMatchObject({ value: 2, tone: "good" });
    expect(result.notes.join("\n")).toContain(
      "2 accounts hold the role now, and 2 more can activate it",
    );
  });

  it("cites Microsoft's guidance and links to it", async () => {
    const { result } = await run({ assignments: [assignment(1, id(1), GLOBAL_ADMIN)] });
    const note = result.notes.find((item) => item.startsWith("Global Administrators:")) ?? "";
    expect(note).toContain("fewer than five Global Administrators");
    expect(note).toContain("at least two emergency access accounts");
    expect(note).toContain(
      "https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/best-practices",
    );
    expect(note).toContain("security-emergency-access");
  });
});

describe("the other tiles", () => {
  const alice = id(1);
  const bob = id(2);
  const gus = id(3);
  const dana = id(4);
  const app = id(5);
  const tenant: Tenant = {
    assignments: [
      assignment(1, alice, GLOBAL_ADMIN),
      assignment(2, alice, EXCHANGE_ADMIN),
      assignment(3, bob, GLOBAL_ADMIN),
      assignment(4, gus, USER_ADMIN),
      assignment(5, dana, EXCHANGE_ADMIN),
      assignment(6, app, EXCHANGE_ADMIN),
      assignment(7, dana, HELPDESK),
    ],
    active: [
      assigned(alice, GLOBAL_ADMIN, "a-1"),
      assigned(alice, EXCHANGE_ADMIN, "a-2", { endDateTime: inDays(5) }),
      assigned(bob, GLOBAL_ADMIN, "a-3"),
      assigned(gus, USER_ADMIN, "a-4"),
      assigned(dana, EXCHANGE_ADMIN, "a-5"),
      assigned(app, EXCHANGE_ADMIN, "a-6"),
      assigned(dana, HELPDESK, "a-7"),
    ],
    eligible: [instance(bob, USER_ADMIN)],
    users: [
      person(1),
      person(2),
      person(3, { userType: "Guest" }),
      person(4, { accountEnabled: false }),
    ],
    applications: [{ id: app, displayName: "Larkspur Backup", accountEnabled: true }],
  };

  it("counts holders, not rows", async () => {
    const { result } = await run(tenant);
    const values = Object.fromEntries(result.summary.map((item) => [item.label, item.value]));
    expect(values).toEqual({
      "Global Administrators": 2,
      // Alice, Bob, Gus, Dana and the app: Dana's Helpdesk role is not privileged.
      "Privileged role holders": 5,
      // Alice's Global Administrator and Bob's, Gus's, Dana's and the app's roles; Alice's
      // Exchange role has an end date.
      "Permanent privileged assignments": 5,
      "Disabled accounts with a privileged role": 1,
      "Guests with a privileged role": 1,
      "Service principals with a privileged role": 1,
    });
  });

  it("does not count a non-privileged role against anyone", async () => {
    const { result } = await run({
      assignments: [assignment(1, id(4), HELPDESK), assignment(2, id(3), CUSTOM)],
      users: [person(4, { accountEnabled: false }), person(3, { userType: "Guest" })],
    });
    for (const label of [
      "Privileged role holders",
      "Disabled accounts with a privileged role",
      "Guests with a privileged role",
    ])
      expect(tile(result, label)?.value, label).toBe(0);
  });

  it("opens views that hold the rows behind each tile", async () => {
    const definition = auditDefinition(privilegedRolesAudit);
    const { result } = await run(tenant);
    const views = new Set(definition.views.map((view) => view.id));
    for (const item of result.summary)
      if (item.view) expect(views.has(item.view), item.label).toBe(true);
    const rows = (view: string) => queryAuditRows(definition, result.rows, { view });
    expect(rows("disabled").map((row) => row.cells.holderType)).toEqual(["User"]);
    expect(rows("guests")).toHaveLength(1);
    expect(rows("service-principals")).toHaveLength(1);
    expect(rows("permanent")).toHaveLength(5);
    expect(rows("global-admins")).toHaveLength(2);
  });

  it("keeps eligible assignments out of the Global Administrator view but in the privileged one", async () => {
    const definition = auditDefinition(privilegedRolesAudit);
    const { result } = await run({
      assignments: [],
      eligible: [instance(id(1), GLOBAL_ADMIN)],
    });
    expect(queryAuditRows(definition, result.rows, { view: "global-admins" })).toHaveLength(0);
    expect(queryAuditRows(definition, result.rows, { view: "privileged" })).toHaveLength(1);
  });
});

describe("MFA registration", () => {
  const tenant: Tenant = {
    assignments: [
      assignment(1, id(1), GLOBAL_ADMIN),
      assignment(2, id(2), GLOBAL_ADMIN),
      assignment(3, id(5), GLOBAL_ADMIN),
      assignment(4, id(6), GLOBAL_ADMIN),
    ],
    users: [person(1), person(2)],
    applications: [{ id: id(5), displayName: "Larkspur Backup", accountEnabled: true }],
    registrations: [
      { id: id(1), isMfaRegistered: true },
      { id: id(2), isMfaRegistered: false },
    ],
  };
  const granted = [
    "RoleManagement.Read.Directory",
    "User.Read.All",
    "Application.Read.All",
    "AuditLog.Read.All",
  ];

  it("fills the column for accounts from the report, and leaves apps and groups empty", async () => {
    const { rowsOf } = await run(tenant, { granted });
    expect(rowsOf(id(1))[0]?.cells.mfa).toBe(true);
    expect(rowsOf(id(2))[0]?.cells.mfa).toBe(false);
    expect(rowsOf(id(5))[0]?.cells.mfa).toBeNull();
    expect(rowsOf(id(6))[0]?.cells.mfa).toBeNull();
  });

  it("reads nothing from the report without the optional permission", async () => {
    const { requests, result } = await run(tenant);
    expect(requests.some((url) => url.pathname.includes("authenticationMethods"))).toBe(false);
    expect(result.rows.every((row) => row.cells.mfa === null)).toBe(true);
    expect(result.complete).toBe(true);
  });

  it("does not read the report when no account holds a role", async () => {
    const { requests } = await run(
      { assignments: [assignment(1, id(6), GLOBAL_ADMIN)] },
      { granted },
    );
    expect(requests.some((url) => url.pathname.includes("authenticationMethods"))).toBe(false);
  });

  it("goes on when the report is refused, and says the result is partial", async () => {
    const { result, rowsOf } = await run(tenant, {
      granted,
      override: only("userRegistrationDetails", 403),
    });
    expect(rowsOf(id(1))[0]?.cells.mfa).toBeNull();
    expect(result.complete).toBe(false);
    expect(result.notes.join("\n")).toContain(
      "Microsoft refused the authentication methods report",
    );
  });
});

describe("what a run reads", () => {
  const tenant: Tenant = {
    assignments: [assignment(1, id(1), GLOBAL_ADMIN)],
    active: [assigned(id(1), GLOBAL_ADMIN, "a-1")],
    users: [person(1)],
  };

  it("only asks, and without a page size the role lists do not document", async () => {
    const { requests, methods } = await run(tenant);
    expect(new Set(methods)).toEqual(new Set(["GET"]));
    const lists = requests.filter((url) => url.pathname.includes("/roleManagement/"));
    expect(lists.map((url) => url.pathname.split("/").pop()).sort()).toEqual([
      "roleAssignmentScheduleInstances",
      "roleAssignments",
      "roleDefinitions",
      "roleEligibilityScheduleInstances",
    ]);
    for (const url of lists) {
      expect(url.searchParams.has("$top"), url.pathname).toBe(false);
      expect(url.searchParams.has("$select"), url.pathname).toBe(true);
    }
  });

  it("selects the fields it joins on", async () => {
    const { requests } = await run(tenant);
    const select = (name: string) =>
      requests
        .find((url) => url.pathname.endsWith(name))
        ?.searchParams.get("$select")
        ?.split(",");
    expect(select("/roleAssignmentScheduleInstances")).toEqual(
      expect.arrayContaining([
        "principalId",
        "roleDefinitionId",
        "directoryScopeId",
        "startDateTime",
        "endDateTime",
        "memberType",
        "assignmentType",
        "roleAssignmentOriginId",
      ]),
    );
    expect(select("/roleAssignments")).toEqual(
      expect.arrayContaining(["id", "principalId", "roleDefinitionId", "directoryScopeId"]),
    );
  });

  it("leaves a record it cannot read out and says the result is partial", async () => {
    const { result } = await run({
      assignments: [
        assignment(1, id(1), GLOBAL_ADMIN),
        assignment(2, "not-a-guid", GLOBAL_ADMIN),
        { id: "a-3" },
      ],
      active: [assigned(id(1), GLOBAL_ADMIN, "a-1", { endDateTime: "someday" })],
      definitions: [...definitions, { displayName: "No id" }],
    });
    // The instance and two assignments and one definition could not be read.
    expect(result.notes.join("\n")).toContain("4 records could not be read and are left out");
    expect(result.complete).toBe(false);
    expect(result.rows).toHaveLength(1);
    // The one assignment that can be read is listed, as active: its instance was unreadable.
    expect(result.rows[0]?.cells.held).toBe("Active");
  });

  it("never reads a malformed end date as a permanent assignment", async () => {
    const { result } = await run({
      assignments: [],
      active: [assigned(id(1), GLOBAL_ADMIN, "a-1", { endDateTime: "tomorrow-ish" })],
    });
    expect(result.rows).toHaveLength(0);
    expect(tile(result, "Permanent privileged assignments")?.value).toBe(0);
  });

  it("fails with Microsoft's refusal so the page can say what is needed", async () => {
    await expect(run({ assignments: [] }, { override: refused(403) })).rejects.toMatchObject({
      status: 403,
    });
    expect(privilegedRolesAudit.forbidden).toContain("RoleManagement.Read.Directory");
  });
});

describe("the published table", () => {
  const alice = id(1);
  const tenant: Tenant = {
    assignments: [assignment(1, alice, GLOBAL_ADMIN), assignment(2, id(2), HELPDESK)],
    active: [assigned(alice, GLOBAL_ADMIN, "a-1")],
    eligible: [instance(id(3), EXCHANGE_ADMIN, { endDateTime: inDays(20) })],
    users: [person(1), person(2), person(3)],
  };

  it("is registered, derives its permissions from its packs and takes no scope", () => {
    expect(builtInAudits.map((audit) => audit.definition.id)).toContain(
      "directory.privileged-roles",
    );
    const definition = auditDefinition(privilegedRolesAudit);
    expect(definition.permissions).toEqual(["RoleManagement.Read.Directory"]);
    expect(definition.optionalPermissions).toEqual([
      "User.Read.All",
      "Application.Read.All",
      "AuditLog.Read.All",
    ]);
    expect(definition.scoped).toBe(false);
    expect(definition.handoff).toBe(false);
  });

  it("has rows, views and a sort that agree with its columns", async () => {
    const definition = auditDefinition(privilegedRolesAudit);
    const { result } = await run(tenant);
    const keys = new Set(definition.columns.map((column) => column.key));
    expect(keys.has(definition.defaultSort.column)).toBe(true);
    expect(definition.views.map((view) => view.id)).toContain(definition.defaultView);
    for (const row of result.rows) {
      auditRowSchema.parse(row);
      for (const key of Object.keys(row.cells)) expect(keys.has(key), key).toBe(true);
    }
    expect(definition.columns.filter((column) => column.facet).map((column) => column.key)).toEqual(
      expect.arrayContaining(["holderType", "role", "held", "through"]),
    );
  });

  it("keeps every note inside the 500 characters the result stores", async () => {
    const full = await run(tenant, {
      granted: ["RoleManagement.Read.Directory", "AuditLog.Read.All"],
      override: only("ScheduleInstances", 403),
    });
    const plain = await run(tenant);
    for (const result of [full.result, plain.result])
      for (const note of result.notes) expect(note.length, note).toBeLessThanOrEqual(500);
  });
});

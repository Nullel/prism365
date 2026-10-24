import { queryAuditRows } from "@prism365/contracts";
import { GraphRequestError, GraphRetryDeferredError } from "@prism365/graph";
import { describe, expect, it } from "vitest";
import {
  appConsentsAudit,
  auditDefinition,
  HIGH_RISK_PERMISSIONS,
  highRiskIn,
  MAX_REQUESTS,
  MAX_SERVICE_PRINCIPALS,
  MICROSOFT_TENANTS,
  ownerOf,
} from "../src/index.ts";
import { expectValidTable, fake, find, paged, type Route } from "./apps-devices-usage-support.ts";

// A synthetic tenant: invented applications and publishers, made-up IDs. The two Microsoft tenant
// IDs are the ones Microsoft documents as owning its first-party apps.
const TENANT = "6f1c2d8e-4b7a-4e35-9a0c-3d5b7e91c204";
const MICROSOFT = "f8cdef31-a31e-4b4a-93e4-5f571e91255a";
const OTHER = "2b3c4d5e-0000-4000-8000-0000000000aa";
const ZERO = "00000000-0000-0000-0000-000000000000";

const role = (n: number) => `a1b2c3d4-0000-4000-8000-${String(n).padStart(12, "0")}`;
const ROLE = {
  mailRead: role(1),
  userReadAll: role(2),
  sitesReadWrite: role(3),
  fullAccess: role(4),
  unknown: role(99),
};

const resources: Record<string, unknown> = {
  "sp-graph": {
    id: "sp-graph",
    displayName: "Microsoft Graph",
    appRoles: [
      { id: ROLE.mailRead, value: "Mail.Read", displayName: "Read mail in all mailboxes" },
      { id: ROLE.userReadAll, value: "User.Read.All", displayName: "Read all users" },
      { id: ROLE.sitesReadWrite, value: "Sites.ReadWrite.All", displayName: "Edit sites" },
      // A role without a value (assignable to users, not an API permission) cannot name anything.
      { id: role(50), value: null, displayName: "Reader" },
      "not a role",
    ],
  },
  "sp-exo": {
    id: "sp-exo",
    displayName: "Office 365 Exchange Online",
    appRoles: [{ id: ROLE.fullAccess, value: "full_access_as_app", displayName: "Use EWS" }],
  },
};

const app = (
  id: string,
  displayName: string,
  extra: Record<string, unknown> = {},
): Record<string, unknown> => ({
  id,
  appId: `appid-${id}`,
  displayName,
  accountEnabled: true,
  servicePrincipalType: "Application",
  appOwnerOrganizationId: OTHER,
  verifiedPublisher: null,
  // Fields the audit does not ask for are tolerated.
  appRoles: [],
  tags: ["x"],
  ...extra,
});

const principals = [
  app("sp-graph", "Microsoft Graph", { appOwnerOrganizationId: MICROSOFT }),
  app("sp-exo", "Office 365 Exchange Online", { appOwnerOrganizationId: MICROSOFT }),
  app("app-quill", "Quillnet Mail Archiver"),
  app("app-orrery", "Orrery Analytics", {
    verifiedPublisher: { displayName: "Orrery Software Ltd", verifiedPublisherId: "1234567" },
  }),
  app("app-tess", "Tessellate Planner", { verifiedPublisher: { displayName: "Tessellate Labs" } }),
  app("app-pin", "Pinwheel Forms"),
  app("app-internal", "Larkspur Lab Portal", { appOwnerOrganizationId: TENANT }),
  app("app-cli", "Graph Command Line Tools", { appOwnerOrganizationId: MICROSOFT }),
  app("app-msidle", "Office 365 Portal", { appOwnerOrganizationId: MICROSOFT }),
  app("app-empty", "Driftwood Notes Sync"),
  app("app-off", "Fernwood Timesheets", {
    accountEnabled: false,
    verifiedPublisher: { displayName: "Fernwood Ltd" },
  }),
  app("app-msapp", "Exchange Hybrid Agent", { appOwnerOrganizationId: MICROSOFT }),
  app("app-mi", "pipeline identity", { servicePrincipalType: "ManagedIdentity" }),
];

const assign = (resourceId: string, appRoleId: string) => ({
  id: `a-${resourceId}-${appRoleId}`,
  resourceId,
  appRoleId,
  principalType: "ServicePrincipal",
});
const assignments: Record<string, unknown[]> = {
  "app-quill": [assign("sp-graph", ROLE.mailRead), assign("sp-graph", ROLE.userReadAll)],
  "app-internal": [
    // Microsoft writes IDs in either case.
    assign("sp-graph", ROLE.sitesReadWrite.toUpperCase()),
    assign("sp-graph", ROLE.unknown),
    assign("sp-graph", ZERO),
  ],
  "app-msapp": [assign("sp-exo", ROLE.fullAccess)],
  "app-mi": [assign("sp-graph", ROLE.mailRead)],
};

const grant = (
  clientId: string,
  consentType: string,
  scope: string,
  principalId: string | null = null,
) => ({ id: `g-${clientId}-${principalId ?? "all"}`, clientId, consentType, principalId, scope });
const grants = [
  grant("app-orrery", "AllPrincipals", "User.Read Calendars.Read"),
  grant("app-orrery", "Principal", "User.Read", "user-1"),
  grant("app-orrery", "Principal", "User.Read", "user-2"),
  grant("app-tess", "AllPrincipals", "Files.ReadWrite.All offline_access"),
  grant("app-pin", "Principal", "User.Read mail.send", "user-1"),
  grant("app-pin", "Principal", "User.Read", "user-1"),
  grant("app-pin", "Principal", "User.Read Mail.Send", "user-2"),
  grant("app-pin", "Principal", "User.Read", "user-3"),
  grant("app-cli", "AllPrincipals", "Directory.Read.All"),
  grant("app-off", "AllPrincipals", "User.Read"),
  // Everything was revoked: no consent left.
  grant("app-empty", "AllPrincipals", " "),
];

interface Data {
  principals: unknown[];
  assignments: Record<string, unknown[]>;
  grants: unknown[];
  resources: Record<string, unknown>;
}
const data: Data = { principals, assignments, grants, resources };

function tenantRoute(from: Data = data): Route {
  return (url) => {
    const path = url.pathname.replace("/v1.0/", "");
    if (path === "servicePrincipals") return paged(url, from.principals);
    const assigned = /^servicePrincipals\/([^/]+)\/appRoleAssignments$/.exec(path);
    if (assigned) return paged(url, from.assignments[assigned[1] ?? ""] ?? []);
    const one = /^servicePrincipals\/([^/]+)$/.exec(path);
    if (one) {
      const resource = from.resources[one[1] ?? ""];
      return resource ? Response.json(resource) : undefined;
    }
    if (path === "oauth2PermissionGrants") return paged(url, from.grants);
    return undefined;
  };
}

const granted = ["Application.Read.All", "Directory.Read.All"];
const withoutGrants = ["Application.Read.All"];
const run = (route: Route = tenantRoute(), grantedScopes: string[] = granted, tenantId = TENANT) =>
  appConsentsAudit.run(fake(appConsentsAudit, route, { granted: grantedScopes, tenantId }).context);
const note = (result: { notes: string[] }) => result.notes.join("\n");

describe("app consents", () => {
  it("publishes a valid table", async () => {
    await expectValidTable(appConsentsAudit, tenantRoute(), { granted, tenantId: TENANT });
  });

  it("declares the application pack, with delegated grants as an optional pack", () => {
    expect(auditDefinition(appConsentsAudit)).toMatchObject({
      id: "apps.consents",
      category: "Applications",
      permissions: ["Application.Read.All"],
      optionalPermissions: ["Directory.Read.All"],
      scoped: false,
    });
  });

  it("lists the applications that hold any permission and skips the rest", async () => {
    const result = await run();
    expect(result.rows.map((row) => row.key).sort()).toEqual([
      "app-cli",
      "app-internal",
      "app-msapp",
      "app-off",
      "app-orrery",
      "app-pin",
      "app-quill",
      "app-tess",
    ]);
    // Nothing for an app with no permission, a managed identity, or a grant whose scope is blank.
    for (const key of ["app-empty", "app-msidle", "app-mi", "sp-graph", "sp-exo"])
      expect(
        result.rows.some((row) => row.key === key),
        key,
      ).toBe(false);
    for (const row of result.rows) expect(row.personId).toBeNull();
  });

  it("names application permissions from the resource's app roles", async () => {
    const result = await run();
    expect(find(result.rows, "app-quill").cells.appPermissions).toEqual([
      "Mail.Read",
      "User.Read.All",
    ]);
    expect(find(result.rows, "app-msapp").cells.appPermissions).toEqual(["full_access_as_app"]);
    // The ID is matched whatever its letter case, and "no specific role" is not a permission.
    expect(find(result.rows, "app-internal").cells.appPermissions).toEqual([
      ROLE.unknown,
      "Sites.ReadWrite.All",
    ]);
  });

  it("shows a permission ID it cannot name as the ID, says so, and marks the result partial", async () => {
    const result = await run();
    expect(find(result.rows, "app-internal").cells.appPermissions).toContain(ROLE.unknown);
    expect(note(result)).toContain("1 application permission could not be matched to a name");
    expect(note(result)).toContain("shown as an ID");
    expect(result.complete).toBe(false);
  });

  it("reads each resource once, however many applications use it", async () => {
    const ctx = fake(appConsentsAudit, tenantRoute(), { granted, tenantId: TENANT });
    await appConsentsAudit.run(ctx.context);
    const lookups = ctx.requests.filter(
      (url) => url.pathname === "/v1.0/servicePrincipals/sp-graph",
    );
    expect(lookups).toHaveLength(1);
    expect(lookups[0]?.searchParams.get("$select")).toBe("id,displayName,appRoles");
    expect(
      ctx.requests.filter((url) => url.pathname === "/v1.0/servicePrincipals/sp-exo"),
    ).toHaveLength(1);
    // The list asks for the fields it needs, not each principal's role and scope lists.
    const list = ctx.requests.find((url) => url.pathname === "/v1.0/servicePrincipals");
    expect(list?.searchParams.get("$top")).toBe("100");
    expect(list?.searchParams.get("$select")).not.toContain("appRoles");
  });

  it("reads one request at a time, third-party apps before Microsoft's", async () => {
    const ctx = fake(appConsentsAudit, tenantRoute(), { granted, tenantId: TENANT });
    await appConsentsAudit.run(ctx.context);
    const order = ctx.requests
      .map(
        (url) =>
          /^\/v1\.0\/servicePrincipals\/([^/]+)\/appRoleAssignments$/.exec(url.pathname)?.[1],
      )
      .filter((id): id is string => Boolean(id));
    // Application-type principals only: the managed identity is never read.
    expect(order).not.toContain("app-mi");
    const firstMicrosoft = order.findIndex((id) =>
      ["sp-graph", "sp-exo", "app-cli", "app-msidle", "app-msapp"].includes(id),
    );
    expect(order.slice(0, firstMicrosoft)).toEqual([
      "app-quill",
      "app-orrery",
      "app-tess",
      "app-pin",
      "app-internal",
      "app-empty",
      "app-off",
    ]);
    expect(order).toHaveLength(12);
  });

  it("says who owns each application", async () => {
    const result = await run();
    const owner = (key: string) => find(result.rows, key).cells.owner;
    expect(owner("app-cli")).toBe("Microsoft");
    expect(owner("app-internal")).toBe("This tenant");
    expect(owner("app-quill")).toBe("Another organisation");
    expect(find(result.rows, "app-cli").tags).not.toContain("other-org");
    expect(find(result.rows, "app-quill").tags).toContain("other-org");
    for (const id of MICROSOFT_TENANTS) expect(ownerOf(id.toUpperCase(), TENANT)).toBe("Microsoft");
    expect(ownerOf("72f988bf-86f1-41af-91ab-2d7cd011db47", TENANT)).toBe("Microsoft");
    expect(ownerOf(null, TENANT)).toBe("Unknown");
  });

  it("cannot tell this tenant's apps from others without the tenant ID, and says Unknown", async () => {
    const result = await run(tenantRoute(), granted, "");
    expect(find(result.rows, "app-internal").cells.owner).toBe("Unknown");
    expect(find(result.rows, "app-quill").cells.owner).toBe("Unknown");
    // Microsoft's own apps are still recognised by their documented tenant IDs.
    expect(find(result.rows, "app-cli").cells.owner).toBe("Microsoft");
    expect(ownerOf(OTHER, undefined)).toBe("Unknown");
  });

  it("names the verified publisher, and leaves Microsoft's own apps and this tenant's out of unverified", async () => {
    const result = await run();
    const publisher = (key: string) => find(result.rows, key).cells.publisher;
    expect(publisher("app-orrery")).toBe("Orrery Software Ltd");
    expect(publisher("app-quill")).toBe("Unverified");
    expect(publisher("app-cli")).toBe("Microsoft");
    const unverified = result.rows.filter((row) => row.tags.includes("unverified"));
    expect(unverified.map((row) => row.key).sort()).toEqual(["app-pin", "app-quill"]);
    expect(find(result.rows, "app-internal").tags).not.toContain("unverified");
    expect(find(result.rows, "app-cli").tags).not.toContain("unverified");
  });

  it("keeps delegated permissions for all users apart from what individual users consented to", async () => {
    const result = await run();
    expect(find(result.rows, "app-orrery").cells).toMatchObject({
      delegated: ["Calendars.Read", "User.Read"],
      consentedUsers: 2,
      userScopes: ["User.Read"],
    });
    // One user consenting twice is one user.
    expect(find(result.rows, "app-pin").cells).toMatchObject({
      delegated: [],
      consentedUsers: 3,
      userScopes: ["mail.send", "Mail.Send", "User.Read"],
    });
    expect(find(result.rows, "app-tess").cells).toMatchObject({
      delegated: ["Files.ReadWrite.All", "offline_access"],
      consentedUsers: 0,
    });
    expect(find(result.rows, "app-off").cells.enabled).toBe(false);
  });

  it("marks high-risk permissions, application or delegated, whatever their case", async () => {
    const result = await run();
    const risk = (key: string) => find(result.rows, key).cells.highRisk;
    expect(risk("app-quill")).toEqual(["Mail.Read"]);
    expect(risk("app-internal")).toEqual(["Sites.ReadWrite.All"]);
    expect(risk("app-msapp")).toEqual(["full_access_as_app"]);
    // Admin consent for everyone, and a permission users consented to themselves.
    expect(risk("app-tess")).toEqual(["Files.ReadWrite.All"]);
    expect(risk("app-pin")).toEqual(["Mail.Send"]);
    for (const key of ["app-orrery", "app-cli", "app-off"]) expect(risk(key)).toEqual([]);
    expect(find(result.rows, "app-pin").tags).toContain("high-risk");
    expect(find(result.rows, "app-pin").tags).not.toContain("app-high-risk");
    expect(find(result.rows, "app-quill").tags).toEqual(
      expect.arrayContaining(["high-risk", "app-high-risk", "app-permissions", "unverified"]),
    );
  });

  it("keeps the documented high-risk list, matched without regard to case", () => {
    for (const name of [
      "Mail.ReadWrite",
      "Mail.Read",
      "Mail.Send",
      "Files.ReadWrite.All",
      "Sites.ReadWrite.All",
      "Directory.ReadWrite.All",
      "RoleManagement.ReadWrite.Directory",
      "AppRoleAssignment.ReadWrite.All",
      "Application.ReadWrite.All",
      "User.ReadWrite.All",
      "full_access_as_app",
      "EWS.AccessAsUser.All",
      "Mail.ReadBasic.All",
      "Chat.Read.All",
    ])
      expect(HIGH_RISK_PERMISSIONS).toContain(name);
    expect(highRiskIn(["mail.send", "MAIL.SEND", "User.Read", "Chat.Read.All"])).toEqual([
      "Chat.Read.All",
      "Mail.Send",
    ]);
    expect(highRiskIn([])).toEqual([]);
  });

  it("opens each view on the applications that belong in it", async () => {
    const result = await run();
    const definition = auditDefinition(appConsentsAudit);
    const keys = (view: string) =>
      queryAuditRows(definition, result.rows, { view })
        .map((row) => row.key)
        .sort();
    expect(keys("high-risk")).toEqual([
      "app-internal",
      "app-msapp",
      "app-pin",
      "app-quill",
      "app-tess",
    ]);
    expect(keys("app-high-risk")).toEqual(["app-internal", "app-msapp", "app-quill"]);
    expect(keys("unverified-high-risk")).toEqual(["app-pin", "app-quill"]);
    expect(keys("app-permissions")).toEqual(["app-internal", "app-msapp", "app-quill"]);
    expect(keys("admin-consent")).toEqual(["app-cli", "app-off", "app-orrery", "app-tess"]);
    expect(keys("user-consent")).toEqual(["app-orrery", "app-pin"]);
    expect(keys("unverified")).toEqual(["app-pin", "app-quill"]);
    expect(keys("other-org")).toEqual([
      "app-off",
      "app-orrery",
      "app-pin",
      "app-quill",
      "app-tess",
    ]);
    expect(keys("all")).toHaveLength(8);
    expect(definition.defaultView).toBe("high-risk");
  });

  it("summarises with the tiles in the contract's order and severity", async () => {
    const result = await run();
    expect(result.summary).toEqual([
      {
        label: "Unverified apps with high-risk permissions",
        value: 2,
        tone: "critical",
        view: "unverified-high-risk",
      },
      {
        label: "Apps with high-risk application permissions",
        value: 3,
        tone: "warning",
        view: "app-high-risk",
      },
      {
        label: "Apps with admin consent for everyone",
        value: 4,
        tone: "warning",
        view: "admin-consent",
      },
      {
        label: "Apps consented by individual users",
        value: 2,
        tone: "neutral",
        view: "user-consent",
      },
    ]);
  });

  it("shows no finding as good", async () => {
    const result = await run(
      tenantRoute({
        principals: [app("sp-graph", "Microsoft Graph", { appOwnerOrganizationId: MICROSOFT })],
        assignments: {},
        grants: [],
        resources,
      }),
    );
    expect(result.rows).toEqual([]);
    expect(result.summary.map((tile) => tile.tone)).toEqual(["good", "good", "good", "neutral"]);
  });

  it("explains how Microsoft's own apps, replication, high-risk and unverified are treated", async () => {
    const text = note(await run());
    expect(text).toContain("Microsoft's own apps are left out unless they hold permissions");
    expect(text).toContain("takes time to replicate");
    expect(text).toContain("High-risk means the permission is on Prism365's short list");
    for (const name of HIGH_RISK_PERMISSIONS) expect(text).toContain(name);
    expect(text).toContain("Unverified means Microsoft has not verified");
  });

  describe("without the optional pack", () => {
    it("never asks for delegated grants, says which permission adds them, and is partial", async () => {
      const ctx = fake(appConsentsAudit, tenantRoute(), {
        granted: withoutGrants,
        tenantId: TENANT,
      });
      const result = await appConsentsAudit.run(ctx.context);
      expect(ctx.requests.some((url) => url.pathname === "/v1.0/oauth2PermissionGrants")).toBe(
        false,
      );
      expect(note(result)).toContain("Directory.Read.All");
      expect(note(result)).toContain("Only application permissions are listed");
      expect(result.complete).toBe(false);
      // Only the apps with application permissions are rows, and delegated columns are not read.
      expect(result.rows.map((row) => row.key).sort()).toEqual([
        "app-internal",
        "app-msapp",
        "app-quill",
      ]);
      expect(find(result.rows, "app-quill").cells).toMatchObject({
        delegated: null,
        consentedUsers: null,
        userScopes: null,
      });
      // A count of zero admin consents says nothing when they were not read.
      expect(result.summary[2]).toMatchObject({ value: 0, tone: "neutral" });
    });

    it("carries on when Microsoft refuses the delegated grants, and says so", async () => {
      const refusing: Route = (url) =>
        url.pathname === "/v1.0/oauth2PermissionGrants"
          ? new Response("{}", { status: 403 })
          : tenantRoute()(url);
      const result = await run(refusing);
      expect(note(result)).toContain("Microsoft refused to list delegated permission grants");
      expect(result.rows.map((row) => row.key)).toContain("app-quill");
      expect(find(result.rows, "app-quill").cells.delegated).toBeNull();
      expect(result.complete).toBe(false);
    });

    it("still fails the run when delegated grants hit throttling", async () => {
      const throttled: Route = (url) =>
        url.pathname === "/v1.0/oauth2PermissionGrants"
          ? new Response("{}", { status: 429 })
          : tenantRoute()(url);
      await expect(run(throttled)).rejects.toBeInstanceOf(GraphRetryDeferredError);
    });
  });

  describe("when Microsoft refuses", () => {
    it("fails with the refusal when the list is refused", async () => {
      const failure = run(() => new Response("{}", { status: 403 }));
      await expect(failure).rejects.toBeInstanceOf(GraphRequestError);
      await expect(failure).rejects.toMatchObject({ status: 403 });
    });

    it("fails with the refusal when app role assignments are refused", async () => {
      const route: Route = (url) =>
        url.pathname.endsWith("/appRoleAssignments")
          ? new Response("{}", { status: 403 })
          : tenantRoute()(url);
      await expect(run(route)).rejects.toMatchObject({ status: 403 });
    });

    it("tells the administrator which roles read app role assignments", () => {
      const text = appConsentsAudit.forbidden;
      expect(text).toContain("Application.Read.All");
      expect(text).toContain("Directory Readers");
      expect(text).toContain("Application Administrator");
      expect(text).toContain("Global Reader");
    });
  });

  describe("unreadable records", () => {
    const messy: Data = {
      principals: [
        ...principals,
        { displayName: "no id", servicePrincipalType: "Application" },
        { id: "no-type", displayName: "no type" },
        "not an object",
      ],
      assignments: {
        ...assignments,
        "app-quill": [
          ...(assignments["app-quill"] ?? []),
          { id: "broken", appRoleId: ROLE.mailRead },
        ],
      },
      grants: [...grants, { id: "no client" }, grant("app-orrery", "Mystery", "User.Read")],
      resources,
    };

    // Everything here resolves to a name, so a result is partial only for the record that is
    // added to it.
    const clean: Data = {
      principals,
      assignments: { "app-quill": assignments["app-quill"] ?? [] },
      grants,
      resources,
    };

    it("is complete when every record can be read", async () => {
      const result = await run(tenantRoute(clean));
      expect(result.complete).toBe(true);
      expect(note(result)).not.toContain("could not be read");
    });

    it("does not call a record with no type unreadable, and does not examine it", async () => {
      const ctx = fake(
        appConsentsAudit,
        tenantRoute({
          ...clean,
          principals: [...principals, { id: "no-type", displayName: "no type" }],
        }),
        {
          granted,
          tenantId: TENANT,
        },
      );
      const result = await appConsentsAudit.run(ctx.context);
      expect(result.complete).toBe(true);
      expect(
        ctx.requests.some(
          (url) => url.pathname === "/v1.0/servicePrincipals/no-type/appRoleAssignments",
        ),
      ).toBe(false);
    });

    it("is partial for a service principal alone that cannot be read", async () => {
      const result = await run(
        tenantRoute({ ...clean, principals: [...principals, { displayName: "no id" }] }),
      );
      expect(result.complete).toBe(false);
      expect(note(result)).toContain("1 service principal could not be read and is left out");
    });

    it("is partial for an assignment alone that cannot be read", async () => {
      const result = await run(
        tenantRoute({
          ...clean,
          assignments: { "app-quill": [...(assignments["app-quill"] ?? []), { id: "broken" }] },
        }),
      );
      expect(result.complete).toBe(false);
      expect(note(result)).toContain("1 permission record could not be read and is left out");
    });

    it("is partial for a grant alone that cannot be read", async () => {
      const result = await run(tenantRoute({ ...clean, grants: [...grants, { id: "no client" }] }));
      expect(result.complete).toBe(false);
      expect(note(result)).toContain("1 permission record could not be read and is left out");
    });

    it("skips and counts them, and marks the result partial", async () => {
      const result = await run(tenantRoute(messy));
      expect(result.complete).toBe(false);
      expect(note(result)).toContain("2 service principals could not be read and are left out");
      expect(note(result)).toContain("3 permission records could not be read and are left out");
      // The readable ones are still there.
      expect(result.rows).toHaveLength(8);
    });
  });

  describe("limits", () => {
    const crowd = (count: number, extra: (index: number) => Record<string, unknown> = () => ({})) =>
      Array.from({ length: count }, (_, index) =>
        app(`crowd-${index}`, `Crowd app ${index}`, extra(index)),
      );

    it("stops at 600 requests, third-party apps first, and says how far it got", async () => {
      const apps = crowd(700);
      const withRoles: Data = {
        principals: apps,
        assignments: {
          "crowd-0": [assign("sp-graph", ROLE.mailRead)],
          "crowd-1": [assign("sp-graph", ROLE.mailRead)],
        },
        grants: [],
        resources,
      };
      const ctx = fake(appConsentsAudit, tenantRoute(withRoles), { granted, tenantId: TENANT });
      const result = await appConsentsAudit.run(ctx.context);
      expect(MAX_REQUESTS).toBe(600);
      expect(ctx.requests.length).toBeLessThanOrEqual(MAX_REQUESTS);
      // Seven pages of principals, one page of grants, then applications until 40 are left.
      expect(ctx.requests.length).toBe(MAX_REQUESTS - 40 + 1);
      expect(result.complete).toBe(false);
      expect(note(result)).toContain("of 700 applications");
      expect(note(result)).toContain("limit of 600 requests");
      expect(note(result)).toContain("Microsoft's own apps are read last");
      // The requests kept back still turn the permission IDs into names.
      expect(find(result.rows, "crowd-0").cells.appPermissions).toEqual(["Mail.Read"]);
      expect(note(result)).not.toContain("could not be matched");
    });

    it("reads Microsoft's own apps last when the budget runs out", async () => {
      const apps = [
        ...crowd(100, () => ({ appOwnerOrganizationId: MICROSOFT })),
        ...crowd(560).map((item, index) => ({ ...item, id: `third-${index}` })),
      ];
      const ctx = fake(
        appConsentsAudit,
        tenantRoute({ principals: apps, assignments: {}, grants: [], resources }),
        { granted: withoutGrants, tenantId: TENANT },
      );
      await appConsentsAudit.run(ctx.context);
      const read = ctx.requests
        .map((url) => /servicePrincipals\/([^/]+)\/appRoleAssignments/.exec(url.pathname)?.[1])
        .filter((id): id is string => Boolean(id));
      expect(read.every((id) => id.startsWith("third-"))).toBe(true);
      expect(read.some((id) => id.startsWith("crowd-"))).toBe(false);
    });

    it("stops at 2,000 service principals", async () => {
      const identities = Array.from({ length: MAX_SERVICE_PRINCIPALS + 100 }, (_, index) =>
        app(`mi-${index}`, `identity ${index}`, { servicePrincipalType: "ManagedIdentity" }),
      );
      const ctx = fake(
        appConsentsAudit,
        tenantRoute({
          principals: [...principals.slice(0, 5), ...identities],
          assignments,
          grants: [],
          resources,
        }),
        { granted, tenantId: TENANT },
      );
      const result = await appConsentsAudit.run(ctx.context);
      expect(ctx.requests.filter((url) => url.pathname === "/v1.0/servicePrincipals")).toHaveLength(
        20,
      );
      expect(result.complete).toBe(false);
      expect(note(result)).toContain("Stopped after 2,000 service principals");
      expect(find(result.rows, "app-quill").cells.appPermissions).toContain("Mail.Read");
    });

    it("keeps what it read when Microsoft asks it to slow down", async () => {
      const route: Route = (url) =>
        url.pathname === "/v1.0/servicePrincipals/app-tess/appRoleAssignments"
          ? new Response("{}", { status: 429 })
          : tenantRoute()(url);
      const result = await run(route);
      expect(result.complete).toBe(false);
      expect(note(result)).toContain(
        "slow down after the application permissions of 2 of 12 applications",
      );
      expect(find(result.rows, "app-quill").cells.appPermissions).toEqual([
        "Mail.Read",
        "User.Read.All",
      ]);
    });

    it("fails when Microsoft throttles the very first read", async () => {
      await expect(run(() => new Response("{}", { status: 429 }))).rejects.toBeInstanceOf(
        GraphRetryDeferredError,
      );
    });

    it("tolerates an application deleted after it was listed", async () => {
      const route: Route = (url) =>
        url.pathname === "/v1.0/servicePrincipals/app-quill/appRoleAssignments"
          ? new Response("{}", { status: 404 })
          : tenantRoute()(url);
      const result = await run(route);
      expect(result.rows.some((row) => row.key === "app-quill")).toBe(false);
      expect(result.rows).toHaveLength(7);
    });

    it("cuts a very long permission list but still checks all of it for high-risk names", async () => {
      const roles = [
        ...Array.from({ length: 249 }, (_, index) => ({
          id: role(1_000 + index),
          value: `A.Perm.${String(index).padStart(3, "0")}`,
        })),
        { id: role(2_000), value: "Mail.Read" },
      ];
      const result = await run(
        tenantRoute({
          principals: [app("app-big", "Big app")],
          assignments: { "app-big": roles.map((item) => assign("sp-big", item.id)) },
          grants: [],
          resources: { "sp-big": { id: "sp-big", displayName: "Big", appRoles: roles } },
        }),
      );
      const cells = find(result.rows, "app-big").cells;
      const list = cells.appPermissions as string[];
      expect(list).toHaveLength(200);
      expect(list.at(-1)).toBe("and 51 more");
      expect(list).not.toContain("Mail.Read");
      expect(cells.highRisk).toEqual(["Mail.Read"]);
    });
  });
});

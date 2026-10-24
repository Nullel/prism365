import { describe, expect, it, vi } from "vitest";
import { z } from "zod";
import { GraphHttpClient, GraphRequestError, GraphRetryDeferredError } from "../src/client.ts";

const context = {
  feature: "audit.signins.read",
  permission: "AuditLog.Read.All",
} as const;
const userContext = { feature: "directory.users.read", permission: "User.Read.All" } as const;

describe("read packs added for the change feed and the consent-batch audits", () => {
  const client = () =>
    new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation: async () => Response.json({ value: [] }),
    });
  const allowed = async (path: string, feature: string, permission: string) =>
    client()
      .request(path, { feature, permission } as never)
      .then(
        () => true,
        (error: unknown) => {
          if (error instanceof Error && /does not belong/.test(error.message)) return false;
          throw error;
        },
      );

  it("keeps the directory audit log out of the sign-in pack", async () => {
    expect(
      await allowed("auditLogs/directoryAudits", "audit.directory.read", "AuditLog.Read.All"),
    ).toBe(true);
    expect(
      await allowed("auditLogs/directoryAudits", "audit.signins.read", "AuditLog.Read.All"),
    ).toBe(false);
    expect(await allowed("auditLogs/signIns", "audit.directory.read", "AuditLog.Read.All")).toBe(
      false,
    );
  });

  it("lets the application packs read applications, their app role assignments and grants only", async () => {
    const apps = ["directory.applications.read", "Application.Read.All"] as const;
    expect(await allowed("servicePrincipals", ...apps)).toBe(true);
    expect(await allowed("servicePrincipals/sp-1/appRoleAssignments", ...apps)).toBe(true);
    expect(await allowed("servicePrincipals/sp-1/appRoleAssignedTo", ...apps)).toBe(false);
    expect(await allowed("servicePrincipals/sp-1/owners", ...apps)).toBe(false);
    expect(await allowed("applications", ...apps)).toBe(false);
    const grants = ["directory.permission-grants.read", "Directory.Read.All"] as const;
    expect(await allowed("oauth2PermissionGrants", ...grants)).toBe(true);
    expect(await allowed("users", ...grants)).toBe(false);
  });

  it("reads Conditional Access policies and named locations only", async () => {
    const pack = ["identity.conditional-access.read", "Policy.Read.All"] as const;
    expect(await allowed("identity/conditionalAccess/policies", ...pack)).toBe(true);
    expect(await allowed("identity/conditionalAccess/namedLocations", ...pack)).toBe(true);
    expect(
      await allowed("identity/conditionalAccess/authenticationStrength/policies", ...pack),
    ).toBe(false);
  });

  it("refuses $expand in any spelling, so navigations a pack leaves out stay unreachable", async () => {
    const apps = ["directory.applications.read", "Application.Read.All"] as const;
    const refused = async (path: string) =>
      client()
        .request(path, { feature: apps[0], permission: apps[1] })
        .then(
          () => false,
          (error: unknown) => error instanceof Error && /\$expand/.test(error.message),
        );
    expect(await refused("servicePrincipals/sp-1?$expand=owners")).toBe(true);
    expect(await refused("servicePrincipals/sp-1?%24expand=owners")).toBe(true);
    expect(await refused("servicePrincipals/sp-1?$EXPAND=owners")).toBe(true);
    expect(await refused("servicePrincipals/sp-1?expand=owners")).toBe(true);
    expect(await refused("servicePrincipals/sp-1?$select=id")).toBe(false);
  });

  it("keeps sibling navigations, deeper paths and encoded separators out of the new packs", async () => {
    const apps = ["directory.applications.read", "Application.Read.All"] as const;
    for (const path of [
      "servicePrincipals/sp-1/oauth2PermissionGrants",
      "servicePrincipals/sp-1/memberOf",
      "servicePrincipals/sp-1/transitiveMemberOf",
      "servicePrincipals/sp-1/appRoleAssignments/extra",
      "servicePrincipals/x%2Fowners",
    ])
      expect(await allowed(path, ...apps)).toBe(false);
    const audit = ["audit.directory.read", "AuditLog.Read.All"] as const;
    expect(await allowed("auditLogs/provisioning", ...audit)).toBe(false);
    expect(await allowed("auditLogs/directoryAudits/a/b", ...audit)).toBe(false);
    const grants = ["directory.permission-grants.read", "Directory.Read.All"] as const;
    expect(await allowed("oauth2PermissionGrants/grant-1", ...grants)).toBe(false);
    expect(await allowed("x/oauth2PermissionGrants", ...grants)).toBe(false);
    const ca = ["identity.conditional-access.read", "Policy.Read.All"] as const;
    expect(await allowed("identity/conditionalAccess/policies/p-1/extra", ...ca)).toBe(false);
    const usage = ["reports.usage.read", "Reports.Read.All"] as const;
    expect(await allowed("reports/getM365AppUserDetail(period='D90')/extra", ...usage)).toBe(false);
    expect(await allowed("reports/getM365AppUserCounts(period='D90')", ...usage)).toBe(false);
  });

  it("allows the usage reports by period only; report.test.ts covers the full list", async () => {
    const pack = ["reports.usage.read", "Reports.Read.All"] as const;
    expect(await allowed("reports/getM365AppUserDetail(period='D90')", ...pack)).toBe(true);
    expect(await allowed("reports/getOffice365ActiveUserDetail(period='D90')", ...pack)).toBe(true);
    expect(await allowed("reports/getM365AppUserDetail(period='D365')", ...pack)).toBe(false);
    expect(await allowed("reports/getOffice365ActiveUserDetail(date=2026-10-01)", ...pack)).toBe(
      false,
    );
  });
});

describe("read packs for the roles, groups and password audits", () => {
  const client = () =>
    new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation: async () => Response.json({ value: [] }),
    });
  const allowed = async (path: string, feature: string, permission: string) =>
    client()
      .request(path, { feature, permission } as never)
      .then(
        () => true,
        (error: unknown) => {
          if (error instanceof Error && /does not belong/.test(error.message)) return false;
          throw error;
        },
      );
  const roles = ["directory.roles.read", "RoleManagement.Read.Directory"] as const;
  const groups = ["directory.groups.read", "Group.Read.All"] as const;
  const domains = ["directory.domains.read", "Domain.Read.All"] as const;
  const group = "0f0e0d0c-0b0a-4908-8706-050403020100";

  it("reads the four role lists and nothing else under role management", async () => {
    for (const list of [
      "roleDefinitions",
      "roleAssignments",
      "roleAssignmentScheduleInstances",
      "roleEligibilityScheduleInstances",
    ])
      expect(await allowed(`roleManagement/directory/${list}?$select=id`, ...roles), list).toBe(
        true,
      );
    for (const path of [
      "roleManagement/directory/roleAssignments/a-1",
      "roleManagement/directory/roleDefinitions/62e90394-69f5-4237-9190-012177145e10",
      "roleManagement/directory/roleEligibilityScheduleRequests",
      "roleManagement/directory/roleAssignmentScheduleRequests",
      "roleManagement/directory/roleEligibilitySchedules",
      "roleManagement/directory/roleAssignmentApprovals",
      "roleManagement/entitlementManagement/roleAssignments",
      "roleManagement/directory/roleAssignments%2Fa-1",
      "directoryRoles",
      "users",
    ])
      expect(await allowed(path, ...roles), path).toBe(false);
  });

  it("reads the group list and one group's owners, never members or other navigations", async () => {
    expect(await allowed("groups?$select=id,displayName", ...groups)).toBe(true);
    expect(await allowed(`groups/${group}/owners?$select=id`, ...groups)).toBe(true);
    expect(await allowed(`groups/${group.toUpperCase()}/owners`, ...groups)).toBe(true);
    for (const path of [
      `groups/${group}`,
      `groups/${group}/members`,
      `groups/${group}/transitiveMembers`,
      `groups/${group}/owners/extra`,
      `groups/${group}/team`,
      `groups/${group}/drive`,
      `groups/${group}/conversations`,
      "groups/not-a-guid/owners",
      `groups/${group}%2Fowners`,
      "groups/delta",
      "groupLifecyclePolicies",
      "teams",
      "users",
    ])
      expect(await allowed(path, ...groups), path).toBe(false);
  });

  it("reads the domain list only", async () => {
    expect(await allowed("domains?$select=id,passwordValidityPeriodInDays", ...domains)).toBe(true);
    for (const path of ["domains/contoso.com", "domains/contoso.com/serviceConfigurationRecords"])
      expect(await allowed(path, ...domains), path).toBe(false);
  });

  it("keeps each new list out of every other pack", async () => {
    const users = ["directory.users.read", "User.Read.All"] as const;
    const grants = ["directory.permission-grants.read", "Directory.Read.All"] as const;
    for (const pack of [users, grants, groups, domains])
      expect(await allowed("roleManagement/directory/roleAssignments", pack[0], pack[1])).toBe(
        false,
      );
    for (const pack of [users, grants, roles, domains])
      expect(await allowed(`groups/${group}/owners`, pack[0], pack[1])).toBe(false);
    for (const pack of [users, grants, roles, groups])
      expect(await allowed("domains", pack[0], pack[1])).toBe(false);
    expect(await allowed("users?$select=id", ...groups)).toBe(false);
  });
});

describe("GraphHttpClient", () => {
  it("blocks mutations by default", async () => {
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation: vi.fn(),
    });

    await expect(client.request("users/1", context, { method: "DELETE" })).rejects.toThrow(
      "read-only boundary",
    );
  });

  it("honors Retry-After before retrying a throttled read", async () => {
    const sleep = vi.fn(async () => undefined);
    const fetchImplementation = vi
      .fn()
      .mockResolvedValueOnce(new Response("", { status: 429, headers: { "Retry-After": "2" } }))
      .mockResolvedValueOnce(Response.json({ value: [{ id: "synthetic-1" }] }));
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation,
      sleep,
    });

    const result = await client.request(
      "auditLogs/signIns",
      context,
      {},
      z.object({ value: z.array(z.object({ id: z.string() })) }),
    );

    expect(result.value[0]?.id).toBe("synthetic-1");
    expect(sleep).toHaveBeenCalledWith(2_000);
    expect(fetchImplementation).toHaveBeenCalledTimes(2);
  });

  it("follows an allowed opaque nextLink", async () => {
    const fetchImplementation = vi
      .fn()
      .mockResolvedValueOnce(
        Response.json({
          value: [{ id: "a" }],
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=synthetic",
        }),
      )
      .mockResolvedValueOnce(Response.json({ value: [{ id: "b" }] }));
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation,
    });

    await expect(
      client.getAll("users", userContext, { schema: z.object({ id: z.string() }) }),
    ).resolves.toEqual([{ id: "a" }, { id: "b" }]);
  });

  it("rejects a nextLink that escapes the Graph host", async () => {
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation: vi.fn(),
    });

    await expect(client.request("https://example.invalid/v1.0/users", context)).rejects.toThrow(
      "host boundary",
    );
  });

  it("uses the allowlisted GCC High Graph endpoint", async () => {
    const fetchImplementation = vi.fn().mockResolvedValue(Response.json({ value: [] }));
    const client = new GraphHttpClient({
      cloud: "usGovernmentL4",
      getAccessToken: async () => "synthetic-token",
      fetchImplementation,
    });

    await client.request("auditLogs/signIns", context);

    expect(String(fetchImplementation.mock.calls[0]?.[0])).toBe(
      "https://graph.microsoft.us/v1.0/auditLogs/signIns",
    );
  });

  it("rejects a nextLink from a different Microsoft cloud", async () => {
    const client = new GraphHttpClient({
      cloud: "usGovernmentL4",
      getAccessToken: async () => "synthetic-token",
      fetchImplementation: vi.fn(),
    });

    await expect(client.request("https://graph.microsoft.com/v1.0/users", context)).rejects.toThrow(
      "host boundary",
    );
  });

  it("rejects a permission that does not belong to the feature", async () => {
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation: vi.fn(),
    });

    await expect(
      client.request("users", {
        feature: "directory.users.read",
        permission: "Reports.Read.All",
      }),
    ).rejects.toThrow("does not match");
  });

  it("rejects a feature that Microsoft does not expose in the selected cloud", async () => {
    const client = new GraphHttpClient({
      cloud: "usGovernmentL5",
      getAccessToken: async () => "synthetic-token",
      fetchImplementation: vi.fn(),
    });

    await expect(
      client.request("reports/getOffice365ActiveUserDetail(period='D7')", {
        feature: "reports.usage.read",
        permission: "Reports.Read.All",
      }),
    ).rejects.toThrow("not supported");
  });

  it("returns a redacted provider error with request ID", async () => {
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation: async () =>
        new Response('{"error":{"message":"potentially sensitive"}}', {
          status: 403,
          headers: { "request-id": "synthetic-request-id" },
        }),
    });

    const error = await client
      .request("auditLogs/signIns", context)
      .catch((caught: unknown) => caught);
    expect(error).toBeInstanceOf(GraphRequestError);
    expect(error).toMatchObject({ status: 403, requestId: "synthetic-request-id" });
    expect(String(error)).not.toContain("potentially sensitive");
  });

  it("does not follow Graph redirects or trust a caller's redirect option", async () => {
    const fetchImplementation = vi.fn().mockResolvedValue(
      new Response(null, {
        status: 302,
        headers: { Location: "https://example.invalid/download" },
      }),
    );
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation,
    });
    await expect(client.request("users", userContext, { redirect: "follow" })).rejects.toThrow(
      "redirects are disabled",
    );
    expect(fetchImplementation.mock.calls[0]?.[1]?.redirect).toBe("manual");
    expect(fetchImplementation).toHaveBeenCalledTimes(1);
  });

  it("defers a long Retry-After without shortening it or issuing an early retry", async () => {
    const sleep = vi.fn(async () => undefined);
    const fetchImplementation = vi
      .fn()
      .mockResolvedValue(new Response(null, { status: 429, headers: { "Retry-After": "120" } }));
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      sleep,
      fetchImplementation,
    });
    const error = await client.request("users", userContext).catch((caught: unknown) => caught);
    expect(error).toBeInstanceOf(GraphRetryDeferredError);
    expect(error).toMatchObject({ retryAfterMilliseconds: 120_000 });
    expect(sleep).not.toHaveBeenCalled();
    expect(fetchImplementation).toHaveBeenCalledTimes(1);
  });

  it("binds an endpoint to its feature permission pack", async () => {
    const fetchImplementation = vi.fn();
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation,
    });
    await expect(client.request("users", context)).rejects.toThrow("does not belong");
    expect(fetchImplementation).not.toHaveBeenCalled();
  });

  it("validates item and envelope shapes before using a continuation", async () => {
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation: async () => Response.json({ value: [{ id: 7 }] }),
    });
    await expect(
      client.getAll("users", userContext, { schema: z.object({ id: z.string() }) }),
    ).rejects.toThrow("invalid collection item");
    const malformed = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation: async () => Response.json({ value: "not-an-array" }),
    });
    await expect(malformed.getAll("users", userContext)).rejects.toThrow("invalid response shape");
  });

  it("returns bounded partial pages explicitly and getAll rejects incomplete results", async () => {
    const fetchImplementation = vi.fn().mockImplementation(async () =>
      Response.json({
        value: [{ id: "synthetic-1" }, { id: "synthetic-2" }],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=synthetic",
      }),
    );
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation,
    });
    await expect(client.getCollection("users", userContext, { maxRecords: 1 })).resolves.toEqual({
      items: [{ id: "synthetic-1" }],
      complete: false,
    });
    await expect(client.getAll("users", userContext, { maxPages: 1 })).rejects.toThrow(
      "safety limit",
    );
    await expect(client.getAll("users", userContext, { maxPages: 0 })).rejects.toThrow("maxPages");
  });

  it("rejects repeated pagination links", async () => {
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation: async () =>
        Response.json({ value: [], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users" }),
    });
    await expect(client.getAll("users", userContext)).rejects.toThrow("repeated");
  });

  it("caps streamed response bytes even without a content-length header", async () => {
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      maxResponseBytes: 128,
      fetchImplementation: async () => Response.json({ value: "x".repeat(200) }),
    });
    await expect(client.request("users", userContext)).rejects.toThrow("size budget");
  });

  it("ends a stalled request within its timeout", async () => {
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      requestTimeoutMilliseconds: 100,
      fetchImplementation: async () => new Promise<Response>(() => {}),
    });
    await expect(client.request("users", userContext)).rejects.toThrow("time budget");
  });

  it("rejects invalid limits and all attempted request bodies", async () => {
    expect(
      () => new GraphHttpClient({ getAccessToken: async () => "synthetic-token", maxRetries: -1 }),
    ).toThrow("maxRetries");
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation: vi.fn(),
    });
    await expect(
      client.request("users", userContext, { method: "GET", body: "synthetic" }),
    ).rejects.toThrow("read-only");
  });
});

describe("GraphHttpClient continuation casing", () => {
  it("follows a lower-case next-page link for the same endpoint and still rejects other endpoints", async () => {
    const signIns = { feature: "audit.signins.read", permission: "AuditLog.Read.All" } as const;
    const pages: Record<string, unknown> = {
      "/v1.0/auditLogs/signIns": {
        value: [{ id: "a" }],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/auditLogs/signins?$skiptoken=next",
      },
      "/v1.0/auditLogs/signins": { value: [{ id: "b" }] },
    };
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation: async (input) =>
        Response.json(pages[new URL(String(input)).pathname] ?? {}),
    });
    await expect(client.getCollection("auditLogs/signIns", signIns)).resolves.toEqual({
      items: [{ id: "a" }, { id: "b" }],
      complete: true,
    });

    const escaping = new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation: async () =>
        Response.json({
          value: [{ id: "a" }],
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=next",
        }),
    });
    await expect(escaping.getCollection("auditLogs/signIns", signIns)).rejects.toMatchObject({
      name: "GraphBoundaryError",
      reason: "continuation",
    });
  });
});

describe("GraphHttpClient deadlines and cancellation", () => {
  const paging = (onRequest: () => void) =>
    (async (input: RequestInfo | URL, init?: RequestInit) => {
      onRequest();
      if (init?.signal?.aborted) throw new DOMException("Aborted", "AbortError");
      const url = new URL(String(input));
      const page = Number(url.searchParams.get("$skiptoken") ?? "0");
      url.searchParams.set("$skiptoken", String(page + 1));
      return Response.json({ value: [{ id: String(page) }], "@odata.nextLink": url.toString() });
    }) as typeof fetch;

  it("starts no page after the deadline", async () => {
    let requests = 0;
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic",
      fetchImplementation: paging(() => {
        requests += 1;
      }),
    });
    // The first page always runs; the deadline only stops later ones, and what was read is kept.
    await expect(
      client.getCollection("users", userContext, { maxPages: 5, deadline: Date.now() - 1 }),
    ).resolves.toEqual({ items: [{ id: "0" }], complete: false, stopped: "deadline" });
    expect(requests).toBe(1);
  });

  it("aborts the request in flight, not just the next page, when the caller's signal aborts", async () => {
    const controller = new AbortController();
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic",
      requestTimeoutMilliseconds: 60_000,
      // A request that only ends when its own signal aborts.
      fetchImplementation: ((_input: RequestInfo | URL, init?: RequestInit) =>
        new Promise<Response>((_resolve, reject) => {
          init?.signal?.addEventListener("abort", () =>
            reject(new DOMException("Aborted", "AbortError")),
          );
        })) as typeof fetch,
    });
    const started = Date.now();
    const pending = client.getCollection("users", userContext, { signal: controller.signal });
    setTimeout(() => controller.abort(new Error("stopped mid-request")), 50);
    await expect(pending).rejects.toThrow("stopped mid-request");
    expect(Date.now() - started).toBeLessThan(5_000);
  });

  it("stops paging and the request in flight when the caller's signal aborts", async () => {
    const controller = new AbortController();
    let requests = 0;
    const client = new GraphHttpClient({
      getAccessToken: async () => "synthetic",
      fetchImplementation: paging(() => {
        requests += 1;
        if (requests === 2) controller.abort(new Error("stopped by caller"));
      }),
    });
    await expect(
      client.getCollection("users", userContext, { maxPages: 50, signal: controller.signal }),
    ).rejects.toThrow("stopped by caller");
    expect(requests).toBe(2);
  });
});

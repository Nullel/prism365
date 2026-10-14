import { describe, expect, it, vi } from "vitest";
import { z } from "zod";
import { GraphHttpClient, GraphRequestError, GraphRetryDeferredError } from "../src/client.ts";

const context = {
  feature: "audit.signins.read",
  permission: "AuditLog.Read.All",
} as const;
const userContext = { feature: "directory.users.read", permission: "User.Read.All" } as const;

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

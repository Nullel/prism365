import { describe, expect, it, vi } from "vitest";
import { GraphHttpClient, GraphRequestError } from "../src/client.ts";

const context = {
  feature: "audit.signins.read",
  permission: "AuditLog.Read.All",
} as const;

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

    const result = await client.request<{ value: Array<{ id: string }> }>(
      "auditLogs/signIns",
      context,
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

    await expect(client.getAll<{ id: string }>("users", context)).resolves.toEqual([
      { id: "a" },
      { id: "b" },
    ]);
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

    const error = await client.request("users", context).catch((caught: unknown) => caught);
    expect(error).toBeInstanceOf(GraphRequestError);
    expect(error).toMatchObject({ status: 403, requestId: "synthetic-request-id" });
    expect(String(error)).not.toContain("potentially sensitive");
  });
});

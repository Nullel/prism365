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

import { describe, expect, it } from "vitest";
import { GraphHttpClient } from "../src/client.ts";
import { lookupPersonId, readPersonDirectory } from "../src/person.ts";

// Synthetic directory records only.
const userId = "44444444-4444-4444-8444-444444444444";

function client(respond: (path: string, select: string | null) => Response) {
  const requests: string[] = [];
  const graph = new GraphHttpClient({
    getAccessToken: async () => "synthetic-token",
    sleep: async () => {},
    maxRetries: 0,
    fetchImplementation: async (input) => {
      const url = new URL(String(input));
      requests.push(decodeURIComponent(url.pathname));
      return respond(decodeURIComponent(url.pathname), url.searchParams.get("$select"));
    },
  });
  return { graph, requests };
}

const profile = {
  id: userId,
  displayName: "Synthetic Person",
  userPrincipalName: "person@example.invalid",
  accountEnabled: false,
  createdDateTime: "not a date",
};

describe("person directory", () => {
  it("asks for nothing without the user read permission", async () => {
    const { graph, requests } = client(() => Response.json(profile));
    const result = await readPersonDirectory({ key: { userId }, client: graph, grantedScopes: [] });
    expect(result).toMatchObject({ state: "missingConsent", profile: null });
    expect(requests).toEqual([]);
  });

  it("reports a deleted account as not found", async () => {
    const { graph } = client(() => new Response("{}", { status: 404 }));
    const result = await readPersonDirectory({
      key: { userId },
      client: graph,
      grantedScopes: ["User.Read.All"],
    });
    expect(result.state).toBe("notFound");
  });

  it("keeps the profile when other parts fail or lack permission", async () => {
    const { graph, requests } = client((path) => {
      if (path.endsWith("/manager")) return new Response("{}", { status: 503 });
      return Response.json(profile);
    });
    const result = await readPersonDirectory({
      key: { userId },
      client: graph,
      grantedScopes: ["User.Read.All"],
    });
    expect(result).toMatchObject({
      state: "available",
      profile: { displayName: "Synthetic Person", accountEnabled: false, createdDateTime: null },
      manager: { state: "unavailable", data: null },
      signInActivity: { state: "missingConsent" },
      licenses: { state: "missingConsent", data: [] },
    });
    // Neither the sign-in activity nor the licenses were requested without their permissions.
    expect(requests.some((path) => path.endsWith("/licenseDetails"))).toBe(false);
    expect(requests).toHaveLength(2);
  });

  it("looks a principal name up and treats a missing one as no match", async () => {
    const found = client(() => Response.json({ id: userId }));
    expect(
      await lookupPersonId({ userPrincipalName: "o'brien@example.invalid", client: found.graph }),
    ).toBe(userId);
    expect(found.requests[0]).toBe("/v1.0/users('o''brien@example.invalid')");
    const missing = client(() => new Response("{}", { status: 404 }));
    expect(
      await lookupPersonId({ userPrincipalName: "nobody@example.invalid", client: missing.graph }),
    ).toBeNull();
  });
});

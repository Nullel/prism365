import { describe, expect, it } from "vitest";
import { GraphHttpClient } from "../src/client.ts";
import { lookupPersonId, readPersonDirectory } from "../src/person.ts";

// Synthetic directory records only.
const userId = "44444444-4444-4444-8444-444444444444";

function client(
  respond: (path: string, select: string | null, params: URLSearchParams) => Response,
) {
  const requests: string[] = [];
  const graph = new GraphHttpClient({
    getAccessToken: async () => "synthetic-token",
    sleep: async () => {},
    maxRetries: 0,
    fetchImplementation: async (input) => {
      const url = new URL(String(input));
      requests.push(decodeURIComponent(url.pathname));
      return respond(
        decodeURIComponent(url.pathname),
        url.searchParams.get("$select"),
        url.searchParams,
      );
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
      directReports: { state: "unavailable", more: false, data: [] },
      signInActivity: { state: "missingConsent" },
      licenses: { state: "missingConsent", data: [] },
    });
    // Neither the sign-in activity nor the licenses were requested without their permissions.
    expect(requests.some((path) => path.endsWith("/licenseDetails"))).toBe(false);
    expect(requests).toHaveLength(3);
  });

  describe("direct reports", () => {
    const grantedScopes = ["User.Read.All"];
    const reports = (respond: (path: string) => Response) =>
      client((path) => (path.endsWith("/directReports") ? respond(path) : Response.json(profile)));
    const read = (graph: GraphHttpClient) =>
      readPersonDirectory({ key: { userId }, client: graph, grantedScopes });

    it("asks for one page of 50 users and keeps users only", async () => {
      const seen: (string | null)[] = [];
      const tops: (string | null)[] = [];
      const { graph, requests } = client((path, select, params) => {
        if (!path.endsWith("/directReports")) return Response.json(profile);
        seen.push(select);
        tops.push(params.get("$top"));
        return Response.json({
          value: [
            {
              "@odata.type": "#microsoft.graph.user",
              id: "r-1",
              displayName: "Ines Calder",
              userPrincipalName: "ines@example.invalid",
              jobTitle: "Analyst",
            },
            { "@odata.type": "#microsoft.graph.contact", id: "c-1", displayName: "A Contact" },
            { id: "r-2", displayName: "Ravi Okoye", userPrincipalName: "ravi@example.invalid" },
            { id: "x-1", displayName: "No type and no name" },
          ],
        });
      });
      const result = await read(graph);
      expect(seen).toEqual(["id,displayName,userPrincipalName,jobTitle"]);
      expect(tops).toEqual(["50"]);
      expect(requests.filter((path) => path.endsWith("/directReports"))).toHaveLength(1);
      expect(result.directReports).toEqual({
        state: "available",
        more: false,
        // The contact and the entry with no type and no principal name are not accounts.
        skipped: 2,
        data: [
          {
            id: "r-1",
            displayName: "Ines Calder",
            userPrincipalName: "ines@example.invalid",
            jobTitle: "Analyst",
          },
          {
            id: "r-2",
            displayName: "Ravi Okoye",
            userPrincipalName: "ravi@example.invalid",
            jobTitle: null,
          },
        ],
      });
    });

    it("says there are more when Microsoft returns a next link, without following it", async () => {
      const { graph, requests } = reports(() =>
        Response.json({
          value: [
            { id: "r-1", displayName: "Ines Calder", userPrincipalName: "i@example.invalid" },
          ],
          "@odata.nextLink": `https://graph.microsoft.com/v1.0/users/${userId}/directReports?$skiptoken=abc`,
        }),
      );
      const result = await read(graph);
      expect(result.directReports).toMatchObject({ state: "available", more: true });
      expect(requests.filter((path) => path.endsWith("/directReports"))).toHaveLength(1);
    });

    it("accepts a next link in the directoryObjects form without resolving it", async () => {
      // Graph uses this form for directory object collections. It is not on the pack's paths, so
      // validating it, as getCollection does, would fail a read that only needs the first page.
      const { graph, requests } = reports(() =>
        Response.json({
          value: [
            { id: "r-1", displayName: "Ines Calder", userPrincipalName: "i@example.invalid" },
          ],
          "@odata.nextLink": `https://graph.microsoft.com/v1.0/directoryObjects/${userId}/Microsoft.DirectoryServices.User/directReports?$skiptoken=abc`,
        }),
      );
      const result = await read(graph);
      expect(result.directReports).toMatchObject({ state: "available", more: true });
      expect(result.directReports?.data).toHaveLength(1);
      expect(requests.filter((path) => path.includes("directReports"))).toHaveLength(1);
    });

    it("caps a page longer than asked for at 50 and says there are more", async () => {
      const many = Array.from({ length: 80 }, (_, n) => ({
        "@odata.type": "#microsoft.graph.user",
        id: `r-${n}`,
        displayName: `Person ${n}`,
        userPrincipalName: `p${n}@example.invalid`,
      }));
      const result = await read(reports(() => Response.json({ value: many })).graph);
      expect(result.state).toBe("available");
      expect(result.directReports?.data).toHaveLength(50);
      expect(result.directReports?.more).toBe(true);
    });

    it("counts the entries that are not accounts", async () => {
      const result = await read(
        reports(() =>
          Response.json({
            value: [
              { "@odata.type": "#microsoft.graph.contact", id: "c-1", displayName: "Contact" },
            ],
          }),
        ).graph,
      );
      expect(result.directReports).toEqual({
        state: "available",
        more: false,
        skipped: 1,
        data: [],
      });
    });

    it("reads a 404 as unavailable, not as no reports, and a 403 as missing consent", async () => {
      const gone = await read(reports(() => new Response("{}", { status: 404 })).graph);
      expect(gone.directReports).toMatchObject({ state: "unavailable", data: [] });
      const denied = await read(reports(() => new Response("{}", { status: 403 })).graph);
      expect(denied.directReports).toMatchObject({ state: "missingConsent", data: [] });
      const failed = await read(reports(() => new Response("{}", { status: 503 })).graph);
      expect(failed.directReports).toMatchObject({ state: "unavailable" });
      const empty = await read(reports(() => Response.json({ value: [] })).graph);
      expect(empty.directReports).toEqual({
        state: "available",
        more: false,
        skipped: 0,
        data: [],
      });
    });

    it("keeps the kind of the manager, so a contact is not taken for an account", async () => {
      const managerOf = async (entry: Record<string, unknown>) => {
        const { graph } = client((path) =>
          path.endsWith("/manager") ? Response.json(entry) : Response.json(profile),
        );
        return (await read(graph)).manager?.data;
      };
      expect(
        await managerOf({ "@odata.type": "#microsoft.graph.user", id: "m-1", displayName: "Mia" }),
      ).toMatchObject({ isUser: true });
      expect(
        await managerOf({
          "@odata.type": "#microsoft.graph.orgContact",
          id: "m-2",
          displayName: "Mo",
        }),
      ).toMatchObject({ isUser: false });
      expect(
        await managerOf({ id: "m-3", displayName: "Mi", userPrincipalName: "mi@example.invalid" }),
      ).toMatchObject({ isUser: true });
      expect(await managerOf({ id: "m-4", displayName: "No type" })).toMatchObject({
        isUser: false,
      });
    });

    it("is null, and never requested, without the user read permission", async () => {
      const { graph, requests } = client(() => Response.json(profile));
      const result = await readPersonDirectory({
        key: { userId },
        client: graph,
        grantedScopes: [],
      });
      expect(result.directReports).toBeNull();
      expect(requests).toEqual([]);
    });
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

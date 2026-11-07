import { describe, expect, it } from "vitest";
import { buildSyntheticTenant, createSyntheticGraph, guid } from "../tenant.ts";

// The synthetic tenant's answers for the Access tab: nesting, roles with their templates, and the
// objects some accounts own.
const now = new Date("2026-10-22T12:00:00Z");
const tenant = buildSyntheticTenant(now);
const graph = createSyntheticGraph({ now });
const base = "https://graph.microsoft.com/v1.0";

async function list(id: string, edge: string, query = "") {
  const response = await graph(`${base}/users/${id}/${edge}${query}`);
  expect(response.status).toBe(200);
  return ((await response.json()) as { value: Record<string, unknown>[] }).value;
}
const idOf = (name: string) => {
  const found = tenant.users.find((user) => user.displayName === name);
  if (!found) throw new Error(`No user named ${name}.`);
  return found.id;
};

describe("the synthetic Access tab data", () => {
  it("gives the technician a directory role with its template, and a Microsoft 365 group", async () => {
    const entries = await list(idOf("Tobias Vance"), "transitiveMemberOf");
    const role = entries.find((entry) => entry.displayName === "Helpdesk Administrator");
    expect(role).toMatchObject({
      "@odata.type": "#microsoft.graph.directoryRole",
      roleTemplateId: "729827e3-9c14-49f7-bb1b-9608f156bbb8",
    });
    expect(entries.find((entry) => entry.displayName === "Larkspur staff")).toMatchObject({
      "@odata.type": "#microsoft.graph.group",
      groupTypes: ["Unified"],
      mailEnabled: true,
    });
  });

  it("nests the file share group under the licensing group, so only some accounts are in it", async () => {
    const nested = guid("group:file-share-readers");
    const licensed = tenant.users.find((user) =>
      user.licenseAssignmentStates.some(
        (state) => state.assignedByGroup === guid("group:licensing"),
      ),
    );
    if (!licensed) throw new Error("No account is licensed through the group.");
    const transitive = await list(licensed.id, "transitiveMemberOf");
    const direct = await list(licensed.id, "memberOf");
    expect(transitive.some((entry) => entry.id === nested)).toBe(true);
    expect(direct.some((entry) => entry.id === nested)).toBe(false);
    // Someone licensed directly is not in it.
    const outside = await list(idOf("Demo administrator"), "transitiveMemberOf");
    expect(outside.some((entry) => entry.id === nested)).toBe(false);
  });

  it("lists what the technician owns, a group and an application", async () => {
    const owned = await list(idOf("Tobias Vance"), "ownedObjects");
    const types = owned.map((entry) => entry["@odata.type"]);
    expect(types).toContain("#microsoft.graph.group");
    expect(types).toContain("#microsoft.graph.application");
    expect(owned.find((entry) => entry.displayName === "Larkspur Lab Portal")).toBeDefined();
  });

  it("answers 404 for an account that does not exist and serves a select", async () => {
    const missing = await graph(`${base}/users/00000000-0000-4000-8000-000000000000/ownedObjects`);
    expect(missing.status).toBe(404);
    const [first] = await list(
      idOf("Tobias Vance"),
      "transitiveMemberOf",
      "?$select=id,displayName",
    );
    // The type stays, as it does in Graph.
    expect(Object.keys(first ?? {}).sort()).toEqual(["@odata.type", "displayName", "id"]);
  });
});

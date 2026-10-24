import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { permissionPacks } from "@prism365/graph";
import { describe, expect, it } from "vitest";
import { supportedReadFeatures } from "../../../apps/api/src/config.ts";

// docs/SETUP.md's "What each feature needs" table must name every pack the API accepts, with
// the permission the catalog gives it.
describe("the setup guide", () => {
  it("lists every read pack with its catalog permission", async () => {
    const guide = await readFile(join(import.meta.dirname, "../../../docs/SETUP.md"), "utf8");
    const rows = new Map(
      [...guide.matchAll(/^\| `([a-z.-]+)` \| `([A-Za-z.]+)` \|/gm)].map((match) => [
        match[1],
        match[2],
      ]),
    );
    expect([...rows.keys()]).toEqual(Object.keys(supportedReadFeatures));
    for (const [feature, permission] of rows)
      expect(permission, feature).toBe(
        permissionPacks.find((pack) => pack.id === feature)?.permission,
      );
  });

  it("says the setup's consent outlives the run, and how to revoke it", async () => {
    const root = join(import.meta.dirname, "../../..");
    // Lines wrap anywhere in the Markdown, so compare with whitespace collapsed.
    const read = async (file: string) =>
      (await readFile(join(root, file), "utf8")).replace(/\s+/g, " ");
    const guide = await read("docs/SETUP.md");
    const permissions = await read("docs/PERMISSIONS.md");
    for (const text of [guide, permissions]) {
      expect(text).not.toContain("for the length of one setup run");
      expect(text).toContain("Consent on behalf of your organization");
    }
    expect(guide).toContain("### Revoke the setup's access");
    // Only Application.ReadWrite.All comes out of the grant, not every consent with it.
    expect(guide).toContain(
      "PATCH https://graph.microsoft.com/v1.0/oauth2PermissionGrants/<grant ID>",
    );
    expect(guide).toContain("the same list without Application.ReadWrite.All");
    // The registration is named, never searched for, and the checks are a snapshot.
    expect(guide).toContain("never looks a registration up by its label or its name");
    expect(guide).toContain("Its owners can change it at any time afterwards");
    expect(guide).not.toContain("never makes two");
    expect(guide).toContain("14d82eec-204b-4c2f-b7e8-296a70dab67e");
    // The one-year limit is Microsoft's, and the guide says where Microsoft says it.
    expect(guide).not.toContain("the longest Microsoft accepts");
    expect(guide).toContain(
      "https://learn.microsoft.com/en-us/graph/applications-how-to-add-certificate",
    );
  });
});

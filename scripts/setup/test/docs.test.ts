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
});

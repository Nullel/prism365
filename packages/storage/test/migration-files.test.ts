import { readdir, readFile } from "node:fs/promises";
import { describe, expect, it } from "vitest";
import { loadCoreMigrations } from "../src/index.ts";

// The Docker build context lets through only the core migrations it names, and the build guard
// refuses any other SQL file, so a new migration has to be listed in three places.
describe("core migration files", () => {
  it("are loaded, let into the build context, and allowed by the build guard", async () => {
    const root = new URL("../../../", import.meta.url);
    const files = (await readdir(new URL("packages/storage/migrations/", root)))
      .filter((name) => name.endsWith(".sql"))
      .sort();
    const loaded = (await loadCoreMigrations()).map((migration) => `${migration.id}.sql`).sort();
    const ignore = await readFile(new URL(".dockerignore", root), "utf8");
    const dockerfile = await readFile(new URL("Dockerfile", root), "utf8");
    expect(loaded).toEqual(files);
    for (const file of files) {
      expect(ignore, file).toContain(`!packages/storage/migrations/${file}`);
      expect(dockerfile, file).toContain(`./packages/storage/migrations/${file}`);
    }
  });
});

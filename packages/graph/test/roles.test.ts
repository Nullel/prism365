import { describe, expect, it } from "vitest";
import { GLOBAL_ADMINISTRATOR_ROLE_ID, privilegedRoles } from "../src/index.ts";

// The shared privileged role map. The audits and the directory change rules all read it, so a
// role added or renamed here changes every one of them at once.

describe("the privileged roles", () => {
  it("are twelve built-in roles named by their template ID", () => {
    expect(privilegedRoles.size).toBe(12);
    for (const id of privilegedRoles.keys()) expect(id).toMatch(/^[0-9a-f-]{36}$/);
    expect(new Set(privilegedRoles.values()).size).toBe(12);
  });

  it("include Global Administrator under the ID Microsoft documents", () => {
    expect(GLOBAL_ADMINISTRATOR_ROLE_ID).toBe("62e90394-69f5-4237-9190-012177145e10");
    expect(privilegedRoles.get(GLOBAL_ADMINISTRATOR_ROLE_ID)).toBe("Global Administrator");
  });
});

import { exportedAtProblem, mailboxAccessExportSchema } from "@prism365/contracts";
import { describe, expect, it } from "vitest";
import { demoMailboxAccessExport } from "./mailbox-access.ts";
import {
  assertSynthetic,
  buildSyntheticTenant,
  createSyntheticGraph,
  guid,
  SYNTHETIC_TENANT_ID,
} from "./tenant.ts";

// The demo's mailbox access export: one the import accepts, made of synthetic addresses only,
// with a mailbox reached directly, one through a group, a Send As, a Send on Behalf and one
// unresolved trustee.
const now = new Date("2026-11-06T12:00:00Z");
const tenant = buildSyntheticTenant(now);
const file = demoMailboxAccessExport(
  tenant,
  new Date(now.getTime() - 2 * 86_400_000),
  SYNTHETIC_TENANT_ID,
);

describe("the demo mailbox access export", () => {
  it("is accepted by the import and names only synthetic addresses", () => {
    const parsed = mailboxAccessExportSchema.safeParse(file);
    expect(parsed.success, JSON.stringify(parsed.error?.issues[0])).toBe(true);
    expect(exportedAtProblem(file.exportedAt, now.getTime())).toBeNull();
    expect(() => assertSynthetic(JSON.stringify(file), tenant.addresses)).not.toThrow();
  });

  it("covers each case the pages show", () => {
    const rights = new Set(file.grants.map((grant) => grant.right));
    expect([...rights].sort()).toEqual(["FullAccess", "SendAs", "SendOnBehalf"]);
    expect(file.grants.filter((grant) => grant.trustee.kind === "unresolved")).toHaveLength(1);
    expect(file.grants.some((grant) => grant.trustee.kind === "group")).toBe(true);
    expect(file.grants.some((grant) => grant.trustee.id === guid("user:tvance"))).toBe(true);
  });

  it("grants to a group the synthetic directory really puts people in", async () => {
    const graph = createSyntheticGraph({ now });
    const member = tenant.users.find(
      (user) =>
        user.department === "Laboratory" && user.userType === "Member" && user.accountEnabled,
    );
    expect(member).toBeDefined();
    const response = await graph(
      `https://graph.microsoft.com/v1.0/users/${member?.id}/transitiveMemberOf?$select=id`,
    );
    const groups = ((await response.json()) as { value: { id: string }[] }).value;
    expect(groups.map((group) => group.id)).toContain(guid("group:laboratory"));
  });
});

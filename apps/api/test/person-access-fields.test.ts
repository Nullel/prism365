import { auditDefinition, builtInAudits } from "@prism365/audits";
import { ACCESS_AUDIT_FIELDS, ENTITY_AUDIT_FIELDS } from "@prism365/contracts";
import { describe, expect, it } from "vitest";

// The columns the Access tab shows are named in one list. A column renamed in its audit must not
// quietly empty the tab.
describe("the Access tab's audit fields", () => {
  it("names only audits and columns that exist, with no repeats", () => {
    for (const [auditId, keys] of Object.entries(ACCESS_AUDIT_FIELDS)) {
      const audit = builtInAudits.find((candidate) => candidate.definition.id === auditId);
      expect(audit, auditId).toBeDefined();
      const columns = new Set(audit?.definition.columns.map((column) => column.key));
      for (const key of keys) expect(columns.has(key), `${auditId} ${key}`).toBe(true);
      expect(new Set(keys).size).toBe(keys.length);
    }
  });

  it("keeps each audit to a handful of fields", () => {
    for (const [auditId, keys] of Object.entries(ACCESS_AUDIT_FIELDS))
      expect(keys.length, auditId).toBeLessThanOrEqual(7);
  });
});

// The same for the panel of a group, role, policy, device or app.
describe("the entity panel's audit fields", () => {
  it("names only audits and columns that exist, with no repeats and no internal id column", () => {
    for (const [auditId, keys] of Object.entries(ENTITY_AUDIT_FIELDS)) {
      const audit = builtInAudits.find((candidate) => candidate.definition.id === auditId);
      expect(audit, auditId).toBeDefined();
      const columns = new Map(
        audit ? auditDefinition(audit).columns.map((column) => [column.key, column]) : [],
      );
      for (const key of keys) {
        expect(columns.has(key), `${auditId} ${key}`).toBe(true);
        expect(columns.get(key)?.internal, `${auditId} ${key}`).toBe(false);
      }
      expect(new Set(keys).size).toBe(keys.length);
    }
  });

  it("marks every column a panel matches on as internal, and leaves the older hidden ones alone", () => {
    const column = (auditId: string, key: string) =>
      auditDefinition(
        builtInAudits.find((candidate) => candidate.definition.id === auditId) as never,
      ).columns.find((item) => item.key === key);
    const internal = [
      ["groups.teams", "ownerList"],
      ["directory.privileged-roles", "roleId"],
      ["directory.privileged-roles", "holderId"],
      ["directory.privileged-roles", "throughId"],
      ...["includeUserIds", "excludeUserIds", "includeGroupIds", "excludeGroupIds"].map(
        (key) => ["identity.conditional-access", key] as const,
      ),
      ...["includeRoleIds", "excludeRoleIds"].map(
        (key) => ["identity.conditional-access", key] as const,
      ),
    ] as const;
    for (const [auditId, key] of internal)
      expect(column(auditId, key), `${auditId} ${key}`).toMatchObject({
        hidden: true,
        internal: true,
      });
    for (const [auditId, key] of [
      ["groups.teams", "ownerIds"],
      ["groups.teams", "renewed"],
      ["devices.managed", "serial"],
      ["apps.consents", "appId"],
    ] as const)
      expect(column(auditId, key), `${auditId} ${key}`).toMatchObject({
        hidden: true,
        internal: false,
      });
  });
});

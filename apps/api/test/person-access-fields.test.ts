import { builtInAudits } from "@prism365/audits";
import { ACCESS_AUDIT_FIELDS } from "@prism365/contracts";
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

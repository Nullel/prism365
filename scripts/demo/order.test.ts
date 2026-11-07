import { COMBINED_RISKS_AUDIT_ID } from "@prism365/contracts";
import { describe, expect, it } from "vitest";
import { auditRecordingOrder } from "./order.ts";

describe("the order the recorder runs audits in", () => {
  it("puts Combined risks after every other audit, wherever the catalog lists it", () => {
    expect(auditRecordingOrder([COMBINED_RISKS_AUDIT_ID, "a.one", "b.two"])).toEqual([
      "a.one",
      "b.two",
      COMBINED_RISKS_AUDIT_ID,
    ]);
    expect(auditRecordingOrder(["a.one", COMBINED_RISKS_AUDIT_ID, "b.two"]).at(-1)).toBe(
      COMBINED_RISKS_AUDIT_ID,
    );
  });

  it("keeps the other audits in the catalog's order and loses none", () => {
    const ids = ["z.last", "a.first", "m.middle"];
    expect(auditRecordingOrder([...ids, COMBINED_RISKS_AUDIT_ID])).toEqual([
      ...ids,
      COMBINED_RISKS_AUDIT_ID,
    ]);
    expect(auditRecordingOrder(ids)).toEqual(ids);
    expect(auditRecordingOrder([])).toEqual([]);
  });
});

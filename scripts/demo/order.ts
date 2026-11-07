import { COMBINED_RISKS_AUDIT_ID } from "@prism365/contracts";

// The order the recorder runs the audits in. Combined risks joins the stored results of the
// others, so it goes after every other audit in each round, whatever order the catalog lists them
// in; the rest keep the catalog's order.
export function auditRecordingOrder(ids: readonly string[]): string[] {
  return [
    ...ids.filter((id) => id !== COMBINED_RISKS_AUDIT_ID),
    ...ids.filter((id) => id === COMBINED_RISKS_AUDIT_ID),
  ];
}

import type { AuditQuery } from "@prism365/contracts";
import { z } from "zod";

// The rows an export takes from an audit run, as the page shows them: ?view=, ?search=,
// ?sort=&direction=, and f.<column>= filters. The CSV export and the handoff export both read
// their query through here, so the same parameters always select the same rows in the same
// order.
export const auditExportParamsSchema = z.record(z.string().max(60), z.string().max(200));
export type AuditExportParams = z.infer<typeof auditExportParamsSchema>;

export function auditExportQuery(params: AuditExportParams): AuditQuery {
  const { view, search, sort, direction } = params;
  const facets = Object.fromEntries(
    Object.entries(params)
      .filter(([key]) => key.startsWith("f."))
      .map(([key, value]) => [key.slice(2), value]),
  );
  return {
    view: view ?? null,
    facets,
    search: search ?? "",
    sort: sort ? { column: sort, direction: direction === "asc" ? "asc" : "desc" } : null,
  };
}

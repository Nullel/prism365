import type { AuditDefinitionInput, AuditRow, AuditSummaryItem } from "@prism365/contracts";
import type { GraphFeatureId } from "@prism365/graph";
import type { ScopedGraph } from "@prism365/modules";

// The audit contract. An audit reads the tenant once, on demand, and returns rows for a table it
// describes itself. It sees a Graph client narrowed to the read packs it declares and nothing
// else: no credentials, no storage, no tenant choice. The runner stores what it returns.

export interface AuditContext {
  readonly graph: ScopedGraph;
  readonly now: Date;
  // Microsoft permissions the session has granted, already narrowed to the enabled read packs.
  readonly granted: ReadonlySet<string>;
  // Aborted when the run exceeds its time limit; checked between reads.
  readonly signal: AbortSignal;
  // The account domains the administrator limited the audits to, as saved when the run started,
  // in lower case. null means every domain. Only an audit that declares `scoped` looks at it.
  readonly scope: ReadonlySet<string> | null;
  // The Microsoft tenant ID of the signed-in session, for an audit that has to tell this tenant's
  // own objects from other organisations'. Unset when the caller does not know it.
  readonly tenantId?: string;
}

export interface AuditResult {
  rows: AuditRow[];
  summary: AuditSummaryItem[];
  notes: string[];
  // false when Microsoft returned less than everything.
  complete: boolean;
  // For audits that use the scope: every account domain read, with its account count, before the
  // scope was applied. Lets the scope chooser list domains the current scope leaves out.
  domains?: { domain: string; accounts: number }[];
}

export interface Audit {
  // Permissions are derived from the packs, and `scoped` is published from the field below, so
  // neither is written twice.
  definition: Omit<AuditDefinitionInput, "permissions" | "optionalPermissions" | "scoped">;
  // Covers only the domains in AuditContext.scope. Such an audit filters its accounts before it
  // computes rows, tiles and notes, adds scopeNote to its notes, and has a "domain" column.
  scoped: boolean;
  // Read packs the audit cannot run without, and packs that only add detail when granted.
  packs: readonly GraphFeatureId[];
  optionalPacks: readonly GraphFeatureId[];
  // Shown when Microsoft refuses the read (HTTP 403): what it needs, in plain words.
  forbidden: string;
  run(context: AuditContext): Promise<AuditResult>;
}

// A failure the audit explains itself. Its message is shown to the administrator, so it must
// never carry provider text or tenant data.
export class AuditError extends Error {
  readonly code: string;
  constructor(code: string, message: string) {
    super(message);
    this.name = "AuditError";
    this.code = code;
  }
}

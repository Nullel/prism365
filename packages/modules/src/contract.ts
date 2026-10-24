import type { Severity } from "@prism365/contracts";
import type { GraphFeatureId, GraphRequestContext } from "@prism365/graph";
import type { FindingPurpose, FindingSubject } from "@prism365/policy";
import type { SqlClient } from "@prism365/storage";
import type { z } from "zod";
import type { AttentionChange, AttentionInput } from "./attention.ts";

// The module contract. A module is the only way Prism365 learns about a tenant: it declares the
// read packs it needs, owns tables in its own namespace, runs collectors on a cadence, and
// emits findings. It never sees credentials, never chooses a tenant, and never writes to the
// tenant, because the context it receives has no way to do any of those things.

export type CredentialKind = "delegated" | "application";

export type ModuleCapability = "entra.p1" | "entra.p2" | "role.reportsReader";

export interface ModuleMigration {
  // "001-events": three digits, then a slug. Applied once, in order, checksummed.
  id: string;
  sql: string;
}

export interface RetentionRule {
  // Table name without the module prefix ("event" for mod_signins_audit_event).
  table: string;
  timeColumn: string;
  days: number;
}

export interface CollectorResult {
  // false when records were left out and coverage is incomplete.
  complete: boolean;
  // true when the collector stopped at its own time budget with more to read now, such as
  // while loading history. The scheduler runs it again within seconds instead of waiting for
  // its cadence.
  more?: boolean;
}

export interface Collector {
  id: string;
  title: string;
  cadenceSeconds: number;
  timeoutSeconds: number;
  run(context: CollectorContext): Promise<CollectorResult | undefined>;
}

export interface PrismModule {
  id: string;
  title: string;
  description: string;
  version: number;
  packs: readonly GraphFeatureId[];
  requires?: readonly ModuleCapability[];
  migrations: readonly ModuleMigration[];
  collectors: readonly Collector[];
  retention: readonly RetentionRule[];
}

export interface EvidenceItem {
  source: string;
  ref: string;
  observedAt: string;
  fields: Record<string, string | number | boolean | null>;
}

export type ModuleFindingPurpose = Exclude<FindingPurpose, "workforce-performance">;

export interface FindingInput {
  ruleId: string;
  ruleVersion: number;
  // Stable identity of "the same problem": repeat observations update one finding.
  dedupeKey: string;
  purpose: ModuleFindingPurpose;
  severity: Severity;
  score: number;
  subject: { kind: FindingSubject; id: string; label?: string };
  title: string;
  summary: string;
  evidence: EvidenceItem[];
  limitations: string[];
  window: { start: Date; end: Date };
  observedAt: Date;
}

export type FindingChange = "created" | "updated" | "reopened" | "escalated";

// Graph access for a module: only its declared packs, and every response is schema-validated.
export interface ScopedGraph {
  get<T>(path: string, context: GraphRequestContext, schema: z.ZodType<T>): Promise<T>;
  getCollection<T>(
    path: string,
    context: GraphRequestContext,
    options: {
      schema: z.ZodType<T>;
      maxPages?: number;
      maxRecords?: number;
      // Epoch milliseconds after which no further page is requested.
      deadline?: number;
    },
  ): Promise<{ items: T[]; complete: boolean; stopped?: "deadline" }>;
}

export interface CollectorContext {
  readonly tenantId: string;
  readonly moduleId: string;
  readonly collectorId: string;
  readonly credentialKind: CredentialKind;
  readonly now: Date;
  readonly signal: AbortSignal;
  readonly graph: ScopedGraph;
  // Fully qualified, quoted name of one of this module's tables: table("event").
  table(name: string): string;
  // A tenant-scoped transaction. Do not call Graph inside it.
  db<T>(operation: (sql: SqlClient) => Promise<T>): Promise<T>;
  readonly cursor: {
    get(): Promise<string | null>;
    set(value: string): Promise<void>;
  };
  readonly findings: {
    upsert(input: FindingInput): Promise<FindingChange>;
  };
  // Notification-center items: things an administrator may need to act on, as the provider
  // reports them. Repeating the same facts is free; only meaningful changes notify.
  readonly attention: {
    upsert(input: AttentionInput): Promise<AttentionChange>;
    // Marks this module's active items of one kind that are not in `keep` as inactive, with a
    // note saying why. Call it only after a complete read, and say in the note whether the
    // provider stopped listing them or Prism365's own rule no longer selects them.
    retire(kind: string, keep: readonly string[], note: string): Promise<number>;
  };
  readonly count: {
    fetched(count: number): void;
    inserted(count: number): void;
    // Records the source returned that could not be read; they are left out and the run is
    // reported as partial coverage rather than failing on them forever.
    skipped(count: number): void;
  };
}

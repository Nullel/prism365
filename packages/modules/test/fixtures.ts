import { randomUUID } from "node:crypto";
import { GraphHttpClient } from "@prism365/graph";
import { PGliteModuleDatabase } from "@prism365/storage";
import type { CollectorContext, CollectorResult, FindingInput, PrismModule } from "../src/index.ts";
import { ModuleRegistry } from "../src/index.ts";

// Synthetic fixtures only: invented tenants, invented accounts, documentation-range values.

export const tenantA = randomUUID();
export const tenantB = randomUUID();
export const actor = randomUUID();

export const itemsMigration = {
  id: "001-items",
  sql: `CREATE TABLE IF NOT EXISTS mod_test_probe_item (
          tenant_id uuid NOT NULL,
          id text NOT NULL,
          value text NOT NULL,
          created_at timestamptz NOT NULL,
          PRIMARY KEY (tenant_id, id)
        );
        CREATE INDEX IF NOT EXISTS mod_test_probe_item_created ON mod_test_probe_item (tenant_id, created_at);
        SELECT prism365_tenant_table('mod_test_probe_item');`,
};

export type ProbeBehavior = (context: CollectorContext) => Promise<CollectorResult | undefined>;

export function probeModule(behavior: { run: ProbeBehavior }): PrismModule {
  return {
    id: "test.probe",
    title: "Probe",
    description: "Synthetic module used to exercise the module runtime.",
    version: 1,
    packs: ["directory.users.read"],
    migrations: [itemsMigration],
    collectors: [
      {
        id: "poll",
        title: "Poll",
        cadenceSeconds: 300,
        timeoutSeconds: 30,
        run: (context) => behavior.run(context),
      },
    ],
    retention: [{ table: "item", timeColumn: "created_at", days: 30 }],
  };
}

export async function openRuntime(module: PrismModule) {
  const registry = new ModuleRegistry().register(module);
  const database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  return { registry, database };
}

export function finding(overrides: Partial<FindingInput> = {}): FindingInput {
  const observedAt = overrides.observedAt ?? new Date("2026-10-22T12:00:00Z");
  return {
    ruleId: "probe.synthetic",
    ruleVersion: 1,
    dedupeKey: "account-1",
    purpose: "security",
    severity: "medium",
    score: 40,
    subject: { kind: "person", id: "synthetic-user-1", label: "Synthetic User" },
    title: "Synthetic sign-in pattern",
    summary: "A synthetic account signed in from a location not seen before for it.",
    evidence: [
      {
        source: "Synthetic fixture",
        ref: "event-1",
        observedAt: observedAt.toISOString(),
        fields: { country: "XA", city: "Example City", success: true },
      },
    ],
    limitations: ["Synthetic data; locations are provider-reported context."],
    window: { start: new Date(observedAt.getTime() - 3_600_000), end: observedAt },
    observedAt,
    ...overrides,
  };
}

export const userPage = {
  value: [{ id: "synthetic-user-1", displayName: "Synthetic User" }],
};

export function fakeGraph(body: unknown = userPage) {
  const calls: string[] = [];
  const fetchImplementation = async (input: RequestInfo | URL) => {
    calls.push(String(input));
    return new Response(JSON.stringify(body), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  };
  const factory = (getAccessToken: () => Promise<string>) =>
    new GraphHttpClient({ getAccessToken, fetchImplementation, sleep: async () => {} });
  return { calls, factory };
}

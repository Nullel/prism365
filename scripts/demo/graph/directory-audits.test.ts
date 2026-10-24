import { randomUUID } from "node:crypto";
import { GraphHttpClient } from "@prism365/graph";
import {
  createDirectoryChangesModule,
  listChanges,
  personChanges,
} from "@prism365/module-directory-changes";
import { listFindings, ModuleRegistry, ModuleScheduler } from "@prism365/modules";
import { PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import { assertSynthetic, buildSyntheticTenant, createSyntheticGraph } from "../tenant.ts";
import { syntheticDirectoryAudits } from "./directory-audits.ts";

vi.setConfig({ testTimeout: 60_000 });

// The synthetic directory audit log, collected by the real module and judged by its real rules:
// the planted stories must be flagged, the window filter must be honoured, and nothing collected
// may come from outside the invented tenant.

const anchor = new Date("2026-10-20T12:00:00.000Z");
const tenant = buildSyntheticTenant(anchor);
const graph = createSyntheticGraph({ now: anchor });
const tenantId = randomUUID();
let database: PGliteModuleDatabase;
let scheduler: ModuleScheduler;

beforeAll(async () => {
  const registry = new ModuleRegistry().register(
    createDirectoryChangesModule(undefined, { sleep: async () => {} }),
  );
  database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  scheduler = new ModuleScheduler({
    enabledPacks: ["audit.directory.read"],
    registry,
    database,
    now: () => anchor,
    graphClient: (getAccessToken) =>
      new GraphHttpClient({ getAccessToken, fetchImplementation: graph, sleep: async () => {} }),
  });
  scheduler.bind({
    kind: "delegated",
    tenantId,
    actorId: randomUUID(),
    grantedPermissions: async () => ["AuditLog.Read.All"],
    accessToken: async () => "synthetic",
  });
  for (let attempt = 0; attempt < 10; attempt += 1) await scheduler.tick();
}, 120_000);

afterAll(async () => {
  await scheduler?.stop();
  await database?.close();
});

const user = (name: string) => tenant.users.find((item) => item.displayName === name);

describe("the synthetic directory audit log", () => {
  it("answers only inside the requested window", async () => {
    const from = new Date(anchor.getTime() - 3 * 86_400_000).toISOString();
    const to = new Date(anchor.getTime() - 2 * 86_400_000).toISOString();
    const query = new URLSearchParams({
      $filter: `activityDateTime ge ${from} and activityDateTime le ${to}`,
      $top: "999",
    });
    const response = await graph(
      `https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?${query}`,
    );
    const body = (await response.json()) as { value: { activityDateTime: string }[] };
    expect(body.value.length).toBeGreaterThan(0);
    for (const record of body.value) {
      expect(record.activityDateTime >= from && record.activityDateTime <= to).toBe(true);
    }
    expect(body.value.length).toBeLessThan(syntheticDirectoryAudits(tenant).length);
  });

  it("is collected in full and flags the four planted stories", async () => {
    const stored = await database.tenant(tenantId, (sql) =>
      listChanges(sql, tenantId, { now: anchor, days: 30, limit: 200 }),
    );
    expect(stored.items).toHaveLength(syntheticDirectoryAudits(tenant).length);
    expect(stored.coverage.phase).toBe("live");
    const findings = await database.tenant(tenantId, (sql) =>
      listFindings(sql, tenantId, { module: "directory.changes", limit: 200 }),
    );
    const byRule = (ruleId: string) => findings.filter((item) => item.ruleId === ruleId);
    expect(findings.map((item) => item.ruleId).sort()).toEqual([
      "app.consent-granted",
      "app.credential-added",
      "policy.conditional-access-changed",
      "role.privileged-assigned",
    ]);
    expect(byRule("role.privileged-assigned")[0]).toMatchObject({
      severity: "critical",
      subject: { kind: "person", id: user("Tobias Vance")?.id },
    });
    expect(byRule("app.credential-added")[0]?.subject.label).toBe("Larkspur Payroll Connector");
    expect(byRule("app.consent-granted")[0]?.severity).toBe("high");
    expect(byRule("policy.conditional-access-changed")[0]?.subject.label).toBe(
      "Require MFA for all users",
    );
    assertSynthetic(JSON.stringify({ stored, findings }), tenant.addresses);
  });

  it("shows changes on the people they involve", async () => {
    const technician = user("Tobias Vance")?.id ?? "";
    const page = await database.tenant(tenantId, (sql) =>
      personChanges(sql, tenantId, technician, { limit: 200 }),
    );
    expect(page.items.some((item) => item.involvement.actor)).toBe(true);
    expect(page.items.some((item) => item.activity === "Add member to role")).toBe(true);
    const newcomer = user("Brigid Alder")?.id ?? "";
    const joined = await database.tenant(tenantId, (sql) => personChanges(sql, tenantId, newcomer));
    expect(joined.items.map((item) => item.activity)).toEqual(
      expect.arrayContaining(["Add user", "Change user license", "Add member to group"]),
    );
  });
});

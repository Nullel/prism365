import { randomBytes, randomUUID } from "node:crypto";
import type { FindingNotification } from "@prism365/modules";
import { MemoryStateStore, PostgresModuleDatabase, runPostgresMigrations } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import {
  disposableAdminUrl,
  holdPostgresSuiteLock,
  ownerPool,
} from "../../../packages/storage/test/postgres-suite-lock.ts";
import { DeliveryService } from "../src/delivery.ts";
import type { WebhookConnector } from "../src/delivery-webhook.ts";
import { createModuleRegistry } from "../src/modules.ts";

// Alert delivery on real PostgreSQL, as the restricted runtime role, where transactions overlap
// and row security is enforced by the server. Runs only with TEST_DATABASE_ADMIN_URL, against a
// database it creates and drops. Synthetic tenant, administrator and hosts only.

const adminUrl = process.env.TEST_DATABASE_ADMIN_URL;
// This public password belongs only to the disposable test database; it is not a secret.
const syntheticPassword = "Prism365_Synthetic_Database_Test_Only_12345";
const tenantId = "11111111-1111-4111-8111-111111111111";
const otherTenant = "55555555-5555-4555-8555-555555555555";
const actorId = "22222222-2222-4222-8222-222222222222";

function identifier(value: string) {
  return `"${value.replaceAll('"', '""')}"`;
}

describe.skipIf(!adminUrl)("alert delivery on PostgreSQL", () => {
  const suffix = randomBytes(8).toString("hex");
  const databaseName = `prism365_delivery_${suffix}`;
  const memberRole = `prism365_delivery_member_${suffix}`;
  let releaseSuiteLock: (() => Promise<void>) | undefined;
  let bootstrap: ReturnType<typeof ownerPool> | undefined;
  let database: PostgresModuleDatabase | undefined;
  let service: DeliveryService | undefined;
  let createdAppRole = false;
  const sent: string[] = [];
  const connect: WebhookConnector = async (_target, request) => {
    sent.push(request.body);
    return { status: 200 };
  };

  beforeAll(async () => {
    releaseSuiteLock = await holdPostgresSuiteLock(adminUrl ?? "");
    const parsed = disposableAdminUrl(adminUrl ?? "");
    bootstrap = ownerPool(parsed.toString());
    const existing = await bootstrap.query("SELECT 1 FROM pg_roles WHERE rolname='prism365_app'");
    createdAppRole = existing.rowCount === 0;
    await bootstrap.query(`CREATE DATABASE ${identifier(databaseName)}`);
    parsed.pathname = `/${databaseName}`;
    await runPostgresMigrations({
      adminUrl: parsed.toString(),
      appPassword: syntheticPassword,
      scopes: createModuleRegistry().migrationScopes(),
    });
    await bootstrap.query(
      `CREATE ROLE ${identifier(memberRole)} LOGIN NOSUPERUSER NOBYPASSRLS PASSWORD '${syntheticPassword}' IN ROLE prism365_app`,
    );
    parsed.username = memberRole;
    parsed.password = syntheticPassword;
    database = new PostgresModuleDatabase({ connectionString: parsed.toString() });
    await database.verify();
    service = new DeliveryService({
      tenantId,
      database,
      store: new MemoryStateStore(),
      origin: "https://prism365.example",
      config: { email: { smtp: null, domains: [] }, webhookPorts: [] },
      encryptionKey: "c".repeat(64),
      lookup: async () => [{ address: "100.128.0.1", family: 4 }],
      connect,
      flushDelayMilliseconds: 0,
      log: () => {},
    });
    service.start();
  }, 180_000);

  afterAll(async () => {
    await service?.stop();
    await database?.close();
    try {
      if (bootstrap) {
        await bootstrap.query(`DROP DATABASE IF EXISTS ${identifier(databaseName)} WITH (FORCE)`);
        await bootstrap.query(`DROP ROLE IF EXISTS ${identifier(memberRole)}`);
        if (createdAppRole) await bootstrap.query("DROP ROLE prism365_app");
      }
    } finally {
      await bootstrap?.end();
      await releaseSuiteLock?.();
    }
  }, 60_000);

  const live = () => {
    if (!service || !database) throw new Error("The suite did not start.");
    return { service, database };
  };

  it("keeps two rules created at the same moment", async () => {
    const { service } = live();
    const hook = await service.createDestination(
      actorId,
      { kind: "generic", label: "Hook", url: "https://hooks.example/x", includeDetails: false },
      async () => {},
    );
    const input = (name: string) => ({
      kind: "finding" as const,
      name,
      destinations: [hook.destination.id],
      quietMinutes: 60,
      minimumSeverity: "high" as const,
      modules: [],
    });
    await Promise.all(
      Array.from({ length: 6 }, (_, index) =>
        service.createRule(actorId, input(`Rule ${index}`), async () => {}),
      ),
    );
    const overview = await service.overview();
    expect(overview.rules.map((rule) => rule.name).sort()).toEqual(
      Array.from({ length: 6 }, (_, index) => `Rule ${index}`),
    );
  });

  it("delivers through the restricted role and never across tenants", async () => {
    const { service, database } = live();
    const notification: FindingNotification = {
      tenantId,
      module: "signins.audit",
      findingId: randomUUID(),
      ruleId: "signins.unfamiliar_signin",
      severity: "critical",
      title: "Sign-in that does not match this account's history",
      summary: "Synthetic",
      subjectLabel: "Avery Stone",
      change: "created",
      observedAt: new Date().toISOString(),
    };
    await service.sink.deliver(notification);
    // Another tenant's finding is not this installation's to alert on.
    await service.sink.deliver({ ...notification, tenantId: otherTenant, findingId: randomUUID() });
    await service.idle();
    // Six rules matched one finding; each sends once to the one destination.
    expect(sent).toHaveLength(6);
    const rows = await database.tenant(tenantId, (sql) =>
      sql.query<{ status: string }>("SELECT status FROM prism365_delivery"),
    );
    expect(rows.rows.map((row) => row.status)).toEqual(Array(6).fill("sent"));
    const other = await database.tenant(otherTenant, (sql) =>
      sql.query("SELECT 1 FROM prism365_delivery"),
    );
    expect(other.rowCount).toBe(0);
  });
});

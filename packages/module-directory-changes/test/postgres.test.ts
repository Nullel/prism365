import { randomBytes, randomUUID } from "node:crypto";
import { GraphHttpClient } from "@prism365/graph";
import { listFindings, ModuleRegistry, ModuleScheduler } from "@prism365/modules";
import { PostgresModuleDatabase, runPostgresMigrations } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import {
  disposableAdminUrl,
  holdPostgresSuiteLock,
  ownerPool,
} from "../../storage/test/postgres-suite-lock.ts";
import {
  createDirectoryChangesModule,
  defaultDirectoryChangesConfig,
  listChanges,
  personChanges,
} from "../src/index.ts";

// The directory change feed on real PostgreSQL: the module's migration under the restricted
// runtime role and row security, collection through the scheduler, the feed and person queries,
// tenant isolation, and the retention rule. Runs only with TEST_DATABASE_ADMIN_URL, against a
// database it creates and drops. Synthetic tenant, people and changes only.

vi.setConfig({ testTimeout: 60_000 });

const adminUrl = process.env.TEST_DATABASE_ADMIN_URL;
// This public password belongs only to the disposable test database; it is not a secret.
const syntheticPassword = "Prism365_Synthetic_Database_Test_Only_12345";
const tenantId = randomUUID();
const otherTenant = randomUUID();
const personId = "0a1b2c3d-0000-4000-8000-00000000b002";
const now = new Date("2026-10-22T12:00:00.000Z");
const HOUR = 3_600_000;
const DAY = 24 * HOUR;

function identifier(value: string) {
  return `"${value.replaceAll('"', '""')}"`;
}

function audit(id: string, at: number, overrides: Record<string, unknown> = {}) {
  return {
    id,
    activityDateTime: new Date(at).toISOString(),
    activityDisplayName: "Update user",
    category: "UserManagement",
    result: "success",
    initiatedBy: {
      user: { id: "0a1b2c3d-0000-4000-8000-00000000a001", displayName: "Synthetic Admin" },
    },
    targetResources: [
      { id: personId, displayName: "Synthetic Person", type: "User", modifiedProperties: [] },
    ],
    additionalDetails: [{ key: "User-Agent", value: "Synthetic browser" }],
    ...overrides,
  };
}

// Five changes, two of them in one second so a small window limit makes the collector split.
const log = [
  audit("Directory_pg_1", now.getTime() - 5 * DAY),
  audit("Directory_pg_2", now.getTime() - 2 * DAY, {
    activityDisplayName: "Add member to role",
    category: "RoleManagement",
    targetResources: [
      {
        id: personId,
        displayName: "Synthetic Person",
        type: "User",
        modifiedProperties: [
          {
            displayName: "Role.TemplateId",
            oldValue: null,
            newValue: '"62e90394-69f5-4237-9190-012177145e10"',
          },
        ],
      },
    ],
  }),
  audit("Directory_pg_3", now.getTime() - DAY),
  audit("Directory_pg_4", now.getTime() - DAY + 300),
  audit("Directory_pg_5", now.getTime() - DAY + 600),
];

async function graph(input: RequestInfo | URL): Promise<Response> {
  const url = new URL(String(input));
  const match = /activityDateTime ge (\S+) and activityDateTime le (\S+)/.exec(
    url.searchParams.get("$filter") ?? "",
  );
  const from = Date.parse(match?.[1] ?? "");
  const to = Date.parse(match?.[2] ?? "");
  const value = log.filter((record) => {
    const at = Date.parse(record.activityDateTime);
    return at >= from && at <= to;
  });
  const skip = Number(url.searchParams.get("$skiptoken") ?? "0");
  const next = new URL(url);
  next.searchParams.set("$skiptoken", String(skip + 1));
  return Response.json({
    value: value.slice(skip, skip + 1),
    ...(skip + 1 < value.length ? { "@odata.nextLink": next.href } : {}),
  });
}

describe.skipIf(!adminUrl)("directory changes on PostgreSQL", () => {
  const suffix = randomBytes(8).toString("hex");
  const databaseName = `prism365_changes_${suffix}`;
  const memberRole = `prism365_changes_member_${suffix}`;
  let releaseSuiteLock: (() => Promise<void>) | undefined;
  let bootstrap: ReturnType<typeof ownerPool> | undefined;
  let owner: ReturnType<typeof ownerPool> | undefined;
  let database: PostgresModuleDatabase | undefined;
  let scheduler: ModuleScheduler | undefined;
  let createdAppRole = false;

  beforeAll(async () => {
    releaseSuiteLock = await holdPostgresSuiteLock(adminUrl ?? "");
    const parsed = disposableAdminUrl(adminUrl ?? "");
    bootstrap = ownerPool(parsed.toString());
    const existing = await bootstrap.query("SELECT 1 FROM pg_roles WHERE rolname='prism365_app'");
    createdAppRole = existing.rowCount === 0;
    await bootstrap.query(`CREATE DATABASE ${identifier(databaseName)}`);
    parsed.pathname = `/${databaseName}`;
    const registry = new ModuleRegistry().register(
      createDirectoryChangesModule(
        { ...defaultDirectoryChangesConfig, windowRecordLimit: 2, retentionDays: 365 },
        { sleep: async () => {} },
      ),
    );
    await runPostgresMigrations({
      adminUrl: parsed.toString(),
      appPassword: syntheticPassword,
      scopes: registry.migrationScopes(),
    });
    await bootstrap.query(
      `CREATE ROLE ${identifier(memberRole)} LOGIN NOSUPERUSER NOBYPASSRLS PASSWORD '${syntheticPassword}' IN ROLE prism365_app`,
    );
    owner = ownerPool(parsed.toString());
    // A change older than the retention, stored by an earlier install.
    await owner.query(
      `INSERT INTO mod_directory_changes_event (tenant_id, id, activity_at, initiated_by_kind,
         targets, details, search_text, processed, collected_at)
       VALUES ($1, 'Directory_pg_old', $2, 'unknown', '[]'::jsonb, '[]'::jsonb, '', true, $2)`,
      [tenantId, new Date(now.getTime() - 400 * DAY).toISOString()],
    );
    parsed.username = memberRole;
    parsed.password = syntheticPassword;
    database = new PostgresModuleDatabase({ connectionString: parsed.toString() });
    await database.verify();
    scheduler = new ModuleScheduler({
      enabledPacks: ["audit.directory.read"],
      registry,
      database,
      now: () => now,
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
    await owner?.end();
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

  it("collects every change, splitting the busy second, and judges them", async () => {
    const page = await database?.tenant(tenantId, (sql) =>
      listChanges(sql, tenantId, { now, days: 30 }),
    );
    expect(page?.items.map((item) => item.id).sort()).toEqual(log.map((record) => record.id));
    expect(page?.coverage.phase).toBe("live");
    expect(JSON.stringify(page)).not.toContain("Synthetic browser");
    const findings = await database?.tenant(tenantId, (sql) =>
      listFindings(sql, tenantId, { module: "directory.changes" }),
    );
    expect(findings?.map((item) => item.ruleId)).toEqual(["role.privileged-assigned"]);
  });

  it("searches, pages and lists a person's changes", async () => {
    const found = await database?.tenant(tenantId, (sql) =>
      listChanges(sql, tenantId, { now, days: 30, search: "ADD MEMBER" }),
    );
    expect(found?.items.map((item) => item.id)).toEqual(["Directory_pg_2"]);
    const first = await database?.tenant(tenantId, (sql) =>
      listChanges(sql, tenantId, { now, days: 30, limit: 2 }),
    );
    const second = await database?.tenant(tenantId, (sql) =>
      listChanges(sql, tenantId, { now, days: 30, limit: 10, before: first?.next ?? null }),
    );
    expect([...(first?.items ?? []), ...(second?.items ?? [])].map((item) => item.id)).toEqual([
      "Directory_pg_5",
      "Directory_pg_4",
      "Directory_pg_3",
      "Directory_pg_2",
      "Directory_pg_1",
    ]);
    const person = await database?.tenant(tenantId, (sql) =>
      personChanges(sql, tenantId, personId, { limit: 200 }),
    );
    expect(person?.items).toHaveLength(5);
  });

  it("keeps another tenant out, and deletes what is older than the retention", async () => {
    const other = await database?.tenant(otherTenant, (sql) =>
      listChanges(sql, otherTenant, { now, days: 3_650 }),
    );
    expect(other?.items).toEqual([]);
    const old = await owner?.query(
      "SELECT count(*)::int AS n FROM mod_directory_changes_event WHERE id = 'Directory_pg_old'",
    );
    expect(old?.rows[0]?.n).toBe(0);
  });
});

import { randomUUID } from "node:crypto";
import { ModuleRegistry } from "@prism365/modules";
import { PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import {
  changeCoverage,
  changeSubjects,
  createDirectoryChangesModule,
  type DirectoryChangeEvent,
  listChanges,
  personChanges,
} from "../src/index.ts";

vi.setConfig({ testTimeout: 30_000 });

// Synthetic tenant, people and changes only.
const tenantId = randomUUID();
const now = new Date("2026-10-22T12:00:00.000Z");
const HOUR = 3_600_000;
const DAY = 24 * HOUR;
const adminId = "0a1b2c3d-0000-4000-8000-00000000a001";
const personId = "0a1b2c3d-0000-4000-8000-00000000b002";
const otherId = "0a1b2c3d-0000-4000-8000-00000000c003";

let database: PGliteModuleDatabase;

function change(
  id: string,
  at: Date,
  overrides: Partial<DirectoryChangeEvent> = {},
): DirectoryChangeEvent {
  return {
    id,
    activityAt: at,
    activity: "Update user",
    category: "UserManagement",
    operationType: "Update",
    result: "success",
    resultReason: null,
    loggedByService: "Core Directory",
    correlationId: null,
    initiatedBy: {
      kind: "user",
      id: adminId,
      name: "Synthetic Admin",
      principal: "admin@example.invalid",
    },
    targets: [
      {
        id: otherId,
        name: "Other Person",
        type: "User",
        principal: "other@example.invalid",
        properties: [{ name: "JobTitle", oldValue: "Analyst", newValue: "Lead analyst" }],
      },
    ],
    details: [],
    ...overrides,
  };
}

async function store(event: DirectoryChangeEvent) {
  await database.tenant(tenantId, async (sql) => {
    await sql.query(
      `INSERT INTO mod_directory_changes_event (tenant_id, id, activity_at, category, activity,
         operation_type, result, initiated_by_kind, initiated_by_id, initiated_by_name,
         initiated_by_principal, targets, details, processed, collected_at)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12::jsonb,'[]'::jsonb,true,$13)`,
      [
        tenantId,
        event.id,
        event.activityAt.toISOString(),
        event.category,
        event.activity,
        event.operationType,
        event.result,
        event.initiatedBy.kind,
        event.initiatedBy.id,
        event.initiatedBy.name,
        event.initiatedBy.principal,
        JSON.stringify(event.targets),
        now.toISOString(),
      ],
    );
    for (const subject of changeSubjects(event))
      await sql.query(
        `INSERT INTO mod_directory_changes_subject (tenant_id, event_id, subject_id, subject_kind, subject_type, activity_at)
         VALUES ($1,$2,$3,$4,$5,$6)`,
        [
          tenantId,
          event.id,
          subject.id,
          subject.kind,
          subject.type,
          event.activityAt.toISOString(),
        ],
      );
  });
}

const ago = (milliseconds: number) => new Date(now.getTime() - milliseconds);
const feed = (filter: Partial<Parameters<typeof listChanges>[2]> = {}) =>
  database.tenant(tenantId, (sql) => listChanges(sql, tenantId, { now, days: 7, ...filter }));

beforeAll(async () => {
  const registry = new ModuleRegistry().register(createDirectoryChangesModule());
  database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  await store(change("a-oldest", ago(20 * DAY)));
  await store(
    change("b-role", ago(3 * DAY), {
      activity: "Add member to role",
      category: "RoleManagement",
      targets: [
        {
          id: personId,
          name: "Synthetic Person",
          type: "User",
          principal: "person@example.invalid",
          properties: [
            { name: "Role.DisplayName", oldValue: null, newValue: "Global Administrator" },
          ],
        },
      ],
    }),
  );
  await store(
    change("c-failed", ago(2 * DAY), {
      activity: "Add group",
      category: "GroupManagement",
      result: "failure",
    }),
  );
  await store(change("d-timeout", ago(DAY), { category: "GroupManagement", result: "timeout" }));
  // Two changes in the same millisecond, and one the person made to their own account.
  await store(change("e-same", ago(HOUR)));
  await store(change("f-same", ago(HOUR)));
  await store(
    change("g-self", ago(30 * 60_000), {
      initiatedBy: { kind: "user", id: personId, name: "Synthetic Person", principal: null },
      targets: [
        { id: personId, name: "Synthetic Person", type: "User", principal: null, properties: [] },
      ],
    }),
  );
}, 60_000);

afterAll(async () => {
  await database?.close();
});

describe("change feed queries", () => {
  it("lists the window newest first with counts by category", async () => {
    const page = await feed();
    expect(page.items.map((item) => item.id)).toEqual([
      "g-self",
      "f-same",
      "e-same",
      "d-timeout",
      "c-failed",
      "b-role",
    ]);
    expect(page.next).toBeNull();
    expect(page.categories).toEqual([
      { category: "UserManagement", count: 3 },
      { category: "GroupManagement", count: 2 },
      { category: "RoleManagement", count: 1 },
    ]);
    expect(page.items[5]?.targets[0]?.properties[0]).toEqual({
      name: "Role.DisplayName",
      oldValue: null,
      newValue: "Global Administrator",
      oldCut: false,
      newCut: false,
    });
    expect((await feed({ days: 30 })).items).toHaveLength(7);
  });

  it("filters by category, result and text", async () => {
    expect((await feed({ category: "groupmanagement" })).items.map((item) => item.id)).toEqual([
      "d-timeout",
      "c-failed",
    ]);
    // Anything that did not succeed counts as a failure, timeouts included.
    expect((await feed({ result: "failure" })).items.map((item) => item.id)).toEqual([
      "d-timeout",
      "c-failed",
    ]);
    expect((await feed({ result: "success" })).items).toHaveLength(4);
    expect((await feed({ search: "ADD GROUP" })).items.map((item) => item.id)).toEqual([
      "c-failed",
    ]);
    // A target's name, and a value in what changed.
    expect((await feed({ search: "person@example" })).items.map((item) => item.id)).toEqual([
      "b-role",
    ]);
    expect((await feed({ search: "global administrator" })).items.map((item) => item.id)).toEqual([
      "b-role",
    ]);
    expect((await feed({ search: "synthetic admin" })).items).toHaveLength(5);
    // Counts stay those of the whole window.
    expect((await feed({ search: "add group" })).categories).toHaveLength(3);
  });

  it("pages with `before` without skipping or repeating changes in the same millisecond", async () => {
    const seen: string[] = [];
    let before: string | null = null;
    for (let page = 0; page < 10; page += 1) {
      const result = await feed({ limit: 1, ...(before ? { before } : {}) });
      seen.push(...result.items.map((item) => item.id));
      before = result.next;
      if (!before) break;
    }
    expect(seen).toEqual(["g-self", "f-same", "e-same", "d-timeout", "c-failed", "b-role"]);
    await expect(feed({ before: "not a position" })).rejects.toThrow(RangeError);
  });

  it("lists a person's changes as actor and as target, once each", async () => {
    const page = await database.tenant(tenantId, (sql) =>
      personChanges(sql, tenantId, personId.toUpperCase()),
    );
    expect(page.items.map((item) => [item.id, item.involvement])).toEqual([
      ["g-self", { actor: true, target: true }],
      ["b-role", { actor: false, target: true }],
    ]);
    const admin = await database.tenant(tenantId, (sql) =>
      personChanges(sql, tenantId, adminId, { limit: 2 }),
    );
    expect(admin.items.map((item) => [item.id, item.involvement.actor])).toEqual([
      ["f-same", true],
      ["e-same", true],
    ]);
    const older = await database.tenant(tenantId, (sql) =>
      personChanges(sql, tenantId, adminId, { limit: 2, before: admin.next }),
    );
    expect(older.items.map((item) => item.id)).toEqual(["d-timeout", "c-failed"]);
  });

  it("reports no coverage before collection, and the cursor's after", async () => {
    expect(await database.tenant(tenantId, (sql) => changeCoverage(sql, tenantId))).toEqual({
      from: ago(20 * DAY).toISOString(),
      through: null,
      lastCollected: now.toISOString(),
      phase: "none",
    });
    await database.tenant(tenantId, (sql) =>
      sql.query(
        `INSERT INTO prism365_cursor (tenant_id, module, collector, value, updated_at)
         VALUES ($1, 'directory.changes', 'audits', $2, $3)`,
        [
          tenantId,
          JSON.stringify({
            v: 1,
            phase: "history",
            start: ago(30 * DAY).toISOString(),
            through: ago(10 * DAY).toISOString(),
            settle: { at: ago(10 * DAY).toISOString() },
          }),
          now.toISOString(),
        ],
      ),
    );
    expect(await database.tenant(tenantId, (sql) => changeCoverage(sql, tenantId))).toEqual({
      from: ago(30 * DAY).toISOString(),
      through: ago(10 * DAY).toISOString(),
      lastCollected: now.toISOString(),
      phase: "history",
    });
  });
});

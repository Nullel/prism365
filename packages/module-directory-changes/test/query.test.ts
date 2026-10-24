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
  searchText,
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
         initiated_by_principal, targets, details, search_text, processed, collected_at)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12::jsonb,$13::jsonb,$14,true,$15)`,
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
        JSON.stringify(event.details),
        searchText(event),
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
    // A target's sign-in name. Changed values are not searched.
    expect((await feed({ search: "person@example" })).items.map((item) => item.id)).toEqual([
      "b-role",
    ]);
    expect((await feed({ search: "global administrator" })).items).toEqual([]);
    expect((await feed({ search: "synthetic admin" })).items).toHaveLength(5);
    // Counts stay those of the whole window.
    expect((await feed({ search: "add group" })).categories).toHaveLength(3);
  });

  it("searches names as a person reads them: no JSON keys, quotes and accents found", async () => {
    await store(
      change("h-quote", ago(40 * DAY), {
        targets: [
          {
            id: "group-1",
            name: 'Field "Ops" O\'Brien',
            type: "Group",
            principal: null,
            properties: [{ name: "Description", oldValue: 'R"D', newValue: "Sales" }],
          },
        ],
      }),
    );
    await store(
      change("i-accent", ago(41 * DAY), {
        targets: [
          {
            id: "user-9",
            name: "Élodie Synthétique",
            type: "User",
            principal: null,
            properties: [],
          },
        ],
      }),
    );
    const ids = async (search: string) =>
      (await feed({ days: 60, search })).items.map((item) => item.id);
    for (const key of ["properties", "oldvalue", "principal", "name", "null"])
      expect(await ids(key), key).toEqual([]);
    expect(await ids('"ops"')).toEqual(["h-quote"]);
    expect(await ids("o'brien")).toEqual(["h-quote"]);
    expect(await ids("élodie")).toEqual(["i-accent"]);
    expect(await ids("ÉLODIE".toLowerCase())).toEqual(["i-accent"]);
    expect(await ids('r"d')).toEqual([]);
  });

  it("finds names whatever their case: dotted capitals, final sigma, and the 30th of 50 targets", async () => {
    const named = (id: string, at: number, name: string) =>
      change(id, ago(at * DAY), {
        targets: [{ id: `user-${id}`, name, type: "User", principal: null, properties: [] }],
      });
    await store(named("k-dotted", 42, "İlker Örnek"));
    await store(named("l-greek", 43, "ΟΔΥΣΣΕΥΣ Δοκιμή"));
    await store(
      change("m-bulk", ago(44 * DAY), {
        activity: "Add member to group",
        targets: Array.from({ length: 50 }, (_value, index) => ({
          id: `bulk-${index}`,
          name: `Bulk Member ${String(index).padStart(2, "0")} ${"x".repeat(60)}`,
          type: "User",
          principal: `bulk${index}@example.invalid`,
          properties: [],
        })),
      }),
    );
    const ids = async (search: string) =>
      (await feed({ days: 60, search })).items.map((item) => item.id);
    for (const search of ["ilker", "İlker", "ILKER", "örnek", "ornek"])
      expect(await ids(search), search).toEqual(["k-dotted"]);
    for (const search of ["ΟΔΥΣ", "οδυσ", "οδυσσευς", "ΟΔΥΣΣΕΥΣ"])
      expect(await ids(search), search).toEqual(["l-greek"]);
    expect(await ids("bulk member 30")).toEqual(["m-bulk"]);
    expect(await ids("bulk49@example")).toEqual(["m-bulk"]);
  });

  it("cuts lists a row holds beyond today's bounds instead of failing the page", async () => {
    const many = change("j-many", ago(50 * DAY), {
      activity: "Bulk update",
      targets: Array.from({ length: 60 }, (_value, index) => ({
        id: `t-${index}`,
        name: `Target ${index}`,
        type: "User",
        principal: null,
        properties: Array.from({ length: 60 }, () => ({
          name: "Note",
          oldValue: null,
          newValue: "x",
        })),
      })),
      details: Array.from({ length: 25 }, (_value, index) => ({ key: `k${index}`, value: "v" })),
    });
    await store(many);
    const page = await feed({ days: 60, search: "bulk update" });
    const [item] = page.items;
    expect(item?.targets).toHaveLength(50);
    expect(item?.targets[0]?.properties).toHaveLength(50);
    expect(item?.details).toHaveLength(20);
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
      // The oldest change stored, before any cursor exists.
      from: ago(50 * DAY).toISOString(),
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
    // From the older of where reading started and the oldest change kept.
    expect(await database.tenant(tenantId, (sql) => changeCoverage(sql, tenantId))).toEqual({
      from: ago(50 * DAY).toISOString(),
      through: ago(10 * DAY).toISOString(),
      lastCollected: now.toISOString(),
      phase: "history",
    });
  });
});

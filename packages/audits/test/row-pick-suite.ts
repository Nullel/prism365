import { randomUUID } from "node:crypto";
import type { AuditRow } from "@prism365/contracts";
import type { ModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import {
  finishRun,
  insertRun,
  latestSucceededRun,
  pickRows,
  type RowPick,
  rowMatches,
} from "../src/index.ts";

// pickRows against a real database, shared by the PGlite run (row-pick.test.ts) and the
// PostgreSQL run (storage's postgres.test.ts), so both prove that the SQL row pick returns exactly
// what filtering the whole run in memory returns. Synthetic tenant and people only.

const tenant = randomUUID();
const otherTenant = randomUUID();
const actor = { id: randomUUID(), name: null };
const ana = "aaaaaaaa-0000-4000-8000-000000000001";
const ben = "aaaaaaaa-0000-4000-8000-000000000002";
const cai = "aaaaaaaa-0000-4000-8000-000000000003";

const row = (key: string, personId: string | null, cells: AuditRow["cells"], tags: string[] = []) =>
  ({ key, personId, tags, cells }) satisfies AuditRow;

const mailboxRows: AuditRow[] = [
  row("m1", null, { mailbox: "Ana.Example@example.invalid" }),
  row("m2", null, { mailbox: "ben@example.invalid" }),
  // Full-width letters fold to the same name only under NFKC, which the database cannot do.
  row("m3", null, { mailbox: "ＡＮＡ.example@example.invalid" }),
  row("m4", null, { mailbox: "0123456789abcdef0123456789abcdef" }),
  row("m5", null, { mailbox: 42 }),
  row("m6", null, { mailbox: "\tana.example@example.invalid\n" }),
  row("i1", null, { mailbox: "Ivan.Inks@example.invalid" }),
  row("i2", null, { mailbox: "ivan.inks@example.invalid" }),
];
const accountRows: AuditRow[] = [
  row("a1", ana, { person: { name: "Ana", principal: "ana@example.invalid" } }),
  row("b1", ben, {}),
  row("a2", ana, {}),
  row("n1", null, {}),
  row("c1", cai, {}),
  ...Array.from({ length: 30 }, (_value, index) => row(`a-more-${index}`, ana, {})),
];
const groupRows: AuditRow[] = [
  row("g1", null, { ownerIds: [ana, ben] }),
  row("g2", null, { ownerIds: [cai] }),
  row("g3", null, { ownerIds: null }),
  row("g4", null, { ownerIds: [ben, ana] }),
];
const appRows: AuditRow[] = [
  row("SP-1", null, { appId: "app-one" }),
  row("sp-2", null, { appId: "APP-TWO" }),
];

const entityRows: AuditRow[] = [
  row("E-One", null, { name: "Sales_team 100%", ids: ["AAA-1", "bbb-2"], single: "CCC-3" }),
  row("e-two", ana, { name: "Salesxteam 1000", ids: [], single: "ddd-4" }),
  row("e-three", null, { name: "Back\\slash", ids: "not a list", single: 7 }),
  row("e-four", ben, { name: "SALES lead", ids: ["bbb-2", "EEE-5"], single: null }),
  row("e-five", null, { ids: null }),
];

export function rowPickSuite(name: string, open: () => Promise<ModuleDatabase>, enabled = true) {
  describe.skipIf(!enabled)(`${name}: picking rows of the latest run in SQL`, () => {
    let database: ModuleDatabase;
    const store = async (
      tenantId: string,
      auditId: string,
      day: number,
      rows: AuditRow[],
      status: "succeeded" | "failed" = "succeeded",
    ) => {
      const id = randomUUID();
      await database.tenant(tenantId, async (sql) => {
        await insertRun(sql, tenantId, {
          id,
          auditId,
          auditVersion: 1,
          startedAt: new Date(Date.UTC(2026, 9, day, 9)),
          startedBy: actor,
        });
        await finishRun(
          sql,
          tenantId,
          id,
          new Date(Date.UTC(2026, 9, day, 10)),
          status === "succeeded"
            ? { status, complete: true, summary: [], notes: [], rows }
            : { status, code: "boom", message: "Failed." },
        );
      });
    };
    const pick = (auditId: string, how: RowPick, limit = 25, tenantId = tenant) =>
      database.tenant(tenantId, (sql) => pickRows(sql, tenantId, auditId, how, limit));
    // The reference answer: the whole latest run, filtered in memory.
    const reference = async (auditId: string, how: RowPick) => {
      const run = await database.tenant(tenant, (sql) => latestSucceededRun(sql, tenant, auditId));
      return (run?.rows ?? []).filter((candidate) => rowMatches(candidate, how)).map((r) => r.key);
    };
    const keys = (picked: Awaited<ReturnType<typeof pick>>) => picked?.rows.map((r) => r.key);

    beforeAll(async () => {
      database = await open();
      // An older run, which must never be read, then the real one, then a newer failed run.
      await store(tenant, "mail.mailbox-storage", 1, [
        row("old", null, { mailbox: "ana.example@example.invalid" }),
      ]);
      await store(tenant, "mail.mailbox-storage", 2, mailboxRows);
      await store(tenant, "mail.mailbox-storage", 3, [], "failed");
      await store(tenant, "accounts.activity", 2, accountRows);
      await store(tenant, "mail.names", 2, [
        ...Array.from({ length: 250 }, (_value, index) =>
          row(`n${index}`, null, { mailbox: `é${index}@example.invalid` }),
        ),
        row("target", null, { mailbox: "ＴＡＲＧＥＴ@example.invalid" }),
      ]);
      await store(tenant, "accounts.broken", 2, accountRows);
      await database.tenant(tenant, (sql) =>
        sql.query(
          `UPDATE prism365_audit_run SET result_rows = '{"not":"a list"}'::jsonb WHERE tenant_id=$1 AND audit_id='accounts.broken'`,
          [tenant],
        ),
      );
      await store(tenant, "groups.teams", 2, groupRows);
      await store(tenant, "apps.consents", 2, appRows);
      await store(tenant, "entities.demo", 2, entityRows);
      await store(tenant, "entities.roles", 2, [
        ...Array.from({ length: 41 }, (_value, index) =>
          row(`r${index}`, null, { name: "Crowd role", rid: "role-a" }),
        ),
        row("r41", null, { name: "Quiet role", rid: "role-b" }),
        row("r42", null, { name: "Other role", rid: "role-c" }),
        row("r43", null, { name: "Not it", rid: "role-d" }),
      ]);
      await store(tenant, "entities.many", 2, [
        ...Array.from({ length: 30 }, (_value, index) =>
          row(`p${index}`, null, { name: `Plain ${index}` }),
        ),
        row("odd", null, { name: "Odd_one 50%" }),
      ]);
      await store(otherTenant, "accounts.activity", 2, [row("x", ana, {})]);
    }, 120_000);
    afterAll(async () => {
      await database?.close();
    });

    it("picks a person's rows by account id, in the run's order, with the run's header", async () => {
      const how: RowPick = { by: "personId", personId: ana };
      const picked = await pick("accounts.activity", how, 100);
      expect(keys(picked)).toEqual(await reference("accounts.activity", how));
      expect(keys(picked)?.slice(0, 2)).toEqual(["a1", "a2"]);
      expect(picked?.run).toMatchObject({ auditId: "accounts.activity", complete: true });
      expect(picked?.more).toBe(false);
    });

    it("cuts at the limit and says there were more", async () => {
      const picked = await pick("accounts.activity", { by: "personId", personId: ana }, 25);
      expect(picked?.rows).toHaveLength(25);
      expect(picked?.more).toBe(true);
      expect(keys(await pick("accounts.activity", { by: "personId", personId: ben }))).toEqual([
        "b1",
      ]);
    });

    it("gives the header and no rows for a person the run does not mention", async () => {
      const picked = await pick("accounts.activity", { by: "personId", personId: randomUUID() });
      expect(picked?.rows).toEqual([]);
      expect(picked?.run.id).toBeTruthy();
    });

    it("is null for an audit that never succeeded, and keeps tenants apart", async () => {
      expect(await pick("mfa.registration", { by: "personId", personId: ana })).toBeNull();
      expect(
        keys(await pick("accounts.activity", { by: "personId", personId: ana }, 5, otherTenant)),
      ).toEqual(["x"]);
    });

    it("matches names the way the audits do: case, full-width letters and white space, never a hash", async () => {
      const how: RowPick = {
        by: "name",
        path: ["mailbox"],
        principal: "ana.example@EXAMPLE.invalid",
      };
      const picked = await pick("mail.mailbox-storage", how);
      expect(keys(picked)).toEqual(["m1", "m3", "m6"]);
      expect(keys(picked)).toEqual(await reference("mail.mailbox-storage", how));
      const hashed: RowPick = {
        by: "name",
        path: ["mailbox"],
        principal: "0123456789abcdef0123456789abcdef",
      };
      expect(keys(await pick("mail.mailbox-storage", hashed))).toEqual([]);
      expect(await reference("mail.mailbox-storage", hashed)).toEqual([]);
    });

    it("reads a name inside a nested cell", async () => {
      const how: RowPick = {
        by: "name",
        path: ["person", "principal"],
        principal: "ANA@example.invalid",
      };
      expect(keys(await pick("accounts.activity", how))).toEqual(["a1"]);
    });

    it("compares ASCII letters by code point, whatever the database's locale", async () => {
      // A Turkish or Azeri locale lower-cases I to a dotless i; the comparison must not.
      const how: RowPick = {
        by: "name",
        path: ["mailbox"],
        principal: "IVAN.INKS@example.invalid",
      };
      const picked = await pick("mail.mailbox-storage", how);
      expect(keys(picked)).toEqual(["i1", "i2"]);
      expect(keys(picked)).toEqual(await reference("mail.mailbox-storage", how));
    });

    it("reads the run whole when more cells than the candidate limit need a closer look", async () => {
      const how: RowPick = { by: "name", path: ["mailbox"], principal: "target@example.invalid" };
      // 250 cells with a character outside printable ASCII are all candidates, and the one that
      // folds to the name comes after them.
      const picked = await pick("mail.names", how);
      expect(keys(picked)).toEqual(["target"]);
      expect(keys(picked)).toEqual(await reference("mail.names", how));
    });

    it("reads a run whose rows are not a list as having none, and keeps the other audits working", async () => {
      expect(await pick("accounts.broken", { by: "personId", personId: ana })).toMatchObject({
        rows: [],
        more: false,
      });
      expect(keys(await pick("accounts.activity", { by: "personId", personId: ben }))).toEqual([
        "b1",
      ]);
    });

    it("finds a list cell that holds the id exactly", async () => {
      const how: RowPick = { by: "list", path: ["ownerIds"], value: ana };
      expect(keys(await pick("groups.teams", how))).toEqual(["g1", "g4"]);
      expect(keys(await pick("groups.teams", how))).toEqual(await reference("groups.teams", how));
      expect(
        keys(
          await pick("groups.teams", { by: "list", path: ["ownerIds"], value: ana.toUpperCase() }),
        ),
      ).toEqual([]);
    });

    it("finds rows by key or by a cell, ignoring case", async () => {
      const how: RowPick = {
        by: "keys",
        keys: ["sp-1"],
        cell: { path: ["appId"], values: ["app-two"] },
      };
      expect(keys(await pick("apps.consents", how))).toEqual(["SP-1", "sp-2"]);
      expect(keys(await pick("apps.consents", how))).toEqual(await reference("apps.consents", how));
      expect(keys(await pick("apps.consents", { by: "keys", keys: [] }))).toEqual([]);
    });
    it("finds rows by key, account id or an id cell, whether it is text or a list, ignoring case", async () => {
      const cases: RowPick[] = [
        { by: "ids", keys: ["e-one", "E-THREE"] },
        { by: "ids", personIds: [ana.toUpperCase()] },
        { by: "ids", cells: [{ path: ["ids"], values: ["BBB-2"] }] },
        { by: "ids", cells: [{ path: ["single"], values: ["ccc-3", "DDD-4"] }] },
        {
          by: "ids",
          keys: ["e-five"],
          personIds: [ben],
          cells: [
            { path: ["ids"], values: ["aaa-1"] },
            { path: ["single"], values: ["ddd-4"] },
          ],
        },
      ];
      const expected = [
        ["E-One", "e-three"],
        ["e-two"],
        ["E-One", "e-four"],
        ["E-One", "e-two"],
        ["E-One", "e-two", "e-four", "e-five"],
      ];
      for (const [index, how] of cases.entries()) {
        const picked = keys(await pick("entities.demo", how));
        expect(picked, `case ${index}`).toEqual(expected[index]);
        expect(picked, `case ${index} against memory`).toEqual(
          await reference("entities.demo", how),
        );
      }
    });

    it("finds nothing for an id pick with nothing to look for, and never reads a text cell as a list", async () => {
      const nothing: RowPick[] = [
        { by: "ids" },
        { by: "ids", keys: [], personIds: [], cells: [{ path: ["ids"], values: [] }] },
      ];
      for (const how of nothing) {
        expect(keys(await pick("entities.demo", how))).toEqual([]);
        expect(await reference("entities.demo", how)).toEqual([]);
      }
      // A cell that holds text is compared as text, and one that holds a number is not a match.
      const text: RowPick = { by: "ids", cells: [{ path: ["ids"], values: ["not a list"] }] };
      expect(keys(await pick("entities.demo", text))).toEqual(["e-three"]);
      expect(await reference("entities.demo", text)).toEqual(["e-three"]);
      const number: RowPick = { by: "ids", cells: [{ path: ["single"], values: ["7"] }] };
      expect(keys(await pick("entities.demo", number))).toEqual([]);
      expect(await pick("accounts.broken", { by: "ids", keys: ["a1"] })).toMatchObject({
        rows: [],
      });
      expect(await pick("mfa.registration", { by: "ids", keys: ["a1"] })).toBeNull();
    });

    it("finds a cell that holds the text, ignoring case, and treats % _ and the backslash as plain characters", async () => {
      const find = async (text: string) => {
        const how: RowPick = { by: "contains", path: ["name"], text };
        const picked = keys(await pick("entities.demo", how));
        expect(picked, text).toEqual(await reference("entities.demo", how));
        return picked;
      };
      expect(await find("sales")).toEqual(["E-One", "e-two", "e-four"]);
      expect(await find("Sales_t")).toEqual(["E-One"]);
      expect(await find("100%")).toEqual(["E-One"]);
      expect(await find("%")).toEqual(["E-One"]);
      expect(await find("_")).toEqual(["E-One"]);
      expect(await find("k\\s")).toEqual(["e-three"]);
      expect(await find("zzz")).toEqual([]);
    });

    it("lets the database narrow by the characters typed, so % and _ cannot crowd out the rows that match", async () => {
      // If LIKE read them as wildcards, the first rows of the run would fill the limit and the one
      // row that holds the characters would never come back.
      for (const text of ["_", "%", "_one 50%"]) {
        const picked = await pick("entities.many", { by: "contains", path: ["name"], text }, 3);
        expect(keys(picked), text).toEqual(["odd"]);
        expect(picked?.more, text).toBe(false);
      }
    });

    it("reads one row for each value of a cell when asked, in the run's order", async () => {
      const how: RowPick = { by: "contains", path: ["name"], text: "role", distinct: ["rid"] };
      const picked = await pick("entities.roles", how, 5);
      expect(keys(picked)).toEqual(["r0", "r41", "r42"]);
      expect(picked?.more).toBe(false);
      expect(keys(picked)).toEqual(
        await reference("entities.roles", { by: "contains", path: ["name"], text: "role" }).then(
          (all) => all.filter((key) => ["r0", "r41", "r42"].includes(key)),
        ),
      );
      // Cut by distinct roles, not by rows: forty rows of one role do not use up the limit.
      const cut = await pick("entities.roles", how, 2);
      expect(keys(cut)).toEqual(["r0", "r41"]);
      expect(cut?.more).toBe(true);
    });

    it("cuts a contains pick at the limit and says there were more", async () => {
      const picked = await pick("entities.demo", { by: "contains", path: ["name"], text: "s" }, 2);
      expect(picked?.rows).toHaveLength(2);
      expect(picked?.more).toBe(true);
    });
  });
}

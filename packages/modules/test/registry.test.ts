import { describe, expect, it } from "vitest";
import {
  ModuleDefinitionError,
  ModuleRegistry,
  modulePermissions,
  moduleTablePrefix,
  type PrismModule,
} from "../src/index.ts";
import { itemsMigration, probeModule } from "./fixtures.ts";

const valid = () => probeModule({ run: async () => undefined });

function withMigration(sql: string): PrismModule {
  return { ...valid(), migrations: [{ id: "001-items", sql }], retention: [] };
}

describe("module registry", () => {
  it("registers a valid module and derives its namespace, scope, tables, and permissions", () => {
    const registry = new ModuleRegistry().register(valid());
    expect(moduleTablePrefix("test.probe")).toBe("mod_test_probe_");
    expect(registry.migrationScopes()).toEqual([
      { scope: "module.test.probe", migrations: [itemsMigration] },
    ]);
    expect(registry.tables("test.probe")).toEqual(["mod_test_probe_item"]);
    expect(modulePermissions(valid())).toEqual(["User.Read.All"]);
  });

  it("rejects duplicate ids, bad ids, and packs outside the Graph catalog", () => {
    const registry = new ModuleRegistry().register(valid());
    expect(() => registry.register(valid())).toThrow("already registered");
    expect(() => new ModuleRegistry().register({ ...valid(), id: "Bad Id" })).toThrow(
      ModuleDefinitionError,
    );
    expect(() =>
      new ModuleRegistry().register({
        ...valid(),
        packs: ["exchange.configuration.read" as never],
      }),
    ).toThrow("unknown Microsoft Graph pack");
    expect(() => new ModuleRegistry().register({ ...valid(), packs: [] })).toThrow("at least one");
  });

  it("lets a module declare saved settings, but only in its own namespace", () => {
    const withSettings = (settings: string[]): PrismModule => ({ ...valid(), settings });
    expect(() => new ModuleRegistry().register(withSettings(["test.canaries"]))).not.toThrow();
    // Another feature's setting, a bare word, odd casing and a repeat are all refused.
    for (const bad of [
      ["attention"],
      ["audits.scope"],
      ["other.canaries"],
      ["test"],
      ["Test.Canaries"],
      ["test.canaries", "test.canaries"],
    ])
      expect(() => new ModuleRegistry().register(withSettings(bad)), bad.join()).toThrow(
        ModuleDefinitionError,
      );
  });

  it("bounds collector cadence and timeout", () => {
    const collector = valid().collectors[0];
    if (!collector) throw new Error("fixture collector missing");
    for (const bad of [
      { cadenceSeconds: 30 },
      { cadenceSeconds: 100_000 },
      { timeoutSeconds: 5 },
      { timeoutSeconds: 300 },
    ])
      expect(() =>
        new ModuleRegistry().register({ ...valid(), collectors: [{ ...collector, ...bad }] }),
      ).toThrow(ModuleDefinitionError);
  });

  it("refuses tables outside the module namespace", () => {
    expect(() =>
      new ModuleRegistry().register(
        withMigration(
          "CREATE TABLE mod_other_thing (tenant_id uuid NOT NULL); SELECT prism365_tenant_table('mod_other_thing');",
        ),
      ),
    ).toThrow("mod_test_probe_");
    expect(() =>
      new ModuleRegistry().register(
        withMigration(
          "CREATE TABLE mod_test_probe_x (tenant_id uuid NOT NULL); SELECT prism365_tenant_table('mod_test_probe_x'); INSERT INTO prism365_finding SELECT * FROM mod_test_probe_x;",
        ),
      ),
    ).toThrow("InsertStmt is not allowed");
    expect(() =>
      new ModuleRegistry().register(
        withMigration(
          "CREATE TABLE mod_test_probe_x (tenant_id uuid NOT NULL, parent text REFERENCES prism365_finding (id)); SELECT prism365_tenant_table('mod_test_probe_x');",
        ),
      ),
    ).toThrow("references prism365_finding");
  });

  it("refuses a table created without the tenant policy", () => {
    expect(() =>
      new ModuleRegistry().register(
        withMigration("CREATE TABLE mod_test_probe_open (tenant_id uuid NOT NULL);"),
      ),
    ).toThrow("without SELECT prism365_tenant_table");
  });

  it("refuses privilege, policy, function, view, drop, and session statements", () => {
    const table =
      "CREATE TABLE mod_test_probe_x (tenant_id uuid NOT NULL); SELECT prism365_tenant_table('mod_test_probe_x');";
    for (const statement of [
      "GRANT SELECT ON mod_test_probe_x TO PUBLIC;",
      "CREATE POLICY open ON mod_test_probe_x USING (true);",
      "ALTER TABLE mod_test_probe_x DISABLE ROW LEVEL SECURITY;",
      "CREATE FUNCTION mod_test_probe_f() RETURNS int LANGUAGE sql AS 'select 1';",
      "CREATE VIEW mod_test_probe_v AS SELECT * FROM mod_test_probe_x;",
      "DROP TABLE mod_test_probe_x;",
      "ALTER TABLE mod_test_probe_x DROP COLUMN tenant_id;",
      "SET search_path TO evil;",
      "CREATE ROLE intruder;",
      "DO $$ BEGIN END $$;",
      "SELECT pg_read_file('/etc/passwd');",
    ])
      expect(() => new ModuleRegistry().register(withMigration(`${table}\n${statement}`))).toThrow(
        ModuleDefinitionError,
      );
  });

  it("is not fooled by comments, quoting, or casing in migration SQL", () => {
    const table =
      "CREATE TABLE mod_test_probe_x (tenant_id uuid NOT NULL); SELECT prism365_tenant_table('mod_test_probe_x');";
    for (const statement of [
      "CREATE/**/POLICY leak ON prism365_finding FOR SELECT USING (true);",
      "create   policy leak on mod_test_probe_x using (true);",
      'CREATE FUNCTION "mod_test_probe_f"() RETURNS int LANGUAGE sql SECURITY/**/DEFINER AS $$ SELECT 1 $$;',
      'ALTER TABLE "mod_test_probe_x" NO FORCE ROW LEVEL SECURITY;',
      "SELECT prism365_tenant_table('mod_test_probe_x'), set_config('role', 'postgres', false);",
      "CREATE TABLE mod_test_probe_y (tenant_id uuid NOT NULL DEFAULT (query_to_xml('select 1', true, true, '')::text)::uuid); SELECT prism365_tenant_table('mod_test_probe_y');",
    ])
      expect(
        () => new ModuleRegistry().register(withMigration(`${table}\n${statement}`)),
        statement,
      ).toThrow(ModuleDefinitionError);
  });

  it("accepts ordinary foreign-key actions and column names that resemble keywords", () => {
    const sql = `CREATE TABLE mod_test_probe_parent (tenant_id uuid NOT NULL, id text NOT NULL, PRIMARY KEY (tenant_id, id));
      SELECT prism365_tenant_table('mod_test_probe_parent');
      CREATE TABLE mod_test_probe_child (
        tenant_id uuid NOT NULL, id text NOT NULL, parent_id text, policy_name text, role_name text,
        FOREIGN KEY (tenant_id, parent_id) REFERENCES mod_test_probe_parent (tenant_id, id) ON DELETE CASCADE ON UPDATE CASCADE
      );
      SELECT prism365_tenant_table('mod_test_probe_child');`;
    expect(() => new ModuleRegistry().register(withMigration(sql))).not.toThrow();
  });

  it("requires retention rules to name a table the module creates", () => {
    expect(() =>
      new ModuleRegistry().register({
        ...valid(),
        retention: [{ table: "missing", timeColumn: "created_at", days: 30 }],
      }),
    ).toThrow("no migration creates");
    expect(() =>
      new ModuleRegistry().register({
        ...valid(),
        retention: [{ table: "item", timeColumn: "created_at; drop", days: 30 }],
      }),
    ).toThrow("plain identifiers");
  });

  it("requires ordered, unique migration ids", () => {
    expect(() =>
      new ModuleRegistry().register({
        ...valid(),
        migrations: [
          { ...itemsMigration, id: "002-b" },
          { ...itemsMigration, id: "001-a" },
        ],
      }),
    ).toThrow("ascending");
  });
});

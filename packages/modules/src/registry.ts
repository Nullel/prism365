import { type GraphFeatureId, permissionPacks } from "@prism365/graph";
import {
  functionName,
  isAllowedFunction,
  type MigrationScope,
  type ParseNode,
  parseStatements,
  walkParseTree,
} from "@prism365/storage";
import type { PrismModule } from "./contract.ts";

export class ModuleDefinitionError extends Error {
  constructor(moduleId: string, message: string) {
    super(`Module ${moduleId}: ${message}`);
    this.name = "ModuleDefinitionError";
  }
}

const moduleIdPattern = /^[a-z][a-z0-9]{1,30}(?:\.[a-z][a-z0-9-]{1,30}){0,2}$/;
const collectorIdPattern = /^[a-z][a-z0-9-]{1,40}$/;
const migrationIdPattern = /^\d{3}-[a-z0-9-]{1,40}$/;
const identifierPattern = /^[a-z][a-z0-9_]{0,40}$/;
const settingKeyPattern = /^[a-z][a-z0-9]{1,30}(?:\.[a-z][a-z0-9-]{1,30}){1,2}$/;

export function moduleTablePrefix(moduleId: string): string {
  return `mod_${moduleId.replace(/[.-]/g, "_")}_`;
}

export function moduleScope(moduleId: string): string {
  return `module.${moduleId}`;
}

// Module SQL is first-party code, but it is checked like untrusted input, on PostgreSQL's own
// parse tree: comments, quoting and casing cannot change what a statement is. A migration may
// only create tables and indexes in its own namespace, add columns, defaults, NOT NULL or
// constraints to them, and apply the tenant policy with prism365_tenant_table(). Every relation
// any statement names, including foreign-key targets, must be in the module's namespace.
const allowedAlterations = new Set([
  "AT_AddColumn",
  "AT_ColumnDefault",
  "AT_AddConstraint",
  "AT_SetNotNull",
]);

function validateMigrationSql(module: PrismModule, migrationId: string, sql: string): string[] {
  const prefix = moduleTablePrefix(module.id);
  const fail = (reason: string): never => {
    throw new ModuleDefinitionError(module.id, `migration ${migrationId}: ${reason}`);
  };
  let statements: ParseNode[];
  try {
    statements = parseStatements(sql);
  } catch {
    return fail("could not be parsed as PostgreSQL.");
  }
  const created: string[] = [];
  const isolated = new Set<string>();
  const inNamespace = (name: unknown) => typeof name === "string" && name.startsWith(prefix);

  for (const statement of statements) {
    const [kind = "", body = {}] = Object.entries(statement)[0] ?? [];
    const node = body as ParseNode;

    if (kind === "SelectStmt") {
      const targets = node.targetList as ParseNode[] | undefined;
      const call = (targets?.[0]?.ResTarget as ParseNode | undefined)?.val as ParseNode | undefined;
      const funcCall = call?.FuncCall as ParseNode | undefined;
      const args = funcCall?.args as ParseNode[] | undefined;
      const table = ((args?.[0]?.A_Const as ParseNode | undefined)?.sval as ParseNode | undefined)
        ?.sval;
      const otherClauses = Object.keys(node).filter(
        (key) => !["targetList", "limitOption", "op"].includes(key),
      );
      if (
        targets?.length !== 1 ||
        otherClauses.length > 0 ||
        functionName(funcCall?.funcname).join(".") !== "prism365_tenant_table" ||
        args?.length !== 1 ||
        !inNamespace(table)
      )
        fail("the only SELECT allowed is SELECT prism365_tenant_table('<module table>').");
      if (!created.includes(String(table)))
        fail(`prism365_tenant_table('${String(table)}') must follow that table's CREATE TABLE.`);
      isolated.add(String(table));
      continue;
    }

    if (kind === "CreateStmt") {
      const relation = node.relation as ParseNode;
      if (
        !inNamespace(relation.relname) ||
        (relation.schemaname && relation.schemaname !== "public")
      )
        fail(`tables must be named ${prefix}<name> in the public schema.`);
      if (
        node.inhRelations ||
        node.partbound ||
        node.partspec ||
        node.ofTypename ||
        node.accessMethod
      )
        fail("inheritance, partitioning, typed tables and access methods are not allowed.");
      if (relation.relpersistence !== "p") fail("temporary and unlogged tables are not allowed.");
      created.push(String(relation.relname));
    } else if (kind === "IndexStmt") {
      const relation = node.relation as ParseNode;
      if (!inNamespace(node.idxname) || !inNamespace(relation.relname))
        fail(`indexes and their tables must be named ${prefix}<name>.`);
    } else if (kind === "AlterTableStmt") {
      const relation = node.relation as ParseNode;
      if (!inNamespace(relation.relname) || node.objtype !== "OBJECT_TABLE")
        fail(`only ${prefix} tables may be altered.`);
      for (const command of (node.cmds as ParseNode[] | undefined) ?? []) {
        const subtype = String((command.AlterTableCmd as ParseNode | undefined)?.subtype ?? "");
        if (!allowedAlterations.has(subtype))
          fail(
            `${subtype || "that change"} is not allowed; add a reviewed core migration instead.`,
          );
      }
    } else {
      fail(
        `${kind || "that statement"} is not allowed; only tables, indexes and the tenant policy.`,
      );
    }

    walkParseTree(node, (key, value) => {
      if (typeof value.relname === "string" && !inNamespace(value.relname))
        fail(`references ${value.relname}; module objects must start with ${prefix}.`);
      if (key === "FuncCall" && !isAllowedFunction(functionName(value.funcname)))
        fail(
          `calls ${functionName(value.funcname).join(".")}, which is not allowed in a migration.`,
        );
    });
  }

  for (const table of created)
    if (!isolated.has(table))
      fail(`creates ${table} without SELECT prism365_tenant_table('${table}').`);
  return created;
}

export function validateModule(module: PrismModule): string[] {
  const id = typeof module.id === "string" ? module.id : "(invalid)";
  if (!moduleIdPattern.test(id))
    throw new ModuleDefinitionError(id, "id must look like area.name in lowercase.");
  if (moduleTablePrefix(id).length > 40)
    throw new ModuleDefinitionError(id, "id is too long for its table namespace.");
  if (!module.title?.trim() || !module.description?.trim())
    throw new ModuleDefinitionError(id, "title and description are required.");
  if (!Number.isInteger(module.version) || module.version < 1)
    throw new ModuleDefinitionError(id, "version must be a positive integer.");

  if (module.packs.length === 0)
    throw new ModuleDefinitionError(id, "declare at least one read permission pack.");
  for (const packId of module.packs) {
    const pack = permissionPacks.find((candidate) => candidate.id === packId);
    if (pack?.source !== "microsoftGraph")
      throw new ModuleDefinitionError(id, `unknown Microsoft Graph pack ${String(packId)}.`);
  }

  if (module.collectors.length === 0)
    throw new ModuleDefinitionError(id, "declare at least one collector.");
  const collectorIds = new Set<string>();
  for (const collector of module.collectors) {
    if (!collectorIdPattern.test(collector.id) || collectorIds.has(collector.id))
      throw new ModuleDefinitionError(id, `collector id ${collector.id} is invalid or repeated.`);
    collectorIds.add(collector.id);
    if (
      !Number.isInteger(collector.cadenceSeconds) ||
      collector.cadenceSeconds < 60 ||
      collector.cadenceSeconds > 86_400
    )
      throw new ModuleDefinitionError(id, `collector ${collector.id} cadence must be 60s to 24h.`);
    if (
      !Number.isInteger(collector.timeoutSeconds) ||
      collector.timeoutSeconds < 10 ||
      collector.timeoutSeconds > 900 ||
      collector.timeoutSeconds >= collector.cadenceSeconds
    )
      throw new ModuleDefinitionError(
        id,
        `collector ${collector.id} timeout must be 10s to 15min and shorter than its cadence.`,
      );
  }

  // The namespace is the module id up to its first dot, so "signins.audit" can read only
  // "signins." keys. Module ids are first-party and the registry refuses a repeated id, so no
  // module is named after a core setting's namespace such as "audits" or "attention".
  const namespace = `${id.split(".")[0]}.`;
  const settings = new Set<string>();
  for (const key of module.settings ?? []) {
    if (!settingKeyPattern.test(key) || !key.startsWith(namespace) || settings.has(key))
      throw new ModuleDefinitionError(
        id,
        `setting ${key} is invalid, repeated, or outside the ${namespace} namespace.`,
      );
    settings.add(key);
  }

  const migrationIds = new Set<string>();
  const tables = new Set<string>();
  let previous = "";
  for (const migration of module.migrations) {
    if (!migrationIdPattern.test(migration.id) || migrationIds.has(migration.id))
      throw new ModuleDefinitionError(id, `migration id ${migration.id} is invalid or repeated.`);
    if (migration.id <= previous)
      throw new ModuleDefinitionError(id, "migrations must be listed in ascending order.");
    previous = migration.id;
    migrationIds.add(migration.id);
    for (const table of validateMigrationSql(module, migration.id, migration.sql))
      tables.add(table);
  }

  for (const rule of module.retention) {
    if (!identifierPattern.test(rule.table) || !identifierPattern.test(rule.timeColumn))
      throw new ModuleDefinitionError(id, "retention table and column must be plain identifiers.");
    if (!tables.has(`${moduleTablePrefix(id)}${rule.table}`))
      throw new ModuleDefinitionError(
        id,
        `retention names ${rule.table}, which no migration creates.`,
      );
    if (!Number.isInteger(rule.days) || rule.days < 1 || rule.days > 3_650)
      throw new ModuleDefinitionError(id, "retention days must be 1 to 3650.");
  }
  return [...tables];
}

export class ModuleRegistry {
  readonly #modules = new Map<string, PrismModule>();
  readonly #tables = new Map<string, readonly string[]>();

  register(module: PrismModule): this {
    const tables = validateModule(module);
    if (this.#modules.has(module.id))
      throw new ModuleDefinitionError(module.id, "is already registered.");
    this.#modules.set(module.id, Object.freeze({ ...module }));
    this.#tables.set(module.id, Object.freeze([...tables]));
    return this;
  }

  // Tables the module's migrations create, for startup checks that migrations were applied.
  tables(id: string): readonly string[] {
    return this.#tables.get(id) ?? [];
  }

  list(): PrismModule[] {
    return [...this.#modules.values()];
  }

  get(id: string): PrismModule | undefined {
    return this.#modules.get(id);
  }

  migrationScopes(): MigrationScope[] {
    return this.list().map((module) => ({
      scope: moduleScope(module.id),
      migrations: module.migrations,
    }));
  }
}

// Microsoft permissions a module needs, for the credential kind collecting it.
export function modulePermissions(module: PrismModule): string[] {
  return [
    ...new Set(
      module.packs.map((packId: GraphFeatureId) => {
        const pack = permissionPacks.find((candidate) => candidate.id === packId);
        if (!pack) throw new ModuleDefinitionError(module.id, `unknown pack ${packId}.`);
        return pack.permission;
      }),
    ),
  ];
}

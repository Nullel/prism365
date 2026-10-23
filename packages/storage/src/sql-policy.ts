import { loadModule, parseSync } from "libpg-query";

// SQL is checked on PostgreSQL's own parse tree (libpg-query, the real parser compiled to
// WebAssembly), never on its text. Quoting, comments, casing and whitespace cannot change what
// a statement is once it has been parsed, which is why the earlier text patterns were replaced.
await loadModule();

export class SqlGuardError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "SqlGuardError";
  }
}

export interface TableAccess {
  read: ReadonlySet<string>;
  write: ReadonlySet<string>;
}

export type ParseNode = Record<string, unknown>;
type Node = ParseNode;

export function parseStatements(text: string): Node[] {
  let tree: { stmts?: { stmt?: Node }[] };
  try {
    tree = parseSync(text) as typeof tree;
  } catch {
    throw new SqlGuardError("The SQL statement could not be parsed.");
  }
  return (tree.stmts ?? []).map((entry) => entry.stmt ?? {});
}

// Functions a module or core query may call. Everything else, including set_config,
// current_setting, query_to_xml, ts_stat and every pg_* function, is refused by name after
// parsing, so quoting it ("set_config") or splitting it with comments changes nothing.
const allowedFunctions: ReadonlySet<string> = new Set([
  "count",
  "sum",
  "min",
  "max",
  "avg",
  "bool_and",
  "bool_or",
  "every",
  "array_agg",
  "string_agg",
  "jsonb_agg",
  "json_agg",
  "jsonb_build_object",
  "jsonb_array_length",
  "jsonb_typeof",
  "lower",
  "upper",
  "length",
  "char_length",
  "btrim",
  "ltrim",
  "rtrim",
  "substring",
  "substr",
  "position",
  "strpos",
  "replace",
  "split_part",
  "concat",
  "abs",
  "round",
  "floor",
  "ceil",
  "ceiling",
  "now",
  "extract",
  "date_part",
  "date_trunc",
  "date_bin",
  "timezone",
  "make_interval",
  "to_char",
  "to_regclass",
]);

// True for an allowlisted function, optionally qualified with pg_catalog.
export function isAllowedFunction(parts: readonly string[]): boolean {
  const name = parts.at(-1) ?? "";
  const schema = parts.length === 2 ? parts[0] : undefined;
  return (
    parts.length <= 2 &&
    (schema === undefined || schema === "pg_catalog") &&
    allowedFunctions.has(name)
  );
}

const dmlStatements = new Set(["SelectStmt", "InsertStmt", "UpdateStmt", "DeleteStmt"]);
const reservedPrefixes = ["prism365_", "mod_", "pg_"];

export function walkParseTree(
  node: unknown,
  visit: (key: string, value: Node, parent: Node) => void,
  parent: Node = {},
) {
  if (Array.isArray(node)) {
    for (const item of node) walkParseTree(item, visit, parent);
    return;
  }
  if (!node || typeof node !== "object") return;
  const object = node as Node;
  for (const [key, value] of Object.entries(object)) {
    if (value && typeof value === "object" && !Array.isArray(value))
      visit(key, value as Node, object);
    walkParseTree(value, visit, object);
  }
}

export function functionName(list: unknown): string[] {
  return Array.isArray(list)
    ? list.map((item) => String((item as { String?: { sval?: unknown } })?.String?.sval ?? ""))
    : [];
}

// Allows exactly one SELECT, INSERT, UPDATE or DELETE (CTEs included) that reads only tables in
// access.read or access.write, writes only tables in access.write, and calls only allowlisted
// functions. Throws SqlGuardError otherwise.
export function assertRuntimeStatement(text: string, access: TableAccess): void {
  if (text.length > 20_000) throw new SqlGuardError("The SQL statement exceeds the size budget.");
  const statements = parseStatements(text);
  if (statements.length !== 1)
    throw new SqlGuardError(
      "Runtime SQL must be exactly one statement; pass values as parameters.",
    );
  const [root] = statements;
  const kind = Object.keys(root ?? {})[0] ?? "";
  if (!dmlStatements.has(kind))
    throw new SqlGuardError("Runtime SQL may only SELECT, INSERT, UPDATE, or DELETE rows.");

  const cteNames = new Set<string>();
  walkParseTree(root, (key, value) => {
    if (key !== "CommonTableExpr") return;
    const name = String(value.ctename ?? "");
    if (
      !/^[a-z_][a-z0-9_]*$/.test(name) ||
      reservedPrefixes.some((prefix) => name.startsWith(prefix)) ||
      access.read.has(name) ||
      access.write.has(name)
    )
      throw new SqlGuardError("A common table expression uses a reserved or table name.");
    cteNames.add(name);
  });

  walkParseTree(root, (key, value) => {
    // Node types are capitalized; lower-case keys such as "selectStmt" are properties.
    if (/^[A-Z]/.test(key) && key.endsWith("Stmt") && !dmlStatements.has(key))
      throw new SqlGuardError("Runtime SQL contains a statement that is not a row read or write.");
    if (key === "SelectStmt" && value.intoClause)
      throw new SqlGuardError("SELECT INTO is not allowed.");
    if ((key === "InsertStmt" || key === "UpdateStmt" || key === "DeleteStmt") && value.relation) {
      const target = value.relation as Node;
      const name = String(target.relname ?? "");
      if ((target.schemaname && target.schemaname !== "public") || !access.write.has(name))
        throw new SqlGuardError(`Writing to ${name || "that relation"} is not allowed here.`);
    }
    if (key === "FuncCall") {
      const parts = functionName(value.funcname).map((part) => part.toLowerCase());
      if (!isAllowedFunction(parts))
        throw new SqlGuardError(`Calling ${parts.join(".") || "that function"} is not allowed.`);
    }
    if (typeof value.relname === "string") {
      const name = value.relname;
      if (value.schemaname !== undefined && value.schemaname !== "public")
        throw new SqlGuardError("Only tables in the public schema may be referenced.");
      if (value.catalogname !== undefined)
        throw new SqlGuardError("Cross-database references are not allowed.");
      if (!cteNames.has(name) && !access.read.has(name) && !access.write.has(name))
        throw new SqlGuardError(`Reading ${name} is not allowed here.`);
    }
  });
}

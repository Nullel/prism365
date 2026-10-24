import { createHash } from "node:crypto";
import { open, readFile, unlink } from "node:fs/promises";
import { join } from "node:path";
import { PGlite, type Transaction } from "@electric-sql/pglite";
import { Pool, type PoolClient } from "pg";
import { assertRuntimeStatement, SqlGuardError, type TableAccess } from "./sql-policy.ts";

export {
  assertRuntimeStatement,
  parseStatements,
  SqlGuardError,
  type TableAccess,
} from "./sql-policy.ts";

// Module data lives in ordinary tenant-scoped tables. Two backends share one set of SQL
// migrations: PGlite (PostgreSQL compiled to WebAssembly, in process) for single-operator
// installs and tests, and PostgreSQL for deployments.
//
// Isolation has three layers. (1) Every statement passes the parse-tree guard in sql-policy.ts,
// which limits it to one row read or write over the tables the caller is allowed, with
// allowlisted functions only. (2) Each transaction runs as the restricted prism365_app role with
// prism365.tenant_id pinned, and forced row-level security filters every row by tenant.
// (3) Startup refuses any tenant table whose policy is not the canonical one.
//
// On PostgreSQL the runtime role also lacks schema and privilege rights. On PGlite the session
// user is a superuser that SET LOCAL ROLE only masks; there, layer (1) is the boundary. Prism365
// connects one tenant per installation, so a PGlite install never holds two tenants' data.

export interface SqlResult<T> {
  rows: T[];
  rowCount: number;
}

export interface SqlClient {
  query<T extends object = Record<string, unknown>>(
    text: string,
    params?: readonly unknown[],
  ): Promise<SqlResult<T>>;
}

export interface Migration {
  id: string;
  sql: string;
}

export interface MigrationScope {
  scope: string;
  migrations: readonly Migration[];
}

export interface TenantOptions {
  // Tables this transaction may read and write. Defaults to every tenant table except the
  // session store, the access ledger, and the migration ledger.
  access?: TableAccess;
  // Aborting rolls the transaction back and disables its client, even if the operation is stuck.
  signal?: AbortSignal;
  // Upper bound for the whole transaction. Default 60 seconds.
  timeoutMilliseconds?: number;
}

export interface ModuleDatabase {
  readonly kind: "pglite" | "postgres";
  readonly durable: boolean;
  tenant<T>(
    tenantId: string,
    operation: (sql: SqlClient) => Promise<T>,
    options?: TenantOptions,
  ): Promise<T>;
  verify(): Promise<void>;
  close(): Promise<void>;
}

export class MigrationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "MigrationError";
  }
}

export class TransactionAbortedError extends Error {
  constructor(reason: "aborted" | "timeout") {
    super(
      reason === "timeout"
        ? "The transaction exceeded its time limit."
        : "The transaction was aborted.",
    );
    this.name = "TransactionAbortedError";
  }
}

const tenantPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function validateTenant(tenantId: string): string {
  if (!tenantPattern.test(tenantId)) throw new Error("Module storage requires a tenant UUID.");
  return tenantId.toLowerCase();
}

// Tables the core never exposes to module or report queries.
const protectedTables: ReadonlySet<string> = new Set([
  "prism365_state",
  "prism365_audit",
  "prism365_schema_migration",
]);

// A client that works only while its transaction is open. A reference kept after the callback
// returns (a missing await, a fire-and-forget) throws instead of running on whatever pooled
// connection, and therefore whichever tenant, comes next.
function guardedClient(
  access: TableAccess,
  run: (text: string, params: unknown[]) => Promise<SqlResult<never>>,
): { client: SqlClient; close: () => void } {
  let open = true;
  return {
    client: {
      async query<T extends object>(text: string, params: readonly unknown[] = []) {
        if (!open)
          throw new SqlGuardError("This database client belongs to a finished transaction.");
        assertRuntimeStatement(text, access);
        return (await run(text, [...params])) as SqlResult<T>;
      },
    },
    close: () => {
      open = false;
    },
  };
}

// Runs the operation but gives up when the signal aborts or the time limit passes. Giving up
// makes the transaction callback throw, which rolls it back and releases the connection; the
// abandoned operation's client is already closed, so nothing it does later reaches storage.
async function bounded<T>(operation: Promise<T>, options: TenantOptions): Promise<T> {
  operation.catch(() => {});
  const limit = options.timeoutMilliseconds ?? 60_000;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let onAbort: (() => void) | undefined;
  try {
    return await Promise.race([
      operation,
      new Promise<never>((_resolve, reject) => {
        timer = setTimeout(() => reject(new TransactionAbortedError("timeout")), limit);
        if (options.signal) {
          if (options.signal.aborted) reject(new TransactionAbortedError("aborted"));
          onAbort = () => reject(new TransactionAbortedError("aborted"));
          options.signal.addEventListener("abort", onAbort, { once: true });
        }
      }),
    ]);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
    if (onAbort) options.signal?.removeEventListener("abort", onAbort);
  }
}

export function migrationChecksum(sql: string): string {
  return createHash("sha256").update(sql.replaceAll("\r\n", "\n")).digest("hex");
}

export async function loadCoreMigrations(): Promise<Migration[]> {
  // Applied in this order on every start. Numbers need not be contiguous: each slice of work
  // claims its own number, and the list, not the numbering, decides what runs.
  const files = [
    "001-state.sql",
    "002-modules.sql",
    "003-attention.sql",
    "004-audit-runs.sql",
    "005-audit-trends.sql",
    "006-handoffs.sql",
    "007-delivery.sql",
  ];
  return Promise.all(
    files.map(async (file) => ({
      id: file.replace(/\.sql$/, ""),
      sql: await readFile(new URL(`../migrations/${file}`, import.meta.url), "utf8"),
    })),
  );
}

// The one tenant policy, as PostgreSQL 17 and PGlite (PostgreSQL 18) both render it.
const canonicalPolicy = "((tenant_id)::text = current_setting('prism365.tenant_id'::text, true))";
const policyNames = ["prism365_tenant_isolation", "prism365_state_tenant", "prism365_audit_tenant"];

const verifyTablesSql = `
  SELECT c.relname AS name, c.relkind::text AS kind, c.relrowsecurity AS rls,
         c.relforcerowsecurity AS forced,
         (SELECT count(*)::int FROM pg_policy p WHERE p.polrelid = c.oid) AS policies,
         (SELECT count(*)::int FROM pg_policy p
           WHERE p.polrelid = c.oid AND p.polname = ANY($1::text[]) AND p.polpermissive
             AND p.polcmd = '*' AND p.polroles = '{0}'::oid[]
             AND pg_get_expr(p.polqual, p.polrelid) = $2
             AND pg_get_expr(p.polwithcheck, p.polrelid) = $2) AS canonical,
         (SELECT count(*)::int FROM pg_trigger t WHERE t.tgrelid = c.oid AND NOT t.tgisinternal) AS triggers,
         EXISTS (SELECT 1 FROM pg_attribute a WHERE a.attrelid = c.oid AND a.attname = 'tenant_id'
                 AND a.attnotnull AND NOT a.attisdropped
                 AND a.atttypid = 'uuid'::regtype) AS has_tenant
    FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace
   WHERE n.nspname = 'public' AND c.relkind IN ('r','p','v','m','f')
     AND (c.relname LIKE 'prism365\\_%' OR c.relname LIKE 'mod\\_%')
     AND c.relname <> 'prism365_schema_migration'`;

// Functions in the public schema other than the core helper mean someone added code that could
// run with its owner's rights; definer functions in particular bypass row security.
const verifyFunctionsSql = `
  SELECT p.proname AS name, p.prosecdef AS definer
    FROM pg_proc p JOIN pg_namespace n ON n.oid = p.pronamespace
   WHERE n.nspname = 'public' AND (p.prosecdef OR p.proname <> 'prism365_tenant_table')`;

interface TableCheck {
  name: string;
  kind: string;
  rls: boolean;
  forced: boolean;
  policies: number;
  canonical: number;
  triggers: number;
  has_tenant: boolean;
}

function assertSchemaIsolated(
  tables: readonly TableCheck[],
  functions: readonly { name: string }[],
) {
  const unsafe = tables.filter(
    (row) =>
      row.kind !== "r" ||
      !row.rls ||
      !row.forced ||
      row.policies !== 1 ||
      row.canonical !== 1 ||
      row.triggers !== 0 ||
      !row.has_tenant,
  );
  if (unsafe.length > 0)
    throw new Error(
      `Tenant tables without exactly the canonical tenant policy: ${unsafe.map((row) => row.name).join(", ")}. Apply prism365_tenant_table() and remove other policies, views, or triggers.`,
    );
  if (functions.length > 0)
    throw new Error(
      `Unexpected functions in the public schema: ${functions.map((row) => row.name).join(", ")}.`,
    );
}

function coreAccess(tables: readonly TableCheck[]): TableAccess {
  const names = new Set(
    tables.map((table) => table.name).filter((name) => !protectedTables.has(name)),
  );
  return { read: names, write: names };
}

interface OwnerExecutor {
  exec(sql: string): Promise<void>;
  query<T>(sql: string, params: unknown[]): Promise<T[]>;
  transaction(operation: (executor: OwnerExecutor) => Promise<void>): Promise<void>;
}

// Applies migrations once per scope. A migration whose text changed after it was applied is a
// hard error: silently re-running edited SQL against live tenant data is how schemas drift.
// Core migrations (001, 002) are idempotent and re-run on every start instead.
async function applyLedgerMigrations(owner: OwnerExecutor, scopes: readonly MigrationScope[]) {
  for (const { scope, migrations } of scopes) {
    if (!/^[a-z][a-z0-9.-]{1,80}$/.test(scope))
      throw new MigrationError("Invalid migration scope.");
    for (const migration of migrations) {
      const checksum = migrationChecksum(migration.sql);
      const existing = await owner.query<{ checksum: string }>(
        "SELECT checksum FROM prism365_schema_migration WHERE scope=$1 AND migration_id=$2",
        [scope, migration.id],
      );
      const applied = existing[0];
      if (applied) {
        if (applied.checksum !== checksum)
          throw new MigrationError(
            `Migration ${scope}/${migration.id} changed after it was applied. Add a new migration instead.`,
          );
        continue;
      }
      try {
        await owner.transaction(async (tx) => {
          await tx.exec(migration.sql);
          await tx.query(
            "INSERT INTO prism365_schema_migration(scope,migration_id,checksum) VALUES($1,$2,$3)",
            [scope, migration.id, checksum],
          );
        });
      } catch {
        // The identifiers are ours; the driver message may carry connection details.
        throw new MigrationError(`Migration ${scope}/${migration.id} failed and was rolled back.`);
      }
    }
  }
}

function pgliteExecutor(db: PGlite | Transaction): OwnerExecutor {
  return {
    async exec(sql) {
      await db.exec(sql);
    },
    async query<T>(sql: string, params: unknown[]) {
      return (await db.query<T>(sql, params)).rows;
    },
    async transaction(operation) {
      if (!("transaction" in db)) throw new Error("Nested migration transactions are not used.");
      await db.transaction((tx) => operation(pgliteExecutor(tx)));
    },
  };
}

function processAlive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    return (error as NodeJS.ErrnoException).code === "EPERM";
  }
}

// One process per data directory. Two PGlite instances on the same files corrupt them silently.
async function lockDataDir(dataDir: string): Promise<string> {
  const path = join(dataDir, "prism365.lock");
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      const handle = await open(path, "wx", 0o600);
      await handle.writeFile(String(process.pid));
      await handle.close();
      return path;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
      const owner = Number((await readFile(path, "utf8").catch(() => "")).trim());
      if (Number.isInteger(owner) && owner > 0 && owner !== process.pid && processAlive(owner))
        throw new Error(
          `PRISM_DATA_DIR is in use by process ${owner}. Stop it before starting another.`,
        );
      await unlink(path).catch(() => {});
    }
  }
  throw new Error("Could not lock PRISM_DATA_DIR.");
}

export class PGliteModuleDatabase implements ModuleDatabase {
  readonly kind = "pglite" as const;
  readonly durable: boolean;
  readonly #db: PGlite;
  readonly #lock: string | null;
  #access: TableAccess = { read: new Set(), write: new Set() };

  private constructor(db: PGlite, durable: boolean, lock: string | null) {
    this.#db = db;
    this.durable = durable;
    this.#lock = lock;
  }

  // dataDir omitted: an in-memory database that disappears with the process. Set it to keep
  // module data across restarts; the directory then holds tenant data and must be protected.
  static async open(
    options: { dataDir?: string; scopes?: readonly MigrationScope[] } = {},
  ): Promise<PGliteModuleDatabase> {
    const lock = options.dataDir ? await lockDataDir(options.dataDir) : null;
    let db: PGlite | undefined;
    try {
      db = options.dataDir ? await PGlite.create(options.dataDir) : await PGlite.create();
      await db.exec(
        "DO $$ BEGIN IF NOT EXISTS (SELECT 1 FROM pg_roles WHERE rolname='prism365_app') THEN CREATE ROLE prism365_app NOLOGIN NOSUPERUSER NOBYPASSRLS; END IF; END $$;",
      );
      for (const migration of await loadCoreMigrations()) await db.exec(migration.sql);
      const database = new PGliteModuleDatabase(db, Boolean(options.dataDir), lock);
      await applyLedgerMigrations(pgliteExecutor(db), options.scopes ?? []);
      await database.verify();
      return database;
    } catch (error) {
      await db?.close().catch(() => {});
      if (lock) await unlink(lock).catch(() => {});
      throw error;
    }
  }

  async tenant<T>(
    tenantId: string,
    operation: (sql: SqlClient) => Promise<T>,
    options: TenantOptions = {},
  ): Promise<T> {
    const tenant = validateTenant(tenantId);
    const access = options.access ?? this.#access;
    return this.#db.transaction(async (tx) => {
      await tx.query("SET LOCAL ROLE prism365_app");
      await tx.query("SELECT set_config('prism365.tenant_id', $1, true)", [tenant]);
      const { client, close } = guardedClient(access, async (text, params) => {
        const result = await tx.query<never>(text, params);
        return { rows: result.rows, rowCount: result.affectedRows ?? result.rows.length };
      });
      try {
        return await bounded(operation(client), options);
      } finally {
        close();
      }
    });
  }

  async verify(): Promise<void> {
    const tables = (
      await this.#db.query<TableCheck>(verifyTablesSql, [policyNames, canonicalPolicy])
    ).rows;
    const functions = (await this.#db.query<{ name: string }>(verifyFunctionsSql)).rows;
    assertSchemaIsolated(tables, functions);
    this.#access = coreAccess(tables);
  }

  // Test-only escape hatch for building fixtures that the runtime role cannot create.
  async ownerExecForTests(sql: string): Promise<void> {
    await this.#db.exec(sql);
  }

  async close(): Promise<void> {
    await this.#db.close();
    if (this.#lock) await unlink(this.#lock).catch(() => {});
  }
}

export class PostgresModuleDatabase implements ModuleDatabase {
  readonly kind = "postgres" as const;
  readonly durable = true;
  readonly #pool: Pool;
  #access: TableAccess = { read: new Set(), write: new Set() };

  constructor(options: { connectionString: string }) {
    this.#pool = new Pool({
      connectionString: options.connectionString,
      max: 5,
      connectionTimeoutMillis: 5_000,
      statement_timeout: 30_000,
    });
    // Idle-client failures are handled by pg; do not re-emit credential-bearing driver objects.
    this.#pool.on("error", () => {});
  }

  async tenant<T>(
    tenantId: string,
    operation: (sql: SqlClient) => Promise<T>,
    options: TenantOptions = {},
  ): Promise<T> {
    const tenant = validateTenant(tenantId);
    const access = options.access ?? this.#access;
    const connection: PoolClient = await this.#pool.connect();
    const { client, close } = guardedClient(access, async (text, params) => {
      const value = await connection.query(text, params);
      return { rows: value.rows as never[], rowCount: value.rowCount ?? value.rows.length };
    });
    try {
      await connection.query("BEGIN");
      await connection.query("SELECT set_config('prism365.tenant_id', $1, true)", [tenant]);
      const result = await bounded(operation(client), options);
      close();
      await connection.query("COMMIT");
      return result;
    } catch (error) {
      close();
      await connection.query("ROLLBACK").catch(() => {});
      throw error;
    } finally {
      close();
      connection.release();
    }
  }

  // Refuses to run with a role that could see across tenants or alter the schema.
  async verify(): Promise<void> {
    const role = await this.#pool.query<{ rolsuper: boolean; rolbypassrls: boolean }>(
      "SELECT rolsuper, rolbypassrls FROM pg_roles WHERE rolname = current_user",
    );
    if (!role.rows[0] || role.rows[0].rolsuper || role.rows[0].rolbypassrls)
      throw new Error("The module database role must not be superuser or bypass row security.");
    const owned = await this.#pool.query(
      "SELECT 1 FROM pg_tables WHERE schemaname='public' AND (tablename LIKE 'prism365\\_%' OR tablename LIKE 'mod\\_%') AND pg_has_role(current_user, tableowner, 'MEMBER')",
    );
    if (owned.rowCount) throw new Error("The module database role must not own Prism365 tables.");
    const ledger = await this.#pool.query<{ readable: boolean }>(
      "SELECT has_table_privilege(current_user, 'public.prism365_schema_migration', 'SELECT') AS readable",
    );
    if (ledger.rows[0]?.readable)
      throw new Error("The runtime role must not access the migration ledger.");
    const tables = (
      await this.#pool.query<TableCheck>(verifyTablesSql, [policyNames, canonicalPolicy])
    ).rows;
    const functions = (await this.#pool.query<{ name: string }>(verifyFunctionsSql)).rows;
    assertSchemaIsolated(tables, functions);
    this.#access = coreAccess(tables);
  }

  async close(): Promise<void> {
    await this.#pool.end();
  }
}

// Administrative migration for PostgreSQL: creates the restricted runtime role when missing,
// applies the idempotent core migrations, then each module scope through the ledger.
export async function runPostgresMigrations(options: {
  adminUrl: string;
  appPassword: string;
  scopes: readonly MigrationScope[];
}): Promise<void> {
  if (!/^[A-Za-z0-9_-]{24,}$/.test(options.appPassword))
    throw new MigrationError("DATABASE_APP_PASSWORD must be URL-safe and at least 24 characters.");
  const pool = new Pool({ connectionString: options.adminUrl, connectionTimeoutMillis: 5_000 });
  pool.on("error", () => {});
  const client = await pool.connect();
  const executor = (target: PoolClient): OwnerExecutor => ({
    async exec(sql) {
      await target.query(sql);
    },
    async query<T>(sql: string, params: unknown[]) {
      return (await target.query(sql, params)).rows as T[];
    },
    async transaction(operation) {
      await target.query("BEGIN");
      try {
        await operation(executor(target));
        await target.query("COMMIT");
      } catch (error) {
        await target.query("ROLLBACK").catch(() => {});
        throw error;
      }
    },
  });
  try {
    await client.query("BEGIN");
    const role = await client.query("SELECT 1 FROM pg_roles WHERE rolname='prism365_app'");
    if (role.rowCount === 0) {
      // Utility statements do not bind parameters. The password is validated above and quoted
      // here as well, so loosening the validation later cannot open an injection.
      const quoted = `'${options.appPassword.replaceAll("'", "''")}'`;
      await client.query(
        `CREATE ROLE prism365_app LOGIN NOSUPERUSER NOBYPASSRLS PASSWORD ${quoted}`,
      );
    }
    for (const migration of await loadCoreMigrations()) await client.query(migration.sql);
    await client.query("COMMIT");
    await applyLedgerMigrations(executor(client), options.scopes);
  } catch (error) {
    await client.query("ROLLBACK").catch(() => {});
    // Driver errors can carry connection details; only our own messages are passed through.
    if (error instanceof MigrationError) throw error;
    throw new MigrationError(
      "Database migration failed. Check database access and migration ownership.",
    );
  } finally {
    client.release();
    await pool.end();
  }
}

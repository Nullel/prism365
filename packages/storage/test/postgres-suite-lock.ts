import { Pool } from "pg";

// The PostgreSQL suites of different packages (this one, apps/api's handoff suite) create, use
// and drop the same cluster role, prism365_app, and vitest runs their files side by side. Each
// suite holds this advisory lock for as long as it runs, so they take turns instead of racing
// to create the role or dropping it while another still holds grants to it.
const SUITE_LOCK = 365_006;

// Throws for anything but a disposable local prism365_ci database, as the suites themselves do.
export function disposableAdminUrl(adminUrl: string): URL {
  const parsed = new URL(adminUrl);
  if (
    !["postgres:", "postgresql:"].includes(parsed.protocol) ||
    !["localhost", "127.0.0.1"].includes(parsed.hostname) ||
    parsed.pathname !== "/prism365_ci"
  )
    throw new Error(
      "TEST_DATABASE_ADMIN_URL must point to a disposable local prism365_ci database.",
    );
  return parsed;
}

// A connection pool for fixtures that need the database's owner (past row security) or a lock
// held open, for suites outside this package that have no pg dependency of their own.
export function ownerPool(connectionString: string): Pool {
  const pool = new Pool({ connectionString, connectionTimeoutMillis: 5_000 });
  pool.on("error", () => {});
  return pool;
}

// Waits for the lock and returns the function that releases it.
export async function holdPostgresSuiteLock(adminUrl: string): Promise<() => Promise<void>> {
  const pool = new Pool({
    connectionString: disposableAdminUrl(adminUrl).toString(),
    max: 1,
    connectionTimeoutMillis: 5_000,
  });
  pool.on("error", () => {});
  const client = await pool.connect();
  await client.query("SELECT pg_advisory_lock($1)", [SUITE_LOCK]);
  return async () => {
    try {
      await client.query("SELECT pg_advisory_unlock($1)", [SUITE_LOCK]);
    } finally {
      client.release();
      await pool.end();
    }
  };
}

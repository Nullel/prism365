import { readFile } from "node:fs/promises";
import { Pool } from "pg";

const connectionString = process.env.DATABASE_ADMIN_URL;
const appPassword = process.env.DATABASE_APP_PASSWORD;
if (!connectionString || !appPassword || !/^[A-Za-z0-9_-]{24,}$/.test(appPassword)) {
  throw new Error(
    "Migration requires DATABASE_ADMIN_URL and a URL-safe DATABASE_APP_PASSWORD of 24+ characters.",
  );
}
const pool = new Pool({ connectionString, connectionTimeoutMillis: 5_000 });
const sql = await readFile(new URL("../migrations/001-state.sql", import.meta.url), "utf8");
const client = await pool.connect();
try {
  await client.query("BEGIN");
  const role = await client.query("SELECT 1 FROM pg_roles WHERE rolname='prism365_app'");
  if (role.rowCount === 0) {
    // PostgreSQL utility statements do not bind parameters; quote this operator-owned secret.
    const quotedPassword = `'${appPassword.replaceAll("'", "''")}'`;
    await client.query(
      `CREATE ROLE prism365_app LOGIN NOSUPERUSER NOBYPASSRLS PASSWORD ${quotedPassword}`,
    );
  }
  await client.query(sql);
  await client.query("COMMIT");
  console.log("Prism365 state schema is ready. Use the restricted prism365_app role at runtime.");
} catch {
  await client.query("ROLLBACK");
  throw new Error("Database migration failed. Check database access and migration ownership.");
} finally {
  client.release();
  await pool.end();
}

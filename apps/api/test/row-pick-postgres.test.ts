import { randomBytes } from "node:crypto";
import { PostgresModuleDatabase, runPostgresMigrations } from "@prism365/storage";
import { beforeAll, afterAll } from "vitest";
import {
  disposableAdminUrl,
  holdPostgresSuiteLock,
  ownerPool,
} from "../../../packages/storage/test/postgres-suite-lock.ts";
import { rowPickSuite } from "../../../packages/audits/test/row-pick-suite.ts";

const adminUrl = process.env.TEST_DATABASE_ADMIN_URL;
// Held for the whole file: see postgres-suite-lock.ts.
let releaseSuiteLock: (() => Promise<void>) | undefined;
beforeAll(async () => {
  if (adminUrl) releaseSuiteLock = await holdPostgresSuiteLock(adminUrl);
}, 180_000);
afterAll(async () => {
  await releaseSuiteLock?.();
});
// This public password belongs only to the disposable test database; it is not a secret.
const syntheticPassword = "Prism365_Synthetic_Database_Test_Only_12345";
function identifier(value: string) {
  return `"${value.replaceAll('"', '""')}"`;
}

// Picking one person's rows out of a stored audit run in SQL, on real PostgreSQL: the same suite
// the PGlite run uses (packages/audits/test/row-pick.test.ts). Connects as a throwaway member of
// the app role, like the isolation suite, and runs last so it never races the suites above.
rowPickSuite(
  "PostgreSQL",
  async () => {
    const parsed = disposableAdminUrl(adminUrl ?? "");
    const suffix = randomBytes(8).toString("hex");
    const pickDatabase = `prism365_pick_${suffix}`;
    const memberRole = `prism365_pick_member_${suffix}`;
    const bootstrap = ownerPool(parsed.toString());
    const existing = await bootstrap.query("SELECT 1 FROM pg_roles WHERE rolname='prism365_app'");
    const createdAppRole = existing.rowCount === 0;
    await bootstrap.query(`CREATE DATABASE ${identifier(pickDatabase)}`);
    parsed.pathname = `/${pickDatabase}`;
    await runPostgresMigrations({
      adminUrl: parsed.toString(),
      appPassword: syntheticPassword,
      scopes: [],
    });
    await bootstrap.query(
      `CREATE ROLE ${identifier(memberRole)} LOGIN NOSUPERUSER NOBYPASSRLS PASSWORD '${syntheticPassword}' IN ROLE prism365_app`,
    );
    parsed.username = memberRole;
    parsed.password = syntheticPassword;
    const database = new PostgresModuleDatabase({ connectionString: parsed.toString() });
    await database.verify();
    const close = database.close.bind(database);
    database.close = async () => {
      await close();
      try {
        await bootstrap.query(`DROP DATABASE ${identifier(pickDatabase)} WITH (FORCE)`);
        await bootstrap.query(`DROP ROLE ${identifier(memberRole)}`);
        if (createdAppRole) await bootstrap.query("DROP ROLE prism365_app");
      } finally {
        await bootstrap.end();
      }
    };
    return database;
  },
  Boolean(adminUrl),
);

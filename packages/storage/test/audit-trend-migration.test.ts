import { PGlite } from "@electric-sql/pglite";
import { loadCoreMigrations } from "../src/index.ts";
import { migrationsBeforeTrends, trendMigrationSuite } from "./audit-trend-migration-suite.ts";

// The audit trend migration on PGlite, starting from a database migrated up to 004. The same
// suite runs on PostgreSQL in postgres.test.ts when TEST_DATABASE_ADMIN_URL is set.
trendMigrationSuite("PGlite", async () => {
  const db = await PGlite.create();
  await db.exec(
    "DO $$ BEGIN IF NOT EXISTS (SELECT 1 FROM pg_roles WHERE rolname='prism365_app') THEN CREATE ROLE prism365_app NOLOGIN NOSUPERUSER NOBYPASSRLS; END IF; END $$;",
  );
  for (const migration of await migrationsBeforeTrends()) await db.exec(migration.sql);
  return {
    owner: async (sql, params) => {
      // exec runs a migration file (several statements); query binds parameters.
      if (params === undefined) {
        const results = await db.exec(sql);
        return (results.at(-1)?.rows ?? []) as Record<string, unknown>[];
      }
      return (await db.query<Record<string, unknown>>(sql, params)).rows;
    },
    // PGliteModuleDatabase.open runs every core migration again on each start.
    migrateAll: async () => {
      for (const migration of await loadCoreMigrations()) await db.exec(migration.sql);
    },
    close: () => db.close(),
  };
});

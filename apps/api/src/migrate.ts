import { runPostgresMigrations } from "@prism365/storage";
import { createModuleRegistry } from "./modules.ts";

// Applies the core schema and every registered module's migrations to PostgreSQL with an
// administrator connection. The runtime never receives this URL. PGlite installs migrate
// themselves at startup and do not need this command.
const adminUrl = process.env.DATABASE_ADMIN_URL;
const appPassword = process.env.DATABASE_APP_PASSWORD;
if (!adminUrl || !appPassword) {
  throw new Error("Migration requires DATABASE_ADMIN_URL and DATABASE_APP_PASSWORD.");
}
await runPostgresMigrations({
  adminUrl,
  appPassword,
  scopes: createModuleRegistry().migrationScopes(),
});
console.log("Prism365 schema is ready. Use the restricted prism365_app role at runtime.");

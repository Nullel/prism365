import { resolve } from "node:path";
import { MemoryStateStore, PostgresStateStore } from "@prism365/storage";
import { createApp } from "./app.ts";
import { loadConfig } from "./config.ts";
import { type ModuleRuntime, openModuleRuntime } from "./modules.ts";

// Last resort only: collector-facing promises are already marked handled. A stray rejection is
// logged without its message (which could carry tenant or provider details) instead of ending
// the process and every tenant's collection with it.
process.on("unhandledRejection", (reason) => {
  const name = reason instanceof Error ? reason.name : typeof reason;
  console.error(`Prism365 ignored an unhandled promise rejection (${name}).`);
});

const host = process.env.HOST ?? "127.0.0.1";
const port = Number(process.env.PORT ?? "8787");

if (!Number.isInteger(port) || port < 1 || port > 65_535) {
  throw new Error("PORT must be an integer between 1 and 65535.");
}

const config = await loadConfig();
const store =
  config.databaseUrl && config.encryptionKey
    ? new PostgresStateStore({
        connectionString: config.databaseUrl,
        encryptionKey: config.encryptionKey,
      })
    : new MemoryStateStore();
if (store instanceof PostgresStateStore) await store.verifyRole();
// Module collection exists only in live mode with a configured tenant connection.
let modules: ModuleRuntime | undefined;
if (config.mode === "live" && config.identity) {
  modules = await openModuleRuntime(config);
  modules.scheduler.start();
}
const app = createApp({
  config,
  store,
  ...(modules ? { modules } : {}),
  ...(process.env.WEB_DIST_PATH ? { webRoot: resolve(process.env.WEB_DIST_PATH) } : {}),
});

try {
  await app.listen({ host, port });
  console.log(`Prism365 ${config.mode} API listening on http://${host}:${port}`);
  if (config.mode === "live" && !store.durable)
    console.log(
      "Development sessions are held in memory and reset on restart. Configure PostgreSQL before deployment.",
    );
  if (modules)
    console.log(
      modules.database.durable
        ? `Module data: ${modules.database.kind === "postgres" ? "PostgreSQL" : "PGlite on disk (PRISM_DATA_DIR)"}.`
        : "Module data is held in memory and is lost on restart. Set PRISM_DATA_DIR to keep it.",
    );
} catch (error) {
  app.log.error(error);
  process.exitCode = 1;
}

for (const signal of ["SIGTERM", "SIGINT"] as const)
  process.once(signal, () => {
    void app.close().then(() => {
      process.exitCode = 0;
    });
  });

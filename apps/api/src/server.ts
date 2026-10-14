import { createApp } from "./app.ts";
import { loadConfig } from "./config.ts";
import { MemoryStateStore, PostgresStateStore } from "@prism365/storage";
import { resolve } from "node:path";

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
const app = createApp({
  config,
  store,
  ...(process.env.WEB_DIST_PATH ? { webRoot: resolve(process.env.WEB_DIST_PATH) } : {}),
});

try {
  await app.listen({ host, port });
  console.log(`Prism365 ${config.mode} API listening on http://${host}:${port}`);
  if (config.mode === "live" && !store.durable)
    console.log(
      "Development sessions are held in memory and reset on restart. Configure PostgreSQL before deployment.",
    );
} catch (error) {
  const code = error instanceof Error && "code" in error ? error.code : null;
  console.error(
    code === "EADDRINUSE"
      ? "Prism365 could not start: PORT is already in use. Choose a free port."
      : "Prism365 could not start. Check the configured host, port, and deployment settings.",
  );
  process.exitCode = 1;
}

for (const signal of ["SIGTERM", "SIGINT"] as const)
  process.once(signal, () => {
    void app.close().then(() => {
      process.exitCode = 0;
    });
  });

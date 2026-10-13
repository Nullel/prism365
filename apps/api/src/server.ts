import { createApp } from "./app.ts";

const host = process.env.HOST ?? "127.0.0.1";
const port = Number(process.env.PORT ?? "8787");

if (!Number.isInteger(port) || port < 1 || port > 65_535) {
  throw new Error("PORT must be an integer between 1 and 65535.");
}

if (process.env.DATA_MODE && process.env.DATA_MODE !== "demo") {
  throw new Error(
    "Only DATA_MODE=demo is implemented. Production tenant connections are disabled.",
  );
}

if (process.env.AI_ANALYSIS_ENABLED && process.env.AI_ANALYSIS_ENABLED !== "false") {
  throw new Error(
    "AI analysis is scaffolded but not connected to a real provider or evidence reader.",
  );
}

const app = createApp();

try {
  await app.listen({ host, port });
  console.log(`TenantLens demo API listening on http://${host}:${port}`);
} catch (error) {
  app.log.error(error);
  process.exitCode = 1;
}

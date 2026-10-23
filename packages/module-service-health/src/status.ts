import type { ServiceHealthStatus } from "@prism365/contracts";
import { listAttentionWindow, type ModuleRunRecord } from "@prism365/modules";
import type { SqlClient } from "@prism365/storage";
import { HEALTH_MODULE, INCIDENT_KIND, MESSAGES_MODULE } from "./collector.ts";

// The state of the "Is Microsoft down?" panel. A successful check that found nothing is "ok"
// with no incidents; it is never reported as "healthy", and a check that did not happen, failed
// or is old says so instead of showing an empty list.

type CheckState = ServiceHealthStatus["state"];

export interface CheckInput {
  enabled: boolean;
  granted: boolean;
  collecting: boolean;
  run: ModuleRunRecord | null;
  lastSuccess: string | null;
  // A successful check older than this is stale.
  staleAfterMs: number;
}

export function checkState(
  input: CheckInput,
  now: Date,
): { state: CheckState; detail: string | null } {
  if (!input.enabled)
    return { state: "not_enabled", detail: "This installation has not enabled this source." };
  if (!input.granted)
    return {
      state: "missing_permission",
      detail: "Your Microsoft session has not granted this permission yet.",
    };
  const lastFailed =
    input.run?.status === "failed" &&
    (!input.lastSuccess || Date.parse(input.run.startedAt) > Date.parse(input.lastSuccess));
  if (lastFailed)
    return {
      state: "failed",
      detail: `The last check failed (${input.run?.errorCategory ?? "unknown error"}).`,
    };
  if (!input.lastSuccess)
    return {
      state: "never",
      detail: input.collecting
        ? "The first check has not finished yet."
        : "Not checked yet. Start collecting to check.",
    };
  const latestPartial =
    input.run?.status === "partial" &&
    Date.parse(input.run.startedAt) >= Date.parse(input.lastSuccess) - 60_000;
  if (now.getTime() - Date.parse(input.lastSuccess) > input.staleAfterMs)
    return {
      state: "stale",
      detail: input.collecting
        ? "The last successful check is older than expected."
        : "Collection is stopped, so this is the last thing Microsoft reported.",
    };
  if (latestPartial)
    return {
      state: "partial",
      detail:
        "The last check was incomplete: some of Microsoft's records could not be read, so an incident may be missing.",
    };
  return { state: "ok", detail: null };
}

async function lastSuccess(sql: SqlClient, tenantId: string, module: string, collector: string) {
  const result = await sql.query<{ value: string }>(
    "SELECT value FROM prism365_cursor WHERE tenant_id=$1 AND module=$2 AND collector=$3",
    [tenantId, module, collector],
  );
  try {
    const parsed = JSON.parse(result.rows[0]?.value ?? "null") as { lastSuccess?: unknown } | null;
    return typeof parsed?.lastSuccess === "string" ? parsed.lastSuccess : null;
  } catch {
    return null;
  }
}

export async function serviceHealthStatus(
  sql: SqlClient,
  tenantId: string,
  options: {
    now: Date;
    collecting: boolean;
    runs: readonly ModuleRunRecord[];
    health: { enabled: boolean; granted: boolean };
    messages: { enabled: boolean; granted: boolean };
  },
): Promise<ServiceHealthStatus> {
  const run = (module: string) => options.runs.find((item) => item.module === module) ?? null;
  const healthSuccess = await lastSuccess(sql, tenantId, HEALTH_MODULE, "health");
  const messageSuccess = await lastSuccess(sql, tenantId, MESSAGES_MODULE, "messages");
  const health = checkState(
    {
      ...options.health,
      collecting: options.collecting,
      run: run(HEALTH_MODULE),
      lastSuccess: healthSuccess,
      staleAfterMs: 15 * 60_000,
    },
    options.now,
  );
  const messages = checkState(
    {
      ...options.messages,
      collecting: options.collecting,
      run: run(MESSAGES_MODULE),
      lastSuccess: messageSuccess,
      staleAfterMs: 90 * 60_000,
    },
    options.now,
  );
  const services = await sql.query<{ name: string; status: string }>(
    `SELECT name, status FROM mod_service_health_service WHERE tenant_id=$1 ORDER BY name`,
    [tenantId],
  );
  // Active first, filtered in the query itself, so read or reviewed items can never push an open
  // incident out of the panel.
  const incidents = await listAttentionWindow(sql, tenantId, {
    module: HEALTH_MODULE,
    kind: INCIDENT_KIND,
    since: new Date(options.now.getTime() - 86_400_000),
    limit: 300,
  });
  const attempt = run(HEALTH_MODULE);
  return {
    state: health.state,
    detail: health.detail,
    lastSuccess: healthSuccess,
    lastAttempt: attempt ? (attempt.finishedAt ?? attempt.startedAt) : null,
    services: services.rows.map((row) => ({ name: row.name, status: row.status })),
    incidents,
    messages: { state: messages.state, lastSuccess: messageSuccess },
  };
}

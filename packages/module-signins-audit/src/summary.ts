import { moduleTablePrefix } from "@prism365/modules";
import type { SqlClient } from "@prism365/storage";
import { parseCursor } from "./cursor.ts";
import {
  conditionalAccessBlockCode,
  credentialFailureCodes,
  legacyClients,
  maskIp,
} from "./source.ts";

// Tenant-level sign-in picture for the module's view. Aggregates only: no names, and addresses
// are reduced to their network before they leave this function.

export interface SigninSummary {
  days: number;
  totals: {
    successes: number;
    failures: number;
    blockedByPolicy: number;
    legacySuccesses: number;
    accounts: number;
  };
  countries: { country: string; successes: number; failures: number }[];
  failingNetworks: { network: string; failures: number; accounts: number }[];
  coverage: {
    events: number;
    firstEvent: string | null;
    lastEvent: string | null;
    lastCollected: string | null;
  };
  // How far collection has read Microsoft's log. While history loads (backfill), everything
  // from `from` to `through` is in; newer days are still to come.
  history: {
    phase: "none" | "backfill" | "live";
    from: string | null;
    through: string | null;
  };
}

const table = `"${moduleTablePrefix("signins.audit")}event"`;
const failureCodes = [...credentialFailureCodes].join(",");
const legacyList = [...legacyClients]
  .map((client) => `'${client.replaceAll("'", "''")}'`)
  .join(",");

const isoOrNull = (value: Date | string | null) => (value ? new Date(value).toISOString() : null);

export async function summarizeSignins(
  sql: SqlClient,
  tenantId: string,
  now: Date,
  days: number,
): Promise<SigninSummary> {
  const since = new Date(now.getTime() - days * 86_400_000).toISOString();
  const totals = await sql.query<{
    successes: number;
    failures: number;
    blocked: number;
    legacy: number;
    accounts: number;
  }>(
    `SELECT count(*) FILTER (WHERE error_code = 0)::int AS successes,
            count(*) FILTER (WHERE error_code IN (${failureCodes}))::int AS failures,
            count(*) FILTER (WHERE error_code = ${conditionalAccessBlockCode})::int AS blocked,
            count(*) FILTER (WHERE error_code = 0 AND lower(client_app) IN (${legacyList}))::int AS legacy,
            count(DISTINCT user_id) FILTER (WHERE error_code = 0)::int AS accounts
       FROM ${table} WHERE tenant_id = $1 AND created_at >= $2`,
    [tenantId, since],
  );
  const countries = await sql.query<{ country: string; successes: number; failures: number }>(
    `SELECT COALESCE(country, '??') AS country,
            count(*) FILTER (WHERE error_code = 0)::int AS successes,
            count(*) FILTER (WHERE error_code IN (${failureCodes}))::int AS failures
       FROM ${table} WHERE tenant_id = $1 AND created_at >= $2
      GROUP BY 1 ORDER BY count(*) DESC, 1 LIMIT 15`,
    [tenantId, since],
  );
  const failing = await sql.query<{ ip: string; failures: number; accounts: number }>(
    `SELECT ip, count(*)::int AS failures, count(DISTINCT user_id)::int AS accounts
       FROM ${table}
      WHERE tenant_id = $1 AND created_at >= $2 AND ip IS NOT NULL AND error_code IN (${failureCodes})
      GROUP BY ip ORDER BY 2 DESC LIMIT 500`,
    [tenantId, since],
  );
  const networks = new Map<string, { failures: number; accounts: number }>();
  for (const row of failing.rows) {
    const network = maskIp(row.ip);
    if (!network) continue;
    const current = networks.get(network) ?? { failures: 0, accounts: 0 };
    networks.set(network, {
      failures: current.failures + row.failures,
      accounts: Math.max(current.accounts, row.accounts),
    });
  }
  const coverage = await sql.query<{
    events: number;
    first_event: Date | string | null;
    last_event: Date | string | null;
    last_collected: Date | string | null;
  }>(
    `SELECT count(*)::int AS events, min(created_at) AS first_event, max(created_at) AS last_event,
            max(collected_at) AS last_collected
       FROM ${table} WHERE tenant_id = $1`,
    [tenantId],
  );
  const position = await sql.query<{ value: string }>(
    "SELECT value FROM prism365_cursor WHERE tenant_id = $1 AND module = $2 AND collector = $3",
    [tenantId, "signins.audit", "signins"],
  );
  const cursor = parseCursor(position.rows[0]?.value);
  const total = totals.rows[0];
  const cover = coverage.rows[0];
  return {
    days,
    totals: {
      successes: total?.successes ?? 0,
      failures: total?.failures ?? 0,
      blockedByPolicy: total?.blocked ?? 0,
      legacySuccesses: total?.legacy ?? 0,
      accounts: total?.accounts ?? 0,
    },
    countries: countries.rows,
    failingNetworks: [...networks]
      .map(([network, value]) => ({ network, ...value }))
      .sort((a, b) => b.failures - a.failures || a.network.localeCompare(b.network))
      .slice(0, 10),
    coverage: {
      events: cover?.events ?? 0,
      firstEvent: isoOrNull(cover?.first_event ?? null),
      lastEvent: isoOrNull(cover?.last_event ?? null),
      lastCollected: isoOrNull(cover?.last_collected ?? null),
    },
    history: !cursor
      ? { phase: "none", from: null, through: null }
      : cursor.phase === "backfill"
        ? { phase: "backfill", from: cursor.start, through: cursor.next }
        : { phase: "live", from: null, through: cursor.through },
  };
}

import { type CanaryList, canaryAccountSchema, MAX_CANARY_ACCOUNTS } from "@prism365/contracts";
import { moduleTablePrefix } from "@prism365/modules";
import type { SqlClient } from "@prism365/storage";

// The canary list: one setting per tenant in prism365_preference, written through the API and read
// by the collector once per run. The module declares the key, so the collector can read it
// without any access to the rest of the preference table.

export const CANARY_SETTING = "signins.canaries";

const events = `"${moduleTablePrefix("signins.audit")}event"`;

// The list as stored. A missing, damaged or oddly shaped value reads as no canaries, and a
// damaged entry is left out without taking the others with it: one bad row must not switch off
// every decoy.
export function parseCanaryList(raw: unknown): CanaryList {
  let value = raw;
  if (typeof value === "string") {
    try {
      value = JSON.parse(value) as unknown;
    } catch {
      return { accounts: [] };
    }
  }
  const stored = (value as { accounts?: unknown } | null)?.accounts;
  if (!Array.isArray(stored)) return { accounts: [] };
  const seen = new Set<string>();
  const accounts: CanaryList["accounts"] = [];
  for (const item of stored) {
    const parsed = canaryAccountSchema.safeParse(item);
    if (!parsed.success || seen.has(parsed.data.userId)) continue;
    seen.add(parsed.data.userId);
    accounts.push(parsed.data);
    if (accounts.length === MAX_CANARY_ACCOUNTS) break;
  }
  return { accounts };
}

export async function readCanaries(sql: SqlClient, tenantId: string): Promise<CanaryList> {
  const result = await sql.query<{ value: unknown }>(
    "SELECT value FROM prism365_preference WHERE tenant_id=$1 AND key=$2",
    [tenantId, CANARY_SETTING],
  );
  return parseCanaryList(result.rows[0]?.value);
}

export async function writeCanaries(
  sql: SqlClient,
  tenantId: string,
  list: CanaryList,
  actorId: string,
): Promise<CanaryList> {
  const value = parseCanaryList(list);
  await sql.query(
    `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
     VALUES ($1,$2,$3::jsonb,now(),$4)
     ON CONFLICT (tenant_id, key) DO UPDATE SET value=EXCLUDED.value, updated_at=EXCLUDED.updated_at,
       updated_by=EXCLUDED.updated_by`,
    [tenantId, CANARY_SETTING, JSON.stringify(value), actorId],
  );
  return value;
}

// The newest sign-in name stored for each of these accounts, from the sign-ins already collected.
// An account with no stored sign-in is simply absent.
export async function storedPrincipals(
  sql: SqlClient,
  tenantId: string,
  userIds: readonly string[],
): Promise<Map<string, string>> {
  const found = new Map<string, string>();
  const ids = [...new Set(userIds)].slice(0, MAX_CANARY_ACCOUNTS);
  if (ids.length === 0) return found;
  const result = await sql.query<{ user_id: string; upn: string }>(
    `SELECT user_id, (array_agg(upn ORDER BY created_at DESC))[1] AS upn
       FROM ${events}
      WHERE tenant_id = $1 AND upn IS NOT NULL AND user_id IN (${ids.map((_id, index) => `$${index + 2}`).join(",")})
      GROUP BY user_id`,
    [tenantId, ...ids],
  );
  for (const row of result.rows) found.set(row.user_id, row.upn);
  return found;
}

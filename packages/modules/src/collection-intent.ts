import type { SqlClient } from "@prism365/storage";
import { z } from "zod";

// Whether background collection should be running, and on whose session. Set when an
// administrator starts collection and cleared only when someone stops it, so a restart, a
// sign-out or an expired session pauses collection instead of ending it: it resumes on that
// administrator's next signed-in request. Only who chose to collect is stored, never a session
// or token.

const KEY = "collection.intent";

const intentSchema = z.object({
  actorId: z.uuid(),
  since: z.string().datetime({ offset: true }),
});
export type CollectionIntent = z.infer<typeof intentSchema>;

export async function readCollectionIntent(
  sql: SqlClient,
  tenantId: string,
): Promise<CollectionIntent | null> {
  const result = await sql.query<{ value: unknown }>(
    "SELECT value FROM prism365_preference WHERE tenant_id=$1 AND key=$2",
    [tenantId, KEY],
  );
  const raw = result.rows[0]?.value;
  if (raw === undefined) return null;
  const parsed = intentSchema.safeParse(typeof raw === "string" ? JSON.parse(raw) : raw);
  return parsed.success ? parsed.data : null;
}

export async function writeCollectionIntent(
  sql: SqlClient,
  tenantId: string,
  intent: CollectionIntent,
): Promise<void> {
  const value = intentSchema.parse(intent);
  await sql.query(
    `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
     VALUES ($1,$2,$3::jsonb,now(),$4)
     ON CONFLICT (tenant_id, key) DO UPDATE SET value=EXCLUDED.value, updated_at=EXCLUDED.updated_at,
       updated_by=EXCLUDED.updated_by`,
    [tenantId, KEY, JSON.stringify(value), value.actorId],
  );
}

export async function clearCollectionIntent(sql: SqlClient, tenantId: string): Promise<void> {
  await sql.query("DELETE FROM prism365_preference WHERE tenant_id=$1 AND key=$2", [tenantId, KEY]);
}

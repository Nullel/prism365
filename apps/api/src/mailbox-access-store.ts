import {
  cleanText,
  type MailboxAccessExport,
  type MailboxAccessImport,
  type MailboxHolder,
  type MailboxRight,
  type MailboxType,
  type PersonMailboxGrant,
  type TrusteeKind,
} from "@prism365/contracts";
import type { SqlClient } from "@prism365/storage";

// The latest mailbox access import in prism365_mailbox_access_import and
// prism365_mailbox_access_grant (migration 008-mailbox-access.sql). Every query is tenant-scoped
// twice: by the transaction's row-security setting and by the tenant_id predicate. Rows are never
// updated: an import replaces the previous one whole, and a clear deletes it.

// Grants per INSERT. Ten values a row keeps each statement well under the SQL guard's size budget
// (20,000 characters), so the largest export is 1,000 statements in one transaction.
const GRANT_BATCH = 200;

// The unique index that keeps one import per tenant. Two imports that race fail on it, the later
// one keeping nothing.
export const LATEST_IMPORT_INDEX = "prism365_mailbox_access_import_latest";

const iso = (value: Date | string) => new Date(value).toISOString();

interface ImportRow {
  id: string;
  exported_at: Date | string;
  imported_at: Date | string;
  imported_by: string;
  imported_by_name: string | null;
  include_user_mailboxes: boolean;
  mailbox_count: number;
  grant_count: number;
  unresolved_count: number;
}

export async function currentImport(
  sql: SqlClient,
  tenantId: string,
): Promise<MailboxAccessImport | null> {
  const result = await sql.query<ImportRow>(
    `SELECT id, exported_at, imported_at, imported_by, imported_by_name, include_user_mailboxes,
            mailbox_count, grant_count, unresolved_count
       FROM prism365_mailbox_access_import
      WHERE tenant_id=$1`,
    [tenantId],
  );
  const row = result.rows[0];
  if (!row) return null;
  return {
    id: row.id,
    exportedAt: iso(row.exported_at),
    importedAt: iso(row.imported_at),
    importedBy: { id: row.imported_by, name: row.imported_by_name },
    includeUserMailboxes: row.include_user_mailboxes,
    mailboxes: row.mailbox_count,
    grants: row.grant_count,
    unresolved: row.unresolved_count,
  };
}

// Takes the tenant's lock row, creating it the first time. Imports and clears take it first, so
// they never interleave, even before any import exists: the second waits for the first to commit,
// then sees what it committed. Two first-time callers meet on the row's key, and the second waits
// there for the first.
async function lockTenant(sql: SqlClient, tenantId: string): Promise<void> {
  await sql.query(
    "INSERT INTO prism365_mailbox_access_lock (tenant_id) VALUES ($1) ON CONFLICT (tenant_id) DO NOTHING",
    [tenantId],
  );
  await sql.query(
    "SELECT tenant_id FROM prism365_mailbox_access_lock WHERE tenant_id=$1 FOR UPDATE",
    [tenantId],
  );
}

// Deletes the tenant's import and its grants, holding the tenant's lock.
async function deleteCurrent(sql: SqlClient, tenantId: string): Promise<boolean> {
  await lockTenant(sql, tenantId);
  await sql.query("DELETE FROM prism365_mailbox_access_grant WHERE tenant_id=$1", [tenantId]);
  const deleted = await sql.query("DELETE FROM prism365_mailbox_access_import WHERE tenant_id=$1", [
    tenantId,
  ]);
  return deleted.rowCount > 0;
}

// Replaces the tenant's import with this validated export. Call it inside one database.tenant()
// transaction: any grant that fails to store throws, and the transaction then keeps the previous
// import exactly as it was. imported_at is the database's now().
export async function replaceImport(
  sql: SqlClient,
  tenantId: string,
  file: MailboxAccessExport,
  meta: { id: string; importedBy: { id: string; name: string | null } },
): Promise<{ summary: MailboxAccessImport; replaced: boolean }> {
  const replaced = await deleteCurrent(sql, tenantId);
  await sql.query(
    `INSERT INTO prism365_mailbox_access_import (tenant_id, id, imported_at, exported_at,
       imported_by, imported_by_name, include_user_mailboxes, mailbox_count, grant_count,
       unresolved_count)
     VALUES ($1,$2,now(),$3,$4,$5,$6,$7,$8,$9)`,
    [
      tenantId,
      meta.id,
      file.exportedAt,
      meta.importedBy.id,
      // The administrator's name as the sign-in token gave it: bounded, then cleaned.
      meta.importedBy.name === null ? null : cleanText(meta.importedBy.name.slice(0, 500)),
      file.includeUserMailboxes,
      file.mailboxes.length,
      file.grants.length,
      file.unresolved,
    ],
  );
  const mailboxes = new Map(file.mailboxes.map((mailbox) => [mailbox.id, mailbox]));
  for (let start = 0; start < file.grants.length; start += GRANT_BATCH) {
    const batch = file.grants.slice(start, start + GRANT_BATCH);
    const params: unknown[] = [tenantId, meta.id];
    const rows = batch.map((grant, offset) => {
      const mailbox = mailboxes.get(grant.mailboxId);
      // The schema refuses a grant for a mailbox the file does not list.
      if (!mailbox) throw new Error("A grant names a mailbox the export does not list.");
      const base = params.length;
      params.push(
        start + offset,
        mailbox.id,
        mailbox.address,
        mailbox.displayName,
        mailbox.type,
        grant.right,
        grant.trustee.id,
        grant.trustee.address,
        grant.trustee.name,
        grant.trustee.kind,
      );
      const values = Array.from({ length: 10 }, (_, index) => `$${base + index + 1}`);
      return `($1,$2,${values.join(",")})`;
    });
    await sql.query(
      `INSERT INTO prism365_mailbox_access_grant (tenant_id, import_id, position, mailbox_id,
         mailbox_address, mailbox_name, mailbox_type, access_right, trustee_id, trustee_address,
         trustee_name, trustee_kind)
       VALUES ${rows.join(",")}`,
      params,
    );
  }
  const summary = await currentImport(sql, tenantId);
  if (!summary || summary.id !== meta.id) throw new Error("The import was not stored.");
  return { summary, replaced };
}

export async function clearImport(sql: SqlClient, tenantId: string): Promise<boolean> {
  return deleteCurrent(sql, tenantId);
}

interface PersonReadRow extends ImportRow {
  // null on the one row of an import that has nothing for this person.
  list: "grant" | "holder" | null;
  mailbox_id: string;
  mailbox_address: string;
  mailbox_name: string;
  mailbox_type: MailboxType;
  access_right: MailboxRight;
  trustee_id: string | null;
  trustee_address: string | null;
  trustee_name: string;
  trustee_kind: TrusteeKind;
}

export interface PersonMailboxRead {
  stored: MailboxAccessImport;
  grants: PersonMailboxGrant[];
  grantsTruncated: boolean;
  holders: MailboxHolder[];
  holdersTruncated: boolean;
}

// The columns both lists carry. k1 to k5 are each list's sort keys under common names, so the
// outer query can keep both lists in the order their LIMIT chose.
const grantColumns = `g.mailbox_id, g.mailbox_address, g.mailbox_name, g.mailbox_type,
  g.access_right, g.trustee_id, g.trustee_address, g.trustee_name, g.trustee_kind`;

// What one person sees, in ONE statement, so the import and both lists come from one snapshot:
// at READ COMMITTED every statement takes its own, and an import committed between two statements
// would pair one export's date with another's grants (or none).
//
// - grants: the rights the person holds, granted to the account itself or to a group in
//   `groupIds` (the person's groups, nested ones included, as Microsoft listed them). A grant to a
//   group is matched only when the export says the trustee is a group.
// - holders: when the person is a mailbox in the import, who holds which right on it.
//
// null when nothing is imported.
export async function personMailboxRead(
  sql: SqlClient,
  tenantId: string,
  userId: string,
  groupIds: readonly string[],
  limits: { grants: number; holders: number },
): Promise<PersonMailboxRead | null> {
  const result = await sql.query<PersonReadRow>(
    `WITH latest AS (
       SELECT id, exported_at, imported_at, imported_by, imported_by_name, include_user_mailboxes,
              mailbox_count, grant_count, unresolved_count
         FROM prism365_mailbox_access_import
        WHERE tenant_id=$1),
     mine AS (
       SELECT 'grant' AS list, g.mailbox_name AS k1, g.mailbox_address AS k2,
              g.mailbox_id::text AS k3, g.access_right AS k4, g.position AS k5, ${grantColumns}
         FROM prism365_mailbox_access_grant g JOIN latest l ON g.import_id = l.id
        WHERE g.tenant_id=$1
          AND (g.trustee_id = $2::uuid
               OR (g.trustee_kind = 'group' AND g.trustee_id = ANY($3::uuid[])))
        ORDER BY k1, k2, k3, k4, k5
        LIMIT $4),
     held AS (
       SELECT 'holder' AS list, g.access_right AS k1, g.trustee_name AS k2, '' AS k3, '' AS k4,
              g.position AS k5, ${grantColumns}
         FROM prism365_mailbox_access_grant g JOIN latest l ON g.import_id = l.id
        WHERE g.tenant_id=$1 AND g.mailbox_id = $2::uuid
        ORDER BY k1, k2, k5
        LIMIT $5)
     SELECT l.id, l.exported_at, l.imported_at, l.imported_by, l.imported_by_name,
            l.include_user_mailboxes, l.mailbox_count, l.grant_count, l.unresolved_count,
            r.list, r.mailbox_id, r.mailbox_address, r.mailbox_name, r.mailbox_type,
            r.access_right, r.trustee_id, r.trustee_address, r.trustee_name, r.trustee_kind
       FROM latest l
       LEFT JOIN (SELECT * FROM mine UNION ALL SELECT * FROM held) r ON true
      ORDER BY r.list, r.k1, r.k2, r.k3, r.k4, r.k5`,
    [tenantId, userId, [...groupIds], limits.grants + 1, limits.holders + 1],
  );
  const first = result.rows[0];
  if (!first) return null;
  const mine = result.rows.filter((row) => row.list === "grant");
  const held = result.rows.filter((row) => row.list === "holder");
  return {
    stored: {
      id: first.id,
      exportedAt: iso(first.exported_at),
      importedAt: iso(first.imported_at),
      importedBy: { id: first.imported_by, name: first.imported_by_name },
      includeUserMailboxes: first.include_user_mailboxes,
      mailboxes: first.mailbox_count,
      grants: first.grant_count,
      unresolved: first.unresolved_count,
    },
    grants: mine.slice(0, limits.grants).map((row) => ({
      mailbox: {
        id: row.mailbox_id,
        address: row.mailbox_address,
        name: row.mailbox_name,
        type: row.mailbox_type,
      },
      right: row.access_right,
      via:
        row.trustee_id === userId
          ? { kind: "direct" as const }
          : { kind: "group" as const, id: row.trustee_id ?? "", name: row.trustee_name },
    })),
    grantsTruncated: mine.length > limits.grants,
    holders: held.slice(0, limits.holders).map((row) => ({
      right: row.access_right,
      trustee: {
        id: row.trustee_id,
        address: row.trustee_address,
        name: row.trustee_name,
        kind: row.trustee_kind,
      },
    })),
    holdersTruncated: held.length > limits.holders,
  };
}

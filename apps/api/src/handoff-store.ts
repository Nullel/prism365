import {
  type HandoffRecord,
  type PersonReceiptItem,
  principalKey,
  type Receipt,
  type ReceiptStatus,
  type ReceiptSummary,
  type StoredReceiptItem,
} from "@prism365/contracts";
import type { SqlClient } from "@prism365/storage";

// Handoffs and receipts in prism365_handoff, prism365_receipt and prism365_receipt_item
// (migration 006-handoffs.sql). Every query is tenant-scoped twice: by the transaction's
// row-security setting and by the tenant_id predicate. Nothing here updates a row: a receipt
// stays as the tool reported it until its retention period ends (pruneHandoffs).

export const LIST_LIMIT = 100;
// Items per INSERT. Seven values a row keeps each statement well under the SQL guard's size
// budget (20,000 characters), so a 20,000-item receipt is 80 statements in one transaction.
const ITEM_BATCH = 250;

const iso = (value: Date | string) => new Date(value).toISOString();
const json = (value: unknown): unknown =>
  typeof value === "string" ? (JSON.parse(value) as unknown) : value;

function scopeOf(value: unknown): string[] | null {
  const parsed = json(value);
  if (!Array.isArray(parsed)) return null;
  const domains = parsed.filter((item): item is string => typeof item === "string");
  return domains.length > 0 ? domains.slice(0, 100) : null;
}

// Rows a single clean-up deletes at most from each table: a backlog clears over several calls,
// and no one call comes near the transaction's time limit.
export const PRUNE_BATCH = 5_000;

export interface PruneOutcome {
  items: number;
  receipts: number;
  handoffs: number;
  // Why nothing was deleted, when the clean-up chose not to run.
  skipped: "clock" | null;
}

// Deletes, in one short transaction of its own, a batch of what is older than the retention
// period: items of old receipts, the old receipts with no items left, then the old handoffs no
// kept receipt answers (a receipt arrives after its handoff, so a handoff stays as long as any
// receipt that answers it). Age is measured by the database's clock, which also stamped every
// imported_at; if that clock reads earlier than the newest receipt, it was set back, and nothing
// is deleted. Rows another transaction holds (an import answering an old handoff holds it for its
// foreign key) are skipped rather than waited for, so a clean-up never blocks, deadlocks with or
// fails a request; whatever it skips, a later one deletes.
export async function pruneHandoffs(
  sql: SqlClient,
  tenantId: string,
  options: { days: number; batch?: number },
): Promise<PruneOutcome> {
  const batch = options.batch ?? PRUNE_BATCH;
  const clock = await sql.query<{ behind: boolean | null }>(
    "SELECT now() < max(imported_at) AS behind FROM prism365_receipt WHERE tenant_id=$1",
    [tenantId],
  );
  if (clock.rows[0]?.behind) return { items: 0, receipts: 0, handoffs: 0, skipped: "clock" };
  const cutoff = "now() - make_interval(days => $2)";
  const items = await sql.query(
    `DELETE FROM prism365_receipt_item
      WHERE tenant_id=$1 AND (receipt_id, position) IN (
        SELECT receipt_id, position FROM prism365_receipt_item
         WHERE tenant_id=$1
           AND receipt_id IN (SELECT id FROM prism365_receipt
                               WHERE tenant_id=$1 AND imported_at < ${cutoff})
         ORDER BY receipt_id, position
         LIMIT $3
           FOR UPDATE SKIP LOCKED)`,
    [tenantId, options.days, batch],
  );
  const receipts = await sql.query(
    `DELETE FROM prism365_receipt
      WHERE tenant_id=$1 AND id IN (
        SELECT r.id FROM prism365_receipt r
         WHERE r.tenant_id=$1 AND r.imported_at < ${cutoff}
           AND NOT EXISTS (SELECT 1 FROM prism365_receipt_item i
                            WHERE i.tenant_id = r.tenant_id AND i.receipt_id = r.id)
         ORDER BY r.imported_at
         LIMIT $3
           FOR UPDATE SKIP LOCKED)`,
    [tenantId, options.days, batch],
  );
  const handoffs = await sql.query(
    `DELETE FROM prism365_handoff
      WHERE tenant_id=$1 AND id IN (
        SELECT h.id FROM prism365_handoff h
         WHERE h.tenant_id=$1 AND h.generated_at < ${cutoff}
           AND NOT EXISTS (SELECT 1 FROM prism365_receipt r
                            WHERE r.tenant_id = h.tenant_id AND r.handoff_id = h.id)
         ORDER BY h.generated_at
         LIMIT $3
           FOR UPDATE SKIP LOCKED)`,
    [tenantId, options.days, batch],
  );
  return {
    items: items.rowCount,
    receipts: receipts.rowCount,
    handoffs: handoffs.rowCount,
    skipped: null,
  };
}

interface HandoffRow {
  id: string;
  audit_id: string;
  audit_title: string;
  run_id: string;
  view_id: string | null;
  view_label: string | null;
  filtered: boolean;
  run_scope: unknown;
  complete: boolean;
  items: number;
  generated_at: Date | string;
  generated_by: string;
  generated_by_name: string | null;
}

export async function insertHandoff(
  sql: SqlClient,
  tenantId: string,
  handoff: HandoffRecord,
): Promise<void> {
  await sql.query(
    `INSERT INTO prism365_handoff (tenant_id, id, audit_id, audit_title, run_id, view_id,
       view_label, filtered, run_scope, complete, items, generated_at, generated_by,
       generated_by_name)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9::jsonb,$10,$11,$12,$13,$14)`,
    [
      tenantId,
      handoff.id,
      handoff.auditId,
      handoff.auditTitle,
      handoff.runId,
      handoff.view?.id ?? null,
      handoff.view?.label ?? null,
      handoff.filtered,
      handoff.scope ? JSON.stringify(handoff.scope) : null,
      handoff.complete,
      handoff.items,
      handoff.generatedAt,
      handoff.generatedBy.id,
      handoff.generatedBy.name,
    ],
  );
}

export async function listHandoffs(
  sql: SqlClient,
  tenantId: string,
  limit = LIST_LIMIT,
): Promise<HandoffRecord[]> {
  const result = await sql.query<HandoffRow>(
    `SELECT id, audit_id, audit_title, run_id, view_id, view_label, filtered, run_scope,
            complete, items, generated_at, generated_by, generated_by_name
       FROM prism365_handoff
      WHERE tenant_id=$1
      ORDER BY generated_at DESC, id DESC
      LIMIT $2`,
    [tenantId, limit],
  );
  return result.rows.map((row) => ({
    id: row.id,
    auditId: row.audit_id,
    auditTitle: row.audit_title,
    runId: row.run_id,
    view: row.view_id === null ? null : { id: row.view_id, label: row.view_label ?? row.view_id },
    filtered: row.filtered,
    scope: scopeOf(row.run_scope),
    complete: row.complete,
    items: row.items,
    generatedAt: iso(row.generated_at),
    generatedBy: { id: row.generated_by, name: row.generated_by_name },
  }));
}

export async function handoffExists(sql: SqlClient, tenantId: string, id: string) {
  const result = await sql.query("SELECT 1 FROM prism365_handoff WHERE tenant_id=$1 AND id=$2", [
    tenantId,
    id,
  ]);
  return result.rows.length > 0;
}

interface ReceiptRow {
  id: string;
  handoff_id: string | null;
  tool_name: string;
  tool_version: string;
  performed_at: Date | string;
  performed_by: string;
  items: number;
  done_count: number;
  failed_count: number;
  skipped_count: number;
  imported_at: Date | string;
  digest: string;
  digest_version: number;
  handoff_audit_id: string | null;
  handoff_generated_at: Date | string | null;
  handoff_items: number | null;
}

const receiptColumns = `r.id, r.handoff_id, r.tool_name, r.tool_version, r.performed_at,
  r.performed_by, r.items, r.done_count, r.failed_count, r.skipped_count, r.imported_at, r.digest,
  r.digest_version, h.audit_id AS handoff_audit_id, h.generated_at AS handoff_generated_at,
  h.items AS handoff_items`;
// The handoff a receipt answers, for the list. The foreign key guarantees it exists.
const receiptSource = `prism365_receipt r
  LEFT JOIN prism365_handoff h ON h.tenant_id = r.tenant_id AND h.id = r.handoff_id`;

function toSummary(row: ReceiptRow): ReceiptSummary {
  return {
    id: row.id,
    handoff:
      row.handoff_id && row.handoff_audit_id && row.handoff_generated_at !== null
        ? {
            id: row.handoff_id,
            auditId: row.handoff_audit_id,
            generatedAt: iso(row.handoff_generated_at),
            items: row.handoff_items ?? 0,
          }
        : null,
    tool: { name: row.tool_name, version: row.tool_version },
    performedAt: iso(row.performed_at),
    performedBy: row.performed_by,
    importedAt: iso(row.imported_at),
    items: row.items,
    done: row.done_count,
    failed: row.failed_count,
    skipped: row.skipped_count,
  };
}

export interface StoredReceipt {
  summary: ReceiptSummary;
  digest: string;
  // How the digest was taken (see the import route). Only digests of one version compare.
  digestVersion: number;
}

export async function findReceipt(
  sql: SqlClient,
  tenantId: string,
  id: string,
): Promise<StoredReceipt | null> {
  const result = await sql.query<ReceiptRow>(
    `SELECT ${receiptColumns} FROM ${receiptSource} WHERE r.tenant_id=$1 AND r.id=$2`,
    [tenantId, id],
  );
  const row = result.rows[0];
  return row
    ? { summary: toSummary(row), digest: row.digest, digestVersion: row.digest_version }
    : null;
}

export async function listReceipts(
  sql: SqlClient,
  tenantId: string,
  limit = LIST_LIMIT,
): Promise<ReceiptSummary[]> {
  const result = await sql.query<ReceiptRow>(
    `SELECT ${receiptColumns} FROM ${receiptSource}
      WHERE r.tenant_id=$1
      ORDER BY r.imported_at DESC, r.id DESC
      LIMIT $2`,
    [tenantId, limit],
  );
  return result.rows.map(toSummary);
}

export function statusCounts(receipt: Receipt): Record<ReceiptStatus, number> {
  const counts = { done: 0, failed: 0, skipped: 0 };
  for (const item of receipt.items) counts[item.status] += 1;
  return counts;
}

// Stores a validated receipt and every item. Call it inside one database.tenant() transaction:
// an item that fails to store throws, and the transaction then keeps nothing of the receipt.
// imported_at is the database's now(), the clock retention measures by.
// A receipt whose id is already stored (imported by another request a moment earlier) is left
// alone and returned as stored, with its digest, for the caller to compare.
export async function insertReceipt(
  sql: SqlClient,
  tenantId: string,
  receipt: Receipt,
  meta: { digest: string; digestVersion: number; importedBy: string },
): Promise<({ inserted: true } | { inserted: false }) & StoredReceipt> {
  const counts = statusCounts(receipt);
  const inserted = await sql.query(
    `INSERT INTO prism365_receipt (tenant_id, id, handoff_id, tool_name, tool_version,
       performed_at, performed_by, items, done_count, failed_count, skipped_count, imported_at,
       imported_by, digest, digest_version)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,now(),$12,$13,$14)
     ON CONFLICT (tenant_id, id) DO NOTHING
     RETURNING id`,
    [
      tenantId,
      receipt.id,
      receipt.handoffId,
      receipt.tool.name,
      receipt.tool.version,
      receipt.performedAt,
      receipt.performedBy,
      receipt.items.length,
      counts.done,
      counts.failed,
      counts.skipped,
      meta.importedBy,
      meta.digest,
      meta.digestVersion,
    ],
  );
  if (inserted.rows.length === 0) {
    const existing = await findReceipt(sql, tenantId, receipt.id);
    if (!existing) throw new Error("The receipt was neither stored nor found.");
    return { inserted: false, ...existing };
  }
  for (let start = 0; start < receipt.items.length; start += ITEM_BATCH) {
    const batch = receipt.items.slice(start, start + ITEM_BATCH);
    const params: unknown[] = [tenantId, receipt.id];
    const rows = batch.map((item, offset) => {
      const base = params.length;
      params.push(
        start + offset,
        item.userId,
        item.userPrincipalName,
        item.userPrincipalName === null ? null : principalKey(item.userPrincipalName),
        item.action,
        item.status,
        item.detail,
      );
      const values = Array.from({ length: 7 }, (_, index) => `$${base + index + 1}`);
      return `($1,$2,${values.join(",")})`;
    });
    await sql.query(
      `INSERT INTO prism365_receipt_item (tenant_id, receipt_id, position, user_id,
         user_principal_name, user_principal_key, action, status, detail)
       VALUES ${rows.join(",")}`,
      params,
    );
  }
  const stored = await findReceipt(sql, tenantId, receipt.id);
  if (!stored) throw new Error("The receipt was not stored.");
  return { inserted: true, ...stored };
}

interface ItemRow {
  position: number;
  user_id: string | null;
  user_principal_name: string | null;
  action: string;
  status: ReceiptStatus;
  detail: string;
}

const toItem = (row: ItemRow): StoredReceiptItem => ({
  position: row.position,
  userId: row.user_id,
  userPrincipalName: row.user_principal_name,
  action: row.action,
  status: row.status,
  detail: row.detail,
});

// One page of a receipt's items in the receipt's order, after the given position.
export async function receiptItems(
  sql: SqlClient,
  tenantId: string,
  receiptId: string,
  options: { after: number | null; limit: number },
): Promise<{ items: StoredReceiptItem[]; next: number | null }> {
  const result = await sql.query<ItemRow>(
    `SELECT position, user_id, user_principal_name, action, status, detail
       FROM prism365_receipt_item
      WHERE tenant_id=$1 AND receipt_id=$2 AND position > $3
      ORDER BY position
      LIMIT $4`,
    [tenantId, receiptId, options.after ?? -1, options.limit + 1],
  );
  const items = result.rows.slice(0, options.limit).map(toItem);
  return {
    items,
    next: result.rows.length > options.limit ? (items.at(-1)?.position ?? null) : null,
  };
}

interface PersonItemRow extends ItemRow {
  receipt_id: string;
  tool_name: string;
  tool_version: string;
  performed_at: Date | string;
  performed_by: string;
  imported_at: Date | string;
}

// What receipts report about one account, newest first: items that name its id, and items that
// name only a principal name with the same key as `principal` (principalKey: NFKC, lower case).
// An item that names another account's id never matches by principal name, even when the names
// agree.
export async function personReceiptItems(
  sql: SqlClient,
  tenantId: string,
  userId: string,
  principal: string | null,
  limit = LIST_LIMIT,
): Promise<{ items: PersonReceiptItem[]; truncated: boolean }> {
  const result = await sql.query<PersonItemRow>(
    `SELECT i.receipt_id, i.position, i.user_id, i.user_principal_name, i.action, i.status,
            i.detail, r.tool_name, r.tool_version, r.performed_at, r.performed_by, r.imported_at
       FROM prism365_receipt_item i
       JOIN prism365_receipt r ON r.tenant_id = i.tenant_id AND r.id = i.receipt_id
      WHERE i.tenant_id=$1
        AND (i.user_id = $2
             OR ($3::text IS NOT NULL AND i.user_id IS NULL AND i.user_principal_key = $3::text))
      ORDER BY r.performed_at DESC, r.imported_at DESC, i.receipt_id, i.position
      LIMIT $4`,
    [tenantId, userId, principal === null ? null : principalKey(principal), limit + 1],
  );
  return {
    items: result.rows.slice(0, limit).map((row) => ({
      ...toItem(row),
      receiptId: row.receipt_id,
      tool: { name: row.tool_name, version: row.tool_version },
      performedAt: iso(row.performed_at),
      performedBy: row.performed_by,
      importedAt: iso(row.imported_at),
      matchedBy: row.user_id === userId ? "userId" : "userPrincipalName",
    })),
    truncated: result.rows.length > limit,
  };
}

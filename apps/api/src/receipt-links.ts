import type { AuditRow, StoredReceiptItem } from "@prism365/contracts";
import { joinKey } from "@prism365/audits";

// A receipt is the action tool's own report, so the account an item names is its claim. Prism365
// links an item to a person's page only when the id and the principal name it gives both belong
// to one account in the run behind the handoff the receipt answers: an account Prism365 itself
// listed. The handoff keeps no copy of the file that was downloaded, so the stored run is the
// record of what could have been exported (the export is a view of it).

// Each account id in the run, with the principal names its cells show.
export function runAccounts(rows: readonly AuditRow[]): Map<string, Set<string>> {
  const accounts = new Map<string, Set<string>>();
  for (const row of rows) {
    if (!row.personId) continue;
    const names = accounts.get(joinKey(row.personId)) ?? new Set<string>();
    for (const cell of Object.values(row.cells)) {
      if (cell && typeof cell === "object" && !Array.isArray(cell)) {
        const principal = (cell as { principal?: unknown }).principal;
        if (typeof principal === "string" && principal.trim()) names.add(joinKey(principal));
      }
    }
    accounts.set(joinKey(row.personId), names);
  }
  return accounts;
}

export function itemIsLinked(
  item: Pick<StoredReceiptItem, "userId" | "userPrincipalName">,
  accounts: Map<string, Set<string>> | null,
): boolean {
  if (!accounts || !item.userId) return false;
  const names = accounts.get(joinKey(item.userId));
  if (!names) return false;
  return item.userPrincipalName === null || names.has(joinKey(item.userPrincipalName));
}

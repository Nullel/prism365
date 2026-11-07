import { HASHED, joinKey } from "@prism365/audits";
import {
  type AccessMembership,
  type AccessMembershipKind,
  type AuditRow,
  cleanText,
  MAX_ACCESS_AUDIT_ROWS,
  MAX_ACCESS_OWNED,
} from "@prism365/contracts";
import { GraphBoundaryError, GraphRequestError, GraphRetryDeferredError } from "@prism365/graph";
import { z } from "zod";

// What the Access tab does with what Microsoft and the audit store hand back: sorting objects into
// kinds, reading a failure in plain words, and picking one person's rows out of a stored audit
// run. Everything here is pure, so the route stays short and the rules can be tested alone.

const text = z
  .string()
  .max(1_000)
  .nullish()
  .catch(null)
  .transform((value) => (value ? cleanText(value.slice(0, 500)).trim() || null : null));
const flag = z.boolean().nullish().catch(null);

// A directory object as Microsoft returns it from a membership or ownership list. Without
// Directory.Read.All, Microsoft may return only the type and the id, so every other field is
// optional and a bad value reads as missing.
export const directoryObjectSchema = z.object({
  "@odata.type": text,
  id: z.string().min(1).max(200),
  displayName: text,
  groupTypes: z.array(z.unknown()).max(20).nullish().catch(null),
  mailEnabled: flag,
  securityEnabled: flag,
  membershipRule: text,
  roleTemplateId: text,
  resourceProvisioningOptions: z.array(z.unknown()).max(20).nullish().catch(null),
});
export type DirectoryObject = z.output<typeof directoryObjectSchema>;

const has = (list: readonly unknown[] | null | undefined, value: string) =>
  (list ?? []).some((item) => typeof item === "string" && item.toLowerCase() === value);

const typeOf = (object: DirectoryObject) => (object["@odata.type"] ?? "").toLowerCase();

// The kind of a group, from the three flags and the rule. A Microsoft 365 group is a Microsoft 365
// group first, even when its members come from a rule, because the mailbox is what an
// administrator looks for; the dynamic flag still says so.
function groupKind(object: DirectoryObject): AccessMembershipKind {
  if (has(object.groupTypes, "unified")) return "microsoft365";
  if (has(object.groupTypes, "dynamicmembership") || object.membershipRule) return "dynamic";
  if (object.mailEnabled && object.securityEnabled) return "mail-security";
  if (object.mailEnabled) return "distribution";
  return "security";
}

export function classifyMembership(
  object: DirectoryObject,
  directIds: ReadonlySet<string>,
): AccessMembership {
  const type = typeOf(object);
  const relation = directIds.has(object.id) ? "direct" : "nested";
  const base = {
    id: object.id,
    relation,
    hasGroupMailbox: false,
    dynamic: false,
    roleTemplateId: null,
  } as const;
  // Without Directory.Read.All Microsoft returns an object as a bare directoryObject with only an
  // id. One with neither a type nor a name is that; an object that has either is classified from it.
  if (type === "" || type === "#microsoft.graph.directoryobject") {
    if (object.displayName === null) return { ...base, kind: "hidden", name: null };
  }
  if (type === "#microsoft.graph.group")
    return {
      ...base,
      kind: groupKind(object),
      name: object.displayName,
      hasGroupMailbox: has(object.groupTypes, "unified"),
      dynamic: has(object.groupTypes, "dynamicmembership") || Boolean(object.membershipRule),
    };
  if (type === "#microsoft.graph.directoryrole")
    return {
      ...base,
      kind: "role",
      name: object.displayName,
      roleTemplateId: object.roleTemplateId,
    };
  if (type === "#microsoft.graph.administrativeunit")
    return { ...base, kind: "administrative-unit", name: object.displayName };
  return { ...base, kind: "hidden", name: null };
}

// Groups first, in the order the page lists kinds, then by name.
const kindOrder: readonly AccessMembershipKind[] = [
  "microsoft365",
  "security",
  "mail-security",
  "distribution",
  "dynamic",
  "role",
  "administrative-unit",
  "hidden",
];
export function sortMemberships(entries: readonly AccessMembership[]): AccessMembership[] {
  return [...entries].sort(
    (a, b) =>
      kindOrder.indexOf(a.kind) - kindOrder.indexOf(b.kind) ||
      (a.name ?? "").localeCompare(b.name ?? "") ||
      a.id.localeCompare(b.id),
  );
}

export type OwnedKind = "group" | "team" | "application" | "service-principal";
export interface OwnedEntry {
  id: string;
  kind: OwnedKind;
  name: string | null;
}

// Groups, teams, applications and service principals. Devices and other objects a person can own
// are not part of this page.
export function classifyOwned(object: DirectoryObject): OwnedEntry | null {
  const type = typeOf(object);
  const name = object.displayName;
  if (type === "#microsoft.graph.group")
    return {
      id: object.id,
      kind: has(object.resourceProvisioningOptions, "team") ? "team" : "group",
      name,
    };
  if (type === "#microsoft.graph.application") return { id: object.id, kind: "application", name };
  if (type === "#microsoft.graph.serviceprincipal")
    return { id: object.id, kind: "service-principal", name };
  return null;
}

// At most MAX_ACCESS_OWNED entries. `more` says the list was cut at the limit, which is the only
// way it is cut here: an object of another kind is left out because it is not part of this page,
// not because anything was missed.
export function ownedEntries(objects: readonly DirectoryObject[]): {
  entries: OwnedEntry[];
  more: boolean;
} {
  const entries: OwnedEntry[] = [];
  for (const object of objects) {
    const entry = classifyOwned(object);
    if (!entry) continue;
    if (entries.length >= MAX_ACCESS_OWNED) return { entries, more: true };
    entries.push(entry);
  }
  return { entries, more: false };
}

// A failed read in words an administrator can act on. Nothing from Microsoft's answer is quoted.
export function readFailure(error: unknown): string {
  if (error instanceof GraphRetryDeferredError)
    return "Microsoft asked Prism365 to slow down. Try again in a few minutes.";
  if (error instanceof GraphRequestError) {
    if (error.status === 404) return "Microsoft has no record of this account.";
    if (error.status === 403 || error.status === 401)
      return "Microsoft refused this read. The signed-in administrator may need a directory role such as Global Reader.";
    return "Microsoft could not answer this read just now. Try again.";
  }
  if (error instanceof GraphBoundaryError)
    return "This installation is not set up to make this read.";
  return "This could not be read just now. Try again.";
}

// The part of a stored audit run this page needs.
export interface StoredRun {
  finishedAt: string | null;
  scope: string[] | null;
  rows: readonly AuditRow[];
}

const mailboxAudit = "mail.mailbox-storage";

// The sign-in name used to find a mailbox row, compared the way the audits compare names. A name
// that is a hash is Microsoft hiding it, and is never matched against anything.
export function mailboxMatches(cell: unknown, principal: string): boolean {
  if (typeof cell !== "string") return false;
  const name = cell.trim();
  if (name === "" || HASHED.test(name)) return false;
  return joinKey(name) === joinKey(principal);
}

// The rows of a stored run that belong to one person: by the account id the row carries, or, for
// the mailbox report, which carries none, by sign-in name.
export function rowsFor(
  auditId: string,
  run: StoredRun,
  personId: string,
  principal: string | null,
): { rows: AuditRow[]; matchedBy: "account" | "sign-in name" } | "no_name" {
  if (auditId === mailboxAudit) {
    if (principal === null) return "no_name";
    return {
      matchedBy: "sign-in name",
      rows: run.rows.filter((row) => mailboxMatches(row.cells.mailbox, principal)),
    };
  }
  return { matchedBy: "account", rows: run.rows.filter((row) => row.personId === personId) };
}

export const auditRowLimit = MAX_ACCESS_AUDIT_ROWS;

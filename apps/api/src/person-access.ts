import { type Audit, RULES, type RowPick } from "@prism365/audits";
import {
  ACCESS_AUDIT_FIELDS,
  type AccessAuditResult,
  type AccessMembership,
  type AccessMembershipKind,
  type AuditRow,
  type CombinedRisksOk,
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
  // Read with the owned objects, to find an application registration in App consents.
  appId: text,
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
// An owned object as the route keeps it: the entry the page shows, and the application id an
// application registration is looked up by. The id is not sent to the page.
export interface OwnedObject extends OwnedEntry {
  appId: string | null;
}

// Groups, teams, applications and service principals. Devices and other objects a person can own
// are not part of this page.
export function classifyOwned(object: DirectoryObject): OwnedObject | null {
  const type = typeOf(object);
  const name = object.displayName;
  const appId = object.appId;
  if (type === "#microsoft.graph.group")
    return {
      id: object.id,
      kind: has(object.resourceProvisioningOptions, "team") ? "team" : "group",
      name,
      appId: null,
    };
  if (type === "#microsoft.graph.application")
    return { id: object.id, kind: "application", name, appId };
  if (type === "#microsoft.graph.serviceprincipal")
    return { id: object.id, kind: "service-principal", name, appId };
  return null;
}

// At most MAX_ACCESS_OWNED entries. `more` says the list was cut at the limit, which is the only
// way it is cut here: an object of another kind is left out because it is not part of this page,
// not because anything was missed.
export function ownedEntries(objects: readonly DirectoryObject[]): {
  entries: OwnedObject[];
  more: boolean;
} {
  const entries: OwnedObject[] = [];
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

// Audits whose rows carry no account id. The mailbox report names mailboxes by sign-in name, and
// the Microsoft 365 apps report names accounts by sign-in name inside the account cell, so a
// person's rows are found by their sign-in name, compared the way the audits compare names. A name
// that is a hash is Microsoft hiding it, and is never matched against anything.
export const NAME_MATCHED: Readonly<Record<string, readonly string[]>> = {
  "mail.mailbox-storage": ["mailbox"],
  "usage.m365-apps": ["person", "principal"],
};

// How to pick one person's rows out of an audit's run, or "no_name" when the audit needs a sign-in
// name and there is none.
export function pickFor(
  auditId: string,
  personId: string,
  principal: string | null,
): { pick: RowPick; matchedBy: "account" | "sign-in name" } | "no_name" {
  const path = NAME_MATCHED[auditId];
  if (!path) return { pick: { by: "personId", personId }, matchedBy: "account" };
  if (principal === null) return "no_name";
  return { pick: { by: "name", path, principal }, matchedBy: "sign-in name" };
}

export const auditRowLimit = MAX_ACCESS_AUDIT_ROWS;

type AccessRow = Extract<AccessAuditResult, { status: "ok" }>["rows"][number];

// Only the columns the person page shows for an audit (ACCESS_AUDIT_FIELDS), in that order, and
// only the ones a row has a value for.
export function accessRows(audit: Audit, rows: readonly AuditRow[]): AccessRow[] {
  const shown = ACCESS_AUDIT_FIELDS[audit.definition.id] ?? [];
  const columns = shown.flatMap((key) => {
    const column = audit.definition.columns.find((candidate) => candidate.key === key);
    return column ? [column] : [];
  });
  return rows.map((row) => ({
    key: row.key,
    cells: columns.flatMap((column) => {
      const value = row.cells[column.key];
      if (value === null || value === undefined) return [];
      if (Array.isArray(value) && value.length === 0) return [];
      return [{ key: column.key, label: column.label, kind: column.kind, value }];
    }),
  }));
}

// The part of a Combined risks run the person page needs.
export interface CombinedRun extends StoredRun {
  complete: boolean;
  notes: readonly string[];
}

const SEVERITY_RANK = { critical: 3, high: 2, medium: 1 } as const;

// The run's notes that say an input was missing or too old, or a rule was not checked. They are
// the audit's own sentences; the other notes describe coverage, not gaps.
const isGapNote = (note: string) =>
  note.startsWith("Not used:") || note.includes(" was not checked");

// One person's answer from a Combined risks run. The rules come from the row's tags, in the rule
// list's order, so a tag this version does not know is ignored rather than guessed at.
export function combinedRisksFor(run: CombinedRun, personId: string): CombinedRisksOk {
  const row = run.rows.find((candidate) => candidate.personId === personId);
  const tags = new Set(row?.tags ?? []);
  const rules = RULES.filter((rule) => tags.has(rule.id)).map((rule) => ({
    id: rule.id,
    label: rule.label,
    severity: rule.severity,
  }));
  const severity = rules.reduce<CombinedRisksOk["severity"]>(
    (best, rule) =>
      best === null || SEVERITY_RANK[rule.severity] > SEVERITY_RANK[best] ? rule.severity : best,
    null,
  );
  const reasons = row?.cells.reasons;
  return {
    status: "ok",
    finishedAt: run.finishedAt,
    severity,
    rules,
    reasons:
      rules.length > 0 && typeof reasons === "string" ? cleanText(reasons.slice(0, 2_000)) : "",
    gaps: run.notes
      .filter(isGapNote)
      .slice(0, 10)
      .map((note) => cleanText(note.slice(0, 500))),
    complete: run.complete,
  };
}

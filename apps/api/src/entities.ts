import type { Audit } from "@prism365/audits";
import {
  type AuditCell,
  type AuditRow,
  cleanText,
  ENTITY_AUDIT_FIELDS,
  type EntityList,
  type EntityMention,
  type EntityRow,
  MAX_ENTITY_LIST,
} from "@prism365/contracts";

// What the entity panel does with a stored row: reading its cells safely, choosing the columns the
// panel shows, and turning ids and names into things the page can link. Everything here is pure,
// so the route stays short and the rules can be tested alone. Cells were cleaned when the audit
// stored them; they are cleaned again here because this is the last place before the page.

const clean = (value: string, max: number) => cleanText(value.slice(0, max));

// A cell as short text, or null when it is not text or holds only white space.
export function textOf(cell: unknown, max = 500): string | null {
  if (typeof cell !== "string") return null;
  const value = clean(cell, max).trim();
  return value === "" ? null : value;
}

// A cell as an id in the one spelling the audits search by: lower case.
export function idOf(cell: unknown): string | null {
  const value = textOf(cell, 200);
  return value === null ? null : value.toLowerCase();
}

// The name and sign-in name of an account cell ({ name, principal }), either of which may be
// missing. A hashed or placeholder value is passed on as the audit stored it.
export function personCell(cell: unknown): { name: string | null; principal: string | null } {
  if (typeof cell !== "object" || cell === null || Array.isArray(cell))
    return { name: null, principal: null };
  const { name, principal } = cell as { name?: unknown; principal?: unknown };
  return { name: textOf(name), principal: textOf(principal) };
}

function cleanCell(value: AuditCell): AuditCell {
  if (typeof value === "string") return clean(value, 2_000);
  if (Array.isArray(value)) return value.slice(0, 200).map((item) => clean(item, 500));
  if (typeof value === "object" && value !== null)
    return {
      name: value.name === null ? null : clean(value.name, 500),
      principal: value.principal === null ? null : clean(value.principal, 500),
    };
  return value;
}

// The columns of a row that the panel shows for `auditId` (or the `fields` given), in order, and
// only the ones the row has a value for.
export function rowCells(
  audit: Audit | undefined,
  auditId: string,
  row: AuditRow,
  fields: readonly string[] = ENTITY_AUDIT_FIELDS[auditId] ?? [],
): EntityRow["cells"] {
  if (!audit) return [];
  return fields.flatMap((key) => {
    const column = audit.definition.columns.find((candidate) => candidate.key === key);
    const value = row.cells[key];
    if (!column || value === null || value === undefined) return [];
    if (Array.isArray(value) && value.length === 0) return [];
    return [{ key, label: column.label, kind: column.kind, value: cleanCell(value) }];
  });
}

export const mention = (
  kind: EntityMention["kind"],
  id: string | null,
  name: string | null,
  principal: string | null = null,
): EntityMention => ({ kind, id, name, principal });

// A list of things that other things name: at most MAX_ENTITY_LIST of them, and how many were left
// out. `names` is where a name is looked up; an id it does not know is shown without one.
export function listOf(
  label: string,
  kind: EntityMention["kind"],
  ids: readonly string[],
  names: ReadonlyMap<string, { name: string | null; principal?: string | null }>,
): EntityList {
  return {
    label,
    items: ids.slice(0, MAX_ENTITY_LIST).map((id) => {
      const known = names.get(id);
      return mention(kind, id, known?.name ?? null, known?.principal ?? null);
    }),
    more: Math.max(0, ids.length - MAX_ENTITY_LIST),
  };
}

// The strings of a list cell, as ids.
export function idsOf(cell: unknown): string[] {
  if (!Array.isArray(cell)) return [];
  return cell.flatMap((item) => {
    const id = idOf(item);
    return id === null ? [] : [id];
  });
}

// Who holds a Privileged roles row: an account (with its page), an app, or something the audit
// could not name or sort. An unnamed holder is shown by its id in the audit, which says nothing,
// so its name is left out here and the page says it is unknown.
export function holderMention(row: AuditRow): EntityMention {
  const holder = personCell(row.cells.holder);
  const holderId = idOf(row.cells.holderId);
  if (row.personId) return mention("person", row.personId, holder.name, holder.principal);
  const name = holder.name !== null && holder.name.toLowerCase() === holderId ? null : holder.name;
  if (row.cells.holderType === "Service principal") return mention("app", holderId, name);
  // "Group or other" is what the audit calls whatever it could not tell apart, which includes
  // accounts and apps it could not look up. It is not claimed to be a group here.
  return mention("other", holderId, name);
}

// The hidden role id and the role's name of a Privileged roles row, as the role it is a row of.
export function roleMention(row: AuditRow): EntityMention {
  return mention("role", idOf(row.cells.roleId), textOf(row.cells.role));
}

// Whether a name is worth showing to someone looking for it: used to order search results so a
// name that starts with what was typed comes before one that only contains it.
export function startsWith(name: string, text: string): boolean {
  return name.toLowerCase().startsWith(text.toLowerCase());
}

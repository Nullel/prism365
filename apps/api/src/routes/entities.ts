import { type Audit, findAudit, type PickedRows, pickRows, type RowPick } from "@prism365/audits";
import {
  type AuditRow,
  COMBINED_RISKS_AUDIT_ID,
  cleanText,
  decodeOwnerList,
  type Entity,
  type EntityKind,
  type EntityList,
  type EntityMention,
  type EntityRow,
  type EntitySearchResult,
  type EntitySection,
  entityParamsSchema,
  entitySchema,
  entitySearchQuerySchema,
  entitySearchSchema,
  MAX_ENTITY_LIST,
  MAX_ENTITY_ROWS,
  MAX_ENTITY_SEARCH,
  type OwnerEntry,
} from "@prism365/contracts";
import { privilegedRoles } from "@prism365/graph";
import type { SqlClient } from "@prism365/storage";
import type { FastifyInstance } from "fastify";
import {
  holderMention,
  idOf,
  idsOf,
  listOf,
  mention,
  personCell,
  roleMention,
  rowCells,
  startsWith,
  textOf,
} from "../entities.ts";
import { combinedRulesOf } from "../person-access.ts";
import type { RouteContext } from "./context.ts";

// The panel that opens when a group, role, Conditional Access policy, managed device or consented
// app is clicked (docs/audits.md, "Entity panel"). It reads only what the audits stored: the rows
// of each audit's latest succeeded run are picked in SQL by id, never by loading a whole run, and
// nothing is asked of Microsoft. The access-ledger entry is written before the first read and names
// the kind, never the thing. Every section stands alone: one that cannot be read says so and the
// rest still come back.

const AUDIT = {
  groups: "groups.teams",
  roles: "directory.privileged-roles",
  policies: "identity.conditional-access",
  devices: "devices.managed",
  apps: "apps.consents",
  credentials: "apps.credentials",
  combined: COMBINED_RISKS_AUDIT_ID,
  activity: "accounts.activity",
} as const;

// Each read runs in a transaction of its own: a statement that fails inside a PostgreSQL
// transaction ends the whole transaction, which would take the other sections down with it.
type Read = <T>(work: (sql: SqlClient) => Promise<T>) => Promise<T>;
type Picked = PickedRows | null | "failed";

const auditName = (auditId: string) => findAudit(auditId)?.definition.title ?? auditId;

async function pick(
  read: Read,
  tenantId: string,
  auditId: string,
  how: RowPick,
  limit: number,
): Promise<Picked> {
  try {
    return await read((sql) => pickRows(sql, tenantId, auditId, how, limit));
  } catch {
    return "failed";
  }
}

interface Spec {
  id: string;
  title: string;
  auditId: string;
}

// A section from what was picked. `rows` turns the picked rows into the panel's rows; `note` is
// a plain line for an answer that has no rows to show.
// The audit version from which a run keeps the hidden ids the panels match on (Privileged roles,
// Conditional Access, and the owner list of Groups and Teams). A run before it cannot say a thing
// is absent, only that it does not know.
const IDS_FROM_VERSION = 2;
const predatesIds = (picked: Picked) =>
  picked !== null && picked !== "failed" && picked.run.auditVersion < IDS_FROM_VERSION;

// `needsIds` marks a section that reads those hidden ids: a run from before them gives "outdated",
// never a negative sentence.
function sectionOf(
  spec: Spec,
  picked: Picked,
  rows: (found: AuditRow[]) => EntityRow[],
  note: (found: PickedRows) => string | null = () => null,
  needsIds = false,
): EntitySection {
  const base = {
    id: spec.id,
    title: spec.title,
    auditId: spec.auditId,
    auditName: auditName(spec.auditId),
  };
  if (picked === "failed")
    return {
      status: "failed",
      ...base,
      message: "The stored result could not be read just now. Try again.",
    };
  if (picked === null) return { status: "no_run", ...base };
  if (needsIds && predatesIds(picked))
    return { status: "outdated", ...base, finishedAt: picked.run.finishedAt };
  if (picked.rows.length === 0 && note(picked) === null)
    return {
      status: "not_in_run",
      ...base,
      finishedAt: picked.run.finishedAt,
      complete: picked.run.complete,
    };
  const built = rows(picked.rows);
  return {
    status: "ok",
    ...base,
    finishedAt: picked.run.finishedAt,
    complete: picked.run.complete,
    rows: built.slice(0, MAX_ENTITY_ROWS),
    more: picked.more || built.length > MAX_ENTITY_ROWS,
    note: note(picked),
  };
}

const rowOf = (
  key: string,
  parts: Partial<Omit<EntityRow, "key">> & Pick<EntityRow, "cells">,
): EntityRow => ({
  key: key.slice(0, 200),
  subject: null,
  tags: [],
  lists: [],
  risks: [],
  ...parts,
});

const auditOf = (auditId: string): Audit | undefined => findAudit(auditId);

// The first row of a single-row pick, if the section found one.
const firstRow = (picked: Picked): AuditRow | null =>
  picked !== null && picked !== "failed" ? (picked.rows[0] ?? null) : null;

// ---------------------------------------------------------------------------------------------
// Names for the ids a policy or an owner list holds
// ---------------------------------------------------------------------------------------------

type Names = Map<string, { name: string | null; principal?: string | null }>;

// Account names from the Account activity audit, which lists every account. An account it does not
// list stays unnamed and the page says so.
async function personNames(read: Read, tenantId: string, ids: readonly string[]): Promise<Names> {
  const names: Names = new Map();
  if (ids.length === 0) return names;
  const found = await pick(
    read,
    tenantId,
    AUDIT.activity,
    { by: "ids", personIds: ids },
    ids.length,
  );
  if (found !== null && found !== "failed")
    for (const row of found.rows) {
      const who = personCell(row.cells.person);
      if (row.personId) names.set(row.personId.toLowerCase(), who);
    }
  return names;
}

async function groupNames(read: Read, tenantId: string, ids: readonly string[]): Promise<Names> {
  const names: Names = new Map();
  if (ids.length === 0) return names;
  const found = await pick(read, tenantId, AUDIT.groups, { by: "ids", keys: ids }, ids.length);
  if (found !== null && found !== "failed")
    for (const row of found.rows)
      names.set(row.key.toLowerCase(), { name: textOf(row.cells.group) });
  return names;
}

// Role names: the built-in administrator roles Prism365 knows by id, then any role the Privileged
// roles audit has a row for. A role it has neither for stays unnamed.
async function roleNames(read: Read, tenantId: string, ids: readonly string[]): Promise<Names> {
  const names: Names = new Map();
  const unknown: string[] = [];
  for (const id of ids) {
    const known = privilegedRoles.get(id);
    if (known) names.set(id, { name: known });
    else unknown.push(id);
  }
  if (unknown.length > 0) {
    const found = await pick(
      read,
      tenantId,
      AUDIT.roles,
      { by: "ids", cells: [{ path: ["roleId"], values: unknown }] },
      MAX_ENTITY_ROWS * 4,
    );
    if (found !== null && found !== "failed")
      for (const row of found.rows) {
        const id = idOf(row.cells.roleId);
        if (id !== null) names.set(id, { name: textOf(row.cells.role) });
      }
  }
  return names;
}

// ---------------------------------------------------------------------------------------------
// Rows
// ---------------------------------------------------------------------------------------------

// A Conditional Access policy as a row for the group or role being looked at, saying how the
// policy reaches it: includes it, leaves it out, or both.
function policyRow(row: AuditRow, id: string, listKeys: { include: string; exclude: string }) {
  return rowOf(row.key, {
    subject: mention("policy", idOf(row.key), textOf(row.cells.policy)),
    tags: [
      ...(idsOf(row.cells[listKeys.include]).includes(id) ? ["Includes it"] : []),
      ...(idsOf(row.cells[listKeys.exclude]).includes(id) ? ["Leaves it out"] : []),
    ],
    cells: rowCells(auditOf(AUDIT.policies), AUDIT.policies, row, [
      "state",
      "appliesTo",
      "apps",
      "grant",
    ]),
  });
}

// A Privileged roles row as seen from the role: who holds it, and how. `resolved` holds the
// holders and groups the audit could not name or sort, as the other audits know them.
function holderRow(row: AuditRow, resolved: ReadonlyMap<string, EntityMention>): EntityRow {
  const through = idOf(row.cells.throughId);
  const holder = holderMention(row);
  return rowOf(row.key, {
    subject: (holder.id ? resolved.get(holder.id) : undefined) ?? holder,
    cells: rowCells(auditOf(AUDIT.roles), AUDIT.roles, row, [
      "holderType",
      "held",
      "through",
      "scope",
      "starts",
      "ends",
      "enabled",
      "mfa",
    ]),
    lists:
      through === null
        ? []
        : [
            {
              label: "Through group",
              items: [resolved.get(`through:${through}`) ?? mention("group", through, null)],
              more: 0,
            },
          ],
  });
}

// A Privileged roles row as seen from a group: the role, and the holder it reaches.
function roleOfGroupRow(row: AuditRow, groupId: string): EntityRow {
  const through = idOf(row.cells.throughId) === groupId;
  const direct = idOf(row.cells.holderId) === groupId;
  return rowOf(row.key, {
    subject: roleMention(row),
    tags: [...(through ? ["Held through it"] : []), ...(direct ? ["It holds the role"] : [])],
    cells: rowCells(auditOf(AUDIT.roles), AUDIT.roles, row, ["held", "through", "scope", "ends"]),
    lists: through ? [{ label: "Holder", items: [holderMention(row)], more: 0 }] : [],
  });
}

const tagsOf = (...tags: (string | null | false | undefined)[]): string[] =>
  tags.flatMap((tag) => (tag ? [tag.slice(0, 80)] : []));

// ---------------------------------------------------------------------------------------------
// One thing of each kind
// ---------------------------------------------------------------------------------------------

type Built = Omit<Entity, "kind" | "id">;

async function groupPanel(read: Read, tenantId: string, id: string): Promise<Built> {
  const own = await pick(read, tenantId, AUDIT.groups, { by: "ids", keys: [id] }, 1);
  const policies = await pick(
    read,
    tenantId,
    AUDIT.policies,
    {
      by: "ids",
      cells: [
        { path: ["includeGroupIds"], values: [id] },
        { path: ["excludeGroupIds"], values: [id] },
      ],
    },
    MAX_ENTITY_ROWS,
  );
  const roles = await pick(
    read,
    tenantId,
    AUDIT.roles,
    {
      by: "ids",
      cells: [
        { path: ["throughId"], values: [id] },
        { path: ["holderId"], values: [id] },
      ],
    },
    MAX_ENTITY_ROWS,
  );
  const row = firstRow(own);
  return {
    name: row ? textOf(row.cells.group) : null,
    tags: row
      ? tagsOf(
          textOf(row.cells.kind),
          textOf(row.cells.status) !== "OK" && textOf(row.cells.status),
          textOf(row.cells.visibility),
          row.cells.roleAssignable === true && "Role-assignable",
        )
      : [],
    notes: [
      "Members are not read, so this cannot say who is in the group. To see which groups a person belongs to, open the Access tab on their page.",
    ],
    sections: [
      sectionOf(
        { id: "group", title: "The group", auditId: AUDIT.groups },
        own,
        (found) =>
          found.map((item) =>
            rowOf(item.key, {
              cells: rowCells(auditOf(AUDIT.groups), AUDIT.groups, item),
              lists: ownerLists(item),
            }),
          ),
        () =>
          predatesIds(own) ? "Run Groups and Teams again to see which owners are accounts." : null,
      ),
      sectionOf(
        { id: "policies", title: "Conditional Access policies", auditId: AUDIT.policies },
        policies,
        (found) =>
          found.map((item) =>
            policyRow(item, id, { include: "includeGroupIds", exclude: "excludeGroupIds" }),
          ),
        (found) =>
          found.rows.length === 0 ? "No policy includes this group or leaves it out." : null,
        true,
      ),
      sectionOf(
        { id: "roles", title: "Roles it gives", auditId: AUDIT.roles },
        roles,
        (found) => found.map((item) => roleOfGroupRow(item, id)),
        (found) =>
          found.rows.length === 0
            ? "No role in the latest run is known to be held through this group."
            : null,
        true,
      ),
    ],
  };
}

// Owners as people to open: an account opens its page, an app opens the app, anything else is shown
// by name only.
function ownersAsList(found: readonly EntityMention[]): EntityList[] {
  if (found.length === 0) return [];
  return [
    {
      label: "Owners",
      items: found.slice(0, MAX_ENTITY_LIST),
      more: Math.max(0, found.length - MAX_ENTITY_LIST),
    },
  ];
}

const ownerMentions = (owners: readonly OwnerEntry[]): EntityMention[] =>
  owners.map((owner) =>
    mention(
      owner.kind === "user" ? "person" : owner.kind === "servicePrincipal" ? "app" : "other",
      owner.id.toLowerCase(),
      owner.name === null ? null : cleanText(owner.name),
    ),
  );

// The owners of a group row as people to open. Older results hold only the ids, which are shown
// without a name or a link, because nothing says what kind of object they are.
function ownerLists(row: AuditRow): EntityList[] {
  const owners = decodeOwnerList(row.cells.ownerList);
  return ownersAsList(
    owners.length > 0
      ? ownerMentions(owners)
      : idsOf(row.cells.ownerIds).map((ownerId) => mention("other", ownerId, null)),
  );
}

// The most holder rows read to check a role's holders against Combined risks, of which the panel
// shows the first MAX_ENTITY_ROWS. A role with more than this is checked in part, and says so.
const MAX_CHECKED_HOLDERS = 1_000;

// Holders the Privileged roles audit could not tell apart are looked up in the audits that know
// groups and apps, and are a group or an app only when one of them lists the id. Anything else
// stays an unnamed object that is not linked. A group a role comes through is named the same way.
async function holderMentions(
  read: Read,
  tenantId: string,
  rows: readonly AuditRow[],
): Promise<Map<string, EntityMention>> {
  const resolved = new Map<string, EntityMention>();
  const unknown = [
    ...new Set(
      rows.flatMap((row) => {
        const found = holderMention(row);
        return found.id && (found.kind === "other" || (found.kind === "app" && !found.name))
          ? [found.id]
          : [];
      }),
    ),
  ];
  const through = [...new Set(rows.flatMap((row) => idOf(row.cells.throughId) ?? []))].filter(
    (id) => !unknown.includes(id),
  );
  const groups = await groupNames(read, tenantId, [...unknown, ...through]);
  const apps = new Map<string, string | null>();
  const rest = unknown.filter((id) => !groups.has(id));
  if (rest.length > 0) {
    const found = await pick(read, tenantId, AUDIT.apps, { by: "ids", keys: rest }, rest.length);
    if (found !== null && found !== "failed")
      for (const row of found.rows) apps.set(row.key.toLowerCase(), textOf(row.cells.app));
  }
  for (const id of unknown) {
    const group = groups.get(id);
    const app = apps.get(id);
    const base = rows.map(holderMention).find((found) => found.id === id);
    if (group) resolved.set(id, mention("group", id, group.name));
    else if (apps.has(id) || base?.kind === "app")
      resolved.set(id, mention("app", id, base?.name ?? app ?? null));
  }
  for (const id of through)
    resolved.set(`through:${id}`, mention("group", id, groups.get(id)?.name ?? null));
  return resolved;
}

async function rolePanel(read: Read, tenantId: string, id: string): Promise<Built> {
  const holders = await pick(
    read,
    tenantId,
    AUDIT.roles,
    { by: "ids", cells: [{ path: ["roleId"], values: [id] }] },
    MAX_CHECKED_HOLDERS,
  );
  const rowsRead = holders !== null && holders !== "failed" ? holders.rows : [];
  const people = [
    ...new Set(rowsRead.flatMap((row) => (row.personId ? [row.personId.toLowerCase()] : []))),
  ];
  const outdated = predatesIds(holders);
  // Combined risks says something only about holders there are, so it is read, and its section
  // shown, only when the role has some. Every holder read is checked, not only those shown.
  const hasHolders = !outdated && rowsRead.length > 0;
  const risks = hasHolders
    ? await pick(
        read,
        tenantId,
        AUDIT.combined,
        { by: "ids", personIds: people },
        Math.max(1, people.length),
      )
    : null;
  const policies = await pick(
    read,
    tenantId,
    AUDIT.policies,
    {
      by: "ids",
      cells: [
        { path: ["includeRoleIds"], values: [id] },
        { path: ["excludeRoleIds"], values: [id] },
      ],
    },
    MAX_ENTITY_ROWS,
  );
  const first = firstRow(holders);
  const name = first ? textOf(first.cells.role) : (privilegedRoles.get(id) ?? null);
  const cut = holders !== null && holders !== "failed" && holders.more;
  const shownRows = rowsRead.slice(0, MAX_ENTITY_ROWS);
  const resolved = await holderMentions(read, tenantId, shownRows);
  const shown =
    holders !== null && holders !== "failed"
      ? { ...holders, rows: shownRows, more: holders.more || rowsRead.length > MAX_ENTITY_ROWS }
      : holders;
  const checked = rowsRead.length.toLocaleString("en-US");
  return {
    name,
    tags: tagsOf(
      first?.cells.privileged === true && !outdated && "Privileged role",
      !outdated &&
        rowsRead.length > 0 &&
        `${rowsRead.length}${cut ? " or more" : ""} ${rowsRead.length === 1 && !cut ? "assignment" : "assignments"}`,
    ),
    notes: [],
    sections: [
      sectionOf(
        { id: "holders", title: "Who holds it", auditId: AUDIT.roles },
        shown,
        (found) => found.map((item) => holderRow(item, resolved)),
        () => null,
        true,
      ),
      ...(hasHolders
        ? [
            sectionOf(
              {
                id: "risks",
                title: "Combined risks among the people who hold it",
                auditId: AUDIT.combined,
              },
              risks !== null && risks !== "failed"
                ? {
                    ...risks,
                    more:
                      risks.rows.filter((item) => combinedRulesOf(item.tags).length > 0).length >
                      MAX_ENTITY_ROWS,
                  }
                : risks,
              (found) =>
                found.flatMap((item) => {
                  const rules = combinedRulesOf(item.tags);
                  return item.personId && rules.length > 0
                    ? [
                        rowOf(item.key, {
                          subject: mention(
                            "person",
                            item.personId,
                            personCell(item.cells.person).name,
                            personCell(item.cells.person).principal,
                          ),
                          risks: rules,
                          cells: [],
                        }),
                      ]
                    : [];
                }),
              (found) => {
                const tripped = found.rows.some((item) => combinedRulesOf(item.tags).length > 0);
                // A negative is only said over the holders that were all checked.
                if (cut)
                  return `Checked the first ${checked} holders. The role has more, which were not checked.`;
                if (tripped) return null;
                return people.length === 0
                  ? "None of the holders in this list is a person."
                  : "None of the people who hold this role trips a rule in this run.";
              },
            ),
          ]
        : []),
      sectionOf(
        { id: "policies", title: "Conditional Access policies", auditId: AUDIT.policies },
        policies,
        (found) =>
          found.map((item) =>
            policyRow(item, id, { include: "includeRoleIds", exclude: "excludeRoleIds" }),
          ),
        (found) =>
          found.rows.length === 0 ? "No policy includes this role or leaves it out." : null,
        true,
      ),
    ],
  };
}

async function policyPanel(read: Read, tenantId: string, id: string): Promise<Built> {
  const own = await pick(read, tenantId, AUDIT.policies, { by: "ids", keys: [id] }, 1);
  const row = firstRow(own);
  const listed = (key: string) => (row ? idsOf(row.cells[key]) : []);
  // Names are looked up for the ids the lists show, which is the first MAX_ENTITY_LIST of each.
  const shown = (...keys: string[]) => [
    ...new Set(keys.flatMap((key) => listed(key).slice(0, MAX_ENTITY_LIST))),
  ];
  const keys = {
    userIds: shown("includeUserIds", "excludeUserIds"),
    groupIds: shown("includeGroupIds", "excludeGroupIds"),
    roleIds: shown("includeRoleIds", "excludeRoleIds"),
  };
  const people = await personNames(read, tenantId, keys.userIds);
  const groups = await groupNames(read, tenantId, keys.groupIds);
  const roles = await roleNames(read, tenantId, keys.roleIds);
  return {
    name: row ? textOf(row.cells.policy) : null,
    tags: row ? tagsOf(textOf(row.cells.state)) : [],
    notes: [],
    sections: [
      sectionOf(
        { id: "policy", title: "The policy", auditId: AUDIT.policies },
        own,
        (found) =>
          found.map((item) => {
            const lists = (
              [
                ["Includes these people", "person", "includeUserIds", people],
                ["Includes these groups", "group", "includeGroupIds", groups],
                ["Includes these roles", "role", "includeRoleIds", roles],
                ["Leaves out these people", "person", "excludeUserIds", people],
                ["Leaves out these groups", "group", "excludeGroupIds", groups],
                ["Leaves out these roles", "role", "excludeRoleIds", roles],
              ] as const
            ).flatMap(([label, kind, key, names]) => {
              const ids = idsOf(item.cells[key]);
              return ids.length === 0 ? [] : [listOf(label, kind, ids, names)];
            });
            return rowOf(item.key, {
              cells: rowCells(auditOf(AUDIT.policies), AUDIT.policies, item),
              lists,
            });
          }),
        () =>
          predatesIds(own)
            ? "Run Conditional Access again to see who this policy includes and leaves out."
            : null,
      ),
    ],
  };
}

async function devicePanel(read: Read, tenantId: string, id: string): Promise<Built> {
  const own = await pick(read, tenantId, AUDIT.devices, { by: "ids", keys: [id] }, 1);
  const row = firstRow(own);
  return {
    name: row ? textOf(row.cells.device) : null,
    tags: row
      ? tagsOf(textOf(row.cells.os), textOf(row.cells.compliance), textOf(row.cells.ownership))
      : [],
    notes: [],
    sections: [
      sectionOf({ id: "device", title: "The device", auditId: AUDIT.devices }, own, (found) =>
        found.map((item) => {
          const owner = personCell(item.cells.person);
          return rowOf(item.key, {
            cells: rowCells(auditOf(AUDIT.devices), AUDIT.devices, item),
            lists: item.personId
              ? [
                  {
                    label: "Owner",
                    items: [mention("person", item.personId, owner.name, owner.principal)],
                    more: 0,
                  },
                ]
              : [],
          });
        }),
      ),
    ],
  };
}

async function appPanel(read: Read, tenantId: string, id: string): Promise<Built> {
  // An app is listed by its service principal's id, and also carries its application id.
  const own = await pick(
    read,
    tenantId,
    AUDIT.apps,
    { by: "ids", keys: [id], cells: [{ path: ["appId"], values: [id] }] },
    1,
  );
  const row = firstRow(own);
  // The credentials audit lists the registration, whose own id is not this one. It is found by the
  // application id, or by the service principal id it keeps for the link back to this panel.
  const appId = row ? idOf(row.cells.appId) : null;
  const credentials = await pick(
    read,
    tenantId,
    AUDIT.credentials,
    {
      by: "ids",
      cells: [
        ...(appId ? [{ path: ["appId"], values: [appId] }] : []),
        { path: ["servicePrincipalId"], values: [id] },
      ],
    },
    1,
  );
  const highRisk = row ? idsOfText(row.cells.highRisk) : [];
  return {
    name: row ? textOf(row.cells.app) : textOf(firstRow(credentials)?.cells.app),
    tags: row
      ? tagsOf(
          textOf(row.cells.publisher),
          textOf(row.cells.owner),
          highRisk.length > 0 && "High-risk permissions",
          row.cells.enabled === false && "Disabled",
        )
      : [],
    notes: [],
    sections: [
      sectionOf({ id: "app", title: "The app", auditId: AUDIT.apps }, own, (found) =>
        found.map((item) =>
          rowOf(item.key, { cells: rowCells(auditOf(AUDIT.apps), AUDIT.apps, item) }),
        ),
      ),
      sectionOf(
        { id: "credentials", title: "Credentials", auditId: AUDIT.credentials },
        credentials,
        (found) =>
          found.map((item) =>
            rowOf(item.key, {
              cells: rowCells(auditOf(AUDIT.credentials), AUDIT.credentials, item),
              lists: ownersAsList(ownerMentions(decodeOwnerList(item.cells.ownerIds))),
            }),
          ),
      ),
    ],
  };
}

function idsOfText(cell: unknown): string[] {
  return Array.isArray(cell) ? cell.filter((item): item is string => typeof item === "string") : [];
}

const panels: Record<EntityKind, (read: Read, tenantId: string, id: string) => Promise<Built>> = {
  group: groupPanel,
  role: rolePanel,
  policy: policyPanel,
  device: devicePanel,
  app: appPanel,
};

// ---------------------------------------------------------------------------------------------
// Search
// ---------------------------------------------------------------------------------------------

// Where each kind's name lives in its audit, how to read the id of a result, and the one word that
// tells two results with the same name apart.
const SEARCHED: readonly {
  kind: EntityKind;
  auditId: string;
  name: string;
  id: (row: AuditRow) => string | null;
  detail: (row: AuditRow) => string | null;
  // A cell that is the same on every row about one thing, so a thing is read once.
  distinct?: string[];
}[] = [
  {
    kind: "group",
    auditId: AUDIT.groups,
    name: "group",
    id: (row) => idOf(row.key),
    detail: (row) => textOf(row.cells.kind, 100),
  },
  {
    kind: "role",
    auditId: AUDIT.roles,
    name: "role",
    id: (row) => idOf(row.cells.roleId),
    detail: () => null,
    distinct: ["roleId"],
  },
  {
    kind: "policy",
    auditId: AUDIT.policies,
    name: "policy",
    id: (row) => idOf(row.key),
    detail: (row) => textOf(row.cells.state, 100),
  },
  {
    kind: "device",
    auditId: AUDIT.devices,
    name: "device",
    id: (row) => idOf(row.key),
    detail: (row) => textOf(row.cells.os, 100),
  },
  {
    kind: "app",
    auditId: AUDIT.apps,
    name: "app",
    id: (row) => idOf(row.key),
    detail: (row) => textOf(row.cells.publisher, 100),
  },
  // Application registrations, found by the enterprise application the panel opens. One that has
  // none cannot be opened and is left out.
  {
    kind: "app",
    auditId: AUDIT.credentials,
    name: "app",
    id: (row) => idOf(row.cells.servicePrincipalId),
    detail: (row) => textOf(row.cells.audience, 100),
  },
];

// Matching rows read from one audit, which is more than are shown so the kinds can share the 20.
const SEARCH_READ = 100;

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

export function registerEntityRoutes(app: FastifyInstance, context: RouteContext): void {
  app.get("/v1/entities/search", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session)
      return reply.status(401).send({
        error: "authentication_required",
        message: "Sign in with an authorized administrator account.",
      });
    const query = entitySearchQuerySchema.safeParse(request.query);
    if (!query.success)
      return reply.status(400).send({
        error: "invalid_request",
        message: "Search for 2 to 100 characters.",
      });
    const { modules } = context;
    if (!modules)
      return reply.status(409).send({
        error: "modules_unavailable",
        message:
          "Searching groups, roles and devices reads stored results, which need module storage.",
      });
    const { tenantId } = session;
    // The ledger entry comes before anything is read. It names the datasets, never the words.
    await context.store.appendAudit(tenantId, session.actorId, {
      type: "entity.searched",
      datasets: ["auditResults"],
      occurredAt: new Date().toISOString(),
    });
    const read: Read = (work) => modules.database.tenant(tenantId, work);
    const text = query.data.q;
    // What each kind found, whichever audits it came from, and the ids already listed: an app is
    // in two audits, and is listed once, by its enterprise application id or its application id.
    const found = new Map<EntityKind, EntitySearchResult[]>();
    const listed = new Map<EntityKind, Set<string>>();
    // True when a source held more matching rows than were read, so there is more to find.
    let more = false;
    let partial = false;
    for (const source of SEARCHED) {
      const picked = await pick(
        read,
        tenantId,
        source.auditId,
        {
          by: "contains",
          path: [source.name],
          text,
          ...(source.distinct ? { distinct: source.distinct } : {}),
        },
        SEARCH_READ,
      );
      if (picked === "failed") {
        partial = true;
        continue;
      }
      if (picked?.more) more = true;
      // A run from before roles kept their ids cannot be searched for roles.
      if (source.kind === "role" && predatesIds(picked)) partial = true;
      const results = found.get(source.kind) ?? [];
      const seen = listed.get(source.kind) ?? new Set<string>();
      found.set(source.kind, results);
      listed.set(source.kind, seen);
      for (const row of picked?.rows ?? []) {
        const id = source.id(row);
        const name = textOf(row.cells[source.name]);
        if (id === null || name === null || !GUID.test(id) || seen.has(id)) continue;
        const appId = source.kind === "app" ? idOf(row.cells.appId) : null;
        if (appId !== null && seen.has(`app:${appId}`)) continue;
        seen.add(id);
        if (appId !== null) seen.add(`app:${appId}`);
        results.push({ kind: source.kind, id, name, detail: source.detail(row) });
      }
    }
    const queues = [...found.values()].map((results) =>
      // A name that starts with the words typed comes before one that only holds them.
      results.sort(
        (a, b) =>
          Number(startsWith(b.name, text)) - Number(startsWith(a.name, text)) ||
          a.name.localeCompare(b.name),
      ),
    );
    // One of each kind in turn, so a query that matches fifty groups still shows the one device.
    const results: EntitySearchResult[] = [];
    for (let round = 0; queues.some((queue) => queue.length > round); round += 1)
      for (const queue of queues) {
        const next = queue[round];
        if (!next) continue;
        if (results.length < MAX_ENTITY_SEARCH) results.push(next);
        else more = true;
      }
    return entitySearchSchema.parse({ results, more, partial });
  });

  app.get("/v1/entities/:kind/:id", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session)
      return reply.status(401).send({
        error: "authentication_required",
        message: "Sign in with an authorized administrator account.",
      });
    const params = entityParamsSchema.safeParse(request.params);
    if (!params.success)
      return reply
        .status(400)
        .send({ error: "invalid_request", message: "Unknown group, role, policy, device or app." });
    const { modules } = context;
    if (!modules)
      return reply.status(409).send({
        error: "modules_unavailable",
        message: "This page is read from stored results, which need module storage.",
      });
    const { tenantId } = session;
    const { kind, id } = params.data;
    // The ledger entry comes before anything is read. It names the kind, never the thing.
    await context.store.appendAudit(tenantId, session.actorId, {
      type: "entity.viewed",
      kind,
      datasets: ["auditResults"],
      occurredAt: new Date().toISOString(),
    });
    const read: Read = (work) => modules.database.tenant(tenantId, work);
    const built = await panels[kind](read, tenantId, id);
    return entitySchema.parse({ kind, id, ...built });
  });
}

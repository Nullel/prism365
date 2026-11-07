import type { AuditRow } from "@prism365/contracts";
import { privilegedRoles } from "@prism365/graph";
import type { StoredFinding, StoredResults, StoredRun } from "./definition.ts";
import { HASHED } from "./m365-app-usage.ts";

// The audits the Combined risks audit reads, the columns it reads from each, and how their rows
// become one record per person. The column keys are listed once, here, and a test checks every
// one against the audit that publishes it, so renaming a column fails a test instead of silently
// emptying a rule.

export const INPUT_IDS = [
  "activity",
  "inactive",
  "reclaim",
  "mfa",
  "guests",
  "risky",
  "roles",
  "passwords",
  "devices",
  "mailbox",
  "groups",
] as const;
export type InputId = (typeof INPUT_IDS)[number];

export interface InputSpec {
  auditId: string;
  title: string;
  // Every column of that audit the joiner reads.
  columns: readonly string[];
}

export const INPUTS: Record<InputId, InputSpec> = {
  activity: {
    auditId: "accounts.activity",
    title: "Account activity",
    columns: ["person", "status", "type", "lastSuccess", "days", "verdict", "licenses", "groupIds"],
  },
  inactive: {
    auditId: "accounts.inactive",
    title: "Inactive accounts",
    columns: ["person", "status", "type", "lastSuccess", "days", "licenses"],
  },
  reclaim: {
    auditId: "licenses.reclaim",
    title: "Licenses to reclaim",
    columns: ["person", "status", "licenses"],
  },
  mfa: {
    auditId: "mfa.registration",
    title: "MFA registration",
    columns: ["person", "registered", "type"],
  },
  guests: { auditId: "guests.review", title: "Guest accounts", columns: ["person", "status"] },
  risky: {
    auditId: "identity.risky-users",
    title: "Risky users",
    columns: ["person", "level", "state"],
  },
  roles: {
    auditId: "directory.privileged-roles",
    title: "Privileged roles",
    columns: ["holder", "holderType", "role", "privileged", "held", "enabled"],
  },
  passwords: {
    auditId: "accounts.passwords",
    title: "Passwords",
    columns: ["person", "enabled", "days", "neverExpires", "weak"],
  },
  devices: {
    auditId: "devices.managed",
    title: "Managed devices",
    columns: ["device", "compliance", "encrypted"],
  },
  mailbox: {
    auditId: "mail.mailbox-storage",
    title: "Mailbox storage",
    columns: ["mailbox", "recipientType"],
  },
  groups: {
    auditId: "groups.teams",
    title: "Groups and Teams",
    columns: ["group", "ownerIds"],
  },
};

// The column Groups and Teams carries for ownership: the directory IDs of a group's owners. It is
// hidden on that audit's page and exists only for this join.
export const OWNER_IDS_COLUMN = "ownerIds";

// The module whose open findings feed the sign-in rule.
export const SIGNINS_MODULE = "signins.audit";

export const FRESH_DAYS = 7;
const DAY = 86_400_000;

export type Freshness =
  | { state: "fresh"; finishedAt: string; run: StoredRun }
  | { state: "stale"; finishedAt: string }
  | { state: "never" };

// An input is usable when its newest successful run finished no more than seven days before this
// run. Exactly seven days is still fresh. A finish time that cannot be read, or lies in the
// future, is not trusted as fresh either way: unreadable is stale, and the future is read as now.
export function freshnessOf(run: StoredRun | null, now: Date): Freshness {
  if (!run) return { state: "never" };
  const finished = Date.parse(run.finishedAt);
  if (Number.isNaN(finished)) return { state: "stale", finishedAt: run.finishedAt };
  return now.getTime() - finished > FRESH_DAYS * DAY
    ? { state: "stale", finishedAt: run.finishedAt }
    : { state: "fresh", finishedAt: run.finishedAt, run };
}

export type Inputs = Record<InputId, Freshness>;

export async function readInputs(stored: StoredResults, now: Date): Promise<Inputs> {
  const entries = await Promise.all(
    INPUT_IDS.map(async (id) => [id, freshnessOf(await stored.latestRun(INPUTS[id].auditId), now)]),
  );
  return Object.fromEntries(entries) as Inputs;
}

// One person's facts, joined from every fresh input. A fact that is null or empty is unknown or
// absent, and no rule treats absence as evidence: a rule trips only on something an input says.
export interface RoleHold {
  role: string;
  privileged: boolean;
  held: string;
  // The role template ID, known for the privileged built-in roles only.
  templateId: string | null;
}

export interface DeviceFact {
  name: string | null;
  compliance: string | null;
  encrypted: boolean | null;
}

export interface Person {
  id: string;
  name: string | null;
  principal: string | null;
  enabled: boolean | null;
  guest: boolean;
  lastSuccess: string | null;
  // Days since the last successful sign-in, counted as the Inactive accounts audit counts them.
  days: number | null;
  verdict: string | null;
  licenses: string[];
  // Direct group IDs, or null when memberships were not read.
  groupIds: string[] | null;
  mfa: boolean | null;
  risk: { level: string; state: string } | null;
  roles: RoleHold[];
  password: { neverExpires: boolean; weak: boolean; days: number | null } | null;
  devices: DeviceFact[];
  mailboxType: string | null;
  findings: StoredFinding[];
  // Group IDs the person owns, from Groups and Teams. Empty when that audit is not usable, and
  // for groups whose owners it did not read.
  ownedGroups: string[];
}

const templateByRole = new Map([...privilegedRoles].map(([id, name]) => [name, id]));

const text = (cell: unknown): string | null =>
  typeof cell === "string" && cell.trim() !== "" ? cell : null;
const bool = (cell: unknown): boolean | null => (typeof cell === "boolean" ? cell : null);
const whole = (cell: unknown): number | null =>
  typeof cell === "number" && Number.isFinite(cell) ? cell : null;
const strings = (cell: unknown): string[] | null =>
  Array.isArray(cell) ? cell.filter((item): item is string => typeof item === "string") : null;

function personOf(cell: unknown): { name: string | null; principal: string | null } {
  if (typeof cell !== "object" || cell === null || Array.isArray(cell))
    return { name: null, principal: null };
  const { name, principal } = cell as { name?: unknown; principal?: unknown };
  return { name: text(name), principal: text(principal) };
}

export function emptyPerson(id: string): Person {
  return {
    id,
    name: null,
    principal: null,
    enabled: null,
    guest: false,
    lastSuccess: null,
    days: null,
    verdict: null,
    licenses: [],
    groupIds: null,
    mfa: null,
    risk: null,
    roles: [],
    password: null,
    devices: [],
    mailboxType: null,
    findings: [],
    ownedGroups: [],
  };
}

export interface Joined {
  people: Map<string, Person>;
  // Mailbox rows skipped because their name is a hash: they can never be matched to anyone.
  hashedMailboxes: number;
  // Fresh inputs whose run was limited to some domains, or only partly read.
  limited: InputId[];
  partial: InputId[];
}

const enabledOf = (status: unknown): boolean | null =>
  status === "Enabled" ? true : status === "Disabled" ? false : null;

// Joins the fresh inputs by person ID. Where two inputs state the same fact, the earlier one in
// INPUT_IDS wins, which puts Account activity (the most complete) first. The mailbox is matched by
// sign-in name, because that is all the report carries.
export function joinInputs(
  inputs: Inputs,
  findings: readonly StoredFinding[] | null,
  usable: ReadonlySet<InputId>,
): Joined {
  const people = new Map<string, Person>();
  const person = (id: string, name?: { name: string | null; principal: string | null }) => {
    let found = people.get(id);
    if (!found) {
      found = emptyPerson(id);
      people.set(id, found);
    }
    if (name) {
      found.name ??= name.name;
      found.principal ??= name.principal;
    }
    return found;
  };
  const rowsOf = (id: InputId): readonly AuditRow[] => {
    const input = inputs[id];
    return usable.has(id) && input.state === "fresh" ? input.run.rows : [];
  };
  const identified = (row: AuditRow): row is AuditRow & { personId: string } =>
    row.personId !== null;

  const limited: InputId[] = [];
  const partial: InputId[] = [];
  for (const id of INPUT_IDS) {
    const input = inputs[id];
    if (!usable.has(id) || input.state !== "fresh") continue;
    if (input.run.scope && input.run.scope.length > 0) limited.push(id);
    if (!input.run.complete) partial.push(id);
  }

  // Sign-in dates, license names and status: the first of these inputs that names the person.
  const dated = new Set<string>();
  for (const id of ["activity", "inactive", "reclaim"] as const) {
    for (const row of rowsOf(id).filter(identified)) {
      const who = person(row.personId, personOf(row.cells.person));
      who.enabled ??= enabledOf(row.cells.status);
      if (row.cells.type === "Guest") who.guest = true;
      if (id !== "reclaim" && !dated.has(row.personId)) {
        dated.add(row.personId);
        who.lastSuccess = text(row.cells.lastSuccess);
        who.days = whole(row.cells.days);
        who.verdict = id === "activity" ? text(row.cells.verdict) : null;
      }
      const licenses = strings(row.cells.licenses);
      if (licenses && who.licenses.length === 0) who.licenses = licenses;
      if (id === "activity") who.groupIds = strings(row.cells.groupIds);
    }
  }
  for (const row of rowsOf("mfa").filter(identified)) {
    const who = person(row.personId, personOf(row.cells.person));
    who.mfa = bool(row.cells.registered);
    if (row.cells.type === "Guest") who.guest = true;
  }
  for (const row of rowsOf("guests").filter(identified))
    person(row.personId, personOf(row.cells.person)).guest = true;
  for (const row of rowsOf("risky").filter(identified)) {
    const level = text(row.cells.level);
    const state = text(row.cells.state);
    if (level && state) person(row.personId, personOf(row.cells.person)).risk = { level, state };
  }
  for (const row of rowsOf("roles").filter(identified)) {
    const role = text(row.cells.role);
    if (!role) continue;
    const who = person(row.personId, personOf(row.cells.holder));
    who.enabled ??= bool(row.cells.enabled);
    if (row.cells.holderType === "Guest") who.guest = true;
    who.roles.push({
      role,
      privileged: row.cells.privileged === true,
      held: text(row.cells.held) ?? "Active",
      templateId: templateByRole.get(role) ?? null,
    });
  }
  for (const row of rowsOf("passwords").filter(identified)) {
    const who = person(row.personId, personOf(row.cells.person));
    who.enabled ??= enabledOf(row.cells.enabled);
    who.password = {
      neverExpires: row.cells.neverExpires === true,
      weak: row.cells.weak === true,
      days: whole(row.cells.days),
    };
  }
  for (const row of rowsOf("devices").filter(identified))
    person(row.personId).devices.push({
      name: text(row.cells.device),
      compliance: text(row.cells.compliance),
      encrypted: bool(row.cells.encrypted),
    });

  // Group ownership: each owner ID that is a person already known from the other inputs. An owner
  // nobody else mentions cannot trip a rule, so no record is made for them.
  for (const row of rowsOf("groups")) {
    const owners = strings(row.cells[OWNER_IDS_COLUMN]);
    for (const ownerId of new Set(owners ?? [])) people.get(ownerId)?.ownedGroups.push(row.key);
  }

  // Mailboxes by sign-in name. A hashed name matches nobody, however it was produced.
  let hashedMailboxes = 0;
  const byPrincipal = new Map<string, Person>();
  for (const who of people.values())
    if (who.principal) byPrincipal.set(who.principal.toLowerCase(), who);
  for (const row of rowsOf("mailbox")) {
    const mailbox = text(row.cells.mailbox);
    if (!mailbox) continue;
    if (HASHED.test(mailbox)) {
      hashedMailboxes += 1;
      continue;
    }
    const owner = byPrincipal.get(mailbox.toLowerCase());
    if (owner) owner.mailboxType = text(row.cells.recipientType);
  }

  for (const finding of findings ?? [])
    if (finding.subjectKind === "person" && people.has(finding.subjectId))
      people.get(finding.subjectId)?.findings.push(finding);
  return { people, hashedMailboxes, limited, partial };
}

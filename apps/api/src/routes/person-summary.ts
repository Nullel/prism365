import {
  type Audit,
  builtInAudits,
  HASHED,
  joinKey,
  type PickedRows,
  pickRows,
  type RowPick,
} from "@prism365/audits";
import {
  type AuditRow,
  COMBINED_RISKS_AUDIT_ID,
  cleanText,
  MAX_ACCESS_AUDIT_ROWS,
  MAX_ACCESS_OWNED,
  MAX_MAILBOX_HOLDERS,
  MAX_PERSON_MAILBOX_GRANTS,
  MAX_SUMMARY_FINDINGS,
  normalizeUserId,
  type PersonSummary,
  personIdSchema,
  personSummarySchema,
  SUMMARY_AUDIT_IDS,
  SUMMARY_CHANGE_DAYS,
  type Severity,
  type SummaryAudit,
} from "@prism365/contracts";
import { countPersonChanges, DIRECTORY_CHANGES_MODULE } from "@prism365/module-directory-changes";
import {
  MODULE_ID as SIGNINS_MODULE,
  principalTakenSince,
  storedPrincipals,
} from "@prism365/module-signins-audit";
import { listFindings, severityAtLeast } from "@prism365/modules";
import type { SqlClient } from "@prism365/storage";
import type { FastifyInstance } from "fastify";
import { z } from "zod";
import { personReceiptItems } from "../handoff-store.ts";
import { currentImport, personMailboxRead } from "../mailbox-access-store.ts";
import { accessRows, combinedRisksFor, NAME_MATCHED, pickFor } from "../person-access.ts";
import type { RouteContext } from "./context.ts";

// The Overview's "Where they show up" panel (docs/people.md): every place Prism365 knows one
// person from, in one answer, so the page can say what each source thinks of them before the
// administrator opens any tab. It reads only what was stored, never asks Microsoft anything, and
// writes its access-ledger entry before the first read. Every section stands alone: one that
// cannot be read says so and the rest still come back.

const GROUPS_AUDIT = "groups.teams";
const DAY_MS = 86_400_000;
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// Each section is read in a transaction of its own: a statement that fails inside a PostgreSQL
// transaction ends the whole transaction, which would take the other sections down with it.
type Read = <T>(work: (sql: SqlClient) => Promise<T>) => Promise<T>;

const failed = (what: string) => ({
  status: "failed" as const,
  message: `${what} could not be read just now. Try again.`,
});

// The sign-in names the rows carry in their account cell, hashed ones left out.
export function principalsFromRows(rows: readonly AuditRow[]): string[] {
  const names: string[] = [];
  for (const row of rows) {
    const person = row.cells.person;
    if (typeof person !== "object" || person === null || Array.isArray(person)) continue;
    const principal = (person as { principal?: unknown }).principal;
    if (typeof principal === "string" && principal.trim() !== "" && !HASHED.test(principal.trim()))
      names.push(principal.trim());
  }
  return names;
}

// Every person-level audit's answer for one account, in the order the page lists them, and the
// sign-in name found along the way. The audits that carry an account id are read first, since
// their rows can give the sign-in name that the audits without one need.
async function auditStates(
  read: Read,
  tenantId: string,
  userId: string,
  storedName: string | null,
  nameTaken: (name: string) => Promise<boolean>,
): Promise<{ states: SummaryAudit[]; principal: string | null }> {
  const wanted = SUMMARY_AUDIT_IDS.flatMap((id) => {
    const audit = builtInAudits.find((candidate) => candidate.definition.id === id);
    return audit ? [audit] : [];
  });
  const pickFrom = async (audit: Audit, pick: RowPick): Promise<PickedRows | null | "failed"> => {
    try {
      return await read((sql) =>
        pickRows(sql, tenantId, audit.definition.id, pick, MAX_ACCESS_AUDIT_ROWS),
      );
    } catch {
      return "failed";
    }
  };
  const byId = new Map<string, SummaryAudit>();
  const picks = new Map<string, PickedRows | null | "failed">();
  const idKeyed = wanted.filter((audit) => !NAME_MATCHED[audit.definition.id]);
  for (const audit of idKeyed)
    picks.set(audit.definition.id, await pickFrom(audit, { by: "personId", personId: userId }));
  const principal = await currentName(idKeyed, picks);

  // This account's current sign-in name, as far as stored data can tell, or null. A name is used
  // only when the stored sources agree and nothing says it now belongs to someone else, because a
  // name is reused when an account is renamed and the old name goes to a new hire: guessing would
  // show one person's mailbox, app usage and receipts on another's page.
  async function currentName(
    audits: readonly Audit[],
    picked: ReadonlyMap<string, PickedRows | null | "failed">,
  ): Promise<string | null> {
    // 1. The account's own rows in the newest runs.
    let newest = -1;
    let names: string[] = [];
    for (const audit of audits) {
      const found = picked.get(audit.definition.id);
      if (!found || found === "failed") continue;
      const own = principalsFromRows(found.rows);
      if (own.length === 0) continue;
      const at = found.run.finishedAt ? Date.parse(found.run.finishedAt) : 0;
      if (at > newest) {
        newest = at;
        names = own;
      } else if (at === newest) names = names.concat(own);
    }
    let candidate: string | null = null;
    if (names.length > 0) {
      // The rows of one run disagreeing about this account is a reason to say nothing.
      if (new Set(names.map(joinKey)).size > 1) return null;
      candidate = names[0] ?? null;
    } else if (storedName !== null && !(await nameTaken(storedName))) {
      // 2. Stored sign-ins, when no other account has signed in with the name since.
      candidate = storedName;
    }
    if (candidate === null) return null;
    // 3. Another account's row in a run that carries the same name settles it the other way.
    for (const audit of audits) {
      const rival = await pickFrom(audit, {
        by: "name",
        path: ["person", "principal"],
        principal: candidate,
      });
      if (
        rival &&
        rival !== "failed" &&
        rival.rows.some((row) => row.personId !== null && row.personId !== userId)
      )
        return null;
    }
    return candidate;
  }
  for (const audit of wanted) {
    const id = audit.definition.id;
    const base = { auditId: id, name: audit.definition.title };
    let picked = picks.get(id);
    if (NAME_MATCHED[id]) {
      const how = pickFor(id, userId, principal);
      if (how === "no_name") {
        // Without a name there is nothing to look for, but whether the audit ran still shows.
        const run = await pickFrom(audit, { by: "keys", keys: [] });
        byId.set(
          id,
          run === "failed"
            ? { ...base, ...failed("The stored result") }
            : run === null
              ? { status: "no_run", ...base }
              : {
                  status: "no_name",
                  ...base,
                  finishedAt: run.run.finishedAt,
                  complete: run.run.complete,
                },
        );
        continue;
      }
      picked = await pickFrom(audit, how.pick);
    }
    if (picked === "failed" || picked === undefined) {
      byId.set(id, { ...base, ...failed("The stored result") });
      continue;
    }
    if (picked === null) {
      byId.set(id, { status: "no_run", ...base });
      continue;
    }
    const run = {
      ...base,
      finishedAt: picked.run.finishedAt,
      complete: picked.run.complete,
      scope: picked.run.scope,
      matchedBy: NAME_MATCHED[id] ? ("sign-in name" as const) : ("account" as const),
    };
    byId.set(
      id,
      picked.rows.length === 0
        ? { status: "not_in_run", ...run }
        : { status: "ok", ...run, rows: accessRows(audit, picked.rows), truncated: picked.more },
    );
  }
  return {
    states: SUMMARY_AUDIT_IDS.flatMap((id) => {
      const state = byId.get(id);
      return state ? [state] : [];
    }),
    principal,
  };
}

async function combinedSection(
  read: Read,
  tenantId: string,
  userId: string,
): Promise<PersonSummary["combinedRisks"]> {
  try {
    const picked = await read((sql) =>
      pickRows(sql, tenantId, COMBINED_RISKS_AUDIT_ID, { by: "personId", personId: userId }, 5),
    );
    if (!picked) return { status: "no_run" };
    return combinedRisksFor({ ...picked.run, rows: picked.rows }, userId);
  } catch {
    return { status: "failed", message: "The stored Combined risks result could not be read." };
  }
}

async function findingsSection(
  read: Read,
  tenantId: string,
  userId: string,
  available: boolean,
): Promise<PersonSummary["signinFindings"]> {
  if (!available) return { status: "unavailable" };
  try {
    const found = await read((sql) =>
      listFindings(sql, tenantId, {
        module: SIGNINS_MODULE,
        personId: userId,
        limit: MAX_SUMMARY_FINDINGS,
      }),
    );
    // Open means not resolved, as the page header counts it.
    const open = found.filter((finding) => finding.state !== "resolved");
    let highest: Severity | null = null;
    for (const finding of open)
      if (highest === null || severityAtLeast(finding.severity, highest))
        highest = finding.severity;
    return { status: "ok", open: open.length, highest, more: found.length >= MAX_SUMMARY_FINDINGS };
  } catch {
    return failed("The stored sign-in findings");
  }
}

const cleaned = (value: unknown, fallback: string) =>
  typeof value === "string" && value !== "" ? cleanText(value.slice(0, 500)) : fallback;

// The groups and teams that list this person as an owner, from the hidden owner ids of the
// Groups and Teams audit.
async function ownedSection(
  read: Read,
  tenantId: string,
  userId: string,
): Promise<PersonSummary["owned"]> {
  try {
    const picked = await read((sql) =>
      pickRows(
        sql,
        tenantId,
        GROUPS_AUDIT,
        { by: "list", path: ["ownerIds"], value: userId },
        MAX_ACCESS_OWNED,
      ),
    );
    if (!picked) return { status: "no_run" };
    return {
      status: "ok",
      finishedAt: picked.run.finishedAt,
      complete: picked.run.complete,
      more: picked.more,
      entries: picked.rows.map((row) => ({
        id: row.key.slice(0, 200),
        name: cleaned(row.cells.group, "Unnamed group"),
        kind: cleaned(row.cells.kind, "Group"),
        status: cleaned(row.cells.status, ""),
        tags: row.tags.slice(0, 20),
      })),
    };
  } catch {
    return failed("The stored Groups and Teams result");
  }
}

// Direct grants only: the person's groups would have to be read from Microsoft, which this page
// does not do. The Access tab expands groups.
async function mailboxSection(
  read: Read,
  tenantId: string,
  userId: string,
): Promise<PersonSummary["sharedMailboxes"]> {
  try {
    const found = await read(async (sql) => {
      // The grants query casts the id to a uuid, so an id that is not one has no grants.
      if (!GUID.test(userId)) {
        const only = await currentImport(sql, tenantId);
        return only
          ? {
              stored: only,
              grants: [],
              grantsTruncated: false,
              holders: [],
              holdersTruncated: false,
            }
          : null;
      }
      return personMailboxRead(sql, tenantId, userId, [], {
        grants: MAX_PERSON_MAILBOX_GRANTS,
        holders: MAX_MAILBOX_HOLDERS,
      });
    });
    if (!found) return { status: "no_import" };
    const holders = new Set(
      found.holders.map((holder) => holder.trustee.id ?? `name:${holder.trustee.name}`),
    );
    return {
      status: "ok",
      exportedAt: found.stored.exportedAt,
      importedAt: found.stored.importedAt,
      includesUserMailboxes: found.stored.includeUserMailboxes,
      mailboxes: new Set(found.grants.map((grant) => grant.mailbox.id)).size,
      mailboxesMore: found.grantsTruncated,
      holders: found.holders.length > 0 ? holders.size : null,
      holdersMore: found.holdersTruncated,
    };
  } catch {
    return failed("The stored mailbox access");
  }
}

async function changesSection(
  read: Read,
  tenantId: string,
  userId: string,
  available: boolean,
): Promise<PersonSummary["changes"]> {
  if (!available) return { status: "unavailable" };
  try {
    const since = new Date(Date.now() - SUMMARY_CHANGE_DAYS * DAY_MS);
    const counted = await read((sql) => countPersonChanges(sql, tenantId, userId, since));
    return { status: "ok", days: SUMMARY_CHANGE_DAYS, ...counted };
  } catch {
    return failed("The stored directory changes");
  }
}

async function handoffSection(
  read: Read,
  tenantId: string,
  userId: string,
  principal: string | null,
): Promise<PersonSummary["handoff"]> {
  try {
    const found = await read((sql) => personReceiptItems(sql, tenantId, userId, principal, 1));
    const item = found.items[0];
    return {
      status: "ok",
      latest: item
        ? {
            state: item.status,
            action: cleanText(item.action.slice(0, 80)),
            tool: cleanText(item.tool.name.slice(0, 80)),
            performedAt: item.performedAt,
          }
        : null,
    };
  } catch {
    return failed("The stored receipts");
  }
}

export function registerPersonSummaryRoutes(app: FastifyInstance, context: RouteContext): void {
  const personParams = z.object({
    userId: personIdSchema.refine((id) => !/^\.+$/.test(id)),
  });

  app.get("/v1/people/:userId/summary", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session)
      return reply.status(401).send({
        error: "authentication_required",
        message: "Sign in with an authorized administrator account.",
      });
    const params = personParams.safeParse(request.params);
    if (!params.success)
      return reply.status(400).send({ error: "invalid_request", message: "Unknown account." });
    const { modules } = context;
    if (!modules)
      return reply.status(409).send({
        error: "modules_unavailable",
        message:
          "Where this person shows up is read from stored results, which need module storage.",
      });
    const { tenantId } = session;
    const userId = normalizeUserId(params.data.userId);
    const signins = Boolean(modules.registry.get(SIGNINS_MODULE));
    const changes = Boolean(modules.registry.get(DIRECTORY_CHANGES_MODULE));

    // The ledger entry comes before anything is read. It names the datasets, never the person.
    await context.store.appendAudit(tenantId, session.actorId, {
      type: "person.summary.viewed",
      datasets: [
        "auditResults",
        ...(signins ? ["signIns", "findings"] : []),
        "mailboxAccess",
        ...(changes ? ["directoryAudit"] : []),
        "receipts",
      ],
      occurredAt: new Date().toISOString(),
    });

    const read: Read = (work) => modules.database.tenant(tenantId, work);
    // The sign-in name, from stored sign-ins when there are any. The audits' own rows are the
    // fallback once they have been read.
    let known: string | null = null;
    if (signins)
      try {
        const names = await read((sql) => storedPrincipals(sql, tenantId, [userId]));
        known = names.get(userId) ?? null;
      } catch {
        // The audits that carry an account id do not need it, and the others say so.
      }
    const audits = await auditStates(read, tenantId, userId, known, async (name) => {
      try {
        return await read((sql) => principalTakenSince(sql, tenantId, userId, name));
      } catch {
        // If it cannot be told, the name is not used.
        return true;
      }
    });

    const summary: PersonSummary = {
      combinedRisks: await combinedSection(read, tenantId, userId),
      signinFindings: await findingsSection(read, tenantId, userId, signins),
      audits: audits.states,
      owned: await ownedSection(read, tenantId, userId),
      sharedMailboxes: await mailboxSection(read, tenantId, userId),
      changes: await changesSection(read, tenantId, userId, changes),
      handoff: await handoffSection(read, tenantId, userId, audits.principal),
    };
    return personSummarySchema.parse(summary);
  });
}

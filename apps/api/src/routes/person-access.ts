import {
  type Audit,
  builtInAudits,
  createPacer,
  grantText,
  latestSucceededRun,
  type Pacer,
  type Policy,
  policyContext,
  policyReach,
  policySchema,
  readPaced,
} from "@prism365/audits";
import {
  ACCESS_AUDIT_FIELDS,
  type AccessAuditResult,
  type AccessMembership,
  type AccessPolicy,
  MAX_ACCESS_AUDIT_ROWS,
  COMBINED_RISKS_AUDIT_ID,
  MAX_ACCESS_MEMBERSHIPS,
  type PersonAccess,
  type PolicyReason,
  personAccessSchema,
  personIdSchema,
} from "@prism365/contracts";
import type { GraphFeatureId, GraphRequestContext } from "@prism365/graph";
import { MODULE_ID as SIGNINS_MODULE, storedPrincipals } from "@prism365/module-signins-audit";
import type { FastifyInstance } from "fastify";
import { z } from "zod";
import {
  classifyMembership,
  combinedRisksFor,
  type DirectoryObject,
  directoryObjectSchema,
  ownedEntries,
  readFailure,
  rowsFor,
  sortMemberships,
} from "../person-access.ts";
import { RouteBoundaryError, type RouteContext } from "./context.ts";

// The person page's Access tab: everything Prism365 can see about one account's access, in one
// answer (docs/people.md). Groups, roles and ownership are read live from the directory with the
// administrator's own session, Conditional Access through the same paced reader the audit uses, and
// audit results from what was stored, so nothing here asks Microsoft more than the page needs.
// Every section stands alone: one that cannot be read says why and the rest still come back.

const usersContext = { feature: "directory.users.read", permission: "User.Read.All" } as const;
const MEMBERSHIP_SELECT =
  "id,displayName,groupTypes,mailEnabled,securityEnabled,membershipRule,roleTemplateId,description";
const OWNED_SELECT = "id,displayName,resourceProvisioningOptions";
const MEMBERSHIP_PAGE = 100;
const OWNED_PAGE = 100;
const MAX_PAGES = 20;
const MAX_POLICIES = 1_000;
// Pacing alone is 1.1 seconds a page, so a tenant with many pages needs a minute or more.
const POLICY_READ_MILLISECONDS = 120_000;

// The audits whose rows describe one account, in the order the tab shows them.
const AUDIT_IDS = [
  "directory.privileged-roles",
  "mfa.registration",
  "devices.managed",
  "mail.mailbox-storage",
  "accounts.passwords",
  "identity.risky-users",
  "accounts.activity",
  "accounts.inactive",
  "licenses.reclaim",
] as const;

const userSchema = z.object({
  userPrincipalName: z.string().max(320).nullish().catch(null),
  userType: z.string().max(40).nullish().catch(null),
});
const pageObjects = (items: readonly unknown[]): DirectoryObject[] =>
  items.flatMap((item) => {
    const parsed = directoryObjectSchema.safeParse(item);
    return parsed.success ? [parsed.data] : [];
  });

// Does not wait between requests: the directory endpoints have their own, generous limits.
const unpaced: Pacer = { beforeRequest: async (signal) => signal.throwIfAborted() };

type Scoped = ReturnType<RouteContext["sessionGraph"]>;
type NotGranted = { status: "not_granted"; pack: string; permission: string; enabled: boolean };

function policyState(state: string): AccessPolicy["state"] {
  switch (state.toLowerCase()) {
    case "enabled":
      return "on";
    case "enabledforreportingbutnotenforced":
      return "report-only";
    case "disabled":
      return "off";
    default:
      return "unknown";
  }
}

export function registerPersonAccessRoutes(app: FastifyInstance, context: RouteContext): void {
  const personParams = z.object({
    userId: personIdSchema.refine((id) => !/^\.+$/.test(id)),
  });

  // The session's Graph client for one pack, or the reason it cannot be had.
  function scopedFor(
    session: NonNullable<Awaited<ReturnType<RouteContext["sessionFor"]>>>,
    handle: string,
    request: GraphRequestContext,
  ): { scoped: Scoped } | NotGranted {
    const pack = request.feature as GraphFeatureId;
    const refusal = (enabled: boolean): NotGranted => ({
      status: "not_granted",
      pack,
      permission: request.permission,
      enabled,
    });
    try {
      const scoped = context.sessionGraph(session, handle, [pack]);
      return scoped.granted.includes(request.permission) ? { scoped } : refusal(true);
    } catch (error) {
      if (error instanceof RouteBoundaryError) return refusal(false);
      throw error;
    }
  }

  async function readMemberships(scoped: Scoped, userId: string, signal: AbortSignal) {
    const id = encodeURIComponent(userId);
    const options = { pacer: unpaced, signal, maxPages: MAX_PAGES };
    const all = await readPaced(
      scoped.graph,
      `users/${id}/transitiveMemberOf?$select=${MEMBERSHIP_SELECT}&$top=${MEMBERSHIP_PAGE}`,
      usersContext,
      { ...options, maxRecords: MAX_ACCESS_MEMBERSHIPS },
    );
    const direct = await readPaced(
      scoped.graph,
      `users/${id}/memberOf?$select=id&$top=${MEMBERSHIP_PAGE}`,
      usersContext,
      { ...options, maxRecords: MAX_ACCESS_MEMBERSHIPS },
    );
    const directIds = new Set(pageObjects(direct.items).map((object) => object.id));
    const seen = new Set<string>();
    const entries: AccessMembership[] = [];
    for (const object of pageObjects(all.items)) {
      if (seen.has(object.id)) continue;
      seen.add(object.id);
      entries.push(classifyMembership(object, directIds));
    }
    return { entries: sortMemberships(entries), truncated: !all.complete };
  }

  async function readOwned(scoped: Scoped, userId: string, signal: AbortSignal) {
    const read = await readPaced(
      scoped.graph,
      `users/${encodeURIComponent(userId)}/ownedObjects?$select=${OWNED_SELECT}&$top=${OWNED_PAGE}`,
      usersContext,
      { pacer: unpaced, signal, maxPages: MAX_PAGES, maxRecords: 1_000 },
    );
    const { entries, more } = ownedEntries(pageObjects(read.items));
    return { entries, truncated: !read.complete || more };
  }

  async function readPolicies(scoped: Scoped) {
    const signal = AbortSignal.timeout(POLICY_READ_MILLISECONDS);
    const read = await readPaced(
      scoped.graph,
      "identity/conditionalAccess/policies",
      policyContext,
      { pacer: createPacer(), signal, maxPages: MAX_PAGES, maxRecords: MAX_POLICIES },
    );
    const policies: Policy[] = [];
    const seen = new Set<string>();
    let unreadable = 0;
    for (const raw of read.items) {
      const parsed = policySchema.safeParse(raw);
      if (!parsed.success || seen.has(parsed.data.id)) {
        unreadable += 1;
        continue;
      }
      seen.add(parsed.data.id);
      policies.push(parsed.data);
    }
    return { policies, unreadable, truncated: !read.complete };
  }

  function reasonsWithNames(
    reasons: { kind: PolicyReason["kind"]; id: string | null }[],
    names: Map<string, string>,
  ): PolicyReason[] {
    return reasons.slice(0, 20).map(({ kind, id }) => ({
      kind,
      id,
      name: kind === "group" || kind === "role" ? (names.get(`${kind}:${id}`) ?? null) : null,
    }));
  }

  app.get("/v1/people/:userId/access", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    const handle = context.sessionHandle(request);
    if (!handle) return reply.status(401).send({ error: "authentication_required" });
    const params = personParams.safeParse(request.params);
    if (!params.success)
      return reply.status(400).send({ error: "invalid_request", message: "Unknown account." });
    const { userId } = params.data;

    // Which live sections can be read is decided first, so the ledger entry names what is about to
    // be read. Nothing has been asked of Microsoft or the store yet.
    const users = scopedFor(session, handle, usersContext);
    const policyAccess = scopedFor(session, handle, policyContext);
    const modules = context.modules;
    const datasets = [
      ...("scoped" in users ? ["directory"] : []),
      ...("scoped" in policyAccess ? ["conditionalAccess"] : []),
      ...(modules ? ["auditResults"] : []),
    ];
    await context.store.appendAudit(session.tenantId, session.actorId, {
      type: "person.access.viewed",
      datasets,
      sections: datasets.length,
      occurredAt: new Date().toISOString(),
    });
    const readAt = () => new Date().toISOString();
    const signal = AbortSignal.timeout(60_000);

    // Who the account is, in the directory: the sign-in name and whether it is a guest.
    let principal: string | null = null;
    let userType: string | null = null;
    if ("scoped" in users) {
      try {
        const user = await users.scoped.graph.get(
          `users/${encodeURIComponent(userId)}?$select=id,userPrincipalName,userType`,
          usersContext,
          userSchema,
        );
        principal = user.userPrincipalName?.trim() || null;
        userType = user.userType ?? null;
      } catch {
        // The memberships read below reports why; the name stays unknown.
      }
    }

    let memberships: PersonAccess["memberships"];
    if (!("scoped" in users)) memberships = users;
    else
      try {
        const read = await readMemberships(users.scoped, userId, signal);
        memberships = { status: "ok", readAt: readAt(), ...read };
      } catch (error) {
        memberships = { status: "failed", message: readFailure(error) };
      }

    let owned: PersonAccess["owned"];
    if (!("scoped" in users)) owned = users;
    else
      try {
        const read = await readOwned(users.scoped, userId, signal);
        owned = { status: "ok", readAt: readAt(), ...read };
      } catch (error) {
        owned = { status: "failed", message: readFailure(error) };
      }

    let conditionalAccess: PersonAccess["conditionalAccess"];
    if (!("scoped" in policyAccess)) conditionalAccess = policyAccess;
    else
      try {
        const read = await readPolicies(policyAccess.scoped);
        const held = memberships.status === "ok" ? memberships.entries : [];
        const names = new Map<string, string>();
        for (const entry of held) {
          if (entry.name === null) continue;
          if (entry.kind === "role" && entry.roleTemplateId)
            names.set(`role:${entry.roleTemplateId}`, entry.name);
          else if (entry.kind !== "administrative-unit") names.set(`group:${entry.id}`, entry.name);
        }
        const person = {
          id: userId,
          userType,
          groupIds: held.filter((entry) => entry.kind !== "role").map((entry) => entry.id),
          roleTemplateIds: held.flatMap((entry) =>
            entry.kind === "role" && entry.roleTemplateId ? [entry.roleTemplateId] : [],
          ),
        };
        const listed: AccessPolicy[] = [];
        let unaffected = 0;
        for (const policy of read.policies) {
          const reach = policyReach(policy, person);
          if (reach.effect === "none") {
            unaffected += 1;
            continue;
          }
          listed.push({
            id: policy.id,
            name: policy.displayName,
            state: policyState(policy.state),
            controls: grantText(policy),
            effect: reach.effect,
            targeted: reach.targeted,
            includedBy: reasonsWithNames(reach.included, names),
            excludedBy: reasonsWithNames(reach.excluded, names),
          });
        }
        listed.sort(
          (a, b) => (a.name ?? "").localeCompare(b.name ?? "") || a.id.localeCompare(b.id),
        );
        conditionalAccess = {
          status: "ok",
          readAt: readAt(),
          policies: listed,
          unaffected,
          unreadable: read.unreadable,
          truncated: read.truncated,
          membershipsKnown: memberships.status === "ok",
        };
      } catch (error) {
        conditionalAccess = { status: "failed", message: readFailure(error) };
      }

    let audits: PersonAccess["audits"];
    if (!modules)
      audits = {
        status: "failed",
        message:
          "Audit results are kept with module storage, which this installation does not have.",
      };
    else
      try {
        audits = await storedAudits(modules, session.tenantId, userId, principal);
      } catch {
        audits = { status: "failed", message: "The stored audit results could not be read." };
      }

    let combinedRisks: PersonAccess["combinedRisks"];
    if (!modules)
      combinedRisks = {
        status: "failed",
        message:
          "Audit results are kept with module storage, which this installation does not have.",
      };
    else
      try {
        const run = await modules.database.tenant(session.tenantId, (sql) =>
          latestSucceededRun(sql, session.tenantId, COMBINED_RISKS_AUDIT_ID),
        );
        combinedRisks = run ? combinedRisksFor(run, userId) : { status: "no_run" };
      } catch {
        combinedRisks = {
          status: "failed",
          message: "The stored Combined risks result could not be read.",
        };
      }

    return personAccessSchema.parse({
      combinedRisks,
      memberships,
      owned,
      conditionalAccess,
      audits,
    });
  });

  async function storedAudits(
    modules: NonNullable<RouteContext["modules"]>,
    tenantId: string,
    userId: string,
    directoryName: string | null,
  ): Promise<PersonAccess["audits"]> {
    const wanted = AUDIT_IDS.flatMap((id) => {
      const audit = builtInAudits.find((candidate) => candidate.definition.id === id);
      return audit ? [audit] : [];
    });
    return modules.database.tenant(tenantId, async (sql) => {
      // The mailbox report names mailboxes by sign-in name. Stored sign-ins give it without asking
      // Microsoft; the directory read above is the fallback.
      let principal = directoryName;
      if (principal === null && modules.registry.get(SIGNINS_MODULE)) {
        const stored = await storedPrincipals(sql, tenantId, [userId]);
        principal = stored.get(userId) ?? null;
      }
      const results: AccessAuditResult[] = [];
      for (const audit of wanted)
        results.push(await resultFor(sql, tenantId, audit, userId, principal));
      return { status: "ok" as const, results };
    });
  }

  async function resultFor(
    sql: Parameters<typeof latestSucceededRun>[0],
    tenantId: string,
    audit: Audit,
    userId: string,
    principal: string | null,
  ): Promise<AccessAuditResult> {
    const auditId = audit.definition.id;
    const name = audit.definition.title;
    const run = await latestSucceededRun(sql, tenantId, auditId);
    if (!run) return { status: "no_run", auditId, name };
    const found = rowsFor(auditId, run, userId, principal);
    if (found === "no_name")
      return {
        status: "failed",
        auditId,
        name,
        message:
          "Prism365 does not know this account's sign-in name yet, so its mailbox cannot be found in the report.",
      };
    // Only the columns this page shows (ACCESS_AUDIT_FIELDS), in that order.
    const shown = ACCESS_AUDIT_FIELDS[auditId] ?? [];
    const columns = shown.flatMap((key) => {
      const column = audit.definition.columns.find((candidate) => candidate.key === key);
      return column ? [column] : [];
    });
    const rows = found.rows.slice(0, MAX_ACCESS_AUDIT_ROWS).map((row) => ({
      key: row.key,
      cells: columns.flatMap((column) => {
        const value = row.cells[column.key];
        if (value === null || value === undefined) return [];
        if (Array.isArray(value) && value.length === 0) return [];
        return [{ key: column.key, label: column.label, kind: column.kind, value }];
      }),
    }));
    return {
      status: "ok",
      auditId,
      name,
      finishedAt: run.finishedAt,
      scope: run.scope,
      matchedBy: found.matchedBy,
      rows,
      truncated: found.rows.length > MAX_ACCESS_AUDIT_ROWS,
    };
  }
}

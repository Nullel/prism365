import {
  type CanaryAccount,
  canaryChangeSchema,
  canaryListSchema,
  cleanText,
  MAX_CANARY_ACCOUNTS,
} from "@prism365/contracts";
import { GraphRequestError } from "@prism365/graph";
import {
  MODULE_ID as SIGNINS_MODULE,
  readCanaries,
  storedPrincipals,
  writeCanaries,
} from "@prism365/module-signins-audit";
import type { FastifyInstance } from "fastify";
import { z } from "zod";
import { RouteBoundaryError, type RouteContext } from "./context.ts";

// Canary accounts (docs/modules/signins-audit.md): the decoy accounts an administrator marks, so
// that any sign-in attempt by one becomes a critical finding. The list is one setting per tenant;
// the collector reads it once per run. Every route checks the session itself, and a change is
// recorded in the access ledger before it is saved, with counts and never an id or a name.
export function registerCanaryRoutes(app: FastifyInstance, context: RouteContext): void {
  const unauthorized = {
    error: "authentication_required",
    message: "Sign in with an authorized administrator account.",
  };
  const unavailable = {
    error: "modules_unavailable",
    message: "Canary accounts need the sign-in audit module, which is not available here.",
  };
  const available = () => Boolean(context.modules?.registry.get(SIGNINS_MODULE));
  // Microsoft names every account by a GUID. Anything else is not looked up, so an id such as
  // ".." can never reach the path of a directory request.
  const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  const directoryContext = {
    feature: "directory.users.read",
    permission: "User.Read.All",
  } as const;
  // A name longer than the list stores is not shortened, because a cut name would match nobody.
  const directoryAccount = z.object({ userPrincipalName: z.string().min(1).max(320) });

  app.get("/v1/settings/canaries", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(unauthorized);
    const modules = context.modules;
    if (!modules || !available()) return reply.status(409).send(unavailable);
    const list = await modules.database.tenant(session.tenantId, (sql) =>
      readCanaries(sql, session.tenantId),
    );
    return canaryListSchema.parse(list);
  });

  // The name of one account in the directory: null when Microsoft has no such account, and
  // "unavailable" when it could not be asked or did not answer.
  async function directoryPrincipal(
    scoped: ReturnType<RouteContext["sessionGraph"]>,
    userId: string,
  ): Promise<string | null | "unavailable"> {
    if (!guid.test(userId)) return null;
    try {
      const account = await scoped.graph.get(
        `users/${encodeURIComponent(userId)}?$select=id,userPrincipalName`,
        directoryContext,
        directoryAccount,
      );
      return cleanText(account.userPrincipalName).trim() || null;
    } catch (error) {
      return error instanceof GraphRequestError && error.status === 404 ? null : "unavailable";
    }
  }

  app.post("/v1/settings/canaries", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(unauthorized);
    const modules = context.modules;
    if (!modules || !available()) return reply.status(409).send(unavailable);
    const body = canaryChangeSchema.safeParse(request.body);
    if (!body.success)
      return reply.status(400).send({
        error: "invalid_request",
        message: `Name up to ${MAX_CANARY_ACCOUNTS} accounts to add or remove, each by its user id, and not the same account in both.`,
      });
    const { tenantId, actorId } = session;
    const ledger = (type: string, fields: Record<string, unknown>) =>
      context.store.appendAudit(tenantId, actorId, {
        type,
        ...fields,
        occurredAt: new Date().toISOString(),
      });
    const tooMany = {
      error: "invalid_request",
      message: `A tenant can have at most ${MAX_CANARY_ACCOUNTS} canary accounts. Remove one first.`,
    };

    const current = await modules.database.tenant(tenantId, (sql) => readCanaries(sql, tenantId));
    const held = new Set(current.accounts.map((account) => account.userId));
    const toAdd = [...new Set(body.data.add ?? [])].filter((id) => !held.has(id));
    const toRemove = [...new Set(body.data.remove ?? [])].filter((id) => held.has(id));
    // Marking a marked account again, or unmarking one that is not, changes nothing.
    if (toAdd.length === 0 && toRemove.length === 0) return canaryListSchema.parse(current);
    if (held.size - toRemove.length + toAdd.length > MAX_CANARY_ACCOUNTS)
      return reply.status(400).send(tooMany);

    // Each new account needs the sign-in name it signs in with, so a record without a user id can
    // still be matched. Stored sign-ins come first; the directory is asked only for the rest.
    const principals = new Map<string, string>();
    if (toAdd.length > 0) {
      await ledger("signins.canaries.lookup", { datasets: ["signIns"], accounts: toAdd.length });
      const stored = await modules.database.tenant(tenantId, (sql) =>
        storedPrincipals(sql, tenantId, toAdd),
      );
      for (const [id, principal] of stored) principals.set(id, cleanText(principal));
      const missing = toAdd.filter((id) => !principals.has(id));
      if (missing.length > 0) {
        const handle = context.sessionHandle(request);
        if (!handle) return reply.status(401).send(unauthorized);
        let scoped: ReturnType<RouteContext["sessionGraph"]> | null = null;
        try {
          scoped = context.sessionGraph(session, handle, ["directory.users.read"]);
        } catch (error) {
          if (!(error instanceof RouteBoundaryError)) throw error;
        }
        if (!scoped?.granted.includes(directoryContext.permission))
          return reply.status(403).send({
            error: "missing_consent",
            message: `Marking an account that has no stored sign-ins needs the ${directoryContext.permission} feature pack.`,
          });
        await ledger("signins.canaries.lookup", {
          datasets: ["directory"],
          accounts: missing.length,
        });
        const answers: (string | null | "unavailable")[] = [];
        for (let start = 0; start < missing.length; start += 5)
          answers.push(
            ...(await Promise.all(
              missing.slice(start, start + 5).map((id) => directoryPrincipal(scoped, id)),
            )),
          );
        if (answers.includes("unavailable"))
          return reply.status(503).send({
            error: "directory_unavailable",
            message:
              "Microsoft's directory could not be read just now. Nothing was changed. Try again.",
          });
        if (answers.includes(null))
          return reply.status(404).send({
            error: "account_not_found",
            message:
              "One of those accounts is not in the sign-in history or the directory, so nothing was changed.",
          });
        for (const [index, id] of missing.entries()) principals.set(id, answers[index] as string);
      }
    }

    // The ledger entry comes first, so no change to the list is ever unrecorded.
    await ledger("signins.canaries.changed", {
      added: toAdd.length,
      removed: toRemove.length,
      total: held.size - toRemove.length + toAdd.length,
    });
    const now = new Date().toISOString();
    const saved = await modules.database.tenant(tenantId, async (sql) => {
      // Applied to the list as it is now, not as it was read above, so a change made in between
      // by another administrator is kept.
      const fresh = await readCanaries(sql, tenantId);
      const removing = new Set(toRemove);
      const kept = fresh.accounts.filter((account) => !removing.has(account.userId));
      const keptIds = new Set(kept.map((account) => account.userId));
      const added: CanaryAccount[] = toAdd
        .filter((id) => !keptIds.has(id))
        .map((userId) => ({
          userId,
          principal: principals.get(userId) as string,
          addedAt: now,
          addedBy: actorId,
        }));
      if (kept.length + added.length > MAX_CANARY_ACCOUNTS) return null;
      return writeCanaries(sql, tenantId, { accounts: [...kept, ...added] }, actorId);
    });
    if (!saved) return reply.status(400).send(tooMany);
    return canaryListSchema.parse(saved);
  });
}

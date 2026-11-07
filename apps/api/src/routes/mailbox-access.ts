import { randomUUID } from "node:crypto";
import { readPaced } from "@prism365/audits";
import {
  exportedAtProblem,
  MAX_MAILBOX_ACCESS_BYTES,
  MAX_MAILBOX_HOLDERS,
  MAX_PERSON_MAILBOX_GRANTS,
  mailboxAccessClearSchema,
  mailboxAccessExportSchema,
  mailboxAccessImportResultSchema,
  mailboxAccessProblem,
  mailboxAccessStatusSchema,
  type PersonMailboxAccess,
  personMailboxAccessSchema,
} from "@prism365/contracts";
import type { GraphFeatureId } from "@prism365/graph";
import type { FastifyInstance } from "fastify";
import { z } from "zod";
import {
  clearImport,
  currentImport,
  LATEST_IMPORT_INDEX,
  personMailboxRead,
  replaceImport,
} from "../mailbox-access-store.ts";
import { readFailure } from "../person-access.ts";
import { RouteBoundaryError, type RouteContext } from "./context.ts";

// Shared mailbox access from a read-only Exchange export (ADR-0012, docs/SETUP.md).
//
// Prism365 never calls Exchange and never holds an Exchange credential. An administrator runs
// Export-MailboxAccess.ps1 in their own session and imports the file here. The file is untrusted:
// it is validated whole before anything is recorded or stored, it must be from this tenant and
// recent, and what is kept is shown as the export's claim with its date. Nothing acts on it.

const unavailable = {
  error: "mailbox_access_unavailable",
  message: "Mailbox access needs live mode with module storage, and a signed-in administrator.",
};
const signIn = {
  error: "authentication_required",
  message: "Sign in with an authorized administrator account.",
};
const usersContext = { feature: "directory.users.read", permission: "User.Read.All" } as const;
const MEMBERSHIP_PAGE = 100;
const MAX_PAGES = 20;
const MAX_GROUPS = 1_000;
// The largest export is 200,000 grants in 1,000 statements, which takes far longer than the
// usual 60-second transaction on a busy disk.
const IMPORT_MILLISECONDS = 300_000;

const personParams = z.object({ userId: z.guid().transform((value) => value.toLowerCase()) });
const membershipSchema = z.object({ id: z.string().min(1).max(200) });

// Does not wait between requests: the directory endpoints have their own, generous limits.
const unpaced = { beforeRequest: async (signal: AbortSignal) => signal.throwIfAborted() };

export function registerMailboxAccessRoutes(app: FastifyInstance, context: RouteContext): void {
  const { store } = context;

  // The session's Graph client for the users pack, or why it cannot be had.
  function usersGraph(
    session: NonNullable<Awaited<ReturnType<RouteContext["sessionFor"]>>>,
    handle: string,
  ) {
    const refusal = (enabled: boolean) => ({
      status: "not_granted" as const,
      pack: usersContext.feature,
      permission: usersContext.permission,
      enabled,
    });
    try {
      const scoped = context.sessionGraph(session, handle, [
        usersContext.feature as GraphFeatureId,
      ]);
      return scoped.granted.includes(usersContext.permission) ? { scoped } : refusal(true);
    } catch (error) {
      if (error instanceof RouteBoundaryError) return refusal(false);
      throw error;
    }
  }

  app.post(
    "/v1/mailbox-access/import",
    {
      // An export may be 25 times the API's usual body limit, so only a signed-in administrator's
      // request is read that far. The global hook still checks the session and the CSRF token
      // once the body is in.
      bodyLimit: MAX_MAILBOX_ACCESS_BYTES,
      onRequest: async (request, reply) => {
        if (!(await context.sessionFor(request))) return reply.status(401).send(signIn);
      },
      // Problems reading the body, in words that fit an export file. The handler answers its own
      // validation problems; anything else is the server's fault and says nothing more.
      errorHandler: (error, _request, reply) => {
        const status =
          error instanceof Error && "statusCode" in error && typeof error.statusCode === "number"
            ? error.statusCode
            : 500;
        if (status === 413)
          return reply.status(413).send({
            error: "request_too_large",
            message:
              "An export can be at most 25 MB. Export shared mailboxes only (without -IncludeUserMailboxes), or ask about a larger limit.",
          });
        if (status === 415)
          return reply
            .status(415)
            .send({ error: "unsupported_media_type", message: "Send the export as JSON." });
        if (status >= 400 && status < 500)
          return reply
            .status(400)
            .send({ error: "invalid_export", message: "The export is not valid JSON." });
        return reply.status(500).send({
          error: "internal_error",
          message:
            "The export could not be imported. The previous import, if any, is unchanged. Try again.",
        });
      },
    },
    async (request, reply) => {
      const session = await context.sessionFor(request);
      if (!session) return reply.status(401).send(signIn);
      const { modules } = context;
      if (!modules) return reply.status(409).send(unavailable);
      const parsed = mailboxAccessExportSchema.safeParse(request.body);
      if (!parsed.success)
        return reply.status(400).send({
          error: "invalid_export",
          message: mailboxAccessProblem(request.body, parsed.error),
        });
      const file = parsed.data;
      const tenantId = session.tenantId.toLowerCase();
      // An export of another tenant would show that tenant's access as this one's.
      if (file.tenantId !== tenantId)
        return reply.status(400).send({
          error: "wrong_tenant",
          message: `This export is from tenant ${file.tenantId}, but this Prism365 reads tenant ${tenantId}. Connect Exchange Online to the right tenant and export again.`,
        });
      const dated = exportedAtProblem(file.exportedAt, Date.now());
      if (dated) return reply.status(400).send({ error: "export_out_of_range", message: dated });
      const id = randomUUID();
      // The ledger entry comes before anything is stored or deleted, so no import or replacement
      // is ever unrecorded. Ids and counts only: nothing the file names.
      await store.appendAudit(session.tenantId, session.actorId, {
        type: "mailbox_access.imported",
        importId: id,
        mailboxes: file.mailboxes.length,
        grants: file.grants.length,
        unresolved: file.unresolved,
        occurredAt: new Date().toISOString(),
      });
      // The old import goes and the new one arrives in one transaction: a grant that fails to
      // store keeps the old import as it was.
      let outcome: Awaited<ReturnType<typeof replaceImport>>;
      try {
        outcome = await modules.database.tenant(
          session.tenantId,
          (sql) =>
            replaceImport(sql, session.tenantId, file, {
              id,
              importedBy: { id: session.actorId, name: session.displayName ?? null },
            }),
          { timeoutMilliseconds: IMPORT_MILLISECONDS },
        );
      } catch (error) {
        // Another import for this tenant committed first. Only that index means so.
        const failure = error as { code?: unknown; constraint?: unknown } | null;
        if (failure?.code === "23505" && failure.constraint === LATEST_IMPORT_INDEX)
          return reply.status(409).send({
            error: "import_conflict",
            message:
              "Another export was imported at the same moment and is kept. Import this file again to replace it.",
          });
        throw error;
      }
      return reply.status(201).send(
        mailboxAccessImportResultSchema.parse({
          import: outcome.summary,
          replaced: outcome.replaced,
        }),
      );
    },
  );

  // The stored import's dates and counts, or none. No mailbox or person appears here.
  app.get("/v1/mailbox-access", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(signIn);
    const { modules } = context;
    if (!modules) return reply.status(409).send(unavailable);
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "mailbox_access.viewed",
      occurredAt: new Date().toISOString(),
    });
    const stored = await modules.database.tenant(session.tenantId, (sql) =>
      currentImport(sql, session.tenantId),
    );
    return mailboxAccessStatusSchema.parse({ import: stored });
  });

  app.delete("/v1/mailbox-access", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(signIn);
    const { modules } = context;
    if (!modules) return reply.status(409).send(unavailable);
    // Recorded before anything is deleted.
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "mailbox_access.cleared",
      occurredAt: new Date().toISOString(),
    });
    const cleared = await modules.database.tenant(session.tenantId, (sql) =>
      clearImport(sql, session.tenantId),
    );
    return mailboxAccessClearSchema.parse({ cleared });
  });

  // The shared mailboxes one person can use according to the import, and when the person is a
  // mailbox in it, who can use theirs. Grants to groups are matched through the person's groups,
  // nested ones included, read live with the administrator's own session.
  app.get("/v1/people/:userId/mailbox-access", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(signIn);
    const handle = context.sessionHandle(request);
    if (!handle) return reply.status(401).send(signIn);
    const params = personParams.safeParse(request.params);
    if (!params.success)
      return reply.status(400).send({ error: "invalid_request", message: "Unknown account." });
    const { modules } = context;
    if (!modules) return reply.status(409).send(unavailable);
    const { userId } = params.data;
    const tenantId = session.tenantId;

    // Whether the groups can be read is decided first, so the ledger entry names what is about to
    // be read. Nothing has been asked of Microsoft or the store yet.
    const users = usersGraph(session, handle);
    await store.appendAudit(tenantId, session.actorId, {
      type: "person.mailbox_access.viewed",
      datasets: ["mailboxAccess", ...("scoped" in users ? ["directory"] : [])],
      occurredAt: new Date().toISOString(),
    });

    const nothing = personMailboxAccessSchema.parse({
      import: null,
      groups: { status: "not_read" },
      grants: [],
      grantsTruncated: false,
      holders: [],
      holdersTruncated: false,
    } satisfies PersonMailboxAccess);
    // Without an import there is nothing to match the groups against, so Microsoft is not asked.
    const any = await modules.database.tenant(tenantId, (sql) => currentImport(sql, tenantId));
    if (!any) return nothing;

    let groups: PersonMailboxAccess["groups"];
    let groupIds: string[] = [];
    if (!("scoped" in users)) groups = users;
    else
      try {
        const read = await readPaced(
          users.scoped.graph,
          `users/${encodeURIComponent(userId)}/transitiveMemberOf?$select=id&$top=${MEMBERSHIP_PAGE}`,
          usersContext,
          {
            pacer: unpaced,
            signal: AbortSignal.timeout(60_000),
            maxPages: MAX_PAGES,
            maxRecords: MAX_GROUPS,
          },
        );
        // Only GUIDs can match a trustee id; anything else Microsoft returned is left out.
        const ids = new Set<string>();
        for (const item of read.items) {
          const parsed = membershipSchema.safeParse(item);
          if (parsed.success && z.guid().safeParse(parsed.data.id).success)
            ids.add(parsed.data.id.toLowerCase());
        }
        groupIds = [...ids];
        groups = {
          status: "ok",
          readAt: new Date().toISOString(),
          count: groupIds.length,
          truncated: !read.complete,
        };
      } catch (error) {
        groups = { status: "failed", message: readFailure(error) };
      }

    // The import is read again after Microsoft has answered, together with its grants and holders
    // in one statement: another administrator may have replaced or cleared it meanwhile, and the
    // page must show one export's date with that same export's grants.
    const found = await modules.database.tenant(tenantId, (sql) =>
      personMailboxRead(sql, tenantId, userId, groupIds, {
        grants: MAX_PERSON_MAILBOX_GRANTS,
        holders: MAX_MAILBOX_HOLDERS,
      }),
    );
    if (!found) return nothing;
    return personMailboxAccessSchema.parse({
      import: found.stored,
      groups,
      grants: found.grants,
      grantsTruncated: found.grantsTruncated,
      holders: found.holders,
      holdersTruncated: found.holdersTruncated,
    } satisfies PersonMailboxAccess);
  });
}

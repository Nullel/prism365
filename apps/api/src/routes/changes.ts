import {
  changeFeedQuerySchema,
  changeFeedSchema,
  personChangePageSchema,
  personChangeQuerySchema,
  personIdSchema,
} from "@prism365/contracts";
import {
  DIRECTORY_CHANGES_MODULE,
  listChanges,
  parsePosition,
  personChanges,
} from "@prism365/module-directory-changes";
import type { FastifyInstance } from "fastify";
import { z } from "zod";
import type { RouteContext } from "./context.ts";

// The directory change feed (docs/modules/directory-changes.md): stored directory audit events,
// filtered and paged, and the changes involving one person. Both read only what collection
// stored; neither asks Microsoft anything. Each writes its access-ledger entry before reading,
// naming the dataset and never a person, a filter's text or a value.
export function registerChangeRoutes(app: FastifyInstance, context: RouteContext): void {
  const unavailable = {
    error: "modules_unavailable",
    message: "The directory change feed is not available in this installation.",
  };
  const available = () => Boolean(context.modules?.registry.get(DIRECTORY_CHANGES_MODULE));
  const personParams = z.object({ userId: personIdSchema });
  // A page position is checked before the ledger entry, so a stale link is refused without
  // recording a read that never happened.
  const position = z
    .string()
    .optional()
    .refine((value) => {
      if (value === undefined) return true;
      try {
        parsePosition(value);
        return true;
      } catch {
        return false;
      }
    });

  app.get("/v1/changes", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session)
      return reply.status(401).send({
        error: "authentication_required",
        message: "Sign in with an authorized administrator account.",
      });
    const modules = context.modules;
    if (!modules || !available()) return reply.status(409).send(unavailable);
    const query = changeFeedQuerySchema
      .extend({ before: changeFeedQuerySchema.shape.before.pipe(position) })
      .safeParse(request.query);
    if (!query.success)
      return reply.status(400).send({
        error: "invalid_request",
        message:
          "Check the filters: 1 to 3650 days, a result of success or failure, and a search of up to 200 characters. Reload the page if this keeps happening.",
      });
    const filter = query.data;
    await context.store.appendAudit(session.tenantId, session.actorId, {
      type: "changes.viewed",
      datasets: ["directoryAudit"],
      days: filter.days,
      occurredAt: new Date().toISOString(),
    });
    const page = await modules.database.tenant(session.tenantId, (sql) =>
      listChanges(sql, session.tenantId, {
        now: new Date(),
        days: filter.days,
        category: filter.category ?? null,
        result: filter.result ?? null,
        search: filter.search ?? null,
        before: filter.before ?? null,
        limit: filter.limit,
      }),
    );
    return changeFeedSchema.parse(page);
  });

  app.get("/v1/people/:userId/changes", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session)
      return reply.status(401).send({
        error: "authentication_required",
        message: "Sign in with an authorized administrator account.",
      });
    const modules = context.modules;
    if (!modules || !available()) return reply.status(409).send(unavailable);
    const params = personParams.safeParse(request.params);
    const query = personChangeQuerySchema
      .extend({ before: personChangeQuerySchema.shape.before.pipe(position) })
      .safeParse(request.query);
    if (!params.success || !query.success)
      return reply
        .status(400)
        .send({ error: "invalid_request", message: "Check the account and choose 1 to 200." });
    await context.store.appendAudit(session.tenantId, session.actorId, {
      type: "person.changes.viewed",
      datasets: ["directoryAudit"],
      occurredAt: new Date().toISOString(),
    });
    const page = await modules.database.tenant(session.tenantId, (sql) =>
      personChanges(sql, session.tenantId, params.data.userId, {
        before: query.data.before ?? null,
        limit: query.data.limit,
      }),
    );
    return personChangePageSchema.parse(page);
  });
}

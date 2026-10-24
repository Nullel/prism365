import {
  alertExportChangeSchema,
  alertOverviewSchema,
  alertRuleInputSchema,
  alertRuleListSchema,
  alertRuleUpdateSchema,
  alertWeeklyChangeSchema,
  deliveryRecordSchema,
  destinationCreatedSchema,
  destinationCreateSchema,
} from "@prism365/contracts";
import type { FastifyInstance, FastifyReply } from "fastify";
import { z } from "zod";
import { DeliveryInputError } from "../delivery.ts";
import type { RouteContext } from "./context.ts";

// Alerts (docs/alerts.md): destinations, rules and the delivery log. Every route checks the
// session itself; the POST and DELETE routes are also CSRF-checked by the global hook in app.ts.
// Each change, and each test send, is written to the access ledger before it is made, with ids
// and counts only: never a webhook address, a recipient or a rule's name.
//
// A webhook's address is accepted here once and never returned: the answer names its host only,
// and its signing secret is in the answer to the request that created it and nowhere else.

const unauthorized = {
  error: "authentication_required",
  message: "Sign in with an authorized administrator account.",
};
const unavailable = {
  error: "alerts_unavailable",
  message: "Alerts need live mode with module storage, and a signed-in administrator.",
};
const idParams = z.object({ id: z.uuid() });

export function registerDeliveryRoutes(app: FastifyInstance, context: RouteContext): void {
  const refuse = (reply: FastifyReply, error: unknown) => {
    if (error instanceof DeliveryInputError)
      return reply.status(error.status).send({ error: error.code, message: error.message });
    throw error;
  };

  app.get("/v1/alerts", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(unauthorized);
    const delivery = context.delivery;
    if (!delivery) return reply.status(409).send(unavailable);
    return alertOverviewSchema.parse(await delivery.overview());
  });

  app.post("/v1/alerts/destinations", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(unauthorized);
    const delivery = context.delivery;
    if (!delivery) return reply.status(409).send(unavailable);
    const body = destinationCreateSchema.safeParse(request.body);
    if (!body.success)
      return reply.status(400).send({
        error: "invalid_request",
        message:
          "Give the destination a name, and either up to 10 email addresses or a webhook address.",
      });
    const { tenantId, actorId } = session;
    try {
      const created = await delivery.createDestination(actorId, body.data, (destinationId) =>
        context.store.appendAudit(tenantId, actorId, {
          type: "alerts.destination.created",
          destinationId,
          kind: body.data.kind,
          recipients: body.data.kind === "email" ? body.data.recipients.length : 0,
          includeDetails: body.data.includeDetails,
          occurredAt: new Date().toISOString(),
        }),
      );
      return reply.status(201).send(destinationCreatedSchema.parse(created));
    } catch (error) {
      return refuse(reply, error);
    }
  });

  app.delete("/v1/alerts/destinations/:id", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(unauthorized);
    const delivery = context.delivery;
    if (!delivery) return reply.status(409).send(unavailable);
    const params = idParams.safeParse(request.params);
    if (!params.success)
      return reply.status(400).send({ error: "invalid_request", message: "Check the address." });
    const { tenantId, actorId } = session;
    try {
      const deleted = await delivery.deleteDestination(params.data.id, () =>
        context.store.appendAudit(tenantId, actorId, {
          type: "alerts.destination.deleted",
          destinationId: params.data.id,
          occurredAt: new Date().toISOString(),
        }),
      );
      if (!deleted)
        return reply
          .status(404)
          .send({ error: "unknown_destination", message: "That destination does not exist." });
      return { deleted: true };
    } catch (error) {
      return refuse(reply, error);
    }
  });

  // The delivery service writes the ledger entry (source "test") before anything is sent.
  app.post("/v1/alerts/destinations/:id/test", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(unauthorized);
    const delivery = context.delivery;
    if (!delivery) return reply.status(409).send(unavailable);
    const params = idParams.safeParse(request.params);
    if (!params.success)
      return reply.status(400).send({ error: "invalid_request", message: "Check the address." });
    const record = await delivery.sendTest(session.actorId, params.data.id);
    if (!record)
      return reply
        .status(404)
        .send({ error: "unknown_destination", message: "That destination does not exist." });
    return { delivery: deliveryRecordSchema.parse(record) };
  });

  app.post("/v1/alerts/rules", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(unauthorized);
    const delivery = context.delivery;
    if (!delivery) return reply.status(409).send(unavailable);
    const body = alertRuleInputSchema.safeParse(request.body);
    if (!body.success)
      return reply.status(400).send({
        error: "invalid_request",
        message:
          "Check the rule: a name, what it watches, one to five destinations and a quiet period of 5 minutes to 7 days.",
      });
    const { tenantId, actorId } = session;
    try {
      const rules = await delivery.createRule(actorId, body.data, (ruleId) =>
        context.store.appendAudit(tenantId, actorId, {
          type: "alerts.rule.created",
          ruleId,
          kind: body.data.kind,
          destinations: body.data.destinations.length,
          occurredAt: new Date().toISOString(),
        }),
      );
      return reply.status(201).send(alertRuleListSchema.parse({ rules }));
    } catch (error) {
      return refuse(reply, error);
    }
  });

  app.post("/v1/alerts/rules/:id", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(unauthorized);
    const delivery = context.delivery;
    if (!delivery) return reply.status(409).send(unavailable);
    const params = idParams.safeParse(request.params);
    const body = alertRuleUpdateSchema.safeParse(request.body);
    if (!params.success || !body.success)
      return reply
        .status(400)
        .send({ error: "invalid_request", message: "Say whether the rule is on or off." });
    const { tenantId, actorId } = session;
    const rules = await delivery.changeRule(actorId, params.data.id, body.data, () =>
      context.store.appendAudit(tenantId, actorId, {
        type: "alerts.rule.changed",
        ruleId: params.data.id,
        enabled: body.data.enabled,
        occurredAt: new Date().toISOString(),
      }),
    );
    if (!rules)
      return reply
        .status(404)
        .send({ error: "unknown_rule", message: "That rule does not exist." });
    return alertRuleListSchema.parse({ rules });
  });

  app.delete("/v1/alerts/rules/:id", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(unauthorized);
    const delivery = context.delivery;
    if (!delivery) return reply.status(409).send(unavailable);
    const params = idParams.safeParse(request.params);
    if (!params.success)
      return reply.status(400).send({ error: "invalid_request", message: "Check the address." });
    const { tenantId, actorId } = session;
    const rules = await delivery.changeRule(actorId, params.data.id, "delete", () =>
      context.store.appendAudit(tenantId, actorId, {
        type: "alerts.rule.deleted",
        ruleId: params.data.id,
        occurredAt: new Date().toISOString(),
      }),
    );
    if (!rules)
      return reply
        .status(404)
        .send({ error: "unknown_rule", message: "That rule does not exist." });
    return alertRuleListSchema.parse({ rules });
  });

  // Where an audit's scheduled runs are sent. The run's CSV goes only to email with details on.
  app.post("/v1/alerts/exports", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(unauthorized);
    const delivery = context.delivery;
    if (!delivery) return reply.status(409).send(unavailable);
    const body = alertExportChangeSchema.safeParse(request.body);
    if (!body.success)
      return reply.status(400).send({
        error: "invalid_request",
        message: "Name the audit, and a destination or null to stop sending it.",
      });
    const { tenantId, actorId } = session;
    try {
      await delivery.setExport(actorId, body.data.auditId, body.data.destinationId, () =>
        context.store.appendAudit(tenantId, actorId, {
          type: "alerts.export.changed",
          auditId: body.data.auditId,
          destinationId: body.data.destinationId,
          occurredAt: new Date().toISOString(),
        }),
      );
      return alertOverviewSchema.parse(await delivery.overview());
    } catch (error) {
      return refuse(reply, error);
    }
  });

  // Where the weekly This week summary goes.
  app.post("/v1/alerts/weekly", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(unauthorized);
    const delivery = context.delivery;
    if (!delivery) return reply.status(409).send(unavailable);
    const body = alertWeeklyChangeSchema.safeParse(request.body);
    if (!body.success)
      return reply
        .status(400)
        .send({ error: "invalid_request", message: "Choose a destination, or null for none." });
    const { tenantId, actorId } = session;
    try {
      await delivery.setWeekly(actorId, body.data.destinationId, () =>
        context.store.appendAudit(tenantId, actorId, {
          type: "alerts.weekly.changed",
          destinationId: body.data.destinationId,
          occurredAt: new Date().toISOString(),
        }),
      );
      return alertOverviewSchema.parse(await delivery.overview());
    } catch (error) {
      return refuse(reply, error);
    }
  });
}

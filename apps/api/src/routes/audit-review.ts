import { auditReadScope, auditDefinition, builtInAudits } from "@prism365/audits";
import {
  auditReviewResponseSchema,
  auditReviewScheduleSchema,
  auditReviewScheduleSettingsSchema,
  auditReviewStartSchema,
} from "@prism365/contracts";
import type { FastifyInstance } from "fastify";
import {
  readReviewSchedule,
  ReviewRunningError,
  type StoredReviewSchedule,
  writeReviewSchedule,
} from "../audit-review.ts";
import { auditUnavailableMessage } from "../audits.ts";
import type { RouteContext } from "./context.ts";

// The full review and its weekly schedule (docs/audits.md, section 22). A review runs the audits
// the administrator could run one by one, through the same runner, so each run is an ordinary run
// with its own access-ledger entry. Starting and stopping a review are ledgered before they act.

const auditsUnavailable = {
  error: "audits_unavailable",
  message: "Audits need live mode with module storage, and a signed-in administrator.",
};

export function registerAuditReviewRoutes(app: FastifyInstance, context: RouteContext): void {
  const { store } = context;

  // The review in progress, or the last one that ended. Counts and states only, so reading it is
  // not an access-ledger event.
  app.get("/v1/audits/review", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    const { reviewer, modules } = context;
    if (!reviewer || !modules) return reply.status(409).send(auditsUnavailable);
    return auditReviewResponseSchema.parse({ review: await reviewer.last(session.tenantId) });
  });

  app.post("/v1/audits/review", async (request, reply) => {
    const session = await context.sessionFor(request);
    const handle = context.sessionHandle(request);
    if (!session || !handle) return reply.status(401).send({ error: "authentication_required" });
    const { reviewer, modules } = context;
    if (!reviewer || !modules) return reply.status(409).send(auditsUnavailable);
    // Combined risks reads no permission of its own, so it is always available; a review with
    // only that to run would join old results and read nothing new.
    const runnable = builtInAudits.filter(
      (audit) =>
        !audit.usesStoredResults && context.auditAvailability(audit, session).state === "available",
    );
    if (runnable.length === 0)
      return reply.status(403).send({
        error: "nothing_to_run",
        message:
          "None of the audits can run with your access. Grant access on the Microsoft access page and try again.",
      });
    try {
      const started = await reviewer.start({
        tenantId: session.tenantId,
        actor: { id: session.actorId, name: session.displayName ?? null },
        trigger: "manual",
        unavailable: (audit) => {
          const availability = context.auditAvailability(audit, session);
          return availability.state === "available" ? null : auditUnavailableMessage(availability);
        },
        reading: async (audit) => {
          const { client, scopes } = context.auditClient(session, handle);
          const read = auditReadScope(audit, context.settings.enabledFeatures, scopes);
          const definition = auditDefinition(audit);
          // The ledger names what this run reads, as a single run's entry does.
          const datasets = [
            ...definition.permissions,
            ...definition.optionalPermissions.filter((permission) =>
              read.granted.includes(permission),
            ),
          ];
          return {
            client,
            packs: read.packs,
            granted: read.granted,
            beforeRead: (run) =>
              store.appendAudit(session.tenantId, session.actorId, {
                type: "audit.run.started",
                auditId: audit.definition.id,
                runId: run.id,
                datasets,
                review: true,
                occurredAt: new Date().toISOString(),
              }),
          };
        },
        ledger: (event) => store.appendAudit(session.tenantId, session.actorId, event),
      });
      return reply.status(202).send(auditReviewStartSchema.parse({ review: started.review }));
    } catch (error) {
      if (error instanceof ReviewRunningError)
        return reply.status(409).send({
          error: "review_running",
          message: "A full review is already running. Wait for it to finish or stop it first.",
        });
      throw error;
    }
  });

  // The ledger entry comes first. The audit reading now is stopped the way a scheduled run is
  // when collection ends; the audits after it are skipped.
  app.post("/v1/audits/review/stop", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    const { reviewer, modules } = context;
    if (!reviewer || !modules) return reply.status(409).send(auditsUnavailable);
    if (!reviewer.running(session.tenantId))
      return reply
        .status(409)
        .send({ error: "no_review_running", message: "No full review is running." });
    const current = await reviewer.last(session.tenantId);
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "audit.review.stopped",
      reviewId: current?.id ?? null,
      occurredAt: new Date().toISOString(),
    });
    reviewer.stop(session.tenantId);
    return reply
      .status(202)
      .send(auditReviewResponseSchema.parse({ review: await reviewer.last(session.tenantId) }));
  });

  const settingsOf = (stored: StoredReviewSchedule) =>
    auditReviewScheduleSettingsSchema.parse({
      frequency: stored.frequency,
      weekday: stored.weekday,
      hour: stored.hour,
      timeZone: context.settings.timeZone ?? "UTC",
      collecting: context.collectingAdmin() !== null,
      by: stored.frequency === "off" ? null : stored.name,
    });

  app.get("/v1/settings/audit-review", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    const { reviewer, modules } = context;
    if (!reviewer || !modules) return reply.status(409).send(auditsUnavailable);
    return settingsOf(
      await modules.database.tenant(session.tenantId, (sql) =>
        readReviewSchedule(sql, session.tenantId),
      ),
    );
  });

  // Saves the weekly day and hour. The review reads, when it runs on its own, with no more than
  // the permissions the saving administrator's own sign-in had granted for each audit, so you
  // cannot schedule a review that reads more than you could run yourself. The ledger entry is
  // written before the setting is saved.
  app.post("/v1/settings/audit-review", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    const { reviewer, modules } = context;
    if (!reviewer || !modules) return reply.status(409).send(auditsUnavailable);
    const body = auditReviewScheduleSchema.safeParse(request.body);
    if (!body.success)
      return reply.status(400).send({
        error: "invalid_request",
        message: "Choose off or weekly, a day of the week and an hour from 0 to 23.",
      });
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "audits.review.schedule.changed",
      frequency: body.data.frequency,
      weekday: body.data.weekday,
      hour: body.data.hour,
      occurredAt: new Date().toISOString(),
    });
    const grants = Object.fromEntries(
      builtInAudits.map((audit) => [
        audit.definition.id,
        auditReadScope(audit, context.settings.enabledFeatures, session.scopes).granted,
      ]),
    );
    const saved = await modules.database.tenant(session.tenantId, (sql) =>
      writeReviewSchedule(sql, session.tenantId, body.data, session.actorId, {
        name: session.displayName ?? null,
        grants,
        now: new Date(),
      }),
    );
    return settingsOf(saved);
  });
}

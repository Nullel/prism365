import {
  buildAuditTrend,
  builtInAudits,
  findAudit,
  knownAuditSchedule,
  readAuditSchedule,
  readTrendPoints,
  scheduledAuditCount,
  writeAuditSchedule,
} from "@prism365/audits";
import {
  auditIdSchema,
  auditScheduleSchema,
  auditScheduleSettingsSchema,
  auditTrendQuerySchema,
  MAX_AUDIT_TREND_DAYS,
} from "@prism365/contracts";
import type { FastifyInstance } from "fastify";
import { z } from "zod";
import type { RouteContext } from "./context.ts";

// Audit trends and schedules (docs/audits.md, sections 10 and 11). Trends are read from stored
// points, never from Microsoft. Schedules are saved here and run by the audit scheduler
// (audit-schedule.ts) while collection is on.

const DAY_MILLISECONDS = 86_400_000;

const auditsUnavailable = {
  error: "audits_unavailable",
  message: "Audits need live mode with module storage, and a signed-in administrator.",
};
const unknownAudit = { error: "unknown_audit", message: "There is no such audit." };
const auditParams = z.object({ auditId: auditIdSchema });

export function registerAuditInsightRoutes(app: FastifyInstance, context: RouteContext): void {
  const { store } = context;

  // Each audit's tiles over time. The points hold counts only, like the run history without its
  // rows, so reading them is not an access-ledger event.
  app.get("/v1/audits/:auditId/trend", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    const { audits, modules } = context;
    if (!audits || !modules) return reply.status(409).send(auditsUnavailable);
    const params = auditParams.safeParse(request.params);
    const audit = params.success ? findAudit(params.data.auditId) : undefined;
    if (!audit) return reply.status(404).send(unknownAudit);
    const query = auditTrendQuerySchema.safeParse(request.query ?? {});
    if (!query.success)
      return reply.status(400).send({
        error: "invalid_request",
        message: `Ask for between 1 and ${MAX_AUDIT_TREND_DAYS.toLocaleString("en-US")} days.`,
      });
    const { days } = query.data;
    const since = new Date(Date.now() - days * DAY_MILLISECONDS);
    const { points, truncated } = await modules.database.tenant(session.tenantId, (sql) =>
      readTrendPoints(sql, session.tenantId, audit.definition.id, since),
    );
    return buildAuditTrend({ auditId: audit.definition.id, days, points, truncated });
  });

  app.get("/v1/settings/audit-schedule", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    const { audits, modules } = context;
    if (!audits || !modules) return reply.status(409).send(auditsUnavailable);
    const schedule = await modules.database.tenant(session.tenantId, (sql) =>
      readAuditSchedule(sql, session.tenantId, builtInAudits),
    );
    return auditScheduleSettingsSchema.parse({
      audits: schedule,
      collecting: context.collectingAdmin() !== null,
    });
  });

  // Saves the whole schedule: an audit the body leaves out is off, an id that is not a known
  // audit is dropped. The CSRF check is the global hook's.
  app.post("/v1/settings/audit-schedule", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    const { audits, modules } = context;
    if (!audits || !modules) return reply.status(409).send(auditsUnavailable);
    const body = auditScheduleSchema.safeParse(request.body);
    if (!body.success)
      return reply.status(400).send({
        error: "invalid_request",
        message: "Choose off, daily or weekly for each audit.",
      });
    // The ledger entry comes first, so no schedule change is ever unrecorded. It holds how many
    // audits run on their own, never which.
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "audits.schedule.changed",
      scheduled: scheduledAuditCount(knownAuditSchedule(body.data, builtInAudits)),
      occurredAt: new Date().toISOString(),
    });
    const saved = await modules.database.tenant(session.tenantId, (sql) =>
      writeAuditSchedule(sql, session.tenantId, body.data, session.actorId, builtInAudits),
    );
    return auditScheduleSettingsSchema.parse({
      audits: saved,
      collecting: context.collectingAdmin() !== null,
    });
  });
}

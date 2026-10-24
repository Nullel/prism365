import {
  auditReadScope,
  auditScheduleChanges,
  buildAuditTrend,
  builtInAudits,
  findAudit,
  knownAuditSchedule,
  readAuditSchedule,
  readStoredAuditSchedule,
  readTrendPoints,
  type StoredAuditSchedule,
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
import { scheduleLimits } from "../audit-schedule.ts";
import { auditUnavailableMessage } from "../audits.ts";
import type { RouteContext } from "./context.ts";

// Audit trends and schedules (docs/audits.md, sections 14 and 15). Trends are read from stored
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
    const stored = await modules.database.tenant(session.tenantId, (sql) =>
      readStoredAuditSchedule(sql, session.tenantId),
    );
    return settingsOf(stored);
  });

  // Changes the audits the body names and leaves the others as they are, so two pages saving
  // different audits never undo each other. An id that is not a known audit is dropped. Turning
  // an audit on, or changing its cadence, needs the saving administrator to be able to run it
  // now, by the Run button's rule, and stores the permissions their own sign-in granted among its
  // packs: its scheduled runs read with no more than those, whoever's sign-in they use. You can
  // schedule only what you could run yourself, and a schedule never reads more than you could.
  // Turning one off is always allowed. The CSRF check is the global hook's.
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
    const current = await modules.database.tenant(session.tenantId, (sql) =>
      readAuditSchedule(sql, session.tenantId, builtInAudits),
    );
    const changes = auditScheduleChanges(current, body.data);
    for (const change of changes) {
      const audit = findAudit(change.auditId);
      if (!audit || change.frequency === "off") continue;
      const availability = context.auditAvailability(audit, session);
      if (availability.state !== "available")
        return reply.status(403).send({
          error: availability.state === "not_enabled" ? "feature_not_enabled" : "missing_consent",
          message: auditUnavailableMessage(availability),
          auditId: change.auditId,
          availability,
        });
    }
    // The ledger entry comes first, so no schedule change is ever unrecorded. It names the audits
    // changed and their cadence, and how many audits run on their own after it.
    const after = {
      ...current,
      ...Object.fromEntries(changes.map((c) => [c.auditId, c.frequency])),
    };
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "audits.schedule.changed",
      changes,
      scheduled: scheduledAuditCount(after),
      occurredAt: new Date().toISOString(),
    });
    // Only the changes are written, so an unchanged audit keeps the administrator who scheduled
    // it.
    const saved = await modules.database.tenant(session.tenantId, (sql) =>
      writeAuditSchedule(
        sql,
        session.tenantId,
        { audits: Object.fromEntries(changes.map((change) => [change.auditId, change.frequency])) },
        session.actorId,
        builtInAudits,
        {
          name: session.displayName ?? null,
          grants: (audit) =>
            auditReadScope(audit, context.settings.enabledFeatures, session.scopes).granted,
        },
      ).then(() => readStoredAuditSchedule(sql, session.tenantId)),
    );
    return settingsOf(saved);
  });

  // Every audit with its frequency, whether collection is on, and the scheduled audits that read
  // less than a fully granted manual run would.
  const settingsOf = (stored: StoredAuditSchedule) =>
    auditScheduleSettingsSchema.parse({
      audits: knownAuditSchedule(stored, builtInAudits),
      collecting: context.collectingAdmin() !== null,
      limits: scheduleLimits(stored, builtInAudits, context.settings.enabledFeatures),
    });
}

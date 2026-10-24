import { digestQuerySchema, digestSchema } from "@prism365/contracts";
import { DIRECTORY_CHANGES_MODULE } from "@prism365/module-directory-changes";
import { HEALTH_MODULE, MESSAGES_MODULE } from "@prism365/module-service-health";
import { MODULE_ID as SIGNINS_MODULE } from "@prism365/module-signins-audit";
import type { FastifyInstance } from "fastify";
import {
  buildDigest,
  type DigestSourceName,
  digestSectionsRead,
  unavailableDigest,
} from "../digest.ts";
import type { RouteContext } from "./context.ts";

// The "This week" page (docs/this-week.md): what Prism365 stored in the last 1, 7, 14 or 30 days.
// It reads stored rows only and never asks Microsoft anything. The access-ledger entry is
// written before the first read and names the sections read, never a person or a value.

const sources: { name: DigestSourceName; module: string; noun: string }[] = [
  { name: "signins", module: SIGNINS_MODULE, noun: "The sign-in audit" },
  { name: "changes", module: DIRECTORY_CHANGES_MODULE, noun: "The directory change feed" },
  { name: "health", module: HEALTH_MODULE, noun: "Service health" },
  { name: "messages", module: MESSAGES_MODULE, noun: "Message center" },
];

// Why a module's section cannot be read, for each module that cannot be: it is not installed, or
// its read pack is not enabled (so nothing new is collected for it, and what is stored is stale).
function unavailableSources(context: RouteContext): Partial<Record<DigestSourceName, string>> {
  const unavailable: Partial<Record<DigestSourceName, string>> = {};
  for (const { name, module: id, noun } of sources) {
    const module = context.modules?.registry.get(id);
    if (!module) {
      unavailable[name] = `${noun} is not available in this installation.`;
      continue;
    }
    const off = module.packs.filter(
      (pack) => !context.settings.enabledFeatures.some((enabled) => enabled === pack),
    );
    if (off.length > 0)
      unavailable[name] =
        `This installation has not enabled ${off.join(", ")}. Add it to ENABLED_READ_FEATURES and restart.`;
  }
  return unavailable;
}

export function registerDigestRoutes(app: FastifyInstance, context: RouteContext): void {
  app.get("/v1/digest", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session)
      return reply.status(401).send({
        error: "authentication_required",
        message: "Sign in with an authorized administrator account.",
      });
    const query = digestQuerySchema.safeParse(request.query ?? {});
    if (!query.success)
      return reply
        .status(400)
        .send({ error: "invalid_request", message: "Choose 1, 7, 14 or 30 days." });
    const { days } = query.data;
    const now = new Date();
    const modules = context.modules;
    if (!modules)
      return digestSchema.parse(
        unavailableDigest(
          now,
          days,
          "Module storage is not configured for this installation, so nothing has been collected.",
        ),
      );
    const unavailable = unavailableSources(context);
    await context.store.appendAudit(session.tenantId, session.actorId, {
      type: "digest.viewed",
      datasets: digestSectionsRead(unavailable),
      days,
      occurredAt: now.toISOString(),
    });
    const digest = await modules.database.tenant(session.tenantId, (sql) =>
      buildDigest(sql, {
        tenantId: session.tenantId,
        now,
        days,
        collecting: modules.status().state === "collecting",
        modules: modules.registry.list().map((module) => ({
          id: module.id,
          title: module.title,
          collectors: module.collectors.map(({ id, title }) => ({ id, title })),
        })),
        unavailable,
      }),
    );
    return digestSchema.parse(digest);
  });
}

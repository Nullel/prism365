import { type Audit, builtInAudits, latestRuns as latestAuditRuns } from "@prism365/audits";
import {
  type AuditAvailability,
  type SetupCheck,
  setupStatusQuerySchema,
  setupStatusSchema,
} from "@prism365/contracts";
import {
  GraphBoundaryError,
  type GraphPermission,
  GraphRequestError,
  getPermissionPack,
} from "@prism365/graph";
import { readCollectionIntent } from "@prism365/modules";
import type { FastifyInstance } from "fastify";
import { z } from "zod";
import { type SupportedReadFeature, supportedReadFeatures } from "../config.ts";
import type { IdentitySession } from "../identity.ts";
import { RouteBoundaryError, type RouteContext, type RouteSettings } from "./context.ts";

// The first-run checklist (docs/SETUP.md): what this installation has set up and what it has
// not, each with its fix. It reads only public facts about the configuration (never a key or a
// secret) and the requesting administrator's own session. `?probe=1` adds one small read per
// granted pack, recorded in the access ledger first, so Microsoft's own answer can be shown.

const DAY = 86_400_000;

// The smallest read each pack allows: one item and only its ID where the endpoint supports it,
// otherwise a list that is short in every tenant. A pack without an entry is not probed.
export const probePaths: Partial<
  Record<SupportedReadFeature, (session: IdentitySession) => string>
> = {
  "directory.users.read": () => "users?$top=1&$select=id",
  // The administrator's own licence details: the pack's own endpoint, on an account that exists.
  "user.licenses.read": (session) => `users/${session.actorId}/licenseDetails?$select=id`,
  "tenant.licenses.read": () => "subscribedSkus",
  "audit.signins.read": () => "auditLogs/signIns?$top=1",
  "audit.directory.read": () => "auditLogs/directoryAudits?$top=1",
  "reports.authentication-methods.read": () =>
    "reports/authenticationMethods/userRegistrationDetails?$top=1&$select=id",
  "service.health.read": () => "admin/serviceAnnouncement/healthOverviews",
  "service.messages.read": () => "admin/serviceAnnouncement/messages?$top=1&$select=id",
  "identity.conditional-access.read": () => "identity/conditionalAccess/policies?$select=id",
  "identity.risk.read": () => "identityProtection/riskyUsers?$top=1&$select=id",
  "devices.intune.read": () => "deviceManagement/managedDevices?$top=1&$select=id",
  "directory.applications.read": () => "servicePrincipals?$top=1&$select=id",
  "directory.permission-grants.read": () => "oauth2PermissionGrants?$top=1",
  "reports.usage.read": () => "reports/getM365AppUserDetail(period='D7')?$format=application/json",
  "reports.settings.read": () => "admin/reportSettings",
};

export type ProbeResult =
  | { outcome: "answered" }
  | { outcome: "refused" }
  | { outcome: "failed"; reason: string }
  | { outcome: "untested" };

export interface SetupFacts {
  settings: RouteSettings;
  now: Date;
  // Microsoft permissions the requesting session granted, among the enabled ones.
  granted: ReadonlySet<string>;
  // Test reads by pack, when the administrator asked for them.
  probes: ReadonlyMap<SupportedReadFeature, ProbeResult> | null;
  // null when this installation has no module runtime.
  collection: {
    state: "stopped" | "collecting";
    boundAt: string | null;
    // Who collection runs as now, and whose standing choice it is to collect.
    collectingAs: string | null;
    chosenBy: string | null;
    viewer: string;
  } | null;
  // null when audits are not available here.
  audits:
    | {
        id: string;
        title: string;
        lastRun: { status: string; startedAt: string } | null;
        availability: AuditAvailability;
      }[]
    | null;
}

const date = (value: string | number) => new Date(value).toISOString().slice(0, 10);
const check = (
  value: Omit<SetupCheck, "grantFeature"> & { grantFeature?: string | null },
): SetupCheck => ({ grantFeature: null, ...value });

function certificateCheck(settings: RouteSettings, now: Date): SetupCheck | null {
  if (settings.credentialKind !== "certificate") return null;
  const rotate =
    "Run npm run setup -- --tenant <your tenant> --rotate-certificate, then restart Prism365. The old certificate stays on the registration until you remove it.";
  const expires = settings.certificate ? Date.parse(settings.certificate.expiresAt) : Number.NaN;
  if (Number.isNaN(expires))
    return check({
      id: "certificate",
      state: "unknown",
      title: "Certificate date unknown",
      detail: "Prism365 could not read when the configured certificate expires.",
      fix: "Check that ENTRA_CERTIFICATE_PATH points to the public certificate, then restart Prism365.",
    });
  if (expires <= now.getTime())
    return check({
      id: "certificate",
      state: "fail",
      title: "The certificate has expired",
      detail: `It expired on ${date(expires)}. Microsoft refuses sign-in with it, so nobody can sign in and collection stops.`,
      fix: rotate,
    });
  if (expires - now.getTime() < 30 * DAY) {
    const days = Math.floor((expires - now.getTime()) / DAY);
    return check({
      id: "certificate",
      state: "warn",
      title: "The certificate expires soon",
      detail: `It expires on ${date(expires)}, ${days < 1 ? "within a day" : `in ${days === 1 ? "1 day" : `${days} days`}`}. After that nobody can sign in.`,
      fix: rotate,
    });
  }
  return check({
    id: "certificate",
    state: "ok",
    title: "The certificate is valid",
    detail: `Valid until ${date(expires)}.`,
    fix: null,
  });
}

function featureCheck(
  feature: SupportedReadFeature,
  granted: ReadonlySet<string>,
  probe: ProbeResult | undefined,
): SetupCheck {
  const pack = getPermissionPack(feature);
  const permission = supportedReadFeatures[feature];
  const base = { id: `feature.${feature}`, title: `${pack.label} (${permission})` };
  if (!granted.has(permission))
    return check({
      ...base,
      state: "warn",
      detail: `Enabled in this installation, but your Microsoft session has not granted ${permission}.`,
      fix: `Grant access. Your account also needs a role Microsoft accepts: ${pack.roleNotes}`,
      grantFeature: feature,
    });
  if (probe?.outcome === "refused")
    return check({
      ...base,
      state: "fail",
      detail: `Microsoft refused the test read although ${permission} is granted, so it needs a role such as these: ${pack.roleNotes}`,
      fix: "Ask a Global Administrator for one of those roles, then sign out and in again. Some packs also need a licence (docs/SETUP.md).",
    });
  if (probe?.outcome === "failed")
    return check({
      ...base,
      state: "unknown",
      detail: `Granted, but the test read did not complete: ${probe.reason}.`,
      fix: "Try again in a few minutes.",
    });
  return check({
    ...base,
    state: "ok",
    detail:
      probe?.outcome === "answered"
        ? "Granted, and Microsoft answered a test read."
        : probe?.outcome === "untested"
          ? "Granted. Prism365 has no test read for this pack."
          : "Granted by your Microsoft session.",
    fix: null,
  });
}

function collectionCheck(collection: SetupFacts["collection"]): SetupCheck {
  if (!collection)
    return check({
      id: "collection",
      state: "unknown",
      title: "Collection is not available",
      detail: "This installation runs without module storage, so nothing is collected.",
      fix: "Use the Docker install (scripts/docker-live.sh).",
    });
  if (collection.state === "collecting") {
    const as =
      collection.collectingAs === collection.viewer
        ? "with your session"
        : "with another administrator's session";
    return check({
      id: "collection",
      state: "ok",
      title: "Prism365 is collecting",
      detail: `Collecting ${as}${collection.boundAt ? ` since ${date(collection.boundAt)}` : ""}. After a restart it resumes when that administrator next opens Prism365.`,
      fix: null,
    });
  }
  if (collection.chosenBy === collection.viewer)
    return check({
      id: "collection",
      state: "warn",
      title: "Collection is paused",
      detail:
        "You started collecting, but it could not resume on your current session. Microsoft may have refused it a token.",
      fix: "Sign out and in again, or press Start collecting on the Sign-in audit page.",
    });
  if (collection.chosenBy)
    return check({
      id: "collection",
      state: "warn",
      title: "Collection is paused",
      detail: "Another administrator started collecting. It resumes when they next open Prism365.",
      fix: "Ask them to open Prism365, or press Start collecting on the Sign-in audit page to collect with your own session.",
    });
  return check({
    id: "collection",
    state: "fail",
    title: "Prism365 is not collecting",
    detail: "Nothing is collected until an administrator starts it.",
    fix: "Open the Sign-in audit page and press Start collecting.",
  });
}

function auditChecks(audits: SetupFacts["audits"]): SetupCheck[] {
  if (!audits)
    return [
      check({
        id: "audits",
        state: "unknown",
        title: "Audits are not available",
        detail:
          "Audits keep their results in module storage, which this installation does not have.",
        fix: "Use the Docker install (scripts/docker-live.sh).",
      }),
    ];
  return audits.map((audit) => {
    const base = { id: `audit.${audit.id}`, title: `${audit.title} has run` };
    if (audit.lastRun?.status === "succeeded")
      return check({
        ...base,
        state: "ok",
        detail: `Last run on ${date(audit.lastRun.startedAt)}.`,
        fix: null,
      });
    const blocked =
      audit.availability.state === "not_enabled"
        ? `Add ${audit.availability.enable.join(", ")} to ENABLED_READ_FEATURES and restart Prism365, then run it from Audits.`
        : audit.availability.state === "missing_permission"
          ? `Grant ${audit.availability.missing.join(", ")}, then run it from Audits.`
          : null;
    return check({
      ...base,
      state: "warn",
      detail: audit.lastRun
        ? `The last run, on ${date(audit.lastRun.startedAt)}, ${audit.lastRun.status === "running" ? "has not finished" : "failed"}.`
        : "It has never run.",
      fix: blocked ?? "Open Audits and press Run.",
      grantFeature:
        audit.availability.state === "missing_permission" ? audit.availability.grantFeature : null,
    });
  });
}

// Every check, in the order an administrator would work through them.
export function setupChecks(facts: SetupFacts): SetupCheck[] {
  const { settings } = facts;
  const checks: SetupCheck[] = [
    settings.configured
      ? check({
          id: "configuration",
          state: "ok",
          title: "Configuration is complete",
          detail:
            "The tenant, the app registration, the administrator list and a credential are set.",
          fix: null,
        })
      : check({
          id: "configuration",
          state: "fail",
          title: "Configuration is incomplete",
          detail: `Missing: ${settings.missingConfiguration.join(", ") || "unknown settings"}.`,
          fix: "Run npm run setup, or set these in .env and restart Prism365 (docs/SETUP.md).",
        }),
    settings.credentialKind === "certificate"
      ? check({
          id: "credential",
          state: "ok",
          title: "Signs in to Microsoft with a certificate",
          detail: "The app registration's credential is a certificate, as production mode needs.",
          fix: null,
        })
      : settings.credentialKind === "secret"
        ? check({
            id: "credential",
            state: "warn",
            title: "Signs in to Microsoft with a client secret",
            detail: "A client secret is for local development only; production mode refuses it.",
            fix: "Run npm run setup to make a certificate (or follow the certificate steps in docs/SETUP.md), then remove ENTRA_CLIENT_SECRET from .env and the secret from the registration.",
          })
        : check({
            id: "credential",
            state: "fail",
            title: "No credential is configured",
            detail: "Prism365 has no certificate or client secret to sign in to Microsoft with.",
            fix: "Run npm run setup (docs/SETUP.md).",
          }),
  ];
  const certificate = certificateCheck(settings, facts.now);
  if (certificate) checks.push(certificate);
  checks.push(
    check({
      id: "signin",
      state: "ok",
      title: "You are signed in and on the administrator list",
      detail:
        "Your account is in ENTRA_ADMIN_OBJECT_IDS. Microsoft's own role checks still apply to every read.",
      fix: null,
    }),
  );
  for (const feature of settings.enabledFeatures)
    checks.push(featureCheck(feature, facts.granted, facts.probes?.get(feature)));
  checks.push(
    settings.durableStorage
      ? check({
          id: "storage",
          state: "ok",
          title: "Storage survives restarts",
          detail: "Sessions and collected data are kept on disk.",
          fix: null,
        })
      : check({
          id: "storage",
          state: "warn",
          title: "Storage is lost on restart",
          detail:
            "Sessions and collected data are held in memory, so a restart signs everyone out and forgets what was collected.",
          fix: "Use the Docker install (scripts/docker-live.sh), or set DATABASE_URL and STATE_ENCRYPTION_KEY, or PRISM_DATA_DIR.",
        }),
    collectionCheck(facts.collection),
    ...auditChecks(facts.audits),
  );
  return checks;
}

// One test read on one pack, through a Graph client narrowed to that pack alone.
async function probe(
  context: RouteContext,
  session: IdentitySession,
  handle: string,
  feature: SupportedReadFeature,
): Promise<ProbeResult> {
  const path = probePaths[feature]?.(session);
  if (!path) return { outcome: "untested" };
  const permission = getPermissionPack(feature).permission as GraphPermission;
  let scoped: ReturnType<RouteContext["sessionGraph"]>;
  try {
    scoped = context.sessionGraph(session, handle, [feature]);
  } catch (error) {
    if (error instanceof RouteBoundaryError)
      return { outcome: "failed", reason: "the pack is not enabled" };
    throw error;
  }
  if (!scoped.granted.includes(permission))
    return { outcome: "failed", reason: "your session has not granted it" };
  try {
    await scoped.graph.get(path, { feature, permission }, z.unknown());
    return { outcome: "answered" };
  } catch (error) {
    if (error instanceof GraphRequestError)
      return error.status === 403
        ? { outcome: "refused" }
        : { outcome: "failed", reason: `Microsoft answered HTTP ${error.status}` };
    if (error instanceof GraphBoundaryError)
      return {
        outcome: "failed",
        reason:
          error.reason === "timeout"
            ? "Microsoft did not answer in time"
            : "Prism365 could not use Microsoft's answer",
      };
    // The token request itself failed: Microsoft refused to renew the session.
    return { outcome: "failed", reason: "Microsoft did not issue a token; sign out and in again" };
  }
}

export function registerSetupRoutes(app: FastifyInstance, context: RouteContext): void {
  app.get("/v1/setup/status", async (request, reply) => {
    const query = setupStatusQuerySchema.safeParse(request.query ?? {});
    if (!query.success)
      return reply
        .status(400)
        .send({ error: "invalid_request", message: "Use probe=1 to check with Microsoft." });
    const { settings } = context;
    const visible = {
      mode: settings.mode,
      configured: settings.configured,
      missingConfiguration: [...settings.missingConfiguration],
    };
    const session = settings.mode === "live" ? await context.sessionFor(request) : null;
    const handle = context.sessionHandle(request);
    // Without an allow-listed session: nothing beyond what /v1/session tells anyone.
    if (!session || !handle)
      return setupStatusSchema.parse({ ...visible, checks: [], probed: false });

    const enabled = new Set(context.enabledPermissions());
    const granted = new Set(session.scopes.filter((scope) => enabled.has(scope)));
    let probes: Map<SupportedReadFeature, ProbeResult> | null = null;
    if (query.data.probe === "1") {
      probes = new Map();
      const packs = settings.enabledFeatures.filter(
        (feature) => granted.has(supportedReadFeatures[feature]) && probePaths[feature],
      );
      if (packs.length > 0) {
        // The ledger entry comes first: no test read without one. It names packs, nothing read.
        await context.store.appendAudit(session.tenantId, session.actorId, {
          type: "setup.probed",
          packs,
          occurredAt: new Date().toISOString(),
        });
        const results = await Promise.all(
          packs.map((feature) => probe(context, session, handle, feature)),
        );
        for (const [index, feature] of packs.entries())
          probes.set(feature, results[index] ?? { outcome: "untested" });
      }
      for (const feature of settings.enabledFeatures)
        if (!probes.has(feature)) probes.set(feature, { outcome: "untested" });
    }

    const modules = context.modules;
    const tenantId = session.tenantId;
    const stored = modules
      ? await modules.database.tenant(tenantId, async (sql) => ({
          intent: await readCollectionIntent(sql, tenantId),
          runs: context.audits ? await latestAuditRuns(sql, tenantId) : [],
        }))
      : null;
    const status = modules?.status();
    return setupStatusSchema.parse({
      ...visible,
      probed: probes !== null,
      checks: setupChecks({
        settings,
        now: new Date(),
        granted,
        probes,
        collection: status
          ? {
              state: status.state,
              boundAt: status.boundAt,
              collectingAs: context.collectingAdmin()?.actorId ?? null,
              chosenBy: stored?.intent?.actorId ?? null,
              viewer: session.actorId,
            }
          : null,
        audits:
          context.audits && stored
            ? builtInAudits.map((audit: Audit) => {
                const run = stored.runs.find((item) => item.auditId === audit.definition.id);
                return {
                  id: audit.definition.id,
                  title: audit.definition.title,
                  lastRun: run ? { status: run.status, startedAt: run.startedAt } : null,
                  availability: context.auditAvailability(audit, session),
                };
              })
            : null,
      }),
    });
  });
}

import { type Audit, builtInAudits, latestRuns as latestAuditRuns } from "@prism365/audits";
import {
  type AuditAvailability,
  type SetupCheck,
  setupProbeBodySchema,
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
// secret) and the requesting administrator's own session. GET /v1/setup/status only reads what is
// stored here; POST /v1/setup/probe (CSRF-checked like every POST) adds one small read per granted
// pack, recorded in the access ledger first, so Microsoft's own answer can be shown.

const DAY = 86_400_000;

// The smallest read each pack allows, by Microsoft's documentation of each endpoint: one item and
// only its ID where the endpoint supports $top and $select, a filter that matches nothing where it
// supports only $filter, and the endpoint's own small answer otherwise. A pack without an entry is
// not probed: the usage report returns a row per person, so it is checked when its audit runs.
export const probePaths: Partial<
  Record<SupportedReadFeature, (session: IdentitySession) => string>
> = {
  "directory.users.read": () => "users?$top=1&$select=id",
  // The administrator's own licence details: the pack's own endpoint, on an account that exists.
  "user.licenses.read": (session) => `users/${session.actorId}/licenseDetails?$select=id`,
  "tenant.licenses.read": () => "subscribedSkus?$select=skuId",
  "audit.signins.read": () => "auditLogs/signIns?$top=1",
  "audit.directory.read": () => "auditLogs/directoryAudits?$top=1",
  "reports.authentication-methods.read": () =>
    `reports/authenticationMethods/userRegistrationDetails?$filter=${encodeURIComponent("userPrincipalName eq 'probe@prism365.invalid'")}`,
  "service.health.read": () => "admin/serviceAnnouncement/healthOverviews?$select=id",
  "service.messages.read": () => "admin/serviceAnnouncement/messages?$top=1&$select=id",
  "identity.conditional-access.read": () => "identity/conditionalAccess/policies?$top=1&$select=id",
  "identity.risk.read": () => "identityProtection/riskyUsers?$top=1&$select=id",
  "devices.intune.read": () => "deviceManagement/managedDevices?$top=1&$select=id",
  "directory.applications.read": () => "servicePrincipals?$top=1&$select=id",
  "directory.permission-grants.read": () =>
    `oauth2PermissionGrants?$filter=${encodeURIComponent("clientId eq '00000000-0000-0000-0000-000000000000'")}`,
  "directory.roles.read": () =>
    `roleManagement/directory/roleDefinitions?$filter=${encodeURIComponent("id eq '62e90394-69f5-4237-9190-012177145e10'")}&$select=id`,
  "directory.groups.read": () => "groups?$top=1&$select=id",
  // Microsoft documents a known issue with $top on the domain list; a tenant has few domains.
  "directory.domains.read": () => "domains?$select=id",
  "reports.settings.read": () => "admin/reportSettings",
};

export type ProbeResult =
  | { outcome: "answered" }
  | { outcome: "refused" }
  | { outcome: "failed"; reason: string }
  | { outcome: "untested" };

export interface SetupFacts {
  settings: RouteSettings;
  // Whether collected module data survives a restart; null without a module runtime.
  moduleDataDurable: boolean | null;
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
          ? feature === "reports.usage.read"
            ? "Granted. The usage report lists every person, so it is checked when its audit runs, not here."
            : "Granted. Prism365 has no test read for this pack."
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
    settings.durableSessions
      ? check({
          id: "storage.sessions",
          state: "ok",
          title: "Sessions and the access ledger survive restarts",
          detail: "They are kept, encrypted, in PostgreSQL.",
          fix: null,
        })
      : check({
          id: "storage.sessions",
          state: "warn",
          title: "Sessions and the access ledger are lost on restart",
          detail:
            "They are held in memory, so a restart signs everyone out and the record of who read what is gone. PRISM_DATA_DIR does not change that; it keeps only collected data.",
          fix: "Use the Docker install (scripts/docker-live.sh), or set DATABASE_URL and STATE_ENCRYPTION_KEY.",
        }),
    ...(facts.moduleDataDurable === null
      ? []
      : [
          facts.moduleDataDurable
            ? check({
                id: "storage.data",
                state: "ok",
                title: "Collected data survives restarts",
                detail: "Sign-ins, findings and audit results are kept on disk.",
                fix: null,
              })
            : check({
                id: "storage.data",
                state: "warn",
                title: "Collected data is lost on restart",
                detail: "Sign-ins, findings and audit results are held in memory.",
                fix: "Use the Docker install (scripts/docker-live.sh), or set PRISM_DATA_DIR or a PostgreSQL database.",
              }),
        ]),
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

// The checklist for one signed-in administrator, with the probe results when there are any.
async function checklist(
  context: RouteContext,
  session: IdentitySession,
  granted: ReadonlySet<string>,
  probes: ReadonlyMap<SupportedReadFeature, ProbeResult> | null,
) {
  const { settings, modules } = context;
  const tenantId = session.tenantId;
  const stored = modules
    ? await modules.database.tenant(tenantId, async (sql) => ({
        intent: await readCollectionIntent(sql, tenantId),
        runs: context.audits ? await latestAuditRuns(sql, tenantId) : [],
      }))
    : null;
  const status = modules?.status();
  return setupChecks({
    settings,
    moduleDataDurable: modules ? modules.database.durable : null,
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
  });
}

export function registerSetupRoutes(app: FastifyInstance, context: RouteContext): void {
  const { settings } = context;
  const visible = () => ({
    mode: settings.mode,
    configured: settings.configured,
    missingConfiguration: [...settings.missingConfiguration],
  });
  const grantedFor = (session: IdentitySession) => {
    const enabled = new Set(context.enabledPermissions());
    return new Set(session.scopes.filter((scope) => enabled.has(scope)));
  };

  // Read-only: it never reads Microsoft and writes nothing, so a link to it can do no harm.
  app.get("/v1/setup/status", async (request, reply) => {
    if (!setupStatusQuerySchema.safeParse(request.query ?? {}).success)
      return reply.status(400).send({
        error: "invalid_request",
        message:
          "This address takes no options. To check with Microsoft, use POST /v1/setup/probe.",
      });
    const session = settings.mode === "live" ? await context.sessionFor(request) : null;
    // Without an allow-listed session: nothing beyond what /v1/session tells anyone.
    if (!session) return setupStatusSchema.parse({ ...visible(), checks: [], probed: false });
    return setupStatusSchema.parse({
      ...visible(),
      probed: false,
      checks: await checklist(context, session, grantedFor(session), null),
    });
  });

  // One small read per granted pack. A POST, so the global hook checks the session's CSRF value
  // and origin: a link or a page on another site cannot make an administrator's browser run it.
  app.post("/v1/setup/probe", async (request, reply) => {
    if (!setupProbeBodySchema.safeParse(request.body ?? {}).success)
      return reply.status(400).send({ error: "invalid_request", message: "Send no options." });
    const session = settings.mode === "live" ? await context.sessionFor(request) : null;
    const handle = context.sessionHandle(request);
    if (!session || !handle) return reply.status(401).send({ error: "authentication_required" });
    const granted = grantedFor(session);
    const probes = new Map<SupportedReadFeature, ProbeResult>();
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
    return setupStatusSchema.parse({
      ...visible(),
      probed: true,
      checks: await checklist(context, session, granted, probes),
    });
  });
}

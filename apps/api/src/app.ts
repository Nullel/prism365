import { randomUUID } from "node:crypto";
import staticFiles from "@fastify/static";
import {
  AiAnalysisBroker,
  type AiEvidenceQuery,
  type AiProvider,
  LocalOllamaProvider,
} from "@prism365/ai";
import {
  type Audit,
  auditDefinition,
  auditReadScope,
  builtInAudits,
  findAudit,
  getRun,
  latestRuns as latestAuditRuns,
  listRuns,
  readAuditSettings,
  writeAuditScope,
} from "@prism365/audits";
import {
  type AiEvidenceRecord,
  type AuditAvailability,
  aiAccessGrantSchema,
  attentionItemSchema,
  attentionListSchema,
  attentionMuteRequestSchema,
  attentionPreferencesSchema,
  attentionReviewRequestSchema,
  attentionSettingsSchema,
  auditCatalogSchema,
  auditCsv,
  auditIdSchema,
  auditRunDetailSchema,
  auditRunListSchema,
  auditScopeSchema,
  auditSettingsSchema,
  auditStartSchema,
  capabilitySummarySchema,
  collectionStatusSchema,
  findingListSchema,
  findingStateRequestSchema,
  findingStateSchema,
  investigationEmailSchema,
  moduleListSchema,
  personIdSchema,
  personListSchema,
  personLookupSchema,
  personRecordSchema,
  personSigninPageSchema,
  queryAuditRows,
  type SystemAlert,
  scopeFileSuffix,
  serviceHealthStatusSchema,
  sessionSchema,
  signinSummarySchema,
  type UserInvestigation,
} from "@prism365/contracts";
import {
  createDemoUserInvestigation,
  type GraphFeatureId,
  GraphHttpClient,
  investigateUser,
  lookupPersonId,
  microsoftCloudProfiles,
  permissionPacks,
  readPersonDirectory,
} from "@prism365/graph";
import {
  HEALTH_MODULE,
  MESSAGES_MODULE,
  serviceHealthStatus,
} from "@prism365/module-service-health";
import {
  listPeople,
  personHistory,
  personSignins,
  MODULE_ID as SIGNINS_MODULE,
  summarizeSignins,
} from "@prism365/module-signins-audit";
import {
  attentionCounts,
  type CollectionIntent,
  clearCollectionIntent,
  latestRuns,
  listAttention,
  listFindings,
  markAttentionRead,
  modulePermissions,
  readAttentionPreferences,
  readCollectionIntent,
  scopeGraph,
  setAttentionMuted,
  setAttentionReview,
  setFindingState,
  writeAttentionPreferences,
  writeCollectionIntent,
} from "@prism365/modules";
import { MemoryStateStore, type SqlClient, type StateStore } from "@prism365/storage";
import Fastify, { type FastifyRequest } from "fastify";
import { z } from "zod";
import { auditExportParamsSchema, auditExportQuery } from "./audit-export.ts";
import { AuditScheduler, type AuditSchedulerOptions } from "./audit-schedule.ts";
import { AuditRunner, auditUnavailableMessage } from "./audits.ts";
import {
  type AppConfig,
  DEFAULT_HANDOFF_RETENTION_DAYS,
  type SupportedReadFeature,
  supportedReadFeatures,
} from "./config.ts";
import { DeliveryService, type DeliveryServiceOptions } from "./delivery.ts";
import { buildDigest, digestSectionsRead } from "./digest.ts";
import { demoDashboard } from "./demo-data.ts";
import {
  digest,
  equalSecret,
  IdentityService,
  type IdentitySession,
  LoginRejectedError,
} from "./identity.ts";
import { DelegatedSessionCredential, type ModuleRuntime } from "./modules.ts";
import { registerAuditInsightRoutes } from "./routes/audit-insights.ts";
import { registerCanaryRoutes } from "./routes/canaries.ts";
import { registerChangeRoutes } from "./routes/changes.ts";
import { RouteBoundaryError, type RouteContext } from "./routes/context.ts";
import { registerDeliveryRoutes } from "./routes/delivery.ts";
import { digestUnavailable, registerDigestRoutes } from "./routes/digest.ts";
import { registerHandoffRoutes } from "./routes/handoffs.ts";
import { registerSetupRoutes } from "./routes/setup.ts";

export interface AppOptions {
  config?: AppConfig;
  store?: StateStore;
  identity?: IdentityService;
  graphFetch?: typeof fetch;
  webRoot?: string;
  aiProvider?: AiProvider;
  modules?: ModuleRuntime;
  // Defaults to a runner on the module database. Tests pass one with a short time limit.
  audits?: AuditRunner;
  // Timing for scheduled audits. Tests pass their own clock and timers.
  auditSchedule?: Pick<
    AuditSchedulerOptions,
    "now" | "intervalMilliseconds" | "startDelayMilliseconds" | "timers" | "log"
  >;
  // Transports, clock and timers for alert delivery. Tests pass fakes; nothing here can change
  // which addresses a webhook may reach.
  delivery?: Pick<
    DeliveryServiceOptions,
    | "lookup"
    | "connect"
    | "mailSender"
    | "now"
    | "timers"
    | "retryDelaysMilliseconds"
    | "flushDelayMilliseconds"
    | "checkIntervalMilliseconds"
    | "log"
  >;
}

function readCookie(request: FastifyRequest, name: string) {
  const match = (request.headers.cookie ?? "")
    .split(";")
    .map((part) => part.trim())
    .find((part) => part.startsWith(`${name}=`));
  return match?.slice(name.length + 1);
}

export function createApp(options: AppOptions = {}) {
  const config =
    options.config ??
    ({
      mode: "demo",
      origin: "http://localhost:5173",
      production: false,
      identity: null,
      missingConfiguration: [],
      enabledFeatures: ["directory.users.read"],
    } satisfies AppConfig);
  const store = options.store ?? new MemoryStateStore();
  const identity =
    options.identity ?? (config.identity ? new IdentityService(config.identity, store) : null);
  const secureCookies = config.origin.startsWith("https:");
  const cookieName = secureCookies ? "__Host-prism365_session" : "prism365_session";
  const oauthCookieName = secureCookies ? "__Host-prism365_oauth" : "prism365_oauth";
  const cookie = (name: string, value: string, maxAge: number) =>
    `${name}=${value}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${maxAge}${secureCookies ? "; Secure" : ""}`;
  const sessionFor = (request: FastifyRequest) =>
    identity?.session(readCookie(request, cookieName)) ?? Promise.resolve(null);
  const rates = new Map<string, { count: number; resetAt: number }>();
  const aiAvailable = Boolean(
    config.mode === "live" && identity && store.durable && config.aiModel,
  );
  const modules = config.mode === "live" && identity ? (options.modules ?? null) : null;
  // Audits store their runs in the module database, so they exist wherever modules do.
  const audits = modules
    ? (options.audits ?? new AuditRunner({ database: modules.database }))
    : null;
  // Read packs this installation enabled, as Microsoft permission names. Background collection
  // never uses a consented scope the operator did not enable here.
  const enabledPermissions = () => [
    "User.Read",
    ...config.enabledFeatures.map((feature) => supportedReadFeatures[feature]),
  ];
  const unbindSession = (handle: string | undefined) => {
    if (!modules || !handle) return;
    const key = digest(handle);
    if (
      modules.scheduler.isBound(
        (credential) =>
          credential instanceof DelegatedSessionCredential && credential.sessionKey === key,
      )
    )
      modules.scheduler.unbind("signed_out");
  };
  // The standing choice to collect (see packages/modules/src/collection-intent.ts), cached
  // after the first read. undefined until read.
  let intent: CollectionIntent | null | undefined;
  // Start, stop and resume change the choice and the binding together, one at a time, so a stop
  // can never be overtaken by a resume that read the choice before it.
  let collectionQueue: Promise<unknown> = Promise.resolve();
  const exclusive = <T>(work: () => Promise<T>): Promise<T> => {
    const next = collectionQueue.then(work, work);
    collectionQueue = next.catch(() => {});
    return next;
  };
  const collectionIntent = async (): Promise<CollectionIntent | null> => {
    if (intent !== undefined || !modules || !identity) return intent ?? null;
    const tenantId = identity.tenantId;
    intent = await modules.database.tenant(tenantId, (sql) => readCollectionIntent(sql, tenantId));
    return intent;
  };
  // Sessions Microsoft refused a token for while collecting (a sign-in frequency policy, a
  // revoked refresh token). They never resume collection again: otherwise every page load would
  // rebind a session Microsoft keeps refusing, and every run would fail and unbind it. Signing in
  // again gives a new session that may; pressing Start retries this one.
  const refusedSessions = new Set<string>();
  let collecting: DelegatedSessionCredential | null = null;
  const bindCollection = (handle: string, session: IdentitySession) => {
    if (!modules || !identity) return;
    const credential = new DelegatedSessionCredential(
      identity,
      handle,
      session,
      enabledPermissions,
    );
    modules.scheduler.bind(credential);
    collecting = credential;
  };
  // Resumes collection on the session of the administrator who started it, when nothing is
  // collecting: after a restart, a sign-out or an expired session.
  const resumeCollection = (session: IdentitySession, handle: string | undefined) =>
    exclusive(async () => {
      if (!modules || !identity || !handle) return;
      if (modules.scheduler.status().state === "collecting") return;
      if (refusedSessions.has(digest(handle))) return;
      const standing = await collectionIntent();
      if (!standing || standing.actorId !== session.actorId) return;
      // The ledger entry comes first: collection never reads without one.
      await store.appendAudit(session.tenantId, session.actorId, {
        type: "collection.resumed",
        credentialKind: "delegated",
        occurredAt: new Date().toISOString(),
      });
      bindCollection(handle, session);
      void modules.scheduler.tick().catch(() => {});
    });
  if (modules)
    modules.scheduler.subscribe((event) => {
      if (event.type === "collection.unbound") {
        if (event.reason === "credential_unavailable" && collecting?.refused) {
          if (refusedSessions.size >= 100) refusedSessions.clear();
          refusedSessions.add(collecting.sessionKey);
        }
        collecting = null;
      }
      // Operator stops are recorded by the stop route with the operator who stopped it.
      if (
        event.type === "collection.unbound" &&
        event.actorId &&
        event.reason !== "shutdown" &&
        event.reason !== "operator_stopped"
      )
        void store
          .appendAudit(event.tenantId, event.actorId, {
            type: "collection.stopped",
            reason: event.reason,
            occurredAt: new Date().toISOString(),
          })
          .catch(() => {});
    });

  const trustedHops = config.trustProxyHops ?? 0;
  const app = Fastify({
    logger: false,
    // Behind a proxy, request.ip would be the proxy for every caller, so all of them would share
    // one sign-in rate-limit bucket. Trust only as many hops as are configured.
    trustProxy: trustedHops > 0 ? (_address: string, hop: number) => hop < trustedHops : false,
    requestIdHeader: "x-request-id",
    bodyLimit: 1_000_000,
    rewriteUrl: (request) =>
      request.url?.startsWith("/api/") ? request.url.slice(4) : (request.url ?? "/"),
  });

  app.setErrorHandler((error, _request, reply) => {
    const statusCode = error instanceof Error && "statusCode" in error ? error.statusCode : 500;
    const status =
      typeof statusCode === "number" && statusCode >= 400 && statusCode < 500 ? statusCode : 500;
    reply.status(status).send({
      error: status === 413 ? "request_too_large" : "internal_error",
      message: "The request could not be completed. Check the server configuration and retry.",
    });
  });

  app.addHook("preHandler", async (request, reply) => {
    // Authorization keys on the route the router matched, never on the raw URL. The router
    // decodes percent-encoding, so "/%76%31/findings" reaches the /v1/findings handler; a string
    // test on request.url would let that request skip the session and CSRF checks below.
    // An unmatched request has no route and is answered by the 404 handler with no data.
    const pathname = request.routeOptions.url;
    if (pathname === undefined) return;
    if (config.mode === "live" && pathname === "/v1/auth/login") {
      const key = `login:${digest(request.ip)}`;
      const now = Date.now();
      for (const [existing, value] of rates) if (value.resetAt <= now) rates.delete(existing);
      if (rates.size >= 2_000 && !rates.has(key))
        return reply.status(429).send({ error: "rate_limited" });
      const rate = rates.get(key) ?? { count: 0, resetAt: now + 60_000 };
      rate.count += 1;
      rates.set(key, rate);
      if (rate.count > 10)
        return reply.header("Retry-After", "60").status(429).send({
          error: "rate_limited",
          message: "Wait a minute before starting another sign-in.",
        });
    }
    if (
      config.mode !== "live" ||
      !pathname?.startsWith("/v1/") ||
      ["/health", "/v1/session", "/v1/auth/login", "/v1/auth/callback"].includes(pathname ?? "")
    )
      return;
    const session = await sessionFor(request);
    // The setup checklist also answers without a session, so an installation that cannot sign
    // anyone in yet can still say what is missing. It then tells no more than /v1/session; a
    // signed-in request goes through every check below like any other.
    if (!session && pathname === "/v1/setup/status") return;
    if (!session)
      return reply.status(401).send({
        error: "authentication_required",
        message: "Sign in with an authorized administrator account.",
      });
    const now = Date.now();
    for (const [key, value] of rates) if (value.resetAt <= now) rates.delete(key);
    const key = readCookie(request, cookieName) ?? "";
    const rate = rates.get(key) ?? { count: 0, resetAt: now + 60_000 };
    rate.count += 1;
    rates.set(key, rate);
    if (rate.count > 60)
      return reply
        .header("Retry-After", "60")
        .status(429)
        .send({ error: "rate_limited", message: "Wait a minute before trying again." });
    // Collection paused by a restart or a sign-out picks up again on the next page load. Only
    // reads trigger it, and only for the administrator whose standing choice it is; a failure to
    // resume never fails the request.
    if (request.method === "GET")
      await Promise.race([
        resumeCollection(session, readCookie(request, cookieName)).catch(() => {}),
        // A stuck storage call must not hold every page load: wait at most five seconds.
        new Promise((resolve) => setTimeout(resolve, 5_000).unref()),
      ]);
    if (request.method !== "GET" && request.method !== "HEAD") {
      if (
        request.headers.origin !== config.origin ||
        typeof request.headers["x-csrf-token"] !== "string" ||
        !equalSecret(request.headers["x-csrf-token"], session.csrfToken)
      ) {
        return reply.status(403).send({
          error: "request_verification_failed",
          message: "Refresh the page before submitting this request.",
        });
      }
    }
  });

  app.addHook("onSend", async (request, reply, payload) => {
    reply.header("Cache-Control", "no-store");
    reply.header("X-Content-Type-Options", "nosniff");
    reply.header("X-Frame-Options", "DENY");
    reply.header("Referrer-Policy", "no-referrer");
    reply.header(
      "Content-Security-Policy",
      request.url.startsWith("/v1/") || request.url === "/health"
        ? "default-src 'none'; frame-ancestors 'none'"
        : "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; connect-src 'self'; object-src 'none'; base-uri 'self'; frame-ancestors 'none'",
    );
    return payload;
  });

  app.get("/health", async () => ({
    status: "ok",
    mode: config.mode,
    tenantConnected: false,
    connectionConfigured: Boolean(identity),
  }));

  app.get("/v1/session", async (request) => {
    const session = await sessionFor(request);
    return sessionSchema.parse({
      mode: config.mode,
      configured: Boolean(identity),
      authenticated: Boolean(session),
      displayName: session?.displayName ?? null,
      csrfToken: session?.csrfToken ?? null,
      enabledFeatures: config.mode === "live" ? config.enabledFeatures : [],
      missingConfiguration: config.missingConfiguration,
    });
  });

  app.get("/v1/auth/login", async (request, reply) => {
    if (config.mode !== "live" || !identity)
      return reply.status(503).send({
        error: "connection_not_configured",
        message: "Configure the Entra application before connecting a tenant.",
      });
    const query = z
      .object({
        // Any supported read pack. The service-health and audit "Grant access" links use this.
        feature: z
          .enum(
            Object.keys(supportedReadFeatures) as [SupportedReadFeature, ...SupportedReadFeature[]],
          )
          .default("directory.users.read"),
      })
      .safeParse(request.query);
    if (!query.success || !config.enabledFeatures.includes(query.data.feature))
      return reply.status(400).send({
        error: "feature_not_enabled",
        message: "The operator must enable this read feature before consent can be requested.",
      });
    const login = await identity.startLogin(query.data.feature, await sessionFor(request));
    reply.header("Set-Cookie", cookie(oauthCookieName, login.state, 300));
    return reply.redirect(login.url);
  });

  app.get("/v1/auth/callback", async (request, reply) => {
    if (!identity) return reply.status(503).send({ error: "connection_not_configured" });
    const parsed = z
      .object({ code: z.string().min(1).max(20_000), state: z.string().min(32).max(100) })
      .safeParse(request.query);
    reply.header("Set-Cookie", cookie(oauthCookieName, "", 0));
    if (!parsed.success) {
      console.warn("Prism365 sign-in rejected: invalid_callback_parameters");
      return reply.redirect(`${config.origin}/?auth=failed`);
    }
    try {
      const result = await identity.finishLogin(
        parsed.data.code,
        parsed.data.state,
        readCookie(request, oauthCookieName) ?? "",
      );
      const oldHandle = readCookie(request, cookieName);
      const oldSession = await sessionFor(request);
      if (oldHandle && oldSession) {
        unbindSession(oldHandle);
        await identity.logout(oldHandle, oldSession);
      }
      reply.header("Set-Cookie", [
        cookie(oauthCookieName, "", 0),
        cookie(cookieName, result.handle, 8 * 60 * 60),
      ]);
      return reply.redirect(config.origin);
    } catch (error) {
      // The browser only learns "failed". The operator's console gets the category so an
      // allowlist mistake can be told apart from a token problem without logging either.
      console.warn(`Prism365 sign-in rejected: ${signInFailureCategory(error)}`);
      return reply.redirect(`${config.origin}/?auth=failed`);
    }
  });

  app.post("/v1/auth/logout", async (request, reply) => {
    const handle = readCookie(request, cookieName);
    const session = await sessionFor(request);
    unbindSession(handle);
    if (identity && handle && session) await identity.logout(handle, session);
    reply.header("Set-Cookie", cookie(cookieName, "", 0));
    return { signedOut: true };
  });

  // `viewerHandle`: the session asking. A session that will not resume collection is told so.
  const collectionStatus = (viewerHandle?: string) => {
    const status = modules?.scheduler.status();
    return collectionStatusSchema.parse({
      available: Boolean(modules),
      state: status?.state ?? "stopped",
      credentialKind: status?.credentialKind ?? null,
      boundAt: status?.boundAt ?? null,
      lastStopReason: status?.lastStopReason ?? null,
      durableStorage: modules?.database.durable ?? false,
      resumes:
        status?.state !== "collecting" &&
        Boolean(intent) &&
        !(viewerHandle && refusedSessions.has(digest(viewerHandle))),
      reason:
        config.mode === "demo"
          ? "Demo mode does not collect tenant data."
          : modules
            ? null
            : "Module collection is not configured for this installation.",
    });
  };

  app.get("/v1/modules", async (request, reply) => {
    const session = await sessionFor(request);
    if (config.mode === "live" && !session)
      return reply.status(401).send({ error: "authentication_required" });
    const runs =
      modules && identity
        ? await modules.database.tenant(identity.tenantId, (sql) =>
            latestRuns(sql, identity.tenantId),
          )
        : [];
    const allowed = new Set(enabledPermissions());
    const granted = new Set((session?.scopes ?? []).filter((scope) => allowed.has(scope)));
    return moduleListSchema.parse({
      mode: config.mode,
      collection: collectionStatus(readCookie(request, cookieName)),
      modules: (modules?.registry.list() ?? []).map((module) => {
        const permissions = modulePermissions(module);
        return {
          id: module.id,
          title: module.title,
          description: module.description,
          packs: [...module.packs],
          permissions,
          missingPermissions: permissions.filter((permission) => !granted.has(permission)),
          collectors: module.collectors.map((collector) => {
            const run = runs.find(
              (candidate) => candidate.module === module.id && candidate.collector === collector.id,
            );
            return {
              id: collector.id,
              title: collector.title,
              cadenceSeconds: collector.cadenceSeconds,
              lastRun: run
                ? {
                    status: run.status,
                    credentialKind: run.credentialKind,
                    startedAt: run.startedAt,
                    finishedAt: run.finishedAt,
                    fetched: run.fetched,
                    inserted: run.inserted,
                    findings: run.findings,
                    skipped: run.skipped,
                    errorCategory: run.errorCategory,
                  }
                : null,
            };
          }),
        };
      }),
    });
  });

  app.post("/v1/collection/start", async (request, reply) => {
    const session = await sessionFor(request);
    const handle = readCookie(request, cookieName);
    if (!modules || !identity || !session || !handle)
      return reply.status(409).send({
        error: "collection_unavailable",
        message: "Background collection needs live mode and a signed-in administrator.",
      });
    await exclusive(async () => {
      const previous = await collectionIntent();
      const standing = { actorId: session.actorId, since: new Date().toISOString() };
      await modules.database.tenant(session.tenantId, (sql) =>
        writeCollectionIntent(sql, session.tenantId, standing),
      );
      intent = standing;
      try {
        await store.appendAudit(session.tenantId, session.actorId, {
          type: "collection.started",
          credentialKind: "delegated",
          modules: modules.registry.list().map((module) => module.id),
          occurredAt: new Date().toISOString(),
        });
      } catch (error) {
        // No ledger entry, no start: put back the choice that was there before (possibly another
        // administrator's), rather than erasing it.
        intent = previous;
        await modules.database
          .tenant(session.tenantId, (sql) =>
            previous
              ? writeCollectionIntent(sql, session.tenantId, previous)
              : clearCollectionIntent(sql, session.tenantId),
          )
          .catch(() => {});
        throw error;
      }
      // An explicit start retries a session Microsoft refused before.
      refusedSessions.delete(digest(handle));
      // Start again from the session already collecting keeps that binding: rebinding would
      // count as collection passing to someone else and stop the scheduled run it is reading
      // with. A binding Microsoft refused is replaced, so the start retries it.
      const current = collecting;
      if (
        current &&
        current.sessionKey === digest(handle) &&
        !current.refused &&
        modules.scheduler.status().state === "collecting"
      )
        return;
      bindCollection(handle, session);
    });
    void modules.scheduler.tick().catch(() => {});
    return collectionStatus();
  });

  app.post("/v1/collection/stop", async (request, reply) => {
    const session = await sessionFor(request);
    if (!modules || !session) return reply.status(409).send({ error: "collection_unavailable" });
    const saved = await exclusive(async () => {
      // Collection stops first, whatever happens to the stored choice after it.
      if (modules.scheduler.status().state === "collecting") {
        modules.scheduler.unbind("operator_stopped");
        await store
          .appendAudit(session.tenantId, session.actorId, {
            type: "collection.stopped",
            reason: "operator_stopped",
            occurredAt: new Date().toISOString(),
          })
          .catch(() => {});
      }
      // Stopping ends the standing choice too, so nothing resumes until someone starts it again.
      intent = null;
      try {
        await modules.database.tenant(session.tenantId, (sql) =>
          clearCollectionIntent(sql, session.tenantId),
        );
        return true;
      } catch {
        return false;
      }
    });
    if (!saved)
      return reply.status(500).send({
        error: "stop_not_saved",
        message:
          "Collection stopped, but the choice could not be saved. It may resume after the server restarts; press Stop again then.",
      });
    return collectionStatus();
  });

  app.get("/v1/findings", async (request, reply) => {
    if (config.mode === "demo") return findingListSchema.parse({ mode: "demo", items: [] });
    if (!(await sessionFor(request)))
      return reply.status(401).send({ error: "authentication_required" });
    if (!modules || !identity)
      return reply.status(409).send({
        error: "modules_unavailable",
        message: "Module findings are not configured for this installation.",
      });
    const query = z
      .object({
        module: z
          .string()
          .regex(/^[a-z][a-z0-9.-]{1,80}$/)
          .optional(),
        state: findingStateSchema.optional(),
        limit: z.coerce.number().int().min(1).max(200).optional(),
      })
      .safeParse(request.query);
    if (!query.success)
      return reply
        .status(400)
        .send({ error: "invalid_request", message: "Check the finding filters and retry." });
    const filter = {
      ...(query.data.module ? { module: query.data.module } : {}),
      ...(query.data.state ? { state: query.data.state } : {}),
      ...(query.data.limit ? { limit: query.data.limit } : {}),
    };
    const items = await modules.database.tenant(identity.tenantId, (sql) =>
      listFindings(sql, identity.tenantId, filter),
    );
    return findingListSchema.parse({ mode: "live", items });
  });

  app.get("/v1/modules/signins.audit/summary", async (request, reply) => {
    const session = await sessionFor(request);
    if (!modules || !session || !modules.registry.get(SIGNINS_MODULE))
      return reply.status(409).send({
        error: "modules_unavailable",
        message: "The sign-in audit module is not available in this installation.",
      });
    const query = z
      .object({ days: z.coerce.number().int().min(1).max(90).default(7) })
      .safeParse(request.query);
    if (!query.success)
      return reply.status(400).send({ error: "invalid_request", message: "Choose 1 to 90 days." });
    const summary = await modules.database.tenant(session.tenantId, (sql) =>
      summarizeSignins(sql, session.tenantId, new Date(), query.data.days),
    );
    return signinSummarySchema.parse(summary);
  });

  // People: one page per account, joining stored sign-ins with the live directory.
  const peopleAvailable = () => Boolean(modules?.registry.get(SIGNINS_MODULE));
  const delegatedGraph = (session: IdentitySession, handle: string) => {
    if (!identity) throw new Error("Live mode is not configured.");
    const permitted = new Set(enabledPermissions());
    const scopes = session.scopes.filter((scope) => permitted.has(scope));
    const client = new GraphHttpClient({
      requestTimeoutMilliseconds: 30_000,
      getAccessToken: () => identity.accessToken(handle, scopes),
      ...(options.graphFetch ? { fetchImplementation: options.graphFetch } : {}),
    });
    return { client, scopes };
  };
  const personParams = z.object({ userId: personIdSchema });

  app.get("/v1/people", async (request, reply) => {
    if (config.mode === "demo")
      return personListSchema.parse({ mode: "demo", days: 30, items: [], truncated: false });
    const session = await sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    if (!modules || !peopleAvailable())
      return reply.status(409).send({
        error: "modules_unavailable",
        message: "People needs the sign-in audit module, which is not available here.",
      });
    const query = z
      .object({
        search: z.string().max(200).optional(),
        days: z.coerce.number().int().min(1).max(90).default(30),
      })
      .safeParse(request.query);
    if (!query.success)
      return reply.status(400).send({ error: "invalid_request", message: "Check the search." });
    const result = await modules.database.tenant(session.tenantId, (sql) =>
      listPeople(sql, session.tenantId, {
        now: new Date(),
        days: query.data.days,
        ...(query.data.search ? { search: query.data.search } : {}),
      }),
    );
    return personListSchema.parse({ mode: "live", days: query.data.days, ...result });
  });

  // Finds the account behind a work email in the directory, for people with no stored sign-ins.
  app.get("/v1/people/lookup", async (request, reply) => {
    const session = await sessionFor(request);
    const handle = readCookie(request, cookieName);
    if (config.mode !== "live" || !session || !handle)
      return reply.status(401).send({ error: "authentication_required" });
    const query = z.object({ email: investigationEmailSchema }).safeParse(request.query);
    if (!query.success)
      return reply
        .status(400)
        .send({ error: "invalid_input", message: "Enter a valid work email address." });
    const { client, scopes } = delegatedGraph(session, handle);
    if (!scopes.includes("User.Read.All"))
      return reply.status(403).send({
        error: "missing_consent",
        message: "Looking people up needs the User.Read.All feature pack.",
      });
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "person.lookup",
      datasets: ["directory"],
      occurredAt: new Date().toISOString(),
    });
    try {
      const userId = await lookupPersonId({ userPrincipalName: query.data.email, client });
      if (!userId)
        return reply.status(404).send({
          error: "user_not_found",
          message: "No account has that email in the connected tenant.",
        });
      return personLookupSchema.parse({ userId });
    } catch {
      return reply.status(503).send({
        error: "directory_unavailable",
        message: "Microsoft's directory could not be read just now. Try again.",
      });
    }
  });

  app.get("/v1/people/:userId", async (request, reply) => {
    const session = await sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    if (!modules || !peopleAvailable())
      return reply.status(409).send({ error: "modules_unavailable" });
    const params = personParams.safeParse(request.params);
    if (!params.success)
      return reply.status(400).send({ error: "invalid_request", message: "Unknown account." });
    const { userId } = params.data;
    const record = await modules.database.tenant(session.tenantId, async (sql) => ({
      history: await personHistory(sql, session.tenantId, userId),
      findings: await listFindings(sql, session.tenantId, {
        module: SIGNINS_MODULE,
        personId: userId,
        limit: 200,
      }),
    }));
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "person.viewed",
      datasets: ["signIns", "findings"],
      occurredAt: new Date().toISOString(),
    });
    return personRecordSchema.parse(record);
  });

  app.get("/v1/people/:userId/signins", async (request, reply) => {
    const session = await sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    if (!modules || !peopleAvailable())
      return reply.status(409).send({ error: "modules_unavailable" });
    const params = personParams.safeParse(request.params);
    const query = z
      .object({
        before: z.string().max(600).optional(),
        result: z.enum(["all", "success", "failed"]).default("all"),
        limit: z.coerce.number().int().min(1).max(500).default(100),
      })
      .safeParse(request.query);
    if (!params.success || !query.success)
      return reply.status(400).send({ error: "invalid_request", message: "Check the filters." });
    try {
      const page = await modules.database.tenant(session.tenantId, (sql) =>
        personSignins(sql, session.tenantId, params.data.userId, {
          before: query.data.before ?? null,
          result: query.data.result,
          limit: query.data.limit,
        }),
      );
      await store.appendAudit(session.tenantId, session.actorId, {
        type: "person.signins.viewed",
        datasets: ["signIns"],
        occurredAt: new Date().toISOString(),
      });
      return personSigninPageSchema.parse(page);
    } catch (error) {
      if (error instanceof RangeError)
        return reply.status(400).send({ error: "invalid_request", message: "Reload the page." });
      throw error;
    }
  });

  app.get("/v1/people/:userId/directory", async (request, reply) => {
    const session = await sessionFor(request);
    const handle = readCookie(request, cookieName);
    if (config.mode !== "live" || !session || !handle)
      return reply.status(401).send({ error: "authentication_required" });
    const params = personParams.safeParse(request.params);
    if (!params.success)
      return reply.status(400).send({ error: "invalid_request", message: "Unknown account." });
    const { client, scopes } = delegatedGraph(session, handle);
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "person.directory.viewed",
      datasets: ["directory", "licenses", "signInActivity"],
      occurredAt: new Date().toISOString(),
    });
    return readPersonDirectory({
      key: { userId: params.data.userId },
      client,
      grantedScopes: scopes,
    });
  });

  app.post("/v1/findings/:id/state", async (request, reply) => {
    const session = await sessionFor(request);
    const params = z.object({ id: z.uuid() }).safeParse(request.params);
    const body = findingStateRequestSchema.safeParse(request.body);
    if (!modules || !session) return reply.status(409).send({ error: "modules_unavailable" });
    if (!params.success || !body.success)
      return reply
        .status(400)
        .send({ error: "invalid_request", message: "Choose open, acknowledged, or resolved." });
    const updated = await modules.database.tenant(session.tenantId, (sql) =>
      setFindingState(sql, session.tenantId, params.data.id, body.data.state, session.actorId),
    );
    if (!updated)
      return reply
        .status(404)
        .send({ error: "finding_not_found", message: "The finding is not available." });
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "finding.state.changed",
      findingId: updated.id,
      state: updated.state,
      occurredAt: new Date().toISOString(),
    });
    return findingListSchema.shape.items.element.parse(updated);
  });

  // Notification center ("attention") and the Microsoft service-health panel.
  // The enabled read pack that carries a permission. Several packs can share one permission
  // (AuditLog.Read.All, LicenseAssignment.Read.All); consent can only be asked for an enabled one.
  const featureFor = (permission: string) =>
    config.enabledFeatures.find((feature) => supportedReadFeatures[feature] === permission);
  // What each source can do right now. While collection runs, that is decided by the credential
  // collecting (the scheduler's own skipped list), not by whoever is looking; when it is stopped,
  // by the viewer's session, which is what starting collection would use.
  const sourceStates = (session: IdentitySession) => {
    const granted = new Set(
      session.scopes.filter((scope) => new Set(enabledPermissions()).has(scope)),
    );
    const status = modules?.scheduler.status();
    const collecting = status?.state === "collecting";
    const skipped = new Map((status?.skipped ?? []).map((entry) => [entry.module, entry.missing]));
    return (modules?.registry.list() ?? []).map((module) => {
      const permissions = modulePermissions(module);
      const notEnabled = module.packs.some(
        (pack) => !config.enabledFeatures.includes(pack as SupportedReadFeature),
      );
      const missing = collecting
        ? (skipped.get(module.id) ?? [])
        : permissions.filter((permission) => !granted.has(permission));
      return {
        module: module.id,
        title: module.title,
        collectors: module.collectors,
        state: notEnabled
          ? ("not_enabled" as const)
          : missing.length > 0
            ? ("missing_permission" as const)
            : collecting
              ? ("available" as const)
              : ("not_collecting" as const),
        missing,
      };
    });
  };
  // Problems that stop Prism365 from checking the environment, worked out from collection state
  // each time they are asked for. Sources the installation never enabled are not alerts.
  const systemAlerts = async (session: IdentitySession) => {
    if (!modules) return [];
    const status = modules.scheduler.status();
    const runs = await modules.database.tenant(session.tenantId, (sql) =>
      latestRuns(sql, session.tenantId),
    );
    const alerts: SystemAlert[] = [];
    const sources = sourceStates(session);
    if (status.state === "stopped" && sources.some((source) => source.state !== "not_enabled"))
      alerts.push({
        id: "collection.stopped",
        severity: status.lastStopReason === "credential_unavailable" ? "high" : "medium",
        title: "Prism365 is not collecting",
        detail:
          status.lastStopReason === "credential_unavailable"
            ? "The session collection ran on ended or lost access, so nothing is being checked. Sign in and start collecting again."
            : "Nothing is being checked until an administrator starts collecting.",
        observedAt: null,
        action: { label: "Start collecting", href: "view:signins" },
      });
    for (const source of sources) {
      if (source.state === "missing_permission") {
        const feature = featureFor(source.missing[0] ?? "");
        alerts.push({
          id: `permission.${source.module}`,
          severity: "medium",
          title: `${source.title} cannot be checked`,
          detail: `${status.state === "collecting" ? "The session Prism365 collects with" : "Your Microsoft session"} has not granted ${source.missing.join(", ")}. Grant it, and make sure the account has a role Microsoft accepts for it.`,
          observedAt: null,
          action: feature
            ? {
                label: "Grant access",
                href: `/api/v1/auth/login?feature=${encodeURIComponent(feature)}`,
              }
            : null,
        });
        continue;
      }
      if (source.state !== "available") continue;
      for (const collector of source.collectors) {
        const run = runs.find(
          (item) => item.module === source.module && item.collector === collector.id,
        );
        if (run?.status === "failed")
          alerts.push({
            id: `failed.${source.module}.${collector.id}`,
            severity: "medium",
            title: `The last ${collector.title.toLowerCase()} check failed`,
            detail: `It failed with ${run.errorCategory ?? "an unknown error"}. The next run retries.`,
            observedAt: run.startedAt,
            action: null,
          });
        else if (
          run &&
          Date.now() - Date.parse(run.startedAt) > collector.cadenceSeconds * 3_000 + 120_000
        )
          alerts.push({
            id: `stale.${source.module}.${collector.id}`,
            severity: "low",
            title: `${collector.title} has not been checked recently`,
            detail: `Prism365 checks it every ${Math.round(collector.cadenceSeconds / 60)} minutes while collecting.`,
            observedAt: run.startedAt,
            action: null,
          });
      }
    }
    return alerts;
  };
  const attentionRoute = (request: FastifyRequest) => {
    if (config.mode !== "live" || !modules) return null;
    return sessionFor(request);
  };

  app.get("/v1/attention", async (request, reply) => {
    if (config.mode === "demo")
      return attentionListSchema.parse({
        mode: "demo",
        items: [],
        system: [],
        sources: [],
        counts: { unread: 0, active: 0 },
      });
    const session = await attentionRoute(request);
    if (!session || !modules) return reply.status(401).send({ error: "authentication_required" });
    const query = z
      .object({
        filter: z.enum(["open", "unread", "all", "muted"]).default("open"),
        limit: z.coerce.number().int().min(1).max(500).default(100),
      })
      .safeParse(request.query);
    if (!query.success)
      return reply.status(400).send({ error: "invalid_request", message: "Check the filter." });
    const stored = await modules.database.tenant(session.tenantId, async (sql) => ({
      items: await listAttention(sql, session.tenantId, query.data),
      counts: await attentionCounts(sql, session.tenantId),
    }));
    return attentionListSchema.parse({
      mode: "live",
      ...stored,
      system: await systemAlerts(session),
      sources: sourceStates(session)
        .filter((source) => source.module !== SIGNINS_MODULE)
        .map(({ module, title, state, missing }) => ({ module, title, state, missing })),
    });
  });

  const attentionParams = z.object({ id: z.uuid() });
  app.post("/v1/attention/:id/review", async (request, reply) => {
    const session = await attentionRoute(request);
    const params = attentionParams.safeParse(request.params);
    const body = attentionReviewRequestSchema.safeParse(request.body);
    if (!session || !modules) return reply.status(401).send({ error: "authentication_required" });
    if (!params.success || !body.success)
      return reply
        .status(400)
        .send({ error: "invalid_request", message: "Choose new, reviewed or done." });
    const item = await modules.database.tenant(session.tenantId, (sql) =>
      setAttentionReview(sql, session.tenantId, params.data.id, body.data.state, session.actorId),
    );
    if (!item)
      return reply.status(404).send({ error: "not_found", message: "The item is not available." });
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "attention.reviewed",
      itemId: item.id,
      state: item.review.state,
      occurredAt: new Date().toISOString(),
    });
    return attentionItemSchema.parse(item);
  });

  app.post("/v1/attention/:id/mute", async (request, reply) => {
    const session = await attentionRoute(request);
    const params = attentionParams.safeParse(request.params);
    const body = attentionMuteRequestSchema.safeParse(request.body);
    if (!session || !modules) return reply.status(401).send({ error: "authentication_required" });
    if (!params.success || !body.success)
      return reply.status(400).send({ error: "invalid_request", message: "Choose muted or not." });
    const item = await modules.database.tenant(session.tenantId, (sql) =>
      setAttentionMuted(sql, session.tenantId, params.data.id, body.data.muted),
    );
    if (!item)
      return reply.status(404).send({ error: "not_found", message: "The item is not available." });
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "attention.muted",
      itemId: item.id,
      muted: item.muted,
      occurredAt: new Date().toISOString(),
    });
    return attentionItemSchema.parse(item);
  });

  app.post("/v1/attention/:id/read", async (request, reply) => {
    const session = await attentionRoute(request);
    const params = attentionParams.safeParse(request.params);
    if (!session || !modules) return reply.status(401).send({ error: "authentication_required" });
    if (!params.success) return reply.status(400).send({ error: "invalid_request" });
    const updated = await modules.database.tenant(session.tenantId, (sql) =>
      markAttentionRead(sql, session.tenantId, params.data.id),
    );
    return { updated };
  });

  app.post("/v1/attention/read-all", async (request, reply) => {
    const session = await attentionRoute(request);
    if (!session || !modules) return reply.status(401).send({ error: "authentication_required" });
    const updated = await modules.database.tenant(session.tenantId, (sql) =>
      markAttentionRead(sql, session.tenantId, null),
    );
    return { updated };
  });

  const serviceNames = (sql: SqlClient, tenantId: string) =>
    sql
      .query<{ name: string }>(
        "SELECT name FROM mod_service_health_service WHERE tenant_id=$1 ORDER BY name",
        [tenantId],
      )
      .then((result) => result.rows.map((row) => row.name));

  app.get("/v1/settings/attention", async (request, reply) => {
    const session = await attentionRoute(request);
    if (!session || !modules) return reply.status(401).send({ error: "authentication_required" });
    const settings = await modules.database.tenant(session.tenantId, async (sql) => ({
      preferences: await readAttentionPreferences(sql, session.tenantId),
      services: await serviceNames(sql, session.tenantId),
    }));
    return attentionSettingsSchema.parse(settings);
  });

  app.post("/v1/settings/attention", async (request, reply) => {
    const session = await attentionRoute(request);
    if (!session || !modules) return reply.status(401).send({ error: "authentication_required" });
    const body = attentionPreferencesSchema.safeParse(request.body);
    if (!body.success)
      return reply.status(400).send({ error: "invalid_request", message: "Check the settings." });
    const settings = await modules.database.tenant(session.tenantId, async (sql) => ({
      preferences: await writeAttentionPreferences(
        sql,
        session.tenantId,
        body.data,
        session.actorId,
      ),
      services: await serviceNames(sql, session.tenantId),
    }));
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "attention.settings.changed",
      occurredAt: new Date().toISOString(),
    });
    return attentionSettingsSchema.parse(settings);
  });

  app.get("/v1/modules/service.health/status", async (request, reply) => {
    if (config.mode === "demo")
      return serviceHealthStatusSchema.parse({
        state: "not_enabled",
        detail: "Demo mode does not read Microsoft service health.",
        lastSuccess: null,
        lastAttempt: null,
        services: [],
        incidents: [],
        messages: { state: "not_enabled", lastSuccess: null },
      });
    const session = await attentionRoute(request);
    if (!session || !modules) return reply.status(401).send({ error: "authentication_required" });
    const sources = sourceStates(session);
    const source = (id: string) => sources.find((item) => item.module === id);
    const flags = (id: string) => ({
      enabled: source(id) !== undefined && source(id)?.state !== "not_enabled",
      granted: source(id)?.state === "available" || source(id)?.state === "not_collecting",
    });
    const status = await modules.database.tenant(session.tenantId, async (sql) =>
      serviceHealthStatus(sql, session.tenantId, {
        now: new Date(),
        collecting: modules.scheduler.status().state === "collecting",
        runs: await latestRuns(sql, session.tenantId),
        health: flags(HEALTH_MODULE),
        messages: flags(MESSAGES_MODULE),
      }),
    );
    return serviceHealthStatusSchema.parse(status);
  });

  // Audits (TL-025): checks an administrator runs on demand. Each run reads Microsoft with the
  // delegated session of whoever started it, and its rows are kept for history and export.
  const packPermission = (id: string) =>
    permissionPacks.find((pack) => pack.id === id)?.permission ?? null;
  // An audit is available when every pack it requires is enabled (the pack itself, not just a
  // permission another enabled pack happens to share) and the scopes granted their permissions.
  // The scopes are a session's, or the collecting credential's for a scheduled run.
  const auditAvailabilityFor = (audit: Audit, scopes: readonly string[]): AuditAvailability => {
    const enabledPacks = new Set<string>(config.enabledFeatures);
    const notEnabledPacks = audit.packs.filter((pack) => !enabledPacks.has(pack));
    if (notEnabledPacks.length > 0)
      return {
        state: "not_enabled",
        missing: [...new Set(notEnabledPacks.map((pack) => packPermission(pack) ?? pack))],
        grantFeature: null,
        enable: notEnabledPacks,
      };
    const { permissions } = auditDefinition(audit);
    const granted = new Set(auditReadScope(audit, config.enabledFeatures, scopes).granted);
    const missing = permissions.filter((permission) => !granted.has(permission));
    if (missing.length > 0)
      return {
        state: "missing_permission",
        missing,
        grantFeature: featureFor(missing[0] ?? "") ?? null,
        enable: [],
      };
    return { state: "available", missing: [], grantFeature: null, enable: [] };
  };
  const auditAvailability = (audit: Audit, session: IdentitySession) =>
    auditAvailabilityFor(audit, session.scopes);
  // Audits read whole directories, so they get the background reader's budgets.
  const auditGraphClient = (getAccessToken: () => Promise<string>) =>
    new GraphHttpClient({
      requestTimeoutMilliseconds: 60_000,
      maxResponseBytes: 8_388_608,
      maxRetries: 3,
      maxRetryWaitMilliseconds: 30_000,
      getAccessToken,
      ...(options.graphFetch ? { fetchImplementation: options.graphFetch } : {}),
    });
  const auditGraph = (session: IdentitySession, handle: string) => {
    if (!identity) throw new Error("Live mode is not configured.");
    const permitted = new Set(enabledPermissions());
    const scopes = session.scopes.filter((scope) => permitted.has(scope));
    const client = auditGraphClient(() => identity.accessToken(handle, scopes));
    return { client, scopes };
  };
  // Scheduled audits (docs/audits.md, section 15): due audits run one at a time on the sign-in
  // collection uses, read from the live credential here, so none runs while collection is off.
  const auditSchedule =
    audits && modules
      ? new AuditScheduler({
          runner: audits,
          database: modules.database,
          audits: builtInAudits,
          credential: () => collecting,
          enabledPacks: config.enabledFeatures,
          availability: auditAvailabilityFor,
          graphClient: auditGraphClient,
          ledger: (tenantId, actorId, event) => store.appendAudit(tenantId, actorId, event),
          ...options.auditSchedule,
        })
      : null;
  if (auditSchedule && modules) {
    auditSchedule.start();
    // Once shortly after collection starts, so a due audit does not wait a whole tick. When
    // collection stops or passes to another administrator, a scheduled run reading on the old
    // sign-in stops at once.
    modules.scheduler.subscribe((event) => {
      if (event.type === "collection.bound") auditSchedule.soon();
      if (event.type === "collection.unbound") auditSchedule.collectionEnded();
    });
  }
  // Alerts: rules over findings, audit results and collection, sent to the operator's own mail
  // relay or a webhook. Off every request path: the sink only buffers, and audit and scheduler
  // events only start work that nothing waits for.
  const delivery =
    modules && identity
      ? new DeliveryService({
          tenantId: identity.tenantId,
          database: modules.database,
          store,
          origin: config.origin,
          config: config.delivery ?? { email: { smtp: null, domains: [] }, webhookPorts: [] },
          encryptionKey: config.encryptionKey ?? null,
          collection: () => modules.scheduler.status(),
          modules: () =>
            modules.registry.list().map((module) => ({ id: module.id, title: module.title })),
          timeZone: config.signinsAudit?.timeZone ?? "UTC",
          // The This week digest, read the way its page reads it, for the weekly summary.
          weeklyDigest: {
            build: (sql, now) =>
              buildDigest(sql, {
                tenantId: identity.tenantId,
                now,
                days: 7,
                collecting: modules.scheduler.status().state === "collecting",
                modules: modules.registry.list().map((module) => ({
                  id: module.id,
                  title: module.title,
                  collectors: module.collectors.map(({ id, title }) => ({ id, title })),
                })),
                unavailable: digestUnavailable(modules.registry, config.enabledFeatures),
              }),
            datasets: () =>
              digestSectionsRead(digestUnavailable(modules.registry, config.enabledFeatures)),
          },
          ...options.delivery,
        })
      : null;
  const detachDelivery: (() => void)[] = [];
  if (delivery && modules) {
    detachDelivery.push(
      modules.scheduler.addSink(delivery.sink),
      modules.scheduler.subscribe((event) => delivery.onSchedulerEvent(event)),
    );
    if (audits) detachDelivery.push(audits.subscribe((event) => delivery.onAuditStored(event)));
    delivery.start();
  }
  const auditsUnavailable = {
    error: "audits_unavailable",
    message: "Audits need live mode with module storage, and a signed-in administrator.",
  };
  const runParams = z.object({ runId: z.uuid() });

  app.get("/v1/audits", async (request, reply) => {
    if (config.mode === "demo")
      return auditCatalogSchema.parse({
        mode: "demo",
        audits: builtInAudits.map((audit) => ({
          definition: auditDefinition(audit),
          availability: { state: "available", missing: [], grantFeature: null, enable: [] },
          lastRun: null,
          runningId: null,
        })),
      });
    const session = await sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    if (!audits || !modules) return reply.status(409).send(auditsUnavailable);
    await audits.recover(session.tenantId);
    const latest = await modules.database.tenant(session.tenantId, (sql) =>
      latestAuditRuns(sql, session.tenantId),
    );
    return auditCatalogSchema.parse({
      mode: "live",
      audits: builtInAudits.map((audit) => {
        const id = audit.definition.id;
        return {
          definition: auditDefinition(audit),
          availability: auditAvailability(audit, session),
          lastRun: latest.find((run) => run.auditId === id) ?? null,
          runningId: audits.runningId(session.tenantId, id),
        };
      }),
    });
  });

  app.post("/v1/audits/:auditId/runs", async (request, reply) => {
    const session = await sessionFor(request);
    const handle = readCookie(request, cookieName);
    if (!session || !handle) return reply.status(401).send({ error: "authentication_required" });
    if (!audits || !modules) return reply.status(409).send(auditsUnavailable);
    const params = z.object({ auditId: auditIdSchema }).safeParse(request.params);
    const audit = params.success ? findAudit(params.data.auditId) : undefined;
    if (!audit)
      return reply.status(404).send({ error: "unknown_audit", message: "There is no such audit." });
    const availability = auditAvailability(audit, session);
    if (availability.state !== "available")
      return reply.status(403).send({
        error: availability.state === "not_enabled" ? "feature_not_enabled" : "missing_consent",
        message: auditUnavailableMessage(availability),
        availability,
      });
    const id = audit.definition.id;
    const { client, scopes } = auditGraph(session, handle);
    const read = auditReadScope(audit, config.enabledFeatures, scopes);
    const definition = auditDefinition(audit);
    // The ledger names what this run will read: required permissions, and optional ones the
    // session granted for an enabled pack (license names, sign-in dates).
    const datasets = [
      ...definition.permissions,
      ...definition.optionalPermissions.filter((permission) => read.granted.includes(permission)),
    ];
    const result = await audits.start({
      tenantId: session.tenantId,
      actor: { id: session.actorId, name: session.displayName ?? null },
      audit,
      client,
      packs: read.packs,
      granted: read.granted,
      beforeRead: (started) =>
        store.appendAudit(session.tenantId, session.actorId, {
          type: "audit.run.started",
          auditId: id,
          runId: started.id,
          datasets,
          occurredAt: new Date().toISOString(),
        }),
    });
    let run = result.run;
    if (!run) {
      const detail = await modules.database.tenant(session.tenantId, (sql) =>
        getRun(sql, session.tenantId, result.runId),
      );
      if (detail) {
        const { rows: _rows, ...summary } = detail;
        run = summary;
      }
    }
    if (!run)
      return reply.status(409).send({ error: "run_unavailable", message: "Reload the page." });
    // A Run that joined a scheduled run in progress shares its end, and a scheduled run stops
    // when collection stops: the administrator is told, so a stopped result is no surprise.
    const message =
      !result.started && run.trigger === "schedule"
        ? "This audit was already running on its schedule, so your Run joined that run. It stops if collection stops; run it again then."
        : null;
    return reply
      .status(result.started ? 202 : 200)
      .send(auditStartSchema.parse({ run, started: result.started, message }));
  });

  app.get("/v1/audits/:auditId/runs", async (request, reply) => {
    const session = await sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    if (!audits || !modules) return reply.status(409).send(auditsUnavailable);
    const params = z.object({ auditId: auditIdSchema }).safeParse(request.params);
    if (!params.success || !findAudit(params.data.auditId))
      return reply.status(404).send({ error: "unknown_audit", message: "There is no such audit." });
    await audits.recover(session.tenantId);
    const runs = await modules.database.tenant(session.tenantId, (sql) =>
      listRuns(sql, session.tenantId, params.data.auditId),
    );
    return auditRunListSchema.parse({ runs });
  });

  const loadRun = async (session: IdentitySession, runId: string) => {
    if (!modules || !audits) return null;
    await audits.recover(session.tenantId);
    return modules.database.tenant(session.tenantId, (sql) => getRun(sql, session.tenantId, runId));
  };

  app.get("/v1/audit-runs/:runId", async (request, reply) => {
    const session = await sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    if (!audits || !modules) return reply.status(409).send(auditsUnavailable);
    const params = runParams.safeParse(request.params);
    if (!params.success)
      return reply.status(400).send({ error: "invalid_request", message: "Unknown audit run." });
    const run = await loadRun(session, params.data.runId);
    if (!run)
      return reply.status(404).send({ error: "run_not_found", message: "That run was deleted." });
    const current = findAudit(run.auditId);
    if (current && run.auditVersion !== current.definition.version)
      run.notes = [
        `Made by version ${run.auditVersion} of this audit; the table now uses version ${current.definition.version}, so some columns may be empty. Run it again for a current result.`,
        ...run.notes,
      ].slice(0, 20);
    // Polling a run in progress returns no rows, so only a finished result counts as viewed.
    if (run.status !== "running")
      await store.appendAudit(session.tenantId, session.actorId, {
        type: "audit.run.viewed",
        auditId: run.auditId,
        runId: run.id,
        occurredAt: new Date().toISOString(),
      });
    return auditRunDetailSchema.parse(run);
  });

  app.get("/v1/audit-runs/:runId/export.csv", async (request, reply) => {
    const session = await sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    if (!audits || !modules) return reply.status(409).send(auditsUnavailable);
    const params = runParams.safeParse(request.params);
    const query = auditExportParamsSchema.safeParse(request.query ?? {});
    if (!params.success || !query.success)
      return reply.status(400).send({ error: "invalid_request", message: "Check the filters." });
    const run = await loadRun(session, params.data.runId);
    const audit = run ? findAudit(run.auditId) : undefined;
    if (!run || !audit || run.status !== "succeeded")
      return reply
        .status(404)
        .send({ error: "run_not_found", message: "That run has no results." });
    const definition = auditDefinition(audit);
    // Export what the page shows. The handoff export parses the same way (audit-export.ts).
    const rows = queryAuditRows(definition, run.rows, auditExportQuery(query.data));
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "audit.run.exported",
      auditId: run.auditId,
      runId: run.id,
      rows: rows.length,
      occurredAt: new Date().toISOString(),
    });
    const day = run.startedAt.slice(0, 10);
    return reply
      .header("Content-Type", "text/csv; charset=utf-8")
      .header(
        "Content-Disposition",
        // A limited result says so in its name, so a file passed on is never mistaken for the
        // whole tenant.
        `attachment; filename="prism365-${run.auditId}-${day}${scopeFileSuffix(run.scope)}.csv"`,
      )
      .send(auditCsv(definition, rows));
  });

  // The audit scope (docs/audits.md): which account domains the account audits cover. The
  // domains to choose from come from the stored results of the audits that use it, so reading
  // them makes no call to Microsoft.
  app.get("/v1/settings/audits", async (request, reply) => {
    const session = await sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    if (!audits || !modules) return reply.status(409).send(auditsUnavailable);
    const settings = await modules.database.tenant(session.tenantId, (sql) =>
      readAuditSettings(sql, session.tenantId, builtInAudits),
    );
    return auditSettingsSchema.parse(settings);
  });

  app.post("/v1/settings/audits", async (request, reply) => {
    const session = await sessionFor(request);
    if (!session) return reply.status(401).send({ error: "authentication_required" });
    if (!audits || !modules) return reply.status(409).send(auditsUnavailable);
    const body = auditScopeSchema.safeParse(request.body);
    if (!body.success)
      return reply.status(400).send({ error: "invalid_request", message: "Check the domains." });
    // The ledger entry comes first, so no scope change is ever unrecorded. It holds how many
    // domains, never which.
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "audits.scope.changed",
      domains: body.data.domains.length,
      occurredAt: new Date().toISOString(),
    });
    const settings = await modules.database.tenant(session.tenantId, async (sql) => {
      await writeAuditScope(sql, session.tenantId, body.data, session.actorId);
      return readAuditSettings(sql, session.tenantId, builtInAudits);
    });
    return auditSettingsSchema.parse(settings);
  });

  // Features with their own route files (routes/). Each checks the session itself, and gets
  // nothing that could act as anyone else: see routes/context.ts.
  const routeGraph = (
    session: IdentitySession,
    handle: string,
    packs: readonly GraphFeatureId[],
  ) => {
    const enabled = new Set<string>(config.enabledFeatures);
    const off = packs.filter((pack) => !enabled.has(pack));
    if (off.length > 0) throw new RouteBoundaryError(`Read packs not enabled: ${off.join(", ")}.`);
    const { client, scopes } = auditGraph(session, handle);
    const carried = new Set<string | null>(packs.map((pack) => packPermission(pack)));
    return {
      graph: scopeGraph(client, packs),
      granted: scopes.filter((scope) => carried.has(scope)),
    };
  };
  const credential = config.identity?.credential ?? null;
  const routes: RouteContext = {
    settings: {
      mode: config.mode,
      origin: config.origin,
      production: config.production,
      configured: Boolean(config.identity),
      missingConfiguration: [...config.missingConfiguration],
      enabledFeatures: [...config.enabledFeatures],
      credentialKind: !credential
        ? null
        : "clientCertificate" in credential
          ? "certificate"
          : "secret",
      certificate: config.certificate ? { ...config.certificate } : null,
      durableStorage: Boolean(config.databaseUrl || config.moduleDataDir),
      durableSessions: store.durable,
      handoffRetentionDays: config.handoffRetentionDays ?? DEFAULT_HANDOFF_RETENTION_DAYS,
    },
    store,
    modules: modules
      ? {
          registry: modules.registry,
          database: modules.database,
          status: () => modules.scheduler.status(),
        }
      : null,
    audits,
    sessionFor,
    sessionHandle: (request) => readCookie(request, cookieName),
    enabledPermissions,
    sessionGraph: routeGraph,
    auditAvailability,
    collectingAdmin: () => (collecting ? { actorId: collecting.actorId } : null),
    delivery,
  };
  registerChangeRoutes(app, routes);
  registerAuditInsightRoutes(app, routes);
  registerDigestRoutes(app, routes);
  registerSetupRoutes(app, routes);
  registerHandoffRoutes(app, routes);
  registerCanaryRoutes(app, routes);
  registerDeliveryRoutes(app, routes);

  app.get("/v1/dashboard/summary", async (_request, reply) => {
    if (config.mode === "demo") return demoDashboard;
    return reply.status(409).send({
      error: "dashboard_not_collected",
      message: "Use People to look up current tenant data. Tenant-wide collection is not enabled.",
    });
  });

  app.get("/v1/platform/capabilities", async () =>
    capabilitySummarySchema.parse({
      dataMode: config.mode,
      writesEnabled: false,
      tenantConnectionEnabled: Boolean(identity),
      permissionPacks,
      cloudProfiles: Object.values(microsoftCloudProfiles),
      ai: {
        analysisEnabled: aiAvailable,
        providerConfigured: Boolean(config.aiModel),
        writeToolsEnabled: false,
        supportedDeployments: ["local", "tenantHosted", "external"],
      },
    }),
  );

  async function investigation(
    email: string,
    request: FastifyRequest,
    now?: Date,
  ): Promise<UserInvestigation> {
    if (config.mode === "demo")
      return createDemoUserInvestigation({ email, ...(now ? { now } : {}) });
    const session = await sessionFor(request);
    const handle = readCookie(request, cookieName);
    if (!session || !identity || !handle) throw new Error("Authentication required.");
    const permitted = new Set<string>([
      "User.Read",
      ...config.enabledFeatures.map(
        (feature) => supportedReadFeatures[feature as SupportedReadFeature],
      ),
    ]);
    const scopes = session.scopes.filter((scope) => permitted.has(scope));
    const client = new GraphHttpClient({
      // An interactive lookup waits longer than the default for Microsoft's slow sign-in log.
      requestTimeoutMilliseconds: 30_000,
      getAccessToken: () => identity.accessToken(handle, scopes),
      ...(options.graphFetch ? { fetchImplementation: options.graphFetch } : {}),
    });
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "investigation.requested",
      datasets: ["directory", "licenses", "signIns"],
      occurredAt: new Date().toISOString(),
    });
    return investigateUser({ email, client, grantedScopes: scopes, ...(now ? { now } : {}) });
  }

  app.get("/v1/ai/status", async () => ({
    available: aiAvailable,
    provider: config.aiModel ? `ollama-local:${config.aiModel}` : null,
    datasets: ["directory", "licenses", "signIns"],
    externalProcessing: false,
    reason: aiAvailable
      ? null
      : "AI requires a connected tenant, encrypted PostgreSQL state, and AI_LOCAL_MODEL on the server.",
  }));

  app.post("/v1/ai/grants", async (request, reply) => {
    const session = await sessionFor(request);
    if (!aiAvailable || !session || !config.aiModel)
      return reply.status(409).send({
        error: "ai_not_configured",
        message: "Configure the local AI provider and durable storage first.",
      });
    const parsed = z
      .object({
        datasets: z
          .array(z.enum(["directory", "licenses", "signIns"]))
          .min(1)
          .max(3),
      })
      .safeParse(request.body);
    if (!parsed.success)
      return reply
        .status(400)
        .send({ error: "invalid_request", message: "Select the datasets to authorize." });
    const now = new Date();
    const grant = aiAccessGrantSchema.parse({
      grantId: randomUUID(),
      tenantId: session.tenantId,
      providerId: `ollama-local:${config.aiModel}`,
      datasets: [...new Set(parsed.data.datasets)],
      purposes: ["securityInvestigation"],
      allowExternalProcessing: false,
      allowDirectIdentifiers: false,
      allowRawRecords: false,
      maxLookbackDays: 7,
      maxEvidenceRecords: 102,
      approvedByActorId: session.actorId,
      approvedAt: now.toISOString(),
      expiresAt: new Date(now.getTime() + 60 * 60_000).toISOString(),
    });
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "ai.grant.approved",
      grantId: grant.grantId,
      datasets: grant.datasets,
      expiresAt: grant.expiresAt,
    });
    await store.set(
      session.tenantId,
      "aiGrant",
      `${session.actorId}:${grant.grantId}`,
      grant,
      new Date(grant.expiresAt),
    );
    return { grantId: grant.grantId, datasets: grant.datasets, expiresAt: grant.expiresAt };
  });

  app.delete("/v1/ai/grants/:id", async (request, reply) => {
    const session = await sessionFor(request);
    const parsed = z.object({ id: z.uuid() }).safeParse(request.params);
    if (!session || !parsed.success) return reply.status(400).send({ error: "invalid_request" });
    await store.delete(session.tenantId, "aiGrant", `${session.actorId}:${parsed.data.id}`);
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "ai.grant.revoked",
      grantId: parsed.data.id,
    });
    return { revoked: true };
  });

  app.post("/v1/ai/investigations", async (request, reply) => {
    const session = await sessionFor(request);
    if (!aiAvailable || !session || !config.aiModel || !identity)
      return reply
        .status(409)
        .send({ error: "ai_not_configured", message: "The local AI provider is not ready." });
    const parsed = z
      .object({
        email: investigationEmailSchema,
        question: z.string().trim().min(1).max(4_000),
        grantId: z.uuid(),
      })
      .safeParse(request.body);
    if (!parsed.success)
      return reply.status(400).send({
        error: "invalid_request",
        message: "Enter a question and select an active AI dataset grant.",
      });
    const handle = readCookie(request, cookieName);
    const value = aiAccessGrantSchema.safeParse(
      await store.get(session.tenantId, "aiGrant", `${session.actorId}:${parsed.data.grantId}`),
    );
    if (!value.success)
      return reply.status(403).send({
        error: "grant_unavailable",
        message: "The AI dataset grant has expired or was revoked.",
      });
    const observedAt = new Date();
    const timeRange = {
      start: new Date(observedAt.getTime() - 7 * 86_400_000).toISOString(),
      end: observedAt.toISOString(),
    };
    const provider = options.aiProvider ?? new LocalOllamaProvider({ model: config.aiModel });
    const cancellation = new AbortController();
    request.raw.once("aborted", () => cancellation.abort());
    reply.raw.once("close", () => {
      if (!reply.raw.writableFinished) cancellation.abort();
    });
    const reader = {
      read: async (query: AiEvidenceQuery) =>
        investigationEvidence(
          await investigation(parsed.data.email, request, observedAt),
          session,
          query,
        ),
    };
    const broker = new AiAnalysisBroker({
      provider,
      evidenceReader: reader,
      grantRepository: {
        getActiveGrant: (query) =>
          store.get(query.tenantId, "aiGrant", `${query.actorId}:${query.grantId}`),
      },
      actorAuthorizer: {
        authorize: async (query) => {
          const current = handle ? await identity.session(handle) : null;
          const scopeByDataset = {
            directory: "User.Read.All",
            licenses: "LicenseAssignment.Read.All",
            signIns: "AuditLog.Read.All",
          };
          return Boolean(
            current &&
              current.tenantId === query.actor.tenantId &&
              current.actorId === query.actor.actorId &&
              query.datasets.every(
                (dataset) =>
                  dataset in scopeByDataset &&
                  current.scopes.includes(scopeByDataset[dataset as keyof typeof scopeByDataset]),
              ),
          );
        },
      },
      auditSink: {
        append: (event) => store.appendAudit(event.tenantId, event.actorId, { ...event }),
      },
    });
    try {
      return await broker.analyze(
        {
          tenantId: session.tenantId,
          actorId: session.actorId,
          question: parsed.data.question,
          purpose: "securityInvestigation",
          datasets: value.data.datasets,
          timeRange,
        },
        {
          actor: { tenantId: session.tenantId, actorId: session.actorId },
          grantId: value.data.grantId,
          signal: cancellation.signal,
        },
      );
    } catch {
      return reply.status(503).send({
        error: "analysis_unavailable",
        message:
          "AI analysis was not completed. Check provider readiness, dataset access, and the access ledger.",
      });
    }
  });

  app.addHook("onClose", async () => {
    // No alert is taken from here on; deliveries still waiting are recorded as interrupted
    // while the database is open.
    for (const detach of detachDelivery) detach();
    await delivery?.stop();
    // No scheduled audit starts from here on.
    auditSchedule?.stop();
    // Audits first: an interrupted run is recorded before the database closes.
    await audits?.close();
    // A scheduled pass waiting on a run sees it stopped and returns.
    await auditSchedule?.settled();
    if (options.modules) {
      await options.modules.scheduler.stop();
      await options.modules.database.close();
    }
    await store.close();
  });

  if (identity) {
    const timer = setInterval(() => {
      void store.purgeExpired(identity.tenantId).catch(() => {});
    }, 60_000);
    timer.unref();
    app.addHook("onReady", async () => store.purgeExpired(identity.tenantId));
    app.addHook("onClose", async () => clearInterval(timer));
  }

  if (options.webRoot) {
    app.register(staticFiles, { root: options.webRoot, prefix: "/", wildcard: false });
  }

  return app;
}

// Categories only: never the provider message, authorization code, token, or account identifier.
export function signInFailureCategory(error: unknown): string {
  return error instanceof LoginRejectedError ? error.reason : "token_exchange_failed";
}

function investigationEvidence(
  data: UserInvestigation,
  session: IdentitySession,
  query: AiEvidenceQuery,
): AiEvidenceRecord[] {
  if (query.tenantId !== session.tenantId || query.actorId !== session.actorId)
    throw new Error("Evidence actor mismatch.");
  const common = {
    tenantId: session.tenantId,
    source: "Microsoft Graph v1.0",
    sensitivity: "personLinked" as const,
    directIdentifiers: {
      email: data.targetEmail,
      displayName: data.profile.data?.displayName ?? "",
    },
    limitations: data.limitations,
    raw: false,
  };
  const records: AiEvidenceRecord[] = [];
  const states = {
    directory: data.profile.state,
    licenses: data.licenses.state,
    signIns: data.signIns.state,
  };
  if (
    query.datasets.some(
      (dataset) => !(dataset in states) || states[dataset as keyof typeof states] !== "available",
    )
  )
    throw new Error("A requested evidence dataset is unavailable.");
  if (query.datasets.includes("directory") && data.profile.state === "available")
    records.push({
      ...common,
      dataset: "directory",
      evidenceId: "directory-profile",
      observedAt: data.generatedAt,
      fields: { userType: data.profile.data?.userType ?? "unknown", userCount: 1 },
    });
  if (query.datasets.includes("licenses") && data.licenses.state === "available")
    records.push({
      ...common,
      dataset: "licenses",
      evidenceId: "license-summary",
      observedAt: data.generatedAt,
      fields: {
        licenseCount: data.licenses.data.length,
        servicePlanCount: data.licenses.data.reduce(
          (total, license) => total + license.servicePlans.length,
          0,
        ),
        complete: true,
      },
    });
  if (query.datasets.includes("signIns") && data.signIns.state === "available") {
    if (data.signIns.data.length === 0)
      records.push({
        ...common,
        dataset: "signIns",
        evidenceId: "empty-signin-window",
        observedAt: data.generatedAt,
        fields: { successCount: 0, failureCount: 0, complete: data.signIns.complete },
      });
    for (const event of data.signIns.data)
      records.push({
        ...common,
        dataset: "signIns",
        evidenceId: event.id,
        observedAt: event.createdDateTime,
        fields: {
          status: event.status.errorCode === 0 ? "success" : "failure",
          errorCode: event.status.errorCode,
          complete: data.signIns.complete,
          ...(event.isInteractive === null ? {} : { isInteractive: event.isInteractive }),
        },
      });
  }
  return records.slice(0, query.limit);
}

import { randomUUID } from "node:crypto";
import staticFiles from "@fastify/static";
import {
  AiAnalysisBroker,
  LocalOllamaProvider,
  type AiEvidenceQuery,
  type AiProvider,
} from "@prism365/ai";
import {
  createAutomationPreview,
  listAutomationWorkflows,
  inspectAutomationPackage,
  summarizeAutomationPackage,
} from "@prism365/automation";
import {
  aiAccessGrantSchema,
  automationPackageSummarySchema,
  automationPackageListSchema,
  investigationEmailSchema,
  automationPreviewRequestSchema,
  capabilitySummarySchema,
  type AiEvidenceRecord,
  type UserInvestigation,
  sessionSchema,
} from "@prism365/contracts";
import {
  createDemoUserInvestigation,
  GraphHttpClient,
  investigateUser,
  microsoftCloudProfiles,
  permissionPacks,
  UserInvestigationError,
} from "@prism365/graph";
import { MemoryStateStore, type StateStore } from "@prism365/storage";
import Fastify, { type FastifyRequest } from "fastify";
import { z } from "zod";
import { type AppConfig, type SupportedReadFeature, supportedReadFeatures } from "./config.ts";
import { demoDashboard } from "./demo-data.ts";
import { digest, equalSecret, IdentityService, type IdentitySession } from "./identity.ts";

export interface AppOptions {
  config?: AppConfig;
  store?: StateStore;
  identity?: IdentityService;
  graphFetch?: typeof fetch;
  webRoot?: string;
  aiProvider?: AiProvider;
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

  const app = Fastify({
    logger: false,
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
    const pathname = request.url.split("?")[0];
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
        feature: z
          .enum(["directory.users.read", "user.licenses.read", "audit.signins.read"])
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
    if (!parsed.success) return reply.redirect(`${config.origin}/?auth=failed`);
    try {
      const result = await identity.finishLogin(
        parsed.data.code,
        parsed.data.state,
        readCookie(request, oauthCookieName) ?? "",
      );
      const oldHandle = readCookie(request, cookieName);
      const oldSession = await sessionFor(request);
      if (oldHandle && oldSession) await identity.logout(oldHandle, oldSession);
      reply.header("Set-Cookie", [
        cookie(oauthCookieName, "", 0),
        cookie(cookieName, result.handle, 8 * 60 * 60),
      ]);
      return reply.redirect(config.origin);
    } catch {
      return reply.redirect(`${config.origin}/?auth=failed`);
    }
  });

  app.post("/v1/auth/logout", async (request, reply) => {
    const handle = readCookie(request, cookieName);
    const session = await sessionFor(request);
    if (identity && handle && session) await identity.logout(handle, session);
    reply.header("Set-Cookie", cookie(cookieName, "", 0));
    return { signedOut: true };
  });

  app.get("/v1/dashboard/summary", async (_request, reply) => {
    if (config.mode === "demo") return demoDashboard;
    return reply.status(409).send({
      error: "dashboard_not_collected",
      message:
        "Use User Intelligence to inspect current tenant data. Tenant-wide collection is not enabled.",
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

  app.get("/v1/automation/workflows", async () => listAutomationWorkflows());

  app.post("/v1/automation/previews", async (request, reply) => {
    if (config.mode !== "demo")
      return reply.status(409).send({
        error: "live_automation_unavailable",
        message:
          "Live lifecycle automation requires a reviewed write identity and isolated runner.",
      });
    const parsed = automationPreviewRequestSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send({
        error: "invalid_request",
        message: "Choose an approved workflow and enter a valid work email address.",
      });
    }

    try {
      return createAutomationPreview(parsed.data);
    } catch (error) {
      if (error instanceof Error && error.message === "Unknown lifecycle workflow.") {
        return reply.status(404).send({
          error: "workflow_not_found",
          message: "The selected workflow is not available.",
        });
      }
      throw error;
    }
  });

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

  app.get("/v1/investigations/user", async (request, reply) => {
    const query = z.object({ email: investigationEmailSchema }).safeParse(request.query);
    if (!query.success)
      return reply
        .status(400)
        .send({ error: "invalid_input", message: "Enter a valid work email address." });
    try {
      return await investigation(query.data.email, request);
    } catch (error) {
      if (error instanceof UserInvestigationError)
        return reply.status(error.status).send({ error: error.code, message: error.message });
      return reply.status(503).send({
        error: "investigation_unavailable",
        message: "The tenant session or data source is unavailable. Sign in again and retry.",
      });
    }
  });

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

  app.addHook("onClose", async () => store.close());

  app.get("/v1/automation/packages", async () => {
    const values = identity ? await store.list(identity.tenantId, "automationPackage") : [];
    const schema = z.object({ summary: automationPackageSummarySchema });
    return automationPackageListSchema.parse({
      importEnabled: Boolean(identity && store.durable),
      executionEnabled: false,
      items: values.map((value) => schema.parse(value).summary),
    });
  });

  app.post("/v1/automation/packages", async (request, reply) => {
    const session = await sessionFor(request);
    if (!session || !store.durable)
      return reply.status(409).send({
        error: "import_not_ready",
        message:
          "Private script imports require an authenticated tenant and encrypted PostgreSQL storage.",
      });
    try {
      const record = inspectAutomationPackage(request.body, {
        tenantId: session.tenantId,
        actorId: session.actorId,
      });
      const summary = summarizeAutomationPackage(record);
      await store.appendAudit(session.tenantId, session.actorId, {
        type: "automation.package.imported",
        packageId: record.id,
        digest: record.digest,
        trust: "unreviewed",
      });
      await store.set(
        session.tenantId,
        "automationPackage",
        record.id,
        { summary, record },
        new Date(Date.now() + 30 * 86_400_000),
      );
      return reply.status(201).send(summary);
    } catch {
      return reply.status(400).send({
        error: "package_rejected",
        message:
          "Use a valid PowerShell 7 package with a fixed entrypoint, safe file paths, and no direct network hosts. Imports expire after 30 days and require review before execution.",
      });
    }
  });

  app.delete("/v1/automation/packages/:id", async (request, reply) => {
    const session = await sessionFor(request);
    const parsed = z.object({ id: z.uuid() }).safeParse(request.params);
    if (!session || !parsed.success) return reply.status(400).send({ error: "invalid_request" });
    const existing = await store.get(session.tenantId, "automationPackage", parsed.data.id);
    if (!existing)
      return reply
        .status(404)
        .send({ error: "package_not_found", message: "The private package is not available." });
    await store.appendAudit(session.tenantId, session.actorId, {
      type: "automation.package.deleted",
      packageId: parsed.data.id,
    });
    await store.delete(session.tenantId, "automationPackage", parsed.data.id);
    return { deleted: true };
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

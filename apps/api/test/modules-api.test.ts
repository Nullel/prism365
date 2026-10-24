import { GraphHttpClient } from "@prism365/graph";
import {
  ModuleRegistry,
  writeCollectionIntent,
  readCollectionIntent,
  ModuleScheduler,
  type PrismModule,
  upsertFinding,
} from "@prism365/modules";
import {
  MemoryStateStore,
  PGliteModuleDatabase,
  type SqlClient,
  type TenantOptions,
} from "@prism365/storage";
import { afterEach, describe, expect, it, vi } from "vitest";
import { z } from "zod";

// Each test opens PGlite with every built-in module's migrations; under a loaded machine that
// cold start alone can pass the default 5 s budget.
vi.setConfig({ testTimeout: 30_000 });
import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig, SupportedReadFeature } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import type { ModuleRuntime } from "../src/modules.ts";

const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";

const closers: (() => Promise<void>)[] = [];
afterEach(async () => {
  while (closers.length) await closers.pop()?.();
});

function probe(counter: { runs: number }): PrismModule {
  return {
    id: "test.probe",
    title: "Probe",
    description: "Synthetic module for API tests.",
    version: 1,
    packs: ["directory.users.read"],
    migrations: [
      {
        id: "001-items",
        sql: "CREATE TABLE IF NOT EXISTS mod_test_probe_item (tenant_id uuid NOT NULL, id text NOT NULL); SELECT prism365_tenant_table('mod_test_probe_item');",
      },
    ],
    collectors: [
      {
        id: "poll",
        title: "Poll",
        cadenceSeconds: 300,
        timeoutSeconds: 30,
        run: async () => {
          counter.runs += 1;
          return undefined;
        },
      },
    ],
    retention: [],
  };
}

// A module whose collector reads Graph, so a refused token stops collection.
function reader(): PrismModule {
  return {
    ...probe({ runs: 0 }),
    id: "test.reader",
    migrations: [],
    collectors: [
      {
        id: "read",
        title: "Read",
        cadenceSeconds: 300,
        timeoutSeconds: 30,
        run: async (context) => {
          await context.graph.get(
            "users?$top=1",
            { feature: "directory.users.read", permission: "User.Read.All" },
            z.unknown(),
          );
          return undefined;
        },
      },
    ],
  };
}

let refuseTokens = false;

async function setup(
  enabledFeatures: SupportedReadFeature[] = ["directory.users.read"],
  module?: PrismModule,
) {
  const identityConfig: IdentityConfig = {
    tenantId,
    clientId: "33333333-3333-4333-8333-333333333333",
    origin,
    redirectUri: `${origin}/api/v1/auth/callback`,
    allowedActorIds: [actorId],
    enabledFeatures,
    credential: { clientSecret: "synthetic-only" },
  };
  const config: AppConfig = {
    mode: "live",
    origin,
    identity: identityConfig,
    production: false,
    missingConfiguration: [],
    enabledFeatures,
  };
  const store = new MemoryStateStore();
  const provider: MicrosoftIdentityProvider = {
    authorizationUrl: async () => "https://login.microsoftonline.com/synthetic",
    redeem: async () => ({
      tenantId,
      actorId,
      displayName: "Synthetic admin",
      accountId: "synthetic-account",
      cache: "synthetic-cache",
      scopes: ["User.Read", "User.Read.All"],
    }),
    token: async (session) => {
      if (refuseTokens) {
        const refused = new Error("Microsoft refused the token.");
        refused.name = "InteractionRequiredAuthError";
        throw refused;
      }
      return { accessToken: "synthetic-token", cache: "synthetic-cache", scopes: session.scopes };
    },
  };
  const identity = new IdentityService(identityConfig, store, provider);
  const login = await identity.startLogin(enabledFeatures[0] ?? "directory.users.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);

  const counter = { runs: 0 };
  const registry = new ModuleRegistry().register(module ?? probe(counter));
  const database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  const scheduler = new ModuleScheduler({
    registry,
    database,
    graphClient: (getAccessToken) =>
      new GraphHttpClient({
        getAccessToken,
        fetchImplementation: async () => Response.json({ value: [] }),
      }),
  });
  const modules: ModuleRuntime = { registry, database, scheduler };
  const app = createApp({ config, store, identity, modules });
  closers.push(() => app.close());
  const headers = {
    cookie: `prism365_session=${loggedIn.handle}`,
    origin,
    "x-csrf-token": loggedIn.session.csrfToken,
  };
  return { app, store, modules, counter, headers, identity, config };
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 50));

describe("module API in live mode", () => {
  it("requires a signed-in administrator", async () => {
    const { app } = await setup();
    const response = await app.inject({ method: "GET", url: "/v1/modules" });
    expect(response.statusCode).toBe(401);
    const findings = await app.inject({ method: "GET", url: "/v1/findings" });
    expect(findings.statusCode).toBe(401);
  });

  it("lists modules with permission readiness and collection status", async () => {
    const { app, headers } = await setup();
    const response = await app.inject({ method: "GET", url: "/v1/modules", headers });
    expect(response.statusCode).toBe(200);
    expect(response.json()).toMatchObject({
      mode: "live",
      collection: { available: true, state: "stopped", durableStorage: false },
      modules: [
        {
          id: "test.probe",
          permissions: ["User.Read.All"],
          missingPermissions: [],
          collectors: [{ id: "poll", lastRun: null }],
        },
      ],
    });
  });

  it("starts collection only with CSRF protection, runs it, and audits it", async () => {
    const { app, headers, store, counter } = await setup();
    const forged = await app.inject({
      method: "POST",
      url: "/v1/collection/start",
      headers: { cookie: headers.cookie, origin },
    });
    expect(forged.statusCode).toBe(403);

    const started = await app.inject({ method: "POST", url: "/v1/collection/start", headers });
    expect(started.statusCode).toBe(200);
    expect(started.json()).toMatchObject({ state: "collecting", credentialKind: "delegated" });
    await settle();
    expect(counter.runs).toBe(1);
    expect(store.audit.map((entry) => entry.event.type)).toContain("collection.started");
    const listed = await app.inject({ method: "GET", url: "/v1/modules", headers });
    expect(listed.json().modules[0].collectors[0].lastRun).toMatchObject({
      status: "succeeded",
      credentialKind: "delegated",
    });
  });

  it("stops collection when the operator signs out and records why", async () => {
    const { app, headers, store, modules } = await setup();
    await app.inject({ method: "POST", url: "/v1/collection/start", headers });
    await settle();
    const signedOut = await app.inject({ method: "POST", url: "/v1/auth/logout", headers });
    expect(signedOut.statusCode).toBe(200);
    expect(modules.scheduler.status()).toMatchObject({
      state: "stopped",
      lastStopReason: "signed_out",
    });
    await settle();
    expect(store.audit).toContainEqual(
      expect.objectContaining({
        event: expect.objectContaining({ type: "collection.stopped", reason: "signed_out" }),
      }),
    );
  });

  it("resumes collection for the administrator who started it, until someone stops it", async () => {
    const { app, headers, store, modules, identity, config } = await setup();
    await app.inject({ method: "POST", url: "/v1/collection/start", headers });
    await settle();
    // A restart: nothing is bound, and a new server reads the standing choice from storage.
    modules.scheduler.unbind("shutdown");
    const restarted = createApp({ config, store, identity, modules });
    const paused = await restarted.inject({
      method: "POST",
      url: "/v1/collection/stop",
      headers: { cookie: headers.cookie },
    });
    expect(paused.statusCode).toBe(403);
    expect(modules.scheduler.status().state).toBe("stopped");
    const listed = await restarted.inject({ method: "GET", url: "/v1/modules", headers });
    expect(listed.statusCode).toBe(200);
    expect(modules.scheduler.status().state).toBe("collecting");
    expect(store.audit.map((entry) => entry.event.type)).toContain("collection.resumed");

    // Signing out pauses it; signing in again resumes it.
    await restarted.inject({ method: "POST", url: "/v1/auth/logout", headers });
    expect(modules.scheduler.status()).toMatchObject({
      state: "stopped",
      lastStopReason: "signed_out",
    });
    const login = await identity.startLogin("directory.users.read");
    const again = await identity.finishLogin("synthetic-code", login.state, login.state);
    const fresh = {
      cookie: `prism365_session=${again.handle}`,
      origin,
      "x-csrf-token": again.session.csrfToken,
    };
    const status = await restarted.inject({ method: "GET", url: "/v1/modules", headers: fresh });
    expect(status.json().collection).toMatchObject({ state: "collecting" });

    // Stop collecting ends the choice: nothing resumes after it.
    await restarted.inject({ method: "POST", url: "/v1/collection/stop", headers: fresh });
    const after = await restarted.inject({ method: "GET", url: "/v1/modules", headers: fresh });
    expect(after.json().collection).toMatchObject({ state: "stopped", resumes: false });
    expect(modules.scheduler.status().state).toBe("stopped");
  });

  it("does not keep resuming a session Microsoft refuses tokens for", async () => {
    refuseTokens = true;
    try {
      const { app, headers, store, modules, identity } = await setup(
        ["directory.users.read"],
        reader(),
      );
      await app.inject({ method: "POST", url: "/v1/collection/start", headers });
      await settle();
      expect(modules.scheduler.status()).toMatchObject({
        state: "stopped",
        lastStopReason: "credential_unavailable",
      });
      let status: { collection: { resumes: boolean } } | null = null;
      for (let load = 0; load < 3; load += 1)
        status = (await app.inject({ method: "GET", url: "/v1/modules", headers })).json();
      expect(modules.scheduler.status().state).toBe("stopped");
      // The page must not promise this session that collection resumes.
      expect(status?.collection.resumes).toBe(false);
      expect(store.audit.filter((entry) => entry.event.type === "collection.resumed")).toEqual([]);

      // Signing in again gives a new session, which may resume.
      refuseTokens = false;
      const login = await identity.startLogin("directory.users.read");
      const again = await identity.finishLogin("synthetic-code", login.state, login.state);
      await app.inject({
        method: "GET",
        url: "/v1/modules",
        headers: { cookie: `prism365_session=${again.handle}` },
      });
      expect(modules.scheduler.status().state).toBe("collecting");
    } finally {
      refuseTokens = false;
    }
  });

  it("never resumes collection without its ledger entry", async () => {
    const { app, headers, store, modules } = await setup();
    await app.inject({ method: "POST", url: "/v1/collection/start", headers });
    await settle();
    modules.scheduler.unbind("shutdown");
    const append = store.appendAudit.bind(store);
    store.appendAudit = async (tenant, actor, event) => {
      if (event.type === "collection.resumed") throw new Error("ledger unavailable");
      return append(tenant, actor, event);
    };
    await app.inject({ method: "GET", url: "/v1/modules", headers });
    expect(modules.scheduler.status().state).toBe("stopped");
    store.appendAudit = append;
    await app.inject({ method: "GET", url: "/v1/modules", headers });
    expect(modules.scheduler.status().state).toBe("collecting");
  });

  it("does not start collecting when the start cannot be recorded", async () => {
    const { app, headers, store, modules } = await setup();
    const append = store.appendAudit.bind(store);
    store.appendAudit = async (tenant, actor, event) => {
      if (event.type === "collection.started") throw new Error("ledger unavailable");
      return append(tenant, actor, event);
    };
    const started = await app.inject({ method: "POST", url: "/v1/collection/start", headers });
    expect(started.statusCode).toBe(500);
    expect(modules.scheduler.status().state).toBe("stopped");
    store.appendAudit = append;
    // Nothing was left behind to resume.
    await app.inject({ method: "GET", url: "/v1/modules", headers });
    expect(modules.scheduler.status().state).toBe("stopped");
  });

  it("stops collecting even when the stored choice cannot be cleared, and says so", async () => {
    const { app, headers, modules } = await setup();
    await app.inject({ method: "POST", url: "/v1/collection/start", headers });
    const tenant = modules.database.tenant.bind(modules.database);
    modules.database.tenant = ((
      id: string,
      operation: (sql: SqlClient) => Promise<unknown>,
      options?: TenantOptions,
    ) =>
      tenant(
        id,
        (sql) =>
          operation({
            query: (text: string, params?: readonly unknown[]) =>
              text.startsWith("DELETE FROM prism365_preference")
                ? Promise.reject(new Error("storage unavailable"))
                : sql.query(text, params),
          } as SqlClient),
        options,
      )) as typeof modules.database.tenant;
    const stopped = await app.inject({ method: "POST", url: "/v1/collection/stop", headers });
    expect(stopped.statusCode).toBe(500);
    expect(stopped.json().message).toContain("Collection stopped");
    expect(modules.scheduler.status()).toMatchObject({
      state: "stopped",
      lastStopReason: "operator_stopped",
    });
    modules.database.tenant = tenant;
  });

  it("lets a stop win over a resume that read the stored choice before it", async () => {
    const { app, headers, store, modules, identity, config } = await setup();
    await app.inject({ method: "POST", url: "/v1/collection/start", headers });
    await settle();
    modules.scheduler.unbind("shutdown");
    // A new server whose first read of the stored choice is slow.
    const tenant = modules.database.tenant.bind(modules.database);
    modules.database.tenant = ((
      id: string,
      operation: (sql: SqlClient) => Promise<unknown>,
      options?: TenantOptions,
    ) =>
      tenant(
        id,
        (sql) =>
          operation({
            query: async (text: string, params?: readonly unknown[]) => {
              const result = await sql.query(text, params);
              if (text.startsWith("SELECT value FROM prism365_preference"))
                await new Promise((resolve) => setTimeout(resolve, 200));
              return result;
            },
          } as SqlClient),
        options,
      )) as typeof modules.database.tenant;
    const restarted = createApp({ config, store, identity, modules });
    const [, stop] = await Promise.all([
      restarted.inject({ method: "GET", url: "/v1/modules", headers }),
      new Promise((resolve) => setTimeout(resolve, 20)).then(() =>
        restarted.inject({ method: "POST", url: "/v1/collection/stop", headers }),
      ),
    ]);
    modules.database.tenant = tenant;
    expect(stop.statusCode).toBe(200);
    expect(modules.scheduler.status().state).toBe("stopped");
    await restarted.inject({ method: "GET", url: "/v1/modules", headers });
    expect(modules.scheduler.status().state).toBe("stopped");
  });

  it("does not hold a passing storage error against the session", async () => {
    const { app, headers, modules } = await setup();
    await app.inject({ method: "POST", url: "/v1/collection/start", headers });
    await settle();
    // What the scheduler does when a session read fails: the credential's state is unknown, so
    // collection stops, but Microsoft never refused a token for this session.
    modules.scheduler.unbind("credential_unavailable");
    const listed = await app.inject({ method: "GET", url: "/v1/modules", headers });
    expect(listed.json().collection.state).toBe("collecting");
  });

  it("puts back another administrator's choice when a start cannot be recorded", async () => {
    const { app, headers, store, modules } = await setup();
    const other = {
      actorId: "55555555-5555-4555-8555-555555555555",
      since: "2026-10-01T00:00:00.000Z",
    };
    await modules.database.tenant(tenantId, (sql) => writeCollectionIntent(sql, tenantId, other));
    const append = store.appendAudit.bind(store);
    store.appendAudit = async (tenant, actor, event) => {
      if (event.type === "collection.started") throw new Error("ledger unavailable");
      return append(tenant, actor, event);
    };
    const started = await app.inject({ method: "POST", url: "/v1/collection/start", headers });
    store.appendAudit = append;
    expect(started.statusCode).toBe(500);
    expect(
      await modules.database.tenant(tenantId, (sql) => readCollectionIntent(sql, tenantId)),
    ).toEqual(other);
    // The server's own copy is restored too: the other administrator's collection still resumes.
    const status = await app.inject({ method: "GET", url: "/v1/modules", headers });
    expect(status.json().collection).toMatchObject({ state: "stopped", resumes: true });
  });

  it("answers page loads within seconds even when resuming is stuck", async () => {
    const { app, headers, store, modules } = await setup();
    await app.inject({ method: "POST", url: "/v1/collection/start", headers });
    await settle();
    modules.scheduler.unbind("shutdown");
    // Resuming waits on a ledger write that never finishes.
    const append = store.appendAudit.bind(store);
    store.appendAudit = (tenant, actor, event) =>
      event.type === "collection.resumed" ? new Promise(() => {}) : append(tenant, actor, event);
    const began = Date.now();
    const page = await app.inject({ method: "GET", url: "/v1/modules", headers });
    store.appendAudit = append;
    expect(page.statusCode).toBe(200);
    expect(Date.now() - began).toBeLessThan(8_000);
    expect(modules.scheduler.status().state).toBe("stopped");
  }, 20_000);

  it("records an operator stop against the operator who stopped it", async () => {
    const { app, headers, store, modules } = await setup();
    await app.inject({ method: "POST", url: "/v1/collection/start", headers });
    const stopped = await app.inject({ method: "POST", url: "/v1/collection/stop", headers });
    expect(stopped.json()).toMatchObject({ state: "stopped", lastStopReason: "operator_stopped" });
    expect(modules.scheduler.status().state).toBe("stopped");
    const stops = store.audit.filter((entry) => entry.event.type === "collection.stopped");
    expect(stops).toHaveLength(1);
    expect(stops[0]).toMatchObject({ actorId, event: { reason: "operator_stopped" } });
  });

  it("never collects a pack the installation did not enable, even if consented", async () => {
    const { app, headers, counter } = await setup(["user.licenses.read"]);
    const listed = await app.inject({ method: "GET", url: "/v1/modules", headers });
    expect(listed.json().modules[0].missingPermissions).toEqual(["User.Read.All"]);
    await app.inject({ method: "POST", url: "/v1/collection/start", headers });
    await settle();
    expect(counter.runs).toBe(0);
  });

  it("lists findings and changes their state with an audit entry", async () => {
    const { app, headers, store, modules } = await setup();
    const observedAt = new Date("2026-10-22T12:00:00Z");
    const { id } = await modules.database.tenant(tenantId, (sql) =>
      upsertFinding(sql, tenantId, "test.probe", {
        ruleId: "probe.synthetic",
        ruleVersion: 1,
        dedupeKey: "synthetic",
        purpose: "security",
        severity: "high",
        score: 70,
        subject: { kind: "person", id: "synthetic-user", label: "Synthetic User" },
        title: "Synthetic finding",
        summary: "A synthetic finding for the API test.",
        evidence: [
          { source: "Synthetic", ref: "e1", observedAt: observedAt.toISOString(), fields: {} },
        ],
        limitations: ["Synthetic data."],
        window: { start: observedAt, end: observedAt },
        observedAt,
      }),
    );
    const listed = await app.inject({ method: "GET", url: "/v1/findings?state=open", headers });
    expect(listed.json()).toMatchObject({ mode: "live", items: [{ id, severity: "high" }] });

    const invalid = await app.inject({
      method: "POST",
      url: "/v1/findings/not-a-uuid/state",
      headers,
      payload: { state: "resolved" },
    });
    expect(invalid.statusCode).toBe(400);
    const missing = await app.inject({
      method: "POST",
      url: "/v1/findings/44444444-4444-4444-8444-444444444444/state",
      headers,
      payload: { state: "resolved" },
    });
    expect(missing.statusCode).toBe(404);

    const resolved = await app.inject({
      method: "POST",
      url: `/v1/findings/${id}/state`,
      headers,
      payload: { state: "resolved" },
    });
    expect(resolved.statusCode).toBe(200);
    expect(resolved.json()).toMatchObject({ id, state: "resolved" });
    expect(store.audit).toContainEqual(
      expect.objectContaining({
        actorId,
        event: expect.objectContaining({ type: "finding.state.changed", findingId: id }),
      }),
    );
  });
});

describe("module API in demo mode", () => {
  it("collects nothing and says so", async () => {
    const app = createApp();
    closers.push(() => app.close());
    const modules = await app.inject({ method: "GET", url: "/v1/modules" });
    expect(modules.json()).toMatchObject({
      mode: "demo",
      collection: { available: false, state: "stopped" },
      modules: [],
    });
    const findings = await app.inject({ method: "GET", url: "/v1/findings" });
    expect(findings.json()).toEqual({ mode: "demo", items: [] });
    const start = await app.inject({ method: "POST", url: "/v1/collection/start" });
    expect(start.statusCode).toBe(409);
  });
});

import { GraphHttpClient } from "@prism365/graph";
import {
  ModuleBoundaryError,
  ModuleRegistry,
  ModuleScheduler,
  type PrismModule,
} from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase } from "@prism365/storage";
import { afterEach, describe, expect, it, vi } from "vitest";
import { z } from "zod";

vi.setConfig({ testTimeout: 30_000 });

// The context the app hands to every route file, captured from one of them.
const captured: { context: RouteContext | null } = { context: null };
vi.mock("../src/routes/changes.ts", () => ({
  registerChangeRoutes: (_app: unknown, context: RouteContext) => {
    captured.context = context;
  },
}));

import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig, SupportedReadFeature } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import { RouteBoundaryError, type RouteContext } from "../src/routes/context.ts";

const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
const secret = "synthetic-secret-never-shown";

const probe: PrismModule = {
  id: "test.probe",
  title: "Probe",
  description: "Synthetic.",
  version: 1,
  packs: ["directory.users.read"],
  migrations: [],
  collectors: [
    {
      id: "poll",
      title: "Poll",
      cadenceSeconds: 300,
      timeoutSeconds: 30,
      run: async () => undefined,
    },
  ],
  retention: [],
};

const closers: (() => Promise<unknown>)[] = [];
afterEach(async () => {
  while (closers.length) await closers.pop()?.();
  captured.context = null;
});

async function setup(enabledFeatures: SupportedReadFeature[]) {
  const identityConfig: IdentityConfig = {
    tenantId,
    clientId: "33333333-3333-4333-8333-333333333333",
    origin,
    redirectUri: `${origin}/api/v1/auth/callback`,
    allowedActorIds: [actorId],
    enabledFeatures,
    credential: { clientSecret: secret },
  };
  const config: AppConfig = {
    mode: "live",
    origin,
    identity: identityConfig,
    production: false,
    missingConfiguration: [],
    enabledFeatures,
    encryptionKey: "a".repeat(64),
    databaseUrl: "postgres://synthetic:password-never-shown@db.invalid/prism365",
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
      // Entra returns every consented scope, including ones this installation never enabled.
      scopes: ["User.Read", "User.Read.All", "AuditLog.Read.All", "Directory.Read.All"],
    }),
    token: async (session) => ({ accessToken: "synthetic", cache: "c", scopes: session.scopes }),
  };
  const identity = new IdentityService(identityConfig, store, provider);
  const login = await identity.startLogin(enabledFeatures[0] ?? "directory.users.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
  const registry = new ModuleRegistry().register(probe);
  const database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  const scheduler = new ModuleScheduler({
    enabledPacks: enabledFeatures,
    registry,
    database,
    graphClient: (getAccessToken) =>
      new GraphHttpClient({
        getAccessToken,
        fetchImplementation: async () => Response.json({ value: [] }),
      }),
  });
  const requested: string[] = [];
  const app = createApp({
    config,
    store,
    identity,
    modules: { registry, database, scheduler },
    graphFetch: (async (input: RequestInfo | URL) => {
      requested.push(String(input));
      return Response.json({ value: [] });
    }) as typeof fetch,
  });
  // Closing the app stops the scheduler and closes the module database.
  closers.push(() => app.close());
  const context = captured.context;
  if (!context) throw new Error("The route context was not captured.");
  const headers = {
    cookie: `prism365_session=${loggedIn.handle}`,
    origin,
    "x-csrf-token": loggedIn.session.csrfToken,
  };
  return { app, context, session: loggedIn.session, handle: loggedIn.handle, headers, requested };
}

describe("the context route files get", () => {
  it("carries no credential, key or database address", async () => {
    const { context } = await setup(["directory.users.read"]);
    const visible = JSON.stringify(context);
    expect(visible).not.toContain(secret);
    expect(visible).not.toContain("password-never-shown");
    expect(visible).not.toContain("a".repeat(64));
    expect(Object.keys(context)).not.toContain("identity");
    expect(Object.keys(context)).not.toContain("config");
    expect(context.settings).toMatchObject({
      mode: "live",
      configured: true,
      credentialKind: "secret",
      certificate: null,
      durableStorage: true,
      enabledFeatures: ["directory.users.read"],
    });
  });

  it("refuses a Graph client for a pack the operator did not enable", async () => {
    const { context, session, handle } = await setup(["directory.users.read"]);
    expect(() => context.sessionGraph(session, handle, ["audit.directory.read"])).toThrow(
      RouteBoundaryError,
    );
  });

  it("narrows the client to the packs asked for, and the grant to their permissions", async () => {
    const { context, session, handle, requested } = await setup([
      "directory.users.read",
      "audit.directory.read",
    ]);
    const { graph, granted } = context.sessionGraph(session, handle, ["directory.users.read"]);
    expect(granted).toEqual(["User.Read.All"]);
    const schema = z.unknown();
    await expect(
      graph.get(
        "auditLogs/directoryAudits",
        { feature: "audit.directory.read", permission: "AuditLog.Read.All" },
        schema,
      ),
    ).rejects.toThrow(ModuleBoundaryError);
    expect(requested).toEqual([]);
    await graph.get(
      "users",
      { feature: "directory.users.read", permission: "User.Read.All" },
      schema,
    );
    expect(requested).toHaveLength(1);
  });

  it("says who is collecting without handing out their credential", async () => {
    const { app, context, headers } = await setup(["directory.users.read"]);
    expect(context.collectingAdmin()).toBeNull();
    const started = await app.inject({ method: "POST", url: "/v1/collection/start", headers });
    expect(started.statusCode).toBe(200);
    expect(context.collectingAdmin()).toEqual({ actorId });
    await app.inject({ method: "POST", url: "/v1/collection/stop", headers });
    expect(context.collectingAdmin()).toBeNull();
  });
});

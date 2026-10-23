import { GraphHttpClient } from "@prism365/graph";
import { createSigninsAuditModule } from "@prism365/module-signins-audit";
import { ModuleRegistry, ModuleScheduler } from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";

const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
const identityConfig: IdentityConfig = {
  tenantId,
  clientId: "33333333-3333-4333-8333-333333333333",
  origin,
  redirectUri: `${origin}/api/v1/auth/callback`,
  allowedActorIds: [actorId],
  enabledFeatures: ["audit.signins.read"],
  credential: { clientSecret: "synthetic-only" },
};
const config: AppConfig = {
  mode: "live",
  origin,
  identity: identityConfig,
  production: false,
  missingConfiguration: [],
  enabledFeatures: ["audit.signins.read"],
};
const provider: MicrosoftIdentityProvider = {
  authorizationUrl: async () => "https://login.microsoftonline.com/synthetic",
  redeem: async () => ({
    tenantId,
    actorId,
    displayName: "Synthetic admin",
    accountId: "synthetic-account",
    cache: "synthetic-cache",
    scopes: ["User.Read", "AuditLog.Read.All"],
  }),
  token: async (session) => ({ accessToken: "synthetic", cache: "c", scopes: session.scopes }),
};

let app: ReturnType<typeof createApp>;
let cookie: string;

beforeAll(async () => {
  const store = new MemoryStateStore();
  const identity = new IdentityService(identityConfig, store, provider);
  const login = await identity.startLogin("audit.signins.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
  cookie = `prism365_session=${loggedIn.handle}`;
  const registry = new ModuleRegistry().register(createSigninsAuditModule());
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
  app = createApp({ config, store, identity, modules: { registry, database, scheduler } });
}, 60_000);

afterAll(async () => {
  await app?.close();
});

describe("sign-in audit API", () => {
  it("requires a session", async () => {
    const response = await app.inject({ method: "GET", url: "/v1/modules/signins.audit/summary" });
    expect(response.statusCode).toBe(401);
  });

  it("validates the range and returns an empty summary on a fresh install", async () => {
    const bad = await app.inject({
      method: "GET",
      url: "/v1/modules/signins.audit/summary?days=0",
      headers: { cookie },
    });
    expect(bad.statusCode).toBe(400);
    const response = await app.inject({
      method: "GET",
      url: "/v1/modules/signins.audit/summary?days=30",
      headers: { cookie },
    });
    expect(response.statusCode).toBe(200);
    expect(response.json()).toMatchObject({
      days: 30,
      totals: { successes: 0, failures: 0 },
      countries: [],
      failingNetworks: [],
      coverage: { events: 0, firstEvent: null },
    });
  });

  it("lists the module with no missing permission for a consented session", async () => {
    const response = await app.inject({ method: "GET", url: "/v1/modules", headers: { cookie } });
    expect(response.json().modules).toEqual([
      expect.objectContaining({
        id: "signins.audit",
        permissions: ["AuditLog.Read.All"],
        missingPermissions: [],
      }),
    ]);
  });
});

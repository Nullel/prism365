import { MemoryStateStore } from "@prism365/storage";
import { afterAll, describe, expect, it } from "vitest";
import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";

// Regression: authorization must follow the route the router matched. The router decodes
// percent-encoding, so a raw-URL check let "/%76%31/..." reach /v1 handlers without a session.

const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
const identityConfig: IdentityConfig = {
  tenantId,
  clientId: "33333333-3333-4333-8333-333333333333",
  origin,
  redirectUri: `${origin}/api/v1/auth/callback`,
  allowedActorIds: [actorId],
  enabledFeatures: ["directory.users.read"],
  credential: { clientSecret: "synthetic-only" },
};
const config: AppConfig = {
  mode: "live",
  origin,
  identity: identityConfig,
  production: false,
  missingConfiguration: [],
  enabledFeatures: ["directory.users.read"],
};
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
  token: async (session) => ({ accessToken: "synthetic", cache: "c", scopes: session.scopes }),
};
const store = new MemoryStateStore();
const identity = new IdentityService(identityConfig, store, provider);
const app = createApp({ config, store, identity });

afterAll(async () => {
  await app.close();
});

const encode = (path: string) =>
  path.replace(/^\/v1\//, "/%76%31/").replace(/\/api\/v1\//, "/api/%76%31/");

describe("route authorization", () => {
  it("refuses protected reads through encoded paths without a session", async () => {
    for (const path of [
      "/v1/modules",
      "/v1/findings",
      "/v1/automation/packages",
      "/v1/investigations/user?email=alex@example.invalid",
      "/v1/ai/status",
      "/v1/audits/accounts.inactive/trend",
      "/v1/settings/audit-schedule",
    ])
      for (const url of [path, encode(path), `/api${encode(path)}`]) {
        const response = await app.inject({ method: "GET", url });
        expect([401, 404], `${url} returned ${response.statusCode}`).toContain(response.statusCode);
        expect(response.body).not.toContain('"items"');
      }
  });

  it("applies the CSRF check to encoded state-changing paths", async () => {
    const login = await identity.startLogin("directory.users.read");
    const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
    const cookie = `prism365_session=${loggedIn.handle}`;
    for (const path of [
      "/v1/collection/start",
      "/v1/collection/stop",
      "/v1/findings/44444444-4444-4444-8444-444444444444/state",
      "/v1/auth/logout",
      "/v1/settings/audit-schedule",
    ]) {
      const response = await app.inject({
        method: "POST",
        url: encode(path),
        headers: { cookie, origin },
        payload: { state: "resolved" },
      });
      expect(response.statusCode, encode(path)).toBe(403);
    }
  });
});

import { MemoryStateStore } from "@prism365/storage";
import { describe, expect, it } from "vitest";
import { createApp } from "../src/app.ts";
import { type AppConfig, type IdentityConfig, loadConfig } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";

// Synthetic tenant and documentation-range addresses only.
const tenantId = "11111111-1111-4111-8111-111111111111";
const origin = "http://localhost:5173";
const identityConfig: IdentityConfig = {
  tenantId,
  clientId: "33333333-3333-4333-8333-333333333333",
  origin,
  redirectUri: `${origin}/api/v1/auth/callback`,
  allowedActorIds: ["22222222-2222-4222-8222-222222222222"],
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
  redeem: async () => {
    throw new Error("not used");
  },
  token: async () => {
    throw new Error("not used");
  },
};

function liveApp(trustProxyHops?: number) {
  const store = new MemoryStateStore();
  return createApp({
    config: { ...config, ...(trustProxyHops ? { trustProxyHops } : {}) },
    store,
    identity: new IdentityService(identityConfig, store, provider),
  });
}

async function signInStarts(
  app: ReturnType<typeof liveApp>,
  forwardedFor: (index: number) => string,
) {
  const statuses: number[] = [];
  for (let index = 0; index < 11; index += 1) {
    const response = await app.inject({
      method: "GET",
      url: "/v1/auth/login",
      remoteAddress: "172.18.0.5",
      headers: { "x-forwarded-for": forwardedFor(index) },
    });
    statuses.push(response.statusCode);
  }
  return statuses;
}

describe("client address behind a proxy", () => {
  it("limits sign-in starts per client when the proxy is trusted", async () => {
    const app = liveApp(1);
    const distinct = await signInStarts(app, (index) => `198.51.100.${index + 1}`);
    expect(distinct).not.toContain(429);
    const repeated = await signInStarts(app, () => "203.0.113.9");
    expect(repeated.at(-1)).toBe(429);
    await app.close();
  });

  it("ignores forwarded addresses when no proxy is trusted", async () => {
    const app = liveApp();
    const statuses = await signInStarts(app, (index) => `198.51.100.${index + 1}`);
    expect(statuses.at(-1)).toBe(429);
    await app.close();
  });

  it("reads TRUST_PROXY_HOPS and refuses anything but 0 to 3", async () => {
    const live = { DATA_MODE: "live", PUBLIC_ORIGIN: origin };
    expect((await loadConfig({ ...live, TRUST_PROXY_HOPS: "1" })).trustProxyHops).toBe(1);
    await expect(loadConfig({ ...live, TRUST_PROXY_HOPS: "yes" })).rejects.toThrow(
      "TRUST_PROXY_HOPS",
    );
  });
});

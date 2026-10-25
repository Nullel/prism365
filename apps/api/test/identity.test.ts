import { MemoryStateStore } from "@prism365/storage";
import { generateKeyPair, SignJWT } from "jose";
import { describe, expect, it, vi } from "vitest";
import { createApp } from "../src/app.ts";
import { type AppConfig, type IdentityConfig, loadConfig } from "../src/config.ts";
import {
  digest,
  IdentityService,
  type MicrosoftIdentityProvider,
  validateMicrosoftIdToken,
} from "../src/identity.ts";

const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const clientId = "33333333-3333-4333-8333-333333333333";
const identityConfig: IdentityConfig = {
  tenantId,
  clientId,
  origin: "http://localhost:5173",
  redirectUri: "http://localhost:5173/api/v1/auth/callback",
  allowedActorIds: [actorId],
  enabledFeatures: ["directory.users.read", "user.licenses.read", "audit.signins.read"],
  credential: { clientSecret: "synthetic-development-credential" },
};
const config: AppConfig = {
  mode: "live",
  origin: identityConfig.origin,
  production: false,
  identity: identityConfig,
  missingConfiguration: [],
  enabledFeatures: identityConfig.enabledFeatures,
};

function harness() {
  const store = new MemoryStateStore();
  const provider: MicrosoftIdentityProvider = {
    authorizationUrl: async (transaction) =>
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?state=${transaction.state}`,
    redeem: vi.fn(async (_code, transaction) => ({
      tenantId,
      actorId,
      displayName: "Synthetic administrator",
      accountId: "synthetic-account",
      cache: "synthetic-cache",
      scopes: transaction.scopes,
    })),
    token: vi.fn(async (session) => ({
      accessToken: "synthetic-token",
      cache: "synthetic-cache",
      scopes: session.scopes,
    })),
  };
  const identity = new IdentityService(identityConfig, store, provider);
  return { store, provider, identity };
}

describe("tenant identity", () => {
  it("binds login to the initiating browser, consumes state once, and rejects unauthorized actors", async () => {
    const { identity, provider } = harness();
    const login = await identity.startLogin("directory.users.read");
    await expect(identity.finishLogin("fake-code", login.state, "wrong-browser")).rejects.toThrow(
      "browser binding",
    );
    expect(provider.redeem).not.toHaveBeenCalled();
    const result = await identity.finishLogin("fake-code", login.state, login.state);
    expect(await identity.session(result.handle)).toMatchObject({ tenantId, actorId });
    await expect(identity.finishLogin("fake-code", login.state, login.state)).rejects.toThrow();
    expect(provider.redeem).toHaveBeenCalledTimes(1);
    provider.redeem = async () => ({
      tenantId,
      actorId: clientId,
      displayName: "Other actor",
      accountId: "other",
      cache: "x",
      scopes: ["User.Read.All"],
    });
    const other = await identity.startLogin("directory.users.read");
    await expect(identity.finishLogin("fake-code", other.state, other.state)).rejects.toThrow(
      "not an authorized",
    );
  });

  it("does not resurrect a session revoked while a token refresh was in flight", async () => {
    const { identity, provider, store } = harness();
    const login = await identity.startLogin("directory.users.read");
    const result = await identity.finishLogin("fake-code", login.state, login.state);
    let release:
      | ((value: { accessToken: string; cache: string; scopes: string[] }) => void)
      | undefined;
    provider.token = async () =>
      new Promise((resolve) => {
        release = resolve;
      });
    const pending = identity.accessToken(result.handle, ["User.Read.All"]);
    await vi.waitFor(() => expect(release).toBeDefined());
    await identity.logout(result.handle, result.session);
    release?.({ accessToken: "synthetic-token", cache: "new-cache", scopes: ["User.Read.All"] });
    await expect(pending).rejects.toThrow("revoked");
    expect(await store.get(tenantId, "session", digest(result.handle))).toBeNull();
  });

  it("rejects wrong issuer, audience, tenant, nonce and expired signed tokens", async () => {
    const keys = await generateKeyPair("RS256");
    const expected = { tenantId, clientId, nonce: "synthetic-nonce" };
    const sign = async (
      claims: Record<string, unknown>,
      issuer = `https://login.microsoftonline.com/${tenantId}/v2.0`,
      audience = clientId,
      expiry = "1h",
    ) =>
      new SignJWT({ tid: tenantId, oid: actorId, nonce: expected.nonce, ...claims })
        .setProtectedHeader({ alg: "RS256" })
        .setIssuer(issuer)
        .setAudience(audience)
        .setIssuedAt()
        .setExpirationTime(expiry)
        .sign(keys.privateKey);
    await expect(
      validateMicrosoftIdToken(await sign({}), expected, keys.publicKey),
    ).resolves.toMatchObject({ tenantId, actorId });
    for (const token of [
      await sign({ nonce: "other" }),
      await sign({ tid: clientId }),
      await sign({}, "https://untrusted.invalid"),
      await sign({}, undefined, tenantId),
      await sign({}, undefined, undefined, "-1h"),
    ]) {
      await expect(validateMicrosoftIdToken(token, expected, keys.publicKey)).rejects.toThrow();
    }
  });

  it("keeps live routes authenticated and mutations bound to origin and CSRF", async () => {
    const { identity, store } = harness();
    const app = createApp({ config, identity, store });
    const anonymous = await app.inject({
      url: "/api/v1/ai/status",
    });
    expect(anonymous.statusCode).toBe(401);
    const start = await app.inject({ url: "/v1/auth/login?feature=directory.users.read" });
    const state = new URL(start.headers.location ?? "").searchParams.get("state") ?? "";
    const callback = await app.inject({
      url: `/v1/auth/callback?code=fake-code&state=${state}`,
      headers: { cookie: `prism365_oauth=${state}` },
    });
    const cookies = callback.headers["set-cookie"];
    const sessionCookie =
      (Array.isArray(cookies) ? cookies : [cookies ?? ""])
        .find((value) => value.startsWith("prism365_session="))
        ?.split(";")[0] ?? "";
    expect(sessionCookie).not.toBe("");
    const session = await app.inject({
      url: "/api/v1/session",
      headers: { cookie: sessionCookie },
    });
    expect(session.json()).toMatchObject({
      mode: "live",
      authenticated: true,
      displayName: "Synthetic administrator",
    });
    expect(JSON.stringify(session.json())).not.toContain("synthetic-token");
    const rejected = await app.inject({
      method: "POST",
      url: "/v1/auth/logout",
      headers: {
        cookie: sessionCookie,
        origin: "https://attacker.invalid",
        "x-csrf-token": session.json().csrfToken,
      },
    });
    expect(rejected.statusCode).toBe(403);
    const logout = await app.inject({
      method: "POST",
      url: "/v1/auth/logout",
      headers: {
        cookie: sessionCookie,
        origin: config.origin,
        "x-csrf-token": session.json().csrfToken,
      },
    });
    expect(logout.statusCode).toBe(200);
    expect(
      (await app.inject({ url: "/v1/session", headers: { cookie: sessionCookie } })).json()
        .authenticated,
    ).toBe(false);
    await app.close();
  });

  it("reports incomplete configuration and refuses unsupported clouds or insecure production", async () => {
    const incomplete = await loadConfig({
      DATA_MODE: "live",
      PUBLIC_ORIGIN: "http://localhost:5173",
    });
    expect(incomplete.identity).toBeNull();
    expect(incomplete.missingConfiguration).toContain("ENTRA_ADMIN_OBJECT_IDS");
    await expect(loadConfig({ DATA_MODE: "live", MICROSOFT_CLOUD: "china" })).rejects.toThrow(
      "Worldwide",
    );
    await expect(
      loadConfig({
        DATA_MODE: "live",
        NODE_ENV: "production",
        PUBLIC_ORIGIN: "http://localhost:5173",
      }),
    ).rejects.toThrow("HTTPS");
    await expect(loadConfig({ DATA_MODE: "live", ENTRA_TENANT_ID: "common" })).rejects.toThrow(
      "UUID",
    );
  });

  it("reads the directory change margin and retention, and refuses values out of range", async () => {
    const live = { DATA_MODE: "live", PUBLIC_ORIGIN: "http://localhost:5173" };
    expect((await loadConfig(live)).directoryChanges).toMatchObject({
      lateArrivalMinutes: 240,
      retentionDays: 3_650,
    });
    expect(
      (
        await loadConfig({
          ...live,
          DIRECTORY_CHANGES_LATE_MINUTES: "90",
          DIRECTORY_CHANGES_RETENTION_DAYS: "730",
        })
      ).directoryChanges,
    ).toMatchObject({ lateArrivalMinutes: 90, retentionDays: 730 });
    await expect(loadConfig({ ...live, DIRECTORY_CHANGES_RETENTION_DAYS: "30" })).rejects.toThrow(
      "DIRECTORY_CHANGES_RETENTION_DAYS",
    );
    await expect(loadConfig({ ...live, DIRECTORY_CHANGES_LATE_MINUTES: "5" })).rejects.toThrow(
      "DIRECTORY_CHANGES_LATE_MINUTES",
    );
  });
});

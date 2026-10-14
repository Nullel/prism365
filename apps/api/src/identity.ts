import { createHash, randomBytes, timingSafeEqual } from "node:crypto";
import {
  ConfidentialClientApplication,
  type AccountInfo,
  type Configuration,
  ResponseMode,
} from "@azure/msal-node";
import type { StateStore } from "@prism365/storage";
import { createRemoteJWKSet, jwtVerify, type JWTVerifyGetKey } from "jose";
import { z } from "zod";
import {
  type IdentityConfig,
  type SupportedReadFeature,
  supportedReadFeatures,
  isUuid,
} from "./config.ts";

export async function validateMicrosoftIdToken(
  token: string,
  expected: { tenantId: string; clientId: string; nonce: string },
  key: CryptoKey | JWTVerifyGetKey,
) {
  const { payload } = await jwtVerify(token, key, {
    issuer: `https://login.microsoftonline.com/${expected.tenantId}/v2.0`,
    audience: expected.clientId,
    algorithms: ["RS256"],
  });
  if (
    payload.tid !== expected.tenantId ||
    typeof payload.oid !== "string" ||
    !isUuid(payload.oid) ||
    payload.nonce !== expected.nonce
  )
    throw new Error("Microsoft identity claims did not match the login transaction.");
  return {
    tenantId: expected.tenantId,
    actorId: payload.oid.toLowerCase(),
    displayName: typeof payload.name === "string" ? payload.name : "Administrator",
  };
}

export interface OAuthTransaction {
  state: string;
  nonce: string;
  verifier: string;
  scopes: string[];
  createdAt: number;
}

const identitySessionSchema = z.object({
  tenantId: z.uuid(),
  actorId: z.uuid(),
  displayName: z.string(),
  csrfToken: z.string().min(32),
  accountId: z.string(),
  cache: z.string(),
  scopes: z.array(z.string()),
  expiresAt: z.number(),
});
export type IdentitySession = z.infer<typeof identitySessionSchema>;

export interface MicrosoftIdentityProvider {
  authorizationUrl(transaction: OAuthTransaction): Promise<string>;
  redeem(
    code: string,
    transaction: OAuthTransaction,
  ): Promise<{
    tenantId: string;
    actorId: string;
    displayName: string;
    accountId: string;
    cache: string;
    scopes: string[];
  }>;
  token(
    session: IdentitySession,
    scopes: string[],
  ): Promise<{ accessToken: string; cache: string; scopes: string[] }>;
}

export function digest(value: string) {
  return createHash("sha256").update(value).digest("hex");
}
export function equalSecret(left: string, right: string) {
  const leftBytes = Buffer.from(left);
  const rightBytes = Buffer.from(right);
  return leftBytes.length === rightBytes.length && timingSafeEqual(leftBytes, rightBytes);
}
function randomHandle() {
  return randomBytes(32).toString("base64url");
}
function normalizeScopes(scopes: string[]) {
  return [
    ...new Set(scopes.map((scope) => scope.replace(/^https:\/\/graph\.microsoft\.com\//i, ""))),
  ];
}

export class MsalIdentityProvider implements MicrosoftIdentityProvider {
  readonly #config: IdentityConfig;
  readonly #msalConfig: Configuration;
  readonly #jwks = createRemoteJWKSet(
    new URL("https://login.microsoftonline.com/common/discovery/v2.0/keys"),
    { timeoutDuration: 5_000 },
  );
  constructor(config: IdentityConfig) {
    this.#config = config;
    this.#msalConfig = {
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        ...config.credential,
      },
      system: { loggerOptions: { piiLoggingEnabled: false, loggerCallback: () => {} } },
    };
  }
  #application(cache?: string) {
    const app = new ConfidentialClientApplication(this.#msalConfig);
    if (cache) app.getTokenCache().deserialize(cache);
    return app;
  }
  async authorizationUrl(transaction: OAuthTransaction) {
    return this.#application().getAuthCodeUrl({
      scopes: transaction.scopes,
      redirectUri: this.#config.redirectUri,
      responseMode: ResponseMode.QUERY,
      state: transaction.state,
      nonce: transaction.nonce,
      codeChallenge: createHash("sha256").update(transaction.verifier).digest("base64url"),
      codeChallengeMethod: "S256",
    });
  }
  async redeem(code: string, transaction: OAuthTransaction) {
    const app = this.#application();
    const result = await app.acquireTokenByCode({
      code,
      scopes: transaction.scopes,
      redirectUri: this.#config.redirectUri,
      codeVerifier: transaction.verifier,
      nonce: transaction.nonce,
    });
    if (!result?.account || !result.idToken)
      throw new Error("Microsoft did not return an authenticated account.");
    const verified = await validateMicrosoftIdToken(
      result.idToken,
      {
        tenantId: this.#config.tenantId,
        clientId: this.#config.clientId,
        nonce: transaction.nonce,
      },
      this.#jwks,
    );
    return {
      ...verified,
      accountId: result.account.homeAccountId,
      cache: app.getTokenCache().serialize(),
      scopes: normalizeScopes(result.scopes),
    };
  }
  async token(session: IdentitySession, scopes: string[]) {
    const app = this.#application(session.cache);
    const account: AccountInfo | null = await app
      .getTokenCache()
      .getAccountByHomeId(session.accountId);
    if (!account || account.tenantId.toLowerCase() !== this.#config.tenantId)
      throw new Error("The tenant account is no longer available.");
    const result = await app.acquireTokenSilent({ account, scopes });
    if (!result?.accessToken || result.tenantId.toLowerCase() !== this.#config.tenantId)
      throw new Error("Microsoft did not return a tenant-bound token.");
    return {
      accessToken: result.accessToken,
      cache: app.getTokenCache().serialize(),
      scopes: normalizeScopes(result.scopes),
    };
  }
}

export class IdentityService {
  readonly #config: IdentityConfig;
  readonly #store: StateStore;
  readonly #provider: MicrosoftIdentityProvider;
  readonly #locks = new Map<string, Promise<unknown>>();
  constructor(
    config: IdentityConfig,
    store: StateStore,
    provider: MicrosoftIdentityProvider = new MsalIdentityProvider(config),
  ) {
    this.#config = config;
    this.#store = store;
    this.#provider = provider;
  }
  get tenantId() {
    return this.#config.tenantId;
  }

  async startLogin(feature: SupportedReadFeature, session?: IdentitySession | null) {
    if (!this.#config.enabledFeatures.includes(feature))
      throw new Error("This read feature has not been enabled by the operator.");
    const permitted = new Set<string>([
      "User.Read",
      ...this.#config.enabledFeatures.map((item) => supportedReadFeatures[item]),
    ]);
    const scopes = [
      ...new Set([
        "User.Read",
        supportedReadFeatures[feature],
        ...(session?.scopes ?? []).filter((scope) => permitted.has(scope)),
      ]),
    ];
    const transaction: OAuthTransaction = {
      state: randomHandle(),
      nonce: randomHandle(),
      verifier: randomHandle() + randomHandle(),
      scopes,
      createdAt: Date.now(),
    };
    const url = await this.#provider.authorizationUrl(transaction);
    await this.#store.set(
      this.tenantId,
      "oauth",
      digest(transaction.state),
      transaction,
      new Date(Date.now() + 5 * 60_000),
    );
    return { url, state: transaction.state };
  }
  async finishLogin(code: string, state: string, browserState: string) {
    if (!state || !browserState || !equalSecret(state, browserState))
      throw new Error("Login browser binding did not match.");
    const value = await this.#store.take(this.tenantId, "oauth", digest(state));
    const transaction = z
      .object({
        state: z.string(),
        nonce: z.string(),
        verifier: z.string(),
        scopes: z.array(z.string()),
        createdAt: z.number(),
      })
      .parse(value);
    if (!equalSecret(transaction.state, state) || Date.now() - transaction.createdAt > 5 * 60_000)
      throw new Error("Login transaction expired.");
    const identity = await this.#provider.redeem(code, transaction);
    if (
      identity.tenantId !== this.tenantId ||
      !this.#config.allowedActorIds.includes(identity.actorId)
    ) {
      throw new Error("This Microsoft account is not an authorized Prism365 administrator.");
    }
    const handle = randomHandle();
    const session = identitySessionSchema.parse({
      ...identity,
      csrfToken: randomHandle(),
      expiresAt: Date.now() + 8 * 60 * 60_000,
    });
    await this.#store.appendAudit(this.tenantId, session.actorId, {
      type: "session.created",
      occurredAt: new Date().toISOString(),
    });
    await this.#store.set(
      this.tenantId,
      "session",
      digest(handle),
      session,
      new Date(session.expiresAt),
    );
    return { handle, session };
  }
  async session(handle: string | undefined): Promise<IdentitySession | null> {
    if (!handle || !/^[A-Za-z0-9_-]{43}$/.test(handle)) return null;
    const parsed = identitySessionSchema.safeParse(
      await this.#store.get(this.tenantId, "session", digest(handle)),
    );
    if (!parsed.success) return null;
    const session = parsed.data;
    if (
      session.tenantId !== this.tenantId ||
      session.expiresAt <= Date.now() ||
      !this.#config.allowedActorIds.includes(session.actorId)
    ) {
      await this.#store.delete(this.tenantId, "session", digest(handle));
      return null;
    }
    return session;
  }
  async logout(handle: string, session: IdentitySession) {
    await this.#store.delete(this.tenantId, "session", digest(handle));
    await this.#store.appendAudit(this.tenantId, session.actorId, {
      type: "session.revoked",
      occurredAt: new Date().toISOString(),
    });
  }
  async accessToken(handle: string, scopes: string[]) {
    const previous = this.#locks.get(handle) ?? Promise.resolve();
    const operation = previous
      .catch(() => {})
      .then(async () => {
        const session = await this.session(handle);
        if (!session) throw new Error("The authenticated session is unavailable.");
        const token = await this.#provider.token(session, scopes);
        const updated = await this.#store.replace(
          this.tenantId,
          "session",
          digest(handle),
          { ...session, cache: token.cache, scopes: token.scopes },
          new Date(session.expiresAt),
        );
        if (!updated) throw new Error("The authenticated session was revoked.");
        return token.accessToken;
      });
    this.#locks.set(handle, operation);
    try {
      return await operation;
    } finally {
      if (this.#locks.get(handle) === operation) this.#locks.delete(handle);
    }
  }
}

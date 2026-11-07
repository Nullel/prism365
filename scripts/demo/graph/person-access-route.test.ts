import { GraphHttpClient } from "@prism365/graph";
import { ModuleScheduler } from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import { createApp } from "../../../apps/api/src/app.ts";
import type { AppConfig, IdentityConfig } from "../../../apps/api/src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../../../apps/api/src/identity.ts";
import { createModuleRegistry } from "../../../apps/api/src/modules.ts";
import { buildSyntheticTenant, createSyntheticGraph } from "../tenant.ts";

vi.setConfig({ testTimeout: 30_000 });

// The Access route against the synthetic tenant, the way the recorder asks: the real route, the
// real Graph client and the synthetic Graph underneath. What it returns here is what the demo
// shows.
const now = new Date("2026-10-22T12:00:00Z");
const tenant = buildSyntheticTenant(now);
const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
const features = ["directory.users.read", "identity.conditional-access.read", "audit.signins.read"];
const identityConfig: IdentityConfig = {
  tenantId,
  clientId: "33333333-3333-4333-8333-333333333333",
  origin,
  redirectUri: `${origin}/api/v1/auth/callback`,
  allowedActorIds: [actorId],
  enabledFeatures: features as IdentityConfig["enabledFeatures"],
  credential: { clientSecret: "synthetic-only" },
};
const config: AppConfig = {
  mode: "live",
  origin,
  identity: identityConfig,
  production: false,
  missingConfiguration: [],
  enabledFeatures: features as AppConfig["enabledFeatures"],
};
const provider: MicrosoftIdentityProvider = {
  authorizationUrl: async () => "https://login.microsoftonline.com/synthetic",
  redeem: async () => ({
    tenantId,
    actorId,
    displayName: "Demo administrator",
    accountId: "synthetic-account",
    cache: "synthetic-cache",
    scopes: ["User.Read", "User.Read.All", "Policy.Read.All", "AuditLog.Read.All"],
  }),
  token: async (session) => ({ accessToken: "synthetic", cache: "c", scopes: session.scopes }),
};

let app: ReturnType<typeof createApp>;
let cookie = "";
const access = async (name: string) => {
  const user = tenant.users.find((candidate) => candidate.displayName === name);
  if (!user) throw new Error(`No user named ${name}.`);
  const response = await app.inject({
    method: "GET",
    url: `/v1/people/${user.id}/access`,
    headers: { cookie },
  });
  expect(response.statusCode).toBe(200);
  return response.json();
};

beforeAll(async () => {
  const store = new MemoryStateStore();
  const identity = new IdentityService(identityConfig, store, provider);
  const login = await identity.startLogin("directory.users.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
  cookie = `prism365_session=${loggedIn.handle}`;
  const registry = createModuleRegistry(config);
  const database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  const scheduler = new ModuleScheduler({
    enabledPacks: config.enabledFeatures,
    registry,
    database,
    graphClient: (getAccessToken) => new GraphHttpClient({ getAccessToken }),
  });
  app = createApp({
    config,
    store,
    identity,
    graphFetch: createSyntheticGraph({ now }),
    modules: { registry, database, scheduler },
  });
}, 60_000);
afterAll(async () => {
  await app?.close();
});

describe("the Access route against the synthetic tenant", () => {
  it("shows Tobias Vance's role, groups and the nested group by name and kind", async () => {
    const { memberships } = await access("Tobias Vance");
    expect(memberships.status).toBe("ok");
    const entries = memberships.entries as {
      kind: string;
      name: string | null;
      relation: string;
    }[];
    expect(entries.some((entry) => entry.kind === "hidden")).toBe(false);
    expect(entries.find((entry) => entry.kind === "role")).toMatchObject({
      name: "Helpdesk Administrator",
      roleTemplateId: "729827e3-9c14-49f7-bb1b-9608f156bbb8",
    });
    expect(entries.find((entry) => entry.name === "Larkspur staff")).toMatchObject({
      kind: "microsoft365",
      relation: "direct",
      hasGroupMailbox: true,
    });
    expect(entries.find((entry) => entry.name === "IT operations admins")).toMatchObject({
      kind: "security",
      relation: "nested",
    });
  });

  it("lists what he owns, a group and the Lab Portal application, and does not call it cut", async () => {
    const { owned } = await access("Tobias Vance");
    expect(owned.truncated).toBe(false);
    const entries = owned.entries as { kind: string; name: string }[];
    expect(entries.some((entry) => entry.kind === "group")).toBe(true);
    expect(entries.find((entry) => entry.name === "Larkspur Lab Portal")).toMatchObject({
      kind: "application",
    });
  });

  it("finds the policy that leaves the technician out of MFA, named from the data", async () => {
    const { conditionalAccess } = await access("Tobias Vance");
    expect(conditionalAccess.status).toBe("ok");
    const excluded = (conditionalAccess.policies as { effect: string }[]).filter(
      (policy) => policy.effect === "excluded",
    );
    expect(excluded.length).toBeGreaterThan(0);
  });
});

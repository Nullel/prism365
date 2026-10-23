import { GraphHttpClient } from "@prism365/graph";
import { ModuleScheduler } from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase } from "@prism365/storage";
import { expect } from "vitest";
import type { ScheduleTimers } from "../src/audit-schedule.ts";
import { createApp } from "../src/app.ts";
import { AuditRunner } from "../src/audits.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import { createModuleRegistry } from "../src/modules.ts";

// One app on its own in-memory database for the audit trend and schedule tests, with a
// synthetic tenant, administrator and accounts. Every domain ends in .example.

export const tenantId = "11111111-1111-4111-8111-111111111111";
export const actorId = "22222222-2222-4222-8222-222222222222";
export const origin = "http://localhost:5173";
const features = ["directory.users.read", "audit.signins.read"] as const;
const identityConfig: IdentityConfig = {
  tenantId,
  clientId: "33333333-3333-4333-8333-333333333333",
  origin,
  redirectUri: `${origin}/api/v1/auth/callback`,
  allowedActorIds: [actorId],
  enabledFeatures: [...features],
  credential: { clientSecret: "synthetic-only" },
};
const config: AppConfig = {
  mode: "live",
  origin,
  identity: identityConfig,
  production: false,
  missingConfiguration: [],
  enabledFeatures: [...features],
};

const daysAgo = (days: number) => new Date(Date.now() - days * 86_400_000).toISOString();
const users = [
  {
    id: "u-a1",
    displayName: "Ada Alpha",
    userPrincipalName: "ada@a.example",
    userType: "Member",
    accountEnabled: true,
    createdDateTime: daysAgo(900),
    assignedLicenses: [{ skuId: "11111111-0000-4000-8000-0000000000e3" }],
    signInActivity: { lastSuccessfulSignInDateTime: daysAgo(200) },
  },
  {
    id: "u-b1",
    displayName: "Bo Beta",
    userPrincipalName: "bo@b.example",
    userType: "Member",
    accountEnabled: false,
    createdDateTime: daysAgo(900),
    assignedLicenses: [],
    signInActivity: { lastSuccessfulSignInDateTime: daysAgo(10) },
  },
];
const registrations = [
  {
    id: "u-a1",
    userPrincipalName: "ada@a.example",
    userDisplayName: "Ada Alpha",
    isAdmin: true,
    isMfaRegistered: false,
    methodsRegistered: [],
    userType: "member",
  },
];

// Timers the test fires by hand. Each registration is kept, with whether it was cancelled.
export function manualTimers() {
  const registered: { kind: "every" | "after"; run: () => void; cancelled: boolean }[] = [];
  const add = (kind: "every" | "after", run: () => void) => {
    const entry = { kind, run, cancelled: false };
    registered.push(entry);
    return () => {
      entry.cancelled = true;
    };
  };
  const timers: ScheduleTimers = {
    every: (_milliseconds, run) => add("every", run),
    after: (_milliseconds, run) => add("after", run),
  };
  return {
    timers,
    registered,
    active: () => registered.filter((entry) => !entry.cancelled),
    // Fires every live timer of a kind once; an "after" timer is spent once fired.
    fire(kind: "every" | "after") {
      for (const entry of registered.filter((item) => item.kind === kind && !item.cancelled)) {
        if (kind === "after") entry.cancelled = true;
        entry.run();
      }
    },
  };
}

export interface Session {
  cookie: string;
  csrf: string;
}

export async function auditHarness(
  options: { withModules?: boolean; scopes?: string[]; now?: () => Date } = {},
) {
  const scopes = options.scopes ?? ["User.Read", "User.Read.All", "AuditLog.Read.All"];
  // Set to make Microsoft refuse every token from now on, as a sign-in frequency policy does.
  let refuseTokens = false;
  const provider: MicrosoftIdentityProvider = {
    authorizationUrl: async () => "https://login.microsoftonline.com/synthetic",
    redeem: async () => ({
      tenantId,
      actorId,
      displayName: "Synthetic admin",
      accountId: "synthetic-account",
      cache: "synthetic-cache",
      scopes: [...scopes],
    }),
    token: async (session) => {
      if (refuseTokens) {
        const error = new Error("Synthetic refusal.");
        error.name = "InteractionRequiredAuthError";
        throw error;
      }
      return { accessToken: "synthetic", cache: "c", scopes: session.scopes };
    },
  };
  const store = new MemoryStateStore();
  // Each Graph read, with how many audit.run.started ledger entries existed when it was made.
  const reads: { path: string; ledgerStarts: number }[] = [];
  const graphFetch = (async (input: RequestInfo | URL) => {
    const url = new URL(String(input));
    reads.push({
      path: url.pathname,
      ledgerStarts: store.audit.filter((entry) => entry.event.type === "audit.run.started")
        .length,
    });
    if (url.pathname === "/v1.0/users") return Response.json({ value: users });
    if (url.pathname === "/v1.0/reports/authenticationMethods/userRegistrationDetails")
      return Response.json({ value: registrations });
    return new Response("{}", { status: 404 });
  }) as typeof fetch;
  const identity = new IdentityService(identityConfig, store, provider);
  const signIn = async (): Promise<Session> => {
    const login = await identity.startLogin("directory.users.read");
    const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
    return {
      cookie: `prism365_session=${loggedIn.handle}`,
      csrf: (await identity.session(loggedIn.handle))?.csrfToken ?? "",
    };
  };
  const who = await signIn();
  const registry = createModuleRegistry(config);
  const database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  const scheduler = new ModuleScheduler({
    registry,
    database,
    enabledPacks: config.enabledFeatures,
    graphClient: (getAccessToken) => new GraphHttpClient({ getAccessToken }),
  });
  const runner = new AuditRunner({
    database,
    timeoutMilliseconds: 5_000,
    ...(options.now ? { now: options.now } : {}),
  });
  const timers = manualTimers();
  const logs: string[] = [];
  const app = createApp({
    config,
    store,
    identity,
    graphFetch,
    audits: runner,
    auditSchedule: {
      timers: timers.timers,
      log: (message) => logs.push(message),
      ...(options.now ? { now: options.now } : {}),
    },
    ...(options.withModules === false ? {} : { modules: { registry, database, scheduler } }),
  });
  const get = (url: string, session: Session | null = who) =>
    app.inject({ method: "GET", url, headers: session ? { cookie: session.cookie } : {} });
  const post = (url: string, body?: unknown, session: Session | null = who, csrf = true) =>
    app.inject({
      method: "POST",
      url,
      headers: {
        ...(session ? { cookie: session.cookie, origin } : {}),
        ...(session && csrf ? { "x-csrf-token": session.csrf } : {}),
      },
      ...(body === undefined ? {} : { payload: body as object }),
    });
  // Runs an audit by hand to the end and returns its id.
  const finish = async (auditId: string) => {
    const response = await post(`/v1/audits/${auditId}/runs`);
    expect(response.statusCode, auditId).toBe(202);
    await runner.idle();
    return response.json().run.id as string;
  };
  return {
    app,
    store,
    identity,
    database,
    scheduler,
    runner,
    timers,
    logs,
    reads,
    who,
    signIn,
    get,
    post,
    finish,
    refuseTokens: (value: boolean) => {
      refuseTokens = value;
    },
  };
}

export type AuditHarness = Awaited<ReturnType<typeof auditHarness>>;

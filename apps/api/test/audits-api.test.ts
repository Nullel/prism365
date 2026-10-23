import { randomUUID } from "node:crypto";
import {
  finishRun,
  getRun,
  guestAccountsAudit,
  insertRun,
  KEEP_FAILED_RUNS,
  KEEP_RUNS,
  licenseReclaimAudit,
  listRuns,
  pruneRuns,
} from "@prism365/audits";
import { GraphHttpClient } from "@prism365/graph";
import { ModuleScheduler } from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { createApp } from "../src/app.ts";
import { AuditRunner } from "../src/audits.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import { createModuleRegistry } from "../src/modules.ts";

// Synthetic tenant, administrators and accounts only.
const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
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

let nextScopes = ["User.Read", "User.Read.All", "AuditLog.Read.All"];
const provider: MicrosoftIdentityProvider = {
  authorizationUrl: async () => "https://login.microsoftonline.com/synthetic",
  redeem: async () => ({
    tenantId,
    actorId,
    displayName: "Synthetic admin",
    accountId: "synthetic-account",
    cache: "synthetic-cache",
    scopes: [...nextScopes],
  }),
  token: async (session) => ({ accessToken: "synthetic", cache: "c", scopes: session.scopes }),
};

const daysAgo = (days: number) => new Date(Date.now() - days * 86_400_000).toISOString();
const users = [
  {
    id: "u-1",
    displayName: "Avery Active",
    userPrincipalName: "avery@example.invalid",
    userType: "Member",
    accountEnabled: true,
    createdDateTime: daysAgo(500),
    assignedLicenses: [],
    signInActivity: { lastSuccessfulSignInDateTime: daysAgo(1) },
  },
  {
    id: "u-2",
    displayName: '=HYPERLINK("http://example.invalid")',
    userPrincipalName: "sam@example.invalid",
    userType: "Member",
    accountEnabled: false,
    createdDateTime: daysAgo(500),
    assignedLicenses: [{ skuId: "11111111-0000-4000-8000-0000000000e3" }],
    signInActivity: { lastSuccessfulSignInDateTime: daysAgo(200) },
  },
];

let gate: Promise<void> | null = null;
let refuseSignInDates = false;
let hangGuests = false;
const graphFetch = (async (input: RequestInfo | URL) => {
  const url = new URL(String(input));
  if (url.pathname === "/v1.0/users") {
    if (gate) await gate;
    if (hangGuests) return new Promise<Response>(() => {});
    if (refuseSignInDates && url.searchParams.get("$select")?.includes("signInActivity"))
      return new Response("{}", { status: 403, headers: { "request-id": "synthetic-request-1" } });
    return Response.json({ value: users });
  }
  return new Response("{}", { status: 404 });
}) as typeof fetch;

let app: ReturnType<typeof createApp>;
let store: MemoryStateStore;
let database: PGliteModuleDatabase;
let runner: AuditRunner;
let identity: IdentityService;
let full: { cookie: string; csrf: string };
let limited: { cookie: string; csrf: string };

async function signIn(scopes: string[]) {
  nextScopes = scopes;
  const login = await identity.startLogin("directory.users.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
  return {
    cookie: `prism365_session=${loggedIn.handle}`,
    csrf: (await identity.session(loggedIn.handle))?.csrfToken ?? "",
  };
}

beforeAll(async () => {
  store = new MemoryStateStore();
  identity = new IdentityService(identityConfig, store, provider);
  full = await signIn(["User.Read", "User.Read.All", "AuditLog.Read.All"]);
  limited = await signIn(["User.Read", "User.Read.All"]);
  const registry = createModuleRegistry(config);
  database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  const scheduler = new ModuleScheduler({
    enabledPacks: config.enabledFeatures,
    registry,
    database,
    graphClient: (getAccessToken) => new GraphHttpClient({ getAccessToken }),
  });
  runner = new AuditRunner({ database, timeoutMilliseconds: 1_500 });
  app = createApp({
    config,
    store,
    identity,
    graphFetch,
    audits: runner,
    modules: { registry, database, scheduler },
  });
}, 60_000);

afterAll(async () => {
  await app?.close();
});

const get = (url: string, who = full) =>
  app.inject({ method: "GET", url, headers: { cookie: who.cookie } });
const run = (auditId: string, who = full, withCsrf = true) =>
  app.inject({
    method: "POST",
    url: `/v1/audits/${auditId}/runs`,
    headers: { cookie: who.cookie, origin, ...(withCsrf ? { "x-csrf-token": who.csrf } : {}) },
  });

describe("audits API", () => {
  it("requires a session", async () => {
    for (const url of [
      "/v1/audits",
      "/v1/audits/accounts.inactive/runs",
      "/v1/audit-runs/44444444-4444-4444-8444-444444444444",
      "/v1/audit-runs/44444444-4444-4444-8444-444444444444/export.csv",
    ]) {
      const response = await app.inject({ method: "GET", url });
      expect(response.statusCode, url).toBe(401);
    }
  });

  it("lists every audit with what the session and installation allow", async () => {
    const response = await get("/v1/audits", limited);
    expect(response.statusCode).toBe(200);
    const audits = response.json().audits as {
      definition: { id: string };
      availability: {
        state: string;
        missing: string[];
        grantFeature: string | null;
        enable: string[];
      };
    }[];
    const by = (id: string) => audits.find((entry) => entry.definition.id === id)?.availability;
    expect(audits.map((entry) => entry.definition.id)).toEqual([
      "accounts.inactive",
      "licenses.reclaim",
      "mfa.registration",
      "guests.review",
      "licenses.capacity",
      "identity.conditional-access",
      "identity.risky-users",
    ]);
    expect(by("accounts.inactive")).toEqual({
      state: "missing_permission",
      missing: ["AuditLog.Read.All"],
      grantFeature: "audit.signins.read",
      enable: [],
    });
    expect(by("licenses.reclaim")?.state).toBe("available");
    expect(by("licenses.capacity")).toMatchObject({
      state: "not_enabled",
      enable: ["tenant.licenses.read"],
    });
  });

  it("refuses runs without CSRF, for unknown audits, and for permissions not granted", async () => {
    expect((await run("accounts.inactive", full, false)).statusCode).toBe(403);
    expect((await run("accounts.nothing")).statusCode).toBe(404);
    const missing = await run("accounts.inactive", limited);
    expect(missing.statusCode).toBe(403);
    expect(missing.json()).toMatchObject({ error: "missing_consent" });
    expect((await run("licenses.capacity")).json()).toMatchObject({ error: "feature_not_enabled" });
  });

  it("does not let a permission another enabled pack carries open an audit whose own pack is off", async () => {
    // The full session granted AuditLog.Read.All for the enabled sign-in pack; the MFA audit's
    // own pack (reports.authentication-methods.read) is not enabled here.
    const catalog = (await get("/v1/audits")).json().audits as {
      definition: { id: string };
      availability: { state: string; enable: string[] };
    }[];
    expect(
      catalog.find((entry) => entry.definition.id === "mfa.registration")?.availability,
    ).toMatchObject({
      state: "not_enabled",
      enable: ["reports.authentication-methods.read"],
    });
    const refused = await run("mfa.registration");
    expect(refused.statusCode).toBe(403);
    expect(refused.json()).toMatchObject({ error: "feature_not_enabled" });
  });

  it("runs an audit once at a time, then keeps its result", async () => {
    let release = () => {};
    gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    const first = await run("accounts.inactive");
    expect(first.statusCode).toBe(202);
    const started = first.json().run;
    expect(started).toMatchObject({ status: "running", startedBy: { name: "Synthetic admin" } });
    const again = await run("accounts.inactive");
    expect(again.statusCode).toBe(200);
    expect(again.json()).toMatchObject({ started: false, run: { id: started.id } });
    const catalog = (await get("/v1/audits")).json();
    expect(
      catalog.audits.find(
        (entry: { definition: { id: string } }) => entry.definition.id === "accounts.inactive",
      ).runningId,
    ).toBe(started.id);
    const polled = await get(`/v1/audit-runs/${started.id}`);
    expect(polled.json()).toMatchObject({ status: "running", rows: [] });
    release();
    await runner.idle();
    gate = null;

    const done = await get(`/v1/audit-runs/${started.id}`);
    expect(done.json()).toMatchObject({ status: "succeeded", complete: true, rowCount: 2 });
    const events = store.audit.map((entry) => entry.event.type);
    expect(events.filter((type) => type === "audit.run.started")).toHaveLength(1);
    // Polling a running run is not a view; reading the result is.
    expect(events.filter((type) => type === "audit.run.viewed")).toHaveLength(1);
    const history = await get("/v1/audits/accounts.inactive/runs");
    expect(history.json().runs.map((item: { id: string }) => item.id)).toEqual([started.id]);
  });

  it("exports the filtered rows as CSV, with formulas kept as text", async () => {
    const [latest] = (await get("/v1/audits/accounts.inactive/runs")).json().runs;
    const response = await get(
      `/v1/audit-runs/${latest.id}/export.csv?view=d90&sort=days&direction=desc&f.status=Disabled`,
    );
    expect(response.statusCode).toBe(200);
    expect(response.headers["content-type"]).toContain("text/csv");
    expect(response.headers["content-disposition"]).toMatch(
      /^attachment; filename="prism365-accounts\.inactive-\d{4}-\d{2}-\d{2}\.csv"$/,
    );
    const lines = response.body.trimEnd().split("\r\n");
    expect(lines).toHaveLength(2);
    expect(lines[1]).toContain(
      `"'=HYPERLINK(""http://example.invalid"")","sam@example.invalid",200`,
    );
    expect(store.audit.at(-1)?.event).toMatchObject({ type: "audit.run.exported", rows: 1 });
  });

  it("records Microsoft's refusal in plain words", async () => {
    refuseSignInDates = true;
    const started = (await run("accounts.inactive")).json().run;
    await runner.idle();
    refuseSignInDates = false;
    const failed = (await get(`/v1/audit-runs/${started.id}`)).json();
    expect(failed).toMatchObject({ status: "failed", rowCount: 0 });
    expect(failed.error.code).toBe("forbidden");
    expect(failed.error.message).toContain("Microsoft Entra ID P1 or P2");
    expect(failed.error.message).toContain("synthetic-request-1");
    const csv = await get(`/v1/audit-runs/${started.id}/export.csv`);
    expect(csv.statusCode).toBe(404);
  });

  it("stops a run at its time limit", async () => {
    hangGuests = true;
    const started = (await run("guests.review")).json().run;
    await runner.idle();
    hangGuests = false;
    const stopped = (await get(`/v1/audit-runs/${started.id}`)).json();
    expect(stopped).toMatchObject({ status: "failed", error: { code: "timeout" } });
  }, 10_000);
});

describe("audit runner", () => {
  const client = () =>
    new GraphHttpClient({
      getAccessToken: async () => "synthetic",
      fetchImplementation: graphFetch,
    });
  const actor = { id: actorId, name: "Synthetic admin" };

  it("closes runs a stopped process left open, and only those", async () => {
    const orphan = "55555555-5555-4555-8555-555555555555";
    await database.tenant(tenantId, (sql) =>
      insertRun(sql, tenantId, {
        id: orphan,
        auditId: guestAccountsAudit.definition.id,
        auditVersion: 1,
        startedAt: new Date(),
        startedBy: actor,
      }),
    );
    const restarted = new AuditRunner({ database });
    await restarted.recover(tenantId);
    const closed = await database.tenant(tenantId, (sql) => getRun(sql, tenantId, orphan));
    expect(closed).toMatchObject({ status: "failed", error: { code: "interrupted" } });
  });

  it(`keeps the newest ${KEEP_RUNS} finished runs of each audit`, async () => {
    const fresh = new AuditRunner({ database });
    for (let index = 0; index < KEEP_RUNS + 3; index += 1) {
      await fresh.start({
        tenantId,
        actor,
        audit: licenseReclaimAudit,
        client: client(),
        packs: [...licenseReclaimAudit.packs, ...licenseReclaimAudit.optionalPacks],
        granted: ["User.Read.All", "AuditLog.Read.All"],
      });
      await fresh.idle();
    }
    const runs = await database.tenant(tenantId, (sql) =>
      listRuns(sql, tenantId, licenseReclaimAudit.definition.id, 50),
    );
    expect(runs).toHaveLength(KEEP_RUNS);
    expect(runs.every((item) => item.status === "succeeded")).toBe(true);
  });

  it("records shutdown as an interruption", async () => {
    const closing = new AuditRunner({ database });
    let release = () => {};
    gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    const { runId } = await closing.start({
      tenantId,
      actor,
      audit: guestAccountsAudit,
      client: client(),
      packs: [...guestAccountsAudit.packs, ...guestAccountsAudit.optionalPacks],
      granted: ["User.Read.All", "AuditLog.Read.All"],
    });
    await closing.close();
    release();
    gate = null;
    const stopped = await database.tenant(tenantId, (sql) => getRun(sql, tenantId, runId));
    expect(stopped).toMatchObject({ status: "failed", error: { code: "interrupted" } });
  });
});

describe("audit review fixes", () => {
  const actor = { id: actorId, name: "Synthetic admin" };

  it("never lets failures push out the last good result, and prunes one audit only", async () => {
    const other = "77777777-7777-4777-8777-777777777777";
    const running = "88888888-8888-4888-8888-888888888888";
    const fresh = new AuditRunner({ database });
    await fresh.recover(tenantId);
    const at = (minutes: number) => new Date(Date.UTC(2026, 0, 1, 0, minutes));
    await database.tenant(tenantId, async (sql) => {
      // Another audit's run, and a run still in progress, must survive pruning.
      await insertRun(sql, tenantId, {
        id: other,
        auditId: "licenses.capacity",
        auditVersion: 1,
        startedAt: at(0),
        startedBy: actor,
      });
      await finishRun(sql, tenantId, other, at(1), { status: "failed", code: "x", message: "x" });
      await insertRun(sql, tenantId, {
        id: running,
        auditId: "mfa.registration",
        auditVersion: 1,
        startedAt: at(0),
        startedBy: actor,
      });
      const good = randomUUID();
      await insertRun(sql, tenantId, {
        id: good,
        auditId: "mfa.registration",
        auditVersion: 1,
        startedAt: at(2),
        startedBy: actor,
      });
      await finishRun(sql, tenantId, good, at(3), {
        status: "succeeded",
        complete: true,
        summary: [],
        notes: [],
        rows: [],
      });
      // More failures than successful runs are kept, so a quota shared between the two would
      // push the good result out.
      for (let index = 0; index < 12; index += 1) {
        const id = randomUUID();
        await insertRun(sql, tenantId, {
          id,
          auditId: "mfa.registration",
          auditVersion: 1,
          startedAt: at(10 + index),
          startedBy: actor,
        });
        await finishRun(sql, tenantId, id, at(11 + index), {
          status: "failed",
          code: "throttled",
          message: "Throttled.",
        });
      }
      await pruneRuns(sql, tenantId, "mfa.registration");
    });
    const kept = await database.tenant(tenantId, (sql) =>
      listRuns(sql, tenantId, "mfa.registration", 50),
    );
    expect(kept.filter((item) => item.status === "succeeded")).toHaveLength(1);
    expect(kept.filter((item) => item.status === "failed")).toHaveLength(KEEP_FAILED_RUNS);
    expect(kept.map((item) => item.id)).toContain(running);
    const untouched = await database.tenant(tenantId, (sql) => getRun(sql, tenantId, other));
    expect(untouched?.status).toBe("failed");
  });

  it("stops reading Microsoft once a run is stopped", async () => {
    let requests = 0;
    // Microsoft paging forever, slowly; a real fetch rejects when its signal aborts.
    const endless = (async (input: RequestInfo | URL, init?: RequestInit) => {
      requests += 1;
      await new Promise<void>((resolve, reject) => {
        const timer = setTimeout(resolve, 40);
        init?.signal?.addEventListener("abort", () => {
          clearTimeout(timer);
          reject(new DOMException("Aborted", "AbortError"));
        });
      });
      const url = new URL(String(input));
      const page = Number(url.searchParams.get("$skiptoken") ?? "0");
      url.searchParams.set("$skiptoken", String(page + 1));
      return Response.json({ value: [], "@odata.nextLink": url.toString() });
    }) as typeof fetch;
    const quick = new AuditRunner({ database, timeoutMilliseconds: 300 });
    const { runId } = await quick.start({
      tenantId,
      actor,
      audit: guestAccountsAudit,
      client: new GraphHttpClient({
        getAccessToken: async () => "synthetic",
        fetchImplementation: endless,
      }),
      packs: [...guestAccountsAudit.packs, ...guestAccountsAudit.optionalPacks],
      granted: ["User.Read.All"],
    });
    await quick.idle();
    const stopped = await database.tenant(tenantId, (sql) => getRun(sql, tenantId, runId));
    expect(stopped).toMatchObject({ status: "failed", error: { code: "timeout" } });
    const afterStop = requests;
    await new Promise((resolve) => setTimeout(resolve, 300));
    expect(requests).toBe(afterStop);
  });

  it("writes one ledger entry per run, naming every dataset it reads", async () => {
    let release = () => {};
    gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    const before = store.audit.length;
    const [first, second] = await Promise.all([run("licenses.reclaim"), run("licenses.reclaim")]);
    release();
    await runner.idle();
    gate = null;
    const ids = new Set([first.json().run.id, second.json().run.id]);
    expect(ids.size).toBe(1);
    const started = store.audit
      .slice(before)
      .filter((entry) => entry.event.type === "audit.run.started");
    expect(started).toHaveLength(1);
    // Reclaim needs User.Read.All and also reads sign-in dates when the session granted them.
    expect(started[0]?.event).toMatchObject({
      auditId: "licenses.reclaim",
      runId: first.json().run.id,
      datasets: ["User.Read.All", "AuditLog.Read.All"],
    });
  });
});

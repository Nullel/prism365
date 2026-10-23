import { randomUUID } from "node:crypto";
import { GraphHttpClient } from "@prism365/graph";
import { createSigninsAuditModule } from "@prism365/module-signins-audit";
import { ModuleRegistry, ModuleScheduler, upsertFinding } from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";

// Synthetic tenant, accounts and documentation-range addresses only.
const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const personId = "44444444-4444-4444-8444-444444444444";
const otherId = "55555555-5555-4555-8555-555555555555";
const origin = "http://localhost:5173";
const features = ["directory.users.read", "user.licenses.read", "audit.signins.read"] as const;
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
const provider: MicrosoftIdentityProvider = {
  authorizationUrl: async () => "https://login.microsoftonline.com/synthetic",
  redeem: async () => ({
    tenantId,
    actorId,
    displayName: "Synthetic admin",
    accountId: "synthetic-account",
    cache: "synthetic-cache",
    scopes: ["User.Read", "User.Read.All", "LicenseAssignment.Read.All", "AuditLog.Read.All"],
  }),
  token: async (session) => ({ accessToken: "synthetic", cache: "c", scopes: session.scopes }),
};

// Microsoft's directory for the one synthetic person: a profile, no manager, one license.
const graphRequests: string[] = [];
async function graphFetch(input: RequestInfo | URL): Promise<Response> {
  const url = new URL(String(input));
  graphRequests.push(`${url.pathname}${url.search}`);
  const path = decodeURIComponent(url.pathname);
  if (path === `/v1.0/users('person@example.invalid')`) return Response.json({ id: personId });
  if (path === `/v1.0/users/${personId}/manager`) return new Response("{}", { status: 404 });
  if (path === `/v1.0/users/${personId}/licenseDetails`)
    return Response.json({
      value: [
        {
          id: "synthetic-license",
          skuId: "synthetic-sku",
          skuPartNumber: "SPB",
          servicePlans: [
            {
              servicePlanId: "synthetic-plan",
              servicePlanName: "EXCHANGE_S_STANDARD",
              provisioningStatus: "Success",
            },
          ],
        },
      ],
    });
  if (path === `/v1.0/users/${personId}` && url.searchParams.get("$select") === "signInActivity")
    return Response.json({
      signInActivity: {
        lastSignInDateTime: "2026-10-21T09:00:00Z",
        lastNonInteractiveSignInDateTime: "2026-10-21T10:00:00Z",
        lastSuccessfulSignInDateTime: "2026-10-21T09:00:00Z",
      },
    });
  if (path === `/v1.0/users/${personId}`)
    return Response.json({
      id: personId,
      displayName: "Synthetic Person",
      userPrincipalName: "person@example.invalid",
      mail: "person@example.invalid",
      jobTitle: "Analyst",
      department: "Research",
      accountEnabled: true,
      createdDateTime: "2024-01-02T03:04:05Z",
      onPremisesSyncEnabled: null,
      businessPhones: ["+1 555 0100"],
    });
  return new Response("{}", { status: 404 });
}

let app: ReturnType<typeof createApp>;
let store: MemoryStateStore;
let database: PGliteModuleDatabase;
let cookie: string;

async function insertEvent(values: {
  userId: string;
  at: string;
  ip: string;
  errorCode?: number;
  city?: string;
  country?: string;
  device?: string | null;
}) {
  await database.tenant(tenantId, (sql) =>
    sql.query(
      `INSERT INTO mod_signins_audit_event (tenant_id, id, created_at, user_id, upn, display_name,
         app_name, client_app, ip, country, state, city, device_id, device_os, device_browser,
         is_compliant, is_managed, error_code, failure_reason, risk_event_types, backfilled,
         processed, collected_at)
       VALUES ($1,$2,$3,$4,$5,$6,'Office 365 Exchange Online','Browser',$7,$8,'Illinois',$9,$10,
         'Windows','Edge',true,true,$11,$12,'[]'::jsonb,false,true,now())`,
      [
        tenantId,
        randomUUID(),
        values.at,
        values.userId,
        values.userId === personId ? "person@example.invalid" : "other@example.invalid",
        values.userId === personId ? "Synthetic Person" : "Other Synthetic",
        values.ip,
        values.country ?? "US",
        values.city ?? "Chicago",
        values.device === undefined ? "device-1" : values.device,
        values.errorCode ?? 0,
        values.errorCode ? "Invalid username or password." : null,
      ],
    ),
  );
}

beforeAll(async () => {
  store = new MemoryStateStore();
  const identity = new IdentityService(identityConfig, store, provider);
  const login = await identity.startLogin("audit.signins.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
  cookie = `prism365_session=${loggedIn.handle}`;
  const registry = new ModuleRegistry().register(createSigninsAuditModule());
  database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  const scheduler = new ModuleScheduler({
    enabledPacks: config.enabledFeatures,
    registry,
    database,
    graphClient: (getAccessToken) =>
      new GraphHttpClient({ getAccessToken, fetchImplementation: graphFetch }),
  });
  app = createApp({
    config,
    store,
    identity,
    graphFetch,
    modules: { registry, database, scheduler },
  });

  const now = Date.now();
  const ago = (hours: number) => new Date(now - hours * 3_600_000).toISOString();
  for (let hour = 48; hour >= 2; hour -= 2)
    await insertEvent({ userId: personId, at: ago(hour), ip: "192.0.2.10" });
  await insertEvent({
    userId: personId,
    at: ago(1),
    ip: "203.0.113.47",
    errorCode: 50126,
    city: "Amsterdam",
    country: "NL",
    device: null,
  });
  await insertEvent({ userId: otherId, at: ago(3), ip: "198.51.100.5" });
  await database.tenant(tenantId, async (sql) => {
    for (const subject of [personId, otherId])
      await upsertFinding(sql, tenantId, "signins.audit", {
        ruleId: "signins.failed_burst_account",
        ruleVersion: 1,
        dedupeKey: subject,
        purpose: "security",
        severity: "medium",
        score: 50,
        subject: { kind: "person", id: subject, label: "Synthetic" },
        title: "Repeated failed sign-ins for one account",
        summary: "Five failed sign-ins within 15 minutes.",
        evidence: [
          {
            source: "Microsoft Entra sign-in logs",
            ref: `synthetic-${subject}`,
            observedAt: new Date(now).toISOString(),
            fields: { network: "203.0.113.0/24" },
          },
        ],
        limitations: ["Synthetic."],
        window: { start: new Date(now - 3_600_000), end: new Date(now) },
        observedAt: new Date(now),
      });
  });
}, 60_000);

afterAll(async () => {
  // Closing the app closes the module database too.
  await app?.close();
});

describe("people API", () => {
  it("requires a session on every people route", async () => {
    for (const url of [
      "/v1/people",
      `/v1/people/${personId}`,
      `/v1/people/${personId}/signins`,
      `/v1/people/${personId}/directory`,
      "/v1/people/lookup?email=person@example.invalid",
    ]) {
      const response = await app.inject({ method: "GET", url });
      expect(response.statusCode, url).toBe(401);
    }
  });

  it("lists accounts with their counts and filters by search", async () => {
    const response = await app.inject({ method: "GET", url: "/v1/people", headers: { cookie } });
    expect(response.statusCode).toBe(200);
    const body = response.json();
    expect(body.items.map((item: { userId: string }) => item.userId)).toEqual([personId, otherId]);
    expect(body.items[0]).toMatchObject({
      userPrincipalName: "person@example.invalid",
      displayName: "Synthetic Person",
      signIns: 25,
      failures: 1,
      countries: 2,
      openFindings: 1,
    });
    const filtered = await app.inject({
      method: "GET",
      url: "/v1/people?search=OTHER",
      headers: { cookie },
    });
    expect(filtered.json().items.map((item: { userId: string }) => item.userId)).toEqual([otherId]);
    const none = await app.inject({
      method: "GET",
      url: "/v1/people?search=nobody",
      headers: { cookie },
    });
    expect(none.json().items).toEqual([]);
  });

  it("returns one person's places, full addresses, devices and only their findings", async () => {
    const response = await app.inject({
      method: "GET",
      url: `/v1/people/${personId}`,
      headers: { cookie },
    });
    expect(response.statusCode).toBe(200);
    const { history, findings } = response.json();
    expect(history.totals).toMatchObject({ signIns: 25, successes: 24, failures: 1 });
    expect(history.addresses.map((row: { ip: string }) => row.ip).sort()).toEqual([
      "192.0.2.10",
      "203.0.113.47",
    ]);
    expect(history.places).toContainEqual(
      expect.objectContaining({ city: "Amsterdam", country: "NL", failures: 1, successes: 0 }),
    );
    expect(history.devices).toContainEqual(
      expect.objectContaining({ deviceId: "device-1", compliant: true, count: 24 }),
    );
    expect(findings.map((finding: { subject: { id: string } }) => finding.subject.id)).toEqual([
      personId,
    ]);
    expect(store.audit.at(-1)?.event).toMatchObject({ type: "person.viewed" });
  });

  it("returns null history for an account with no stored sign-ins", async () => {
    const response = await app.inject({
      method: "GET",
      url: "/v1/people/66666666-6666-4666-8666-666666666666",
      headers: { cookie },
    });
    expect(response.json()).toEqual({ history: null, findings: [] });
  });

  it("pages sign-ins newest first and filters unsuccessful ones", async () => {
    const first = await app.inject({
      method: "GET",
      url: `/v1/people/${personId}/signins?limit=10`,
      headers: { cookie },
    });
    const page = first.json();
    expect(page.items).toHaveLength(10);
    expect(page.items[0]).toMatchObject({
      ip: "203.0.113.47",
      result: "failure",
      errorCode: 50126,
    });
    expect(page.next).toEqual(expect.any(String));
    const second = await app.inject({
      method: "GET",
      url: `/v1/people/${personId}/signins?limit=10&before=${encodeURIComponent(page.next)}`,
      headers: { cookie },
    });
    const ids = new Set(page.items.map((item: { id: string }) => item.id));
    const more = second.json().items as { id: string; createdAt: string }[];
    expect(more).toHaveLength(10);
    expect(more.some((item) => ids.has(item.id))).toBe(false);
    expect(Date.parse(more[0]?.createdAt ?? "")).toBeLessThan(
      Date.parse(page.items.at(-1)?.createdAt ?? ""),
    );
    const failed = await app.inject({
      method: "GET",
      url: `/v1/people/${personId}/signins?result=failed`,
      headers: { cookie },
    });
    expect(failed.json().items.map((item: { ip: string }) => item.ip)).toEqual(["203.0.113.47"]);
    const bad = await app.inject({
      method: "GET",
      url: `/v1/people/${personId}/signins?before=not-a-position`,
      headers: { cookie },
    });
    expect(bad.statusCode).toBe(400);
  });

  it("rejects account identifiers that are not directory IDs", async () => {
    const response = await app.inject({
      method: "GET",
      url: "/v1/people/a'b",
      headers: { cookie },
    });
    expect(response.statusCode).toBe(400);
  });

  it("reads the directory, manager, last sign-in and licenses live", async () => {
    const response = await app.inject({
      method: "GET",
      url: `/v1/people/${personId}/directory`,
      headers: { cookie },
    });
    expect(response.statusCode).toBe(200);
    expect(response.json()).toMatchObject({
      state: "available",
      profile: {
        displayName: "Synthetic Person",
        jobTitle: "Analyst",
        accountEnabled: true,
        businessPhones: ["+1 555 0100"],
        onPremisesSyncEnabled: null,
      },
      manager: { state: "available", data: null },
      signInActivity: { state: "available", lastSignIn: "2026-10-21T09:00:00Z" },
      licenses: { state: "available", data: [expect.objectContaining({ skuPartNumber: "SPB" })] },
    });
    expect(store.audit.at(-1)?.event).toMatchObject({ type: "person.directory.viewed" });
  });

  it("looks a work email up in the directory", async () => {
    const found = await app.inject({
      method: "GET",
      url: "/v1/people/lookup?email=person@example.invalid",
      headers: { cookie },
    });
    expect(found.json()).toEqual({ userId: personId });
    const missing = await app.inject({
      method: "GET",
      url: "/v1/people/lookup?email=missing@example.invalid",
      headers: { cookie },
    });
    expect(missing.statusCode).toBe(404);
    const invalid = await app.inject({
      method: "GET",
      url: "/v1/people/lookup?email=not-an-email",
      headers: { cookie },
    });
    expect(invalid.statusCode).toBe(400);
  });
  it("cuts an over-long Microsoft failure reason instead of failing the page", async () => {
    await insertEvent({
      userId: otherId,
      at: new Date().toISOString(),
      ip: "198.51.100.6",
      errorCode: 50126,
    });
    await database.tenant(tenantId, (sql) =>
      sql.query(
        "UPDATE mod_signins_audit_event SET failure_reason = $2 WHERE tenant_id = $1 AND ip = '198.51.100.6'",
        [tenantId, "x".repeat(900)],
      ),
    );
    const response = await app.inject({
      method: "GET",
      url: `/v1/people/${otherId}/signins`,
      headers: { cookie },
    });
    expect(response.statusCode).toBe(200);
    expect(response.json().items[0].failureReason).toHaveLength(500);
  });
});

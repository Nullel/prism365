import { MemoryStateStore } from "@prism365/storage";
import { describe, expect, it, vi } from "vitest";
import type { AiProvider } from "@prism365/ai";
import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";

const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
class SyntheticDurableStore extends MemoryStateStore {
  override readonly durable = true;
}

async function setup() {
  const identityConfig: IdentityConfig = {
    tenantId,
    clientId: "33333333-3333-4333-8333-333333333333",
    origin: "http://localhost:5173",
    redirectUri: "http://localhost:5173/api/v1/auth/callback",
    allowedActorIds: [actorId],
    enabledFeatures: ["directory.users.read"],
    credential: { clientSecret: "synthetic-only" },
  };
  const config: AppConfig = {
    mode: "live",
    origin: identityConfig.origin,
    identity: identityConfig,
    production: false,
    missingConfiguration: [],
    enabledFeatures: identityConfig.enabledFeatures,
    aiModel: "synthetic-model",
  };
  const store = new SyntheticDurableStore();
  const provider: MicrosoftIdentityProvider = {
    authorizationUrl: async () => "https://login.microsoftonline.com/synthetic",
    redeem: async () => ({
      tenantId,
      actorId,
      displayName: "Synthetic admin",
      accountId: "fake-account",
      cache: "fake-cache",
      scopes: ["User.Read", "User.Read.All"],
    }),
    token: async (session) => ({
      accessToken: "fake-token",
      cache: "fake-cache",
      scopes: session.scopes,
    }),
  };
  const identity = new IdentityService(identityConfig, store, provider);
  const login = await identity.startLogin("directory.users.read");
  const loggedIn = await identity.finishLogin("fake-code", login.state, login.state);
  const model: AiProvider = {
    descriptor: {
      id: "ollama-local:synthetic-model",
      label: "Synthetic model",
      deployment: "local",
    },
    analyze: vi.fn<AiProvider["analyze"]>(async (input) => ({
      summary: "One synthetic directory record is available.",
      citations: input.evidence.map((record) => record.evidenceId),
      limitations: ["Synthetic integration test only."],
      model: "synthetic-model",
    })),
  };
  const app = createApp({
    config,
    store,
    identity,
    aiProvider: model,
    graphFetch: async () =>
      Response.json({
        id: "fake-user",
        displayName: "Synthetic person",
        userPrincipalName: "alex@example.invalid",
        mail: null,
        jobTitle: null,
        department: null,
        userType: "Member",
      }),
  });
  const headers = {
    cookie: `prism365_session=${loggedIn.handle}`,
    origin: config.origin,
    "x-csrf-token": loggedIn.session.csrfToken,
  };
  return { app, headers, store, model };
}

describe("authenticated platform", () => {
  it("persists actor-bound AI grants, audits before disclosure, and respects revocation", async () => {
    const { app, headers, store, model } = await setup();
    const approved = await app.inject({
      method: "POST",
      url: "/v1/ai/grants",
      headers,
      payload: { datasets: ["directory"] },
    });
    expect(approved.statusCode).toBe(200);
    const grantId = approved.json().grantId;
    const answer = await app.inject({
      method: "POST",
      url: "/v1/ai/investigations",
      headers,
      payload: { email: "alex@example.invalid", question: "Explain these records.", grantId },
    });
    expect(answer.statusCode).toBe(200);
    expect(answer.json()).toMatchObject({ providerDeployment: "local", evidenceCount: 1 });
    expect(model.analyze).toHaveBeenCalledTimes(1);
    expect(store.audit.map((item) => item.event.type)).toEqual(
      expect.arrayContaining([
        "ai.grant.approved",
        "ai.analysis.authorized",
        "ai.analysis.completed",
      ]),
    );
    const submitted = vi.mocked(model.analyze).mock.calls[0]?.[0];
    expect(JSON.stringify(submitted)).not.toContain("alex@example.invalid");
    expect(JSON.stringify(submitted)).not.toContain("Synthetic person");
    expect(
      (await app.inject({ method: "DELETE", url: `/v1/ai/grants/${grantId}`, headers })).statusCode,
    ).toBe(200);
    expect(
      (
        await app.inject({
          method: "POST",
          url: "/v1/ai/investigations",
          headers,
          payload: { email: "alex@example.invalid", question: "Explain these records.", grantId },
        })
      ).statusCode,
    ).toBe(403);
    expect(model.analyze).toHaveBeenCalledTimes(1);
    await app.close();
  });

  it("privately stores unreviewed script source and returns metadata only", async () => {
    const { app, headers, store } = await setup();
    const payload = {
      manifest: {
        schemaVersion: 1,
        id: "example.offboard",
        name: "Example runbook",
        version: "1.0.0",
        description: "Synthetic package",
        lifecycle: "offboarding",
        runtime: "powershell7",
        entrypoint: "scripts/offboard.ps1",
        publisher: "Example publisher",
        parameters: [],
        capabilities: ["identity.lookup"],
        allowedHosts: [],
        timeoutSeconds: 30,
        steps: [
          {
            id: "lookup",
            label: "Inspect target",
            description: "Read only",
            capability: "identity.lookup",
            risk: "readOnly",
            reversible: true,
            recovery: "No changes made.",
          },
        ],
      },
      files: [{ path: "scripts/offboard.ps1", content: "Write-Output 'synthetic-private-source'" }],
    };
    const imported = await app.inject({
      method: "POST",
      url: "/v1/automation/packages",
      headers,
      payload,
    });
    expect(imported.statusCode).toBe(201);
    expect(imported.json()).toMatchObject({ status: "unreviewed", name: "Example runbook" });
    const list = await app.inject({ url: "/v1/automation/packages", headers });
    expect(list.json()).toMatchObject({
      executionEnabled: false,
      items: [{ status: "unreviewed" }],
    });
    expect(list.body).not.toContain("synthetic-private-source");
    expect(imported.body).not.toContain("synthetic-private-source");
    expect(JSON.stringify(await store.list(tenantId, "automationPackage"))).toContain(
      "synthetic-private-source",
    );
    expect(await store.list("44444444-4444-4444-8444-444444444444", "automationPackage")).toEqual(
      [],
    );
    expect(
      (
        await app.inject({
          method: "POST",
          url: "/v1/automation/runs",
          headers,
          payload: { packageId: imported.json().id },
        })
      ).statusCode,
    ).toBe(404);
    expect(
      (
        await app.inject({
          method: "DELETE",
          url: `/v1/automation/packages/${imported.json().id}`,
          headers,
        })
      ).statusCode,
    ).toBe(200);
    expect(await store.list(tenantId, "automationPackage")).toEqual([]);
    await app.close();
  });
});

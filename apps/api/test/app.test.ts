import { afterAll, describe, expect, it } from "vitest";
import { createApp } from "../src/app.ts";

const app = createApp();

afterAll(async () => {
  await app.close();
});

describe("demo API", () => {
  it("reports that no tenant is connected", async () => {
    const response = await app.inject({ method: "GET", url: "/health" });

    expect(response.statusCode).toBe(200);
    expect(response.json()).toMatchObject({ mode: "demo", tenantConnected: false });
    expect(response.headers["cache-control"]).toBe("no-store");
  });

  it("returns runtime-validated synthetic dashboard data", async () => {
    const response = await app.inject({ method: "GET", url: "/v1/dashboard/summary" });

    expect(response.statusCode).toBe(200);
    expect(response.json()).toMatchObject({ mode: "demo", state: "demo" });
  });

  it("advertises read-only capabilities", async () => {
    const response = await app.inject({ method: "GET", url: "/v1/platform/capabilities" });

    expect(response.statusCode).toBe(200);
    const body = response.json();
    expect(body).toMatchObject({ writesEnabled: false, tenantConnectionEnabled: false });
    expect(body.ai).toEqual({
      analysisEnabled: false,
      providerConfigured: false,
      writeToolsEnabled: false,
      supportedDeployments: ["local", "tenantHosted", "external"],
    });
    expect(body.cloudProfiles).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          id: "worldwide",
          graphBaseUrl: "https://graph.microsoft.com/v1.0",
        }),
        expect.objectContaining({
          id: "usGovernmentL4",
          graphBaseUrl: "https://graph.microsoft.us/v1.0",
        }),
      ]),
    );
    expect(body.permissionPacks).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          id: "identity.risk.read",
          permission: "IdentityRiskyUser.Read.All",
          licenseNotes: expect.stringContaining("P2"),
          documentationLastVerified: "2026-10-13",
        }),
      ]),
    );
  });

  it("lists lifecycle workflows without enabling import or execution", async () => {
    const response = await app.inject({ method: "GET", url: "/v1/automation/workflows" });

    expect(response.statusCode).toBe(200);
    expect(response.json()).toMatchObject({
      mode: "demo",
      importEnabled: false,
      executionEnabled: false,
      items: expect.arrayContaining([
        expect.objectContaining({ lifecycle: "onboarding" }),
        expect.objectContaining({ lifecycle: "offboarding", runtime: "powershell7" }),
      ]),
    });
  });

  it("returns a guarded preview and never exposes an execute endpoint", async () => {
    const response = await app.inject({
      method: "POST",
      url: "/v1/automation/previews",
      payload: {
        workflowId: "prism365.secure-offboarding",
        targetEmail: "alex.rivera@example.invalid",
      },
    });

    expect(response.statusCode).toBe(200);
    expect(response.json()).toMatchObject({
      mode: "demo",
      targetEmail: "alex.rivera@example.invalid",
      executionAllowed: false,
      steps: expect.arrayContaining([
        expect.objectContaining({ capability: "identity.user.disable" }),
      ]),
    });

    const executeResponse = await app.inject({
      method: "POST",
      url: "/v1/automation/runs",
      payload: { previewId: response.json().previewId },
    });
    expect(executeResponse.statusCode).toBe(404);
  });

  it("rejects an invalid lifecycle preview request", async () => {
    const response = await app.inject({
      method: "POST",
      url: "/v1/automation/previews",
      payload: { workflowId: "prism365.secure-offboarding", targetEmail: "not-an-email" },
    });

    expect(response.statusCode).toBe(400);
    expect(response.json()).toMatchObject({ error: "invalid_request" });
  });
});

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

  it("has no automation, script package or single-user investigation routes", async () => {
    for (const [method, url] of [
      ["GET", "/v1/automation/workflows"],
      ["POST", "/v1/automation/previews"],
      ["POST", "/v1/automation/runs"],
      ["GET", "/v1/automation/packages"],
      ["POST", "/v1/automation/packages"],
      ["DELETE", "/v1/automation/packages/44444444-4444-4444-8444-444444444444"],
      ["GET", "/v1/investigations/user?email=alex@example.invalid"],
    ] as const) {
      const response = await app.inject({ method, url });
      expect(response.statusCode, `${method} ${url}`).toBe(404);
    }
  });
});

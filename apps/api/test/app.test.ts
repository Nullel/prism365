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
    expect(response.json()).toMatchObject({ writesEnabled: false, tenantConnectionEnabled: false });
  });
});

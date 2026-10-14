import { userInvestigationSchema } from "@prism365/contracts";
import { describe, expect, it, vi } from "vitest";
import { GraphHttpClient } from "../src/client.ts";
import {
  createDemoUserInvestigation,
  investigateUser,
  UserInvestigationError,
} from "../src/investigation.ts";

const email = "alex.rivera@example.invalid";
const now = new Date("2026-10-13T17:00:00.000Z");
const scopes = ["User.Read.All", "LicenseAssignment.Read.All", "AuditLog.Read.All"];
const profile = {
  id: "synthetic-user",
  displayName: "Alex Rivera",
  userPrincipalName: email,
  mail: email,
  jobTitle: "Engineer",
  department: "Technology",
  userType: "Member",
};
const license = {
  id: "synthetic-license",
  skuId: "synthetic-sku",
  skuPartNumber: "DEMO_M365",
  servicePlans: [
    {
      servicePlanId: "synthetic-plan",
      servicePlanName: "DEMO_EXCHANGE",
      provisioningStatus: "Success",
    },
  ],
};
const signIn = {
  id: "synthetic-signin",
  userId: profile.id,
  createdDateTime: "2026-10-13T16:00:00.000Z",
  appDisplayName: "Demo app",
  clientAppUsed: "Browser",
  isInteractive: true,
  status: { errorCode: 500121, failureReason: "Synthetic failed MFA", additionalDetails: null },
  location: { city: "Demo city", countryOrRegion: "Demo region" },
  ipAddress: "192.0.2.1",
  arbitrarySecret: "must-never-leave-provider",
};

function clientFor(overrides: { license?: Response; signIns?: Response; user?: Response } = {}) {
  const fetchImplementation = vi.fn(async (input: RequestInfo | URL) => {
    const path = new URL(String(input)).pathname;
    if (path.endsWith("/licenseDetails"))
      return overrides.license ?? Response.json({ value: [license] });
    if (path.endsWith("/signIns")) return overrides.signIns ?? Response.json({ value: [signIn] });
    return overrides.user ?? Response.json(profile);
  });
  return {
    client: new GraphHttpClient({
      getAccessToken: async () => "synthetic-token",
      fetchImplementation,
      maxRetries: 0,
    }),
    fetchImplementation,
  };
}

describe("delegated user investigation", () => {
  it("reads one resolved user and bounded logs, strips unrelated/sensitive provider fields", async () => {
    const { client, fetchImplementation } = clientFor();
    const result = await investigateUser({ email, client, grantedScopes: scopes, now });
    expect(userInvestigationSchema.safeParse(result).success).toBe(true);
    expect(result).toMatchObject({
      mode: "live",
      profile: { state: "available", data: profile },
      licenses: { state: "available", data: [license] },
      signIns: { state: "available", complete: true },
    });
    expect(result.signIns.data[0]?.status.errorCode).toBe(500121);
    expect(JSON.stringify(result)).not.toContain("192.0.2.1");
    expect(JSON.stringify(result)).not.toContain("must-never-leave-provider");
    expect(fetchImplementation).toHaveBeenCalledTimes(3);
    const logUrl = fetchImplementation.mock.calls
      .map(([url]) => new URL(String(url)))
      .find((url) => url.pathname.endsWith("/signIns"));
    expect(logUrl?.searchParams.get("$filter")).toBe(
      "userId eq 'synthetic-user' and createdDateTime ge 2026-10-06T17:00:00.000Z and createdDateTime le 2026-10-13T17:00:00.000Z",
    );
    expect(logUrl?.searchParams.get("$top")).toBe("50");
  });

  it("does not request optional feature packs without explicit granted scopes", async () => {
    const { client, fetchImplementation } = clientFor();
    const result = await investigateUser({ email, client, grantedScopes: ["User.Read.All"], now });
    expect(result.licenses).toMatchObject({ state: "missingConsent", data: [] });
    expect(result.signIns).toMatchObject({ state: "missingConsent", data: [], complete: false });
    expect(result.profile.state).toBe("available");
    expect(fetchImplementation).toHaveBeenCalledTimes(1);
  });

  it("keeps a denied section independent and does not diagnose a generic 403 as a specific role/license", async () => {
    const { client } = clientFor({
      license: new Response("sensitive provider error", { status: 403 }),
    });
    const result = await investigateUser({ email, client, grantedScopes: scopes, now });
    expect(result.licenses).toMatchObject({ state: "unknown", data: [] });
    expect(result.signIns.state).toBe("available");
    expect(JSON.stringify(result)).not.toContain("sensitive provider error");
  });

  it("maps profile not-found/access-denied to friendly typed errors", async () => {
    const missing = clientFor({ user: new Response(null, { status: 404 }) }).client;
    await expect(
      investigateUser({ email, client: missing, grantedScopes: scopes, now }),
    ).rejects.toMatchObject({ code: "user_not_found", status: 404 });
    const denied = clientFor({ user: new Response(null, { status: 403 }) }).client;
    await expect(
      investigateUser({ email, client: denied, grantedScopes: scopes, now }),
    ).rejects.toBeInstanceOf(UserInvestigationError);
    await expect(
      investigateUser({ email, client: clientFor().client, grantedScopes: [], now }),
    ).rejects.toMatchObject({ code: "missing_consent", status: 403 });
  });

  it("rejects records for another user or outside the requested window", async () => {
    for (const entry of [
      { ...signIn, userId: "synthetic-other-user" },
      { ...signIn, createdDateTime: "2025-10-13T16:00:00.000Z" },
    ]) {
      const { client } = clientFor({ signIns: Response.json({ value: [entry] }) });
      const result = await investigateUser({ email, client, grantedScopes: scopes, now });
      expect(result.signIns).toMatchObject({
        state: "providerUnavailable",
        data: [],
        complete: false,
      });
      expect(result.profile.state).toBe("available");
    }
  });

  it("rejects a lookup response for a different principal name", async () => {
    const { client } = clientFor({
      user: Response.json({ ...profile, userPrincipalName: "other@example.invalid" }),
    });
    await expect(
      investigateUser({ email, client, grantedScopes: scopes, now }),
    ).rejects.toMatchObject({ code: "provider_unavailable", status: 503 });
  });

  it("marks capped sign-in results incomplete instead of claiming full coverage", async () => {
    const { client } = clientFor({
      signIns: Response.json({
        value: Array.from({ length: 101 }, (_, index) => ({ ...signIn, id: `synthetic-${index}` })),
      }),
    });
    const result = await investigateUser({ email, client, grantedScopes: scopes, now });
    expect(result.signIns.data).toHaveLength(100);
    expect(result.signIns.complete).toBe(false);
    expect(result.signIns.message).toContain("100");
  });

  it("contains malformed/throttled optional responses without hiding the user profile", async () => {
    const { client } = clientFor({
      license: Response.json({ value: [{ skuId: 7 }] }),
      signIns: new Response(null, { status: 429, headers: { "Retry-After": "120" } }),
    });
    const result = await investigateUser({ email, client, grantedScopes: scopes, now });
    expect(result.profile.state).toBe("available");
    expect(result.licenses.state).toBe("providerUnavailable");
    expect(result.signIns.state).toBe("providerUnavailable");
  });

  it("encodes principal names as data, including a leading dollar sign", async () => {
    const special = "$alex@example.invalid";
    const { client, fetchImplementation } = clientFor({
      user: Response.json({ ...profile, userPrincipalName: special, mail: special }),
    });
    await investigateUser({ email: special, client, grantedScopes: ["User.Read.All"], now });
    expect(new URL(String(fetchImplementation.mock.calls[0]?.[0])).pathname).toBe(
      "/v1.0/users('%24alex%40example.invalid')",
    );
  });

  it("provides five synthetic sign-ins and rejects invalid input before any provider call", async () => {
    const demo = createDemoUserInvestigation({ email, now });
    expect(demo.mode).toBe("demo");
    expect(demo.signIns.data).toHaveLength(5);
    expect(demo.signIns.data.some((entry) => entry.status.errorCode === 500121)).toBe(true);
    const { client, fetchImplementation } = clientFor();
    await expect(investigateUser({ email: "not-a-user", client, now })).rejects.toMatchObject({
      code: "invalid_input",
    });
    expect(fetchImplementation).not.toHaveBeenCalled();
  });
});

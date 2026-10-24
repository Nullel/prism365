import { afterAll, afterEach, describe, expect, it, vi } from "vitest";
import { createApp, signInFailureCategory } from "../src/app.ts";
import type { AppConfig } from "../src/config.ts";
import { type IdentityService, LoginRejectedError } from "../src/identity.ts";

const tenantId = "11111111-1111-4111-8111-111111111111";
const liveConfig: AppConfig = {
  mode: "live",
  origin: "http://localhost:5173",
  production: false,
  identity: null,
  missingConfiguration: [],
  enabledFeatures: ["directory.users.read"],
};

function rejectingIdentity(failure: Error): IdentityService {
  return {
    tenantId,
    session: async () => null,
    startLogin: async () => ({ url: "https://login.microsoftonline.com/", state: "s".repeat(43) }),
    finishLogin: async () => {
      throw failure;
    },
    logout: async () => {},
    accessToken: async () => {
      throw new Error("unused");
    },
  } as unknown as IdentityService;
}

const app = createApp({
  config: liveConfig,
  identity: rejectingIdentity(
    new LoginRejectedError(
      "operator_not_allowed",
      "This Microsoft account is not an authorized Prism365 administrator.",
    ),
  ),
});

afterEach(() => vi.restoreAllMocks());
afterAll(async () => {
  await app.close();
});

describe("sign-in callback failures", () => {
  it("redirects a rejected operator with a generic marker and logs only a category", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    const response = await app.inject({
      method: "GET",
      url: `/v1/auth/callback?code=synthetic-code&state=${"a".repeat(40)}`,
    });

    expect(response.statusCode).toBe(302);
    expect(response.headers.location).toBe("http://localhost:5173/?auth=failed");
    expect(warn).toHaveBeenCalledWith("Prism365 sign-in rejected: operator_not_allowed");
    const logged = warn.mock.calls.flat().join(" ");
    expect(logged).not.toContain("synthetic-code");
    expect(logged).not.toContain("authorized Prism365 administrator");
  });

  it("rejects malformed callback parameters before touching the identity provider", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    const response = await app.inject({ method: "GET", url: "/v1/auth/callback?code=x" });

    expect(response.statusCode).toBe(302);
    expect(response.headers.location).toBe("http://localhost:5173/?auth=failed");
    expect(warn).toHaveBeenCalledWith("Prism365 sign-in rejected: invalid_callback_parameters");
  });

  it("maps identity failures to fixed categories without echoing the message", () => {
    expect(signInFailureCategory(new LoginRejectedError("transaction_expired", "x"))).toBe(
      "transaction_expired",
    );
    expect(signInFailureCategory(new LoginRejectedError("browser_binding_mismatch", "x"))).toBe(
      "browser_binding_mismatch",
    );
    expect(signInFailureCategory(new LoginRejectedError("token_claims_mismatch", "x"))).toBe(
      "token_claims_mismatch",
    );
    expect(signInFailureCategory(new Error("AADSTS70000 something secret"))).toBe(
      "token_exchange_failed",
    );
    expect(signInFailureCategory("not an error")).toBe("token_exchange_failed");
  });
});

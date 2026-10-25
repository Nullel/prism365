import { type PrismSession, sessionSchema } from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { signInFailedNotice } from "../src/App.tsx";
import { TenantConnection } from "../src/TenantConnection.tsx";

function session(overrides: Partial<PrismSession> = {}): PrismSession {
  return sessionSchema.parse({
    mode: "live",
    configured: true,
    authenticated: false,
    displayName: null,
    csrfToken: null,
    enabledFeatures: ["directory.users.read"],
    missingConfiguration: [],
    ...overrides,
  });
}

function connectionMarkup(value: PrismSession | null) {
  return renderToStaticMarkup(
    createElement(TenantConnection, {
      session: value,
      error: "",
      onRetry: () => {},
      onLogout: () => {},
      loggingOut: false,
    }),
  );
}

describe("connection interface", () => {
  it("shows a sign-in failure sentence on the live card without exposing a reason code", () => {
    const html = renderToStaticMarkup(
      createElement(TenantConnection, {
        session: session(),
        error: signInFailedNotice,
        onRetry: () => {},
        onLogout: () => {},
        loggingOut: false,
      }),
    );
    expect(html).toContain('role="alert"');
    expect(html).toContain("Microsoft sign-in did not complete");
    expect(html).toContain("operator list");
    expect(html).not.toContain("AADSTS");
    expect(html).toContain("Sign in with Microsoft");
  });

  it("waits for the mode without displaying a synthetic tenant", () => {
    const html = connectionMarkup(null);
    expect(html).toContain("Checking your workspace");
    expect(html).not.toContain("Northstar");
    expect(html).not.toContain("Synthetic");
    expect(html).not.toContain("/auth/login");
  });

  it("labels the demo and offers no Microsoft consent action", () => {
    const html = connectionMarkup(session({ mode: "demo", configured: false }));
    expect(html).toContain("synthetic data");
    expect(html).toContain("fictional");
    expect(html).not.toContain("/auth/login");
  });

  it("shows missing setup before presenting a sign-in action", () => {
    const html = connectionMarkup(
      session({ configured: false, missingConfiguration: ["Tenant registration"] }),
    );
    expect(html).toContain("Tenant registration");
    expect(html).not.toContain("/auth/login");
    expect(html).not.toContain("synthetic data");
  });

  it("does not offer profile sign-in when the investigation pack is disabled", () => {
    const html = connectionMarkup(session({ enabledFeatures: [] }));
    expect(html).not.toContain("/auth/login");
    expect(html).not.toContain("Sign in with Microsoft");
    expect(html).toContain("Not enabled in this installation");
  });

  it("offers access review only for enabled features in a live session", () => {
    const html = connectionMarkup(
      session({ authenticated: true, displayName: "Demo operator", csrfToken: "synthetic-csrf" }),
    );
    expect(html).toContain("feature=directory.users.read");
    expect(html).not.toContain("feature=audit.signins.read");
    expect(html).not.toContain("feature=user.licenses.read");
    expect(html).not.toContain("fictional");
    expect(html).toContain("Sign out");
  });
});

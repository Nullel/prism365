import { type PrismSession, sessionSchema } from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { SigninAudit } from "../src/SigninAudit.tsx";

function session(overrides: Partial<PrismSession> = {}): PrismSession {
  return sessionSchema.parse({
    mode: "live",
    configured: true,
    authenticated: true,
    displayName: "Synthetic admin",
    csrfToken: "synthetic-csrf-token-value-000000000000",
    enabledFeatures: ["directory.users.read", "audit.signins.read"],
    missingConfiguration: [],
    ...overrides,
  });
}

const render = (value: PrismSession) =>
  renderToStaticMarkup(createElement(SigninAudit, { session: value }));

describe("sign-in audit view", () => {
  it("shows no invented sign-in data in demo mode", () => {
    const html = render(session({ mode: "demo", authenticated: false, configured: false }));
    expect(html).toContain("does not invent security evidence");
    expect(html).not.toContain("Start collecting");
    expect(html).not.toMatch(/\d+ sign-ins stored/);
  });

  it("asks for sign-in before showing evidence", () => {
    const html = render(session({ authenticated: false, csrfToken: null, displayName: null }));
    expect(html).toContain("Sign in to see sign-in evidence");
    expect(html).not.toContain("Start collecting");
  });

  it("offers collection controls and an honest empty state when signed in", () => {
    const html = render(session());
    expect(html).toContain("Not collecting");
    expect(html).toContain("Start collecting");
    expect(html).toContain("Nothing open");
    expect(html).toContain('aria-pressed="true"');
    expect(html).toContain("not a measure of");
  });
});

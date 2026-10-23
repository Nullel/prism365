import { type PrismSession, sessionSchema } from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { hibpLink, People } from "../src/People.tsx";
import { runLine } from "../src/SigninAudit.tsx";

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

const render = (value: PrismSession, selected: string | null = null) =>
  renderToStaticMarkup(createElement(People, { session: value, selected, onSelect: () => {} }));

const run = {
  credentialKind: "delegated" as const,
  startedAt: "2026-10-22T19:44:01.000Z",
  finishedAt: "2026-10-22T19:46:40.000Z",
  fetched: 1354,
  inserted: 1300,
  findings: 0,
  skipped: 0,
  errorCategory: null,
};

describe("people view", () => {
  it("shows no invented people in demo mode", () => {
    const html = render(session({ mode: "demo", authenticated: false, configured: false }));
    expect(html).toContain("does not invent security evidence");
    expect(html).not.toContain("Find a person");
  });

  it("asks for sign-in first", () => {
    expect(render(session({ authenticated: false, csrfToken: null }))).toContain(
      "Sign in to see people",
    );
  });

  it("offers search when signed in, and a profile page for a selected person", () => {
    expect(render(session())).toContain("Find a person");
    const profile = render(session(), "44444444-4444-4444-8444-444444444444");
    expect(profile).toContain('aria-label="Breadcrumb"');
    expect(profile).toContain("Reading Microsoft");
    expect(profile).toContain("Every sign-in");
  });
});

describe("collection run line", () => {
  it("says a run is in progress instead of showing zero counts", () => {
    const line = runLine({ ...run, status: "running", finishedAt: null, fetched: 0, inserted: 0 });
    expect(line).toContain("Reading sign-ins now");
    expect(line).not.toContain("0 records");
    const stopped = runLine({ ...run, status: "running", finishedAt: null }, false);
    expect(stopped).toContain("cut short");
  });

  it("separates unreadable records from a busy stretch and a restart", () => {
    expect(runLine({ ...run, status: "partial", skipped: 2 })).toContain("2 unreadable");
    expect(runLine({ ...run, status: "partial" })).toContain("busy stretch");
    expect(runLine({ ...run, status: "abandoned" })).toContain("restart");
    expect(runLine({ ...run, status: "succeeded" })).toContain("1,354 records read");
  });
});

describe("Have I Been Pwned link", () => {
  it("opens HIBP's search for exactly the person's address", () => {
    expect(hibpLink("alex.rivera@example.invalid")).toBe(
      "https://haveibeenpwned.com/account/alex.rivera@example.invalid",
    );
    // Characters that could change the path or host are encoded.
    expect(hibpLink(" a/b?c#d@example.invalid ")).toBe(
      "https://haveibeenpwned.com/account/a%2Fb%3Fc%23d@example.invalid",
    );
  });
});

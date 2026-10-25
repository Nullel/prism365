import { type PersonListItem, type PrismSession, sessionSchema } from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { hibpLink, People, PersonRow } from "../src/People.tsx";
import { runLine } from "../src/SigninAudit.tsx";
import { at, expectSkeleton, relativeCell } from "./markup.ts";

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

// The run line is text with the time as a <time> element, so it is read as markup.
const line = (...args: Parameters<typeof runLine>) =>
  renderToStaticMarkup(createElement("p", null, runLine(...args)));

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

describe("people list row", () => {
  const item: PersonListItem = {
    userId: "44444444-4444-4444-8444-444444444444",
    userPrincipalName: "avery@example.invalid",
    displayName: "Avery Stone",
    lastSignIn: "2026-10-22T19:44:01.000Z",
    lastSuccess: "2026-10-22T19:44:01.000Z",
    signIns: 12,
    failures: 1,
    countries: 2,
    openFindings: 0,
  };
  const row = (value: PersonListItem) =>
    renderToStaticMarkup(
      createElement(
        "table",
        null,
        createElement("tbody", null, createElement(PersonRow, { item: value, onSelect: () => {} })),
      ),
    );

  it("keeps the address out of the button that opens the person, and lets it be copied", () => {
    const html = row(item);
    const opener = html.match(/<button class="people-name"[^>]*>(.*?)<\/button>/)?.[1] ?? "";
    expect(opener).toContain("Avery Stone");
    expect(opener).not.toContain("avery@example.invalid");
    expect(html).toMatch(/<span class="copy-line__text">avery@example\.invalid<\/span>/);
    expect(html).toContain('aria-label="Copy avery@example.invalid"');
  });

  it("has no copy button for an account without an address", () => {
    const html = row({ ...item, userPrincipalName: null });
    expect(html).toContain("Avery Stone");
    expect(html).not.toContain("Copy ");
  });
});

describe("collection run line", () => {
  it("says a run is in progress instead of showing zero counts", () => {
    const running = line({ ...run, status: "running", finishedAt: null, fetched: 0, inserted: 0 });
    expect(running).toContain("Reading sign-ins now");
    expect(running).not.toContain("0 records");
    const stopped = line({ ...run, status: "running", finishedAt: null }, false);
    expect(stopped).toContain("cut short");
  });

  it("separates unreadable records from a busy stretch and a restart", () => {
    expect(line({ ...run, status: "partial", skipped: 2 })).toContain("2 unreadable");
    expect(line({ ...run, status: "partial" })).toContain("busy stretch");
    expect(line({ ...run, status: "abandoned" })).toContain("restart");
    expect(line({ ...run, status: "succeeded" })).toContain("1,354 records read");
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

describe("the page speaks the shared visual language", () => {
  const NOW = "2026-10-22T20:00:00.000Z";
  const row = (lastSignIn: string | null) =>
    at(NOW, () =>
      renderToStaticMarkup(
        createElement(
          "table",
          null,
          createElement(
            "tbody",
            null,
            createElement(PersonRow, {
              item: {
                userId: "44444444-4444-4444-8444-444444444444",
                userPrincipalName: "avery@example.invalid",
                displayName: "Avery Stone",
                lastSignIn,
                lastSuccess: lastSignIn,
                signIns: 1234,
                failures: 1,
                countries: 2,
                openFindings: 0,
              },
              onSelect: () => {},
            }),
          ),
        ),
      ),
    );

  it("shows the last sign-in as words, with the exact moment on hover", () => {
    expect(row("2026-10-22T19:44:01.000Z")).toMatch(
      relativeCell("2026-10-22T19:44:01.000Z", "15 min ago"),
    );
  });

  it("says never, in muted text, for an account with no sign-in", () => {
    const html = row(null);
    expect(html).toContain('<span class="person-muted">never</span>');
    expect(html).not.toContain("<time");
  });

  it("draws still skeleton rows while the list loads", () => {
    expectSkeleton(render(session()), "Loading people");
  });
});

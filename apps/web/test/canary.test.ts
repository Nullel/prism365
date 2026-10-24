import { type CanaryAccount, type PrismSession, sessionSchema } from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { CanaryPanel, CanaryPanelView, CanaryToggle, CanaryToggleView } from "../src/Canary.tsx";
import { People } from "../src/People.tsx";
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

const noop = () => {};
// What a reader sees: tags and React's text-node separators removed.
const text = (html: string) =>
  html
    .replace(/<!--.*?-->/g, "")
    .replace(/<[^>]+>/g, " ")
    .replace(/&#x27;/g, "'")
    .replace(/\s+/g, " ");

const toggle = (
  overrides: Partial<Parameters<typeof CanaryToggleView>[0]> = {},
): ReturnType<typeof renderToStaticMarkup> =>
  renderToStaticMarkup(
    createElement(CanaryToggleView, {
      name: "Avery Stone",
      active: false,
      confirming: false,
      busy: false,
      error: "",
      onChange: noop,
      onConfirm: noop,
      onCancel: noop,
      ...overrides,
    }),
  );

describe("canary toggle", () => {
  it("is a checkbox with a one-line explanation, off by default", () => {
    const html = toggle();
    expect(html).toContain("Canary account");
    expect(text(html)).toContain(
      "A decoy account nobody should ever sign in to. Any sign-in attempt on it, successful or not, opens a critical finding.",
    );
    expect(html).toMatch(/<input[^>]*type="checkbox"/);
    expect(html).not.toContain("checked");
    expect(html).not.toContain("Mark as canary");
  });

  it("is on when the account is a canary", () => {
    expect(toggle({ active: true })).toContain("checked");
  });

  it("asks before turning on, naming the account, and stays off until confirmed", () => {
    const html = toggle({ confirming: true });
    expect(text(html)).toContain("Mark Avery Stone as a canary account?");
    expect(html).toContain("Mark as canary");
    expect(html).toContain("Cancel");
    expect(html).not.toContain("checked");
  });

  it("does not ask before turning off", () => {
    expect(toggle({ active: true })).not.toContain("Cancel");
  });

  it("disables the controls while saving and shows a failure", () => {
    const busy = toggle({ busy: true, confirming: true });
    expect(busy.match(/disabled/g)?.length).toBeGreaterThanOrEqual(3);
    const failed = toggle({ error: "That account is not in the directory." });
    expect(failed).toContain('role="alert"');
    expect(failed).toContain("That account is not in the directory.");
  });

  it("reads like a colleague wrote it", () => {
    const words = text(toggle({ confirming: true, error: "Nothing was changed." }));
    expect(words).not.toMatch(/[—–·!]/);
  });
});

const account = (userId: string, principal: string): CanaryAccount => ({
  userId,
  principal,
  addedAt: "2026-10-22T12:00:00.000Z",
  addedBy: "22222222-2222-4222-8222-222222222222",
});

const panel = (
  overrides: Partial<Parameters<typeof CanaryPanelView>[0]> = {},
): ReturnType<typeof renderToStaticMarkup> =>
  renderToStaticMarkup(
    createElement(CanaryPanelView, {
      accounts: [],
      busy: false,
      error: "",
      onRemove: noop,
      ...overrides,
    }),
  );

describe("canary panel", () => {
  it("says what to do when there are none", () => {
    const html = panel();
    expect(html).toContain("Canary accounts");
    expect(text(html)).toContain(
      "No canary accounts yet. Turn on Canary account on a person's page",
    );
    expect(html).not.toContain("<ul");
  });

  it("lists each account with a button that removes it", () => {
    const html = panel({
      accounts: [
        account("a-1", "decoy.one@example.invalid"),
        account("a-2", "decoy.two@example.invalid"),
      ],
    });
    expect(text(html)).toContain("decoy.one@example.invalid");
    expect(text(html)).toContain("decoy.two@example.invalid");
    expect(html).toContain('aria-label="Stop treating decoy.one@example.invalid as a canary"');
    expect(html.match(/Remove/g)).toHaveLength(2);
    expect(panel({ accounts: [account("a-1", "x@example.invalid")], busy: true })).toContain(
      "disabled",
    );
  });

  it("opens a person when it can, and shows a failure", () => {
    const html = panel({
      accounts: [account("a-1", "decoy.one@example.invalid")],
      onOpenPerson: noop,
      error: "The list could not be saved.",
    });
    expect(html).toMatch(/<button[^>]*finding__subject--link[^>]*>decoy\.one@example\.invalid</);
    expect(html).toContain("The list could not be saved.");
  });

  it("reads like a colleague wrote it", () => {
    const words = text(panel({ accounts: [account("a-1", "x@example.invalid")] }));
    expect(words).not.toMatch(/[—–·!]/);
  });
});

describe("canary containers", () => {
  it("show the controls, disabled, while the list loads", () => {
    const html = renderToStaticMarkup(
      createElement(CanaryToggle, { session: session(), userId: "u-1", name: "Avery Stone" }),
    );
    expect(html).toContain("Canary account");
    expect(html).toContain("disabled");
    expect(renderToStaticMarkup(createElement(CanaryPanel, { session: session() }))).toContain(
      "Canary accounts",
    );
  });

  it("sit on the person page and the sign-in audit page", () => {
    const profile = renderToStaticMarkup(
      createElement(People, {
        session: session(),
        selected: "44444444-4444-4444-8444-444444444444",
        onSelect: noop,
      }),
    );
    expect(profile).toContain("Canary account");
    expect(renderToStaticMarkup(createElement(SigninAudit, { session: session() }))).toContain(
      "Canary accounts",
    );
  });
});

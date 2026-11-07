import { type CanaryAccount, type PrismSession, sessionSchema } from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { CanaryPanel, CanaryPanelView, CanaryStepView, canaryExplanation } from "../src/Canary.tsx";
import { People } from "../src/People.tsx";
import { SigninAudit } from "../src/SigninAudit.tsx";
import { at, findAll, hostTree, textOf } from "./markup.ts";

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
const text = (html: string) => {
  // React separates adjacent text with exactly this comment. Tags are removed until none is
  // left, so removing one can never leave another behind.
  let current = html.replaceAll("<!-- -->", "");
  let previous: string;
  do {
    previous = current;
    current = current.replace(/<[^>]*>/g, " ");
  } while (current !== previous);
  return current.replace(/&#x27;/g, "'").replace(/\s+/g, " ");
};

const step = (
  overrides: Partial<Parameters<typeof CanaryStepView>[0]> = {},
): ReturnType<typeof renderToStaticMarkup> =>
  renderToStaticMarkup(
    createElement(CanaryStepView, {
      name: "Avery Stone",
      confirming: false,
      busy: false,
      error: "",
      onConfirm: noop,
      onCancel: noop,
      ...overrides,
    }),
  );

describe("canary explanation", () => {
  it("is one plain sentence pair that says what a canary does", () => {
    expect(canaryExplanation).toBe(
      "A decoy account nobody should ever sign in to. Any sign-in attempt on it, successful or not, opens a critical finding.",
    );
  });
});

describe("canary step after the menu item", () => {
  it("draws nothing until there is a question or a failure to show", () => {
    expect(step()).toBe("");
  });

  it("asks before marking, naming the account", () => {
    const html = step({ confirming: true });
    expect(text(html)).toContain("Mark Avery Stone as a canary account?");
    expect(html).toContain("Mark as canary");
    expect(html).toContain("Cancel");
    expect(html).not.toMatch(/<input/);
  });

  it("wires the two buttons to confirm and cancel", () => {
    const calls: string[] = [];
    const tree = hostTree(
      createElement(CanaryStepView, {
        name: "Avery Stone",
        confirming: true,
        busy: false,
        error: "",
        onConfirm: () => calls.push("confirm"),
        onCancel: () => calls.push("cancel"),
      }),
    );
    const buttons = findAll(tree, (host) => host.tag === "button");
    expect(buttons.map((button) => textOf(button))).toEqual(["Mark as canary", "Cancel"]);
    for (const button of buttons) {
      const onClick = button.props.onClick;
      if (typeof onClick !== "function") throw new Error("The button has no click handler.");
      onClick();
    }
    expect(calls).toEqual(["confirm", "cancel"]);
  });

  it("disables the buttons while saving and shows a failure", () => {
    const busy = step({ busy: true, confirming: true });
    expect(busy.match(/disabled/g)?.length).toBeGreaterThanOrEqual(2);
    const failed = step({ error: "That account is not in the directory." });
    expect(failed).toContain('role="alert"');
    expect(failed).toContain("That account is not in the directory.");
    expect(failed).not.toContain("Mark as canary");
  });

  it("reads like a colleague wrote it", () => {
    const words = text(step({ confirming: true, error: "Nothing was changed." }));
    expect(words).not.toMatch(/[—–·!]/);
    expect(canaryExplanation).not.toMatch(/[—–·!]/);
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
      "No canary accounts yet. Open a person's page and choose Mark as canary in its Actions menu to add one.",
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

  it("says when each account was added in words, with the exact moment on hover", () => {
    const html = at("2026-10-22T15:00:00.000Z", () =>
      panel({ accounts: [account("a-1", "decoy.one@example.invalid")] }),
    );
    expect(html).toMatch(
      /Added <time dateTime="2026-10-22T12:00:00.000Z" title="[^"]+">3 h ago<\/time>/,
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
  it("show the panel's heading while the list loads", () => {
    expect(renderToStaticMarkup(createElement(CanaryPanel, { session: session() }))).toContain(
      "Canary accounts",
    );
  });

  it("sit in the person page's Actions menu and on the sign-in audit page", () => {
    const profile = renderToStaticMarkup(
      createElement(People, {
        session: session(),
        selected: "44444444-4444-4444-8444-444444444444",
        onSelect: noop,
      }),
    );
    // The item is there, off until the list has loaded, and its explanation is its second line.
    expect(profile).toMatch(/<button[^>]*disabled=""[^>]*><span[^>]*>Mark as canary<\/span>/);
    expect(text(profile)).toContain(canaryExplanation);
    expect(profile).not.toMatch(/type="checkbox"/);
    expect(renderToStaticMarkup(createElement(SigninAudit, { session: session() }))).toContain(
      "Canary accounts",
    );
  });
});

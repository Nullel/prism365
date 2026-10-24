import { createElement, type ReactElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { afterEach, describe, expect, it, vi } from "vitest";
import { AccountName } from "../src/common.tsx";
import { CopyButton, copyText } from "../src/CopyButton.tsx";

afterEach(() => {
  vi.unstubAllGlobals();
});

// A document that records what the textarea fallback does to it.
function fakeDocument(execCommand: (command: string) => boolean) {
  const area = {
    value: "",
    style: {} as Record<string, string>,
    attributes: {} as Record<string, string>,
    setAttribute(name: string, value: string) {
      this.attributes[name] = value;
    },
    select: vi.fn(),
    setSelectionRange: vi.fn(),
    remove: vi.fn(),
  };
  const document = {
    activeElement: null,
    body: { appendChild: vi.fn() },
    createElement: vi.fn(() => area),
    execCommand: vi.fn(execCommand),
  };
  return { area, document };
}

describe("copyText", () => {
  it("resolves true when the clipboard API accepts the text", async () => {
    const writeText = vi.fn().mockResolvedValue(undefined);
    vi.stubGlobal("navigator", { clipboard: { writeText } });
    const { document } = fakeDocument(() => true);
    vi.stubGlobal("document", document);
    await expect(copyText("avery@example.invalid")).resolves.toBe(true);
    expect(writeText).toHaveBeenCalledWith("avery@example.invalid");
    // The fallback is only a fallback.
    expect(document.execCommand).not.toHaveBeenCalled();
  });

  it("falls back to a hidden textarea and execCommand when the clipboard API rejects", async () => {
    vi.stubGlobal("navigator", {
      clipboard: { writeText: vi.fn().mockRejectedValue(new Error("denied")) },
    });
    const { area, document } = fakeDocument(() => true);
    vi.stubGlobal("document", document);
    await expect(copyText("avery@example.invalid")).resolves.toBe(true);
    expect(document.createElement).toHaveBeenCalledWith("textarea");
    expect(area.value).toBe("avery@example.invalid");
    expect(area.attributes.readonly).toBe("");
    expect(document.body.appendChild).toHaveBeenCalledWith(area);
    expect(document.execCommand).toHaveBeenCalledWith("copy");
    // The textarea does not stay in the page.
    expect(area.remove).toHaveBeenCalledTimes(1);
  });

  it("falls back the same way when the browser has no clipboard API", async () => {
    vi.stubGlobal("navigator", {});
    const { document } = fakeDocument(() => true);
    vi.stubGlobal("document", document);
    await expect(copyText("a@example.invalid\nb@example.invalid")).resolves.toBe(true);
    expect(document.execCommand).toHaveBeenCalledWith("copy");
  });

  it("resolves false, without throwing, when both ways fail", async () => {
    vi.stubGlobal("navigator", {
      clipboard: { writeText: vi.fn().mockRejectedValue(new Error("denied")) },
    });
    const refused = fakeDocument(() => false);
    vi.stubGlobal("document", refused.document);
    await expect(copyText("avery@example.invalid")).resolves.toBe(false);
    expect(refused.area.remove).toHaveBeenCalledTimes(1);

    const broken = fakeDocument(() => {
      throw new Error("not allowed");
    });
    vi.stubGlobal("document", broken.document);
    await expect(copyText("avery@example.invalid")).resolves.toBe(false);
    // Even a failed copy takes its textarea out again.
    expect(broken.area.remove).toHaveBeenCalledTimes(1);
  });

  it("resolves false when there is no clipboard and no page to fall back on", async () => {
    vi.stubGlobal("navigator", undefined);
    vi.stubGlobal("document", undefined);
    await expect(copyText("avery@example.invalid")).resolves.toBe(false);
  });
});

const render = (node: ReactElement) => renderToStaticMarkup(node);

describe("CopyButton", () => {
  it("is a labelled icon button with a polite status for screen readers", () => {
    const html = render(createElement(CopyButton, { text: "avery@example.invalid" }));
    expect(html).toContain('type="button"');
    expect(html).toContain('aria-label="Copy avery@example.invalid"');
    expect(html).toContain('title="Copy avery@example.invalid"');
    expect(html).toContain("<svg");
    expect(html).toMatch(/<span class="visually-hidden" aria-live="polite"><\/span>/);
    // Nothing is announced before anything is copied.
    expect(html).not.toContain("Copied");
    expect(html).not.toContain("Copy failed");
  });

  it("does not let the click reach the row or the link around it", () => {
    type Click = (event: { stopPropagation: () => void }) => void;
    const captured: { onClick?: Click | undefined } = {};
    // The copy never settles, so no state changes after this render is over.
    vi.stubGlobal("navigator", { clipboard: { writeText: () => new Promise(() => {}) } });
    function Probe() {
      const tree = CopyButton({ text: "avery@example.invalid" }) as ReactElement<{
        children: ReactElement<{ onClick: Click }>[];
      }>;
      captured.onClick = tree.props.children[0]?.props.onClick;
      return tree;
    }
    render(createElement(Probe));
    const stopPropagation = vi.fn();
    expect(captured.onClick).toBeTypeOf("function");
    captured.onClick?.({ stopPropagation });
    expect(stopPropagation).toHaveBeenCalledTimes(1);
  });
});

describe("AccountName", () => {
  it("keeps the address outside the name button, as text, with its own copy button", () => {
    const html = render(
      createElement(AccountName, {
        name: "Avery Stone",
        principal: "avery@example.invalid",
        onOpen: () => {},
      }),
    );
    const button = html.match(/<button class="people-name"[^>]*>(.*?)<\/button>/)?.[1] ?? "";
    expect(button).toContain("Avery Stone");
    expect(button).not.toContain("avery@example.invalid");
    expect(html).toMatch(/<span class="copy-line__text">avery@example\.invalid<\/span>/);
    expect(html).toContain('aria-label="Copy avery@example.invalid"');
    expect(html.match(/aria-label="Copy /g)).toHaveLength(1);
  });

  it("puts the copy button after the name when the name is the address", () => {
    const html = render(
      createElement(AccountName, {
        name: "avery@example.invalid",
        principal: "avery@example.invalid",
        onOpen: () => {},
      }),
    );
    expect(html).not.toContain("copy-line__text");
    expect(html).toMatch(/<\/button><button[^>]*aria-label="Copy avery@example.invalid"/);
  });

  it("puts the copy button after the name when there is no person page", () => {
    const html = render(
      createElement(AccountName, {
        name: "Avery Stone",
        principal: "avery@example.invalid",
        onOpen: null,
      }),
    );
    expect(html).not.toContain('class="people-name"');
    expect(html).toMatch(
      /Avery Stone<\/strong><button[^>]*aria-label="Copy avery@example.invalid"/,
    );
  });

  it("has no copy button without an address", () => {
    for (const principal of [null, "", "  "]) {
      const html = render(
        createElement(AccountName, { name: "Avery Stone", principal, onOpen: () => {} }),
      );
      expect(html).not.toContain("Copy");
      expect(html).toContain("Avery Stone");
    }
    expect(
      render(createElement(AccountName, { name: "Avery Stone", principal: null, onOpen: null })),
    ).not.toContain("<button");
  });
});

import { Fragment, isValidElement, type ReactNode } from "react";
import { expect, vi } from "vitest";

// Helpers for tests that look at what a component draws without a browser: the host elements it
// produces as a plain tree, and a check that no link or button holds another one.

export interface Host {
  tag: string;
  props: Record<string, unknown>;
  children: (Host | string)[];
}

// Runs function components (the ones in these pages use no hooks) until only host elements are
// left, so a test can find a button and call its handler.
export function hostTree(node: ReactNode): (Host | string)[] {
  if (node === null || node === undefined || typeof node === "boolean") return [];
  if (typeof node === "string" || typeof node === "number") return [String(node)];
  if (Array.isArray(node)) return node.flatMap((item) => hostTree(item));
  if (!isValidElement(node)) return [];
  const { type } = node;
  const props = node.props as Record<string, unknown>;
  if (type === Fragment) return hostTree(props.children as ReactNode);
  if (typeof type === "function")
    return hostTree((type as (value: Record<string, unknown>) => ReactNode)(props));
  return [
    { tag: String(type), props, children: hostTree(props.children as ReactNode) } satisfies Host,
  ];
}

export function textOf(node: Host | string): string {
  return typeof node === "string" ? node : node.children.map(textOf).join("");
}

export function findAll(nodes: readonly (Host | string)[], test: (host: Host) => boolean): Host[] {
  const found: Host[] = [];
  for (const node of nodes) {
    if (typeof node === "string") continue;
    if (test(node)) found.push(node);
    found.push(...findAll(node.children, test));
  }
  return found;
}

// A plain click, with nothing held down.
export const plainClick = () => ({
  button: 0,
  metaKey: false,
  ctrlKey: false,
  shiftKey: false,
  altKey: false,
  defaultPrevented: false,
  preventDefault() {},
});

// How many links or buttons sit inside another link or button in this markup. Zero is the goal:
// an interactive element inside another has two meanings for one click.
export function nestedInteractive(html: string): number {
  let depth = 0;
  let nested = 0;
  for (const match of html.matchAll(/<(\/?)(a|button)\b[^>]*>/g)) {
    if (match[1] === "/") depth = Math.max(0, depth - 1);
    else {
      if (depth > 0) nested += 1;
      depth += 1;
    }
  }
  return nested;
}

// Runs a render with the clock fixed, so a time shown as words ("5 min ago") reads the same on
// every day the test runs. Only Date is faked: nothing in a render waits on a timer.
export function at<T>(now: string, render: () => T): T {
  vi.useFakeTimers({ toFake: ["Date"] });
  vi.setSystemTime(new Date(now));
  try {
    return render();
  } finally {
    vi.useRealTimers();
  }
}

// A date shown as words: a time element that holds the exact moment and says `words`.
export function relativeCell(iso: string, words: string): RegExp {
  return new RegExp(`<time dateTime="${iso}" title="[^"]+">${words}</time>`);
}

// Severity as the shared dot and name, with a count when it has one.
export function severityDot(severity: string, label: string): string {
  return `<span class="sev sev--${severity}"><span class="sev__dot" aria-hidden="true"></span>${label}</span>`;
}

// What a loading page draws: still grey rows, a label for assistive technology, and never text
// that says Loading with an ellipsis.
export function expectSkeleton(html: string, label: string) {
  expect(html).toContain("skeleton__row");
  expect(html).toContain('aria-busy="true"');
  expect(html).toContain(`<span class="visually-hidden">${label}</span>`);
  expect(html).not.toContain("Loading…");
}

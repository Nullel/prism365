import type { EntityKind } from "@prism365/contracts";

// What the entity panel is showing, as plain data so every step is a test. The panel holds the
// things it has shown as a stack: opening another thing from inside it adds a step the Back button
// goes through, and anything that leaves the panel (closing it, opening a person's page) empties
// the stack. An empty stack is a closed panel.

export interface EntityTarget {
  kind: EntityKind;
  id: string;
}

export type PanelAction =
  // `fresh` starts over, as when something outside the panel (Quick search) opens a thing. Without
  // it, a panel that is already open keeps its steps and adds one.
  { type: "open"; target: EntityTarget; fresh?: boolean } | { type: "back" } | { type: "close" };

const same = (a: EntityTarget, b: EntityTarget) =>
  a.kind === b.kind && a.id.toLowerCase() === b.id.toLowerCase();

export function panelReducer(
  stack: readonly EntityTarget[],
  action: PanelAction,
): readonly EntityTarget[] {
  switch (action.type) {
    case "open": {
      const top = stack.at(-1);
      if (action.fresh || !top) return [action.target];
      // Opening what is already shown adds nothing to go back to.
      return same(top, action.target) ? stack : [...stack, action.target];
    }
    case "back":
      return stack.length > 1 ? stack.slice(0, -1) : stack;
    case "close":
      return [];
  }
}

// The part of a keyboard event the panel reads, so a test can pass a plain object.
export interface PanelKeyEvent {
  key: string;
  shiftKey: boolean;
  defaultPrevented: boolean;
  isComposing?: boolean;
  nativeEvent?: { isComposing?: boolean };
  preventDefault(): void;
}

interface Focusable {
  focus(): void;
}

// Escape closes the panel wherever focus is, and nothing else sees it. Tab at either end of the
// panel's controls wraps to the other end, and Tab with focus outside the panel (the page behind,
// or nothing after a click on text) brings focus into it, so focus never wanders off behind a
// dialog that covers the page.
export function handlePanelKey(
  event: PanelKeyEvent,
  controls: {
    first: Focusable | null;
    last: Focusable | null;
    active: unknown;
    // Whether focus is somewhere in the panel.
    inside: boolean;
  },
  close: () => void,
): void {
  if (event.defaultPrevented) return;
  if (event.key === "Escape") {
    // Escape that cancels a word in an input method is not a request to close anything.
    if ((event.nativeEvent?.isComposing ?? event.isComposing) === true) return;
    event.preventDefault();
    close();
    return;
  }
  if (event.key !== "Tab" || !controls.first || !controls.last) return;
  if (!controls.inside) {
    event.preventDefault();
    (event.shiftKey ? controls.last : controls.first).focus();
    return;
  }
  if (event.shiftKey && controls.active === controls.first) {
    event.preventDefault();
    controls.last.focus();
  } else if (!event.shiftKey && controls.active === controls.last) {
    event.preventDefault();
    controls.first.focus();
  }
}

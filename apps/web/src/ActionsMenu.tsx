import { type KeyboardEvent, type RefObject, useEffect, useId, useRef, useState } from "react";
import { ArrowUpRightIcon, ChevronIcon } from "./icons.tsx";

// A button that discloses a short list of actions: buttons that do something and links that go
// somewhere. It is a disclosure, not an ARIA menu: the button says whether the list is open and
// which list it controls, the arrow keys move between the items, Escape closes the list and gives
// the focus back to the button, and so does choosing an item or clicking anywhere else.

export type MenuItem =
  | {
      kind: "button";
      id: string;
      label: string;
      // A second line under the label: what the item does, or what it will act on.
      hint?: string;
      disabled?: boolean;
      // Stays open after the choice, for an item that shows its own result.
      keepOpen?: boolean;
      onSelect: () => void;
    }
  | {
      kind: "link";
      id: string;
      label: string;
      hint?: string;
      // Always opens in a new tab: these are other sites, and the page keeps its place.
      href: string;
    };

// Where the arrow keys go: down and up step through the items and wrap at the ends, Home and End
// jump to the first and last. `current` is -1 when the focus is not on an item.
export function menuKeyTarget(key: string, current: number, count: number): number | null {
  if (count <= 0) return null;
  switch (key) {
    case "ArrowDown":
      return current < 0 ? 0 : (current + 1) % count;
    case "ArrowUp":
      return current <= 0 ? count - 1 : current - 1;
    case "Home":
      return 0;
    case "End":
      return count - 1;
    default:
      return null;
  }
}

export interface MenuKeyContext {
  open: boolean;
  // The items that can take focus, in order, and which of them has it (-1 for none).
  items: readonly { focus: () => void }[];
  current: number;
  openMenu: (focus: "first" | "last") => void;
  closeMenu: () => void;
  focusButton: () => void;
}

// What a key does to the menu. Returns whether the key was used, so the caller can stop the
// browser from also scrolling the page with it.
export function handleMenuKey(key: string, context: MenuKeyContext): boolean {
  if (key === "Escape") {
    if (!context.open) return false;
    context.closeMenu();
    context.focusButton();
    return true;
  }
  if (key !== "ArrowDown" && key !== "ArrowUp" && key !== "Home" && key !== "End") return false;
  if (!context.open) {
    // Only the arrows open a closed list; Home and End keep their usual meaning on the page.
    if (key === "ArrowDown") context.openMenu("first");
    else if (key === "ArrowUp") context.openMenu("last");
    else return false;
    return true;
  }
  const target = menuKeyTarget(key, context.current, context.items.length);
  if (target === null) return false;
  context.items[target]?.focus();
  return true;
}

// Whether a click or the focus landed outside the menu, which closes it.
export function isOutside(
  container: { contains: (node: Node | null) => boolean } | null,
  target: EventTarget | null,
): boolean {
  return container !== null && !container.contains(target as Node | null);
}

// Choosing an item runs it, then closes the list and returns the focus to the button, unless the
// item asked to stay open.
export function activateItem(
  item: MenuItem,
  controls: { closeMenu: () => void; focusButton: () => void },
): void {
  if (item.kind === "button") {
    item.onSelect();
    if (item.keepOpen) return;
  }
  controls.closeMenu();
  controls.focusButton();
}

export function ActionsMenuView({
  label,
  open,
  buttonId,
  menuId,
  items,
  onToggle,
  onActivate,
  onKeyDown,
  wrapRef,
  buttonRef,
  listRef,
}: {
  label: string;
  open: boolean;
  buttonId: string;
  menuId: string;
  items: readonly MenuItem[];
  onToggle: () => void;
  onActivate: (item: MenuItem) => void;
  onKeyDown: (event: KeyboardEvent<HTMLElement>) => void;
  wrapRef?: RefObject<HTMLDivElement | null>;
  buttonRef?: RefObject<HTMLButtonElement | null>;
  listRef?: RefObject<HTMLUListElement | null>;
}) {
  return (
    // The key handler is on the wrapper so the arrows work from the button and from the items.
    // biome-ignore lint/a11y/noStaticElementInteractions: a wrapper that only listens for keys that reach its button and items
    <div className="actions-menu" ref={wrapRef} onKeyDown={onKeyDown}>
      <button
        id={buttonId}
        ref={buttonRef}
        className="button button--secondary actions-menu__button"
        type="button"
        aria-expanded={open}
        aria-controls={menuId}
        onClick={onToggle}
      >
        {label}
        <ChevronIcon className="actions-menu__chevron" />
      </button>
      <ul
        id={menuId}
        ref={listRef}
        className="actions-menu__list"
        aria-labelledby={buttonId}
        hidden={!open}
      >
        {items.map((item) => (
          <li key={item.id}>
            {item.kind === "link" ? (
              <a
                className="actions-menu__item"
                href={item.href}
                target="_blank"
                rel="noreferrer noopener"
                onClick={() => onActivate(item)}
              >
                <span className="actions-menu__label">
                  {item.label}
                  <ArrowUpRightIcon />
                </span>
                {item.hint ? <span className="actions-menu__hint">{item.hint}</span> : null}
              </a>
            ) : (
              <button
                className="actions-menu__item"
                type="button"
                disabled={item.disabled ?? false}
                onClick={() => onActivate(item)}
              >
                <span className="actions-menu__label">{item.label}</span>
                {item.hint ? <span className="actions-menu__hint">{item.hint}</span> : null}
              </button>
            )}
          </li>
        ))}
      </ul>
    </div>
  );
}

// The items that can take focus right now: links, and buttons that are not disabled.
function reachable(list: HTMLElement | null): HTMLElement[] {
  return Array.from(list?.querySelectorAll<HTMLElement>("a[href], button:not(:disabled)") ?? []);
}

export function ActionsMenu({
  items,
  label = "Actions",
}: {
  items: readonly MenuItem[];
  label?: string;
}) {
  const [open, setOpen] = useState(false);
  const uid = useId();
  const wrap = useRef<HTMLDivElement>(null);
  const button = useRef<HTMLButtonElement>(null);
  const list = useRef<HTMLUListElement>(null);
  // Set when an arrow key opens the list, so the focus can move once the items are on screen.
  const focusOnOpen = useRef<"first" | "last" | null>(null);

  useEffect(() => {
    if (!open) return;
    const wanted = focusOnOpen.current;
    focusOnOpen.current = null;
    if (wanted) {
      const found = reachable(list.current);
      (wanted === "first" ? found[0] : found[found.length - 1])?.focus();
    }
    const away = (event: Event) => {
      if (isOutside(wrap.current, event.target)) setOpen(false);
    };
    document.addEventListener("pointerdown", away);
    document.addEventListener("focusin", away);
    return () => {
      document.removeEventListener("pointerdown", away);
      document.removeEventListener("focusin", away);
    };
  }, [open]);

  const closeMenu = () => setOpen(false);
  const focusButton = () => button.current?.focus();
  return (
    <ActionsMenuView
      label={label}
      open={open}
      buttonId={`${uid}-button`}
      menuId={`${uid}-list`}
      items={items}
      wrapRef={wrap}
      buttonRef={button}
      listRef={list}
      onToggle={() => setOpen((current) => !current)}
      onActivate={(item) => activateItem(item, { closeMenu, focusButton })}
      onKeyDown={(event) => {
        const found = reachable(list.current);
        const used = handleMenuKey(event.key, {
          open,
          items: found,
          current: found.indexOf(document.activeElement as HTMLElement),
          openMenu: (focus) => {
            focusOnOpen.current = focus;
            setOpen(true);
          },
          closeMenu,
          focusButton,
        });
        if (used) event.preventDefault();
      }}
    />
  );
}

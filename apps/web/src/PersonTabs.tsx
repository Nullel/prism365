import type { KeyboardEvent, ReactNode } from "react";
import { number } from "./common.tsx";

// The tabs on a person's page. The selected tab is part of the address (#/people/<id>/signins),
// and a plain #/people/<id> opens the first one, so links made before the tabs existed still work.

export const personTabs = ["overview", "signins", "licenses", "changes"] as const;
export type PersonTab = (typeof personTabs)[number];

export const personTabLabels: Record<PersonTab, string> = {
  overview: "Overview",
  signins: "Sign-ins",
  licenses: "Licenses",
  changes: "Changes",
};

// The tab a route segment names, or null when it names none.
export function parsePersonTab(segment: string | undefined): PersonTab | null {
  return personTabs.find((tab) => tab === segment) ?? null;
}

export interface TabCount {
  value: number;
  // The number is the most that has been loaded so far, so it is shown as "20+".
  more?: boolean;
}

// The count each tab carries, where the page already has one. A null input means that part of the
// page has not loaded, and the tab shows no number rather than a zero it does not know.
export function personTabCounts(input: {
  openFindings: number | null;
  signIns: number | null;
  licenses: number | null;
  changes: TabCount | null;
}): Record<PersonTab, TabCount | null> {
  return {
    // Findings are the only thing the first tab counts, and "0" would read as an empty page. The
    // header already says "Nothing open".
    overview:
      input.openFindings !== null && input.openFindings > 0 ? { value: input.openFindings } : null,
    signins: input.signIns === null ? null : { value: input.signIns },
    licenses: input.licenses === null ? null : { value: input.licenses },
    changes: input.changes,
  };
}

export function countLabel(count: TabCount): string {
  return `${number.format(count.value)}${count.more ? "+" : ""}`;
}

export const tabId = (base: string, tab: PersonTab) => `${base}-tab-${tab}`;
export const panelId = (base: string, tab: PersonTab) => `${base}-panel-${tab}`;

// Where the arrow keys go on the tab list: left and right step through the tabs and wrap, Home and
// End jump to the ends. Null for any other key.
export function tabKeyTarget(key: string, selected: PersonTab): PersonTab | null {
  const index = personTabs.indexOf(selected);
  const last = personTabs.length - 1;
  switch (key) {
    case "ArrowRight":
      return personTabs[index >= last ? 0 : index + 1] ?? null;
    case "ArrowLeft":
      return personTabs[index <= 0 ? last : index - 1] ?? null;
    case "Home":
      return personTabs[0] ?? null;
    case "End":
      return personTabs[last] ?? null;
    default:
      return null;
  }
}

// A key on the tab list selects the tab it points to and moves the focus there, so the tab under
// the focus is always the one showing. Returns whether the key was used.
export function handleTabKey(
  key: string,
  selected: PersonTab,
  controls: { select: (tab: PersonTab) => void; focusTab: (tab: PersonTab) => void },
): boolean {
  const target = tabKeyTarget(key, selected);
  if (target === null) return false;
  controls.select(target);
  controls.focusTab(target);
  return true;
}

function focusById(id: string) {
  if (typeof document !== "undefined") document.getElementById(id)?.focus();
}

export function PersonTabList({
  idBase,
  selected,
  counts,
  onSelect,
}: {
  idBase: string;
  selected: PersonTab;
  counts: Record<PersonTab, TabCount | null>;
  onSelect: (tab: PersonTab) => void;
}) {
  return (
    // The strip scrolls sideways by itself on a narrow screen; nothing else on the page does.
    <div className="person-tabs-strip">
      <div
        className="person-tabs"
        role="tablist"
        aria-label="Sections of this person's page"
        onKeyDown={(event: KeyboardEvent<HTMLElement>) => {
          const used = handleTabKey(event.key, selected, {
            select: onSelect,
            focusTab: (tab) => focusById(tabId(idBase, tab)),
          });
          if (used) event.preventDefault();
        }}
      >
        {personTabs.map((tab) => {
          const count = counts[tab];
          return (
            <button
              key={tab}
              id={tabId(idBase, tab)}
              className="person-tab"
              type="button"
              role="tab"
              aria-selected={tab === selected}
              aria-controls={panelId(idBase, tab)}
              tabIndex={tab === selected ? 0 : -1}
              onClick={() => onSelect(tab)}
            >
              {personTabLabels[tab]}
              {count ? <span className="person-tab__count">{countLabel(count)}</span> : null}
            </button>
          );
        })}
      </div>
    </div>
  );
}

// Every panel stays on the page and the ones not showing are hidden, so a tab's requests are made
// once, when the page opens, and switching tabs never asks for anything again.
export function PersonTabPanel({
  idBase,
  tab,
  selected,
  children,
}: {
  idBase: string;
  tab: PersonTab;
  selected: PersonTab;
  children?: ReactNode;
}) {
  return (
    <div
      id={panelId(idBase, tab)}
      className="person-panel"
      role="tabpanel"
      aria-labelledby={tabId(idBase, tab)}
      hidden={tab !== selected}
      // biome-ignore lint/a11y/noNoninteractiveTabindex: the tabs pattern puts a panel in the tab order, so its content is reachable even when it holds nothing focusable
      tabIndex={0}
    >
      {children}
    </div>
  );
}

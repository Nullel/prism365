import { readFileSync } from "node:fs";
import {
  type AuditCatalog,
  auditCatalogSchema,
  type PersonList,
  type PrismSession,
  personListSchema,
  sessionSchema,
} from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { App, navGroups } from "../src/App.tsx";
import { createDemoState, type Fixtures, handleRequest, normalizeKey } from "../src/demo/replay.ts";
import {
  auditItems,
  auditsUrl,
  emailToLookUp,
  type FlatResult,
  flatten,
  focusKeeper,
  fold,
  handleDialogKey,
  isMacPlatform,
  isTextField,
  type KeyEventLike,
  MAX_RESULTS,
  noMatchText,
  openFromKeyboard,
  opensQuickSearch,
  optionId,
  pageItems,
  pageSynonyms,
  peopleUrl,
  personItems,
  QuickSearch,
  QuickSearchButton,
  QuickSearchView,
  type QuickSearchViewProps,
  resultStatus,
  type SearchItem,
  type SearchSources,
  scoreItem,
  searchResults,
  shortcutHint,
  sourceNotes,
} from "../src/QuickSearch.tsx";
import { findAll, type Host, hostTree, nestedInteractive, textOf } from "./markup.ts";

// Invented people, audits and addresses only.

const pages = pageItems(navGroups);

function person(name: string | null, principal: string | null, index: number) {
  return {
    userId: `00000000-0000-4000-8000-${String(index).padStart(12, "0")}`,
    userPrincipalName: principal,
    displayName: name,
    lastSignIn: "2026-10-20T10:00:00.000Z",
    lastSuccess: "2026-10-20T10:00:00.000Z",
    signIns: 4,
    failures: 0,
    countries: 1,
    openFindings: 0,
  };
}

function people(...entries: [string | null, string | null][]): PersonList {
  return personListSchema.parse({
    mode: "live",
    days: 30,
    items: entries.map(([name, principal], index) => person(name, principal, index + 1)),
    truncated: false,
  });
}

const directory = people(
  ["Marcus Rowe", "marcus.rowe@example.invalid"],
  ["Rowan Ellis", "rowan.ellis@example.invalid"],
  ["Dana Campbell", "dana.campbell@example.invalid"],
  ["Zoë Álvarez", "zoe.alvarez@example.invalid"],
  [null, "shared.mailbox@example.invalid"],
);

function auditEntry(
  id: string,
  title: string,
  category: string,
  lastRun: Record<string, unknown> | null,
) {
  return {
    definition: {
      id,
      version: 1,
      title,
      description: "Every account and how long it has gone without signing in.",
      category,
      subject: { one: "account", many: "accounts" },
      columns: [{ key: "person", label: "Account", kind: "person" }],
      views: [{ id: "all", label: "All accounts", tags: [] }],
      defaultView: "all",
      defaultSort: { column: "person", direction: "asc" },
      permissions: ["User.Read.All"],
      optionalPermissions: [],
      requirements: [],
    },
    availability: { state: "available", missing: [], grantFeature: null, enable: [] },
    lastRun,
    runningId: null,
  };
}

const run = {
  id: "11111111-1111-4111-8111-111111111111",
  auditId: "accounts.inactive",
  auditVersion: 1,
  status: "succeeded",
  complete: true,
  startedAt: "2026-10-20T09:59:00.000Z",
  finishedAt: "2026-10-20T09:59:03.000Z",
  startedBy: { id: "22222222-2222-4222-8222-222222222222", name: "Synthetic admin" },
  error: null,
  summary: [{ label: "Inactive 30+ days", value: 2, tone: "warning", view: "all" }],
  notes: [],
  rowCount: 3,
};

const catalog: AuditCatalog = auditCatalogSchema.parse({
  mode: "live",
  audits: [
    auditEntry("accounts.inactive", "Inactive accounts", "Accounts", run),
    auditEntry("mfa.registration", "MFA registration", "Security", null),
  ],
});

const sources: SearchSources = {
  pages,
  audits: auditItems(catalog),
  people: personItems(directory),
};

const ids = (groups: ReturnType<typeof searchResults>) =>
  flatten(groups).map((entry) => entry.item.id);
const labels = (groups: ReturnType<typeof searchResults>) =>
  flatten(groups).map((entry) => entry.item.label);
const peopleOnly = (list: PersonList): SearchSources => ({
  pages: [],
  audits: [],
  people: personItems(list),
});

function item(label: string, keywords: string[] = [], id = label): SearchItem {
  return { id, label, detail: null, keywords };
}

describe("matching", () => {
  it("ignores case and accents in both the text and what is typed", () => {
    expect(fold("Zoë ÁLVAREZ")).toBe("zoe alvarez");
    for (const typed of ["zoe", "ZOË", "Zoe Alvarez", "álvarez", "ALVAREZ"])
      expect(scoreItem(item("Zoë Álvarez"), typed), typed).toBeGreaterThan(0);
    expect(ids(searchResults(sources, "zoe"))).toEqual([directory.items[3]?.userId]);
  });

  it("needs every word typed, in any order", () => {
    expect(scoreItem(item("Rowan Ellis"), "ellis rowan")).toBeGreaterThan(0);
    expect(scoreItem(item("Rowan Ellis"), "rowan nobody")).toBe(0);
    expect(scoreItem(item("Rowan Ellis"), "   ")).toBe(0);
  });

  it("prefers the start of a word to the middle of one, even when the middle comes first", () => {
    // "Dana Campbell" holds "ell" inside a word and is listed before "Rowan Ellis".
    const list = people(
      ["Dana Campbell", "dana.campbell@example.invalid"],
      ["Rowan Ellis", "rowan.ellis@example.invalid"],
    );
    expect(labels(searchResults(peopleOnly(list), "ell"))).toEqual([
      "Rowan Ellis",
      "Dana Campbell",
    ]);
  });

  it("prefers the start of the whole name to the start of a later word", () => {
    expect(labels(searchResults(peopleOnly(directory), "ro"))).toEqual([
      "Rowan Ellis",
      "Marcus Rowe",
    ]);
  });

  it("weighs the label above a keyword, and a typed full label above both", () => {
    const things = [item("Audit trail", ["people"]), item("People", []), item("Peoples choice")];
    const found = searchResults({ pages: things, audits: [], people: [] }, "people");
    expect(labels(found)).toEqual(["People", "Peoples choice", "Audit trail"]);
  });

  it("keeps the order it was given when matches are equal", () => {
    const first = people(
      ["Dana Ellis", "dana.ellis@example.invalid"],
      ["Rowan Ellis", "rowan.ellis@example.invalid"],
    );
    expect(labels(searchResults(peopleOnly(first), "ellis"))).toEqual([
      "Dana Ellis",
      "Rowan Ellis",
    ]);
    const reversed = people(
      ["Rowan Ellis", "rowan.ellis@example.invalid"],
      ["Dana Ellis", "dana.ellis@example.invalid"],
    );
    expect(labels(searchResults(peopleOnly(reversed), "ellis"))).toEqual([
      "Rowan Ellis",
      "Dana Ellis",
    ]);
  });

  it("matches a person on their principal name when the display name does not match", () => {
    expect(labels(searchResults(sources, "shared.mail"))).toEqual([
      "shared.mailbox@example.invalid",
    ]);
    expect(labels(searchResults(sources, "rowan.ellis@example"))).toEqual(["Rowan Ellis"]);
  });

  it("matches an audit on its name and its category", () => {
    expect(ids(searchResults(sources, "inactive"))).toContain("accounts.inactive");
    expect(ids(searchResults(sources, "security"))).toEqual(["mfa.registration"]);
  });

  it("shows at most eight results at once, and keeps the best eight", () => {
    const crowd = people(
      ...Array.from({ length: 11 }, (_, index): [string, string] => [
        `Team Alex ${index + 1}`,
        `team.alex${index + 1}@example.invalid`,
      ]),
      ["Alex", "alex@example.invalid"],
    );
    const found = searchResults(peopleOnly(crowd), "alex");
    expect(flatten(found)).toHaveLength(MAX_RESULTS);
    // The best match is the last one listed, and still makes the cut, first.
    expect(labels(found)[0]).toBe("Alex");
  });

  it("applies the limit across groups, not to each", () => {
    const wide: SearchSources = {
      pages: Array.from({ length: 5 }, (_, index) => item(`Report page ${index}`, [], `p${index}`)),
      audits: Array.from({ length: 5 }, (_, index) =>
        item(`Report audit ${index}`, [], `a${index}`),
      ),
      people: Array.from({ length: 5 }, (_, index) => item(`Report ${index}`, [], `u${index}`)),
    };
    expect(flatten(searchResults(wide, "report"))).toHaveLength(8);
  });

  it("with nothing typed, shows the pages and no more than eight", () => {
    expect(pages.length).toBeGreaterThan(MAX_RESULTS);
    for (const typed of ["", "   "]) {
      const found = searchResults(sources, typed);
      expect(found.map((group) => group.id)).toEqual(["pages"]);
      expect(flatten(found)).toHaveLength(MAX_RESULTS);
      expect(ids(found)).toEqual(pages.slice(0, MAX_RESULTS).map((page) => page.id));
    }
  });
});

describe("pages", () => {
  it("finds every sidebar entry by its own label, first", () => {
    for (const group of navGroups)
      for (const entry of group.items)
        expect(ids(searchResults(sources, entry.label))[0], entry.label).toBe(entry.view);
  });

  it("lists every sidebar entry once, in sidebar order", () => {
    expect(pages.map((page) => page.id)).toEqual(
      navGroups.flatMap((group) => group.items.map((entry) => entry.view)),
    );
  });

  it("only has synonyms for pages that exist", () => {
    const views = navGroups.flatMap((group) => group.items.map((entry) => entry.view));
    for (const view of Object.keys(pageSynonyms)) expect(views).toContain(view);
  });

  it("finds Microsoft updates by 'settings', 'notifications' and 'service health'", () => {
    for (const typed of ["settings", "notifications", "service health"])
      expect(ids(searchResults(sources, typed)), typed).toContain("notifications");
    expect(ids(searchResults(sources, "notifications"))).toEqual(["notifications"]);
    expect(ids(searchResults(sources, "service health"))).toEqual(["notifications"]);
  });

  it("finds Microsoft access by 'permissions' and 'consent'", () => {
    expect(ids(searchResults(sources, "permissions"))).toEqual(["access"]);
    expect(ids(searchResults(sources, "consent"))).toEqual(["access"]);
  });

  it("finds a page by the section it sits under in the sidebar", () => {
    expect(ids(searchResults(sources, "monitor"))).toEqual([
      "audits",
      "signins",
      "changes",
      "handoffs",
    ]);
    expect(ids(searchResults(sources, "settings"))).toEqual(["notifications", "alerts", "access"]);
  });
});

describe("what each source shows", () => {
  it("shows an audit's latest result in muted text, and says when it never ran", () => {
    const [inactive, mfa] = sources.audits;
    expect(inactive?.detail).toBe("3 accounts, 2 inactive 30+ days");
    expect(mfa?.detail).toBe("Not run yet");
  });

  it("shows a person's principal name, once", () => {
    const [marcus, , , , shared] = sources.people;
    expect(marcus).toMatchObject({ label: "Marcus Rowe", detail: "marcus.rowe@example.invalid" });
    expect(shared).toMatchObject({ label: "shared.mailbox@example.invalid", detail: null });
  });
});

describe("grouping", () => {
  it("puts results under Pages, Audits and People, the group of the best match first", () => {
    const found = searchResults(peopleOnly(directory), "ro");
    expect(found.map((group) => group.heading)).toEqual(["People"]);
    const mixed = searchResults(sources, "inactive");
    expect(mixed.map((group) => group.heading)).toEqual(["Audits"]);
    const several = searchResults(
      {
        pages: [item("Accounts overview", [], "overview")],
        audits: [item("Accounts", [], "accounts.inactive")],
        people: [item("Accounts team", [], "u1")],
      },
      "accounts",
    );
    expect(several.map((group) => group.heading)).toEqual(["Audits", "Pages", "People"]);
  });

  it("keeps each group best first and flattens in the order shown", () => {
    const found = searchResults(peopleOnly(directory), "ro");
    const flat: FlatResult[] = flatten(found);
    expect(flat.map((entry) => entry.group)).toEqual(["people", "people"]);
    expect(flat.map((entry) => entry.item.label)).toEqual(["Rowan Ellis", "Marcus Rowe"]);
  });
});

describe("looking someone up by email", () => {
  it("accepts a full work email and nothing else", () => {
    expect(emailToLookUp("  lena.park@example.invalid ")).toBe("lena.park@example.invalid");
    expect(emailToLookUp("lena.park")).toBeNull();
    expect(emailToLookUp("lena park@example.invalid")).toBeNull();
    expect(emailToLookUp("@example.invalid")).toBeNull();
  });

  it("offers one lookup, first, when the email matches nobody stored", () => {
    const found = searchResults(sources, "lena.park@example.invalid", {
      email: "lena.park@example.invalid",
    });
    expect(found.map((group) => group.heading)).toEqual(["Directory"]);
    expect(labels(found)).toEqual(["Look up lena.park@example.invalid in the directory"]);
  });

  it("does not offer it when someone stored matches, or before the people have loaded", () => {
    const stored = searchResults(sources, "rowan.ellis@example.invalid", {
      email: "rowan.ellis@example.invalid",
    });
    expect(labels(stored)).toEqual(["Rowan Ellis"]);
    expect(labels(searchResults(sources, "lena.park@example.invalid", { email: null }))).toEqual(
      [],
    );
  });
});

describe("loading and failing", () => {
  const ready = { state: "ready", value: null } as const;

  it("says so under the results, per source, and never for a source that is fine", () => {
    expect(sourceNotes("ro", true, ready, ready)).toEqual([]);
    expect(sourceNotes("ro", true, { state: "failed" }, ready)).toEqual([
      "Audits could not be loaded.",
    ]);
    expect(sourceNotes("ro", true, ready, { state: "failed" })).toEqual([
      "People could not be loaded.",
    ]);
    expect(sourceNotes("ro", true, { state: "loading" }, { state: "loading" })).toEqual([
      "Loading audits…",
      "Loading people…",
    ]);
  });

  it("asks for a sign-in when audits and people cannot be read, and hints before typing", () => {
    expect(sourceNotes("ro", false, { state: "idle" }, { state: "idle" })).toEqual([
      "Sign in to search audits and people.",
    ]);
    expect(sourceNotes("", true, ready, ready)).toEqual([
      "Type to search audits and people as well.",
    ]);
  });

  it("describes the result count for a screen reader, and the empty result in one plain line", () => {
    expect(resultStatus("", 8)).toBe("");
    expect(resultStatus("ro", 0)).toBe("No results.");
    expect(resultStatus("ro", 1)).toBe("1 result.");
    expect(resultStatus("ro", 5)).toBe("5 results.");
    expect(noMatchText("  zzz ")).toBe("No page, audit or person matches “zzz”.");
  });
});

// ---------------------------------------------------------------------------------------------
// The dialog's markup
// ---------------------------------------------------------------------------------------------

function viewProps(overrides: Partial<QuickSearchViewProps> = {}) {
  const calls: string[] = [];
  const props: QuickSearchViewProps = {
    idBase: "qs",
    query: "ro",
    groups: searchResults(peopleOnly(directory), "ro"),
    active: 1,
    notes: [],
    empty: null,
    status: "2 results.",
    working: false,
    error: "",
    onQuery: (value) => calls.push(`query:${value}`),
    onKeyDown: () => calls.push("key"),
    onPick: (index) => calls.push(`pick:${index}`),
    onHover: (index) => calls.push(`hover:${index}`),
    onBackdrop: () => calls.push("backdrop"),
    ...overrides,
  };
  return { props, calls };
}

const html = (overrides: Partial<QuickSearchViewProps> = {}) =>
  renderToStaticMarkup(createElement(QuickSearchView, viewProps(overrides).props));
const tree = (overrides: Partial<QuickSearchViewProps> = {}) =>
  hostTree(createElement(QuickSearchView, viewProps(overrides).props));
const withRole = (nodes: ReturnType<typeof tree>, role: string) =>
  findAll(nodes, (host) => host.props.role === role);
const one = (hosts: Host[]): Host => {
  expect(hosts).toHaveLength(1);
  return hosts[0] as Host;
};

describe("the dialog's roles and aria wiring", () => {
  it("is a modal dialog named by its own heading", () => {
    const nodes = tree();
    const dialog = one(withRole(nodes, "dialog"));
    expect(dialog.props["aria-modal"]).toBe("true");
    const title = one(findAll(nodes, (host) => host.props.id === dialog.props["aria-labelledby"]));
    expect(textOf(title)).toBe("Search pages, audits and people");
  });

  it("is a combobox over a listbox, and points at the marked option", () => {
    const nodes = tree();
    const input = one(withRole(nodes, "combobox"));
    const list = one(withRole(nodes, "listbox"));
    expect(input.tag).toBe("input");
    expect(input.props["aria-expanded"]).toBe(true);
    expect(input.props["aria-autocomplete"]).toBe("list");
    expect(input.props["aria-controls"]).toBe(list.props.id);
    const options = withRole(nodes, "option");
    expect(options.map((option) => option.props.id)).toEqual([
      optionId("qs", 0),
      optionId("qs", 1),
    ]);
    expect(new Set(options.map((option) => option.props.id)).size).toBe(options.length);
    expect(input.props["aria-activedescendant"]).toBe(optionId("qs", 1));
    expect(options.map((option) => option.props["aria-selected"])).toEqual([false, true]);
  });

  it("labels each group by its heading", () => {
    const nodes = tree({
      groups: searchResults(
        {
          pages: [item("Accounts overview", [], "overview")],
          audits: [item("Accounts", [], "accounts.inactive")],
          people: [],
        },
        "accounts",
      ),
    });
    const groups = withRole(nodes, "group");
    expect(groups).toHaveLength(2);
    for (const group of groups) {
      const heading = one(
        findAll(nodes, (host) => host.props.id === group.props["aria-labelledby"]),
      );
      expect(["Audits", "Pages"]).toContain(textOf(heading));
      expect(
        group.children.every((child) => typeof child !== "string" && child.props.role === "option"),
      ).toBe(true);
    }
  });

  it("shows a result's muted text after its label", () => {
    const markup = html({ groups: searchResults(sources, "inactive"), active: 0 });
    expect(markup).toContain("Inactive accounts");
    expect(markup).toMatch(/quick-search__detail">3 accounts, 2 inactive 30\+ days</);
  });

  it("announces the count in a polite live region", () => {
    const status = one(withRole(tree(), "status"));
    expect(status.props["aria-live"]).toBe("polite");
    expect(textOf(status)).toBe("2 results.");
  });

  it("has no link or button inside another, and none of the punctuation the copy rules ban", () => {
    for (const markup of [html(), html({ groups: [], empty: "Nothing here." })]) {
      expect(nestedInteractive(markup)).toBe(0);
      expect(markup).not.toMatch(/[—·!]/);
    }
  });
});

describe("the empty, no-match and failed states", () => {
  it("before typing, lists the pages and a hint about the rest", () => {
    const markup = html({
      query: "",
      groups: searchResults(sources, ""),
      active: 0,
      notes: ["Type to search audits and people as well."],
      status: "",
    });
    expect(markup).toContain("Pages");
    expect(markup).toContain("Overview");
    expect(markup).toContain("Type to search audits and people as well.");
    expect(markup).not.toContain("Audits</div>");
  });

  it("with no match, shows one plain line and no listbox to point at", () => {
    const nodes = tree({
      query: "zzz",
      groups: [],
      empty: noMatchText("zzz"),
      status: "No results.",
    });
    expect(withRole(nodes, "listbox")).toHaveLength(0);
    const input = one(withRole(nodes, "combobox"));
    expect(input.props["aria-expanded"]).toBe(false);
    expect(input.props["aria-controls"]).toBeUndefined();
    expect(input.props["aria-activedescendant"]).toBeUndefined();
    const markup = html({ query: "zzz", groups: [], empty: noMatchText("zzz") });
    expect(markup).toContain("No page, audit or person matches “zzz”.");
  });

  it("when a source failed, still shows the others and says which could not load", () => {
    const markup = html({
      groups: searchResults({ ...sources, people: [] }, "inactive"),
      active: 0,
      notes: ["People could not be loaded."],
    });
    expect(markup).toContain("Inactive accounts");
    expect(markup).toContain("People could not be loaded.");
  });

  it("shows the email lookup as a result, under Directory", () => {
    const groups = searchResults(sources, "lena.park@example.invalid", {
      email: "lena.park@example.invalid",
    });
    const markup = html({ query: "lena.park@example.invalid", groups, active: 0 });
    expect(markup).toContain("Directory");
    expect(markup).toContain("Look up lena.park@example.invalid in the directory");
  });

  it("shows a lookup that is running, and one that failed, as an alert", () => {
    expect(html({ working: true })).toContain("Looking up…");
    const nodes = tree({ error: "No account has that email in the connected tenant." });
    const alert = one(withRole(nodes, "alert"));
    expect(textOf(alert)).toBe("No account has that email in the connected tenant.");
  });
});

// ---------------------------------------------------------------------------------------------
// Handlers
// ---------------------------------------------------------------------------------------------

type Handler = (event: unknown) => void;
const handlerOf = (host: Host, name: string) => host.props[name] as Handler;

function press(
  key: string,
  extra: Partial<Omit<KeyEventLike, "preventDefault">> = {},
): KeyEventLike & { prevented: boolean } {
  return {
    key,
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    shiftKey: false,
    defaultPrevented: false,
    target: null,
    ...extra,
    prevented: false,
    preventDefault() {
      this.prevented = true;
    },
  };
}

function recorder(state: { count: number; active: number }) {
  const log: string[] = [];
  const actions = {
    move: (to: number) => log.push(`move:${to}`),
    choose: (index: number) => log.push(`choose:${index}`),
    close: () => log.push("close"),
  };
  return {
    log,
    send: (event: KeyEventLike) => handleDialogKey(event, state, actions),
  };
}

describe("the dialog's handlers", () => {
  it("passes what is typed on", () => {
    const { props, calls } = viewProps();
    const input = one(withRole(hostTree(createElement(QuickSearchView, props)), "combobox"));
    handlerOf(input, "onChange")({ target: { value: "rowan" } });
    expect(calls).toEqual(["query:rowan"]);
  });

  it("opens the result that is clicked, and marks the one the pointer moves over", () => {
    const { props, calls } = viewProps();
    const options = withRole(hostTree(createElement(QuickSearchView, props)), "option");
    handlerOf(options[1] as Host, "onClick")({});
    handlerOf(options[0] as Host, "onMouseMove")({});
    expect(calls).toEqual(["pick:1", "hover:0"]);
  });

  it("closes on a press on the backdrop, and not on a press that began inside the panel", () => {
    const { props, calls } = viewProps();
    const nodes = hostTree(createElement(QuickSearchView, props));
    const backdrop = nodes[0] as Host;
    expect(backdrop.props.className).toBe("quick-search");
    const panel = one(withRole(nodes, "dialog"));
    const outer = {};
    handlerOf(backdrop, "onMouseDown")({ target: outer, currentTarget: outer });
    expect(calls).toEqual(["backdrop"]);
    // A press inside the panel bubbles to the backdrop with the panel as its target.
    handlerOf(backdrop, "onMouseDown")({ target: panel, currentTarget: outer });
    expect(calls).toEqual(["backdrop"]);
  });

  it("keeps focus in the input when the panel is pressed anywhere else", () => {
    const panel = one(withRole(tree(), "dialog"));
    let prevented = 0;
    const event = (target: unknown) => ({ target, preventDefault: () => (prevented += 1) });
    handlerOf(panel, "onMouseDown")(event({ tagName: "DIV" }));
    handlerOf(panel, "onMouseDown")(event({ tagName: "INPUT", type: "text" }));
    expect(prevented).toBe(1);
  });

  it("hands every key inside the panel to the key handler", () => {
    const { props, calls } = viewProps();
    const panel = one(withRole(hostTree(createElement(QuickSearchView, props)), "dialog"));
    handlerOf(panel, "onKeyDown")({ key: "ArrowDown" });
    expect(calls).toEqual(["key"]);
  });
});

describe("the shortcuts that open it", () => {
  it("opens on Cmd+K and on Ctrl+K, in either case, from anywhere", () => {
    expect(opensQuickSearch(press("k", { metaKey: true }))).toBe(true);
    expect(opensQuickSearch(press("k", { ctrlKey: true }))).toBe(true);
    expect(opensQuickSearch(press("K", { metaKey: true }))).toBe(true);
    // Inside another text box, as other tools do.
    expect(
      opensQuickSearch(press("k", { metaKey: true, target: { tagName: "INPUT", type: "text" } })),
    ).toBe(true);
    expect(opensQuickSearch(press("k", { ctrlKey: true, target: { tagName: "TEXTAREA" } }))).toBe(
      true,
    );
  });

  it("does not open on K alone or on other chords", () => {
    expect(opensQuickSearch(press("k"))).toBe(false);
    expect(opensQuickSearch(press("k", { metaKey: true, shiftKey: true }))).toBe(false);
    expect(opensQuickSearch(press("k", { ctrlKey: true, altKey: true }))).toBe(false);
    expect(opensQuickSearch(press("j", { metaKey: true }))).toBe(false);
  });

  it("opens on / only when focus is not in something you type into", () => {
    expect(opensQuickSearch(press("/"))).toBe(true);
    expect(opensQuickSearch(press("/", { target: { tagName: "BODY" } }))).toBe(true);
    expect(opensQuickSearch(press("/", { target: { tagName: "BUTTON" } }))).toBe(true);
    expect(opensQuickSearch(press("/", { target: { tagName: "INPUT", type: "checkbox" } }))).toBe(
      true,
    );
    // Some keyboards type "/" with Shift held.
    expect(opensQuickSearch(press("/", { shiftKey: true }))).toBe(true);
    for (const target of [
      { tagName: "INPUT", type: "text" },
      { tagName: "INPUT", type: "search" },
      { tagName: "INPUT" },
      { tagName: "TEXTAREA" },
      { tagName: "SELECT" },
      { tagName: "DIV", isContentEditable: true },
    ])
      expect(opensQuickSearch(press("/", { target })), JSON.stringify(target)).toBe(false);
  });

  it("leaves / with a modifier, a handled key and an input method alone", () => {
    expect(opensQuickSearch(press("/", { ctrlKey: true }))).toBe(false);
    expect(opensQuickSearch(press("/", { metaKey: true }))).toBe(false);
    expect(opensQuickSearch(press("/", { defaultPrevented: true }))).toBe(false);
    expect(opensQuickSearch(press("k", { metaKey: true, isComposing: true }))).toBe(false);
  });

  it("recognises what counts as a text field", () => {
    expect(isTextField(null)).toBe(false);
    expect(isTextField("input")).toBe(false);
    expect(isTextField({ tagName: "input", type: "EMAIL" })).toBe(true);
    expect(isTextField({ tagName: "INPUT", type: "radio" })).toBe(false);
    expect(isTextField({ tagName: "A" })).toBe(false);
  });

  it("claims the shortcut so the browser's own Ctrl+K does not fire, and opens once", () => {
    const opened: boolean[] = [];
    const first = press("k", { metaKey: true });
    openFromKeyboard(first, false, (value) => opened.push(value));
    expect(first.prevented).toBe(true);
    expect(opened).toEqual([true]);
    // Already open: still claimed, nothing to open.
    const again = press("k", { ctrlKey: true });
    openFromKeyboard(again, true, (value) => opened.push(value));
    expect(again.prevented).toBe(true);
    expect(opened).toEqual([true]);
    // Not a shortcut: untouched.
    const typing = press("a");
    openFromKeyboard(typing, false, (value) => opened.push(value));
    expect(typing.prevented).toBe(false);
    expect(opened).toEqual([true]);
  });
});

describe("keys inside the dialog", () => {
  it("moves down and up, and wraps at both ends", () => {
    const state = { count: 3, active: 0 };
    const { log, send } = recorder(state);
    send(press("ArrowDown"));
    state.active = 2;
    send(press("ArrowDown"));
    state.active = 0;
    send(press("ArrowUp"));
    state.active = 2;
    send(press("ArrowUp"));
    expect(log).toEqual(["move:1", "move:0", "move:2", "move:1"]);
  });

  it("stops the page and the cursor from moving with the arrows", () => {
    const { send } = recorder({ count: 3, active: 0 });
    const down = press("ArrowDown");
    send(down);
    expect(down.prevented).toBe(true);
  });

  it("does nothing with the arrows or Enter when there is no result", () => {
    const { log, send } = recorder({ count: 0, active: 0 });
    send(press("ArrowDown"));
    send(press("ArrowUp"));
    send(press("Enter"));
    expect(log).toEqual([]);
  });

  it("opens the marked result on Enter", () => {
    const { log, send } = recorder({ count: 4, active: 2 });
    const enter = press("Enter");
    send(enter);
    expect(log).toEqual(["choose:2"]);
    expect(enter.prevented).toBe(true);
  });

  it("does not treat the Enter that confirms an input method's word as a request to open", () => {
    const { log, send } = recorder({ count: 4, active: 2 });
    send(press("Enter", { isComposing: true }));
    send(press("Enter", { nativeEvent: { isComposing: true } }));
    expect(log).toEqual([]);
  });

  it("leaves the arrows and Enter with a modifier to the browser", () => {
    const { log, send } = recorder({ count: 4, active: 1 });
    send(press("ArrowDown", { metaKey: true }));
    send(press("ArrowUp", { altKey: true }));
    send(press("Enter", { ctrlKey: true }));
    expect(log).toEqual([]);
  });

  it("closes on Escape", () => {
    const { log, send } = recorder({ count: 4, active: 1 });
    const escapeKey = press("Escape");
    send(escapeKey);
    expect(log).toEqual(["close"]);
    expect(escapeKey.prevented).toBe(true);
  });

  it("keeps Tab inside, where the input is the only control", () => {
    const { log, send } = recorder({ count: 4, active: 1 });
    for (const shiftKey of [false, true]) {
      const tab = press("Tab", { shiftKey });
      send(tab);
      expect(tab.prevented).toBe(true);
    }
    expect(log).toEqual([]);
  });

  it("ignores other keys, so typing reaches the input", () => {
    const { log, send } = recorder({ count: 4, active: 1 });
    const letter = press("r");
    send(letter);
    expect(log).toEqual([]);
    expect(letter.prevented).toBe(false);
  });
});

describe("giving focus back", () => {
  function element(connected = true) {
    const calls: string[] = [];
    return { calls, isConnected: connected, focus: () => calls.push("focus") };
  }

  it("returns focus to what had it when Escape closes the dialog", () => {
    const opener = element();
    const keeper = focusKeeper(() => opener);
    keeper.remember();
    // The wiring the dialog uses: closing restores focus before the dialog goes.
    const closed: string[] = [];
    const escapeKey = press("Escape");
    handleDialogKey(
      escapeKey,
      { count: 2, active: 0 },
      {
        move: () => {},
        choose: () => {},
        close: () => {
          keeper.restore();
          closed.push("closed");
        },
      },
    );
    expect(closed).toEqual(["closed"]);
    expect(opener.calls).toEqual(["focus"]);
  });

  it("gives focus back once, and not to an element that has left the page", () => {
    const opener = element();
    const keeper = focusKeeper(() => opener);
    keeper.remember();
    keeper.restore();
    keeper.restore();
    expect(opener.calls).toEqual(["focus"]);

    const gone = element(false);
    const second = focusKeeper(() => gone);
    second.remember();
    second.restore();
    expect(gone.calls).toEqual([]);
  });

  it("copes with nothing to return to", () => {
    for (const active of [null, undefined, {}, "body"]) {
      const keeper = focusKeeper(() => active);
      keeper.remember();
      expect(() => keeper.restore()).not.toThrow();
    }
  });
});

// ---------------------------------------------------------------------------------------------
// The button, the sidebar and the demo
// ---------------------------------------------------------------------------------------------

describe("the shortcut hint", () => {
  it("shows Cmd on a Mac and Ctrl elsewhere", () => {
    expect(isMacPlatform({ platform: "MacIntel" })).toBe(true);
    expect(isMacPlatform({ platform: "iPhone" })).toBe(true);
    expect(isMacPlatform({ platform: "Win32" })).toBe(false);
    expect(isMacPlatform({ platform: "Linux x86_64" })).toBe(false);
    expect(isMacPlatform(undefined)).toBe(false);
    // The newer property wins when a browser has both.
    expect(isMacPlatform({ platform: "Win32", userAgentData: { platform: "macOS" } })).toBe(true);
    expect(isMacPlatform({ platform: "MacIntel", userAgentData: { platform: "Windows" } })).toBe(
      false,
    );
    expect(shortcutHint(true)).toBe("⌘K");
    expect(shortcutHint(false)).toBe("Ctrl K");
  });

  it("is a button named Search, with the hint hidden from screen readers", () => {
    const nodes = hostTree(createElement(QuickSearchButton, { mac: true, onOpen: () => {} }));
    const button = one(findAll(nodes, (host) => host.tag === "button"));
    expect(button.props["aria-haspopup"]).toBe("dialog");
    expect(button.props["aria-keyshortcuts"]).toBe("Meta+K");
    const hidden = one(
      findAll(nodes, (host) => host.props["aria-hidden"] === "true" && host.tag === "span"),
    );
    expect(textOf(hidden)).toBe("⌘K");
    const label = one(
      findAll(nodes, (host) => host.props.className === "quick-search-open__label"),
    );
    expect(textOf(label)).toBe("Search");
    const other = hostTree(createElement(QuickSearchButton, { mac: false, onOpen: () => {} }));
    expect(textOf(one(findAll(other, (host) => host.tag === "kbd")))).toBe("Ctrl K");
  });

  it("opens the dialog when pressed", () => {
    let opened = 0;
    const nodes = hostTree(
      createElement(QuickSearchButton, { mac: false, onOpen: () => (opened += 1) }),
    );
    handlerOf(one(findAll(nodes, (host) => host.tag === "button")), "onClick")({});
    expect(opened).toBe(1);
  });
});

describe("the sidebar", () => {
  const markup = renderToStaticMarkup(createElement(App));

  it("has the Search button at the top, above the navigation and outside it", () => {
    const button = markup.indexOf('class="quick-search-open"');
    const nav = markup.indexOf("<nav");
    const brand = markup.indexOf('class="brand"');
    expect(brand).toBeGreaterThan(-1);
    expect(button).toBeGreaterThan(brand);
    expect(nav).toBeGreaterThan(button);
    expect(markup.slice(nav, markup.indexOf("</nav>"))).not.toContain("quick-search-open");
    expect(markup.slice(button, nav)).toContain("Search");
  });

  it("draws no dialog until it is opened", () => {
    expect(markup).not.toContain('role="dialog"');
    expect(markup).not.toContain("quick-search__panel");
  });
});

describe("the dialog component", () => {
  function session(overrides: Partial<PrismSession> = {}): PrismSession {
    return sessionSchema.parse({
      mode: "live",
      configured: true,
      authenticated: true,
      displayName: "Synthetic admin",
      csrfToken: "synthetic-csrf-token-value-000000000000",
      enabledFeatures: ["audit.signins.read"],
      missingConfiguration: [],
      ...overrides,
    });
  }
  const render = (open: boolean, value: PrismSession | null) =>
    renderToStaticMarkup(
      createElement(QuickSearch, {
        open,
        onOpenChange: () => {},
        session: value,
        pages,
        onNavigate: () => {},
        onOpenAudit: () => {},
        onOpenPerson: () => {},
      }),
    );

  it("draws nothing while closed", () => {
    expect(render(false, session())).toBe("");
  });

  it("opens on the pages, with a hint about the rest, and wires its ids together", () => {
    const markup = render(true, session());
    expect(markup).toContain('role="dialog"');
    expect(markup).toContain("Overview");
    expect(markup).toContain("Type to search audits and people as well.");
    const controls = markup.match(/aria-controls="([^"]+)"/)?.[1];
    const active = markup.match(/aria-activedescendant="([^"]+)"/)?.[1];
    expect(controls).toBeTruthy();
    expect(markup).toContain(`id="${controls}"`);
    expect(markup).toContain(`id="${active}"`);
    expect(markup.match(/role="option"/g)).toHaveLength(MAX_RESULTS);
  });

  it("still lists pages when nobody is signed in, and says audits and people need a sign-in", () => {
    for (const value of [null, session({ authenticated: false })]) {
      const markup = render(true, value);
      expect(markup).toContain("Overview");
      expect(markup).toContain("Sign in to search audits and people.");
    }
  });
});

describe("the demo", () => {
  const fixtures = JSON.parse(
    readFileSync(new URL("../demo/fixtures.json", import.meta.url), "utf8"),
  ) as Fixtures;
  const state = createDemoState(fixtures, () => new Date(fixtures.recordedAt));

  it("reads the two requests in exactly the forms the recording holds", () => {
    expect(auditsUrl).toBe("/api/v1/audits");
    expect(peopleUrl).toBe("/api/v1/people?days=30");
    for (const url of [auditsUrl, peopleUrl])
      expect(Object.keys(fixtures.responses)).toContain(normalizeKey("GET", url));
  });

  it("answers both, in shapes the dialog accepts", () => {
    const audits = handleRequest(state, { method: "GET", url: auditsUrl });
    expect(audits.status).toBe(200);
    expect(auditItems(auditCatalogSchema.parse(audits.body)).length).toBeGreaterThan(0);
    const answer = handleRequest(state, { method: "GET", url: peopleUrl });
    expect(answer.status).toBe(200);
    expect(personItems(personListSchema.parse(answer.body)).length).toBeGreaterThan(0);
  });
});

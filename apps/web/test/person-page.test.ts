import { readFileSync } from "node:fs";
import {
  type Finding,
  type FindingState,
  findingSchema,
  type PrismSession,
  personDirectorySchema,
  personHistorySchema,
  personSigninSchema,
  type Severity,
  sessionSchema,
} from "@prism365/contracts";
import { createElement, isValidElement, type ReactNode } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import {
  ActionsMenuView,
  activateItem,
  handleMenuKey,
  isOutside,
  type MenuItem,
  type MenuKeyContext,
  menuKeyTarget,
} from "../src/ActionsMenu.tsx";
import type { PersonCanary } from "../src/Canary.tsx";
import { People, PersonProfileView, type PersonProfileViewProps } from "../src/People.tsx";
import { PersonChanges } from "../src/PersonChanges.tsx";
import {
  entraUserLink,
  hibpLink,
  initials,
  PersonHeader,
  PersonStatusChips,
  personAddress,
  personMenuItems,
  severityCounts,
  userTypeLabel,
} from "../src/PersonHeader.tsx";
import { PersonReceipts } from "../src/PersonReceipts.tsx";
import {
  handleTabKey,
  type PersonTab,
  PersonTabList,
  PersonTabPanel,
  panelId,
  parsePersonTab,
  personTabCounts,
  personTabs,
  type TabCount,
  tabKeyTarget,
} from "../src/PersonTabs.tsx";
import { findAll, type Host, hostTree, nestedInteractive, textOf } from "./markup.ts";

const noop = () => {};

// Calls a handler that a drawn element was given, and fails the test when it has none.
function fire(host: Host | undefined, name: string, ...args: unknown[]): void {
  const handler = host?.props[name];
  if (typeof handler !== "function") throw new Error(`The element has no ${name}.`);
  handler(...args);
}

const personId = "44444444-4444-4444-8444-444444444444";
const stamp = "2026-10-20T10:00:00.000Z";

// What a reader sees: tags and React's text-node separators removed.
const words = (html: string) => {
  let current = html.replaceAll("<!-- -->", "");
  let previous: string;
  do {
    previous = current;
    current = current.replace(/<[^>]*>/g, " ");
  } while (current !== previous);
  return current.replace(/&#x27;/g, "'").replace(/\s+/g, " ");
};

function session(): PrismSession {
  return sessionSchema.parse({
    mode: "live",
    configured: true,
    authenticated: true,
    displayName: "Synthetic admin",
    csrfToken: "synthetic-csrf-token-value-000000000000",
    enabledFeatures: ["directory.users.read", "audit.signins.read"],
    missingConfiguration: [],
  });
}

// ---------------------------------------------------------------------------------------------
// The header
// ---------------------------------------------------------------------------------------------

const account = {
  enabled: true as boolean | null,
  userType: "Member" as string | null,
  synced: false as boolean | null,
};

const chips = (overrides: Partial<Parameters<typeof PersonStatusChips>[0]> = {}) =>
  renderToStaticMarkup(
    createElement(PersonStatusChips, {
      account,
      openSeverities: null,
      canary: false,
      ...overrides,
    }),
  );

describe("status chips", () => {
  it("say whether the account is enabled or disabled", () => {
    expect(chips()).toContain('<span class="person-tag person-tag--ok">Enabled</span>');
    const disabled = chips({ account: { ...account, enabled: false } });
    expect(disabled).toMatch(/person-tag--alert"[^>]*>Disabled</);
    expect(disabled).not.toContain("Enabled");
    // An account whose state the directory did not report gets neither chip.
    const unknown = chips({ account: { ...account, enabled: null } });
    expect(unknown).not.toContain("Enabled");
    expect(unknown).not.toContain("Disabled");
  });

  it("say member or guest, and show any other type as Microsoft gave it", () => {
    expect(chips()).toContain(">Member<");
    expect(chips({ account: { ...account, userType: "guest" } })).toContain(">Guest<");
    expect(chips({ account: { ...account, userType: "Contractor" } })).toContain(">Contractor<");
    expect(chips({ account: { ...account, userType: null } })).not.toMatch(/Member|Guest/);
    expect(userTypeLabel("MEMBER")).toBe("Member");
  });

  it("say synced or cloud", () => {
    expect(chips()).toContain(">Cloud account<");
    expect(chips({ account: { ...account, synced: true } })).toContain(">Synced from on-premises<");
    expect(chips({ account: { ...account, synced: null } })).toContain(">Cloud account<");
  });

  it("leave the account chips out until the directory has answered", () => {
    const html = chips({ account: null });
    expect(html).not.toMatch(/Enabled|Disabled|Member|Cloud account|Synced/);
  });

  it("show the canary chip only when the account is marked", () => {
    expect(chips()).not.toContain("Canary");
    expect(chips({ canary: true })).toMatch(/person-tag--notice"[^>]*>Canary</);
  });

  it("count open findings by severity in the shared severity style, worst first", () => {
    const html = chips({ openSeverities: ["low", "critical", "medium", "critical", "info"] });
    expect(words(html)).toContain("Open findings 2 critical 1 medium 1 low 1 info");
    expect(html).toContain('<span class="sev sev--critical">');
    expect(html).toContain('<span class="sev sev--medium">');
    // Info is named as info but drawn in low's colour.
    expect(html).not.toContain("sev--info");
    expect(html.match(/sev--low/g)).toHaveLength(2);
    expect(html).not.toContain("Nothing open");
  });

  it("say nothing is open once the record has loaded, and nothing before", () => {
    expect(chips({ openSeverities: [] })).toContain("Nothing open");
    expect(chips({ openSeverities: null })).not.toMatch(/Nothing open|Open findings/);
  });

  it("orders and groups severities without depending on the input order", () => {
    expect(severityCounts(["low", "critical", "high", "high"])).toEqual([
      { severity: "critical", count: 1 },
      { severity: "high", count: 2 },
      { severity: "low", count: 1 },
    ]);
    expect(severityCounts([])).toEqual([]);
  });
});

describe("initials", () => {
  it("take the first letters of the first and last words", () => {
    expect(initials("Avery Stone")).toBe("AS");
    expect(initials("avery stone")).toBe("AS");
    expect(initials("Mary Jane Watson")).toBe("MW");
    expect(initials("Avery")).toBe("A");
  });

  it("read an address by its local part, never its domain", () => {
    expect(initials("avery.stone@example.invalid")).toBe("AS");
    expect(initials("support@example.invalid")).toBe("S");
  });

  it("show a question mark for an id or nothing", () => {
    expect(initials(personId)).toBe("?");
    expect(initials("   ")).toBe("?");
    expect(initials("1234")).toBe("?");
  });
});

describe("addresses and links", () => {
  it("links to the person's page in the Entra admin center, encoding the id", () => {
    expect(entraUserLink(personId)).toBe(
      `https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/overview/userId/${personId}`,
    );
    expect(entraUserLink("a/b?c")).toContain("/userId/a%2Fb%3Fc");
  });

  it("keeps the breach link as it was", () => {
    expect(hibpLink("avery.stone@example.invalid")).toBe(
      "https://haveibeenpwned.com/account/avery.stone@example.invalid",
    );
  });

  it("uses the mailbox, else the sign-in name, and skips guest sign-in names", () => {
    expect(personAddress("mail@example.invalid", "upn@example.invalid")).toBe(
      "mail@example.invalid",
    );
    expect(personAddress(null, "upn@example.invalid")).toBe("upn@example.invalid");
    expect(personAddress(null, "guest_other.example#EXT#@tenant.example.invalid")).toBeNull();
    expect(personAddress("", "no-at-sign")).toBeNull();
  });
});

describe("person header", () => {
  const header = (overrides: Partial<Parameters<typeof PersonHeader>[0]> = {}) =>
    renderToStaticMarkup(
      createElement(PersonHeader, {
        titleId: "title",
        name: "Avery Stone",
        principal: "avery.stone@example.invalid",
        jobTitle: "Support analyst",
        department: "Operations",
        account,
        openSeverities: [],
        canary: false,
        actions: createElement("span", { id: "slot" }, "menu goes here"),
        ...overrides,
      }),
    );

  it("shows the avatar, name, address with its copy button, and role", () => {
    const html = header();
    expect(html).toContain('class="person-avatar" aria-hidden="true">AS<');
    expect(html).toContain('<h2 id="title">Avery Stone</h2>');
    expect(html).toContain('<span class="copy-line__text">avery.stone@example.invalid</span>');
    expect(html).toContain('aria-label="Copy avery.stone@example.invalid"');
    // Title and department are separate pieces, divided by the stylesheet and not by a dot.
    expect(html).toContain(
      '<p class="person-header__role"><span>Support analyst</span><span>Operations</span></p>',
    );
    expect(words(html)).not.toMatch(/[—–·!]/);
  });

  it("holds the menu on the right and nothing below it unless asked", () => {
    const html = header();
    expect(html).toMatch(/person-header__actions"><span id="slot">menu goes here<\/span>/);
    expect(html).not.toContain("person-header__below");
    expect(header({ below: createElement("p", null, "Are you sure?") })).toContain(
      '<div class="person-header__below"><p>Are you sure?</p></div>',
    );
  });

  it("does not repeat an address that is already the name", () => {
    const html = header({
      name: "avery.stone@example.invalid",
      jobTitle: null,
      department: null,
    });
    expect(html.match(/aria-label="Copy /g)).toHaveLength(1);
    expect(html.match(/avery\.stone@example\.invalid/g)).toHaveLength(3);
    expect(html).not.toContain("person-header__principal");
  });

  it("says so when the directory has nothing on the person", () => {
    const html = header({ principal: null, jobTitle: null, department: null, account: null });
    expect(html).toContain("No directory details");
    expect(html).not.toContain("Copy ");
  });

  it("shows the chips, the canary one included", () => {
    const html = header({ canary: true, openSeverities: ["high"] });
    expect(html).toContain("Canary");
    expect(words(html)).toContain("Open findings 1 high");
  });
});

// ---------------------------------------------------------------------------------------------
// The Actions menu
// ---------------------------------------------------------------------------------------------

function canaryState(overrides: Partial<PersonCanary> = {}): PersonCanary {
  return {
    available: true,
    active: false,
    busy: false,
    error: "",
    confirming: false,
    onMark: noop,
    onConfirm: noop,
    onCancel: noop,
    onUnmark: noop,
    ...overrides,
  };
}

const menuItems = (overrides: Partial<Parameters<typeof personMenuItems>[0]> = {}) =>
  personMenuItems({
    userId: personId,
    address: "avery.stone@example.invalid",
    canary: canaryState(),
    copy: "idle",
    onCopy: noop,
    ...overrides,
  });

describe("the Actions menu items", () => {
  it("are, in order: canary, breach search, copy email, Entra", () => {
    expect(menuItems().map((item) => item.label)).toEqual([
      "Mark as canary",
      "Check Have I Been Pwned",
      "Copy email",
      "Open in Microsoft Entra",
    ]);
  });

  it("carry the canary explanation as the second line, and flip when the account is a canary", () => {
    const [mark] = menuItems();
    expect(mark?.hint).toContain("A decoy account nobody should ever sign in to.");
    const [stop] = menuItems({ canary: canaryState({ active: true }) });
    expect(stop?.label).toBe("Stop treating as canary");
    expect(stop?.hint).toBe(mark?.hint);
  });

  it("marking asks first and unmarking acts at once, each through its own handler", () => {
    const calls: string[] = [];
    const handlers = { onMark: () => calls.push("mark"), onUnmark: () => calls.push("unmark") };
    const [mark] = menuItems({ canary: canaryState(handlers) });
    const [stop] = menuItems({ canary: canaryState({ ...handlers, active: true }) });
    if (mark?.kind !== "button" || stop?.kind !== "button") throw new Error("Expected buttons.");
    mark.onSelect();
    stop.onSelect();
    expect(calls).toEqual(["mark", "unmark"]);
  });

  it("turns the canary item off while the list is busy, and leaves it out when there is no list", () => {
    const [busy] = menuItems({ canary: canaryState({ busy: true }) });
    expect(busy?.kind === "button" && busy.disabled).toBe(true);
    expect(menuItems({ canary: canaryState({ available: false }) }).map((i) => i.label)).toEqual([
      "Check Have I Been Pwned",
      "Copy email",
      "Open in Microsoft Entra",
    ]);
  });

  it("opens the same breach search and the Entra page as plain links", () => {
    const items = menuItems();
    const breach = items.find((item) => item.id === "breach");
    const entra = items.find((item) => item.id === "entra");
    expect(breach).toMatchObject({
      kind: "link",
      href: "https://haveibeenpwned.com/account/avery.stone@example.invalid",
    });
    expect(breach?.hint).toBe(
      "Opens haveibeenpwned.com in a new tab with avery.stone@example.invalid already searched.",
    );
    expect(entra).toMatchObject({ kind: "link", href: entraUserLink(personId) });
  });

  it("says what the copy did, and stays open to show it", () => {
    const copy = (status: "idle" | "copied" | "failed") =>
      menuItems({ copy: status }).find((item) => item.id === "copy");
    expect(copy("idle")).toMatchObject({
      label: "Copy email",
      hint: "avery.stone@example.invalid",
    });
    expect(copy("copied")).toMatchObject({ label: "Email copied", keepOpen: true });
    expect(copy("failed")?.label).toBe("Copy failed");
    expect(copy("failed")?.hint).toContain("copy it by hand");
  });

  it("leave out the breach search and the copy when the person has no mail address", () => {
    expect(menuItems({ address: null }).map((item) => item.label)).toEqual([
      "Mark as canary",
      "Open in Microsoft Entra",
    ]);
  });

  it("read like a colleague wrote them", () => {
    const text = menuItems()
      .flatMap((item) => [item.label, item.hint ?? ""])
      .join(" ");
    expect(text).not.toMatch(/[—–·!]/);
  });
});

const sampleItems: MenuItem[] = [
  { kind: "button", id: "one", label: "First", hint: "Does one thing.", onSelect: noop },
  { kind: "link", id: "two", label: "Second", href: "https://example.invalid/two" },
];

const menuView = (overrides: Partial<Parameters<typeof ActionsMenuView>[0]> = {}) =>
  hostTree(
    createElement(ActionsMenuView, {
      label: "Actions",
      open: false,
      buttonId: "b",
      menuId: "m",
      items: sampleItems,
      onToggle: noop,
      onActivate: noop,
      onKeyDown: noop,
      ...overrides,
    }),
  );

describe("the Actions menu markup", () => {
  it("is a disclosure: the button says whether the list is open and which list it controls", () => {
    for (const open of [false, true]) {
      const tree = menuView({ open });
      const button = findAll(tree, (host) => host.props.id === "b")[0];
      const list = findAll(tree, (host) => host.tag === "ul")[0];
      expect(button?.tag).toBe("button");
      expect(button?.props["aria-expanded"]).toBe(open);
      expect(button?.props["aria-controls"]).toBe("m");
      expect(list?.props.id).toBe("m");
      expect(list?.props["aria-labelledby"]).toBe("b");
      // Closed, the list stays in the page, hidden, so the button's reference is never dangling.
      expect(list?.props.hidden).toBe(!open);
    }
  });

  it("is a list of buttons and links, the links opening in a new tab", () => {
    const tree = menuView({ open: true });
    const button = findAll(tree, (host) => host.props.className === "actions-menu__item")[0];
    expect(button?.tag).toBe("button");
    expect(textOf(button as never)).toBe("FirstDoes one thing.");
    const link = findAll(tree, (host) => host.tag === "a")[0];
    expect(link?.props.href).toBe("https://example.invalid/two");
    expect(link?.props.target).toBe("_blank");
    expect(link?.props.rel).toBe("noreferrer noopener");
    expect(
      findAll(tree, (host) => host.props.role === "menu" || host.props.role === "menuitem"),
    ).toEqual([]);
  });

  it("hands clicks, the toggle and keys to the handlers it was given", () => {
    const seen: string[] = [];
    const tree = menuView({
      open: true,
      onToggle: () => seen.push("toggle"),
      onActivate: (item) => seen.push(`activate ${item.id}`),
      onKeyDown: (event) => seen.push(`key ${event.key}`),
    });
    fire(findAll(tree, (host) => host.props.id === "b")[0], "onClick");
    for (const host of findAll(tree, (h) => h.props.className === "actions-menu__item"))
      fire(host, "onClick");
    fire(findAll(tree, (host) => host.props.className === "actions-menu")[0], "onKeyDown", {
      key: "ArrowDown",
    });
    expect(seen).toEqual(["toggle", "activate one", "activate two", "key ArrowDown"]);
  });

  it("turns a disabled item off", () => {
    const tree = menuView({
      open: true,
      items: [{ kind: "button", id: "x", label: "X", disabled: true, onSelect: noop }],
    });
    expect(
      findAll(tree, (host) => host.props.className === "actions-menu__item")[0]?.props.disabled,
    ).toBe(true);
  });
});

function keyContext(overrides: Partial<MenuKeyContext> = {}) {
  const log: string[] = [];
  const items = ["a", "b", "c"].map((name) => ({ focus: () => log.push(`focus ${name}`) }));
  const context: MenuKeyContext = {
    open: true,
    items,
    current: -1,
    openMenu: (focus) => log.push(`open ${focus}`),
    closeMenu: () => log.push("close"),
    focusButton: () => log.push("button"),
    ...overrides,
  };
  return { context, log };
}

describe("the Actions menu keys", () => {
  it("steps down and up and wraps at both ends", () => {
    expect(menuKeyTarget("ArrowDown", -1, 3)).toBe(0);
    expect(menuKeyTarget("ArrowDown", 0, 3)).toBe(1);
    expect(menuKeyTarget("ArrowDown", 2, 3)).toBe(0);
    expect(menuKeyTarget("ArrowUp", 2, 3)).toBe(1);
    expect(menuKeyTarget("ArrowUp", 0, 3)).toBe(2);
    expect(menuKeyTarget("ArrowUp", -1, 3)).toBe(2);
    expect(menuKeyTarget("Home", 2, 3)).toBe(0);
    expect(menuKeyTarget("End", 0, 3)).toBe(2);
    expect(menuKeyTarget("a", 0, 3)).toBeNull();
    expect(menuKeyTarget("ArrowDown", 0, 0)).toBeNull();
  });

  it("moves the focus between the items while open", () => {
    const first = keyContext({ current: -1 });
    expect(handleMenuKey("ArrowDown", first.context)).toBe(true);
    expect(first.log).toEqual(["focus a"]);
    const wrap = keyContext({ current: 2 });
    handleMenuKey("ArrowDown", wrap.context);
    expect(wrap.log).toEqual(["focus a"]);
    const up = keyContext({ current: 0 });
    handleMenuKey("ArrowUp", up.context);
    expect(up.log).toEqual(["focus c"]);
    const end = keyContext({ current: 0 });
    handleMenuKey("End", end.context);
    expect(end.log).toEqual(["focus c"]);
  });

  it("closes on Escape and gives the focus back to the button", () => {
    const { context, log } = keyContext({ current: 1 });
    expect(handleMenuKey("Escape", context)).toBe(true);
    expect(log).toEqual(["close", "button"]);
  });

  it("leaves Escape alone when the list is closed, so a dialog or page can use it", () => {
    const { context, log } = keyContext({ open: false });
    expect(handleMenuKey("Escape", context)).toBe(false);
    expect(log).toEqual([]);
  });

  it("opens a closed list with the arrows only, landing on the first or last item", () => {
    const down = keyContext({ open: false });
    expect(handleMenuKey("ArrowDown", down.context)).toBe(true);
    expect(down.log).toEqual(["open first"]);
    const up = keyContext({ open: false });
    expect(handleMenuKey("ArrowUp", up.context)).toBe(true);
    expect(up.log).toEqual(["open last"]);
    const home = keyContext({ open: false });
    expect(handleMenuKey("Home", home.context)).toBe(false);
    expect(home.log).toEqual([]);
  });

  it("ignores other keys, so typing and tabbing keep working", () => {
    for (const key of ["Tab", "Enter", " ", "a"]) {
      const { context, log } = keyContext();
      expect(handleMenuKey(key, context), key).toBe(false);
      expect(log).toEqual([]);
    }
  });

  it("does nothing on an arrow when no item can take focus", () => {
    const { context, log } = keyContext({ items: [] });
    expect(handleMenuKey("ArrowDown", context)).toBe(false);
    expect(log).toEqual([]);
  });
});

describe("closing the Actions menu", () => {
  it("closes on a click or focus outside it, and not inside", () => {
    const inside = {};
    const container = { contains: (node: unknown) => node === inside };
    expect(isOutside(container as never, {} as EventTarget)).toBe(true);
    expect(isOutside(container as never, inside as EventTarget)).toBe(false);
    expect(isOutside(container as never, null)).toBe(true);
    // Nothing to compare against before the menu is on the page.
    expect(isOutside(null, {} as EventTarget)).toBe(false);
  });

  it("closes after a choice and returns the focus, unless the item stays open", () => {
    const log: string[] = [];
    const controls = { closeMenu: () => log.push("close"), focusButton: () => log.push("button") };
    activateItem(
      { kind: "button", id: "a", label: "A", onSelect: () => log.push("run") },
      controls,
    );
    activateItem({ kind: "link", id: "b", label: "B", href: "https://example.invalid" }, controls);
    activateItem(
      { kind: "button", id: "c", label: "C", keepOpen: true, onSelect: () => log.push("stay") },
      controls,
    );
    expect(log).toEqual(["run", "close", "button", "close", "button", "stay"]);
  });
});

// ---------------------------------------------------------------------------------------------
// The tabs
// ---------------------------------------------------------------------------------------------

describe("the tab in the address", () => {
  it("round-trips every tab", async () => {
    const { parseRoute, routeHash } = await import("../src/App.tsx");
    for (const tab of personTabs) {
      const hash = tab === "overview" ? `#/people/${personId}` : `#/people/${personId}/${tab}`;
      const route = parseRoute(hash);
      expect(route?.view).toBe("people");
      expect(route?.person).toBe(personId);
      expect(route?.tab ?? "overview").toBe(tab);
      expect(routeHash({ view: "people", person: personId, audit: null, tab })).toBe(hash);
    }
  });

  it("opens the Overview for a plain link, as links made before the tabs did", async () => {
    const { parseRoute, routeHash } = await import("../src/App.tsx");
    const route = parseRoute(`#/people/${personId}`);
    expect(route).toEqual({ view: "people", person: personId, audit: null });
    expect(route).not.toHaveProperty("tab");
    expect(routeHash({ view: "people", person: personId, audit: null })).toBe(
      `#/people/${personId}`,
    );
    // The Overview is written without a segment, so one address means one page.
    expect(parseRoute(`#/people/${personId}/overview`)).not.toHaveProperty("tab");
    expect(parseRoute("#/people")).toEqual({ view: "people", person: null, audit: null });
  });

  it("keeps ids that were links before, including ones with an encoded slash", async () => {
    const { parseRoute, routeHash } = await import("../src/App.tsx");
    expect(parseRoute("#/people/u%3A1")).toEqual({ view: "people", person: "u:1", audit: null });
    expect(parseRoute("#/people/a%2Fb/signins")).toEqual({
      view: "people",
      person: "a/b",
      audit: null,
      tab: "signins",
    });
    expect(routeHash({ view: "people", person: "a/b", audit: null, tab: "signins" })).toBe(
      "#/people/a%2Fb/signins",
    );
    // An id that happens to be a tab's name is still an id when it stands alone, and a last
    // segment that is no tab stays part of the id, as it always did.
    expect(parseRoute("#/people/signins")).toEqual({
      view: "people",
      person: "signins",
      audit: null,
    });
    expect(parseRoute("#/people/a/b")).toEqual({ view: "people", person: "a/b", audit: null });
  });

  it("writes a tab only for an open person on the People page", async () => {
    const { parseRoute, routeHash } = await import("../src/App.tsx");
    expect(routeHash({ view: "people", person: null, audit: null, tab: "signins" })).toBe(
      "#/people",
    );
    expect(routeHash({ view: "audits", person: personId, audit: "x", tab: "signins" })).toBe(
      "#/audits/x",
    );
    expect(routeHash({ view: "overview", person: personId, audit: null, tab: "changes" })).toBe(
      "#/overview",
    );
    expect(parseRoute("#/people//signins")).toEqual({ view: "people", person: null, audit: null });
    expect(parseRoute("#/audits/accounts.inactive/signins")).toEqual({
      view: "audits",
      person: null,
      audit: "accounts.inactive/signins",
    });
  });

  it("names the tabs the way the labels do", () => {
    expect(parsePersonTab("signins")).toBe("signins");
    expect(parsePersonTab("Sign-ins")).toBeNull();
    expect(parsePersonTab(undefined)).toBeNull();
    expect(personTabs).toEqual(["overview", "signins", "licenses", "changes"]);
  });
});

describe("the tab list keys", () => {
  it("steps right and left and wraps, with Home and End", () => {
    expect(tabKeyTarget("ArrowRight", "overview")).toBe("signins");
    expect(tabKeyTarget("ArrowRight", "changes")).toBe("overview");
    expect(tabKeyTarget("ArrowLeft", "overview")).toBe("changes");
    expect(tabKeyTarget("ArrowLeft", "licenses")).toBe("signins");
    expect(tabKeyTarget("Home", "licenses")).toBe("overview");
    expect(tabKeyTarget("End", "overview")).toBe("changes");
    expect(tabKeyTarget("ArrowDown", "overview")).toBeNull();
    expect(tabKeyTarget("Tab", "overview")).toBeNull();
  });

  it("selects the tab the key points to and moves the focus there", () => {
    const log: string[] = [];
    const controls = {
      select: (tab: PersonTab) => log.push(`select ${tab}`),
      focusTab: (tab: PersonTab) => log.push(`focus ${tab}`),
    };
    expect(handleTabKey("ArrowRight", "signins", controls)).toBe(true);
    expect(handleTabKey("Enter", "signins", controls)).toBe(false);
    expect(log).toEqual(["select licenses", "focus licenses"]);
  });
});

const noCounts: Record<PersonTab, TabCount | null> = {
  overview: null,
  signins: null,
  licenses: null,
  changes: null,
};

describe("the tab list markup", () => {
  const list = (
    selected: PersonTab,
    onSelect: (tab: PersonTab) => void = noop,
    counts = noCounts,
  ) => hostTree(createElement(PersonTabList, { idBase: "x", selected, counts, onSelect }));

  it("is a tablist of tabs with one selected and the rest out of the tab order", () => {
    const tree = list("licenses");
    const tablist = findAll(tree, (host) => host.props.role === "tablist");
    expect(tablist).toHaveLength(1);
    const tabs = findAll(tree, (host) => host.props.role === "tab");
    expect(tabs.map((tab) => tab.props["aria-selected"])).toEqual([false, false, true, false]);
    expect(tabs.map((tab) => tab.props.tabIndex)).toEqual([-1, -1, 0, -1]);
    expect(tabs.map((tab) => textOf(tab))).toEqual(["Overview", "Sign-ins", "Licenses", "Changes"]);
  });

  it("points each tab at its own panel, and the panel back at the tab", () => {
    const tabs = findAll(list("overview"), (host) => host.props.role === "tab");
    for (const [index, tab] of personTabs.entries()) {
      expect(tabs[index]?.props["aria-controls"]).toBe(panelId("x", tab));
      const panel = hostTree(
        createElement(PersonTabPanel, { idBase: "x", tab, selected: "overview" }, "body"),
      )[0];
      expect(typeof panel === "object" && panel.props.id).toBe(panelId("x", tab));
      expect(typeof panel === "object" && panel.props["aria-labelledby"]).toBe(
        tabs[index]?.props.id,
      );
      expect(typeof panel === "object" && panel.props.role).toBe("tabpanel");
    }
  });

  it("selects a tab on click", () => {
    const chosen: PersonTab[] = [];
    const tabs = findAll(
      list("overview", (tab) => chosen.push(tab)),
      (h) => h.props.role === "tab",
    );
    fire(tabs[2], "onClick");
    expect(chosen).toEqual(["licenses"]);
  });

  it("selects the next tab on an arrow key and stops the browser using the key", () => {
    const chosen: PersonTab[] = [];
    const tablist = findAll(
      list("signins", (tab) => chosen.push(tab)),
      (h) => h.props.role === "tablist",
    )[0];
    const onKeyDown = tablist?.props.onKeyDown as (event: {
      key: string;
      preventDefault: () => void;
    }) => void;
    let prevented = 0;
    onKeyDown({ key: "ArrowRight", preventDefault: () => (prevented += 1) });
    onKeyDown({ key: "End", preventDefault: () => (prevented += 1) });
    onKeyDown({ key: "Tab", preventDefault: () => (prevented += 1) });
    expect(chosen).toEqual(["licenses", "changes"]);
    expect(prevented).toBe(2);
  });

  it("shows a count only where there is one", () => {
    const tabs = findAll(
      list("overview", noop, {
        overview: { value: 2 },
        signins: { value: 1234 },
        licenses: { value: 0 },
        changes: { value: 20, more: true },
      }),
      (host) => host.props.role === "tab",
    );
    expect(tabs.map((tab) => textOf(tab))).toEqual([
      "Overview2",
      "Sign-ins1,234",
      "Licenses0",
      "Changes20+",
    ]);
    expect(
      findAll(list("overview"), (host) => host.props.className === "person-tab__count"),
    ).toEqual([]);
  });

  it("hides the panels not showing, and shows the selected one", () => {
    const hidden = personTabs.map((tab) => {
      const panel = hostTree(
        createElement(PersonTabPanel, { idBase: "x", tab, selected: "changes" }, "body"),
      )[0];
      return typeof panel === "object" ? panel.props.hidden : null;
    });
    expect(hidden).toEqual([true, true, true, false]);
  });
});

describe("the counts on the tabs", () => {
  it("count open findings, stored sign-ins, licenses and loaded changes", () => {
    expect(
      personTabCounts({
        openFindings: 3,
        signIns: 42,
        licenses: 2,
        changes: { value: 20, more: true },
      }),
    ).toEqual({
      overview: { value: 3 },
      signins: { value: 42 },
      licenses: { value: 2 },
      changes: { value: 20, more: true },
    });
  });

  it("show no number for what has not loaded, and none on the Overview for no open findings", () => {
    expect(
      personTabCounts({ openFindings: null, signIns: null, licenses: null, changes: null }),
    ).toEqual(noCounts);
    expect(personTabCounts({ openFindings: 0, signIns: 0, licenses: 0, changes: null })).toEqual({
      overview: null,
      signins: { value: 0 },
      licenses: { value: 0 },
      changes: null,
    });
  });
});

// ---------------------------------------------------------------------------------------------
// The whole page
// ---------------------------------------------------------------------------------------------

const finding = (id: string, severity: Severity, state: FindingState = "open"): Finding =>
  findingSchema.parse({
    id,
    module: "signins-audit",
    ruleId: "signins.unfamiliar_signin",
    ruleVersion: 1,
    purpose: "security",
    severity,
    score: 400,
    subject: { kind: "person", id: personId, label: "Avery Stone" },
    title: `Finding ${id.slice(0, 2)}`,
    summary: "A sign-in came from a place not seen before.",
    evidence: [],
    limitations: ["Based on stored sign-ins only."],
    windowStart: stamp,
    windowEnd: stamp,
    state,
    occurrences: 1,
    firstSeen: stamp,
    lastSeen: stamp,
    stateChangedAt: null,
  });

const span = { count: 40, successes: 38, failures: 2, firstSeen: stamp, lastSeen: stamp };
const history = personHistorySchema.parse({
  userId: personId,
  userPrincipalName: "avery.stone@example.invalid",
  displayName: "Avery Stone",
  totals: {
    signIns: 42,
    successes: 40,
    failures: 2,
    blocked: 0,
    firstSeen: stamp,
    lastSeen: stamp,
    lastSuccess: stamp,
  },
  places: [{ country: "CA", state: "Saskatchewan", city: "Moose Jaw", ...span }],
  addresses: [{ ip: "192.0.2.10", country: "CA", state: null, city: null, ...span }],
  devices: [
    {
      deviceId: "device-1",
      operatingSystem: "Windows",
      browser: "Edge",
      compliant: true,
      managed: true,
      trustType: "AzureAd",
      ...span,
    },
  ],
  apps: [{ app: "Office 365", ...span }],
  baseline: [],
});

const plan = (name: string) => ({
  servicePlanId: `plan-${name}`,
  servicePlanName: name,
  provisioningStatus: "Success",
});
const directory = personDirectorySchema.parse({
  state: "available",
  profile: {
    id: personId,
    displayName: "Avery Stone",
    userPrincipalName: "avery.stone@example.invalid",
    mail: "avery.stone@example.invalid",
    jobTitle: "Support analyst",
    department: "Operations",
    companyName: null,
    officeLocation: null,
    employeeId: null,
    userType: "Member",
    accountEnabled: true,
    createdDateTime: stamp,
    lastPasswordChangeDateTime: stamp,
    onPremisesSyncEnabled: true,
    usageLocation: "CA",
    city: null,
    state: null,
    country: null,
    mobilePhone: null,
    businessPhones: [],
  },
  manager: { state: "available", data: null },
  signInActivity: {
    state: "available",
    lastSignIn: stamp,
    lastNonInteractiveSignIn: stamp,
    lastSuccessfulSignIn: stamp,
  },
  licenses: {
    state: "available",
    data: [
      { id: "l1", skuId: "s1", skuPartNumber: "ENTERPRISEPACK", servicePlans: [plan("EXCHANGE")] },
      { id: "l2", skuId: "s2", skuPartNumber: "POWER_BI_PRO", servicePlans: [plan("BI")] },
      { id: "l3", skuId: "s3", skuPartNumber: "FLOW_FREE", servicePlans: [plan("FLOW")] },
    ],
  },
});

const signins = [
  personSigninSchema.parse({
    id: "s-1",
    createdAt: stamp,
    result: "success",
    errorCode: 0,
    failureReason: null,
    app: "Office 365",
    resource: null,
    clientApp: "Browser",
    ip: "192.0.2.10",
    city: "Moose Jaw",
    state: "Saskatchewan",
    country: "CA",
    latitude: null,
    longitude: null,
    deviceId: "device-1",
    operatingSystem: "Windows",
    browser: "Edge",
    compliant: true,
    managed: true,
    trustType: "AzureAd",
    conditionalAccess: "success",
    riskLevel: "none",
    riskState: null,
    riskEventTypes: [],
    interactive: true,
  }),
  personSigninSchema.parse({
    id: "s-2",
    createdAt: "2026-10-19T09:00:00.000Z",
    result: "failure",
    errorCode: 50126,
    failureReason: "Invalid username or password.",
    app: null,
    resource: null,
    clientApp: null,
    ip: null,
    city: null,
    state: null,
    country: null,
    latitude: null,
    longitude: null,
    deviceId: null,
    operatingSystem: null,
    browser: null,
    compliant: null,
    managed: null,
    trustType: null,
    conditionalAccess: null,
    riskLevel: null,
    riskState: null,
    riskEventTypes: [],
    interactive: null,
  }),
];

const findings = [
  finding("11111111-1111-4111-8111-111111111111", "critical"),
  finding("22222222-2222-4222-8222-222222222222", "medium"),
  finding("33333333-3333-4333-8333-333333333333", "high", "resolved"),
];

function viewProps(overrides: Partial<PersonProfileViewProps> = {}): PersonProfileViewProps {
  return {
    session: session(),
    userId: personId,
    titleId: "title",
    tabsId: "x",
    tab: "overview",
    record: { history, findings },
    directory,
    signins,
    next: null,
    filter: "all",
    error: "",
    busy: null,
    changes: null,
    canary: canaryState(),
    onBack: noop,
    onTab: noop,
    onFilter: noop,
    onFindingState: noop,
    onMoreSignins: noop,
    onChanges: noop,
    ...overrides,
  };
}

const page = (overrides: Partial<PersonProfileViewProps> = {}) =>
  renderToStaticMarkup(createElement(PersonProfileView, viewProps(overrides)));

// The markup of each panel, by tab.
function panels(html: string): Record<string, string> {
  const found: Record<string, string> = {};
  const starts = [...html.matchAll(/<div id="x-panel-(\w+)"[^>]*>/g)];
  for (const [index, match] of starts.entries())
    found[match[1] ?? ""] = html.slice(
      (match.index ?? 0) + match[0].length,
      starts[index + 1]?.index ?? html.length,
    );
  return found;
}

// The component types under each panel, found without drawing them.
function panelComponents(
  overrides: Partial<PersonProfileViewProps> = {},
): Record<string, unknown[]> {
  const found: Record<string, unknown[]> = {};
  const walk = (value: ReactNode, into: unknown[]) => {
    if (Array.isArray(value)) {
      for (const item of value) walk(item, into);
      return;
    }
    if (!isValidElement(value)) return;
    into.push(value.type);
    walk((value.props as { children?: ReactNode }).children, into);
  };
  const visit = (value: ReactNode) => {
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    if (!isValidElement(value)) return;
    const props = value.props as { tab?: PersonTab; children?: ReactNode };
    if (value.type === PersonTabPanel && props.tab) {
      found[props.tab] = [];
      walk(props.children, found[props.tab] as unknown[]);
      return;
    }
    visit(props.children);
  };
  visit(PersonProfileView(viewProps(overrides)));
  return found;
}

describe("a person's page", () => {
  it("keeps every section it had, each under one tab", () => {
    const shown = panels(page());
    const where = (name: string) =>
      Object.entries(shown)
        .filter(
          ([, html]) => html.includes(`<h3>${name}</h3>`) || html.includes(`aria-label="${name}"`),
        )
        .map(([tab]) => tab);
    expect(where("Directory")).toEqual(["overview"]);
    expect(where("Stored sign-ins")).toEqual(["overview"]);
    expect(where("Findings")).toEqual(["overview"]);
    for (const heading of ["Places", "Apps", "IP addresses", "Devices", "Every sign-in"])
      expect(where(heading), heading).toEqual(["signins"]);
    expect(where("Licenses")).toEqual(["licenses"]);
    expect(where("Directory changes")).toEqual(["changes"]);
  });

  it("keeps the reports from other tools on the Changes tab, next to directory changes", () => {
    const components = panelComponents();
    expect(components.changes).toContain(PersonChanges);
    expect(components.changes).toContain(PersonReceipts);
    for (const tab of ["overview", "signins", "licenses"])
      expect(components[tab], tab).not.toContain(PersonChanges);
  });

  it("mounts the changes and reports on every tab, so switching tabs asks for nothing", () => {
    // Both fetch when they mount. They sit in the page whichever tab shows, hidden when not
    // selected, so each makes its request once, when the page opens.
    for (const tab of personTabs) {
      expect(panels(page({ tab })).changes, tab).toContain("Directory changes");
      const mounted = panelComponents({ tab }).changes;
      expect(mounted, tab).toContain(PersonChanges);
      expect(mounted, tab).toContain(PersonReceipts);
    }
  });

  it("shows exactly the selected tab's panel", () => {
    for (const tab of personTabs) {
      const html = page({ tab });
      const open = [...html.matchAll(/<div id="x-panel-(\w+)"([^>]*)>/g)]
        .filter((match) => !(match[2] ?? "").includes("hidden"))
        .map((match) => match[1]);
      expect(open, tab).toEqual([tab]);
      expect(html).toContain(
        `id="x-tab-${tab}" class="person-tab" type="button" role="tab" aria-selected="true"`,
      );
    }
  });

  it("shows the counts from what has loaded", () => {
    const html = page({ changes: { value: 20, more: true } });
    expect(html).toContain('>Overview<span class="person-tab__count">2</span>');
    expect(html).toContain('>Sign-ins<span class="person-tab__count">42</span>');
    expect(html).toContain('>Licenses<span class="person-tab__count">3</span>');
    expect(html).toContain('>Changes<span class="person-tab__count">20+</span>');
  });

  it("shows no counts while the page is loading", () => {
    const html = page({ record: null, directory: null, signins: null });
    expect(html).not.toContain("person-tab__count");
  });

  it("counts a person with no stored sign-ins as zero once the record has loaded", () => {
    const html = page({ record: { history: null, findings: [] } });
    expect(html).toContain('>Sign-ins<span class="person-tab__count">0</span>');
    expect(html).not.toContain('>Overview<span class="person-tab__count">');
  });

  it("puts the chips, the canary included, in the header", () => {
    const html = page({ canary: canaryState({ active: true }) });
    const head = html.slice(0, html.indexOf("person-tabs-strip"));
    expect(head).toContain(">Enabled<");
    expect(head).toContain(">Member<");
    expect(head).toContain(">Synced from on-premises<");
    expect(head).toContain(">Canary<");
    expect(words(head)).toContain("Open findings 1 critical 1 medium");
    expect(head).toContain("Stop treating as canary");
    expect(head).toContain('href="https://haveibeenpwned.com/account/avery.stone@example.invalid"');
    expect(head).toContain(`userId/${personId}`);
  });

  it("shows the canary question under the header while one is waiting", () => {
    const asking = page({ canary: canaryState({ confirming: true }) });
    expect(words(asking)).toContain("Mark Avery Stone as a canary account?");
    expect(page()).not.toContain("Mark Avery Stone as a canary account?");
  });

  it("shows dates in the tables as time ago, with the exact time on hover", () => {
    const shown = panels(page());
    const dates = shown.signins?.match(/<time [^>]*>/g) ?? [];
    // Places, apps, two for the address, devices, and each of the two sign-ins.
    expect(dates).toHaveLength(7);
    for (const tag of dates) {
      expect(tag).toMatch(/dateTime="\d{4}-\d\d-\d\dT[^"]+"|datetime="\d{4}-\d\d-\d\dT[^"]+"/);
      expect(tag).toMatch(/title="[^"]+"/);
    }
  });

  it("opens the tab the address names, and the Overview without one", () => {
    const profile = (tab?: PersonTab) =>
      renderToStaticMarkup(
        createElement(People, {
          session: session(),
          selected: personId,
          onSelect: noop,
          ...(tab ? { tab } : {}),
        }),
      );
    expect(profile()).toMatch(/aria-selected="true"[^>]*tabindex="0">Overview</);
    expect(profile("signins")).toMatch(/aria-selected="true"[^>]*tabindex="0">Sign-ins</);
    expect(profile("changes")).toMatch(/aria-selected="true"[^>]*tabindex="0">Changes</);
  });

  it("reads like a colleague wrote the header, the menu and the tabs", () => {
    // The panels hold finding cards and other shared components, which are not this page's copy.
    const html = page({ canary: canaryState({ confirming: true }) });
    const top = html.slice(0, html.indexOf('<div id="x-panel-overview"'));
    expect(words(top)).toContain("Actions");
    expect(words(top)).not.toMatch(/[—–·!]/);
  });

  it("has no link or button inside another", () => {
    expect(nestedInteractive(page({ canary: canaryState({ confirming: true }) }))).toBe(0);
  });
});

// ---------------------------------------------------------------------------------------------
// The stylesheet
// ---------------------------------------------------------------------------------------------

// The rules this page depends on are plain CSS, so they are checked as text: each is easy to lose
// in an edit, and nothing else would notice until someone opened the page on a phone.
const css = readFileSync(new URL("../src/styles.css", import.meta.url), "utf8");

function rule(selector: string): string {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = css.match(new RegExp(`(?:^|\\n)${escaped}\\s*\\{([^}]*)\\}`));
  if (!match) throw new Error(`No rule for ${selector}`);
  return match[1] ?? "";
}

// The text inside every block that opens with the given at-rule header.
function blocks(header: string): string[] {
  const found: string[] = [];
  let from = css.indexOf(header);
  while (from !== -1) {
    let depth = 0;
    let end = from + header.length - 1;
    for (let at = end; at < css.length; at += 1) {
      if (css[at] === "{") depth += 1;
      if (css[at] === "}") depth -= 1;
      if (depth === 0) {
        end = at;
        break;
      }
    }
    found.push(css.slice(from + header.length, end));
    from = css.indexOf(header, end);
  }
  return found;
}

describe("the person page's stylesheet", () => {
  it("hides a panel that is not showing, even though the panel sets its own display", () => {
    expect(rule(".person-panel")).toContain("display: grid;");
    expect(rule(".person-panel[hidden]")).toContain("display: none;");
  });

  it("lets a wide table scroll inside its panel instead of widening the page", () => {
    expect(rule(".person-panel")).toContain("grid-template-columns: minmax(0, 1fr);");
    expect(rule(".person-panel")).toContain("min-width: 0;");
  });

  it("scrolls the tab strip sideways on its own, and keeps each tab on one line", () => {
    expect(rule(".person-tabs-strip")).toContain("overflow-x: auto;");
    expect(rule(".person-tab")).toContain("flex: none;");
    expect(rule(".person-tab")).toContain("white-space: nowrap;");
  });

  it("wraps the header on a narrow screen and lets its text shrink and break", () => {
    expect(rule(".person-header")).toContain("flex-wrap: wrap;");
    expect(rule(".person-header__main")).toContain("min-width: 0;");
    expect(rule(".person-header__text")).toContain("min-width: 0;");
    expect(rule(".person-header__principal .copy-line__text")).toContain("white-space: normal;");
    expect(rule(".person-header__principal .copy-line__text")).toContain(
      "overflow-wrap: anywhere;",
    );
  });

  it("opens the menu from the right edge, and from the left on a phone", () => {
    expect(rule(".actions-menu")).toContain("position: relative;");
    expect(rule(".actions-menu__list")).toContain("position: absolute;");
    expect(rule(".actions-menu__list")).toContain("right: 0;");
    expect(rule(".actions-menu__list")).toContain("max-width: min(360px, calc(100vw - 32px));");
    const phone = blocks("@media (max-width: 560px) {").find((block) =>
      block.includes(".actions-menu__list"),
    );
    expect(phone).toMatch(/\.actions-menu__list \{[^}]*left: 0;/);
  });

  it("draws the avatar as a rounded square, not a circle", () => {
    const avatar = rule(".person-avatar");
    expect(avatar).toMatch(/border-radius: \d+px;/);
    expect(avatar).not.toContain("50%");
  });

  it("gets its tabular figures from the shared rule on the root, not its own", () => {
    expect(rule(":root")).toContain("font-variant-numeric: tabular-nums;");
    expect(rule(".person-tab__count")).not.toContain("font-variant-numeric");
  });

  it("keeps the focus ring inside a tab, where the scrolling strip would clip it", () => {
    expect(rule(".person-tab:focus-visible")).toContain("inset");
  });
});

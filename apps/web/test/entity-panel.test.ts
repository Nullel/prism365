import { readFileSync } from "node:fs";
import {
  type AuditRunDetail,
  auditDefinitionSchema,
  decodeOwnerList,
  type Entity,
  type EntityKind,
  type EntitySearch,
  encodeOwnerEntry,
  entitySchema,
  type PersonSummary,
} from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";
import { entityOfCell, ResultTable } from "../src/Audits.tsx";
import { entityOfTarget } from "../src/Changes.tsx";
import { EntityPanelView, type EntityPanelViewProps, panelOrigin } from "../src/EntityPanel.tsx";
import {
  type EntityTarget,
  handlePanelKey,
  type PanelKeyEvent,
  panelReducer,
} from "../src/entityPanelState.ts";
import {
  EntityLink,
  type Navigation,
  NavigationContext,
  OriginScope,
  type PersonOrigin,
  originTrail,
  useOpenPerson,
} from "../src/navigation.tsx";
import { PersonSummaryView } from "../src/PersonSummary.tsx";
import { summaryItems } from "../src/personSummaryLines.ts";
import {
  entityGroups,
  entityItems,
  entityKindOfGroup,
  entityQuery,
  flatten,
  focusKeeper,
  type SearchSources,
  searchResults,
  sourceNotes,
} from "../src/QuickSearch.tsx";
import { nestedInteractive } from "./markup.ts";

// The panel for a group, role, policy, device or app, and the links that open it. Invented names
// and ids, .invalid addresses only.

const id = (n: number) => `bbbbbbbb-0000-4000-8000-${n.toString(16).padStart(12, "0")}`;
const GROUP = id(10);
const ROLE = "62e90394-69f5-4237-9190-012177145e10";
const POLICY = id(30);
const DEVICE = id(40);
const APP = id(50);
const ANA = id(1);
const T = "2026-10-20T10:05:00.000Z";

const noop = () => {};
const nav = (calls: unknown[][] = []): Navigation => ({
  openPerson: (...args) => void calls.push(["person", ...args]),
  openAudit: (...args) => void calls.push(["audit", ...args]),
  openView: (...args) => void calls.push(["view", ...args]),
  openEntity: (...args) => void calls.push(["entity", ...args]),
  origin: null,
});

const base = { finishedAt: T, complete: true };
const section = (over: Record<string, unknown>) => ({
  id: "details",
  title: "Details",
  auditId: "groups.teams",
  auditName: "Groups and Teams",
  ...over,
});
const row = (over: Record<string, unknown> = {}) => ({
  key: "r",
  subject: null,
  tags: [],
  cells: [],
  lists: [],
  risks: [],
  ...over,
});

const entities: Record<EntityKind, Entity> = {
  group: entitySchema.parse({
    kind: "group",
    id: GROUP,
    name: "Plant team",
    tags: ["Team", "One owner"],
    notes: ["Members are not read, so this cannot say who is in the group."],
    sections: [
      section({
        status: "ok",
        id: "group",
        title: "The group",
        ...base,
        more: false,
        note: null,
        rows: [
          row({
            key: GROUP,
            cells: [
              { key: "status", label: "Status", kind: "enum", value: "One owner" },
              { key: "created", label: "Created", kind: "date", value: "2025-01-05T00:00:00.000Z" },
            ],
            lists: [
              {
                label: "Owners",
                items: [
                  { kind: "person", id: ANA, name: "Ana Ames", principal: "ana@example.invalid" },
                  { kind: "app", id: APP, name: "Plant robot", principal: null },
                  { kind: "other", id: id(3), name: null, principal: null },
                ],
                more: 2,
              },
            ],
          }),
        ],
      }),
      section({
        status: "ok",
        id: "policies",
        title: "Conditional Access policies",
        auditId: "identity.conditional-access",
        auditName: "Conditional Access",
        ...base,
        complete: false,
        more: false,
        note: null,
        rows: [
          row({
            key: POLICY,
            subject: { kind: "policy", id: POLICY, name: "Require MFA", principal: null },
            tags: ["Includes it"],
          }),
        ],
      }),
      section({
        status: "no_run",
        id: "roles",
        title: "Roles it gives",
        auditId: "directory.privileged-roles",
        auditName: "Privileged roles",
      }),
    ],
  }),
  role: entitySchema.parse({
    kind: "role",
    id: ROLE,
    name: "Global Administrator",
    tags: ["Privileged role", "2 assignments"],
    notes: [],
    sections: [
      section({
        status: "ok",
        id: "holders",
        title: "Who holds it",
        auditId: "directory.privileged-roles",
        auditName: "Privileged roles",
        ...base,
        more: true,
        note: null,
        rows: [
          row({
            key: "a",
            subject: { kind: "person", id: ANA, name: "Ana Ames", principal: null },
            risks: [
              { id: "admin-unprotected", label: "Administrator with no MFA", severity: "critical" },
            ],
          }),
          row({ key: "b", subject: { kind: "group", id: GROUP, name: null, principal: null } }),
        ],
      }),
      section({
        status: "failed",
        id: "risks",
        title: "Combined risks among the people who hold it",
        auditId: "identity.combined-risks",
        auditName: "Combined risks",
        message: "The stored result could not be read just now. Try again.",
      }),
    ],
  }),
  policy: entitySchema.parse({
    kind: "policy",
    id: POLICY,
    name: "Require MFA",
    tags: ["On"],
    notes: [],
    sections: [
      section({
        status: "ok",
        id: "policy",
        title: "The policy",
        auditId: "identity.conditional-access",
        auditName: "Conditional Access",
        ...base,
        more: false,
        note: null,
        rows: [
          row({
            key: POLICY,
            lists: [
              {
                label: "Includes these groups",
                items: [{ kind: "group", id: GROUP, name: null, principal: null }],
                more: 0,
              },
            ],
          }),
        ],
      }),
    ],
  }),
  device: entitySchema.parse({
    kind: "device",
    id: DEVICE,
    name: "Plant laptop",
    tags: ["Windows"],
    notes: [],
    sections: [
      section({
        status: "not_in_run",
        id: "device",
        title: "The device",
        auditId: "devices.managed",
        auditName: "Managed devices",
        ...base,
      }),
    ],
  }),
  app: entitySchema.parse({
    kind: "app",
    id: APP,
    name: null,
    tags: [],
    notes: [],
    sections: [
      section({
        status: "ok",
        id: "app",
        title: "The app",
        auditId: "apps.consents",
        auditName: "App consents",
        ...base,
        more: false,
        note: "Nothing more to say.",
        rows: [],
      }),
    ],
  }),
};

const props = (over: Partial<EntityPanelViewProps> = {}): EntityPanelViewProps => ({
  target: { kind: "group", id: GROUP },
  entity: entities.group,
  error: "",
  canGoBack: false,
  titleId: "title",
  onBack: noop,
  onClose: noop,
  ...over,
});
const draw = (over: Partial<EntityPanelViewProps> = {}, navigation: Navigation = nav()) =>
  renderToStaticMarkup(
    createElement(
      NavigationContext.Provider,
      { value: navigation },
      createElement(EntityPanelView, props(over)),
    ),
  );
const words = (html: string) =>
  html
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
const peopleButtons = (html: string) => [...html.matchAll(/<button[^>]*people-name[^>]*>/g)].length;

describe("the panel", () => {
  it("is a labelled dialog with the kind as a small label, the name and the key facts as tags", () => {
    const html = draw();
    expect(html).toContain('role="dialog"');
    expect(html).toContain('aria-modal="true"');
    expect(html).toContain('aria-labelledby="title"');
    expect(html).toMatch(
      /<span class="eyebrow">Group<\/span><h2 id="title" tabindex="-1">Plant team<\/h2>/,
    );
    expect(html).toContain('<span class="person-tag">Team</span>');
    expect(html).toContain('<span class="person-tag">One owner</span>');
    expect(html).toContain('aria-label="Close"');
  });

  it("has one section per source, each saying which audit it is from and when", () => {
    const html = draw();
    expect(html).toContain(
      '<a href="#/audits/groups.teams">Groups and Teams</a> audit, <time dateTime="2026-10-20T10:05:00.000Z"',
    );
    expect((html.match(/<section class="panel person-access entity-section"/g) ?? []).length).toBe(
      3,
    );
    // A run that could not check everything says so, and a source that never ran says how to run it.
    expect(words(html)).toContain("That run could not check everything");
    expect(words(html)).toContain("Run the Privileged roles audit to see this here");
  });

  it("names a group's owners as links: accounts open a person, apps open an app, others are text", () => {
    const html = draw();
    expect(words(html)).toContain("Ana Ames");
    expect(words(html)).toContain("Plant robot");
    expect(words(html)).toContain("Unknown owner");
    expect(words(html)).toContain("and 2 more");
    // Ana, the robot and the policy are links; the unnamed owner is text with a copy button.
    expect(peopleButtons(html)).toBe(3);
    expect(html).toContain(`aria-label="Copy the ID of this owner"`);
    expect(words(html)).toContain("Members are not read");
  });

  it("shows a role's holders with their combined risks, a section that failed, and a count that is cut", () => {
    const html = draw({ target: { kind: "role", id: ROLE }, entity: entities.role });
    expect(html).toContain('<span class="eyebrow">Role</span>');
    expect(html).toContain('<span class="sev sev--critical">');
    expect(words(html)).toContain("Administrator with no MFA");
    expect(words(html)).toContain("The stored result could not be read just now. Try again.");
    expect(words(html)).toContain("The audit holds more.");
    // A group holder nobody named is a link that reads "Unknown group", with its id behind a copy button.
    expect(words(html)).toContain("Unknown group");
    expect(html).toContain("Copy the ID of this group");
  });

  it("shows a policy's lists, a device that is not in the run, and an app no audit names", () => {
    const policy = draw({ target: { kind: "policy", id: POLICY }, entity: entities.policy });
    expect(words(policy)).toContain("Includes these groups");
    expect(words(policy)).toContain("Unknown group");
    const device = draw({ target: { kind: "device", id: DEVICE }, entity: entities.device });
    expect(words(device)).toContain("This device is not in the latest Managed devices run.");
    const app = draw({ target: { kind: "app", id: APP }, entity: entities.app });
    expect(words(app)).toContain("Unknown app");
    expect(app).toContain("Copy the ID of this app");
    expect(words(app)).toContain("Nothing more to say.");
  });

  it("shows a skeleton while it loads and one plain line when it fails", () => {
    const loading = draw({ entity: null });
    expect(loading).toContain("skeleton__row");
    expect(loading).toContain("Reading what is stored about this group");
    expect(loading).not.toContain("Loading…");
    expect(loading).toContain('<h2 id="title" tabindex="-1">Group</h2>');
    const failed = draw({ entity: null, error: "That is not available." });
    expect(failed).toContain('role="alert"');
    expect(words(failed)).toContain("That is not available.");
    expect(failed).not.toContain("skeleton__row");
  });

  it("offers Back only when there is a step to go back to", () => {
    expect(draw()).not.toContain(">Back<");
    expect(draw({ canGoBack: true })).toContain(">Back<");
  });

  it("holds no link or button inside another", () => {
    for (const kind of Object.keys(entities) as EntityKind[])
      expect(
        nestedInteractive(
          draw({ target: { kind, id: entities[kind].id }, entity: entities[kind] }),
        ),
        kind,
      ).toBe(0);
  });
});

describe("the steps inside the panel", () => {
  const a: EntityTarget = { kind: "group", id: GROUP };
  const b: EntityTarget = { kind: "policy", id: POLICY };
  const c: EntityTarget = { kind: "role", id: ROLE };

  it("starts with the thing opened and adds a step for each thing opened from inside", () => {
    let stack = panelReducer([], { type: "open", target: a });
    expect(stack).toEqual([a]);
    stack = panelReducer(stack, { type: "open", target: b });
    stack = panelReducer(stack, { type: "open", target: c });
    expect(stack).toEqual([a, b, c]);
  });

  it("goes back one step at a time and stays on the first", () => {
    let stack = panelReducer([a, b, c], { type: "back" });
    expect(stack).toEqual([a, b]);
    stack = panelReducer(panelReducer(stack, { type: "back" }), { type: "back" });
    expect(stack).toEqual([a]);
  });

  it("adds nothing when the thing shown is opened again, whatever the case of its id", () => {
    const stack = panelReducer([a], { type: "open", target: { ...a, id: GROUP.toUpperCase() } });
    expect(stack).toEqual([a]);
  });

  it("starts over when something outside opens a thing, and closes empty", () => {
    expect(panelReducer([a, b], { type: "open", target: c, fresh: true })).toEqual([c]);
    expect(panelReducer([a, b, c], { type: "close" })).toEqual([]);
  });

  it("leaves the origin a person opened from the panel reports: the thing, by name", () => {
    const origin = panelOrigin({ kind: "group", id: GROUP }, entities.group);
    expect(origin).toEqual({ kind: "entity", entityKind: "group", id: GROUP, name: "Plant team" });
    expect(panelOrigin({ kind: "app", id: APP }, entities.app)).toMatchObject({
      name: "Unknown app",
    });
    expect(panelOrigin({ kind: "app", id: APP }, null)).toMatchObject({ name: "Unknown app" });
  });

  it("says where a person came from, and goes back to the thing", () => {
    const calls: unknown[][] = [];
    const origin: PersonOrigin = {
      kind: "entity",
      entityKind: "group",
      id: GROUP,
      name: "Plant team",
    };
    const trail = originTrail(origin, nav(calls));
    expect(trail.map((crumb) => crumb.label)).toEqual(["Group: Plant team"]);
    trail[0]?.open();
    expect(calls).toEqual([["entity", "group", GROUP]]);
  });

  it("closes the panel whenever a person's page opens, in the App that owns it", () => {
    const source = readFileSync(new URL("../src/App.tsx", import.meta.url), "utf8");
    expect(source).toMatch(
      /const openPerson = [\s\S]*?dispatchEntity\(\{ type: "close" \}\);[\s\S]*?selectPerson\(userId\)/,
    );
    // And every other way off the page: the address changing, an audit, another view.
    expect(source).toMatch(/const follow = [\s\S]*?dispatchEntity\(\{ type: "close" \}\)/);
    expect(source).toMatch(/const openAudit = [\s\S]*?dispatchEntity\(\{ type: "close" \}\)/);
    expect(source).toMatch(/const navigate = [\s\S]*?dispatchEntity\(\{ type: "close" \}\)/);
  });

  it("hands the thing's origin to a person opened from inside the panel", () => {
    const calls: unknown[][] = [];
    const origin = panelOrigin({ kind: "group", id: GROUP }, entities.group);
    // A person link opens people through this hook with whatever origin its scope holds.
    const Probe = () => {
      useOpenPerson()(ANA);
      return null;
    };
    renderToStaticMarkup(
      createElement(
        NavigationContext.Provider,
        { value: nav(calls) },
        createElement(OriginScope, { origin }, createElement(Probe)),
      ),
    );
    expect(calls).toEqual([["person", ANA, origin]]);
    // Outside any panel the same link has no origin of its own.
    calls.length = 0;
    renderToStaticMarkup(
      createElement(NavigationContext.Provider, { value: nav(calls) }, createElement(Probe)),
    );
    expect(calls).toEqual([["person", ANA, undefined]]);
  });
});

describe("keyboard and focus", () => {
  const event = (over: Partial<PanelKeyEvent> = {}): PanelKeyEvent & { prevented: boolean } => {
    const value = {
      key: "Escape",
      shiftKey: false,
      defaultPrevented: false,
      prevented: false,
      preventDefault() {
        value.prevented = true;
      },
      ...over,
    };
    return value;
  };
  const control = () => ({ focus: vi.fn() });

  it("closes on Escape and nothing else sees it", () => {
    const close = vi.fn();
    const pressed = event();
    handlePanelKey(pressed, { first: null, last: null, active: null, inside: true }, close);
    expect(close).toHaveBeenCalledTimes(1);
    expect(pressed.prevented).toBe(true);
  });

  it("does not close for an Escape that ends a word in an input method, or one already handled", () => {
    const close = vi.fn();
    handlePanelKey(
      event({ nativeEvent: { isComposing: true } }),
      { first: null, last: null, active: null, inside: true },
      close,
    );
    handlePanelKey(
      event({ defaultPrevented: true }),
      { first: null, last: null, active: null, inside: true },
      close,
    );
    expect(close).not.toHaveBeenCalled();
  });

  it("keeps Tab inside: from the last control to the first, and Shift+Tab the other way", () => {
    const first = control();
    const last = control();
    const forward = event({ key: "Tab" });
    handlePanelKey(forward, { first, last, active: last, inside: true }, vi.fn());
    expect(first.focus).toHaveBeenCalledTimes(1);
    expect(forward.prevented).toBe(true);
    const backward = event({ key: "Tab", shiftKey: true });
    handlePanelKey(backward, { first, last, active: first, inside: true }, vi.fn());
    expect(last.focus).toHaveBeenCalledTimes(1);
    // In the middle, Tab does its own thing.
    const middle = event({ key: "Tab" });
    handlePanelKey(middle, { first, last, active: {}, inside: true }, vi.fn());
    expect(middle.prevented).toBe(false);
  });

  it("brings focus in from outside the panel, either way, and closes on Escape from anywhere", () => {
    const first = control();
    const last = control();
    const forward = event({ key: "Tab" });
    handlePanelKey(forward, { first, last, active: null, inside: false }, vi.fn());
    expect(first.focus).toHaveBeenCalledTimes(1);
    expect(forward.prevented).toBe(true);
    const backward = event({ key: "Tab", shiftKey: true });
    handlePanelKey(backward, { first, last, active: null, inside: false }, vi.fn());
    expect(last.focus).toHaveBeenCalledTimes(1);
    const close = vi.fn();
    handlePanelKey(event(), { first, last, active: null, inside: false }, close);
    expect(close).toHaveBeenCalledTimes(1);
  });

  it("gives focus back to what had it when the panel opened, unless that left the page", () => {
    const opener = { focus: vi.fn(), isConnected: true };
    const keeper = focusKeeper(() => opener);
    keeper.remember();
    keeper.restore();
    expect(opener.focus).toHaveBeenCalledTimes(1);
    const gone = { focus: vi.fn(), isConnected: false };
    const other = focusKeeper(() => gone);
    other.remember();
    other.restore();
    expect(gone.focus).not.toHaveBeenCalled();
  });
});

describe("links that open a panel", () => {
  it("is a button that opens the thing, and says so when no audit named it", () => {
    const html = renderToStaticMarkup(
      createElement(
        "div",
        null,
        createElement(EntityLink, { kind: "group", id: GROUP, name: "Plant team" }),
        createElement(EntityLink, { kind: "device", id: DEVICE, name: null }),
        createElement(EntityLink, { kind: "role", id: null, name: "Custom role" }),
        createElement(EntityLink, { kind: "role", id: null, name: null }),
      ),
    );
    expect(html).toMatch(/<button class="people-name entity-link" type="button">/);
    expect(words(html)).toContain("Plant team");
    expect(words(html)).toContain("Unknown device");
    expect(html).toContain("Copy the ID of this device");
    // Without an id there is nothing to open: plain text.
    expect(peopleButtons(html)).toBe(2);
    expect(html).toContain("<span>Custom role</span>");
    expect(html).toContain('<span class="person-muted">Unknown role</span>');
  });
});

describe("tables and pages that link to a panel", () => {
  const definition = (idValue: string, columns: Record<string, unknown>[]) =>
    auditDefinitionSchema.parse({
      id: idValue,
      version: 1,
      title: "An audit",
      description: "Rows.",
      category: "Test",
      subject: { one: "row", many: "rows" },
      columns,
      views: [{ id: "all", label: "All", tags: [] }],
      defaultView: "all",
      defaultSort: { column: columns[0]?.key, direction: "asc" },
      permissions: [],
      optionalPermissions: [],
      requirements: [],
    });
  const table = (def: ReturnType<typeof definition>, rows: unknown[]) =>
    renderToStaticMarkup(
      createElement(ResultTable, {
        definition: def,
        run: {
          id: "66666666-6666-4666-8666-666666666666",
          auditId: def.id,
          auditVersion: 1,
          status: "succeeded",
          complete: true,
          startedAt: "2026-10-22T15:30:00.000Z",
          finishedAt: "2026-10-22T15:31:00.000Z",
          startedBy: { id: ANA, name: null },
          error: null,
          summary: [],
          notes: [],
          rowCount: rows.length,
          rows,
        } as unknown as AuditRunDetail,
        onOpenPerson: noop,
      }),
    );

  it("makes the name of each group, policy, app, role and device a link, when the row has its id", () => {
    const cases: [string, string, Record<string, unknown>, EntityKind][] = [
      ["groups.teams", "group", { key: GROUP }, "group"],
      ["identity.conditional-access", "policy", { key: POLICY }, "policy"],
      ["apps.consents", "app", { key: APP }, "app"],
      // A registration is keyed by its own object id; the panel opens by the enterprise application.
      ["apps.credentials", "app", { cells: { servicePrincipalId: APP } }, "app"],
      ["devices.managed", "device", { key: DEVICE }, "device"],
      ["directory.privileged-roles", "role", { key: "held|x", cells: { roleId: ROLE } }, "role"],
    ];
    for (const [auditId, column, over, kind] of cases) {
      const found = entityOfCell(auditId, column, {
        key: GROUP,
        personId: null,
        tags: [],
        cells: {},
        ...over,
      } as never);
      expect(found?.kind, auditId).toBe(kind);
      expect(
        entityOfCell(auditId, column, { key: "not-an-id", personId: null, tags: [], cells: {} }),
      ).toBeNull();
      expect(
        entityOfCell(auditId, "other", { key: GROUP, personId: null, tags: [], cells: {} }),
      ).toBeNull();
    }
    // A registration's own key is not the id the panel opens by: a row with no enterprise
    // application has nothing to open.
    const registration = (cells: Record<string, unknown>) =>
      entityOfCell("apps.credentials", "app", {
        key: GROUP,
        personId: null,
        tags: [],
        cells,
      } as never);
    expect(registration({ servicePrincipalId: APP })?.id).toBe(APP);
    expect(registration({ servicePrincipalId: null })).toBeNull();
    expect(registration({})).toBeNull();
    expect(
      entityOfCell("accounts.inactive", "group", {
        key: GROUP,
        personId: null,
        tags: [],
        cells: {},
      }),
    ).toBeNull();
  });

  it("draws the Groups and Teams name as a link and the owners as person links, users only", () => {
    const def = definition("groups.teams", [
      { key: "group", label: "Group", kind: "text" },
      { key: "ownerNames", label: "Owner names", kind: "list" },
      { key: "ownerList", label: "Owners", kind: "list", hidden: true },
    ]);
    const owners = [
      encodeOwnerEntry({ id: ANA, name: "Ana Ames", kind: "user" }),
      encodeOwnerEntry({ id: APP, name: "Ana robot", kind: "servicePrincipal" }),
      encodeOwnerEntry({ id: id(3), name: "Cai Cole", kind: "user" }),
    ];
    expect(decodeOwnerList(owners)).toHaveLength(3);
    const html = table(def, [
      {
        key: GROUP,
        personId: null,
        tags: [],
        cells: {
          group: "Plant team",
          ownerNames: ["Ana Ames", "Ana robot", "Cai Cole"],
          ownerList: owners,
        },
      },
    ]);
    // The group name, and the first two owners by name: Ana, an account, which opens her page,
    // and an app of the same family name, which does not.
    expect(peopleButtons(html)).toBe(2);
    // Each name keeps the comma that follows it, so no line starts with one.
    expect(html).toMatch(/Ana Ames<\/strong><\/button>,/);
    expect(html).toMatch(/Ana robot,/);
    expect(html).toContain('class="link-list"');
    expect(words(html)).toContain("+1");
    expect(html).toContain('title="Ana Ames\nAna robot\nCai Cole"');
    // A run from before the hidden list shows the names as text.
    const plain = table(def, [
      {
        key: GROUP,
        personId: null,
        tags: [],
        cells: { group: "Plant team", ownerNames: ["Ana Ames"], ownerList: null },
      },
    ]);
    expect(peopleButtons(plain)).toBe(1);
    expect(nestedInteractive(html)).toBe(0);
  });

  it("opens an app that holds a role, and leaves a holder the audit could not sort as it is", () => {
    const def = definition("directory.privileged-roles", [
      { key: "holder", label: "Holder", kind: "person" },
      { key: "role", label: "Role", kind: "enum" },
      { key: "roleId", label: "Role ID", kind: "text", hidden: true, internal: true },
    ]);
    const held = (key: string, type: string, holderId: string, personId: string | null = null) => ({
      key,
      personId,
      tags: [],
      cells: {
        holder: { name: `Holder ${key}`, principal: null },
        holderType: type,
        holderId,
        role: "Lab Reports Reader",
        roleId: ROLE,
      },
    });
    const html = table(def, [
      held("app", "Service principal", APP),
      held("group", "Group or other", GROUP),
      held("person", "User", ANA, ANA),
    ]);
    expect(html).toMatch(/people-name entity-link[^>]*><strong[^>]*>Holder app</);
    expect(html).not.toMatch(/entity-link[^>]*><strong[^>]*>Holder group</);
    expect(html).toContain("Holder group");
    expect(html).toMatch(/people-name"[^>]*><strong[^>]*>Holder person</);
  });

  it("keeps machine data out of the table and the Columns chooser, and shows the other hidden columns", () => {
    const def = definition("groups.teams", [
      { key: "group", label: "Group", kind: "text" },
      { key: "ownerIds", label: "Owner IDs", kind: "list", hidden: true },
      { key: "ownerList", label: "Owners with kinds", kind: "list", hidden: true, internal: true },
    ]);
    const html = table(def, [
      {
        key: GROUP,
        personId: null,
        tags: [],
        cells: { group: "Plant team", ownerIds: [ANA], ownerList: ['{"id":"x","kind":"user"}'] },
      },
    ]);
    expect(html).toContain("Owner IDs");
    expect(html).not.toContain("Owners with kinds");
    expect(html).not.toContain("&quot;kind&quot;");
  });

  it("keeps the owner of a device opening the owner, and the device name opening the device", () => {
    const def = definition("devices.managed", [
      { key: "device", label: "Device", kind: "text" },
      { key: "person", label: "Owner", kind: "person" },
    ]);
    const html = table(def, [
      {
        key: DEVICE,
        personId: ANA,
        tags: [],
        cells: { device: "Plant laptop", person: { name: "Ana Ames", principal: null } },
      },
    ]);
    expect(peopleButtons(html)).toBe(2);
    expect(html).toContain("Plant laptop");
    expect(html).toContain("Ana Ames");
  });

  it("links a change's target only when its id is the one the panel is keyed by", () => {
    const target = (type: string | null, targetId: string | null) => ({
      id: targetId,
      name: "x",
      type,
      principal: null,
      properties: [],
    });
    expect(entityOfTarget(target("Group", GROUP.toUpperCase()))).toEqual({
      kind: "group",
      id: GROUP,
    });
    expect(entityOfTarget(target("Policy", POLICY))).toEqual({ kind: "policy", id: POLICY });
    expect(entityOfTarget(target("ServicePrincipal", APP))).toEqual({ kind: "app", id: APP });
    // A role target is a directory role, a device target is an Entra device, an application is
    // an app registration: none has the id its panel is keyed by.
    for (const type of ["Role", "Device", "Application", "User", null])
      expect(entityOfTarget(target(type, GROUP))).toBeNull();
    expect(entityOfTarget(target("Group", "not-an-id"))).toBeNull();
    expect(entityOfTarget(target("Group", null))).toBeNull();
  });
});

describe("the person Summary", () => {
  const summary = {
    combinedRisks: { status: "no_run" },
    signinFindings: { status: "unavailable" },
    audits: [],
    owned: {
      status: "ok",
      finishedAt: T,
      complete: true,
      more: false,
      entries: [1, 2, 3, 4, 5].map((n) => ({
        id: id(100 + n),
        name: `Group ${n}`,
        kind: "Team",
        status: "OK",
        tags: [],
      })),
    },
    sharedMailboxes: { status: "no_import" },
    changes: { status: "unavailable" },
    handoff: { status: "ok", latest: null },
  } as unknown as PersonSummary;

  it("names the first three groups an account owns under its row, as links outside the row's own", () => {
    const lines = summaryItems(summary).items.find((item) => item.id === "owned-groups");
    expect(lines?.entities?.map((entity) => entity.name)).toEqual([
      "Group 1",
      "Group 2",
      "Group 3",
    ]);
    expect(lines?.entitiesMore).toBe(2);
    const html = renderToStaticMarkup(
      createElement(PersonSummaryView, { summary, error: "", userId: ANA, onJump: noop }),
    );
    expect(peopleButtons(html)).toBe(3);
    expect(words(html)).toContain("and 2 more");
    expect(nestedInteractive(html)).toBe(0);
  });
});

describe("Quick search", () => {
  const found: EntitySearch = {
    results: [
      { kind: "group", id: GROUP, name: "Plant team", detail: "Team" },
      { kind: "device", id: DEVICE, name: "Plant laptop", detail: "Windows" },
      { kind: "role", id: ROLE, name: "Plant administrator", detail: null },
      { kind: "policy", id: POLICY, name: "Plant sign-in", detail: "On" },
      { kind: "app", id: APP, name: "Plant robot", detail: "Unverified" },
    ],
    more: false,
    partial: false,
  };
  const sources = (over: Partial<SearchSources> = {}): SearchSources => ({
    pages: [{ id: "people", label: "People", detail: null, keywords: [] }],
    audits: [{ id: "accounts.inactive", label: "Plant audit", detail: null, keywords: [] }],
    people: [{ id: ANA, label: "Plant Pell", detail: null, keywords: [] }],
    entities: entityItems(found),
    ...over,
  });

  it("lists groups, roles, policies, devices and apps under their own headings, after the people", () => {
    const groups = searchResults(sources(), "plant", { limit: 20 });
    expect(groups.map((group) => group.heading)).toEqual([
      "Audits",
      "People",
      "Groups",
      "Roles",
      "Policies",
      "Devices",
      "Apps",
    ]);
    expect(entityGroups.map((group) => entityKindOfGroup[group])).toEqual([
      "group",
      "role",
      "policy",
      "device",
      "app",
    ]);
    const flat = flatten(groups);
    expect(flat.find((item) => item.group === "devices")?.item).toMatchObject({
      id: DEVICE,
      label: "Plant laptop",
      detail: "Windows",
    });
  });

  it("puts them after the people even when one of them is the best match", () => {
    const better = searchResults(
      sources({
        entities: entityItems({
          ...found,
          results: [{ kind: "group", id: GROUP, name: "Plant", detail: null }],
        }),
      }),
      "plant",
      { limit: 20 },
    );
    expect(better.map((group) => group.id)).toEqual(["audits", "people", "groups"]);
  });

  it("shows nothing under them when the server found nothing, and ranks them like the rest", () => {
    const none = searchResults(sources({ entities: {} }), "plant", { limit: 20 });
    expect(none.map((group) => group.id)).toEqual(["audits", "people"]);
    const typed = searchResults(sources(), "plant robot", { limit: 20 });
    expect(typed.map((group) => group.id)).toEqual(["apps"]);
  });

  it("asks for at least two characters, trimmed and bounded", () => {
    expect(entityQuery("a")).toBeNull();
    expect(entityQuery("  a ")).toBeNull();
    expect(entityQuery(" ab ")).toBe("ab");
    expect(entityQuery("x".repeat(300))).toHaveLength(100);
  });

  it("never cuts a character in half, so the address can always be built", () => {
    // An emoji is two UTF-16 units; one that starts at the 100th unit does not fit.
    const text = `${"a".repeat(99)}😀tail`;
    const sent = entityQuery(text);
    expect(sent).toBe("a".repeat(99));
    expect(sent?.isWellFormed()).toBe(true);
    expect(() => encodeURIComponent(sent ?? "")).not.toThrow();
    expect(entityQuery(`${"a".repeat(98)}😀`)).toHaveLength(100);
    expect(entityQuery(text.slice(0, 100))?.isWellFormed()).toBe(true);
  });

  it("gives each kind the icon that goes with it, like the people", () => {
    const items = entityItems(found);
    for (const group of entityGroups)
      for (const item of items[group]) expect(item.icon, group).toBeDefined();
    const icons = entityGroups.map((group) => items[group][0]?.icon);
    expect(new Set(icons).size).toBe(entityGroups.length);
  });

  it("says quietly when there are more matches than were sent", () => {
    const idle = { state: "idle" } as const;
    const audits = { state: "ready", value: {} } as const;
    const ready = { state: "ready", value: { ...found, more: true } } as const;
    expect(sourceNotes("plant", true, audits, idle, ready)).toEqual(["More matches not shown."]);
  });

  it("says plainly when not everything could be searched", () => {
    const idle = { state: "idle" } as const;
    const ready = (partial: boolean) => ({ state: "ready", value: { ...found, partial } }) as const;
    expect(sourceNotes("plant", true, { state: "ready", value: {} }, idle, ready(false))).toEqual(
      [],
    );
    const partial = sourceNotes("plant", true, { state: "ready", value: {} }, idle, ready(true));
    const failed = sourceNotes("plant", true, { state: "ready", value: {} }, idle, {
      state: "failed",
    });
    expect(partial).toEqual([
      "Groups, roles, policies, devices and apps could not all be searched.",
    ]);
    expect(failed).toEqual(partial);
  });
});

import {
  type AuditRunDetail,
  auditDefinitionSchema,
  type ChangeEvent,
  type PersonDirectory,
  type PersonMailboxAccess,
  type PrismSession,
  personDirectorySchema,
  personMailboxAccessSchema,
  sessionSchema,
} from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";
import { viewOrigin } from "../src/App.tsx";
import { ResultTable } from "../src/Audits.tsx";
import { ChangeRow } from "../src/Changes.tsx";
import { ReceiptItemsTable } from "../src/Handoffs.tsx";
import { canLookUp, openByName } from "../src/MailboxName.tsx";
import {
  NavigationContext,
  OriginScope,
  PersonBreadcrumb,
  PersonLink,
  type PersonOrigin,
  originTrail,
  personOpener,
} from "../src/navigation.tsx";
import { People, PersonProfileView, type PersonProfileViewProps } from "../src/People.tsx";
import { SharedMailboxesSection } from "../src/PersonAccess.tsx";
import { DirectReportsValue, ManagerValue } from "../src/PersonDirectoryLinks.tsx";
import { findAll, hostTree, textOf } from "./markup.ts";

// Every person's name opens their page, and the page says where you came from. Synthetic names
// and .invalid addresses only.

const ID = {
  self: "44444444-4444-4444-8444-444444444444",
  manager: "55555555-5555-4555-8555-555555555551",
  reportA: "55555555-5555-4555-8555-555555555552",
  reportB: "55555555-5555-4555-8555-555555555553",
  billing: "eeeeeeee-0000-4000-8000-000000000002",
  group: "bbbbbbbb-0000-4000-8000-000000000003",
};

const buttons = (html: string) => [...html.matchAll(/<button[^>]*people-name[^>]*>/g)].length;
const noop = () => {};

describe("where you came from", () => {
  const calls: string[] = [];
  const nav = {
    openPerson: (id: string, from?: PersonOrigin | null) =>
      calls.push(`person:${id}:${from === null ? "none" : (from?.kind ?? "default")}`),
    openAudit: (id: string) => calls.push(`audit:${id}`),
    openView: (view: string) => calls.push(`view:${view}`),
  };
  const click = (crumbs: ReturnType<typeof originTrail>, at: number) => {
    calls.length = 0;
    crumbs[at]?.open();
    return [...calls];
  };

  it("has no trail without an origin", () => {
    expect(originTrail(null, nav)).toEqual([]);
  });

  it("names the page for a whole-page origin and returns to it", () => {
    const trail = originTrail({ kind: "view", view: "signins", label: "Sign-ins" }, nav);
    expect(trail.map((crumb) => crumb.label)).toEqual(["Sign-ins"]);
    expect(click(trail, 0)).toEqual(["view:signins"]);
  });

  it("names Audits and then the audit, each returning there", () => {
    const trail = originTrail(
      { kind: "audit", auditId: "accounts.inactive", title: "Inactive accounts" },
      nav,
    );
    expect(trail.map((crumb) => crumb.label)).toEqual(["Audits", "Inactive accounts"]);
    expect(click(trail, 0)).toEqual(["view:audits"]);
    expect(click(trail, 1)).toEqual(["audit:accounts.inactive"]);
  });

  it("names the person you came from, and reopens them with no origin of their own", () => {
    const trail = originTrail({ kind: "person", userId: ID.manager, name: "Mia Brandt" }, nav);
    expect(trail.map((crumb) => crumb.label)).toEqual(["Mia Brandt"]);
    expect(click(trail, 0)).toEqual([`person:${ID.manager}:none`]);
  });

  it("gives a page opened from a view that view as its origin, and the People list none", () => {
    expect(viewOrigin("signins")).toEqual({ kind: "view", view: "signins", label: "Sign-ins" });
    expect(viewOrigin("changes")).toMatchObject({ label: "Directory changes" });
    expect(viewOrigin("people")).toBeNull();
  });

  it("draws the breadcrumb with the trail, or with People and its way back without one", () => {
    const back = vi.fn();
    const withTrail = hostTree(
      createElement(PersonBreadcrumb, {
        trail: originTrail({ kind: "view", view: "signins", label: "Sign-ins" }, nav),
        name: "Tobias Vance",
        onBack: back,
      }),
    );
    expect(textOf(withTrail[0] ?? "")).toBe("Sign-ins/Tobias Vance");
    calls.length = 0;
    findAll(withTrail, (host) => host.tag === "button")[0]?.props.onClick?.call(null);
    expect(calls).toEqual(["view:signins"]);
    expect(back).not.toHaveBeenCalled();

    const plain = hostTree(
      createElement(PersonBreadcrumb, { trail: [], name: "Tobias Vance", onBack: back }),
    );
    expect(textOf(plain[0] ?? "")).toBe("People/Tobias Vance");
    const only = findAll(plain, (host) => host.tag === "button")[0];
    if (!only) throw new Error("The breadcrumb has no button.");
    (only.props.onClick as () => void)();
    expect(back).toHaveBeenCalledTimes(1);
  });
});

describe("person links", () => {
  it("open the person with the origin they were given, or the scope's, and only for a user id", () => {
    const openPerson = vi.fn();
    const from: PersonOrigin = { kind: "person", userId: ID.self, name: "Rowan Ellis" };
    personOpener({ openPerson }, ID.reportA, from)?.();
    expect(openPerson).toHaveBeenCalledWith(ID.reportA, from, undefined);
    personOpener({ openPerson }, ID.reportA, null)?.();
    expect(openPerson).toHaveBeenLastCalledWith(ID.reportA, null, undefined);
    expect(personOpener({ openPerson }, null, from)).toBeNull();
    personOpener({ openPerson }, ID.billing, undefined, "access")?.();
    expect(openPerson).toHaveBeenLastCalledWith(ID.billing, undefined, "access");
  });

  it("draw a button with an id and plain text without one", () => {
    const linked = renderToStaticMarkup(
      createElement(PersonLink, { userId: ID.reportA, name: "Ines Calder" }),
    );
    expect(linked).toContain('<button class="people-name"');
    expect(linked).toContain("Ines Calder");
    const plain = renderToStaticMarkup(createElement(PersonLink, { userId: null, name: "Team A" }));
    expect(plain).not.toContain("<button");
    expect(plain).toContain("Team A");
    const scoped = renderToStaticMarkup(
      createElement(
        OriginScope,
        { origin: null },
        createElement(PersonLink, { userId: ID.reportA, name: "Ines Calder" }),
      ),
    );
    expect(scoped).toContain("<button");
  });
});

describe("the person page", () => {
  const session: PrismSession = sessionSchema.parse({
    mode: "live",
    configured: true,
    authenticated: true,
    displayName: "Synthetic admin",
    csrfToken: "synthetic-csrf-token-value-000000000000",
    enabledFeatures: ["directory.users.read"],
    missingConfiguration: [],
  });
  const directory = (over: Record<string, unknown> = {}): PersonDirectory =>
    personDirectorySchema.parse({
      state: "available",
      profile: {
        id: ID.self,
        displayName: "Rowan Ellis",
        userPrincipalName: "rowan@example.invalid",
        mail: null,
        jobTitle: null,
        department: null,
        companyName: null,
        officeLocation: null,
        employeeId: null,
        userType: "Member",
        accountEnabled: true,
        createdDateTime: null,
        lastPasswordChangeDateTime: null,
        onPremisesSyncEnabled: null,
        usageLocation: null,
        city: null,
        state: null,
        country: null,
        mobilePhone: null,
        businessPhones: [],
      },
      manager: {
        state: "available",
        data: {
          id: ID.manager,
          displayName: "Mia Brandt",
          userPrincipalName: "mia@example.invalid",
          jobTitle: "Director",
        },
      },
      directReports: {
        state: "available",
        more: false,
        data: [
          {
            id: ID.reportA,
            displayName: "Ines Calder",
            userPrincipalName: null,
            jobTitle: "Analyst",
          },
          { id: ID.reportB, displayName: "Ravi Okoye", userPrincipalName: null, jobTitle: null },
        ],
      },
      signInActivity: null,
      licenses: null,
      ...over,
    });
  const props = (over: Partial<PersonProfileViewProps> = {}): PersonProfileViewProps => ({
    session,
    userId: ID.self,
    titleId: "title",
    tabsId: "x",
    tab: "overview",
    record: null,
    directory: directory(),
    signins: null,
    next: null,
    filter: "all",
    error: "",
    busy: null,
    changes: null,
    canary: {
      active: false,
      confirming: false,
      busy: false,
      error: "",
      onToggle: noop,
      onConfirm: noop,
      onCancel: noop,
    } as unknown as PersonProfileViewProps["canary"],
    onBack: noop,
    onTab: noop,
    onFilter: noop,
    onFindingState: noop,
    onMoreSignins: noop,
    onChanges: noop,
    ...over,
  });
  const draw = (over: Partial<PersonProfileViewProps> = {}) =>
    renderToStaticMarkup(createElement(PersonProfileView, props(over))).replaceAll("<!-- -->", "");
  const crumbs = (html: string) =>
    (html.match(/<nav class="breadcrumb"[^>]*>(.*?)<\/nav>/)?.[1] ?? "").replace(/<[^>]+>/g, " ");

  it("links the manager and each direct report, and shows their job titles", () => {
    const html = draw();
    expect(buttons(html)).toBe(3);
    for (const name of ["Mia Brandt", "Ines Calder", "Ravi Okoye"])
      expect(html).toMatch(new RegExp(`<button class="people-name"[^>]*><strong[^>]*>${name}<`));
    expect(html).toContain(", Director");
    expect(html).toMatch(/Analyst<\/span>/);
    expect(html).toContain("Direct reports");
  });

  it("gives direct reports a row of their own under the facts grid", () => {
    const html = draw();
    const grid = html.indexOf("</dl>");
    const row = html.indexOf('class="person-facts__reports"');
    expect(grid).toBeGreaterThan(-1);
    expect(row).toBeGreaterThan(grid);
    // The grid itself no longer carries the label, so many names cannot stretch one of its cells.
    expect(html.slice(0, grid)).not.toContain("Direct reports");
  });

  it("starts the breadcrumb at People when it has no origin", () => {
    expect(crumbs(draw()).replace(/\s+/g, " ").trim()).toBe("People / Rowan Ellis");
  });

  it("reads the origin from the navigation the App provides", () => {
    const opened = (origin: PersonOrigin | null) =>
      renderToStaticMarkup(
        createElement(
          NavigationContext.Provider,
          { value: { openPerson: noop, openAudit: noop, openView: noop, origin } },
          createElement(People, { session, selected: ID.self, onSelect: noop }),
        ),
      );
    expect(crumbs(opened(null)).replace(/\s+/g, " ").trim()).toBe(`People / ${ID.self}`);
    expect(
      crumbs(opened({ kind: "view", view: "signins", label: "Sign-ins" }))
        .replace(/\s+/g, " ")
        .trim(),
    ).toBe(`Sign-ins / ${ID.self}`);
  });

  it("starts the breadcrumb at the origin it was given", () => {
    const trail = originTrail(
      { kind: "audit", auditId: "accounts.inactive", title: "Inactive accounts" },
      { openPerson: noop, openAudit: noop, openView: noop },
    );
    expect(crumbs(draw({ trail })).replace(/\s+/g, " ").trim()).toBe(
      "Audits / Inactive accounts / Rowan Ellis",
    );
  });
});

describe("manager and direct reports", () => {
  const reports = (state: string, data: unknown[] = [], more = false, skipped = 0) =>
    personDirectorySchema.parse({
      state: "available",
      profile: null,
      manager: null,
      directReports: { state, more, skipped, data },
      signInActivity: null,
      licenses: null,
    }).directReports;
  const person = (n: number) => ({
    id: `66666666-6666-4666-8666-${String(n).padStart(12, "0")}`,
    displayName: `Person ${n}`,
    userPrincipalName: null,
    jobTitle: null,
  });
  const html = (node: Parameters<typeof renderToStaticMarkup>[0]) => renderToStaticMarkup(node);

  it("says None when there are none, and why when they could not be read", () => {
    expect(html(createElement(DirectReportsValue, { reports: reports("available") }))).toBe("None");
    expect(html(createElement(DirectReportsValue, { reports: reports("unavailable") }))).toBe(
      "Unavailable",
    );
    expect(
      html(createElement(DirectReportsValue, { reports: reports("missingConsent") })),
    ).toContain("did not allow");
    expect(html(createElement(DirectReportsValue, { reports: null }))).toBe("Unavailable");
  });

  it("counts contacts instead of saying None", () => {
    const only = html(
      createElement(DirectReportsValue, { reports: reports("available", [], false, 3) }),
    );
    expect(only).toContain("3 contacts, no accounts");
    expect(only).not.toContain("None");
    const some = html(
      createElement(DirectReportsValue, { reports: reports("available", [person(1)], false, 1) }),
    );
    expect(some).toContain("and 1 contact");
    expect(buttons(some)).toBe(1);
  });

  it("says and more when Microsoft has another page", () => {
    const drawn = html(
      createElement(DirectReportsValue, { reports: reports("available", [person(1)], true) }),
    );
    expect(drawn).toContain("and more");
    expect(buttons(drawn)).toBe(1);
    expect(
      html(createElement(DirectReportsValue, { reports: reports("available", [person(1)]) })),
    ).not.toContain("and more");
  });

  it("shows no more than 50 names and counts the rest", () => {
    const many = Array.from({ length: 50 }, (_, n) => person(n));
    const drawn = html(createElement(DirectReportsValue, { reports: reports("available", many) }));
    expect(buttons(drawn)).toBe(50);
    expect(drawn).not.toContain("and ");
    // A longer list than the contract allows is cut and counted, not trusted.
    const over = {
      state: "available" as const,
      more: false,
      data: Array.from({ length: 53 }, (_, n) => person(n)),
    };
    const cut = html(createElement(DirectReportsValue, { reports: over }));
    expect(buttons(cut)).toBe(50);
    expect(cut).toContain("and 3 more");
  });

  it("shows a manager that is a contact as text marked as one", () => {
    const drawn = html(
      createElement(ManagerValue, {
        manager: {
          state: "available",
          data: {
            id: ID.manager,
            displayName: "Cass Orr",
            userPrincipalName: null,
            jobTitle: "Adviser",
            isUser: false,
          },
        },
      }),
    );
    expect(drawn).toBe("Cass Orr, Adviser (contact)");
    expect(drawn).not.toContain("<button");
  });

  it("reads no manager as None set when available and Unavailable otherwise", () => {
    expect(html(createElement(ManagerValue, { manager: { state: "available", data: null } }))).toBe(
      "None set",
    );
    expect(
      html(createElement(ManagerValue, { manager: { state: "unavailable", data: null } })),
    ).toBe("Unavailable");
  });
});

describe("the Shared mailboxes section", () => {
  const mailbox = (holders: unknown[]): PersonMailboxAccess =>
    personMailboxAccessSchema.parse({
      import: {
        id: "44444444-4444-4444-8444-444444444444",
        exportedAt: new Date().toISOString(),
        importedAt: new Date().toISOString(),
        importedBy: { id: "22222222-2222-4222-8222-222222222222", name: "Synthetic admin" },
        includeUserMailboxes: false,
        mailboxes: 1,
        grants: 1,
        unresolved: 0,
      },
      groups: { status: "ok", readAt: new Date().toISOString(), count: 1, truncated: false },
      grants: [
        {
          mailbox: {
            id: ID.billing,
            address: "billing@example.invalid",
            name: "Billing",
            type: "SharedMailbox",
          },
          right: "FullAccess",
          via: { kind: "direct" },
        },
      ],
      grantsTruncated: false,
      holders,
      holdersTruncated: false,
    });
  const draw = (holders: unknown[]) =>
    renderToStaticMarkup(
      createElement(SharedMailboxesSection, { mailbox: mailbox(holders), error: "" }),
    ).replaceAll("<!-- -->", "");

  it("links each mailbox the person can use", () => {
    const html = draw([]);
    expect(buttons(html)).toBe(1);
    expect(html).toMatch(/<button class="people-name"[^>]*><strong[^>]*>Billing</);
  });

  it("does not link a mailbox's own row on its own page", () => {
    const own = renderToStaticMarkup(
      createElement(SharedMailboxesSection, {
        mailbox: mailbox([]),
        error: "",
        selfId: ID.billing,
      }),
    );
    expect(buttons(own)).toBe(0);
    expect(own).toContain("Billing");
  });

  it("links a holder that is a user with an id, and no other kind", () => {
    const holder = (kind: string, id: string | null, name: string) => ({
      right: "FullAccess",
      trustee: { id, address: null, name, kind },
    });
    const html = draw([
      holder("user", ID.reportA, "Ines Calder"),
      holder("user", null, "Nameless user"),
      holder("group", ID.group, "All plant staff"),
      holder("other", null, "Service robot"),
      holder("unresolved", null, "S-1-5-21-synthetic"),
    ]);
    // The shared mailbox itself, and the one user holder with an id.
    expect(buttons(html)).toBe(2);
    expect(html).toMatch(/<button class="people-name"[^>]*><strong[^>]*>Ines Calder</);
    for (const name of ["Nameless user", "All plant staff", "Service robot", "S-1-5-21-synthetic"])
      expect(html).toContain(name);
  });
});

describe("a change on a person's page", () => {
  const change = (over: Partial<ChangeEvent> = {}): ChangeEvent => ({
    id: "c-1",
    activityAt: "2026-10-22T15:30:00.000Z",
    activity: "Update user",
    category: "UserManagement",
    operationType: "Update",
    result: "success",
    resultReason: null,
    loggedByService: "Core Directory",
    correlationId: null,
    initiatedBy: { kind: "user", id: ID.manager, name: "Mia Brandt", principal: null },
    targets: [{ id: ID.self, name: "Rowan Ellis", type: "User", principal: null, properties: [] }],
    details: [],
    ...over,
  });
  const draw = (self?: string, over: Partial<ChangeEvent> = {}) =>
    renderToStaticMarkup(
      createElement(ChangeRow, {
        change: change(over),
        onOpenPerson: noop,
        ...(self ? { self } : {}),
      }),
    );

  it("links the other party and leaves the page's own person as text", () => {
    const html = draw(ID.self);
    expect(html.match(/<button class="link-button"/g)).toHaveLength(1);
    expect(html).toMatch(/<button class="link-button"[^>]*>Mia Brandt</);
  });

  it("links both parties on the all-changes page, which has no own person", () => {
    expect(draw().match(/<button class="link-button"/g)).toHaveLength(2);
  });

  it("does not link an app or a group", () => {
    const html = draw(ID.self, {
      initiatedBy: { kind: "app", id: "sp-1", name: "Synthetic Sync", principal: null },
      targets: [
        { id: ID.group, name: "Plant staff", type: "Group", principal: null, properties: [] },
      ],
    });
    expect(html).not.toContain("link-button");
  });
});

describe("a handoff receipt", () => {
  it("links an account with an id and leaves a bare sign-in name as text", () => {
    const html = renderToStaticMarkup(
      createElement(ReceiptItemsTable, {
        page: {
          receipt: {
            id: "77777777-7777-4777-8777-777777777777",
            tool: { name: "Example action tool", version: "1.4.0" },
            performedBy: "Jordan Vale",
            performedAt: "2026-10-22T15:30:00.000Z",
            items: 2,
          },
          items: [
            {
              position: 0,
              userId: ID.reportA,
              userPrincipalName: "ines@example.invalid",
              action: "account.sign-in-blocked",
              status: "done",
              detail: "",
              linked: true,
            },
            {
              position: 2,
              userId: ID.reportB,
              userPrincipalName: "someone.else@example.invalid",
              action: "account.sign-in-blocked",
              status: "done",
              detail: "",
              linked: false,
            },
            {
              position: 1,
              userId: null,
              userPrincipalName: "blake@example.invalid",
              action: "license.removed",
              status: "done",
              detail: "",
            },
          ],
          next: null,
        } as never,
        loading: false,
        onMore: null,
      }),
    );
    expect(buttons(html)).toBe(1);
    expect(html).toMatch(/<button class="people-name"[^>]*><strong[^>]*>ines@example.invalid</);
    expect(html).toContain("blake@example.invalid");
    // An item the server did not match to an account the handoff listed is text, id or not.
    expect(html).toContain("someone.else@example.invalid");
  });
});

describe("the Mailbox storage table", () => {
  const definition = auditDefinitionSchema.parse({
    id: "mail.mailbox-storage",
    version: 1,
    title: "Mailbox storage",
    description: "How full each mailbox is.",
    category: "Mail",
    subject: { one: "mailbox", many: "mailboxes" },
    columns: [
      { key: "mailbox", label: "Mailbox", kind: "text" },
      { key: "used", label: "Used (GB)", kind: "number" },
    ],
    views: [{ id: "all", label: "All", tags: [] }],
    defaultView: "all",
    defaultSort: { column: "used", direction: "desc" },
    permissions: ["Reports.Read.All"],
    optionalPermissions: [],
    requirements: [],
  });
  const row = (key: string, mailbox: string, personId: string | null = null) => ({
    key,
    personId,
    tags: [],
    cells: { mailbox, used: 4.2 },
  });
  const draw = (rows: ReturnType<typeof row>[], over: Partial<typeof definition> = {}) =>
    renderToStaticMarkup(
      createElement(ResultTable, {
        definition: { ...definition, ...over },
        run: {
          id: "66666666-6666-4666-8666-666666666666",
          auditId: definition.id,
          auditVersion: 1,
          status: "succeeded",
          complete: true,
          startedAt: "2026-10-22T15:30:00.000Z",
          finishedAt: "2026-10-22T15:31:00.000Z",
          startedBy: { id: "22222222-2222-4222-8222-222222222222", name: "Synthetic admin" },
          error: null,
          summary: [],
          notes: [],
          rowCount: rows.length,
          rows,
        } as unknown as AuditRunDetail,
        onOpenPerson: noop,
      }),
    );

  it("makes a real sign-in name a link, and a hashed one plain text", () => {
    const hashed = "0123456789abcdef0123456789ABCDEF";
    const html = draw([row("a", "billing@example.invalid"), row("b", hashed)]);
    expect(buttons(html)).toBe(1);
    expect(html).toMatch(/<button class="people-name"[^>]*><strong[^>]*>billing@example.invalid</);
    expect(html).toContain(hashed);
  });

  it("leaves other audits' text cells alone", () => {
    const html = draw([row("a", "billing@example.invalid")], { id: "accounts.inactive" });
    expect(buttons(html)).toBe(0);
  });

  it("only looks up what could be a name", () => {
    expect(canLookUp("billing@example.invalid")).toBe(true);
    expect(canLookUp("0123456789abcdef0123456789abcdef")).toBe(false);
    expect(canLookUp("0123456789abcdef0123456789abcdef@example.invalid")).toBe(true);
    expect(canLookUp("  ")).toBe(false);
    expect(canLookUp(null)).toBe(false);
  });

  it("opens the person when the lookup finds one", async () => {
    const open = vi.fn();
    const lookup = vi.fn(async () => ID.billing);
    await expect(openByName("billing@example.invalid", open, lookup)).resolves.toBeNull();
    expect(lookup).toHaveBeenCalledWith("billing@example.invalid", undefined);
    expect(open).toHaveBeenCalledWith(ID.billing);
  });

  it("says so in plain words when no account has the name", async () => {
    const open = vi.fn();
    const lookup = vi.fn(async () => {
      throw new Error("No account has that email in the connected tenant.");
    });
    await expect(openByName("ghost@example.invalid", open, lookup)).resolves.toBe(
      "No account has that email in the connected tenant.",
    );
    expect(open).not.toHaveBeenCalled();
  });

  it("opens nothing when the address changed before the answer came", async () => {
    const open = vi.fn();
    let here = true;
    await openByName("billing@example.invalid", open, async () => ID.billing, {
      stillWanted: () => here,
    });
    expect(open).toHaveBeenCalledTimes(1);
    here = false;
    await openByName("billing@example.invalid", open, async () => ID.billing, {
      stillWanted: () => here,
    });
    expect(open).toHaveBeenCalledTimes(1);
  });

  it("opens nothing for a lookup that was aborted, and says nothing", async () => {
    const open = vi.fn();
    const abort = new AbortController();
    const answer = openByName("billing@example.invalid", open, async () => ID.billing, {
      signal: abort.signal,
    });
    abort.abort();
    await expect(answer).resolves.toBeNull();
    expect(open).not.toHaveBeenCalled();
    const failing = new AbortController();
    const quiet = openByName(
      "billing@example.invalid",
      open,
      async () => {
        throw new Error("aborted");
      },
      { signal: failing.signal },
    );
    failing.abort();
    await expect(quiet).resolves.toBeNull();
  });

  it("lets only the latest lookup navigate", async () => {
    const open = vi.fn();
    let release: (id: string) => void = () => {};
    const slow = openByName(
      "slow@example.invalid",
      open,
      () => new Promise<string>((resolve) => (release = resolve)),
    );
    await openByName("fast@example.invalid", open, async () => ID.reportA);
    release(ID.reportB);
    await slow;
    expect(open).toHaveBeenCalledTimes(1);
    expect(open).toHaveBeenCalledWith(ID.reportA);
  });

  it("never looks a hashed name up", async () => {
    const open = vi.fn();
    const lookup = vi.fn(async () => ID.billing);
    const message = await openByName("0123456789abcdef0123456789abcdef", open, lookup);
    expect(message).not.toBeNull();
    expect(lookup).not.toHaveBeenCalled();
    expect(open).not.toHaveBeenCalled();
  });
});

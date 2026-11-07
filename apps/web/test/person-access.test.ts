import { type PersonAccess as PersonAccessData, personAccessSchema } from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { PersonAccess, PersonAccessView } from "../src/PersonAccess.tsx";

// The Access tab, drawn from an answer in hand. Synthetic names only.
const text = (html: string) => {
  let current = html.replaceAll("<!-- -->", "");
  let previous: string;
  do {
    previous = current;
    current = current.replace(/<[^>]*>/g, " ");
  } while (current !== previous);
  return current.replace(/&#x27;/g, "'").replace(/\s+/g, " ");
};

const readAt = "2026-10-20T10:00:00.000Z";
const membership = (over: Record<string, unknown>) => ({
  id: "g1",
  kind: "security",
  name: "Plant operators",
  relation: "direct",
  hasGroupMailbox: false,
  dynamic: false,
  roleTemplateId: null,
  ...over,
});
const policy = (over: Record<string, unknown>) => ({
  id: "p1",
  name: "Require MFA",
  state: "on",
  controls: "Require MFA",
  effect: "applies",
  targeted: true,
  includedBy: [{ kind: "all", id: null, name: null }],
  excludedBy: [],
  ...over,
});
const auditResult = (id: string, name: string, over: Record<string, unknown> = {}) => ({
  status: "ok",
  auditId: id,
  name,
  finishedAt: readAt,
  scope: null,
  matchedBy: "account",
  rows: [
    {
      key: "r1",
      cells: [
        { key: "registered", label: "MFA registered", kind: "boolean", value: true },
        { key: "strength", label: "Strongest method", kind: "enum", value: "Authenticator app" },
      ],
    },
  ],
  truncated: false,
  ...over,
});
const everyAudit = [
  ["directory.privileged-roles", "Privileged roles"],
  ["mfa.registration", "MFA registration"],
  ["devices.managed", "Managed devices"],
  ["mail.mailbox-storage", "Mailbox storage"],
  ["accounts.passwords", "Passwords"],
  ["identity.risky-users", "Risky users"],
  ["accounts.activity", "Account activity"],
  ["accounts.inactive", "Inactive accounts"],
  ["licenses.reclaim", "Licenses to reclaim"],
] as const;

function answer(over: Record<string, unknown> = {}): PersonAccessData {
  return personAccessSchema.parse({
    memberships: {
      status: "ok",
      readAt,
      truncated: false,
      entries: [
        membership({
          id: "g2",
          kind: "microsoft365",
          name: "Quality team",
          hasGroupMailbox: true,
        }),
        membership({}),
        membership({ id: "g3", name: "All plant staff", relation: "nested" }),
        membership({
          id: "r1",
          kind: "role",
          name: "Helpdesk Administrator",
          roleTemplateId: "t1",
        }),
        membership({ id: "a1", kind: "administrative-unit", name: "Plant site" }),
        membership({ id: "h1", kind: "hidden", name: null, relation: "nested" }),
      ],
    },
    owned: {
      status: "ok",
      readAt,
      truncated: false,
      entries: [
        { id: "o1", kind: "team", name: "Plant team" },
        { id: "o2", kind: "application", name: "Plant dashboard" },
      ],
    },
    conditionalAccess: {
      status: "ok",
      readAt,
      unaffected: 3,
      unreadable: 0,
      truncated: false,
      membershipsKnown: true,
      policies: [
        policy({}),
        policy({
          id: "p2",
          name: "Managed device",
          state: "off",
          controls: "Require a compliant device",
          effect: "excluded",
          includedBy: [{ kind: "group", id: "g3", name: "All plant staff" }],
          excludedBy: [{ kind: "person", id: "u1", name: null }],
        }),
      ],
    },
    audits: {
      status: "ok",
      results: everyAudit.map(([id, name]) =>
        auditResult(
          id,
          name,
          id === "directory.privileged-roles"
            ? {
                rows: [
                  {
                    key: "role-1",
                    cells: [
                      { key: "role", label: "Role", kind: "enum", value: "Exchange Administrator" },
                      { key: "privileged", label: "Privileged", kind: "boolean", value: true },
                      { key: "held", label: "How held", kind: "enum", value: "Eligible" },
                      {
                        key: "ends",
                        label: "Ends",
                        kind: "date",
                        value: "2026-12-01T00:00:00.000Z",
                      },
                    ],
                  },
                ],
              }
            : {},
        ),
      ),
    },
    ...over,
  });
}

const draw = (access: PersonAccessData | null, error = "") =>
  text(renderToStaticMarkup(createElement(PersonAccessView, { access, error })));

describe("the Access tab", () => {
  it("waits with grey rows while the answer loads, and shows an error as a line", () => {
    const html = renderToStaticMarkup(createElement(PersonAccessView, { access: null, error: "" }));
    expect(html).toContain("skeleton");
    expect(draw(null, "Access could not be loaded.")).toContain("Access could not be loaded.");
  });

  it("shows nothing, and so asks for nothing, until the tab is first opened", () => {
    const session = {
      mode: "live" as const,
      configured: true,
      authenticated: true,
      displayName: "Synthetic admin",
      csrfToken: null,
      enabledFeatures: [],
      missingConfiguration: [],
    };
    const html = renderToStaticMarkup(
      createElement(PersonAccess, { session, userId: "u1", active: false }),
    );
    expect(html).toBe("");
  });

  it("lists groups by kind, direct or nested, marks group mailboxes and keeps hidden objects", () => {
    const words = draw(answer());
    expect(words).toContain("Groups and roles");
    expect(words).toContain("Microsoft 365 groups");
    expect(words).toContain("Quality team");
    expect(words).toContain("Group mailbox");
    expect(words).toContain("Security groups");
    expect(words).toMatch(/All plant staff\s+Nested/);
    expect(words).toMatch(/Plant operators\s+Direct/);
    expect(words).toContain("Directory roles held now");
    expect(words).toContain("Helpdesk Administrator");
    expect(words).toContain("Administrative units");
    expect(words).toContain("Hidden by Microsoft");
    expect(words).toContain("Live from Microsoft");
  });

  it("shows the eligible and time-bound roles from the privileged roles audit", () => {
    const words = draw(answer());
    expect(words).toContain("Eligible and time-bound roles");
    expect(words).toContain("From the Privileged roles audit");
  });

  it("lists what the account owns, with the hand-over note", () => {
    const words = draw(answer());
    expect(words).toContain("Plant team");
    expect(words).toContain("Team");
    expect(words).toContain("Plant dashboard");
    expect(words).toContain("Hand ownership over before this account is removed");
    expect(
      draw(answer({ owned: { status: "ok", readAt, truncated: false, entries: [] } })),
    ).toContain("owns no groups, teams or applications");
  });

  it("shows policies with state, controls and why, and says conditions still decide", () => {
    const words = draw(answer());
    expect(words).toContain("Require MFA");
    expect(words).toContain("Applies through all users");
    expect(words).toContain("Left out by name.");
    expect(words).toContain("through the group All plant staff");
    expect(words).toContain("Off");
    expect(words).toContain(
      "Conditions such as apps, platforms and locations still decide whether a policy applies at a given sign-in.",
    );
    expect(words).toContain("3 other policies do not reach this person");
  });

  it("shows the policies that leave the person out first when one of them is on", () => {
    const risky = answer({
      conditionalAccess: {
        status: "ok",
        readAt,
        unaffected: 0,
        unreadable: 0,
        truncated: false,
        membershipsKnown: true,
        policies: [
          policy({}),
          policy({
            id: "p3",
            name: "Block legacy sign-in",
            effect: "excluded",
            controls: "Block access",
            excludedBy: [{ kind: "group", id: "g9", name: "Break glass" }],
          }),
        ],
      },
    });
    const words = draw(risky);
    expect(words.indexOf("Policies that leave them out")).toBeLessThan(
      words.indexOf("Policies that apply"),
    );
    expect(words).toContain("Excluded");
    expect(words).toContain("Left out through the group Break glass.");
    expect(words).not.toContain("Medium");
    // With the exclusion on an off policy, the applying policies lead.
    const calm = draw(answer());
    expect(calm.indexOf("Policies that apply")).toBeLessThan(
      calm.indexOf("Policies that leave them out"),
    );
  });

  it("notes when memberships were unknown, policies were unreadable or the list was cut", () => {
    const words = draw(
      answer({
        conditionalAccess: {
          status: "ok",
          readAt,
          policies: [],
          unaffected: 0,
          unreadable: 2,
          truncated: true,
          membershipsKnown: false,
        },
      }),
    );
    expect(words).toContain("No Conditional Access policy names or reaches this person");
    expect(words).toContain("Group and role memberships could not be read");
    expect(words).toContain("2 policies could not be read and are not counted");
    expect(words).toContain("more policies than were read");
  });

  it("shows the stored audits as label and value lists under their own headings", () => {
    const words = draw(answer());
    for (const heading of ["Sign-in methods", "Devices", "Mailbox", "Password", "Risk", "Activity"])
      expect(words).toContain(heading);
    expect(words).toContain("MFA registered");
    expect(words).toContain("Strongest method Authenticator app");
    expect(words).toContain("From the MFA registration audit");
    // A role that is only eligible, with the end date, in the shared severity style.
    expect(words).toMatch(/Exchange Administrator Privileged ?: Eligible ?, ends/);
  });

  it("counts the groups of each kind in the headings", () => {
    const words = draw(answer());
    expect(words).toContain("Security groups (2)");
    expect(words).toContain("Microsoft 365 groups (1)");
  });

  it("links an audit that has not run, and says so for one without a row", () => {
    const html = renderToStaticMarkup(
      createElement(PersonAccessView, {
        access: answer({
          audits: {
            status: "ok",
            results: [
              { status: "no_run", auditId: "devices.managed", name: "Managed devices" },
              auditResult("mfa.registration", "MFA registration", { rows: [] }),
              auditResult("mail.mailbox-storage", "Mailbox storage", {
                matchedBy: "sign-in name",
                scope: ["example.invalid"],
              }),
              {
                status: "failed",
                auditId: "accounts.passwords",
                name: "Passwords",
                message: "Something plain.",
              },
            ],
          },
        }),
        error: "",
      }),
    );
    expect(html).toContain('href="#/audits/devices.managed"');
    const words = text(html);
    expect(words).toContain("Run the Managed devices audit");
    expect(words).toContain("Nothing about this person in the latest MFA registration run.");
    expect(words).toContain("matched by sign-in name");
    expect(words).toContain("limited to example.invalid");
    expect(words).toContain("Something plain.");
  });

  it("says what a refused section needs, with the pack and a link to Microsoft access", () => {
    const html = renderToStaticMarkup(
      createElement(PersonAccessView, {
        access: answer({
          memberships: {
            status: "not_granted",
            pack: "directory.users.read",
            permission: "User.Read.All",
            enabled: true,
          },
          owned: {
            status: "not_granted",
            pack: "directory.users.read",
            permission: "User.Read.All",
            enabled: false,
          },
          conditionalAccess: { status: "failed", message: "Microsoft refused this read." },
          audits: { status: "failed", message: "The stored audit results could not be read." },
        }),
        error: "",
      }),
    );
    const words = text(html);
    expect(words).toContain("needs the User.Read.All permission");
    expect(words).toContain("needs the directory.users.read read pack");
    expect(words).toContain("ENABLED_READ_FEATURES");
    expect(words).toContain("Microsoft refused this read.");
    expect(words).toContain("The stored audit results could not be read.");
    expect(html).toContain('href="#/access"');
  });

  it("waits for its own answer in the Shared mailboxes section", () => {
    const html = renderToStaticMarkup(
      createElement(PersonAccessView, { access: answer(), error: "" }),
    );
    expect(html).toContain('<span class="visually-hidden">Reading shared mailbox access</span>');
  });

  it("is written without em dashes, middle dots or exclamation marks", () => {
    const words = draw(answer());
    expect(words).not.toMatch(/[—–·!]/);
  });
});

describe("the Access tab, empty audit wording", () => {
  const withAudits = (results: unknown[]) => draw(answer({ audits: { status: "ok", results } }));

  it("says a person is not flagged in the Risky users run, and when", () => {
    const words = withAudits([auditResult("identity.risky-users", "Risky users", { rows: [] })]);
    expect(words).toMatch(/Not flagged in the latest Risky users run, Oct 20, 2026 ?\./);
  });

  it("says nothing was found in any other audit's latest run, and hides an absent reclaim line", () => {
    const words = withAudits([
      auditResult("devices.managed", "Managed devices", { rows: [] }),
      auditResult("licenses.reclaim", "Licenses to reclaim", { rows: [] }),
    ]);
    expect(words).toContain("Nothing about this person in the latest Managed devices run.");
    expect(words).not.toContain("License to reclaim");
  });
});

import { type PersonSummary, personAccessSchema, personSummarySchema } from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { PersonAccessView } from "../src/PersonAccess.tsx";
import { personHref, PersonSummaryView } from "../src/PersonSummary.tsx";
import { accessibleName, sectionIds, summaryItems } from "../src/personSummaryLines.ts";
import { nestedInteractive } from "./markup.ts";

// The Overview's "Where they show up" panel, drawn from an answer in hand. Synthetic names only.
const at = "2026-10-20T10:00:00.000Z";
const text = (html: string) => {
  let current = html.replaceAll("<!-- -->", "");
  let previous: string;
  do {
    previous = current;
    current = current.replace(/<[^>]*>/g, " ");
  } while (current !== previous);
  return current.replace(/&#x27;/g, "'").replace(/\s+/g, " ");
};
const cells = (values: Record<string, string | number | boolean | string[]>) =>
  Object.entries(values).map(([key, value]) => ({
    key,
    label: key,
    kind: typeof value === "boolean" ? "boolean" : "text",
    value,
  }));
const ok = (
  auditId: string,
  name: string,
  values: Record<string, never> | Record<string, unknown>,
) => ({
  status: "ok",
  auditId,
  name,
  finishedAt: at,
  complete: true,
  scope: null,
  matchedBy: "account",
  rows: [{ key: "r", cells: cells(values as Record<string, string>) }],
  truncated: false,
});
const notInRun = (auditId: string, name: string) => ({
  status: "not_in_run",
  auditId,
  name,
  finishedAt: at,
  complete: true,
  scope: null,
  matchedBy: "account",
});
const noRun = (auditId: string, name: string) => ({ status: "no_run", auditId, name });

const quiet = (over: Record<string, unknown> = {}): PersonSummary =>
  personSummarySchema.parse({
    combinedRisks: {
      status: "ok",
      finishedAt: at,
      severity: null,
      rules: [],
      reasons: "",
      gaps: [],
      complete: true,
    },
    signinFindings: { status: "ok", open: 0, highest: null, more: false },
    audits: [
      notInRun("identity.risky-users", "Risky users"),
      notInRun("directory.privileged-roles", "Privileged roles"),
      ok("accounts.activity", "Account activity", { verdict: "Active", lastSuccess: at }),
      ok("mfa.registration", "MFA registration", {
        registered: true,
        strength: "Authenticator app",
      }),
      notInRun("guests.review", "Guest accounts"),
      notInRun("devices.managed", "Managed devices"),
      notInRun("usage.m365-apps", "Microsoft 365 apps"),
    ],
    owned: { status: "ok", finishedAt: at, complete: true, entries: [], more: false },
    sharedMailboxes: { status: "no_import" },
    changes: { status: "ok", days: 30, count: 0, more: false },
    handoff: { status: "ok", latest: null },
    ...over,
  });
const render = (summary: PersonSummary | null, error = "") =>
  text(
    renderToStaticMarkup(
      createElement(PersonSummaryView, { summary, error, userId: "u1", onJump: () => {} }),
    ),
  );

const critical = quiet({
  combinedRisks: {
    status: "ok",
    finishedAt: at,
    severity: "critical",
    rules: [
      { id: "a", label: "Privileged without MFA", severity: "critical" },
      { id: "b", label: "Dormant owner", severity: "high" },
    ],
    reasons: "Synthetic reason.",
    gaps: [],
    complete: false,
  },
  signinFindings: { status: "ok", open: 2, highest: "high", more: false },
  owned: {
    status: "ok",
    finishedAt: at,
    complete: true,
    more: false,
    entries: [
      { id: "g1", name: "Plant team", kind: "Team", status: "One owner", tags: ["one-owner"] },
      { id: "g2", name: "Other", kind: "Security group", status: "OK", tags: [] },
    ],
  },
  changes: { status: "ok", days: 30, count: 4, more: false },
  handoff: {
    status: "ok",
    latest: {
      state: "failed",
      action: "account.sign-in-blocked",
      tool: "Example tool",
      performedAt: at,
    },
  },
});

const DAY = 86_400_000;
// Two days after the sample run, so "2 days ago" is what the sentences say.
const now = Date.parse(at) + 2 * DAY;
const sentences = (summary: PersonSummary) =>
  new Map(summaryItems(summary, now).items.map((entry) => [entry.id, entry.sentence]));
const html = (summary: PersonSummary | null, error = "", userId = "u1") =>
  renderToStaticMarkup(
    createElement(PersonSummaryView, { summary, error, userId, onJump: () => {}, now }),
  );

describe("the Where they show up panel", () => {
  it("keeps the fixed order, in two columns, and gives a quiet account a short, calm panel", () => {
    const { items, notRun } = summaryItems(quiet(), now);
    expect(items.map((entry) => [entry.column, entry.source])).toEqual([
      ["risk", "Combined risks"],
      ["risk", "Sign-in findings"],
      ["risk", "Risky users"],
      // Without an export, the row says so and links to where one is imported.
      ["risk", "Shared mailboxes"],
      ["account", "Account activity"],
      ["account", "MFA registration"],
      ["account", "Directory changes"],
    ]);
    expect(notRun).toEqual([]);
    const markup = text(html(quiet()));
    expect(markup).toContain("Risk and access");
    expect(markup.indexOf("Combined risks")).toBeLessThan(markup.indexOf("Account activity"));
    const said = sentences(quiet());
    expect(said.get("combined-risks")).toBe("No combined risks");
    expect(said.get("signin-findings")).toBe("No open findings");
    expect(said.get("identity.risky-users")).toBe("Not flagged");
    expect(said.get("accounts.activity")).toBe("Active, signed in 2 days ago");
    expect(said.get("mfa.registration")).toBe("MFA registered, Authenticator app");
    // Privileged roles, Guest, Devices, Microsoft 365 apps and groups are
    // normally empty, so they are left out rather than shown as nothing.
    for (const left of ["Privileged roles", "Guest accounts", "Managed devices", "Groups they own"])
      expect(markup).not.toContain(left);
    expect(markup).not.toContain("Not run yet");
  });

  it("leads a critical account with its combined risk, in words, and says what they own", () => {
    const markup = text(html(critical));
    expect(markup.indexOf("Combined risks")).toBeLessThan(markup.indexOf("Sign-in findings"));
    const said = sentences(critical);
    expect(said.get("combined-risks")).toBe("Privileged without MFA and 1 more rule");
    expect(markup).toContain("Critical");
    expect(said.get("signin-findings")).toBe("2 open findings");
    expect(markup).toContain("High");
    expect(said.get("owned-groups")).toBe("Owns 2 groups and teams, 1 needs attention");
    expect(said.get("directory-changes")).toBe("4 changes in the last 30 days");
    expect(said.get("handoff")).toBe("Account sign in blocked, failed 2 days ago");
    const names = summaryItems(critical, now).items.map(accessibleName);
    expect(names[0]).toBe(
      "Combined risks, critical: Privileged without MFA and 1 more rule. That run could not check everything. Opens the Access tab.",
    );
  });

  it("is a list of links, one per source, with the destination in the name and no repeated hints", () => {
    const markup = html(critical);
    expect(markup).toContain('href="#/people/u1/access"');
    expect(markup).toContain('href="#/people/u1/changes"');
    expect(markup).toContain(
      'aria-label="Directory changes: 4 changes in the last 30 days. Opens the Changes tab."',
    );
    expect(text(markup)).not.toContain("Open on the Access tab");
    expect(text(markup)).not.toContain("Last run");
    expect(nestedInteractive(markup)).toBe(0);
    const rows = markup.match(/class="person-summary__row/g) ?? [];
    expect(rows.length).toBe(summaryItems(critical, now).items.length);
  });

  it("folds sources that never ran into one last line linking to the audits page", () => {
    const summary = quiet({
      combinedRisks: { status: "no_run" },
      audits: [
        noRun("identity.risky-users", "Risky users"),
        noRun("accounts.passwords", "Passwords"),
        noRun("devices.managed", "Managed devices"),
        noRun("guests.review", "Guest accounts"),
        ok("mfa.registration", "MFA registration", { registered: false }),
      ],
      owned: { status: "no_run" },
    });
    const markup = html(summary);
    expect(text(markup)).toContain(
      "Not run yet: Combined risks, Risky users, Groups and Teams, Passwords, Managed devices. Run a full review",
    );
    expect(markup).toContain('href="#/audits"');
    expect(text(markup)).not.toContain("Guest accounts");
    expect(sentences(summary).get("mfa.registration")).toBe("No MFA registered");
  });

  it("says when a source cannot match the account and when one failed, quietly", () => {
    const { items } = summaryItems(
      quiet({
        audits: [
          {
            status: "no_name",
            auditId: "mail.mailbox-storage",
            name: "Mailbox storage",
            finishedAt: at,
            complete: true,
          },
          { status: "failed", auditId: "accounts.passwords", name: "Passwords", message: "x" },
        ],
        handoff: { status: "failed", message: "x" },
      }),
      now,
    );
    const byId = new Map(items.map((entry) => [entry.id, entry]));
    expect(byId.get("mail.mailbox-storage")?.sentence).toMatch(/^Could not be matched/);
    expect(byId.get("mail.mailbox-storage")?.muted).toBe(true);
    expect(byId.get("accounts.passwords")?.sentence).toBe("Passwords could not be read just now");
    expect(byId.get("handoff")?.muted).toBe(true);
  });

  it("shows a skeleton while loading and one plain line for a failed load", () => {
    expect(html(null)).toContain("skeleton");
    const failed = text(html(null, "Where this person shows up could not be loaded."));
    expect(failed).toContain("Where they show up Where this person shows up could not be loaded.");
    expect(failed).not.toContain("Combined risks");
  });

  it("escapes the address and sends every section row to an id the tabs carry", () => {
    expect(personHref("u 1", "access")).toBe("#/people/u%201/access");
    expect(personHref("u 1", "overview")).toBe("#/people/u%201");
    const jumps = summaryItems(critical, now).items.map((entry) => entry.jump);
    for (const jump of jumps)
      if (jump.kind === "section") expect(Object.values(sectionIds)).toContain(jump.id);
  });

  it("words each source as a short sentence", () => {
    const rowsOf = (id: string, name: string, rows: Record<string, unknown>[]) => ({
      ...ok(id, name, {}),
      rows: rows.map((values, index) => ({
        key: `r${index}`,
        cells: cells(values as Record<string, string>),
      })),
    });
    const threeDaysAgo = new Date(now - 3 * DAY).toISOString();
    const said = sentences(
      quiet({
        audits: [
          rowsOf("directory.privileged-roles", "Privileged roles", [
            { role: "Global Administrator", privileged: true },
            { role: "Helpdesk Administrator", privileged: false },
            { role: "Reports Reader", privileged: false },
          ]),
          rowsOf("accounts.passwords", "Passwords", [
            { lastChanged: threeDaysAgo, policy: "Managed on-premises", status: "OK" },
          ]),
          rowsOf("devices.managed", "Managed devices", [
            { compliance: "Compliant", encrypted: true },
            { compliance: "Compliant", encrypted: true },
          ]),
          rowsOf("mail.mailbox-storage", "Mailbox storage", [
            { used: 1.94, sendQuota: 49.5, percent: 4, status: "OK" },
          ]),
          rowsOf("usage.m365-apps", "Microsoft 365 apps", [
            {
              lastActivity: at,
              apps: ["Outlook", "Word", "Excel", "Teams"],
              platforms: ["Windows"],
            },
          ]),
          rowsOf("guests.review", "Guest accounts", [{ invitation: "Not accepted", invited: at }]),
        ],
        sharedMailboxes: {
          status: "ok",
          exportedAt: at,
          importedAt: at,
          includesUserMailboxes: false,
          mailboxes: 2,
          mailboxesMore: false,
          holders: null,
          holdersMore: false,
        },
      }),
    );
    expect(said.get("directory.privileged-roles")).toBe("Global Administrator and 2 more roles");
    expect(said.get("accounts.passwords")).toBe("Changed 3 days ago, managed on-premises");
    expect(said.get("devices.managed")).toBe("2 devices, all compliant and encrypted");
    expect(said.get("mail.mailbox-storage")).toBe("1.9 of 49.5 GB used");
    expect(said.get("usage.m365-apps")).toBe("Active 2 days ago in Outlook, Word, Excel and Teams");
    expect(said.get("guests.review")).toBe("Invitation not accepted, sent 2 days ago");
    expect(said.get("shared-mailboxes")).toBe("Can use 2 shared mailboxes");
    const roles = summaryItems(
      quiet({
        audits: [
          rowsOf("directory.privileged-roles", "Privileged roles", [
            { role: "Global Administrator", privileged: true },
          ]),
        ],
      }),
      now,
    ).items.find((entry) => entry.id === "directory.privileged-roles");
    expect(roles).toMatchObject({
      severity: "high",
      severityLabel: "Privileged",
      sentence: "Global Administrator",
    });
    // A device that fails compliance is a warning with the count in words.
    const device = summaryItems(
      quiet({
        audits: [
          rowsOf("devices.managed", "Managed devices", [
            { compliance: "Compliant", encrypted: true },
            { compliance: "Not compliant", encrypted: true },
          ]),
          rowsOf("mfa.registration", "MFA registration", [{ registered: false }]),
        ],
      }),
      now,
    ).items;
    expect(device.find((entry) => entry.id === "devices.managed")).toMatchObject({
      sentence: "1 of 2 not compliant",
      tone: "warning",
    });
    expect(device.find((entry) => entry.id === "mfa.registration")).toMatchObject({
      sentence: "No MFA registered",
      tone: "bad",
    });
  });
});

describe("the Access tab, with the new sections and owned objects", () => {
  const base = {
    memberships: { status: "failed", message: "x" },
    conditionalAccess: { status: "failed", message: "x" },
  };
  const entries = [
    { id: "g1", kind: "team", name: "Plant team" },
    { id: "obj", kind: "application", name: "Plant dashboard" },
    { id: "sp", kind: "service-principal", name: "Plant sync" },
    { id: "none", kind: "application", name: "No row" },
  ];
  const view = (over: Record<string, unknown>) =>
    text(
      renderToStaticMarkup(
        createElement(PersonAccessView, {
          access: personAccessSchema.parse({
            ...base,
            owned: { status: "ok", readAt: at, truncated: false, entries, ...over },
            audits: { status: "ok", results: [] },
          }),
          error: "",
        }),
      ),
    );

  it("shows each owned object's status, matched application, service and missing row", () => {
    const html = view({
      lookups: { groups: { status: "ok", finishedAt: at }, apps: { status: "ok", finishedAt: at } },
      details: [
        { kind: "group", id: "g1", status: "One owner", tags: ["one-owner", "team", "guests"] },
        {
          kind: "app",
          id: "obj",
          highRisk: true,
          applicationPermissions: true,
          adminConsent: true,
          userConsent: false,
        },
        {
          kind: "app",
          id: "sp",
          highRisk: false,
          applicationPermissions: false,
          adminConsent: false,
          userConsent: true,
        },
      ],
    });
    expect(html).toContain("One owner One owner, Has guests");
    expect(html).toContain(
      "High-risk permissions Holds application permissions, usable with no one signed in. Admin consent for everyone.",
    );
    expect(html).toContain("Holds no application permissions. Consented to by users.");
    expect(html).toContain("Not in the latest App consents run.");
  });

  it("says to run the audit when it never ran, and shows nothing for an older answer", () => {
    const never = view({
      lookups: {
        groups: { status: "no_run" },
        apps: { status: "failed", message: "Could not read." },
      },
    });
    expect(never).toContain("Run Groups and Teams to see this object's status here.");
    expect(never).toContain("Could not read.");
    const older = view({});
    expect(older).not.toContain("App consents");
  });

  it("shows the Guest section only when the answer holds a guest result", () => {
    const withResults = (results: unknown[]) =>
      text(
        renderToStaticMarkup(
          createElement(PersonAccessView, {
            access: personAccessSchema.parse({
              ...base,
              owned: { status: "failed", message: "x" },
              audits: { status: "ok", results },
            }),
            error: "",
          }),
        ),
      );
    expect(withResults([])).not.toContain("Guest");
    expect(withResults([])).toContain("Microsoft 365 apps");
    const guest = withResults([
      {
        status: "ok",
        auditId: "guests.review",
        name: "Guest accounts",
        finishedAt: at,
        scope: null,
        matchedBy: "account",
        truncated: false,
        rows: [
          {
            key: "g",
            cells: [
              { key: "invitation", label: "Invitation", kind: "enum", value: "Not accepted" },
            ],
          },
        ],
      },
    ]);
    expect(guest).toContain("Not accepted");
  });

  it("carries the id every Access jump from the summary scrolls to", () => {
    const results = [
      "identity.risky-users",
      "directory.privileged-roles",
      "accounts.activity",
      "mfa.registration",
      "accounts.passwords",
      "guests.review",
      "devices.managed",
      "mail.mailbox-storage",
      "usage.m365-apps",
    ].map((id) => ({
      status: "no_run",
      auditId: id,
      name: id,
    }));
    const html = renderToStaticMarkup(
      createElement(PersonAccessView, {
        access: personAccessSchema.parse({
          ...base,
          combinedRisks: { status: "no_run" },
          owned: { status: "failed", message: "x" },
          audits: { status: "ok", results },
        }),
        error: "",
      }),
    );
    const wanted = summaryItems(critical)
      .items.concat(summaryItems(quiet()).items)
      .flatMap((entry) =>
        entry.jump.kind === "section" && entry.jump.tab === "access" ? [entry.jump.id] : [],
      );
    expect(wanted.length).toBeGreaterThan(0);
    for (const id of new Set([
      ...wanted,
      sectionIds.risk,
      sectionIds.groups,
      sectionIds.activity,
      sectionIds.methods,
      sectionIds.password,
      sectionIds.guest,
      sectionIds.devices,
      sectionIds.mailbox,
      sectionIds.apps,
      sectionIds.owns,
      sectionIds.combinedRisks,
      sectionIds.sharedMailboxes,
    ]))
      expect(html, id).toContain(`id="${id}"`);
  });
});

describe("the Microsoft 365 apps section on the Access tab", () => {
  it("reads as one sentence, not a list of fields", () => {
    const last = new Date(Date.now() - 2 * DAY).toISOString();
    const markup = text(
      renderToStaticMarkup(
        createElement(PersonAccessView, {
          access: personAccessSchema.parse({
            memberships: { status: "failed", message: "x" },
            conditionalAccess: { status: "failed", message: "x" },
            owned: { status: "failed", message: "x" },
            audits: {
              status: "ok",
              results: [
                {
                  status: "ok",
                  auditId: "usage.m365-apps",
                  name: "Microsoft 365 apps",
                  finishedAt: at,
                  scope: null,
                  matchedBy: "sign-in name",
                  truncated: false,
                  rows: [
                    {
                      key: "u",
                      cells: [
                        { key: "lastActivity", label: "Last activity", kind: "date", value: last },
                        {
                          key: "apps",
                          label: "Apps used",
                          kind: "list",
                          value: ["Outlook", "Word", "Excel", "Teams"],
                        },
                        { key: "platforms", label: "Platforms", kind: "list", value: ["Windows"] },
                      ],
                    },
                  ],
                },
              ],
            },
          }),
          error: "",
        }),
      ),
    );
    expect(markup).toContain(
      "Last active 2 days ago in Outlook, Word, Excel and Teams, on Windows",
    );
    expect(markup).not.toContain("Platforms Windows");
  });
});

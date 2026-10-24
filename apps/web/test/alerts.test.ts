import {
  type AlertOverview,
  alertOverviewSchema,
  type PrismSession,
  sessionSchema,
} from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";
import {
  Alerts,
  AlertsView,
  CreatedSecret,
  DeliveryTable,
  deliveryStatus,
  errorText,
  RuleForm,
  ruleSummary,
} from "../src/Alerts.tsx";

// The Alerts settings page, rendered without a browser. Synthetic destinations and rules only.

const noop = () => {};
const text = (html: string) =>
  html
    .replace(/<!--.*?-->/g, "")
    .replace(/<[^>]+>/g, " ")
    .replace(/&#x27;/g, "'")
    .replace(/\s+/g, " ");

const hookId = "bbbbbbbb-0000-4000-8000-000000000001";
const mailId = "bbbbbbbb-0000-4000-8000-000000000002";
const ruleId = "aaaaaaaa-0000-4000-8000-000000000001";

function overview(overrides: Partial<AlertOverview> = {}): AlertOverview {
  return alertOverviewSchema.parse({
    settings: {
      email: { configured: true, domains: ["example.invalid"] },
      webhookPorts: [],
      temporaryKey: false,
    },
    choices: {
      modules: [{ id: "signins.audit", title: "Sign-in audit" }],
      audits: [
        {
          id: "accounts.inactive",
          title: "Inactive accounts",
          tiles: [{ id: "inactive-accounts", label: "Inactive accounts" }],
        },
      ],
    },
    exports: [{ auditId: "accounts.inactive", destinationId: mailId }],
    weekly: { destinationId: hookId, timeZone: "America/Chicago" },
    destinations: [
      {
        id: hookId,
        kind: "teams",
        label: "Security channel",
        host: "hooks.example",
        recipients: [],
        includeDetails: false,
        readable: true,
        recipientsAllowed: true,
        createdAt: "2026-10-24T10:00:00.000Z",
      },
      {
        id: mailId,
        kind: "email",
        label: "Ops mail",
        host: null,
        recipients: ["ops@example.invalid"],
        includeDetails: true,
        readable: true,
        recipientsAllowed: true,
        createdAt: "2026-10-24T10:00:00.000Z",
      },
    ],
    rules: [
      {
        id: ruleId,
        kind: "finding",
        name: "High findings",
        enabled: true,
        destinations: [hookId],
        quietMinutes: 60,
        minimumSeverity: "high",
        modules: ["signins.audit"],
        createdAt: "2026-10-24T10:00:00.000Z",
        createdBy: "22222222-2222-4222-8222-222222222222",
      },
    ],
    otherRules: [],
    deliveries: [
      {
        id: "cccccccc-0000-4000-8000-000000000001",
        destinationId: hookId,
        source: "rule",
        ruleId,
        items: 3,
        status: "failed",
        attempts: 3,
        lastError: "server_error",
        createdAt: "2026-10-24T11:00:00.000Z",
        updatedAt: "2026-10-24T11:05:00.000Z",
      },
      {
        id: "cccccccc-0000-4000-8000-000000000002",
        destinationId: "dddddddd-0000-4000-8000-000000000009",
        source: "test",
        ruleId: null,
        items: 0,
        status: "sent",
        attempts: 1,
        lastError: null,
        createdAt: "2026-10-24T10:30:00.000Z",
        updatedAt: "2026-10-24T10:30:01.000Z",
      },
    ],
    ...overrides,
  });
}

const page = (value: AlertOverview, extra: Partial<Parameters<typeof AlertsView>[0]> = {}) =>
  renderToStaticMarkup(
    createElement(AlertsView, {
      overview: value,
      busy: false,
      error: "",
      created: null,
      onCreateDestination: noop,
      onCreatedDone: noop,
      onTest: noop,
      onDeleteDestination: noop,
      onCreateRule: noop,
      onToggleRule: noop,
      onDeleteRule: noop,
      onExport: noop,
      onWeekly: noop,
      ...extra,
    }),
  );

describe("alerts page", () => {
  it("lists destinations by host or recipients, with test and delete", () => {
    const html = page(overview());
    const words = text(html);
    expect(words).toContain("Security channel");
    expect(words).toContain("Teams workflow, hooks.example");
    expect(words).toContain("Email, ops@example.invalid, with details");
    expect(html.match(/Send test/g)).toHaveLength(2);
    expect(html).toContain('aria-label="Delete Security channel"');
    expect(words).toContain("Prism365 never sends through Microsoft.");
  });

  it("lists rules with a switch, what they watch and where they send", () => {
    const words = text(page(overview()));
    expect(words).toContain("High findings");
    expect(words).toContain("New or escalated findings at high or above from Sign-in audit.");
    expect(words).toContain("To Security channel. Quiet for 1 hour after each alert.");
    expect(page(overview())).toContain('aria-label="Turn off High findings"');
  });

  it("shows recent deliveries with their status and why one failed", () => {
    const words = text(page(overview()));
    expect(words).toContain("Failed");
    expect(words).toContain(`${errorText.server_error} 3 attempts.`);
    expect(words).toContain("A deleted destination");
    expect(words).toContain("Test");
    expect(
      deliveryStatus({ ...overview().deliveries[0], status: "pending", attempts: 1 } as never),
    ).toBe("Will try again");
  });

  it("says when a destination's recipients are refused now", () => {
    const value = overview();
    const [hook, mail] = value.destinations;
    const words = text(
      page({
        ...value,
        destinations: [
          hook as AlertOverview["destinations"][number],
          { ...(mail as AlertOverview["destinations"][number]), recipientsAllowed: false },
        ],
      }),
    );
    expect(words).toContain("A recipient is refused by this server now, so nothing is sent here.");
    expect(text(page(value))).not.toContain("refused by this server now");
    expect(errorText.outcome_unknown).toContain("It was not sent again.");
    expect(errorText.dropped).toContain("lowest severity first");
  });

  it("lists rules it keeps but does not run", () => {
    const words = text(
      page(
        overview({
          otherRules: [
            {
              id: "aaaaaaaa-0000-4000-8000-000000000009",
              name: "From a newer version",
              newer: true,
            },
            { id: null, name: null, newer: false },
          ],
        }),
      ),
    );
    expect(words).toContain(
      "From a newer version Made by a newer version of Prism365, so this version keeps it but does not run it. Delete",
    );
    expect(words).toContain(
      "A rule without a name This version cannot read it, so it keeps it but does not run it.",
    );
  });

  it("says when there is nothing yet", () => {
    const words = text(page(overview({ destinations: [], rules: [], deliveries: [] })));
    expect(words).toContain("No destinations yet. Add one below.");
    expect(words).toContain("No rules yet.");
    expect(words).toContain("Add a destination first, then a rule that sends to it.");
    expect(words).toContain("Nothing has been sent yet.");
  });

  it("offers email only when the server has a relay", () => {
    const off = page(
      overview({
        settings: {
          email: { configured: false, domains: [] },
          webhookPorts: [],
          temporaryKey: true,
        },
      }),
    );
    expect(off).toMatch(/<option value="email" disabled="">Email \(not set up on this server\)/);
    expect(text(off)).toContain("stop working after a restart");
    expect(page(overview())).toMatch(/<option value="email">Email<\/option>/);
  });

  it("shows a new webhook's signing secret once, with a copy button", () => {
    const html = renderToStaticMarkup(
      createElement(CreatedSecret, {
        created: {
          destination: overview().destinations[0] as AlertOverview["destinations"][number],
          signingSecret: "synthetic-signing-secret-000000000000000000",
        },
        onDone: noop,
      }),
    );
    expect(html).toContain("synthetic-signing-secret-000000000000000000");
    expect(text(html)).toContain("Prism365 shows it only once.");
    expect(html).toContain('aria-label="Copy the signing secret"');
  });

  it("describes each rule kind in plain words", () => {
    const { choices } = overview();
    const base = {
      id: ruleId,
      name: "Rule",
      enabled: true,
      destinations: [hookId],
      quietMinutes: 60,
      createdAt: "2026-10-24T10:00:00.000Z",
      createdBy: "22222222-2222-4222-8222-222222222222",
    };
    expect(
      ruleSummary({ ...base, kind: "directory", ruleIds: ["role.privileged-assigned"] }, choices),
    ).toBe("Directory changes: a privileged role was assigned.");
    expect(ruleSummary({ ...base, kind: "collection", stoppedHours: 6, gaps: true }, choices)).toBe(
      "When collection stopped for 6 hours, or when a module could not read part of its source.",
    );
    expect(
      ruleSummary(
        {
          ...base,
          kind: "audit",
          auditId: "accounts.inactive",
          tileId: "inactive-accounts",
          condition: "change",
          value: 5,
        },
        choices,
      ),
    ).toBe("Inactive accounts: Inactive accounts moves by 5 or more.");
  });

  it("asks for an audit run before an audit rule can pick a number", () => {
    const html = renderToStaticMarkup(
      createElement(RuleForm, {
        destinations: overview().destinations,
        choices: {
          modules: [],
          audits: [{ id: "accounts.inactive", title: "Inactive accounts", tiles: [] }],
        },
        busy: false,
        onCreate: noop,
      }),
    );
    // The form starts on findings; the audit choice is in the kind list.
    expect(html).toContain('<option value="audit">An audit number</option>');
  });

  it("writes its copy plainly", () => {
    const html = page(overview());
    const table = renderToStaticMarkup(
      createElement(DeliveryTable, {
        deliveries: overview().deliveries,
        destinations: overview().destinations,
        rules: overview().rules,
      }),
    );
    for (const sample of [html, table, ...Object.values(errorText)]) {
      expect(sample).not.toMatch(/[\u2014\u00b7!]/);
    }
  });

  it("shows where the weekly summary and each audit's scheduled runs go", () => {
    const html = page(overview());
    const words = text(html);
    expect(words).toContain("every Monday at 07:00 (America/Chicago)");
    expect(words).toContain("Email with details on gets the CSV attached, up to 5 MB");
    expect(html).toMatch(
      new RegExp(`<span>Weekly summary</span><select>.*?<option value="${hookId}" selected="">`),
    );
    expect(html).toMatch(
      new RegExp(`<span>Inactive accounts</span><select>.*?<option value="${mailId}" selected="">`),
    );
  });

  it("tells the demo that it sends nothing", () => {
    const session: PrismSession = sessionSchema.parse({
      mode: "demo",
      configured: false,
      authenticated: false,
      displayName: null,
      csrfToken: null,
      enabledFeatures: [],
      missingConfiguration: [],
    });
    expect(text(renderToStaticMarkup(createElement(Alerts, { session })))).toContain(
      "The demo sends no alerts.",
    );
  });

  it("tells the public demo build too, whose replayed session says live", () => {
    const session: PrismSession = sessionSchema.parse({
      mode: "live",
      configured: true,
      authenticated: true,
      displayName: "Demo administrator",
      csrfToken: "synthetic-csrf",
      enabledFeatures: [],
      missingConfiguration: [],
    });
    vi.stubEnv("VITE_DEMO", "true");
    try {
      expect(text(renderToStaticMarkup(createElement(Alerts, { session })))).toContain(
        "The demo sends no alerts.",
      );
    } finally {
      vi.unstubAllEnvs();
    }
    expect(text(renderToStaticMarkup(createElement(Alerts, { session })))).not.toContain(
      "The demo sends no alerts.",
    );
  });
});

import { type AttentionItem, type PrismSession, sessionSchema } from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { AttentionCard, emptyInbox, Notifications, SystemAlertRow } from "../src/Attention.tsx";
import { at, expectSkeleton, relativeCell, severityDot } from "./markup.ts";

// Microsoft updates (the notification inbox), rendered without a browser. Invented posts only.

const session: PrismSession = sessionSchema.parse({
  mode: "live",
  configured: true,
  authenticated: true,
  displayName: "Synthetic admin",
  csrfToken: "synthetic-csrf-token-value-000000000000",
  enabledFeatures: [],
  missingConfiguration: [],
});

const item: AttentionItem = {
  id: "11111111-1111-4111-8111-111111111111",
  module: "microsoft.service-health",
  kind: "incident",
  class: "incident",
  ref: "EX000001",
  title: "Some users cannot send mail",
  summary: "Synthetic summary.",
  services: ["Exchange Online"],
  provider: {
    status: "Investigating",
    labels: [],
    deadline: "2026-10-30T00:00:00.000Z",
    updatedAt: "2026-10-22T11:30:00.000Z",
    active: true,
    impact: null,
    workaround: null,
  },
  localLabels: [],
  severity: "high",
  link: "https://admin.example.invalid/incidents/EX000001",
  revision: 2,
  changeNote: "Moved from advisory to incident.",
  firstSeen: "2026-10-22T08:00:00.000Z",
  lastSeen: "2026-10-22T11:30:00.000Z",
  changedAt: "2026-10-22T10:00:00.000Z",
  unread: true,
  review: { state: "new", changedAt: null },
  muted: false,
};

const NOW = "2026-10-22T12:00:00.000Z";
const card = (value: AttentionItem = item) =>
  at(NOW, () =>
    renderToStaticMarkup(
      createElement(AttentionCard, { item: value, session, onChanged: () => {} }),
    ),
  );

describe("Microsoft updates", () => {
  it("shows when Microsoft changed an item, and when this installation noticed, as words", () => {
    const html = card();
    expect(html).toMatch(relativeCell("2026-10-22T10:00:00.000Z", "2 h ago"));
    expect(html).toMatch(relativeCell("2026-10-22T11:30:00.000Z", "30 min ago"));
  });

  it("keeps the date Microsoft wants action by as a date", () => {
    const html = card();
    expect(html).toContain("Action required by ");
    expect(html).not.toContain('dateTime="2026-10-30T00:00:00.000Z"');
  });

  it("marks severity with the shared dot, on the card and on a system alert", () => {
    expect(card()).toContain(severityDot("high", "High"));
    expect(card()).not.toContain("finding-severity");
    const alert = at(NOW, () =>
      renderToStaticMarkup(
        createElement(SystemAlertRow, {
          alert: {
            id: "collection-stopped",
            severity: "critical",
            title: "Collection has stopped",
            detail: "Nothing has been read since the sign-in ended.",
            observedAt: "2026-10-22T11:00:00.000Z",
            action: null,
          },
          onNavigate: () => {},
        }),
      ),
    );
    expect(alert).toContain(severityDot("critical", "Critical"));
    expect(alert).toMatch(relativeCell("2026-10-22T11:00:00.000Z", "1 h ago"));
  });

  it("lists a card's facts one after another with a rule between them, not a punctuation mark", () => {
    const meta = card().match(/<p class="incident__meta meta-line">(.*?)<\/p>/)?.[1] ?? "";
    expect(meta).toContain("<span>EX000001</span>");
    expect(meta).toContain("<span>You: Not reviewed</span>");
    expect(meta).toContain('<a href="https://admin.example.invalid/incidents/EX000001"');
    expect(meta).not.toMatch(/[·—]/);
  });

  it("draws still skeleton rows while the inbox loads", () => {
    expectSkeleton(
      renderToStaticMarkup(
        createElement(Notifications, {
          session,
          onNavigate: () => {},
          onCountsChanged: () => {},
        }),
      ),
      "Loading notifications",
    );
  });

  it("says in one plain line what an empty inbox is waiting for, under every filter", () => {
    for (const [filter, line] of Object.entries(emptyInbox)) {
      expect(line, filter).toMatch(/^No [a-z ]+\. [A-Z].*\.$/);
      expect(line, filter).not.toMatch(/[·—!]/);
    }
    expect(Object.keys(emptyInbox)).toEqual(["unread", "open", "all", "muted"]);
  });
});

import type { AttentionItem, ServiceHealthStatus } from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { ServiceHealthPanel } from "../src/ServiceHealth.tsx";

const incident: AttentionItem = {
  id: "6f1c2b8e-0000-4000-8000-000000000101",
  module: "service.health",
  kind: "service.incident",
  class: "incident",
  ref: "EX400001",
  title: "Some users cannot send mail",
  summary: "Users may be unable to send mail.",
  services: ["Exchange Online"],
  provider: {
    status: "Investigating",
    labels: ["Incident"],
    deadline: null,
    updatedAt: "2026-10-22T11:30:00.000Z",
    active: true,
    impact: "Some users in the tenant may be unable to send mail.",
    workaround: "Resend after a short delay.",
  },
  localLabels: [],
  severity: "high",
  link: "https://admin.microsoft.com/#/servicehealth/:/alerts/EX400001",
  revision: 1,
  changeNote: null,
  firstSeen: "2026-10-22T11:30:00.000Z",
  lastSeen: "2026-10-22T11:30:00.000Z",
  changedAt: "2026-10-22T11:30:00.000Z",
  unread: true,
  review: { state: "new", changedAt: null },
  muted: false,
};

function status(overrides: Partial<ServiceHealthStatus> = {}): ServiceHealthStatus {
  return {
    state: "ok",
    detail: null,
    lastSuccess: "2026-10-22T11:58:00.000Z",
    lastAttempt: "2026-10-22T11:58:00.000Z",
    services: [{ name: "Exchange Online", status: "serviceOperational" }],
    incidents: [],
    messages: { state: "ok", lastSuccess: null },
    ...overrides,
  };
}

const render = (value: ServiceHealthStatus | null) =>
  renderToStaticMarkup(
    createElement(ServiceHealthPanel, { status: value, error: "", onNavigate: () => {} }),
  );

describe("service health panel", () => {
  it("says no issues were reported, never that services are healthy", () => {
    const html = render(status());
    expect(html).toContain("No issues reported by Microsoft");
    expect(html).not.toMatch(/healthy|all systems/i);
  });

  it("shows Microsoft's impact, workaround and source link for an incident", () => {
    const html = render(status({ incidents: [incident] }));
    expect(html).toContain("1 issue reported by Microsoft");
    expect(html).toContain("Microsoft says:");
    expect(html).toContain("Some users in the tenant");
    expect(html).toContain("Workaround from Microsoft");
    expect(html).toContain("https://admin.microsoft.com/#/servicehealth/:/alerts/EX400001");
  });

  it("separates a check that could not run from one that found nothing", () => {
    expect(render(status({ state: "missing_permission", lastSuccess: null }))).toContain(
      "Grant access",
    );
    expect(render(status({ state: "missing_permission" }))).not.toContain("No issues reported");
    expect(
      render(status({ state: "failed", detail: "The last check failed (timeout)." })),
    ).toContain("The last check failed");
    expect(render(status({ state: "never", lastSuccess: null }))).toContain("Not checked yet");
    expect(render(status({ state: "not_enabled", lastSuccess: null }))).toContain(
      "ENABLED_READ_FEATURES",
    );
    expect(render(status({ state: "stale" }))).toContain("Data may be out of date");
  });
});

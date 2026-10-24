import {
  type PrismSession,
  type SetupCheck,
  type SetupStatus,
  sessionSchema,
} from "@prism365/contracts";
import { permissionPacks } from "@prism365/graph";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { supportedReadFeatures } from "../../api/src/config.ts";
import { SetupChecklist, SetupChecklistView } from "../src/Setup.tsx";
import { readPacks, TenantConnection } from "../src/TenantConnection.tsx";

function session(overrides: Partial<PrismSession> = {}): PrismSession {
  return sessionSchema.parse({
    mode: "live",
    configured: true,
    authenticated: true,
    displayName: "Synthetic admin",
    csrfToken: "synthetic-csrf",
    enabledFeatures: ["directory.users.read", "audit.signins.read"],
    missingConfiguration: [],
    ...overrides,
  });
}

const check = (overrides: Partial<SetupCheck>): SetupCheck => ({
  id: "configuration",
  state: "ok",
  title: "Configuration is complete",
  detail: "The tenant, the app registration, the administrator list and a credential are set.",
  fix: null,
  grantFeature: null,
  ...overrides,
});

const status = (overrides: Partial<SetupStatus> = {}): SetupStatus => ({
  mode: "live",
  configured: true,
  missingConfiguration: [],
  checks: [],
  probed: false,
  ...overrides,
});

const render = (
  value: SetupStatus | null,
  options: { session?: PrismSession; error?: string; probing?: boolean } = {},
) =>
  renderToStaticMarkup(
    createElement(SetupChecklistView, {
      session: options.session ?? session(),
      status: value,
      error: options.error ?? "",
      probing: options.probing ?? false,
      onProbe: () => {},
    }),
  );

describe("setup checklist", () => {
  it("is hidden in demo mode", () => {
    expect(
      renderToStaticMarkup(
        createElement(SetupChecklist, { session: session({ mode: "demo", configured: false }) }),
      ),
    ).toBe("");
  });

  it("says it is checking before the status arrives", () => {
    const html = render(null);
    expect(html).toContain("Checking this installation");
    expect(html).not.toContain("Check with Microsoft");
  });

  it("names missing settings on an installation nobody can sign in to yet", () => {
    const html = render(
      status({ configured: false, missingConfiguration: ["ENTRA_CLIENT_ID", "ENTRA_TENANT_ID"] }),
      { session: session({ configured: false, authenticated: false }) },
    );
    expect(html).toContain("Configuration is incomplete");
    expect(html).toContain("Missing: ENTRA_CLIENT_ID, ENTRA_TENANT_ID.");
    expect(html).toContain("npm run setup");
    expect(html).toContain("Not working");
    expect(html).not.toContain("Check with Microsoft");
  });

  it("asks a configured but signed-out visitor to sign in", () => {
    const html = render(status(), { session: session({ authenticated: false }) });
    expect(html).toContain("Sign in with an administrator account");
    expect(html).not.toContain("<li");
  });

  it("shows each check with its state in words, its fix and its consent link", () => {
    const html = render(
      status({
        checks: [
          check({}),
          check({
            id: "certificate",
            state: "fail",
            title: "The certificate has expired",
            detail: "It expired on 2026-10-01.",
            fix: "Run npm run setup -- --tenant <your tenant> --rotate-certificate.",
          }),
          check({
            id: "feature.audit.signins.read",
            state: "warn",
            title: "Sign-in logs (AuditLog.Read.All)",
            detail: "Your Microsoft session has not granted AuditLog.Read.All.",
            fix: "Grant access.",
            grantFeature: "audit.signins.read",
          }),
          check({ id: "collection", state: "unknown", title: "Collection is not available" }),
        ],
      }),
    );
    expect(html).toContain("1 of 4 done; 3 need a look.");
    for (const label of ["Done", "Not working", "Needs attention", "Not known"])
      expect(html).toContain(`>${label}</span>`);
    expect(html).toContain("setup-state--fail");
    expect(html).toContain("--rotate-certificate");
    expect(html).toContain('href="/api/v1/auth/login?feature=audit.signins.read"');
    expect(html.match(/>Grant access<\/a>/g)).toHaveLength(1);
    expect(html).toContain("Check with Microsoft");
    expect(html).toContain("records it in the access ledger");
  });

  it("disables the button while checking and says when Microsoft was asked", () => {
    const probing = render(status({ checks: [check({})] }), { probing: true });
    expect(probing).toContain("Checking with Microsoft");
    expect(probing).toMatch(/<button[^>]*disabled/);
    const probed = render(status({ checks: [check({})], probed: true }));
    expect(probed).toContain("Checked with Microsoft just now");
    expect(probed).toContain("1 of 1 done.");
  });

  it("shows a load failure as an alert", () => {
    const html = render(null, { error: "The request could not be completed." });
    expect(html).toContain('role="alert"');
    expect(html).toContain("The request could not be completed.");
  });
});

describe("read packs on the Microsoft access page", () => {
  it("describes every pack the API accepts, with its catalog label and permission", () => {
    expect(readPacks.map((pack) => pack.id)).toEqual(Object.keys(supportedReadFeatures));
    for (const pack of readPacks) {
      const entry = permissionPacks.find((item) => item.id === pack.id);
      expect(pack.label, pack.id).toBe(entry?.label);
      expect(pack.scope, pack.id).toBe(entry?.permission);
    }
  });

  it("lists the enabled packs, including newer ones, with their access links", () => {
    const html = renderToStaticMarkup(
      createElement(TenantConnection, {
        session: session({
          enabledFeatures: ["directory.users.read", "audit.directory.read", "future.pack.read"],
        }),
        error: "",
        onRetry: () => {},
        onLogout: () => {},
        loggingOut: false,
      }),
    );
    expect(html).toContain("Directory audit log");
    expect(html).toContain('href="/api/v1/auth/login?feature=audit.directory.read"');
    expect(html).toContain("future.pack.read");
    expect(html).toContain(
      `Read packs this installation has not enabled (${readPacks.length - 2})`,
    );
    expect(html).not.toContain("feature=service.health.read");
  });
});

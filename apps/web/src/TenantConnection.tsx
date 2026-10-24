import type { PrismSession } from "@prism365/contracts";
import { LockIcon, ShieldIcon } from "./icons.tsx";

// Every read pack an installation can enable (ENABLED_READ_FEATURES, apps/api/src/config.ts), in
// the API's order, with the label and permission of its entry in packages/graph's catalog. A test
// keeps this list in step with both.
export const readPacks: readonly { id: string; label: string; scope: string; detail: string }[] = [
  {
    id: "directory.users.read",
    label: "User inventory",
    scope: "User.Read.All",
    detail: "Selected directory fields for each account, its manager and last sign-in.",
  },
  {
    id: "user.licenses.read",
    label: "User license details",
    scope: "LicenseAssignment.Read.All",
    detail: "The licenses and service plans assigned to a user.",
  },
  {
    id: "tenant.licenses.read",
    label: "Tenant license inventory",
    scope: "LicenseAssignment.Read.All",
    detail: "Products and seats the tenant has bought, for the license capacity audit.",
  },
  {
    id: "audit.signins.read",
    label: "Sign-in logs",
    scope: "AuditLog.Read.All",
    detail: "Sign-ins for the sign-in audit and the People view, subject to role and license.",
  },
  {
    id: "reports.authentication-methods.read",
    label: "Authentication method registration",
    scope: "AuditLog.Read.All",
    detail: "Which sign-in methods each account has registered, for the MFA audit.",
  },
  {
    id: "service.health.read",
    label: "Service health",
    scope: "ServiceHealth.Read.All",
    detail: "Microsoft's incidents and advisories for the services this tenant has.",
  },
  {
    id: "service.messages.read",
    label: "Message center",
    scope: "ServiceMessage.Read.All",
    detail: "Planned changes, retirements and actions Microsoft asks administrators to take.",
  },
  {
    id: "audit.directory.read",
    label: "Directory audit log",
    scope: "AuditLog.Read.All",
    detail: "Who changed which user, group, role, app or policy, for the change feed.",
  },
  {
    id: "identity.conditional-access.read",
    label: "Conditional Access policies",
    scope: "Policy.Read.All",
    detail: "Conditional Access policy definitions.",
  },
  {
    id: "identity.risk.read",
    label: "Identity Protection risk",
    scope: "IdentityRiskyUser.Read.All",
    detail: "Risky-user detections. Microsoft needs Entra ID P2 for them.",
  },
  {
    id: "devices.intune.read",
    label: "Intune managed devices",
    scope: "DeviceManagementManagedDevices.Read.All",
    detail: "Managed devices and their compliance state.",
  },
  {
    id: "directory.applications.read",
    label: "Enterprise applications",
    scope: "Application.Read.All",
    detail: "Enterprise applications and the application permissions granted to them.",
  },
  {
    id: "directory.permission-grants.read",
    label: "Delegated permission grants",
    scope: "Directory.Read.All",
    detail: "Which delegated permissions users and administrators consented to, per application.",
  },
  {
    id: "reports.usage.read",
    label: "Microsoft 365 usage reports",
    scope: "Reports.Read.All",
    detail: "Per-user Microsoft 365 apps usage.",
  },
  {
    id: "reports.settings.read",
    label: "Usage-report privacy setting",
    scope: "ReportSettings.Read.All",
    detail: "Whether the tenant conceals names in usage reports.",
  },
];

export function TenantConnection({
  session,
  error,
  onRetry,
  onLogout,
  loggingOut,
}: {
  session: PrismSession | null;
  error: string;
  onRetry: () => void;
  onLogout: () => void;
  loggingOut: boolean;
}) {
  if (!session) {
    return (
      <section className="connection-panel" aria-label="Connection status">
        <ShieldIcon />
        <div>
          <h2>{error ? "The application API is unavailable." : "Checking your workspace…"}</h2>
          <p>{error || "Prism365 is checking the connection mode before displaying data."}</p>
          {error ? (
            <button className="button button--secondary" type="button" onClick={onRetry}>
              Try again
            </button>
          ) : null}
        </div>
      </section>
    );
  }

  if (session.mode === "demo") {
    return (
      <section className="demo-banner" aria-label="Demo data notice">
        <div className="demo-banner__icon">
          <ShieldIcon />
        </div>
        <div>
          <strong>Exploring with synthetic data</strong>
          <span>
            No Microsoft 365 tenant is connected. Every identity, event, and metric is fictional.
          </span>
        </div>
        <span className="source-label">Local demonstration</span>
      </section>
    );
  }

  // The enabled packs first, in the API's order, including any this page has no description for.
  const known = new Set(readPacks.map((pack) => pack.id));
  const enabledPacks = [
    ...readPacks.filter((pack) => session.enabledFeatures.includes(pack.id)),
    ...session.enabledFeatures
      .filter((id) => !known.has(id))
      .map((id) => ({
        id,
        label: id,
        scope: "",
        detail: "A read pack this page does not describe yet.",
      })),
  ];
  const otherPacks = readPacks.filter((pack) => !session.enabledFeatures.includes(pack.id));

  return (
    <section className="tenant-connection" aria-labelledby="connection-title">
      <div className="tenant-connection__header">
        <div>
          <span className="eyebrow">Tenant connection</span>
          <h2 id="connection-title">
            {session.authenticated
              ? "Your Microsoft session is connected."
              : "Connect the permissions your work needs."}
          </h2>
          <p>
            {session.authenticated
              ? `Signed in as ${session.displayName ?? "an authorized administrator"}.`
              : "This installation uses your organization’s own Microsoft app registration."}
          </p>
        </div>
        {session.authenticated ? (
          <button
            className="button button--secondary"
            onClick={onLogout}
            disabled={loggingOut}
            type="button"
          >
            {loggingOut ? "Signing out…" : "Sign out"}
          </button>
        ) : session.configured && session.enabledFeatures.includes("directory.users.read") ? (
          <a
            className="button button--primary"
            href="/api/v1/auth/login?feature=directory.users.read"
          >
            <LockIcon /> Sign in with Microsoft
          </a>
        ) : null}
      </div>
      {error ? (
        <p className="form-error" role="alert">
          {error}
        </p>
      ) : null}
      {!session.configured ? (
        <div className="setup-needed">
          <strong>Complete installation setup</strong>
          <p>
            The operator must configure a work or school tenant, its application registration, and
            the exact return address before Microsoft sign-in can start. On the server,{" "}
            <code>npm run setup</code> does all of it (docs/SETUP.md).
          </p>
          <ul>
            {session.missingConfiguration.map((item) => (
              <li key={item}>{item}</li>
            ))}
          </ul>
          <p>
            Keep registration credentials on the server. Personal Microsoft accounts are not
            supported.
          </p>
        </div>
      ) : null}
      {enabledPacks.length > 0 ? (
        <div className="connection-features">
          {enabledPacks.map((pack) => (
            <PackCard key={pack.id} pack={pack} enabled session={session} />
          ))}
        </div>
      ) : (
        <p className="connection-none">No read packs are enabled in this installation.</p>
      )}
      {otherPacks.length > 0 ? (
        <details className="connection-more">
          <summary>Read packs this installation has not enabled ({otherPacks.length})</summary>
          <div className="connection-features">
            {otherPacks.map((pack) => (
              <PackCard key={pack.id} pack={pack} enabled={false} session={session} />
            ))}
          </div>
        </details>
      ) : null}
      <p className="connection-footnote">
        Microsoft access also depends on administrator roles, licenses, and service availability.
        Each source reports its own access result. This connection grants no administration write
        actions.
      </p>
    </section>
  );
}

function PackCard({
  pack,
  enabled,
  session,
}: {
  pack: (typeof readPacks)[number];
  enabled: boolean;
  session: PrismSession;
}) {
  return (
    <article>
      <strong>{pack.label}</strong>
      <p>{pack.detail}</p>
      {pack.scope ? <code>{pack.scope}</code> : null}
      <span className="source-label">
        {enabled ? "Enabled in this installation" : "Not enabled in this installation"}
      </span>
      {enabled && session.configured && session.authenticated ? (
        <a href={`/api/v1/auth/login?feature=${encodeURIComponent(pack.id)}`}>
          Review Microsoft access
        </a>
      ) : null}
    </article>
  );
}

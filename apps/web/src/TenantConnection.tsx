import type { PrismSession } from "@prism365/contracts";
import { LockIcon, ShieldIcon } from "./icons.tsx";

const featureDescriptions = [
  {
    id: "directory.users.read",
    label: "Directory profiles",
    scope: "User.Read.All",
    detail: "Look up a user’s selected directory fields.",
  },
  {
    id: "user.licenses.read",
    label: "User license details",
    scope: "LicenseAssignment.Read.All",
    detail: "Show the licenses and service plans assigned to a user.",
  },
  {
    id: "audit.signins.read",
    label: "Sign-in evidence",
    scope: "AuditLog.Read.All",
    detail: "Read recent sign-ins, subject to role and license availability.",
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
            the exact return address before Microsoft sign-in can start.
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
      <div className="connection-features">
        {featureDescriptions.map((feature) => {
          const enabled = session.enabledFeatures.includes(feature.id);
          return (
            <article key={feature.id}>
              <strong>{feature.label}</strong>
              <p>{feature.detail}</p>
              <code>{feature.scope}</code>
              <span className="source-label">
                {enabled ? "Enabled in this installation" : "Not enabled in this installation"}
              </span>
              {enabled && session.configured && session.authenticated ? (
                <a href={`/api/v1/auth/login?feature=${encodeURIComponent(feature.id)}`}>
                  Review Microsoft access
                </a>
              ) : null}
            </article>
          );
        })}
      </div>
      <p className="connection-footnote">
        Microsoft access also depends on administrator roles, licenses, and service availability.
        Each source reports its own access result. This connection grants no administration write
        actions.
      </p>
    </section>
  );
}

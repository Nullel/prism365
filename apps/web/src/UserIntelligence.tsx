import {
  userInvestigationSchema,
  type CapabilityState,
  type PrismSession,
  type UserInvestigation,
} from "@prism365/contracts";
import { useEffect, useId, useRef, useState } from "react";
import { SearchIcon, ShieldIcon, UsersIcon } from "./icons.tsx";
import { AiInvestigation } from "./AiInvestigation.tsx";

const stateLabels: Record<CapabilityState, string> = {
  available: "Available",
  missingConsent: "Permission needed",
  missingRole: "Administrator role needed",
  missingLicense: "Microsoft license needed",
  unsupportedCloud: "Cloud unavailable",
  notEnabled: "Feature not enabled",
  notImplemented: "Coming later",
  providerUnavailable: "Microsoft unavailable",
  unknown: "Could not verify access",
};

function formatDate(value: string) {
  return new Intl.DateTimeFormat(undefined, {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(new Date(value));
}

function AccessState({
  state,
  message,
  feature,
  session,
}: {
  state: CapabilityState;
  message: string | undefined;
  feature: string;
  session: PrismSession;
}) {
  if (state === "available") return null;
  return (
    <div className="source-state" role="status">
      <ShieldIcon />
      <div>
        <strong>{stateLabels[state]}</strong>
        <p>{message ?? "This source could not provide data for this investigation."}</p>
        {state === "missingConsent" && session.mode === "live" ? (
          <a href={`/api/v1/auth/login?feature=${encodeURIComponent(feature)}`}>
            Review this feature’s access
          </a>
        ) : null}
      </div>
    </div>
  );
}

export function UserIntelligence({ session }: { session: PrismSession }) {
  const inputId = useId();
  const [email, setEmail] = useState(session.mode === "demo" ? "alex.rivera@example.invalid" : "");
  const [result, setResult] = useState<UserInvestigation | null>(null);
  const [state, setState] = useState<"idle" | "loading" | "error" | "ready">("idle");
  const [error, setError] = useState("");
  const pending = useRef<AbortController | null>(null);
  const requestNumber = useRef(0);

  useEffect(() => () => pending.current?.abort(), []);

  function changeEmail(value: string) {
    requestNumber.current += 1;
    pending.current?.abort();
    setEmail(value);
    setResult(null);
    setError("");
    setState("idle");
  }

  async function investigate(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const number = ++requestNumber.current;
    pending.current?.abort();
    const controller = new AbortController();
    pending.current = controller;
    setState("loading");
    setResult(null);
    setError("");
    try {
      const response = await fetch(
        `/api/v1/investigations/user?email=${encodeURIComponent(email.trim())}`,
        { signal: controller.signal, credentials: "same-origin", cache: "no-store" },
      );
      const body: unknown = await response.json();
      if (!response.ok) {
        const message =
          typeof body === "object" &&
          body !== null &&
          "message" in body &&
          typeof body.message === "string"
            ? body.message
            : `Investigation could not load (HTTP ${response.status}).`;
        throw new Error(message);
      }
      const data = userInvestigationSchema.parse(body);
      if (data.mode !== session.mode) throw new Error("The data mode changed. Refresh this page.");
      if (number !== requestNumber.current) return;
      setResult(data);
      setState("ready");
    } catch (caught) {
      if (controller.signal.aborted || number !== requestNumber.current) return;
      setError(caught instanceof Error ? caught.message : "Unable to load this investigation.");
      setState("error");
    }
  }

  const canInvestigate = session.mode === "demo" || (session.configured && session.authenticated);
  const profile = result?.profile.data;

  return (
    <section className="investigation" aria-labelledby="investigation-title">
      <div className="investigation__intro">
        <span className="eyebrow">User Intelligence</span>
        <h2 id="investigation-title">Start with a work email.</h2>
        <p>Bring a user’s profile, assigned licenses, and recent sign-in evidence into one view.</p>
      </div>
      <form className="investigation-search" onSubmit={investigate}>
        <label className="field" htmlFor={inputId}>
          <span>User principal name or work email</span>
          <input
            id={inputId}
            type="email"
            required
            maxLength={254}
            autoComplete="off"
            value={email}
            onChange={(event) => changeEmail(event.target.value)}
            placeholder="person@company.com"
            disabled={!canInvestigate}
          />
        </label>
        <button
          className="button button--primary"
          type="submit"
          disabled={!canInvestigate || state === "loading"}
        >
          <SearchIcon /> {state === "loading" ? "Loading evidence…" : "Investigate user"}
        </button>
      </form>
      {!canInvestigate ? (
        <p className="source-notice">
          Complete tenant setup and sign in above to open a live investigation.
        </p>
      ) : null}
      {error ? (
        <p className="form-error investigation-error" role="alert">
          {error}
        </p>
      ) : null}
      <div aria-live="polite" aria-busy={state === "loading"}>
        {result ? (
          <>
            <div className="investigation-meta">
              <span className={`mode-chip mode-chip--${result.mode}`}>
                {result.mode === "demo" ? "Synthetic demonstration" : "Live Microsoft evidence"}
              </span>
              <span>Retrieved {formatDate(result.generatedAt)}</span>
              <span>{result.targetEmail}</span>
            </div>
            <div className="investigation-grid">
              <section className="panel profile-panel" aria-labelledby="profile-title">
                <div className="panel__header">
                  <h3 id="profile-title">Directory profile</h3>
                  <UsersIcon />
                </div>
                <AccessState
                  state={result.profile.state}
                  message={result.profile.message}
                  feature="directory.users.read"
                  session={session}
                />
                {profile ? (
                  <>
                    <div className="profile-identity">
                      <span>{profile.displayName.slice(0, 1)}</span>
                      <div>
                        <h4>{profile.displayName}</h4>
                        <p>{profile.userPrincipalName}</p>
                      </div>
                    </div>
                    <dl className="profile-details">
                      <div>
                        <dt>Department</dt>
                        <dd>{profile.department ?? "Not provided"}</dd>
                      </div>
                      <div>
                        <dt>Job title</dt>
                        <dd>{profile.jobTitle ?? "Not provided"}</dd>
                      </div>
                      <div>
                        <dt>Account type</dt>
                        <dd>{profile.userType ?? "Not provided"}</dd>
                      </div>
                      <div>
                        <dt>Mail address</dt>
                        <dd>{profile.mail ?? "Not provided"}</dd>
                      </div>
                    </dl>
                  </>
                ) : null}
              </section>
              <section className="panel licenses-panel" aria-labelledby="licenses-title">
                <div className="panel__header">
                  <h3 id="licenses-title">Assigned licenses</h3>
                  <span className="source-label">{stateLabels[result.licenses.state]}</span>
                </div>
                <AccessState
                  state={result.licenses.state}
                  message={result.licenses.message}
                  feature="user.licenses.read"
                  session={session}
                />
                {result.licenses.state === "available" && result.licenses.data.length === 0 ? (
                  <p className="empty-data">
                    No direct license details were returned for this user.
                  </p>
                ) : null}
                <ul className="license-list">
                  {result.licenses.data.map((license) => (
                    <li key={license.id}>
                      <strong>{license.skuPartNumber}</strong>
                      <details>
                        <summary>{license.servicePlans.length} service plans</summary>
                        <ul>
                          {license.servicePlans.map((plan) => (
                            <li key={plan.servicePlanId}>
                              <span>{plan.servicePlanName}</span>
                              <span>{plan.provisioningStatus}</span>
                            </li>
                          ))}
                        </ul>
                      </details>
                    </li>
                  ))}
                </ul>
              </section>
            </div>
            <section className="panel signins-panel" aria-labelledby="signins-title">
              <div className="panel__header">
                <div>
                  <span className="eyebrow">Authentication evidence</span>
                  <h3 id="signins-title">Recent sign-ins</h3>
                </div>
                <span className="source-label">{stateLabels[result.signIns.state]}</span>
              </div>
              <AccessState
                state={result.signIns.state}
                message={result.signIns.message}
                feature="audit.signins.read"
                session={session}
              />
              <p className="source-notice">
                {formatDate(result.signIns.window.start)} – {formatDate(result.signIns.window.end)}
                {!result.signIns.complete
                  ? " · Partial coverage: this bounded result is not a complete audit history."
                  : ""}
              </p>
              {result.signIns.state === "available" && result.signIns.data.length === 0 ? (
                <p className="empty-data">
                  No sign-ins were returned in this time window. This does not establish inactivity.
                </p>
              ) : null}
              <ol className="signin-timeline">
                {result.signIns.data.map((event) => (
                  <li key={event.id}>
                    <span
                      className={`signin-outcome ${event.status.errorCode === 0 ? "signin-outcome--success" : "signin-outcome--failure"}`}
                    >
                      {event.status.errorCode === 0 ? "Succeeded" : "Failed"}
                    </span>
                    <div>
                      <strong>{event.appDisplayName}</strong>
                      <time dateTime={event.createdDateTime}>
                        {formatDate(event.createdDateTime)}
                      </time>
                      <p>
                        {event.clientAppUsed} ·{" "}
                        {[event.location.city, event.location.countryOrRegion]
                          .filter(Boolean)
                          .join(", ") || "Location unavailable"}
                      </p>
                      {event.status.failureReason ? <p>{event.status.failureReason}</p> : null}
                      {event.status.additionalDetails ? (
                        <p>{event.status.additionalDetails}</p>
                      ) : null}
                    </div>
                  </li>
                ))}
              </ol>
            </section>
            <AiInvestigation
              key={`${result.targetEmail}-${result.generatedAt}`}
              investigation={result}
              session={session}
            />
            <aside className="investigation-limitations" aria-labelledby="limitations-title">
              <h3 id="limitations-title">Evidence context</h3>
              <ul>
                {result.limitations.map((item) => (
                  <li key={item}>{item}</li>
                ))}
              </ul>
              <p>
                Sign-in and activity evidence describe system events. They do not measure employee
                productivity, effort, or performance.
              </p>
            </aside>
          </>
        ) : state !== "loading" ? (
          <div className="investigation-empty">
            <UsersIcon />
            <h3>A focused view of one account.</h3>
            <p>
              Profile, licenses, and audit access are checked independently, so a missing permission
              does not hide the sources you can read.
            </p>
          </div>
        ) : (
          <p className="source-notice" role="status">
            Retrieving the user’s available evidence…
          </p>
        )}
      </div>
    </section>
  );
}

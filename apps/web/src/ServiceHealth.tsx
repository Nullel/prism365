import {
  type AttentionItem,
  type PrismSession,
  type ServiceHealthStatus,
  serviceHealthStatusSchema,
} from "@prism365/contracts";
import { useCallback, useEffect, useState } from "react";
import { day, readJson, when } from "./common.tsx";

// "Is Microsoft down?": what Microsoft reports for this tenant's services. The panel always says
// whether the check itself worked; an empty list only ever means "No issues reported by
// Microsoft", never that a service is healthy.

export function useServiceHealth(session: PrismSession) {
  const [status, setStatus] = useState<ServiceHealthStatus | null>(null);
  const [error, setError] = useState("");
  const load = useCallback(
    async (signal?: AbortSignal) => {
      if (session.mode !== "live" || !session.authenticated) return;
      try {
        setStatus(
          await readJson(
            "/api/v1/modules/service.health/status",
            (value) => serviceHealthStatusSchema.parse(value),
            signal ? { signal } : {},
          ),
        );
        setError("");
      } catch (loadError) {
        if (signal?.aborted) return;
        setError(loadError instanceof Error ? loadError.message : "Service health could not load.");
      }
    },
    [session.mode, session.authenticated],
  );
  useEffect(() => {
    const controller = new AbortController();
    void load(controller.signal);
    const timer = setInterval(() => void load(), 60_000);
    return () => {
      controller.abort();
      clearInterval(timer);
    };
  }, [load]);
  return { status, error, reload: load };
}

export function IncidentCard({
  item,
  compact = false,
}: {
  item: AttentionItem;
  compact?: boolean;
}) {
  return (
    <article className={`incident${item.provider.active ? "" : " incident--resolved"}`}>
      <div className="incident__topline">
        <strong>{item.services.join(", ")}</strong>
        {item.provider.labels.slice(0, 1).map((label) => (
          <span
            key={label}
            className={`finding-severity finding-severity--${item.class === "incident" ? "high" : "medium"}`}
          >
            {label}
          </span>
        ))}
        <span className="incident__status">{item.provider.status ?? "Status not reported"}</span>
        {item.localLabels.map((label) => (
          <span key={label} className="person-tag" title="Added by Prism365, not Microsoft">
            {label}
          </span>
        ))}
      </div>
      <h3>{item.title}</h3>
      {item.provider.impact ? (
        <p>
          <span className="incident__label">Microsoft says: </span>
          {item.provider.impact}
        </p>
      ) : (
        <p>{item.summary}</p>
      )}
      {!compact && item.provider.workaround ? (
        <p>
          <span className="incident__label">Workaround from Microsoft: </span>
          {item.provider.workaround}
        </p>
      ) : null}
      <p className="incident__meta">
        {item.ref} · Microsoft updated {when(item.provider.updatedAt)}
        {item.changeNote ? ` · ${item.changeNote}` : ""}
        {item.link ? (
          <>
            {" · "}
            <a href={item.link} target="_blank" rel="noreferrer">
              Open in Microsoft 365 admin center
            </a>
          </>
        ) : null}
      </p>
    </article>
  );
}

const stateCopy: Record<ServiceHealthStatus["state"], string> = {
  ok: "",
  partial: "The last check was incomplete.",
  stale: "Data may be out of date.",
  failed: "The last check failed.",
  never: "Not checked yet.",
  missing_permission: "Prism365 cannot check service health.",
  not_enabled: "Service health is not set up in this installation.",
};

export function ServiceHealthPanel({
  status,
  error,
  onNavigate,
}: {
  status: ServiceHealthStatus | null;
  error: string;
  onNavigate: (view: "signins" | "access") => void;
}) {
  const active = status?.incidents.filter((item) => item.provider.active) ?? [];
  const resolved = status?.incidents.filter((item) => !item.provider.active) ?? [];
  const checked =
    status?.state === "ok" || status?.state === "stale" || status?.state === "partial";
  return (
    <section className="panel health-panel" aria-labelledby="health-title">
      <div className="health-panel__header">
        <div>
          <h2 id="health-title">Is Microsoft down?</h2>
          <p>
            {!status
              ? error || "Checking…"
              : status.lastSuccess
                ? `Microsoft service health, last checked ${when(status.lastSuccess)}.`
                : "Microsoft service health for the services this tenant has."}
          </p>
        </div>
        {checked ? (
          <span className={`health-verdict health-verdict--${active.length ? "issues" : "clear"}`}>
            {active.length
              ? `${active.length} ${active.length === 1 ? "issue" : "issues"} reported by Microsoft`
              : "No issues reported by Microsoft"}
          </span>
        ) : null}
      </div>

      {status && status.state !== "ok" ? (
        <div className={`health-state health-state--${status.state}`} role="status">
          <strong>{stateCopy[status.state]}</strong> {status.detail}
          {status.state === "missing_permission" ? (
            <>
              {" "}
              <a href="/api/v1/auth/login?feature=service.health.read">Grant access</a>
            </>
          ) : status.state === "not_enabled" ? (
            <>
              {" "}
              Add <code>service.health.read</code> to <code>ENABLED_READ_FEATURES</code> and grant
              ServiceHealth.Read.All in the app registration.{" "}
              <button className="link-button" type="button" onClick={() => onNavigate("access")}>
                Microsoft access
              </button>
            </>
          ) : status.state === "never" ? (
            <>
              {" "}
              <button className="link-button" type="button" onClick={() => onNavigate("signins")}>
                Start collecting
              </button>
            </>
          ) : null}
        </div>
      ) : null}

      {active.length > 0 ? (
        <div className="incident-list">
          {active.map((item) => (
            <IncidentCard key={item.id} item={item} />
          ))}
        </div>
      ) : checked && status?.services.length ? (
        <p className="empty-data">
          No issues reported by Microsoft for the {status.services.length} services it lists for
          this tenant. That is what Microsoft has published, not a guarantee every service works.
        </p>
      ) : null}

      {resolved.length > 0 ? (
        <details className="health-resolved">
          <summary>Resolved in the last day ({resolved.length})</summary>
          <div className="incident-list">
            {resolved.map((item) => (
              <IncidentCard key={item.id} item={item} compact />
            ))}
          </div>
        </details>
      ) : null}

      {status && status.services.length > 0 ? (
        <details className="health-services">
          <summary>Services Microsoft reports for this tenant ({status.services.length})</summary>
          <ul>
            {status.services.map((service) => (
              <li key={service.name}>
                <span>{service.name}</span>
                <span className="person-muted">
                  {service.status.replace(/([a-z])([A-Z])/g, "$1 $2")}
                </span>
              </li>
            ))}
          </ul>
        </details>
      ) : null}
      {status?.lastAttempt && status.lastSuccess && status.lastAttempt !== status.lastSuccess ? (
        <p className="signin-coverage">Last attempt {day(status.lastAttempt)}.</p>
      ) : null}
    </section>
  );
}

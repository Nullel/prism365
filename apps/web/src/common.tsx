import type { Finding, FindingState, PrismSession } from "@prism365/contracts";
import { LockIcon } from "./icons.tsx";

// Helpers and the finding card shared by the sign-in audit and people views.

export const stateLabels: Record<FindingState, string> = {
  open: "Open",
  acknowledged: "Acknowledged",
  resolved: "Resolved",
};

const regionNames = (() => {
  try {
    return new Intl.DisplayNames(undefined, { type: "region" });
  } catch {
    return null;
  }
})();

export function countryName(code: string | null): string {
  if (!code || code === "??") return "Not reported";
  try {
    return regionNames?.of(code) ?? code;
  } catch {
    return code;
  }
}

export function place(parts: {
  city: string | null;
  state: string | null;
  country: string | null;
}) {
  const named = [parts.city, parts.state].filter(Boolean).join(", ");
  const country = parts.country ? countryName(parts.country) : "";
  return [named, country].filter(Boolean).join(", ") || "Not reported";
}

export function when(value: string | null): string {
  if (!value) return "never";
  return new Intl.DateTimeFormat(undefined, { dateStyle: "medium", timeStyle: "short" }).format(
    new Date(value),
  );
}

export function day(value: string | null): string {
  if (!value) return "never";
  return new Intl.DateTimeFormat(undefined, { dateStyle: "medium" }).format(new Date(value));
}

export const number = new Intl.NumberFormat();

export async function readJson<T>(
  url: string,
  parse: (value: unknown) => T,
  init?: RequestInit,
): Promise<T> {
  const response = await fetch(url, { credentials: "same-origin", cache: "no-store", ...init });
  if (!response.ok) {
    const body = (await response.json().catch(() => null)) as { message?: string } | null;
    throw new Error(body?.message ?? `Request failed with HTTP ${response.status}.`);
  }
  return parse(await response.json());
}

export function post(session: PrismSession, body?: unknown): RequestInit {
  return {
    method: "POST",
    headers: {
      "x-csrf-token": session.csrfToken ?? "",
      ...(body === undefined ? {} : { "Content-Type": "application/json" }),
    },
    ...(body === undefined ? {} : { body: JSON.stringify(body) }),
  };
}

function Evidence({ finding }: { finding: Finding }) {
  return (
    <details className="finding-details">
      <summary>Evidence and limitations</summary>
      {finding.evidence.map((item) => (
        <dl className="finding-evidence" key={item.ref}>
          <div>
            <dt>Observed</dt>
            <dd>{when(item.observedAt)}</dd>
          </div>
          {Object.entries(item.fields)
            .filter(([, value]) => value !== null && value !== "")
            .map(([key, value]) => (
              <div key={key}>
                <dt>{key}</dt>
                <dd>{String(value)}</dd>
              </div>
            ))}
          <div>
            <dt>Source</dt>
            <dd>{item.source}</dd>
          </div>
        </dl>
      ))}
      <ul className="finding-limitations">
        {finding.limitations.map((limitation) => (
          <li key={limitation}>{limitation}</li>
        ))}
      </ul>
    </details>
  );
}

export function FindingItem({
  finding,
  busy,
  onState,
  onOpenPerson,
  showSubject = true,
}: {
  finding: Finding;
  busy: boolean;
  // Off on a person's own page, where the subject is the page.
  showSubject?: boolean;
  onState: (state: FindingState) => void;
  onOpenPerson?: (userId: string) => void;
}) {
  const subject = finding.subject.label ?? finding.subject.id;
  return (
    <article className="finding">
      <div className="finding__topline">
        <span className={`finding-severity finding-severity--${finding.severity}`}>
          {finding.severity}
        </span>
        {!showSubject ? null : finding.subject.kind === "person" && onOpenPerson ? (
          <button
            className="finding__subject finding__subject--link"
            type="button"
            onClick={() => onOpenPerson(finding.subject.id)}
          >
            {subject}
          </button>
        ) : (
          <span className="finding__subject">{subject}</span>
        )}
        {finding.state !== "open" ? (
          <span className="person-tag">{stateLabels[finding.state]}</span>
        ) : null}
      </div>
      <h3>{finding.title}</h3>
      <p>{finding.summary}</p>
      <p className="finding__meta">
        Last seen {when(finding.lastSeen)} · first seen {when(finding.firstSeen)} ·{" "}
        {finding.occurrences} {finding.occurrences === 1 ? "occurrence" : "occurrences"}
      </p>
      <Evidence finding={finding} />
      <div className="finding__actions">
        {finding.state !== "acknowledged" && finding.state !== "resolved" ? (
          <button
            className="button button--secondary"
            type="button"
            disabled={busy}
            onClick={() => onState("acknowledged")}
          >
            Acknowledge
          </button>
        ) : null}
        {finding.state !== "resolved" ? (
          <button
            className="button button--primary"
            type="button"
            disabled={busy}
            onClick={() => onState("resolved")}
          >
            <LockIcon /> Resolve
          </button>
        ) : (
          <button
            className="button button--secondary"
            type="button"
            disabled={busy}
            onClick={() => onState("open")}
          >
            Reopen
          </button>
        )}
      </div>
      {finding.ruleId === "signins.unfamiliar_signin" && finding.state !== "resolved" ? (
        <p className="finding__hint">
          Resolving tells Prism365 this place and device are normal for the account. Acknowledging
          keeps them unfamiliar while you look into it.
        </p>
      ) : null}
    </article>
  );
}

import type { Finding, FindingState, PrismSession, Severity } from "@prism365/contracts";
import { CopyButton } from "./CopyButton.tsx";
import { LockIcon } from "./icons.tsx";

// Helpers, the finding card and the account name shared by the sign-in audit, audits and people
// views.

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

const MINUTE = 60_000;
const HOUR = 60 * MINUTE;
const DAY = 24 * HOUR;
// A timestamp this far ahead of now is a clock that disagrees with the server by a little, and
// reads as "just now". Further ahead than this it is shown as its date, never as a time ago.
const CLOCK_SKEW = 5 * MINUTE;

// How long ago something happened, the way a mail client says it: "just now", "5 min ago",
// "3 h ago", "yesterday", "4 days ago", and from a week on, the date. Whole units are rounded
// down, so "59 min ago" never turns into "60 min ago" a moment early.
export function relativeTime(value: string | null, now: number = Date.now()): string {
  if (!value) return "never";
  const at = Date.parse(value);
  if (Number.isNaN(at)) return "unknown";
  const elapsed = now - at;
  if (elapsed < -CLOCK_SKEW) return day(value);
  if (elapsed < MINUTE) return "just now";
  if (elapsed < HOUR) return `${Math.floor(elapsed / MINUTE)} min ago`;
  if (elapsed < DAY) return `${Math.floor(elapsed / HOUR)} h ago`;
  const days = Math.floor(elapsed / DAY);
  if (days === 1) return "yesterday";
  if (days < 7) return `${days} days ago`;
  return day(value);
}

// The time as words, with the exact moment on hover and in the markup. `now` is for tests.
export function RelativeTime({ value, now }: { value: string | null; now?: number }) {
  if (!value || Number.isNaN(Date.parse(value)))
    return <span className="person-muted">{relativeTime(value, now)}</span>;
  return (
    <time dateTime={value} title={when(value)}>
      {relativeTime(value, now)}
    </time>
  );
}

const severityNames: Record<Severity, string> = {
  critical: "Critical",
  high: "High",
  medium: "Medium",
  low: "Low",
  info: "Info",
};

// The one way severity looks outside a finding's own card: a dot in the severity's colour and
// its name, so the colour never carries the meaning alone. With a count it reads "2 critical",
// for a tile. Info shares low's colour.
export function SeverityDot({ severity, count }: { severity: Severity; count?: number }) {
  const name = severityNames[severity];
  return (
    <span className={`sev sev--${severity === "info" ? "low" : severity}`}>
      <span className="sev__dot" aria-hidden="true" />
      {count === undefined ? name : `${number.format(count)} ${name.toLowerCase()}`}
    </span>
  );
}

// Keys for loading rows. Fixed, because the rows are identical and never reorder.
const SKELETON_KEYS = Array.from({ length: 20 }, (_, position) => `skeleton-${position + 1}`);

// Grey bars where rows will be, drawn without animation: the loading state of a list or table.
// The label is for screen readers, which would otherwise hear nothing. Spans throughout, so it can
// stand inside a button as well as in a panel.
export function SkeletonRows({ rows = 4, label }: { rows?: number; label: string }) {
  return (
    <span className="skeleton" aria-busy="true">
      <span className="visually-hidden">{label}</span>
      {SKELETON_KEYS.slice(0, Math.min(rows, SKELETON_KEYS.length)).map((key) => (
        <span className="skeleton__row" key={key} aria-hidden="true" />
      ))}
    </span>
  );
}

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
        <SeverityDot severity={finding.severity} />
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
      <p className="finding__meta meta-line">
        <span>
          Last seen <RelativeTime value={finding.lastSeen} />
        </span>
        <span>
          First seen <RelativeTime value={finding.firstSeen} />
        </span>
        <span>
          {number.format(finding.occurrences)}{" "}
          {finding.occurrences === 1 ? "occurrence" : "occurrences"}
        </span>
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

// An account in a table: the name, which opens the person's page when there is one, and the user
// principal name. The address is plain text beside its own copy button, outside the name button,
// so it can be selected, copied in one click, and neither opens the person nor loses the click.
// When the name already is the address, or there is no page to open, there is no second line and
// the copy button follows the name.
export function AccountName({
  name,
  principal,
  onOpen,
}: {
  name: string;
  principal: string | null;
  onOpen: (() => void) | null;
}) {
  const label = <strong className="people-name__name">{name}</strong>;
  const title = onOpen ? (
    <button className="people-name" type="button" onClick={onOpen}>
      {label}
    </button>
  ) : (
    label
  );
  const address = principal?.trim() ? principal : null;
  if (!address) return title;
  if (!onOpen || address === name)
    return (
      <span className="copy-line">
        {title}
        <CopyButton text={address} />
      </span>
    );
  return (
    <div className="account-cell">
      {title}
      <span className="copy-line copy-line--address">
        <span className="copy-line__text">{address}</span>
        <CopyButton text={address} />
      </span>
    </div>
  );
}

import {
  type AuditReview,
  type AuditReviewAudit,
  type AuditReviewAuditState,
  type AuditReviewSchedule,
  type AuditReviewScheduleSettings,
  auditReviewProgress,
  REVIEW_WEEKDAYS,
} from "@prism365/contracts";
import { number, RelativeTime } from "./common.tsx";

// The full review on the Audits page (docs/audits.md, section 22): one button that runs every
// audit you can run, in an order that keeps Microsoft's limits, with its progress and a weekly
// schedule. The pieces draw from props only; the Audits page owns the requests.

const stateLabel: Record<AuditReviewAuditState, string> = {
  waiting: "Waiting",
  running: "Running",
  done: "Done",
  failed: "Failed",
  skipped: "Skipped",
};

export function reviewIsActive(review: AuditReview | null): boolean {
  return review?.state === "running" || review?.state === "stopping";
}

export function reviewCounts(review: AuditReview): Record<AuditReviewAuditState, number> {
  const counts: Record<AuditReviewAuditState, number> = {
    waiting: 0,
    running: 0,
    done: 0,
    failed: 0,
    skipped: 0,
  };
  for (const item of review.audits) counts[item.state] += 1;
  return counts;
}

const plural = (value: number, one: string, many: string) =>
  `${number.format(value)} ${value === 1 ? one : many}`;

function listOf(values: readonly string[]): string {
  if (values.length <= 1) return values.join("");
  return `${values.slice(0, -1).join(", ")} and ${values.at(-1)}`;
}

// What the review is doing, or how it ended, in a sentence.
export function reviewStatus(
  review: AuditReview,
  titleOf: (auditId: string) => string,
): { text: string; failures: string | null } {
  const { ended, total } = auditReviewProgress(review);
  const counts = reviewCounts(review);
  const failed = review.audits.filter((item) => item.state === "failed");
  const failures =
    failed.length > 0
      ? `${plural(failed.length, "audit failed", "audits failed")}: ${listOf(failed.map((item) => titleOf(item.auditId)))}.`
      : null;
  const running = review.audits.find((item) => item.state === "running");
  if (review.state === "running")
    return {
      text: `${ended} of ${total} audits finished or skipped. ${
        running
          ? `Running ${titleOf(running.auditId)}.`
          : "Waiting a moment to stay inside Microsoft's limits."
      }`,
      failures,
    };
  if (review.state === "stopping")
    return {
      text: `Stopping. ${ended} of ${total} audits finished or skipped${
        running ? `, and ${titleOf(running.auditId)} is ending` : ""
      }.`,
      failures,
    };
  const parts = [`${plural(counts.done, "audit", "audits")} done`];
  if (counts.skipped > 0) parts.push(`${counts.skipped} skipped`);
  return {
    text: `${review.state === "stopped" ? "Stopped" : "Finished"}: ${parts.join(", ")}.`,
    failures,
  };
}

function ReviewList({
  review,
  titleOf,
}: {
  review: AuditReview;
  titleOf: (auditId: string) => string;
}) {
  return (
    <ul className="audit-review__list">
      {review.audits.map((item: AuditReviewAudit) => (
        <li key={item.auditId} className={`audit-review__item audit-review__item--${item.state}`}>
          <span className="audit-review__name">{titleOf(item.auditId)}</span>
          <span className="audit-review__state">{stateLabel[item.state]}</span>
          {item.reason ? <span className="audit-review__reason">{item.reason}</span> : null}
        </li>
      ))}
    </ul>
  );
}

const hours = Array.from({ length: 24 }, (_, hour) => hour);
const hourLabel = (hour: number) => `${String(hour).padStart(2, "0")}:00`;

export function ReviewScheduleControl({
  settings,
  saving,
  error,
  disabled = false,
  onChange,
}: {
  settings: AuditReviewScheduleSettings;
  saving: boolean;
  error: string;
  disabled?: boolean;
  onChange: (schedule: AuditReviewSchedule) => void;
}) {
  // One review control on the page, so fixed ids; keeping hooks out lets tests walk the tree.
  const labelId = "audit-review-schedule-label";
  const dayId = "audit-review-schedule-day";
  const hourId = "audit-review-schedule-hour";
  const weekly = settings.frequency === "weekly";
  const change = (over: Partial<AuditReviewSchedule>) =>
    onChange({
      frequency: settings.frequency,
      weekday: settings.weekday,
      hour: settings.hour,
      ...over,
    });
  return (
    <section className="audit-schedule audit-review__schedule" aria-labelledby={labelId}>
      <div className="audit-schedule__row">
        <span className="audit-filter__label" id={labelId}>
          Run a full review automatically
        </span>
        <fieldset className="signin-range" aria-labelledby={labelId} disabled={disabled || saving}>
          {(["off", "weekly"] as const).map((value) => (
            <button
              key={value}
              type="button"
              className={value === settings.frequency ? "is-active" : undefined}
              aria-pressed={value === settings.frequency}
              onClick={() => {
                if (value !== settings.frequency) change({ frequency: value });
              }}
            >
              {value === "off" ? "Off" : "Weekly"}
            </button>
          ))}
        </fieldset>
        {weekly ? (
          <>
            <label className="visually-hidden" htmlFor={dayId}>
              Day of the week
            </label>
            <select
              id={dayId}
              value={settings.weekday}
              disabled={disabled || saving}
              onChange={(event) => change({ weekday: Number(event.target.value) })}
            >
              {REVIEW_WEEKDAYS.map((name, index) => (
                <option key={name} value={index}>
                  {name}
                </option>
              ))}
            </select>
            <label className="visually-hidden" htmlFor={hourId}>
              Hour of the day
            </label>
            <select
              id={hourId}
              value={settings.hour}
              disabled={disabled || saving}
              onChange={(event) => change({ hour: Number(event.target.value) })}
            >
              {hours.map((hour) => (
                <option key={hour} value={hour}>
                  {hourLabel(hour)}
                </option>
              ))}
            </select>
          </>
        ) : null}
        <span className="muted-label" aria-live="polite">
          {saving ? "Saving…" : ""}
        </span>
      </div>
      <p className="muted-label">
        {weekly
          ? `Runs every ${REVIEW_WEEKDAYS[settings.weekday] ?? "Monday"} at ${hourLabel(settings.hour)} (${settings.timeZone}) while collection is on, with the collecting administrator's sign-in${settings.by ? `. Set by ${settings.by}` : ""}. It never reads more than the person who set it could.`
          : "Off. A full review runs only when someone starts it."}
      </p>
      {weekly && !settings.collecting ? (
        <p className="audit-schedule__off">
          Collection is off, so the review will not run on its own until someone starts it.
        </p>
      ) : null}
      {error ? (
        <p className="form-error" role="alert">
          {error}
        </p>
      ) : null}
    </section>
  );
}

export function AuditReviewPanel({
  review,
  titleOf,
  busy,
  error,
  note = null,
  canRun,
  schedule = null,
  scheduleSaving = false,
  scheduleError = "",
  now,
  onRun,
  onStop,
  onSchedule,
}: {
  review: AuditReview | null;
  titleOf: (auditId: string) => string;
  // The start or stop request is in flight.
  busy: boolean;
  error: string;
  // Why the review cannot be started from here, when it cannot.
  note?: string | null;
  canRun: boolean;
  schedule?: AuditReviewScheduleSettings | null;
  scheduleSaving?: boolean;
  scheduleError?: string;
  // For tests: the moment the relative times are measured from.
  now?: number;
  onRun: () => void;
  onStop: () => void;
  onSchedule: (schedule: AuditReviewSchedule) => void;
}) {
  const active = reviewIsActive(review);
  const status = review ? reviewStatus(review, titleOf) : null;
  const { ended, total } = review ? auditReviewProgress(review) : { ended: 0, total: 0 };
  return (
    <section className="panel audit-review" aria-label="Full review">
      <div className="audit-review__head">
        <div className="audit-review__intro">
          <h2 className="audit-group__title">Full review</h2>
          <p className="muted-label">
            Runs every audit you can run, one at a time, and finishes with Combined risks so it
            joins fresh results.
          </p>
        </div>
        {active ? (
          <button
            className="button button--secondary"
            type="button"
            disabled={busy || review?.state === "stopping"}
            aria-busy={busy || undefined}
            onClick={onStop}
          >
            Stop
          </button>
        ) : (
          <button
            className="button"
            type="button"
            disabled={busy || !canRun}
            aria-busy={busy || undefined}
            onClick={onRun}
          >
            Run full review
          </button>
        )}
      </div>
      {note ? <p className="audit-note">{note}</p> : null}
      {error ? (
        <p className="form-error" role="alert">
          {error}
        </p>
      ) : null}
      {review && status ? (
        <div className="audit-review__status" aria-live="polite">
          {active ? (
            <progress
              className="audit-review__bar"
              max={total}
              value={ended}
              aria-label="Full review progress"
            />
          ) : null}
          <p>{status.text}</p>
          {status.failures ? <p className="audit-schedule__off">{status.failures}</p> : null}
          {!active && review.finishedAt ? (
            <p className="muted-label">
              Last full review {review.state === "stopped" ? "stopped" : "finished"}{" "}
              <RelativeTime value={review.finishedAt} {...(now === undefined ? {} : { now })} />
              {review.trigger === "schedule" ? ", started on its schedule" : ""}.
            </p>
          ) : null}
          <details className="audit-review__details">
            <summary>What happened to each audit</summary>
            <ReviewList review={review} titleOf={titleOf} />
          </details>
        </div>
      ) : (
        <p className="muted-label">No full review has run yet.</p>
      )}
      {schedule ? (
        <ReviewScheduleControl
          settings={schedule}
          saving={scheduleSaving}
          error={scheduleError}
          onChange={onSchedule}
        />
      ) : null}
    </section>
  );
}

import { type PersonSummary as PersonSummaryData, personSummarySchema } from "@prism365/contracts";
import { useEffect, useState } from "react";
import { RelativeTime, readJson, SeverityDot, SkeletonRows } from "./common.tsx";
import { ChevronIcon } from "./icons.tsx";
import { EntityLink } from "./navigation.tsx";
import { accessibleName, type Jump, type SummaryItem, summaryItems } from "./personSummaryLines.ts";
import type { PersonTab } from "./PersonTabs.tsx";

// The Overview's first panel: an index of every place Prism365 knows this person from. One row per
// source, a short sentence, when it last ran, and a way to the detail. Read from stored results
// only, so it costs Microsoft nothing; the page's other panels load as they always did.

const SCROLL_TRIES = 30;
const SCROLL_WAIT_MILLISECONDS = 100;

// Scrolls to a section of a tab, waiting for it to exist: a tab that has not been opened yet draws
// its sections only once its answer arrives. The tab's top comes into view as soon as the tab
// shows, so a section that never appears (the Guest section of an account that is no longer a
// guest) still leaves the administrator on the right tab, at its top.
export function scrollToSection(id: string, tab: PersonTab, tries = SCROLL_TRIES, landed = false) {
  if (typeof document === "undefined") return;
  const element = document.getElementById(id);
  const panel = document.querySelector<HTMLElement>(`[role=tabpanel][id$="-panel-${tab}"]`);
  const shown = panel !== null && !panel.hidden;
  if (element && shown && !element.closest("[hidden]")) {
    element.scrollIntoView({ block: "start" });
    // The focus follows, so a keyboard user lands where the page scrolled to.
    element.setAttribute("tabindex", "-1");
    element.focus({ preventScroll: true });
    return;
  }
  let reached = landed;
  if (shown && panel && !landed) {
    panel.scrollIntoView({ block: "start" });
    reached = true;
  }
  if (tries > 0)
    setTimeout(() => scrollToSection(id, tab, tries - 1, reached), SCROLL_WAIT_MILLISECONDS);
  else if (shown && panel) panel.focus({ preventScroll: true });
}

export const personHref = (userId: string, tab: PersonTab) =>
  `#/people/${encodeURIComponent(userId)}${tab === "overview" ? "" : `/${tab}`}`;

const hrefOf = (jump: Jump, userId: string) =>
  jump.kind === "page" ? jump.href : personHref(userId, jump.tab);

function Row({
  entry,
  userId,
  now,
  onJump,
}: {
  entry: SummaryItem;
  userId: string;
  now?: number;
  onJump: (tab: PersonTab, sectionId: string) => void;
}) {
  const { jump } = entry;
  const name = accessibleName(entry);
  // The whole row is the link; its name is the full sentence, since the visible one may be cut.
  const said = entry.severity ? (
    <>
      <SeverityDot
        severity={entry.severity}
        {...(entry.severityLabel ? { label: entry.severityLabel } : {})}
      />{" "}
      {entry.sentence}
    </>
  ) : entry.tone === "bad" ? (
    <SeverityDot severity="high" label={entry.sentence} />
  ) : entry.tone === "warning" ? (
    <SeverityDot severity="medium" label={entry.sentence} />
  ) : entry.tone === "good" ? (
    <span className="person-access__good">{entry.sentence}</span>
  ) : (
    entry.sentence
  );
  return (
    <li>
      <a
        className={`person-summary__row${entry.muted ? " person-summary__row--muted" : ""}`}
        href={hrefOf(jump, userId)}
        aria-label={name}
        title={`${entry.source}: ${entry.sentence}`}
        onClick={(event) => {
          if (jump.kind === "page") return;
          // A modified click keeps its usual meaning (a new tab, say).
          if (
            event.metaKey ||
            event.ctrlKey ||
            event.shiftKey ||
            event.altKey ||
            event.button !== 0
          )
            return;
          event.preventDefault();
          onJump(jump.tab, jump.id);
        }}
      >
        <span className="person-summary__source">{entry.source}</span>
        <span className="person-summary__said">{said}</span>
        <span className="person-summary__when">
          {entry.ranAt ? <RelativeTime value={entry.ranAt} {...(now ? { now } : {})} /> : null}
        </span>
        <ChevronIcon className="person-summary__chevron" />
      </a>
      {entry.entities && entry.entities.length > 0 ? (
        // Outside the row's link, which is one link to the detail: a link cannot hold another.
        <p className="person-summary__things link-list">
          {entry.entities.map((entity, index) => (
            <span key={entity.id}>
              <EntityLink kind={entity.kind} id={entity.id} name={entity.name} />
              {index < (entry.entities?.length ?? 0) - 1 || entry.entitiesMore ? "," : null}{" "}
            </span>
          ))}
          {entry.entitiesMore ? (
            <span className="person-muted"> and {entry.entitiesMore} more</span>
          ) : null}
        </p>
      ) : null}
    </li>
  );
}

// The panel as a function of what has loaded. It holds no state, so a test can draw it.
export function PersonSummaryView({
  summary,
  error,
  userId,
  onJump,
  now,
}: {
  // null while it loads.
  summary: PersonSummaryData | null;
  error: string;
  userId: string;
  onJump: (tab: PersonTab, sectionId: string) => void;
  now?: number;
}) {
  const body = (() => {
    if (error) return <p className="empty-data">{error}</p>;
    if (!summary) return <SkeletonRows rows={4} label="Reading where this person shows up" />;
    const { items, notRun } = summaryItems(summary, now);
    const column = (id: "risk" | "account", label: string) => (
      <div className="person-summary__column">
        <h4 className="person-summary__label">{label}</h4>
        <ul className="person-summary__list">
          {items
            .filter((entry) => entry.column === id)
            .map((entry) => (
              <Row
                key={entry.id}
                entry={entry}
                userId={userId}
                onJump={onJump}
                {...(now ? { now } : {})}
              />
            ))}
        </ul>
      </div>
    );
    return (
      <>
        <div className="person-summary__columns">
          {column("risk", "Risk and access")}
          {column("account", "Account")}
        </div>
        {notRun.length > 0 ? (
          <p className="person-summary__not-run person-muted">
            Not run yet: {notRun.join(", ")}. <a href="#/audits">Run a full review</a>
          </p>
        ) : null}
      </>
    );
  })();
  return (
    <section className="panel person-summary" aria-label="Where they show up">
      <h3>Where they show up</h3>
      {body}
    </section>
  );
}

// Loads with the page. A failed load is one plain line and the rest of the page works.
export function PersonSummary({
  live,
  userId,
  onJump,
}: {
  live: boolean;
  userId: string;
  onJump: (tab: PersonTab, sectionId: string) => void;
}) {
  const [summary, setSummary] = useState<PersonSummaryData | null>(null);
  const [error, setError] = useState("");
  useEffect(() => {
    if (!live) return;
    const controller = new AbortController();
    readJson(
      `/api/v1/people/${encodeURIComponent(userId)}/summary`,
      (value) => personSummarySchema.parse(value),
      { signal: controller.signal },
    )
      .then((value) => {
        setSummary(value);
        setError("");
      })
      .catch((loadError: unknown) => {
        if (controller.signal.aborted) return;
        setError(
          loadError instanceof Error && loadError.message
            ? loadError.message
            : "Where this person shows up could not be loaded.",
        );
      });
    return () => controller.abort();
  }, [live, userId]);
  if (!live) return null;
  return <PersonSummaryView summary={summary} error={error} userId={userId} onJump={onJump} />;
}

import {
  type ChangeCoverage,
  type Digest,
  type DigestDays,
  type DigestDifference,
  type DigestFinding,
  type DigestSectionState,
  digestSchema,
  type PrismSession,
} from "@prism365/contracts";
import { type ReactNode, useEffect, useState } from "react";
import type { WorkspaceView } from "./App.tsx";
import { categoryLabel } from "./Changes.tsx";
import {
  AccountName,
  day,
  number,
  RelativeTime,
  readJson,
  SeverityDot,
  SkeletonRows,
  when,
} from "./common.tsx";

// This week: what Prism365 stored in the last 1, 7, 14 or 30 days, one panel per kind of thing,
// each item a link to where it can be looked at. The page reads what collection and the audits
// already stored; opening it asks Microsoft nothing. The data comes from GET /v1/digest
// (docs/this-week.md), and every panel says in one line what it holds or why it is not available.

export const builtFrom =
  "Built from what Prism365 collected. Nothing here is read from Microsoft when you open the page.";

export const ranges: { days: DigestDays; label: string }[] = [
  { days: 1, label: "Last day" },
  { days: 7, label: "Last 7 days" },
  { days: 14, label: "Last 14 days" },
  { days: 30, label: "Last 30 days" },
];

export const digestUrl = (days: DigestDays) => `/api/v1/digest?days=${days}`;

// Where an item opens. The page builds these and the tests read them, so what a click does is
// not left to the markup.
export type WeekTarget =
  | { type: "person"; id: string }
  | { type: "view"; view: WorkspaceView }
  | { type: "audit"; id: string };

// A finding about a person opens that person's page, where their findings are; any other opens
// the page its module owns.
export function findingTarget(finding: DigestFinding, view: WorkspaceView): WeekTarget {
  return finding.personId ? { type: "person", id: finding.personId } : { type: "view", view };
}

export interface WeekLinks {
  onNavigate: (view: WorkspaceView) => void;
  onOpenPerson: (userId: string) => void;
  onOpenAudit: (auditId: string) => void;
}

function open(links: WeekLinks, target: WeekTarget) {
  if (target.type === "person") links.onOpenPerson(target.id);
  else if (target.type === "audit") links.onOpenAudit(target.id);
  else links.onNavigate(target.view);
}

export function coverageText(coverage: ChangeCoverage): ReactNode {
  if (coverage.phase === "none" || coverage.from === null) return null;
  if (coverage.phase === "history")
    return `Loading the last 30 days, oldest first. History has reached ${when(coverage.through)}.`;
  return (
    <>
      History kept since {day(coverage.from)}. Last collected{" "}
      <RelativeTime value={coverage.lastCollected} />.
    </>
  );
}

const differenceWords: Record<DigestDifference, string> = {
  version: "another version of the audit",
  scope: "another account scope",
  incomplete: "a partial result",
};

export function differenceText(differences: readonly DigestDifference[]): string | null {
  if (differences.length === 0) return null;
  return `Not directly comparable: ${differences.map((item) => differenceWords[item]).join(", ")}.`;
}

export function signed(change: number): string {
  if (change === 0) return "0";
  return `${change > 0 ? "+" : "-"}${number.format(Math.abs(change))}`;
}

function Link({
  links,
  target,
  children,
}: {
  links: WeekLinks;
  target: WeekTarget;
  children: ReactNode;
}) {
  return (
    <button className="link-button" type="button" onClick={() => open(links, target)}>
      {children}
    </button>
  );
}

function Note({ state, children }: { state: DigestSectionState; children: string }) {
  return (
    <p className={`week-note${state === "unavailable" ? " week-note--unavailable" : ""}`}>
      {children}
    </p>
  );
}

function Panel({
  id,
  title,
  action,
  state,
  children,
}: {
  id: string;
  title: string;
  action: ReactNode;
  state: DigestSectionState;
  children: ReactNode;
}) {
  return (
    <section className="panel week-panel" aria-labelledby={`week-${id}`} data-state={state}>
      <div className="panel__header">
        <h3 id={`week-${id}`}>{title}</h3>
        {action}
      </div>
      {children}
    </section>
  );
}

function FindingList({
  findings,
  view,
  links,
}: {
  findings: readonly DigestFinding[];
  view: WorkspaceView;
  links: WeekLinks;
}) {
  return (
    <ul className="week-list">
      {findings.map((finding) => (
        <li key={finding.id}>
          <SeverityDot severity={finding.severity} />
          <div>
            <Link links={links} target={findingTarget(finding, view)}>
              {finding.title}
            </Link>
            <span className="week-sub">
              {finding.subjectLabel}, first seen <RelativeTime value={finding.firstSeen} />
            </span>
          </div>
        </li>
      ))}
    </ul>
  );
}

function More({ shown, total, noun }: { shown: number; total: number; noun: string }) {
  if (total <= shown) return null;
  return (
    <p className="week-sub">
      Showing {number.format(shown)} of {number.format(total)} {noun}.
    </p>
  );
}

const severities = ["critical", "high", "medium", "low", "info"] as const;

function SignIns({ digest, links }: { digest: Digest; links: WeekLinks }) {
  const section = digest.signInFindings;
  const opened = severities.filter((severity) => section.opened[severity] > 0);
  return (
    <Panel
      id="signins"
      title="Sign-in findings"
      state={section.state}
      action={
        <Link links={links} target={{ type: "view", view: "signins" }}>
          Open sign-in audit
        </Link>
      }
    >
      <Note state={section.state}>{section.note}</Note>
      {section.state === "ok" ? (
        <>
          {opened.length > 0 ? (
            <ul className="week-counts" aria-label="Opened by severity">
              {opened.map((severity) => (
                <li key={severity}>
                  <SeverityDot severity={severity} count={section.opened[severity]} />
                </li>
              ))}
            </ul>
          ) : null}
          {section.open.length > 0 ? (
            <>
              <h4 className="week-heading">Most severe still open</h4>
              <FindingList findings={section.open} view="signins" links={links} />
              <More shown={section.open.length} total={section.openTotal} noun="open findings" />
            </>
          ) : null}
        </>
      ) : null}
    </Panel>
  );
}

function Changes({ digest, links }: { digest: Digest; links: WeekLinks }) {
  const section = digest.directoryChanges;
  const coverage = coverageText(section.coverage);
  return (
    <Panel
      id="changes"
      title="Directory changes"
      state={section.state}
      action={
        <Link links={links} target={{ type: "view", view: "changes" }}>
          Open Directory changes
        </Link>
      }
    >
      <Note state={section.state}>{section.note}</Note>
      {section.state === "ok" ? (
        <>
          {section.categories.length > 0 ? (
            <ul className="week-counts" aria-label="Changes by category">
              {section.categories.map((item) => (
                <li key={item.category ?? "none"}>
                  <span className="person-tag">
                    {categoryLabel(item.category)} {number.format(item.count)}
                  </span>
                </li>
              ))}
            </ul>
          ) : null}
          {section.flagged.length > 0 ? (
            <>
              <h4 className="week-heading">Flagged by the rules</h4>
              <FindingList findings={section.flagged} view="changes" links={links} />
              <More
                shown={section.flagged.length}
                total={section.flaggedTotal}
                noun="flagged changes"
              />
            </>
          ) : null}
        </>
      ) : null}
      {coverage ? <p className="week-sub">{coverage}</p> : null}
    </Panel>
  );
}

function Quiet({ digest, links }: { digest: Digest; links: WeekLinks }) {
  const section = digest.quietAccounts;
  return (
    <Panel
      id="quiet"
      title="Accounts that went quiet"
      state={section.state}
      action={
        <Link links={links} target={{ type: "audit", id: section.auditId }}>
          Open Inactive accounts
        </Link>
      }
    >
      <Note state={section.state}>{section.note}</Note>
      {section.state === "ok" ? (
        <>
          <ul className="week-list week-list--accounts">
            {section.items.map((item, index) => {
              const personId = item.personId;
              return (
                <li key={personId ?? `row-${index}`}>
                  <AccountName
                    name={item.name ?? item.principal ?? "Unnamed account"}
                    principal={item.principal}
                    onOpen={personId ? () => open(links, { type: "person", id: personId }) : null}
                  />
                  <span className="week-sub">
                    {item.days === null ? "Days unknown" : `${number.format(item.days)} days`}
                  </span>
                </li>
              );
            })}
          </ul>
          <More shown={section.items.length} total={section.total} noun="accounts" />
        </>
      ) : null}
      {section.earlier ? (
        <p className="week-sub">
          Compared with the result from <RelativeTime value={section.earlier.at} />. The audit opens
          on its 30+ days view.
        </p>
      ) : null}
    </Panel>
  );
}

function Drift({ digest, links }: { digest: Digest; links: WeekLinks }) {
  const section = digest.auditDrift;
  return (
    <Panel
      id="drift"
      title="Audit changes"
      state={section.state}
      action={
        <Link links={links} target={{ type: "view", view: "audits" }}>
          Open Audits
        </Link>
      }
    >
      <Note state={section.state}>{section.note}</Note>
      {section.state === "ok" ? (
        <ul className="week-audits">
          {section.audits.map((audit) => (
            <li key={audit.auditId}>
              <div className="week-audit__header">
                <Link links={links} target={{ type: "audit", id: audit.auditId }}>
                  {audit.title}
                </Link>
                <span className="week-sub">
                  Compared with the result from <RelativeTime value={audit.from.at} />
                </span>
              </div>
              {differenceText(audit.differences) ? (
                <p className="week-note week-note--unavailable">
                  {differenceText(audit.differences)}
                </p>
              ) : null}
              <ul className="week-tiles">
                {audit.tiles.map((tile) => (
                  <li key={tile.id} className={`week-tile week-tile--${tile.tone}`}>
                    <span>{tile.label}</span>
                    <span className="week-tile__values">
                      {number.format(tile.from)} to {number.format(tile.to)}
                    </span>
                    <strong>{signed(tile.change)}</strong>
                  </li>
                ))}
              </ul>
            </li>
          ))}
        </ul>
      ) : null}
    </Panel>
  );
}

function Microsoft({ digest, links }: { digest: Digest; links: WeekLinks }) {
  const { posts, incidents, state } = digest.microsoft;
  const notifications: WeekTarget = { type: "view", view: "notifications" };
  return (
    <Panel
      id="microsoft"
      title="Microsoft"
      state={state}
      action={
        <Link links={links} target={notifications}>
          Open Notifications
        </Link>
      }
    >
      <div className="week-part" data-state={posts.state}>
        <h4 className="week-heading">Act-by dates in the next 30 days</h4>
        <Note state={posts.state}>{posts.note}</Note>
        {posts.state === "ok" ? (
          <>
            <ul className="week-list">
              {posts.items.map((post) => (
                <li key={post.id}>
                  <div>
                    <Link links={links} target={notifications}>
                      {post.title}
                    </Link>
                    <span className="week-sub">
                      {[`Act by ${day(post.actBy)}`, post.services.join(", ")]
                        .filter(Boolean)
                        .join(", ")}
                    </span>
                  </div>
                </li>
              ))}
            </ul>
            <More shown={posts.items.length} total={posts.total} noun="posts" />
          </>
        ) : null}
      </div>
      <div className="week-part" data-state={incidents.state}>
        <h4 className="week-heading">Service incidents</h4>
        <Note state={incidents.state}>{incidents.note}</Note>
        {incidents.state === "ok" && incidents.open.length > 0 ? (
          <>
            <ul className="week-list">
              {incidents.open.map((incident) => (
                <li key={incident.id}>
                  <SeverityDot severity={incident.severity} />
                  <div>
                    <Link links={links} target={notifications}>
                      {incident.title}
                    </Link>
                    <span className="week-sub">
                      {[incident.status, incident.services.join(", ")].filter(Boolean).join(", ")}
                      {incident.status || incident.services.length > 0 ? ", " : ""}first seen{" "}
                      <RelativeTime value={incident.firstSeen} />
                    </span>
                  </div>
                </li>
              ))}
            </ul>
            <More shown={incidents.open.length} total={incidents.openTotal} noun="open incidents" />
          </>
        ) : null}
      </div>
    </Panel>
  );
}

function Collection({ digest, links }: { digest: Digest; links: WeekLinks }) {
  const section = digest.collection;
  return (
    <Panel
      id="collection"
      title="Collection"
      state={section.state}
      action={
        <Link links={links} target={{ type: "view", view: "signins" }}>
          {section.collecting ? "Open sign-in audit" : "Start collecting"}
        </Link>
      }
    >
      <Note state={section.state}>{section.note}</Note>
      {section.state === "ok" ? (
        <>
          <p className="week-status">
            <span
              className={`collection-dot collection-dot--${section.collecting ? "on" : "off"}`}
            />
            {section.collecting ? "Collecting" : "Not collecting"}
          </p>
          <ul className="week-list week-list--collectors">
            {section.collectors.map((item) => (
              <li key={`${item.module}/${item.collector}`}>
                <div>
                  <strong className="overview-list__title">{item.moduleTitle}</strong>
                  <span className="week-sub">
                    {item.lastSuccess ? (
                      <>
                        Last success <RelativeTime value={item.lastSuccess} />
                      </>
                    ) : (
                      "No success in the last 30 days of run history"
                    )}
                  </span>
                </div>
                {item.lastRun && item.lastRun.status !== "succeeded" ? (
                  <span
                    className={`person-tag${item.lastRun.status === "failed" ? " person-tag--alert" : ""}`}
                  >
                    Last run {item.lastRun.status}
                  </span>
                ) : null}
              </li>
            ))}
          </ul>
          {section.failedRuns > 0 || section.gaps > 0 ? (
            <p className="week-sub">
              {[
                section.failedRuns > 0
                  ? `${number.format(section.failedRuns)} failed ${section.failedRuns === 1 ? "run" : "runs"}${
                      section.errors.length > 0 ? ` (${section.errors.join(", ")})` : ""
                    }.`
                  : null,
                section.gaps > 0
                  ? `${number.format(section.gaps)} ${section.gaps === 1 ? "gap" : "gaps"} in what was read.`
                  : null,
              ]
                .filter(Boolean)
                .join(" ")}
            </p>
          ) : null}
          {section.audits.length > 0 ? (
            <>
              <h4 className="week-heading">Audits that run on their own</h4>
              <ul className="week-list">
                {section.audits.map((audit) => (
                  <li key={audit.auditId}>
                    <div>
                      <Link links={links} target={{ type: "audit", id: audit.auditId }}>
                        {audit.title}
                      </Link>
                      <span className="week-sub">
                        {audit.frequency === "daily" ? "Daily" : "Weekly"}.{" "}
                        {audit.lastRun ? (
                          <>
                            Last ran <RelativeTime value={audit.lastRun.at} />,{" "}
                            {audit.lastRun.status}.
                          </>
                        ) : (
                          "Has not run yet."
                        )}
                      </span>
                    </div>
                  </li>
                ))}
              </ul>
            </>
          ) : null}
        </>
      ) : null}
    </Panel>
  );
}

export interface WeekViewProps extends WeekLinks {
  // null while the page is loading, or when the range changed and the answer has not arrived.
  digest: Digest | null;
  days: DigestDays;
  error: string;
  onDays: (days: DigestDays) => void;
}

// The page with its data: shown by ThisWeek and rendered directly in tests. It keeps no state, so
// a test can read what each link does without a browser.
export function WeekView(props: WeekViewProps) {
  const { digest, days } = props;
  const links: WeekLinks = {
    onNavigate: props.onNavigate,
    onOpenPerson: props.onOpenPerson,
    onOpenAudit: props.onOpenAudit,
  };
  return (
    <section className="signin-audit week" aria-labelledby="week-title">
      <h2 id="week-title" className="visually-hidden">
        This week
      </h2>
      <div className="week-top">
        <p className="signin-coverage">{builtFrom}</p>
        <fieldset className="signin-range">
          <legend className="visually-hidden">Time range</legend>
          {ranges.map((range) => (
            <button
              key={range.days}
              type="button"
              aria-pressed={days === range.days}
              className={days === range.days ? "is-active" : ""}
              onClick={() => props.onDays(range.days)}
            >
              {range.label}
            </button>
          ))}
        </fieldset>
      </div>
      {props.error ? (
        <p className="form-error" role="alert">
          {props.error}
        </p>
      ) : null}
      {digest === null ? (
        // After an error there is nothing coming, so the error is all the page says.
        props.error ? null : (
          <div aria-live="polite">
            <SkeletonRows rows={6} label="Loading this week" />
          </div>
        )
      ) : (
        <>
          <SignIns digest={digest} links={links} />
          <Changes digest={digest} links={links} />
          <Quiet digest={digest} links={links} />
          <Drift digest={digest} links={links} />
          <Microsoft digest={digest} links={links} />
          <Collection digest={digest} links={links} />
        </>
      )}
    </section>
  );
}

function Intro({ title, children }: { title: string; children: string }) {
  return (
    <section className="signin-audit" aria-labelledby="week-intro">
      <div className="signin-audit__intro">
        <span className="eyebrow">This week</span>
        <h2 id="week-intro">{title}</h2>
        <p>{children}</p>
      </div>
    </section>
  );
}

export function ThisWeek({
  session,
  ...links
}: {
  session: PrismSession;
} & WeekLinks) {
  if (session.mode === "demo")
    return (
      <Intro title="What changed since you last looked.">
        Connected to a tenant, this page gathers what Prism365 collected in the last day, week or
        month: new sign-in findings, directory changes, accounts that went quiet, audit numbers that
        moved, and Microsoft deadlines. Demo mode shows none of it, because Prism365 does not invent
        tenant data.
      </Intro>
    );
  if (!session.authenticated)
    return (
      <Intro title="Sign in to see this week.">
        Use the tenant connection above. This page reads what collection and the audits stored, with
        your own delegated access.
      </Intro>
    );
  return <WeekWorkspace {...links} />;
}

// An answer for another range is not shown under this range's button.
export function digestFor(digest: Digest | null, days: DigestDays): Digest | null {
  return digest?.days === days ? digest : null;
}

function WeekWorkspace(links: WeekLinks) {
  const [days, setDays] = useState<DigestDays>(7);
  const [digest, setDigest] = useState<Digest | null>(null);
  const [error, setError] = useState("");

  useEffect(() => {
    const controller = new AbortController();
    setError("");
    readJson(digestUrl(days), (value) => digestSchema.parse(value), {
      signal: controller.signal,
    })
      .then((loaded) => {
        if (!controller.signal.aborted) setDigest(loaded);
      })
      .catch((loadError: unknown) => {
        if (controller.signal.aborted) return;
        setError(loadError instanceof Error ? loadError.message : "This week could not be loaded.");
      });
    return () => controller.abort();
  }, [days]);

  return (
    <WeekView
      {...links}
      digest={digestFor(digest, days)}
      days={days}
      error={error}
      onDays={setDays}
    />
  );
}

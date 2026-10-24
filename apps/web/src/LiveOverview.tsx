import {
  type Finding,
  findingListSchema,
  type ModuleList,
  moduleListSchema,
  type PersonList,
  type PrismSession,
  personListSchema,
  type SigninSummary,
  signinSummarySchema,
} from "@prism365/contracts";
import { useCallback, useEffect, useState } from "react";
import type { WorkspaceView } from "./App.tsx";
import { AdminAttention } from "./Attention.tsx";
import { number, readJson, when } from "./common.tsx";
import { ServiceHealthPanel, useServiceHealth } from "./ServiceHealth.tsx";
import { runLine } from "./SigninAudit.tsx";

// The live home page: is collection running, what the last week looks like, what is open, and
// who signed in last. Everything here links to the view that owns it.

const MODULE = "signins.audit";

export function LiveOverview({
  session,
  onOpenPerson,
  onNavigate,
}: {
  session: PrismSession;
  onOpenPerson: (userId: string) => void;
  onNavigate: (view: WorkspaceView) => void;
}) {
  const [modules, setModules] = useState<ModuleList | null>(null);
  const [summary, setSummary] = useState<SigninSummary | null>(null);
  const [findings, setFindings] = useState<Finding[] | null>(null);
  const [people, setPeople] = useState<PersonList | null>(null);
  const [error, setError] = useState("");
  const health = useServiceHealth(session);

  const load = useCallback(
    async (signal?: AbortSignal) => {
      if (!session.authenticated) return;
      const init = signal ? { signal } : {};
      try {
        const [moduleList, week, open, recent] = await Promise.all([
          readJson("/api/v1/modules", (value) => moduleListSchema.parse(value), init),
          readJson(
            `/api/v1/modules/${MODULE}/summary?days=7`,
            (value) => signinSummarySchema.parse(value),
            init,
          ),
          readJson(
            `/api/v1/findings?module=${MODULE}&state=open&limit=100`,
            (value) => findingListSchema.parse(value),
            init,
          ),
          readJson("/api/v1/people?days=7", (value) => personListSchema.parse(value), init),
        ]);
        setModules(moduleList);
        setSummary(week);
        setFindings(open.items);
        setPeople(recent);
        setError("");
      } catch (loadError) {
        if (signal?.aborted) return;
        setError(loadError instanceof Error ? loadError.message : "The overview could not load.");
      }
    },
    [session.authenticated],
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

  const module = modules?.modules.find((item) => item.id === MODULE);
  const run = module?.collectors[0]?.lastRun ?? null;
  const collecting = modules?.collection.state === "collecting";
  const open = findings ?? [];
  const recent = (people?.items ?? []).slice(0, 8);

  return (
    <div className="overview">
      {error ? (
        <p className="form-error form-error--banner" role="alert">
          {error}
        </p>
      ) : null}

      <section className="panel overview-status" aria-label="Collection">
        <span className={`collection-dot collection-dot--${collecting ? "on" : "off"}`} />
        <div>
          <strong>
            {collecting
              ? "Collecting sign-ins"
              : modules?.collection.resumes
                ? "Collection paused"
                : "Not collecting"}
          </strong>
          <p>
            {modules ? runLine(run, collecting) : "Checking collection…"}
            {summary?.history.phase === "backfill"
              ? ` History has reached ${when(summary.history.through)}.`
              : ""}
          </p>
        </div>
        <button
          className="button button--secondary"
          type="button"
          onClick={() => onNavigate("signins")}
        >
          {collecting ? "Open sign-in audit" : "Start collecting"}
        </button>
      </section>

      <ServiceHealthPanel status={health.status} error={health.error} onNavigate={onNavigate} />
      <AdminAttention session={session} onNavigate={onNavigate} />

      <section className="signin-stats" aria-label="Last 7 days">
        {(
          [
            ["Successful sign-ins, 7 days", summary?.totals.successes],
            ["Failed password attempts", summary?.totals.failures],
            ["Accounts signing in", summary?.totals.accounts],
            ["Open findings", findings ? open.length : undefined],
          ] as const
        ).map(([label, value]) => (
          <article className="metric-card signin-stat" key={label}>
            <span>{label}</span>
            <strong>
              {value === undefined
                ? "–"
                : `${number.format(value)}${label === "Open findings" && open.length >= 100 ? "+" : ""}`}
            </strong>
          </article>
        ))}
      </section>

      <div className="overview-grid">
        <section className="panel overview-list" aria-labelledby="overview-findings">
          <div className="overview-list__header">
            <h2 id="overview-findings">Open findings</h2>
            <button className="link-button" type="button" onClick={() => onNavigate("signins")}>
              View all
            </button>
          </div>
          {findings === null ? (
            <p className="empty-data">Loading…</p>
          ) : open.length === 0 ? (
            <p className="empty-data">Nothing open. New findings appear here as collection runs.</p>
          ) : (
            <ul>
              {open.slice(0, 6).map((finding) => (
                <li key={finding.id}>
                  <span className={`finding-severity finding-severity--${finding.severity}`}>
                    {finding.severity}
                  </span>
                  <div>
                    <strong className="overview-list__title">{finding.title}</strong>
                    <span className="overview-list__sub">
                      {finding.subject.kind === "person" ? (
                        <button
                          className="link-button"
                          type="button"
                          onClick={() => onOpenPerson(finding.subject.id)}
                        >
                          {finding.subject.label ?? finding.subject.id}
                        </button>
                      ) : (
                        (finding.subject.label ?? finding.subject.id)
                      )}
                      {" · "}
                      {when(finding.lastSeen)}
                    </span>
                  </div>
                </li>
              ))}
            </ul>
          )}
        </section>

        <section className="panel overview-list" aria-labelledby="overview-people">
          <div className="overview-list__header">
            <h2 id="overview-people">Recently signed in</h2>
            <button className="link-button" type="button" onClick={() => onNavigate("people")}>
              All people
            </button>
          </div>
          {people === null ? (
            <p className="empty-data">Loading…</p>
          ) : recent.length === 0 ? (
            <p className="empty-data">No sign-ins stored for the last 7 days yet.</p>
          ) : (
            <ul>
              {recent.map((item) => (
                <li key={item.userId}>
                  <div>
                    <button
                      className="link-button link-button--strong"
                      type="button"
                      onClick={() => onOpenPerson(item.userId)}
                    >
                      {item.displayName ?? item.userPrincipalName ?? item.userId}
                    </button>
                    <span className="overview-list__sub">{item.userPrincipalName}</span>
                  </div>
                  <span className="overview-list__meta">
                    {item.failures > 0 ? (
                      <span className="person-tag person-tag--alert">{item.failures} failed</span>
                    ) : null}
                    {when(item.lastSignIn)}
                  </span>
                </li>
              ))}
            </ul>
          )}
        </section>
      </div>
    </div>
  );
}

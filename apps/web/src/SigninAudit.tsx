import {
  type Finding,
  findingListSchema,
  findingSchema,
  type FindingState,
  type ModuleList,
  moduleListSchema,
  type PrismSession,
  type SigninSummary,
  signinSummarySchema,
} from "@prism365/contracts";
import { useCallback, useEffect, useId, useState } from "react";
import { countryName, FindingItem, number, post, readJson, stateLabels, when } from "./common.tsx";
import { ActivityIcon, ShieldIcon } from "./icons.tsx";

const MODULE = "signins.audit";

const categoryCopy: Record<string, string> = {
  access_denied:
    "Microsoft denied access. Your account needs Reports Reader, Security Reader, or Global Reader, and the tenant needs Microsoft Entra ID P1 or P2.",
  credential_unavailable:
    "Your sign-in ended, so collection paused. It resumes when you sign in again.",
  throttled: "Microsoft asked Prism365 to slow down. The next run will continue.",
  provider_unavailable: "Microsoft Graph was unavailable. The next run will retry.",
  provider_error: "Microsoft Graph rejected the request. The next run will retry.",
  timeout: "The run took too long and stopped. The next run continues where it left off.",
  boundary: "A response did not match what Prism365 expects, so it was not stored.",
  "boundary.timeout":
    "Microsoft Graph took longer than 30 seconds to answer. The next run retries.",
  "boundary.size":
    "A page from Microsoft Graph was larger than Prism365 accepts, so it was not stored.",
  "boundary.continuation":
    "Microsoft Graph's next-page link pointed somewhere Prism365 does not read, so collection stopped.",
  "boundary.item": "Microsoft Graph returned a sign-in record Prism365 could not read.",
  "boundary.shape": "Microsoft Graph returned a response Prism365 could not read.",
  "boundary.content_type": "Microsoft Graph returned something other than JSON.",
  "boundary.redirect": "Microsoft Graph redirected the request, which Prism365 does not follow.",
  "boundary.retry_budget": "Microsoft Graph kept asking Prism365 to wait. The next run retries.",
  "boundary.page_limit":
    "A time window held more records than one request may read. The next run continues.",
  "boundary.module": "The module asked for something outside what it declared.",
  storage_denied: "Storage refused a write. Check the server log.",
  storage_error: "Storage reported an error. Check the server log.",
  collector_error: "The run failed. Check the server log.",
  invalid_finding: "A finding did not pass validation and was not stored.",
  sql_guard: "A storage query was refused by the safety guard.",
};

const stopCopy: Record<string, string> = {
  operator_stopped: "An administrator stopped it.",
  signed_out: "The administrator who started it signed out.",
  credential_unavailable: "The administrator's session ended or lost access.",
  replaced: "Another administrator started collection.",
  shutdown: "The server restarted.",
};

type ModuleRunView = NonNullable<ModuleList["modules"][number]["collectors"][number]["lastRun"]>;

export function runLine(run: ModuleRunView | null, collecting = true): string {
  if (!run) return "No collection has run yet. The first run loads up to 30 days of history.";
  // Counts are saved when a run ends, so a run in progress has none to show yet.
  if (run.status === "running")
    // A run still marked running while nothing collects was cut off by a restart; the next
    // start marks it abandoned and carries on from its saved place.
    return collecting
      ? `Reading sign-ins now (started ${when(run.startedAt)}).`
      : `The run started ${when(run.startedAt)} was cut short. Start collecting to continue where it left off.`;
  if (run.status === "failed")
    return `Last run ${when(run.startedAt)} failed. ${categoryCopy[run.errorCategory ?? ""] ?? categoryCopy[(run.errorCategory ?? "").split(".")[0] ?? ""] ?? "Check the server log."}`;
  if (run.status === "abandoned")
    return `The run started ${when(run.startedAt)} was cut short by a restart. Collection continues where it left off.`;
  const note =
    run.skipped > 0
      ? `, ${number.format(run.skipped)} unreadable and left out`
      : run.status === "partial"
        ? ", but some sign-ins in a busy stretch could not be read (see findings)"
        : "";
  return `Last run ${when(run.startedAt)}: ${number.format(run.fetched)} records read, ${number.format(run.inserted)} new${note}.`;
}

export function SigninAudit({
  session,
  onOpenPerson,
}: {
  session: PrismSession;
  onOpenPerson?: (userId: string) => void;
}) {
  const titleId = useId();
  const [modules, setModules] = useState<ModuleList | null>(null);
  const [summary, setSummary] = useState<SigninSummary | null>(null);
  const [findings, setFindings] = useState<Finding[]>([]);
  const [filter, setFilter] = useState<FindingState>("open");
  const [days, setDays] = useState(7);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState<string | null>(null);

  const live = session.mode === "live" && session.authenticated;

  const load = useCallback(
    async (signal?: AbortSignal) => {
      if (!live) return;
      try {
        const init = signal ? { signal } : {};
        const [moduleList, signins, list] = await Promise.all([
          readJson("/api/v1/modules", (value) => moduleListSchema.parse(value), init),
          readJson(
            `/api/v1/modules/${MODULE}/summary?days=${days}`,
            (value) => signinSummarySchema.parse(value),
            init,
          ),
          readJson(
            `/api/v1/findings?module=${MODULE}&state=${filter}&limit=100`,
            (value) => findingListSchema.parse(value),
            init,
          ),
        ]);
        setModules(moduleList);
        setSummary(signins);
        setFindings(list.items);
        setError("");
      } catch (loadError) {
        if (signal?.aborted) return;
        setError(loadError instanceof Error ? loadError.message : "Unable to load sign-in audit.");
      }
    },
    [live, days, filter],
  );

  useEffect(() => {
    const controller = new AbortController();
    void load(controller.signal);
    return () => controller.abort();
  }, [load]);

  const collecting = modules?.collection.state === "collecting";
  useEffect(() => {
    if (!collecting) return;
    const timer = setInterval(() => void load(), 60_000);
    return () => clearInterval(timer);
  }, [collecting, load]);

  async function act(label: string, operation: () => Promise<unknown>) {
    setBusy(label);
    setError("");
    try {
      await operation();
      await load();
    } catch (actionError) {
      setError(actionError instanceof Error ? actionError.message : "The request failed.");
    } finally {
      setBusy(null);
    }
  }

  if (session.mode === "demo")
    return (
      <section className="signin-audit" aria-labelledby={titleId}>
        <div className="signin-audit__intro">
          <span className="eyebrow">Sign-in audit</span>
          <h2 id={titleId}>Where your accounts sign in from.</h2>
          <p>
            Connected to a tenant, this view collects interactive sign-ins, learns each account's
            usual places and devices, and flags sign-ins and failure patterns that break them. Demo
            mode shows no sign-in data, because Prism365 does not invent security evidence.
          </p>
        </div>
      </section>
    );

  if (!session.authenticated)
    return (
      <section className="signin-audit" aria-labelledby={titleId}>
        <div className="signin-audit__intro">
          <span className="eyebrow">Sign-in audit</span>
          <h2 id={titleId}>Sign in to see sign-in evidence.</h2>
          <p>Use the tenant connection above. Collection runs on your own delegated access.</p>
        </div>
      </section>
    );

  const module = modules?.modules.find((item) => item.id === MODULE);
  const run = module?.collectors[0]?.lastRun ?? null;
  const missing = module?.missingPermissions ?? [];
  const canConsent = session.enabledFeatures.includes("audit.signins.read");
  const topCountry = Math.max(
    1,
    ...(summary?.countries.map((row) => row.successes + row.failures) ?? [1]),
  );

  return (
    <section className="signin-audit" aria-labelledby={titleId}>
      <h2 id={titleId} className="visually-hidden">
        Sign-in audit
      </h2>
      {error ? (
        <p className="form-error" role="alert">
          {error}
        </p>
      ) : null}

      <section className="panel signin-collection" aria-label="Collection">
        <div className="signin-collection__status">
          <span className={`collection-dot collection-dot--${collecting ? "on" : "off"}`} />
          <div>
            <strong>
              {collecting
                ? "Collecting"
                : modules?.collection.resumes
                  ? "Paused"
                  : "Not collecting"}
            </strong>
            <p aria-live="polite">
              {runLine(run, collecting)}
              {!collecting && modules?.collection.lastStopReason
                ? ` ${stopCopy[modules.collection.lastStopReason] ?? ""}`
                : ""}
              {!collecting && modules?.collection.resumes
                ? " It resumes when the administrator who started it next opens Prism365."
                : ""}
            </p>
            {collecting ? (
              <p className="signin-collection__note">
                Runs on your Microsoft session. A restart or sign-out pauses it, and it resumes when
                you next open Prism365. Stop collecting ends it.
              </p>
            ) : null}
            {summary?.history.phase === "backfill" ? (
              <p className="signin-collection__note">
                Loading the last 30 days, oldest first. History has reached{" "}
                {when(summary.history.through)}. Newer days, including the last 7, fill in as it
                catches up.
              </p>
            ) : null}
            {modules && !modules.collection.durableStorage ? (
              <p className="signin-collection__note">
                Data is held in memory on this server and is lost when it restarts. Set
                PRISM_DATA_DIR to keep it.
              </p>
            ) : null}
            {missing.length > 0 ? (
              <p className="signin-collection__note">
                Needs {missing.join(", ")}.{" "}
                {canConsent ? (
                  <a href="/api/v1/auth/login?feature=audit.signins.read">Grant sign-in access</a>
                ) : (
                  "Enable audit.signins.read in ENABLED_READ_FEATURES."
                )}
              </p>
            ) : null}
          </div>
        </div>
        <div className="signin-collection__actions">
          <button
            className="button button--secondary"
            type="button"
            disabled={busy !== null}
            onClick={() => void act("refresh", async () => {})}
          >
            Refresh
          </button>
          {collecting ? (
            <button
              className="button button--secondary"
              type="button"
              disabled={busy !== null}
              onClick={() =>
                void act("stop", () =>
                  readJson("/api/v1/collection/stop", (value) => value, post(session)),
                )
              }
            >
              {busy === "stop" ? "Stopping…" : "Stop collecting"}
            </button>
          ) : (
            <button
              className="button button--primary"
              type="button"
              disabled={busy !== null || !modules?.collection.available || missing.length > 0}
              onClick={() =>
                void act("start", () =>
                  readJson("/api/v1/collection/start", (value) => value, post(session)),
                )
              }
            >
              <ActivityIcon /> {busy === "start" ? "Starting…" : "Start collecting"}
            </button>
          )}
        </div>
      </section>

      {summary ? (
        <>
          <fieldset className="signin-range">
            <legend className="visually-hidden">Time range</legend>
            {[7, 30].map((value) => (
              <button
                key={value}
                type="button"
                aria-pressed={days === value}
                className={days === value ? "is-active" : ""}
                onClick={() => setDays(value)}
              >
                Last {value} days
              </button>
            ))}
          </fieldset>
          {summary.history.phase === "backfill" &&
          summary.history.through &&
          Date.parse(summary.history.through) < Date.now() - summary.days * 86_400_000 ? (
            <p className="signin-coverage">
              The last {summary.days} days have not loaded yet. History has reached{" "}
              {when(summary.history.through)}; these numbers fill in as it catches up.
            </p>
          ) : null}
          <section className="signin-stats" aria-label={`Sign-ins, last ${summary.days} days`}>
            {[
              ["Successful sign-ins", summary.totals.successes],
              ["Failed password attempts", summary.totals.failures],
              ["Blocked by Conditional Access", summary.totals.blockedByPolicy],
              ["Legacy authentication", summary.totals.legacySuccesses],
              ["Accounts signing in", summary.totals.accounts],
            ].map(([label, value]) => (
              <article className="metric-card signin-stat" key={label}>
                <span>{label}</span>
                <strong>{number.format(Number(value))}</strong>
              </article>
            ))}
          </section>
          <div className="signin-tables">
            <section className="panel signin-table">
              <h3>By country</h3>
              {summary.countries.length === 0 ? (
                <p className="empty-data">No sign-ins in this range yet.</p>
              ) : (
                <table>
                  <caption className="visually-hidden">
                    Sign-ins by country, last {summary.days} days
                  </caption>
                  <thead>
                    <tr>
                      <th scope="col">Country</th>
                      <th scope="col">Successful</th>
                      <th scope="col">Failed</th>
                    </tr>
                  </thead>
                  <tbody>
                    {summary.countries.map((row) => (
                      <tr key={row.country}>
                        <th scope="row">
                          <span>{countryName(row.country)}</span>
                          <span
                            className="signin-bar"
                            aria-hidden="true"
                            style={{
                              width: `${((row.successes + row.failures) / topCountry) * 100}%`,
                            }}
                          />
                        </th>
                        <td>{number.format(row.successes)}</td>
                        <td>{number.format(row.failures)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </section>
            <section className="panel signin-table">
              <h3>Networks with failed attempts</h3>
              {summary.failingNetworks.length === 0 ? (
                <p className="empty-data">No failed password attempts in this range.</p>
              ) : (
                <table>
                  <caption className="visually-hidden">
                    Networks with failed password attempts, last {summary.days} days
                  </caption>
                  <thead>
                    <tr>
                      <th scope="col">Network</th>
                      <th scope="col">Attempts</th>
                      <th scope="col">Accounts</th>
                    </tr>
                  </thead>
                  <tbody>
                    {summary.failingNetworks.map((row) => (
                      <tr key={row.network}>
                        <th scope="row">
                          <code>{row.network}</code>
                        </th>
                        <td>{number.format(row.failures)}</td>
                        <td>{number.format(row.accounts)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </section>
          </div>
          <p className="signin-coverage">
            {summary.coverage.events > 0
              ? `${number.format(summary.coverage.events)} sign-ins stored, from ${when(summary.coverage.firstEvent)} to ${when(summary.coverage.lastEvent)}. Last collected ${when(summary.coverage.lastCollected)}.`
              : "No sign-ins stored yet."}
          </p>
        </>
      ) : null}

      <section className="panel signin-findings" aria-labelledby={`${titleId}-findings`}>
        <div className="panel__header">
          <div>
            <span className="eyebrow">Findings</span>
            <h2 id={`${titleId}-findings`}>What needs a look</h2>
          </div>
          <fieldset className="signin-range">
            <legend className="visually-hidden">Finding state</legend>
            {(["open", "acknowledged", "resolved"] as const).map((state) => (
              <button
                key={state}
                type="button"
                aria-pressed={filter === state}
                className={filter === state ? "is-active" : ""}
                onClick={() => setFilter(state)}
              >
                {stateLabels[state]}
              </button>
            ))}
          </fieldset>
        </div>
        {findings.length === 0 ? (
          <div className="signin-findings__empty">
            <ShieldIcon />
            <p>
              {filter === "open"
                ? "Nothing open. New findings appear here as collection runs."
                : `No ${stateLabels[filter].toLowerCase()} findings.`}
            </p>
          </div>
        ) : (
          <ul className="finding-list">
            {findings.map((finding) => (
              <li key={finding.id}>
                <FindingItem
                  finding={finding}
                  busy={busy !== null}
                  onState={(state) =>
                    void act(finding.id, () =>
                      readJson(
                        `/api/v1/findings/${finding.id}/state`,
                        (value) => findingSchema.parse(value),
                        post(session, { state }),
                      ),
                    )
                  }
                  {...(onOpenPerson ? { onOpenPerson } : {})}
                />
              </li>
            ))}
          </ul>
        )}
      </section>
      <p className="signin-coverage">
        Sign-in records are security evidence, not a measure of anyone's work.
      </p>
    </section>
  );
}

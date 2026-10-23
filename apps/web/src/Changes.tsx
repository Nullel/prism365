import {
  type ChangeEvent,
  type ChangeFeed,
  type ChangeProperty,
  type ChangeTarget,
  changeFeedSchema,
  type Finding,
  type FindingState,
  findingListSchema,
  findingSchema,
  type ModuleList,
  moduleListSchema,
  type PrismSession,
} from "@prism365/contracts";
import { type ReactNode, useCallback, useEffect, useId, useRef, useState } from "react";
import { day, FindingItem, number, post, readJson, when } from "./common.tsx";
import { ShieldIcon } from "./icons.tsx";

// Directory changes: every Microsoft Entra directory audit event Prism365 has collected, newest
// first, with the changes the rules flagged on top. Microsoft keeps these records 7 or 30 days;
// Prism365 keeps them for as long as it runs, so this page is where the older history lives.

export const CHANGES_MODULE = "directory.changes";
const PACK = "audit.directory.read";
export const VALUE_CUT_AT = 400;

export const startHistory = "Collect with the directory audit log enabled to start the history.";

const ranges = [7, 30, 90, 365] as const;

export interface ChangeFilters {
  days: number;
  category: string;
  result: "" | "success" | "failure";
  search: string;
}

export const defaultChangeFilters: ChangeFilters = {
  days: 7,
  category: "",
  result: "",
  search: "",
};

export function changesQuery(filters: ChangeFilters, before: string | null = null): string {
  const query = new URLSearchParams({ days: String(filters.days), limit: "100" });
  if (filters.category) query.set("category", filters.category);
  if (filters.result) query.set("result", filters.result);
  if (filters.search.trim()) query.set("search", filters.search.trim().slice(0, 200));
  if (before) query.set("before", before);
  return `/api/v1/changes?${query}`;
}

// "UserManagement" reads as "User management".
export function categoryLabel(category: string | null): string {
  if (!category) return "No category";
  const words = category.replace(/([a-z])([A-Z])/g, "$1 $2").split(" ");
  return [words[0], ...words.slice(1).map((word) => word.toLowerCase())].join(" ");
}

const resultCopy: Record<string, { label: string; tone: "ok" | "alert" | "neutral" }> = {
  success: { label: "Success", tone: "ok" },
  failure: { label: "Failed", tone: "alert" },
  timeout: { label: "Timed out", tone: "alert" },
};

function Result({ value }: { value: string | null }) {
  const copy = value
    ? (resultCopy[value] ?? { label: value, tone: "neutral" as const })
    : { label: "Not reported", tone: "neutral" as const };
  return (
    <span
      className={`person-tag${copy.tone === "ok" ? " person-tag--ok" : copy.tone === "alert" ? " person-tag--alert" : ""}`}
    >
      {copy.label}
    </span>
  );
}

const clock = (value: string) =>
  new Intl.DateTimeFormat(undefined, { timeStyle: "short" }).format(new Date(value));

// The person a link opens: a user with a directory id.
function personLink(
  kind: string | null,
  id: string | null,
  onOpenPerson: ((userId: string) => void) | undefined,
) {
  return onOpenPerson && id && kind?.toLowerCase() === "user" ? () => onOpenPerson(id) : null;
}

function Name({ label, onOpen }: { label: string; onOpen: (() => void) | null }) {
  return onOpen ? (
    <button className="link-button" type="button" onClick={onOpen}>
      {label}
    </button>
  ) : (
    <span>{label}</span>
  );
}

export function actorLabel(change: ChangeEvent): string {
  const actor = change.initiatedBy;
  if (actor.kind === "user") return actor.name ?? actor.principal ?? "A user";
  if (actor.kind === "app") return actor.name ?? "An application";
  return "Microsoft did not say who";
}

const targetLabel = (target: ChangeTarget) =>
  target.name ?? target.principal ?? target.id ?? target.type ?? "Not named";

function Value({ value, cut }: { value: string | null; cut: boolean }) {
  if (value === null || value === "") return <span className="person-muted">Empty</span>;
  return (
    <>
      {value}
      {cut ? (
        <span className="person-muted"> (cut at {number.format(VALUE_CUT_AT)} characters)</span>
      ) : null}
    </>
  );
}

function Properties({ change }: { change: ChangeEvent }) {
  const rows = change.targets.flatMap((target, index) =>
    target.properties.map((property: ChangeProperty, position) => ({
      key: `${index}-${position}`,
      target: change.targets.length > 1 ? targetLabel(target) : null,
      property,
    })),
  );
  const details = change.details
    .filter((item) => item.key || item.value)
    .map((item, index) => ({ ...item, id: `detail-${index}` }));
  if (rows.length === 0 && details.length === 0 && !change.resultReason) return null;
  return (
    <details className="change-row__details">
      <summary>What changed</summary>
      {change.resultReason ? <p className="signin-coverage">{change.resultReason}</p> : null}
      {rows.length > 0 ? (
        <table className="change-props">
          <caption className="visually-hidden">Values before and after the change</caption>
          <thead>
            <tr>
              <th scope="col">Property</th>
              <th scope="col">Before</th>
              <th scope="col">After</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(({ key, target, property }) => (
              <tr key={key}>
                <th scope="row">
                  {property.name ?? "Not named"}
                  {target ? <span className="person-muted"> on {target}</span> : null}
                </th>
                <td>
                  <Value value={property.oldValue} cut={property.oldCut} />
                </td>
                <td>
                  <Value value={property.newValue} cut={property.newCut} />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      ) : null}
      {details.length > 0 ? (
        <dl className="finding-evidence">
          {details.map((item) => (
            <div key={item.id}>
              <dt>{item.key ?? "Detail"}</dt>
              <dd>{item.value ?? "Empty"}</dd>
            </div>
          ))}
        </dl>
      ) : null}
    </details>
  );
}

// One change: when, who, what, to what, and how it went; the values it changed fold out.
export function ChangeRow({
  change,
  onOpenPerson,
  showDate = false,
  note,
}: {
  change: ChangeEvent;
  onOpenPerson?: (userId: string) => void;
  showDate?: boolean;
  note?: ReactNode;
}) {
  const shown = change.targets
    .slice(0, 3)
    .map((target, index) => ({ target, id: `target-${index}`, first: index === 0 }));
  const more = change.targets.length - shown.length;
  return (
    <li className="change-row">
      <div className="change-row__main">
        <time className="change-row__time" dateTime={change.activityAt}>
          {showDate ? when(change.activityAt) : clock(change.activityAt)}
        </time>
        <span className="change-row__actor">
          <Name
            label={actorLabel(change)}
            onOpen={personLink(change.initiatedBy.kind, change.initiatedBy.id, onOpenPerson)}
          />
        </span>
        <span className="change-row__activity">
          <strong>{change.activity ?? "Activity not reported"}</strong>
          <span className="person-muted"> {categoryLabel(change.category)}</span>
          {note}
        </span>
        <span className="change-row__targets">
          {shown.length === 0 ? (
            <span className="person-muted">No target recorded</span>
          ) : (
            shown.map(({ target, id, first }) => (
              <span key={id}>
                {first ? "" : ", "}
                <Name
                  label={targetLabel(target)}
                  onOpen={personLink(target.type, target.id, onOpenPerson)}
                />
              </span>
            ))
          )}
          {more > 0 ? <span className="person-muted"> and {number.format(more)} more</span> : null}
        </span>
        <span className="change-row__result">
          <Result value={change.result} />
        </span>
      </div>
      <Properties change={change} />
    </li>
  );
}

// Changes grouped by the day they happened, newest first.
export function ChangeFeedList({
  items,
  onOpenPerson,
}: {
  items: readonly ChangeEvent[];
  onOpenPerson?: (userId: string) => void;
}) {
  const days: { key: string; label: string; items: ChangeEvent[] }[] = [];
  for (const change of items) {
    const key = new Date(change.activityAt).toDateString();
    const current = days.at(-1);
    if (current?.key === key) current.items.push(change);
    else days.push({ key, label: day(change.activityAt), items: [change] });
  }
  return (
    <div className="change-feed">
      {days.map((group) => (
        <section key={group.key} aria-label={group.label}>
          <h4 className="change-day">{group.label}</h4>
          <ol className="change-list">
            {group.items.map((change) => (
              <ChangeRow
                key={change.id}
                change={change}
                {...(onOpenPerson ? { onOpenPerson } : {})}
              />
            ))}
          </ol>
        </section>
      ))}
    </div>
  );
}

type ModuleView = ModuleList["modules"][number];
type RunView = NonNullable<ModuleView["collectors"][number]["lastRun"]>;

const failureCopy: Record<string, string> = {
  access_denied:
    "Microsoft refused access. Your account needs a role such as Reports Reader, and the tenant may need Microsoft Entra ID P1 or P2 for the audit log to be read through Graph.",
  credential_unavailable:
    "Your sign-in ended, so collection paused. It resumes when you sign in again.",
  throttled: "Microsoft asked Prism365 to slow down. The next run continues.",
  provider_unavailable: "Microsoft Graph was unavailable. The next run continues.",
  timeout: "The run took too long and stopped. The next run continues where it left off.",
};

export function changesRunLine(run: RunView | null, collecting: boolean): string {
  if (!run)
    return "No collection has run yet. The first run loads the 30 days Microsoft keeps, oldest first.";
  if (run.status === "running")
    return collecting
      ? `Reading the audit log now (started ${when(run.startedAt)}).`
      : `The run started ${when(run.startedAt)} was cut short. Start collecting to continue where it left off.`;
  if (run.status === "failed")
    return `Last run ${when(run.startedAt)} failed. ${failureCopy[run.errorCategory ?? ""] ?? failureCopy[(run.errorCategory ?? "").split(".")[0] ?? ""] ?? "The next run retries; the server log has the reason."}`;
  if (run.status === "abandoned")
    return `The run started ${when(run.startedAt)} was cut short by a restart. Collection continues where it left off.`;
  const note =
    run.skipped > 0
      ? `, ${number.format(run.skipped)} unreadable and left out`
      : run.status === "partial"
        ? ", but some changes could not be read (see the flagged changes)"
        : "";
  return `Last run ${when(run.startedAt)}: ${number.format(run.fetched)} records read, ${number.format(run.inserted)} new${note}.`;
}

export function coverageLine(feed: ChangeFeed | null): string {
  const coverage = feed?.coverage;
  if (!coverage || coverage.phase === "none") return startHistory;
  if (coverage.phase === "history")
    return `Loading the last 30 days, oldest first. History has reached ${when(coverage.through)}; newer changes fill in as it catches up.`;
  return `History kept since ${day(coverage.from)}. Last collected ${when(coverage.lastCollected)}. Microsoft keeps 7 days without Entra ID P1 or P2 and 30 days with it, so days before collection started can be empty.`;
}

export interface ChangesViewProps {
  session: PrismSession;
  modules: ModuleList | null;
  feed: ChangeFeed | null;
  items: readonly ChangeEvent[];
  findings: readonly Finding[];
  filters: ChangeFilters;
  error: string;
  busy: string | null;
  onFilters: (next: Partial<ChangeFilters>) => void;
  onMore: () => void;
  onFindingState: (finding: Finding, state: FindingState) => void;
  onOpenPerson?: (userId: string) => void;
}

// The page with its data: shown by DirectoryChanges and rendered directly in tests.
export function ChangesView(props: ChangesViewProps) {
  const { session, modules, feed, items, findings, filters, busy } = props;
  const titleId = useId();
  const searchId = useId();
  const module = modules?.modules.find((item) => item.id === CHANGES_MODULE);
  const run = module?.collectors[0]?.lastRun ?? null;
  const collecting = modules?.collection.state === "collecting";
  const enabled = session.enabledFeatures.includes(PACK);
  const missing = module?.missingPermissions ?? [];
  const filtered = Boolean(filters.category || filters.result || filters.search.trim());
  const categories = feed?.categories ?? [];

  return (
    <section className="signin-audit changes" aria-labelledby={titleId}>
      <h2 id={titleId} className="visually-hidden">
        Directory changes
      </h2>
      {props.error ? (
        <p className="form-error" role="alert">
          {props.error}
        </p>
      ) : null}

      <section className="panel signin-collection" aria-label="Collection">
        <div className="signin-collection__status">
          <span className={`collection-dot collection-dot--${collecting ? "on" : "off"}`} />
          <div>
            <strong>{collecting ? "Collecting" : "Not collecting"}</strong>
            <p aria-live="polite">
              {!enabled
                ? `This installation does not read the directory audit log. Add ${PACK} to ENABLED_READ_FEATURES, then start collecting on Sign-in audit.`
                : changesRunLine(run, collecting)}
            </p>
            {enabled && missing.length > 0 ? (
              <p className="signin-collection__note">
                Needs {missing.join(", ")}.{" "}
                <a href={`/api/v1/auth/login?feature=${PACK}`}>Grant directory audit access</a>
              </p>
            ) : null}
            <p className="signin-collection__note">{coverageLine(feed)}</p>
          </div>
        </div>
      </section>

      <section className="panel signin-findings" aria-labelledby={`${titleId}-flagged`}>
        <div className="panel__header">
          <h2 id={`${titleId}-flagged`}>Flagged changes</h2>
        </div>
        {findings.length === 0 ? (
          <div className="signin-findings__empty">
            <ShieldIcon />
            <p>
              Nothing flagged. Role assignments, new app credentials, consents, Conditional Access
              and domain changes show up here as they are collected.
            </p>
          </div>
        ) : (
          <ul className="finding-list">
            {findings.map((finding) => (
              <li key={finding.id}>
                <FindingItem
                  finding={finding}
                  busy={busy !== null}
                  onState={(state) => props.onFindingState(finding, state)}
                  {...(props.onOpenPerson ? { onOpenPerson: props.onOpenPerson } : {})}
                />
              </li>
            ))}
          </ul>
        )}
      </section>

      <section className="panel signin-table change-panel" aria-labelledby={`${titleId}-feed`}>
        <div className="people-table__header">
          <h3 id={`${titleId}-feed`}>
            {feed
              ? `${filtered ? "Matching changes" : "Changes"}, last ${number.format(filters.days)} days`
              : "Changes"}
          </h3>
          <fieldset className="signin-range">
            <legend className="visually-hidden">Time range</legend>
            {ranges.map((value) => (
              <button
                key={value}
                type="button"
                aria-pressed={filters.days === value}
                className={filters.days === value ? "is-active" : ""}
                onClick={() => props.onFilters({ days: value })}
              >
                {value === 365 ? "Last year" : `Last ${value} days`}
              </button>
            ))}
          </fieldset>
        </div>
        <div className="audit-filters change-filters">
          <label className="audit-filter">
            <span className="audit-filter__label">Category</span>
            <select
              value={filters.category}
              onChange={(event) => props.onFilters({ category: event.target.value })}
            >
              <option value="">All</option>
              {categories
                .filter((item) => item.category)
                .map((item) => (
                  <option key={item.category} value={item.category ?? ""}>
                    {categoryLabel(item.category)} ({number.format(item.count)})
                  </option>
                ))}
            </select>
          </label>
          <label className="audit-filter">
            <span className="audit-filter__label">Result</span>
            <select
              value={filters.result}
              onChange={(event) =>
                props.onFilters({ result: event.target.value as ChangeFilters["result"] })
              }
            >
              <option value="">All</option>
              <option value="success">Succeeded</option>
              <option value="failure">Did not succeed</option>
            </select>
          </label>
          <label className="visually-hidden" htmlFor={searchId}>
            Search changes
          </label>
          <input
            id={searchId}
            className="people-search__input audit-search"
            type="search"
            value={filters.search}
            maxLength={200}
            placeholder="Search people, apps or activity"
            autoComplete="off"
            onChange={(event) => props.onFilters({ search: event.target.value })}
          />
        </div>
        {!feed ? (
          <p className="empty-data">Loading…</p>
        ) : items.length === 0 ? (
          <p className="empty-data">
            {feed.coverage.phase === "none"
              ? startHistory
              : filtered
                ? "No changes match these filters in this range."
                : feed.coverage.phase === "history"
                  ? "No changes in this range yet. History is still loading."
                  : "No directory changes in this range."}
          </p>
        ) : (
          <>
            <ChangeFeedList
              items={items}
              {...(props.onOpenPerson ? { onOpenPerson: props.onOpenPerson } : {})}
            />
            {feed.next ? (
              <button
                className="button button--secondary person-more"
                type="button"
                disabled={busy !== null}
                onClick={props.onMore}
              >
                {busy === "more" ? "Loading…" : "Show older changes"}
              </button>
            ) : null}
          </>
        )}
      </section>
      <p className="signin-coverage">
        Directory changes record who changed what, as Microsoft reports it. They are not a measure
        of anyone's work.
      </p>
    </section>
  );
}

function Intro({ title, children }: { title: string; children: string }) {
  const id = useId();
  return (
    <section className="signin-audit" aria-labelledby={id}>
      <div className="signin-audit__intro">
        <span className="eyebrow">Directory changes</span>
        <h2 id={id}>{title}</h2>
        <p>{children}</p>
      </div>
    </section>
  );
}

export function DirectoryChanges({
  session,
  onOpenPerson,
}: {
  session: PrismSession;
  onOpenPerson: (userId: string) => void;
}) {
  if (session.mode === "demo")
    return (
      <Intro title="Who changed what in your directory.">
        Connected to a tenant, this view keeps every Microsoft Entra directory audit event for as
        long as Prism365 runs: who changed which account, group, role, app or policy, and what the
        values were before and after. Demo mode shows no changes, because Prism365 does not invent
        audit records.
      </Intro>
    );
  if (!session.authenticated)
    return (
      <Intro title="Sign in to see directory changes.">
        Use the tenant connection above. The change feed reads what collection stored, with your own
        delegated access.
      </Intro>
    );
  return <ChangesWorkspace session={session} onOpenPerson={onOpenPerson} />;
}

function ChangesWorkspace({
  session,
  onOpenPerson,
}: {
  session: PrismSession;
  onOpenPerson: (userId: string) => void;
}) {
  const [filters, setFilters] = useState<ChangeFilters>(defaultChangeFilters);
  // The search box updates `filters` at once; the request waits until typing pauses.
  const [query, setQuery] = useState<ChangeFilters>(defaultChangeFilters);
  const [modules, setModules] = useState<ModuleList | null>(null);
  const [feed, setFeed] = useState<ChangeFeed | null>(null);
  const [items, setItems] = useState<ChangeEvent[]>([]);
  const [findings, setFindings] = useState<Finding[]>([]);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState<string | null>(null);
  // A page that arrives for older filters is dropped.
  const generation = useRef(0);

  useEffect(() => {
    const timer = setTimeout(() => setQuery(filters), filters.search === query.search ? 0 : 250);
    return () => clearTimeout(timer);
  }, [filters, query.search]);

  const load = useCallback(
    async (signal?: AbortSignal) => {
      generation.current += 1;
      const mine = generation.current;
      const init = signal ? { signal } : {};
      try {
        const [moduleList, page, flagged] = await Promise.all([
          readJson("/api/v1/modules", (value) => moduleListSchema.parse(value), init),
          readJson(changesQuery(query), (value) => changeFeedSchema.parse(value), init),
          readJson(
            `/api/v1/findings?module=${CHANGES_MODULE}&state=open&limit=100`,
            (value) => findingListSchema.parse(value),
            init,
          ),
        ]);
        if (mine !== generation.current) return;
        setModules(moduleList);
        setFeed(page);
        setItems(page.items);
        setFindings(flagged.items);
        setError("");
      } catch (loadError) {
        if (signal?.aborted) return;
        setError(
          loadError instanceof Error ? loadError.message : "Directory changes could not be loaded.",
        );
      }
    },
    [query],
  );

  useEffect(() => {
    const controller = new AbortController();
    void load(controller.signal);
    return () => controller.abort();
  }, [load]);

  async function more() {
    if (!feed?.next) return;
    const mine = generation.current;
    setBusy("more");
    try {
      const page = await readJson(changesQuery(query, feed.next), (value) =>
        changeFeedSchema.parse(value),
      );
      if (mine !== generation.current) return;
      setFeed(page);
      setItems((current) => [...current, ...page.items]);
    } catch (loadError) {
      setError(
        loadError instanceof Error ? loadError.message : "Older changes could not be loaded.",
      );
    } finally {
      setBusy(null);
    }
  }

  async function changeState(finding: Finding, state: FindingState) {
    setBusy(finding.id);
    try {
      await readJson(
        `/api/v1/findings/${finding.id}/state`,
        (value) => findingSchema.parse(value),
        post(session, { state }),
      );
      await load();
    } catch (stateError) {
      setError(
        stateError instanceof Error ? stateError.message : "The finding could not be updated.",
      );
    } finally {
      setBusy(null);
    }
  }

  return (
    <ChangesView
      session={session}
      modules={modules}
      feed={feed}
      items={items}
      findings={findings}
      filters={filters}
      error={error}
      busy={busy}
      onFilters={(next) => setFilters((current) => ({ ...current, ...next }))}
      onMore={() => void more()}
      onFindingState={(finding, state) => void changeState(finding, state)}
      onOpenPerson={onOpenPerson}
    />
  );
}

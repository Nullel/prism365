import {
  type AuditCatalog,
  type AuditCatalogEntry,
  type AuditCell,
  type AuditColumn,
  type AuditDefinition,
  type AuditRow,
  type AuditRun,
  type AuditRunDetail,
  type AuditScheduleFrequency,
  type AuditScheduleLimit,
  type AuditScheduleSettings,
  type AuditSettings,
  type AuditSummaryItem,
  type AuditTrend,
  type AuditTrendDifference,
  type AuditTrendPoint,
  type AuditTrendTile,
  auditCatalogSchema,
  auditFacetOptions,
  auditRunDetailSchema,
  auditRunListSchema,
  auditScheduleSettingsSchema,
  auditSettingsSchema,
  auditStartSchema,
  auditTrendSchema,
  MAX_AUDIT_SCOPE_DOMAINS,
  nextScheduledRun,
  type PrismSession,
  queryAuditRows,
} from "@prism365/contracts";
import { type ReactNode, useCallback, useEffect, useId, useMemo, useRef, useState } from "react";
import { AccountName, day, number, post, readJson, when } from "./common.tsx";
import { CopyButton, useCopy } from "./CopyButton.tsx";
import { CheckIcon, CopyIcon, DownloadIcon } from "./icons.tsx";

// Audits: checks an administrator runs when they need them. The catalog lists every audit with
// its last result; an audit's page shows a run as a table the audit describes itself, so a new
// audit needs no new page code.

// Each open results page polls while a run is in progress; the API allows 60 requests a minute
// per session, so this leaves room for several tabs. A failed poll backs off instead of stopping.
const POLL_MILLISECONDS = 5_000;
const MAX_POLL_MILLISECONDS = 30_000;
const PAGE_ROWS = 200;

export function ago(value: string, now = Date.now()): string {
  const seconds = Math.max(0, Math.round((now - Date.parse(value)) / 1_000));
  if (seconds < 60) return "just now";
  const minutes = Math.round(seconds / 60);
  if (minutes < 60) return `${minutes} min ago`;
  const hours = Math.round(minutes / 60);
  if (hours < 24) return `${hours} h ago`;
  return day(value);
}

function duration(run: AuditRun): string | null {
  if (!run.finishedAt) return null;
  const seconds = Math.max(
    0,
    Math.round((Date.parse(run.finishedAt) - Date.parse(run.startedAt)) / 1_000),
  );
  return seconds < 90 ? `${seconds} s` : `${Math.round(seconds / 60)} min`;
}

// The line under an audit's name: what its last run found, or why there is nothing.
export function runHeadline(run: AuditRun | null, definition: AuditDefinition): string {
  if (!run) return "Not run yet.";
  if (run.status === "running") return `Running since ${when(run.startedAt)}.`;
  if (run.status === "failed")
    return `Last run failed ${ago(run.startedAt)}: ${run.error?.message ?? "no reason recorded."}`;
  const notable =
    run.summary.find((item) => item.tone !== "neutral" && item.tone !== "good" && item.value > 0) ??
    null;
  const total = `${number.format(run.rowCount)} ${run.rowCount === 1 ? definition.subject.one : definition.subject.many}`;
  return [
    `Ran ${ago(run.startedAt)}`,
    total,
    notable ? `${number.format(notable.value)} ${notable.label.toLowerCase()}` : null,
    run.complete ? null : "partial",
  ]
    .filter(Boolean)
    .join(" · ");
}

function AccessAction({ entry }: { entry: AuditCatalogEntry }) {
  const { availability } = entry;
  if (availability.state === "missing_permission")
    return availability.grantFeature ? (
      <a
        className="button button--secondary"
        href={`/api/v1/auth/login?feature=${encodeURIComponent(availability.grantFeature)}`}
      >
        Grant access
      </a>
    ) : null;
  return null;
}

function accessText(entry: AuditCatalogEntry): string | null {
  const { availability } = entry;
  if (availability.state === "missing_permission")
    return `Needs ${availability.missing.join(" and ")}, which your session has not granted.`;
  if (availability.state === "not_enabled")
    return `Not enabled here. Add ${(availability.enable.length > 0 ? availability.enable : availability.missing).join(", ")} to ENABLED_READ_FEATURES.`;
  return null;
}

// The audit scope: which account domains the account audits cover. Saved per tenant; an audit
// reads it when a run starts, so a change applies from the next run.

export function scopeSummary(domains: readonly string[]): string {
  if (domains.length === 0) return "All domains";
  return `${plural(domains.length, "domain", "domains")}: ${domains.join(", ")}`;
}

// The scope a run used, and whether the saved scope has changed since. Only a run made by the
// current version of the audit records its scope, and the comparison needs the saved scope.
export function RunScopeTags({
  definition,
  run,
  saved,
}: {
  definition: AuditDefinition;
  run: AuditRun;
  saved: readonly string[] | null;
}) {
  if (!definition.scoped || run.auditVersion !== definition.version) return null;
  const used = run.scope ?? [];
  return (
    <>
      {used.length > 0 ? (
        <span className="person-tag" title={used.join(", ")}>
          {scopeLimit(used.length)}
        </span>
      ) : null}
      {saved !== null && run.status === "succeeded" && saved.join(",") !== used.join(",") ? (
        <span className="person-tag">The saved scope has changed since this run</span>
      ) : null}
    </>
  );
}

export function scopeLimit(count: number): string {
  return `Limited to ${plural(count, "domain", "domains")}`;
}

const firstRunHint = "Run Inactive accounts once to list your domains.";

function titles(values: readonly string[]): string {
  if (values.length < 2) return values.join("");
  return `${values.slice(0, -1).join(", ")} and ${values.at(-1)}`;
}

export function AuditScopeEditor({
  domains,
  selected,
  saving,
  error,
  onToggle,
  onAll,
  onSave,
  onCancel,
}: {
  // accounts is null for a saved domain the latest results no longer list.
  domains: readonly { domain: string; accounts: number | null }[];
  selected: ReadonlySet<string>;
  saving: boolean;
  error: string;
  onToggle: (domain: string, checked: boolean) => void;
  onAll: () => void;
  onSave: () => void;
  onCancel: () => void;
}) {
  return (
    <div className="audit-scope__editor">
      <fieldset className="audit-scope__choices">
        <legend className="visually-hidden">Domains the audits cover</legend>
        <label className="check-row check-row--compact">
          <input type="checkbox" checked={selected.size === 0} onChange={onAll} />
          <span>All domains</span>
        </label>
        {domains.length === 0 ? (
          <p className="muted-label">{firstRunHint}</p>
        ) : (
          <div className="service-grid">
            {domains.map((item) => (
              <label key={item.domain} className="check-row check-row--compact">
                <input
                  type="checkbox"
                  checked={selected.has(item.domain)}
                  onChange={(event) => onToggle(item.domain, event.target.checked)}
                />
                <span>
                  {item.domain}
                  <span className="audit-scope__count">
                    {item.accounts === null
                      ? "not in the latest results"
                      : plural(item.accounts, "account", "accounts")}
                  </span>
                </span>
              </label>
            ))}
          </div>
        )}
      </fieldset>
      {error ? (
        <p className="form-error" role="alert">
          {error}
        </p>
      ) : null}
      <div className="finding__actions">
        <button className="button button--primary" type="button" disabled={saving} onClick={onSave}>
          {saving ? "Saving…" : "Save"}
        </button>
        <button
          className="button button--secondary"
          type="button"
          disabled={saving}
          onClick={onCancel}
        >
          Cancel
        </button>
      </div>
    </div>
  );
}

export function AuditScopeSection({
  settings,
  applies = [],
  onSave,
}: {
  settings: AuditSettings;
  // Titles of the audits that use the scope.
  applies?: readonly string[];
  onSave: (domains: string[]) => Promise<void>;
}) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState<ReadonlySet<string>>(new Set());
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const editor = useRef<HTMLDivElement>(null);
  const change = useRef<HTMLButtonElement>(null);
  const wasEditing = useRef(false);
  const saved = settings.scope.domains;
  const counts = new Map(settings.domains.map((item) => [item.domain, item.accounts]));
  // Saved domains stay in the list even when the latest results no longer hold them.
  const choices = [...new Set([...saved, ...counts.keys()])]
    .sort()
    .map((domain) => ({ domain, accounts: counts.get(domain) ?? null }));

  // Keyboard users land in the list when it opens and back on Change when it closes.
  useEffect(() => {
    if (editing) editor.current?.querySelector("input")?.focus();
    else if (wasEditing.current) change.current?.focus();
    wasEditing.current = editing;
  }, [editing]);

  function open() {
    setDraft(new Set(saved));
    setError("");
    setEditing(true);
  }

  async function save() {
    if (draft.size > MAX_AUDIT_SCOPE_DOMAINS) {
      setError(`Choose at most ${MAX_AUDIT_SCOPE_DOMAINS} domains, or All domains.`);
      return;
    }
    setSaving(true);
    setError("");
    try {
      await onSave([...draft].sort());
      setEditing(false);
    } catch (problem) {
      setError(problem instanceof Error ? problem.message : "The scope could not be saved.");
    } finally {
      setSaving(false);
    }
  }

  return (
    <section className="panel audit-group" aria-label="Audit scope">
      <h2 className="audit-group__title">Audit scope</h2>
      {editing ? (
        <div ref={editor}>
          <AuditScopeEditor
            domains={choices}
            selected={draft}
            saving={saving}
            error={error}
            onToggle={(domain, checked) =>
              setDraft((current) => {
                const next = new Set(current);
                if (checked) next.add(domain);
                else next.delete(domain);
                return next;
              })
            }
            onAll={() => setDraft(new Set())}
            onSave={() => void save()}
            onCancel={() => setEditing(false)}
          />
        </div>
      ) : (
        <div className="audit-scope">
          <div>
            <p>{scopeSummary(saved)}</p>
            {applies.length > 0 ? (
              <p className="muted-label">Applies to {titles(applies)} on their next run.</p>
            ) : null}
            {settings.domains.length === 0 ? <p className="muted-label">{firstRunHint}</p> : null}
          </div>
          <button ref={change} className="button button--secondary" type="button" onClick={open}>
            Change
          </button>
        </div>
      )}
    </section>
  );
}

export function AuditCatalogList({
  catalog,
  busy,
  onOpen,
  onRun,
  scope,
  schedule = null,
}: {
  catalog: AuditCatalog;
  busy: string | null;
  onOpen: (auditId: string) => void;
  onRun: (auditId: string) => void;
  // The audit scope section, shown above the audits when the page has loaded it.
  scope?: ReactNode;
  // How often each audit runs on its own, when the page has loaded it.
  schedule?: Readonly<Record<string, AuditScheduleFrequency>> | null;
}) {
  const demo = catalog.mode === "demo";
  const categories = [...new Set(catalog.audits.map((entry) => entry.definition.category))];
  return (
    <div className="audits">
      {demo ? (
        <p className="audit-note">
          Demo mode lists the audits but cannot run them, because every audit reads a connected
          tenant live.
        </p>
      ) : null}
      {scope}
      {categories.map((category) => (
        <section className="panel audit-group" key={category} aria-label={category}>
          <h2 className="audit-group__title">{category}</h2>
          <ul className="audit-list">
            {catalog.audits
              .filter((entry) => entry.definition.category === category)
              .map((entry) => {
                const { definition } = entry;
                const running = Boolean(entry.runningId) || busy === definition.id;
                const access = accessText(entry);
                const tag = scheduleTag(schedule?.[definition.id]);
                return (
                  <li className="audit-item" key={definition.id}>
                    <div className="audit-item__body">
                      <h3>
                        <button
                          className="link-button link-button--strong"
                          type="button"
                          onClick={() => onOpen(definition.id)}
                        >
                          {definition.title}
                        </button>
                        {tag ? <span className="person-tag audit-item__tag">{tag}</span> : null}
                      </h3>
                      <p>{definition.description}</p>
                      <p
                        className={`audit-item__meta${entry.lastRun?.status === "failed" ? " audit-item__meta--failed" : ""}`}
                      >
                        {busy === definition.id && entry.lastRun?.status !== "running"
                          ? "Starting…"
                          : runHeadline(entry.lastRun, definition)}
                      </p>
                      {access ? <p className="audit-item__access">{access}</p> : null}
                    </div>
                    <div className="audit-item__actions">
                      <AccessAction entry={entry} />
                      {entry.lastRun ? (
                        <button
                          className="button button--secondary"
                          type="button"
                          onClick={() => onOpen(definition.id)}
                        >
                          View results
                        </button>
                      ) : null}
                      <button
                        className="button button--primary"
                        type="button"
                        disabled={demo || running || entry.availability.state !== "available"}
                        onClick={() => onRun(definition.id)}
                      >
                        {running ? "Running…" : entry.lastRun ? "Run again" : "Run audit"}
                      </button>
                    </div>
                  </li>
                );
              })}
          </ul>
        </section>
      ))}
    </div>
  );
}

function plural(value: number, one: string, many: string) {
  return `${number.format(value)} ${value === 1 ? one : many}`;
}

function isPerson(
  cell: AuditCell | undefined,
): cell is { name: string | null; principal: string | null } {
  return typeof cell === "object" && cell !== null && !Array.isArray(cell);
}

// A value that is, whole, one email address. Text columns that hold one get a copy button.
const EMAIL = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function isEmail(cell: AuditCell | undefined): cell is string {
  return typeof cell === "string" && EMAIL.test(cell);
}

// The address to copy for a row: its first email-valued text cell (a guest's real address, which
// the #EXT# principal name is not), otherwise the account's user principal name. Takes every
// column of the audit, shown or hidden.
function rowEmail(columns: readonly AuditColumn[], row: AuditRow): string | null {
  for (const column of columns) {
    const cell = row.cells[column.key];
    if (column.kind === "text" && isEmail(cell)) return cell;
  }
  for (const column of columns) {
    const cell = row.cells[column.key];
    if (column.kind === "person" && isPerson(cell) && cell.principal?.trim())
      return cell.principal.trim();
  }
  return null;
}

// One address per row, in the order of the rows, each address once (letter case aside). Rows with
// no address are left out.
export function rowEmails(columns: readonly AuditColumn[], rows: readonly AuditRow[]): string[] {
  const seen = new Set<string>();
  const emails: string[] = [];
  for (const row of rows) {
    const email = rowEmail(columns, row);
    if (!email || seen.has(email.toLowerCase())) continue;
    seen.add(email.toLowerCase());
    emails.push(email);
  }
  return emails;
}

const alertValues = new Set([
  "Disabled",
  "Not accepted",
  "Expired, in grace period",
  "Suspended",
  "Locked out",
]);

function CellValue({
  column,
  cell,
  personId,
  onOpenPerson,
}: {
  column: AuditColumn;
  cell: AuditCell | undefined;
  personId: string | null;
  onOpenPerson: (userId: string) => void;
}) {
  if (column.kind === "person") {
    const person = isPerson(cell) ? cell : { name: null, principal: null };
    const name = person.name ?? person.principal ?? "Unnamed account";
    return (
      <AccountName
        name={name}
        principal={person.principal}
        onOpen={personId ? () => onOpenPerson(personId) : null}
      />
    );
  }
  if (cell === null || cell === undefined || cell === "")
    return (
      <span className="audit-empty">
        {column.empty ?? (column.kind === "date" ? "Never" : "None")}
      </span>
    );
  if (column.kind === "days" && typeof cell === "number")
    return <span className="audit-days">{cell === 0 ? "Today" : plural(cell, "day", "days")}</span>;
  if (column.kind === "date" && typeof cell === "string")
    return (
      <time dateTime={cell} title={when(cell)}>
        {day(cell)}
      </time>
    );
  if (column.kind === "boolean" && typeof cell === "boolean") return <>{cell ? "Yes" : "No"}</>;
  if (Array.isArray(cell)) {
    if (cell.length === 0) return <span className="audit-empty">None</span>;
    const shown = cell.slice(0, 2).join(", ");
    return (
      <span title={cell.join("\n")}>
        {shown}
        {cell.length > 2 ? <span className="audit-more"> +{cell.length - 2}</span> : null}
      </span>
    );
  }
  if (typeof cell === "number")
    return <span className={cell < 0 ? "audit-negative" : undefined}>{number.format(cell)}</span>;
  if (column.kind === "enum" && typeof cell === "string" && alertValues.has(cell))
    return <span className="person-tag person-tag--alert">{cell}</span>;
  if (column.kind === "text" && isEmail(cell))
    return (
      <span className="copy-line">
        <span className="copy-line__text">{cell}</span>
        <CopyButton text={cell} />
      </span>
    );
  return <>{String(cell)}</>;
}

const numeric = (column: AuditColumn) => column.kind === "number" || column.kind === "days";

function SummaryTiles({
  items,
  view,
  onView,
}: {
  items: AuditSummaryItem[];
  view: string;
  onView: (view: string) => void;
}) {
  if (items.length === 0) return null;
  return (
    <section className="audit-tiles" aria-label="Summary">
      {items.map((item) =>
        item.view ? (
          <button
            key={item.label}
            type="button"
            className={`audit-tile audit-tile--${item.tone}${view === item.view ? " is-active" : ""}`}
            aria-pressed={view === item.view}
            onClick={() => item.view && onView(item.view)}
          >
            <span className="audit-tile__label">{item.label}</span>
            <strong className="audit-tile__value">{number.format(item.value)}</strong>
          </button>
        ) : (
          <div key={item.label} className={`audit-tile audit-tile--${item.tone}`}>
            <span className="audit-tile__label">{item.label}</span>
            <strong className="audit-tile__value">{number.format(item.value)}</strong>
          </div>
        ),
      )}
    </section>
  );
}

export function exportHref(
  run: AuditRun,
  query: {
    view: string;
    search: string;
    sort: { column: string; direction: "asc" | "desc" };
    facets: Record<string, string>;
  },
): string {
  const params = new URLSearchParams({ view: query.view });
  if (query.search.trim()) params.set("search", query.search.trim());
  params.set("sort", query.sort.column);
  params.set("direction", query.sort.direction);
  for (const [key, value] of Object.entries(query.facets)) if (value) params.set(`f.${key}`, value);
  return `/api/v1/audit-runs/${encodeURIComponent(run.id)}/export.csv?${params.toString()}`;
}

// Copies the address of every row the current view, search and filters leave, not only the rows
// the page shows so far. The label says how many were copied, then goes back.
function CopyEmailsButton({ emails }: { emails: readonly string[] }) {
  const { status, copy } = useCopy(2_000);
  // The count at the moment of copying, so a filter change does not rewrite the confirmation.
  const [count, setCount] = useState(0);
  const label =
    status === "copied"
      ? `Copied ${plural(count, "email", "emails")}`
      : status === "failed"
        ? "Copy failed"
        : "Copy emails";
  return (
    <>
      <button
        className="button button--secondary"
        type="button"
        disabled={emails.length === 0}
        onClick={() => {
          setCount(emails.length);
          void copy(emails.join("\n"));
        }}
      >
        {status === "copied" ? <CheckIcon /> : <CopyIcon />} {label}
      </button>
      <span className="visually-hidden" aria-live="polite">
        {status === "idle" ? "" : label}
      </span>
    </>
  );
}

export function ResultTable({
  definition,
  run,
  onOpenPerson,
}: {
  definition: AuditDefinition;
  run: AuditRunDetail;
  onOpenPerson: (userId: string) => void;
}) {
  const searchId = useId();
  const [view, setView] = useState(definition.defaultView);
  const [facets, setFacets] = useState<Record<string, string>>({});
  const [search, setSearch] = useState("");
  const [sort, setSort] = useState(definition.defaultSort);
  const [shown, setShown] = useState<Set<string>>(
    () =>
      new Set(definition.columns.filter((column) => !column.hidden).map((column) => column.key)),
  );
  // How many rows are shown, for the current filters; changing a filter starts again at one page.
  const [page, setPage] = useState({ key: "", limit: PAGE_ROWS });
  const options = useMemo(() => auditFacetOptions(definition, run.rows), [definition, run.rows]);
  const rows = useMemo(
    () => queryAuditRows(definition, run.rows, { view, facets, search, sort }),
    [definition, run.rows, view, facets, search, sort],
  );
  // Every column counts here, shown or not: a guest's Email column can be hidden and still be
  // the address to copy.
  const hasPeople = definition.columns.some((column) => column.kind === "person");
  const emails = useMemo(
    () => (hasPeople ? rowEmails(definition.columns, rows) : []),
    [definition.columns, hasPeople, rows],
  );
  const queryKey = JSON.stringify([view, facets, search, sort]);
  const limit = page.key === queryKey ? page.limit : PAGE_ROWS;
  const columns = definition.columns.filter((column) => shown.has(column.key));
  const facetColumns = definition.columns.filter(
    (column) => column.facet && (options[column.key]?.length ?? 0) > 1,
  );
  const activeView = definition.views.find((item) => item.id === view);
  const toggleSort = (column: AuditColumn) =>
    setSort((current) =>
      current.column === column.key
        ? { column: column.key, direction: current.direction === "asc" ? "desc" : "asc" }
        : {
            column: column.key,
            direction: ["days", "number", "boolean"].includes(column.kind) ? "desc" : "asc",
          },
    );

  return (
    <>
      <SummaryTiles items={run.summary} view={view} onView={setView} />
      <section className="panel audit-table" aria-label="Results">
        <div className="audit-toolbar">
          <fieldset className="signin-range audit-views">
            <legend className="visually-hidden">Show</legend>
            {definition.views.map((item) => (
              <button
                key={item.id}
                type="button"
                className={item.id === view ? "is-active" : undefined}
                aria-pressed={item.id === view}
                onClick={() => setView(item.id)}
              >
                {item.label}
              </button>
            ))}
          </fieldset>
          <div className="audit-filters">
            {facetColumns.map((column) => (
              <label className="audit-filter" key={column.key}>
                <span className="audit-filter__label">{column.label}</span>
                <select
                  value={facets[column.key] ?? ""}
                  onChange={(event) =>
                    setFacets((current) => ({ ...current, [column.key]: event.target.value }))
                  }
                >
                  <option value="">All</option>
                  {(options[column.key] ?? []).map((value) => (
                    <option key={value} value={value}>
                      {value}
                    </option>
                  ))}
                </select>
              </label>
            ))}
            <label className="visually-hidden" htmlFor={searchId}>
              Search results
            </label>
            <input
              id={searchId}
              className="people-search__input audit-search"
              type="search"
              value={search}
              placeholder="Search"
              autoComplete="off"
              onChange={(event) => setSearch(event.target.value)}
            />
            <details className="audit-columns">
              <summary className="button button--secondary">Columns</summary>
              <div className="audit-columns__menu">
                {definition.columns.map((column) => (
                  <label key={column.key} className="check-row check-row--compact">
                    <input
                      type="checkbox"
                      checked={shown.has(column.key)}
                      disabled={column.kind === "person"}
                      onChange={(event) =>
                        setShown((current) => {
                          const next = new Set(current);
                          if (event.target.checked) next.add(column.key);
                          else next.delete(column.key);
                          return next;
                        })
                      }
                    />
                    <span>{column.label}</span>
                  </label>
                ))}
              </div>
            </details>
            {hasPeople ? <CopyEmailsButton emails={emails} /> : null}
            <a
              className="button button--secondary"
              href={exportHref(run, { view, search, sort, facets })}
              download
            >
              <DownloadIcon /> Export CSV
            </a>
          </div>
        </div>
        <p className="audit-count" aria-live="polite">
          {plural(rows.length, definition.subject.one, definition.subject.many)}
          {activeView && activeView.tags.length > 0 ? ` in “${activeView.label}”` : ""}
          {rows.length !== run.rows.length ? ` of ${number.format(run.rows.length)}` : ""}
        </p>
        {rows.length === 0 ? (
          <p className="empty-data audit-empty-state">
            {run.rows.length === 0
              ? `This run found no ${definition.subject.many}.`
              : `No ${definition.subject.many} match. Try another range or clear the filters.`}
          </p>
        ) : (
          <div className="audit-scroll">
            <table>
              <caption className="visually-hidden">{definition.title} results</caption>
              <thead>
                <tr>
                  {columns.map((column) => {
                    const active = sort.column === column.key;
                    return (
                      <th
                        scope="col"
                        key={column.key}
                        className={numeric(column) ? "audit-numeric" : undefined}
                        aria-sort={
                          active
                            ? sort.direction === "asc"
                              ? "ascending"
                              : "descending"
                            : undefined
                        }
                      >
                        <button
                          type="button"
                          className={`audit-sort${active ? " is-active" : ""}`}
                          title={column.help ?? undefined}
                          onClick={() => toggleSort(column)}
                        >
                          {column.label}
                          <span className="audit-sort__arrow" aria-hidden="true">
                            {active ? (sort.direction === "asc" ? "↑" : "↓") : ""}
                          </span>
                        </button>
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {rows.slice(0, limit).map((row) => (
                  <tr key={row.key}>
                    {columns.map((column) =>
                      column.kind === "person" ? (
                        <th scope="row" key={column.key}>
                          <CellValue
                            column={column}
                            cell={row.cells[column.key]}
                            personId={row.personId}
                            onOpenPerson={onOpenPerson}
                          />
                        </th>
                      ) : (
                        <td
                          key={column.key}
                          className={numeric(column) ? "audit-numeric" : undefined}
                        >
                          <CellValue
                            column={column}
                            cell={row.cells[column.key]}
                            personId={row.personId}
                            onOpenPerson={onOpenPerson}
                          />
                        </td>
                      ),
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        {rows.length > limit ? (
          <div className="audit-more-rows">
            <button
              className="button button--secondary"
              type="button"
              onClick={() => setPage({ key: queryKey, limit: limit + PAGE_ROWS })}
            >
              Show {number.format(Math.min(PAGE_ROWS, rows.length - limit))} more
            </button>
            <span className="muted-label">
              Showing {number.format(limit)} of {number.format(rows.length)}
            </span>
          </div>
        ) : null}
      </section>
    </>
  );
}

export function runLabel(run: AuditRun): string {
  const notes = [
    run.status === "running" ? "running" : run.status === "failed" ? "failed" : null,
    run.trigger === "schedule" ? "scheduled" : null,
  ].filter(Boolean);
  return `${when(run.startedAt)}${notes.length > 0 ? ` (${notes.join(", ")})` : ""}`;
}

// Schedules (docs/audits.md, section 12): an audit can run on its own, daily or weekly, while
// collection is on, with the collecting administrator's sign-in.

const frequencies = ["off", "daily", "weekly"] as const;
const frequencyLabels: Record<AuditScheduleFrequency, string> = {
  off: "Off",
  daily: "Daily",
  weekly: "Weekly",
};

// The catalog's tag for a scheduled audit; null for one that only runs by hand.
export function scheduleTag(frequency: AuditScheduleFrequency | undefined): string | null {
  return frequency && frequency !== "off" ? frequencyLabels[frequency] : null;
}

// Said when a scheduled run reads less than a manual run with every permission would: it reads
// only with what the administrator who scheduled it had granted then. null when it reads it all.
export function scheduleLimitNote(limit: AuditScheduleLimit | null): string | null {
  if (!limit || limit.without.length === 0) return null;
  const without = titles(limit.without);
  return limit.by
    ? `Runs with the permissions ${limit.by} had when scheduling, so it does not read with ${without}.`
    : `Runs with its required permissions only, so it does not read with ${without}: it was scheduled before Prism365 kept the permissions of whoever scheduled it.`;
}

// Saves one audit's schedule. Only that audit is sent: the server leaves the others as they
// are, so another page changing another audit is never undone by a stale copy here.
export function saveAuditSchedule(
  session: PrismSession,
  auditId: string,
  frequency: AuditScheduleFrequency,
): Promise<AuditScheduleSettings> {
  return readJson(
    "/api/v1/settings/audit-schedule",
    auditScheduleSettingsSchema.parse,
    post(session, { audits: { [auditId]: frequency } }),
  );
}

// Said when a scheduled audit's newest run is dated in the future (the clock was turned back):
// the schedule counts from that start, so nothing runs until a full period after it. null when
// the next run is not held back that way.
export function futureRunNote(
  frequency: AuditScheduleFrequency,
  lastStart: string | null,
  now = Date.now(),
): string | null {
  if (lastStart === null || Date.parse(lastStart) <= now) return null;
  const next = nextScheduledRun(frequency, lastStart);
  return next ? `Next run waits for ${when(next)}: the last run is dated in the future.` : null;
}

export function AuditScheduleControl({
  frequency,
  collecting,
  saving,
  error,
  lastStart = null,
  limit = null,
  disabled = false,
  onChange,
}: {
  frequency: AuditScheduleFrequency;
  collecting: boolean;
  saving: boolean;
  error: string;
  // When the audit's newest run of any kind started, which the schedule counts from.
  lastStart?: string | null;
  // Set when a scheduled run reads less than a fully granted manual run would.
  limit?: AuditScheduleLimit | null;
  disabled?: boolean;
  onChange: (frequency: AuditScheduleFrequency) => void;
}) {
  const labelId = useId();
  const held = futureRunNote(frequency, lastStart);
  const limited = frequency === "off" ? null : scheduleLimitNote(limit);
  return (
    <section className="audit-schedule" aria-labelledby={labelId}>
      <div className="audit-schedule__row">
        <span className="audit-filter__label" id={labelId}>
          Run automatically
        </span>
        <fieldset className="signin-range" aria-labelledby={labelId} disabled={disabled || saving}>
          {frequencies.map((value) => (
            <button
              key={value}
              type="button"
              className={value === frequency ? "is-active" : undefined}
              aria-pressed={value === frequency}
              onClick={() => {
                if (value !== frequency) onChange(value);
              }}
            >
              {frequencyLabels[value]}
            </button>
          ))}
        </fieldset>
        <span className="muted-label" aria-live="polite">
          {saving ? "Saving…" : ""}
        </span>
      </div>
      <p className="muted-label">
        Runs while collection is on, with the collecting administrator's sign-in. You can schedule
        an audit only if you could run it yourself, and it never reads more than you could.
      </p>
      {limited ? <p className="muted-label">{limited}</p> : null}
      {!collecting && frequency !== "off" ? (
        <p className="audit-schedule__off">
          Collection is off, so this audit will not run on its own until someone starts it.
        </p>
      ) : null}
      {held ? <p className="audit-schedule__off">{held}</p> : null}
      {error ? (
        <p className="form-error" role="alert">
          {error}
        </p>
      ) : null}
    </section>
  );
}

// Trends (docs/audits.md, section 11): each tile's value over the audit's successful runs. A
// point that does not compare with the newest run is drawn hollow and says why.

const differenceText: Record<AuditTrendDifference, string> = {
  version: "another version of the audit",
  scope: "another audit scope",
  incomplete: "partial result",
};

// A point's value for one tile; null when that run had no such tile. Own keys only, so a tile
// id such as "constructor" never reads something inherited.
export function trendValue(point: AuditTrendPoint, tileId: string): number | null {
  if (!Object.hasOwn(point.tiles, tileId)) return null;
  const value = point.tiles[tileId];
  return typeof value === "number" ? value : null;
}

// The tile shown first: the first critical or warning tile of the newest run, else the first.
export function defaultTrendTile(tiles: readonly AuditTrendTile[]): string | null {
  return (
    (tiles.find((tile) => tile.tone === "critical" || tile.tone === "warning") ?? tiles[0])?.id ??
    null
  );
}

// Why one point does not compare with the newest run, as a sentence; empty when it does.
export function trendNote(point: AuditTrendPoint): string {
  const text = point.notComparable.map((difference) => differenceText[difference]).join(", ");
  return text ? `${text.charAt(0).toUpperCase()}${text.slice(1)}.` : "";
}

// What the hollow points among these are, counted by reason; null when there are none.
export function trendHollowNote(points: readonly AuditTrendPoint[]): string | null {
  const count = (difference: AuditTrendDifference) =>
    points.filter((point) => point.notComparable.includes(difference)).length;
  const version = count("version");
  const scope = count("scope");
  const incomplete = count("incomplete");
  const parts = [
    version > 0 ? `${number.format(version)} by another version of the audit` : null,
    scope > 0 ? `${number.format(scope)} with another audit scope` : null,
    incomplete > 0 ? plural(incomplete, "partial result", "partial results") : null,
  ].filter((part): part is string => part !== null);
  return parts.length > 0
    ? `Hollow points do not compare directly with the latest run: ${titles(parts)}.`
    : null;
}

const CHART_WIDTH = 600;
const CHART_HEIGHT = 160;
const CHART_PAD = 10;
const round = (value: number) => Math.round(value * 10) / 10;

// Where each point sits in the chart: along time, or evenly when every run finished at the same
// moment, and from zero up to the highest value. A run without the tile leaves a gap.
export function trendGeometry(points: readonly AuditTrendPoint[], tileId: string) {
  const values = points.map((point) => trendValue(point, tileId));
  const times = points.map((point) => Date.parse(point.finishedAt));
  const first = times[0] ?? 0;
  const span = (times.at(-1) ?? first) - first;
  const max = Math.max(1, ...values.filter((value): value is number => value !== null));
  const inner = CHART_WIDTH - 2 * CHART_PAD;
  const x = (index: number) =>
    round(
      CHART_PAD +
        (span > 0
          ? (((times[index] ?? first) - first) / span) * inner
          : points.length > 1
            ? (index / (points.length - 1)) * inner
            : inner / 2),
    );
  const y = (value: number) =>
    round(CHART_PAD + (1 - value / max) * (CHART_HEIGHT - 2 * CHART_PAD));
  const dots = points.flatMap((point, index) => {
    const value = values[index];
    return value === null || value === undefined
      ? []
      : [{ point, value, x: x(index), y: y(value) }];
  });
  const segments: string[] = [];
  let current: string[] = [];
  points.forEach((_point, index) => {
    const value = values[index];
    if (value === null || value === undefined) {
      if (current.length > 1) segments.push(current.join(" "));
      current = [];
      return;
    }
    current.push(`${x(index)},${y(value)}`);
  });
  if (current.length > 1) segments.push(current.join(" "));
  return { dots, segments, max, baseline: y(0), top: y(max) };
}

export function AuditTrendView({
  trend,
  tileId,
  showTable,
  onTile,
  onToggleTable,
}: {
  trend: AuditTrend;
  tileId: string;
  showTable: boolean;
  onTile: (tileId: string) => void;
  onToggleTable: () => void;
}) {
  const headingId = useId();
  const selectId = useId();
  const titleId = useId();
  const descId = useId();
  const tile = trend.tiles.find((item) => item.id === tileId) ?? trend.tiles[0];
  if (trend.points.length < 2 || !tile) return null;
  const geometry = trendGeometry(trend.points, tile.id);
  const first = geometry.dots[0];
  const last = geometry.dots.at(-1);
  const hollow = trendHollowNote(geometry.dots.map((dot) => dot.point));
  const summary =
    first && last
      ? `${tile.label}: ${number.format(last.value)} on ${day(last.point.finishedAt)}, against ${number.format(first.value)} on ${day(first.point.finishedAt)}.`
      : `No run in this range has ${tile.label.toLowerCase()}.`;
  return (
    <section className="panel audit-trend" aria-labelledby={headingId}>
      <div className="audit-trend__head">
        <h3 id={headingId}>Trend</h3>
        <div className="audit-filters">
          <label className="audit-filter" htmlFor={selectId}>
            <span className="audit-filter__label">Show</span>
            <select id={selectId} value={tile.id} onChange={(event) => onTile(event.target.value)}>
              {trend.tiles.map((item) => (
                <option key={item.id} value={item.id}>
                  {item.label}
                </option>
              ))}
            </select>
          </label>
          <button
            className="button button--secondary"
            type="button"
            aria-pressed={showTable}
            onClick={onToggleTable}
          >
            Show as table
          </button>
        </div>
      </div>
      <p className="audit-trend__summary">{summary}</p>
      {showTable ? (
        <div className="audit-scroll">
          <table className="audit-trend__table">
            <caption className="visually-hidden">{tile.label}, run by run</caption>
            <thead>
              <tr>
                <th scope="col">Finished</th>
                <th scope="col" className="audit-numeric">
                  {tile.label}
                </th>
                <th scope="col">Started by</th>
                <th scope="col">Note</th>
              </tr>
            </thead>
            <tbody>
              {trend.points.map((point) => {
                const value = trendValue(point, tile.id);
                return (
                  <tr key={point.runId}>
                    <th scope="row">
                      <time dateTime={point.finishedAt}>{when(point.finishedAt)}</time>
                    </th>
                    <td className="audit-numeric">
                      {value === null ? (
                        <span className="audit-empty">None</span>
                      ) : (
                        number.format(value)
                      )}
                    </td>
                    <td>{point.trigger === "schedule" ? "Schedule" : "An administrator"}</td>
                    <td>{trendNote(point)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      ) : (
        <>
          <svg
            className="audit-trend__chart"
            viewBox={`0 0 ${CHART_WIDTH} ${CHART_HEIGHT}`}
            role="img"
            aria-labelledby={`${titleId} ${descId}`}
          >
            <title
              id={titleId}
            >{`${tile.label} over the last ${number.format(trend.days)} days`}</title>
            <desc id={descId}>
              {`${plural(geometry.dots.length, "run", "runs")} with this value. ${summary}${hollow ? ` ${hollow}` : ""}`}
            </desc>
            <line
              className="audit-trend__grid"
              x1={CHART_PAD}
              x2={CHART_WIDTH - CHART_PAD}
              y1={geometry.top}
              y2={geometry.top}
            />
            <line
              className="audit-trend__grid"
              x1={CHART_PAD}
              x2={CHART_WIDTH - CHART_PAD}
              y1={geometry.baseline}
              y2={geometry.baseline}
            />
            {geometry.segments.map((points) => (
              <polyline key={points} className="audit-trend__line" points={points} />
            ))}
            {geometry.dots.map((dot) => {
              const note = trendNote(dot.point);
              return (
                <circle
                  key={dot.point.runId}
                  className={`audit-trend__point${note ? " audit-trend__point--hollow" : ""}`}
                  cx={dot.x}
                  cy={dot.y}
                  r={4}
                >
                  <title>{`${when(dot.point.finishedAt)}: ${number.format(dot.value)}${note ? `. ${note}` : ""}`}</title>
                </circle>
              );
            })}
          </svg>
          <div className="audit-trend__axis" aria-hidden="true">
            <span>{first ? day(first.point.finishedAt) : ""}</span>
            <span>{`0 to ${number.format(geometry.max)}`}</span>
            <span>{last ? day(last.point.finishedAt) : ""}</span>
          </div>
        </>
      )}
      {hollow ? <p className="muted-label">{hollow}</p> : null}
      {trend.truncated ? (
        <p className="muted-label">
          This range holds more runs than the chart shows; the oldest are left out.
        </p>
      ) : null}
    </section>
  );
}

export function AuditTrendPanel({ trend }: { trend: AuditTrend }) {
  const [tileId, setTileId] = useState<string | null>(null);
  const [showTable, setShowTable] = useState(false);
  const chosen =
    tileId && trend.tiles.some((tile) => tile.id === tileId)
      ? tileId
      : defaultTrendTile(trend.tiles);
  if (!chosen) return null;
  return (
    <AuditTrendView
      trend={trend}
      tileId={chosen}
      showTable={showTable}
      onTile={setTileId}
      onToggleTable={() => setShowTable((current) => !current)}
    />
  );
}

function AuditDetail({
  session,
  entry,
  scope,
  schedule,
  onSchedule,
  onBack,
  onOpenPerson,
  onChanged,
}: {
  session: PrismSession;
  entry: AuditCatalogEntry;
  // The saved audit scope: domains the account audits are limited to, empty for all of them.
  // null when the saved scope could not be loaded.
  scope: readonly string[] | null;
  // How often this audit runs on its own, and whether collection is on; null when the saved
  // schedule could not be loaded.
  schedule: {
    frequency: AuditScheduleFrequency;
    collecting: boolean;
    limit: AuditScheduleLimit | null;
  } | null;
  onSchedule: (frequency: AuditScheduleFrequency) => Promise<void>;
  onBack: () => void;
  onOpenPerson: (userId: string) => void;
  onChanged: () => void;
}) {
  const { definition } = entry;
  const historyId = useId();
  const [history, setHistory] = useState<AuditRun[] | null>(null);
  const [trend, setTrend] = useState<AuditTrend | null>(null);
  const [scheduleSaving, setScheduleSaving] = useState(false);
  const [scheduleError, setScheduleError] = useState("");
  const [runId, setRunId] = useState<string | null>(entry.runningId ?? entry.lastRun?.id ?? null);
  const [run, setRun] = useState<AuditRunDetail | null>(null);
  const [error, setError] = useState("");
  // A plain note about the last Run, such as having joined a scheduled run.
  const [notice, setNotice] = useState("");
  const [starting, setStarting] = useState(false);
  const demo = session.mode === "demo";
  // The parent's refresh, read through a ref so a new callback never restarts the polling.
  const changed = useRef(onChanged);
  changed.current = onChanged;

  const loadHistory = useCallback(async () => {
    if (demo) return;
    try {
      const list = await readJson(
        `/api/v1/audits/${encodeURIComponent(definition.id)}/runs`,
        auditRunListSchema.parse,
      );
      setHistory(list.runs);
      setRunId((current) => current ?? list.runs[0]?.id ?? null);
    } catch (problem) {
      setError(problem instanceof Error ? problem.message : "Run history could not be loaded.");
    }
  }, [definition.id, demo]);

  // The trend is a side panel: if it cannot be loaded, the page works without it.
  const loadTrend = useCallback(async () => {
    if (demo) return;
    try {
      setTrend(
        await readJson(
          `/api/v1/audits/${encodeURIComponent(definition.id)}/trend?days=365`,
          auditTrendSchema.parse,
        ),
      );
    } catch {
      // Keep what was loaded, if anything.
    }
  }, [definition.id, demo]);

  useEffect(() => {
    void loadHistory();
    void loadTrend();
  }, [loadHistory, loadTrend]);

  useEffect(() => {
    if (!runId) {
      setRun(null);
      return;
    }
    let stopped = false;
    let polled = false;
    let delay = POLL_MILLISECONDS;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const load = async () => {
      try {
        const detail = await readJson(
          `/api/v1/audit-runs/${encodeURIComponent(runId)}`,
          auditRunDetailSchema.parse,
        );
        if (stopped) return;
        setRun(detail);
        setError("");
        if (detail.status === "running") {
          polled = true;
          delay = POLL_MILLISECONDS;
          timer = setTimeout(() => void load(), delay);
        } else if (polled) {
          // The run this page watched has just finished.
          void loadHistory();
          void loadTrend();
          changed.current();
        }
      } catch (problem) {
        if (stopped) return;
        setError(problem instanceof Error ? problem.message : "This run could not be loaded.");
        // Keep watching a run in progress: one failed request (a rate limit, a restart) must not
        // leave the page stuck on "Running".
        if (polled) {
          delay = Math.min(MAX_POLL_MILLISECONDS, delay * 2);
          timer = setTimeout(() => void load(), delay);
        }
      }
    };
    void load();
    return () => {
      stopped = true;
      if (timer !== undefined) clearTimeout(timer);
    };
  }, [runId, loadHistory, loadTrend]);

  async function start() {
    setStarting(true);
    setError("");
    setNotice("");
    try {
      const started = await readJson(
        `/api/v1/audits/${encodeURIComponent(definition.id)}/runs`,
        auditStartSchema.parse,
        post(session),
      );
      setRunId(started.run.id);
      setRun({ ...started.run, rows: [] });
      setNotice(started.message ?? "");
      void loadHistory();
      changed.current();
    } catch (problem) {
      setError(problem instanceof Error ? problem.message : "The audit could not be started.");
    } finally {
      setStarting(false);
    }
  }

  async function changeSchedule(frequency: AuditScheduleFrequency) {
    setScheduleSaving(true);
    setScheduleError("");
    try {
      await onSchedule(frequency);
    } catch (problem) {
      setScheduleError(
        problem instanceof Error ? problem.message : "The schedule could not be saved.",
      );
    } finally {
      setScheduleSaving(false);
    }
  }

  const running = run?.status === "running" || starting;
  const access = accessText(entry);
  const took = run ? duration(run) : null;
  const scheduled = run?.trigger === "schedule";

  return (
    <div className="audit-detail">
      <nav className="breadcrumb" aria-label="Breadcrumb">
        <button type="button" onClick={onBack}>
          Audits
        </button>
        <span aria-hidden="true">/</span>
        <span aria-current="page">{definition.title}</span>
      </nav>

      <header className="audit-header">
        <div>
          <h2>{definition.title}</h2>
          <p>{definition.description}</p>
        </div>
        <div className="audit-header__actions">
          <AccessAction entry={entry} />
          <button
            className="button button--primary"
            type="button"
            disabled={demo || running || entry.availability.state !== "available"}
            onClick={() => void start()}
          >
            {running ? "Running…" : run ? "Run again" : "Run audit"}
          </button>
        </div>
      </header>

      {schedule && !demo ? (
        <AuditScheduleControl
          frequency={schedule.frequency}
          collecting={schedule.collecting}
          saving={scheduleSaving}
          error={scheduleError}
          lastStart={history?.[0]?.startedAt ?? entry.lastRun?.startedAt ?? null}
          limit={schedule.limit}
          onChange={(frequency) => void changeSchedule(frequency)}
        />
      ) : null}

      {access ? <p className="audit-note">{access}</p> : null}
      {notice ? (
        <p className="audit-note" role="status">
          {notice}
        </p>
      ) : null}
      {demo ? (
        <p className="audit-note">Connect a tenant to run this audit. Demo mode has no results.</p>
      ) : null}
      {error ? (
        <p className="form-error form-error--banner" role="alert">
          {error}
        </p>
      ) : null}

      {run ? (
        <div className="audit-run-line">
          <p>
            {run.status === "running" ? (
              <>
                <span className="collection-dot collection-dot--on" aria-hidden="true" />
                Reading Microsoft since {when(run.startedAt)}. Large directories take a few minutes.
              </>
            ) : run.status === "failed" ? (
              <>Run of {when(run.startedAt)} failed.</>
            ) : (
              <>
                Ran {when(run.startedAt)}
                {run.startedBy.name
                  ? scheduled
                    ? ` with ${run.startedBy.name}'s sign-in`
                    : ` by ${run.startedBy.name}`
                  : ""}
                {took ? `, took ${took}` : ""}.{" "}
                {run.complete ? null : (
                  <strong className="audit-partial">Partial result: see the notes below.</strong>
                )}
              </>
            )}
            {scheduled ? <span className="person-tag">Scheduled</span> : null}
            <RunScopeTags definition={definition} run={run} saved={scope} />
          </p>
          {history && history.length > 1 ? (
            <label className="audit-filter" htmlFor={historyId}>
              <span className="audit-filter__label">Run</span>
              <select
                id={historyId}
                value={runId ?? ""}
                onChange={(event) => setRunId(event.target.value || null)}
              >
                {history.map((item) => (
                  <option key={item.id} value={item.id}>
                    {runLabel(item)}
                  </option>
                ))}
              </select>
            </label>
          ) : null}
        </div>
      ) : null}

      {run?.status === "failed" && run.error ? (
        <p className="form-error form-error--banner" role="alert">
          {run.error.message}
        </p>
      ) : null}

      {!run && !demo && history !== null && history.length === 0 ? (
        <section className="panel audit-empty-run">
          <p>
            This audit has not been run yet. It reads Microsoft once each time you press Run audit,
            and keeps the last 10 successful results here.
          </p>
        </section>
      ) : null}

      {run?.status === "succeeded" ? (
        <ResultTable key={run.id} definition={definition} run={run} onOpenPerson={onOpenPerson} />
      ) : null}

      {trend ? <AuditTrendPanel trend={trend} /> : null}

      <section className="panel audit-about" aria-label="About this audit">
        <h3>About this audit</h3>
        <ul>
          {(run?.status === "succeeded" ? run.notes : []).map((note) => (
            <li key={note}>{note}</li>
          ))}
          {definition.requirements.map((requirement) => (
            <li key={requirement}>{requirement}</li>
          ))}
          <li>
            Reads with {definition.permissions.join(" and ")}
            {definition.optionalPermissions.length > 0
              ? `, and ${definition.optionalPermissions.join(" and ")} when granted`
              : ""}
            . Read-only: an audit never changes the tenant.
          </li>
        </ul>
      </section>
    </div>
  );
}

export function Audits({
  session,
  selected,
  onSelect,
  onOpenPerson,
}: {
  session: PrismSession;
  selected: string | null;
  onSelect: (auditId: string | null) => void;
  onOpenPerson: (userId: string) => void;
}) {
  const [catalog, setCatalog] = useState<AuditCatalog | null>(null);
  const [settings, setSettings] = useState<AuditSettings | null>(null);
  const [schedule, setSchedule] = useState<AuditScheduleSettings | null>(null);
  const [error, setError] = useState("");
  // A plain note about the last Run, such as having joined a scheduled run.
  const [notice, setNotice] = useState("");
  const [busy, setBusy] = useState<string | null>(null);
  const live = session.mode === "live";

  const refresh = useCallback(async () => {
    try {
      setCatalog(await readJson("/api/v1/audits", auditCatalogSchema.parse));
      setError("");
    } catch (problem) {
      setError(problem instanceof Error ? problem.message : "Audits could not be loaded.");
    }
  }, []);

  // The saved scope and the domains to choose from. Where the server cannot answer (the
  // recorded demo has no scope), the section stays hidden and the audits work as before.
  const loadSettings = useCallback(async () => {
    if (!live) return;
    try {
      setSettings(await readJson("/api/v1/settings/audits", auditSettingsSchema.parse));
    } catch {
      // Keep what was loaded, if anything.
    }
  }, [live]);

  // Which audits run on their own. Where the server cannot answer, the control and the tags stay
  // hidden.
  const loadSchedule = useCallback(async () => {
    if (!live) return;
    try {
      setSchedule(
        await readJson("/api/v1/settings/audit-schedule", auditScheduleSettingsSchema.parse),
      );
    } catch {
      // Keep what was loaded, if anything.
    }
  }, [live]);

  useEffect(() => {
    if (live && !session.authenticated) return;
    void refresh();
    void loadSettings();
    void loadSchedule();
  }, [live, session.authenticated, refresh, loadSettings, loadSchedule]);

  // While any audit runs, keep the catalog's status lines current.
  const anyRunning = Boolean(catalog?.audits.some((entry) => entry.runningId));
  // A finished run changes which domains can be listed. Reloaded once, not on every poll: the
  // domains are counted from stored results.
  const wasRunning = useRef(false);
  useEffect(() => {
    if (wasRunning.current && !anyRunning) void loadSettings();
    wasRunning.current = anyRunning;
  }, [anyRunning, loadSettings]);
  useEffect(() => {
    if (!anyRunning || selected) return;
    const timer = setInterval(() => void refresh(), POLL_MILLISECONDS);
    return () => clearInterval(timer);
  }, [anyRunning, selected, refresh]);

  async function run(auditId: string) {
    setBusy(auditId);
    setError("");
    setNotice("");
    try {
      const started = await readJson(
        `/api/v1/audits/${encodeURIComponent(auditId)}/runs`,
        auditStartSchema.parse,
        post(session),
      );
      setNotice(started.message ?? "");
      await refresh();
    } catch (problem) {
      setError(problem instanceof Error ? problem.message : "The audit could not be started.");
    } finally {
      setBusy(null);
    }
  }

  async function saveSchedule(auditId: string, frequency: AuditScheduleFrequency) {
    setSchedule(await saveAuditSchedule(session, auditId, frequency));
  }

  async function saveScope(domains: string[]) {
    setSettings(
      await readJson(
        "/api/v1/settings/audits",
        auditSettingsSchema.parse,
        post(session, { domains }),
      ),
    );
  }

  if (live && !session.authenticated)
    return <p className="empty-data">Sign in to run audits against your tenant.</p>;
  const entry = selected ? catalog?.audits.find((item) => item.definition.id === selected) : null;
  if (selected && entry)
    return (
      <AuditDetail
        key={entry.definition.id}
        session={session}
        entry={entry}
        scope={settings ? settings.scope.domains : null}
        schedule={
          schedule
            ? {
                frequency: schedule.audits[entry.definition.id] ?? "off",
                collecting: schedule.collecting,
                limit: schedule.limits[entry.definition.id] ?? null,
              }
            : null
        }
        onSchedule={(frequency) => saveSchedule(entry.definition.id, frequency)}
        onBack={() => {
          onSelect(null);
          void refresh();
        }}
        onOpenPerson={onOpenPerson}
        onChanged={() => void refresh()}
      />
    );
  return (
    <>
      {error ? (
        <p className="form-error form-error--banner" role="alert">
          {error}
        </p>
      ) : null}
      {notice ? (
        <p className="audit-note" role="status">
          {notice}
        </p>
      ) : null}
      {catalog ? (
        <AuditCatalogList
          catalog={catalog}
          busy={busy}
          onOpen={onSelect}
          onRun={(id) => void run(id)}
          schedule={schedule?.audits ?? null}
          scope={
            settings ? (
              <AuditScopeSection
                settings={settings}
                applies={catalog.audits
                  .filter((item) => item.definition.scoped)
                  .map((item) => item.definition.title)}
                onSave={saveScope}
              />
            ) : null
          }
        />
      ) : error ? null : (
        <p className="empty-data">Loading audits…</p>
      )}
    </>
  );
}

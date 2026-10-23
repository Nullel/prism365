import {
  type AuditCatalog,
  type AuditCatalogEntry,
  type AuditCell,
  type AuditColumn,
  type AuditDefinition,
  type AuditRow,
  type AuditRun,
  type AuditRunDetail,
  type AuditSettings,
  type AuditSummaryItem,
  auditCatalogSchema,
  auditFacetOptions,
  auditRunDetailSchema,
  auditRunListSchema,
  auditSettingsSchema,
  auditStartSchema,
  MAX_AUDIT_SCOPE_DOMAINS,
  type PrismSession,
  queryAuditRows,
} from "@prism365/contracts";
import { type ReactNode, useCallback, useEffect, useId, useMemo, useRef, useState } from "react";
import { AccountName, day, number, post, readJson, when } from "./common.tsx";
import { CopyButton, useCopy } from "./CopyButton.tsx";
import { downloadHandoff } from "./Handoffs.tsx";
import { ArrowUpRightIcon, CheckIcon, CopyIcon, DownloadIcon } from "./icons.tsx";

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
}: {
  catalog: AuditCatalog;
  busy: string | null;
  onOpen: (auditId: string) => void;
  onRun: (auditId: string) => void;
  // The audit scope section, shown above the audits when the page has loaded it.
  scope?: ReactNode;
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

interface ExportQuery {
  view: string;
  search: string;
  sort: { column: string; direction: "asc" | "desc" };
  facets: Record<string, string>;
}

// The page's view, search, sort and filters as export parameters. The CSV and the handoff take
// the same ones, and the server reads both the same way, so they always hold the same rows.
function exportParams(query: ExportQuery): URLSearchParams {
  const params = new URLSearchParams({ view: query.view });
  if (query.search.trim()) params.set("search", query.search.trim());
  params.set("sort", query.sort.column);
  params.set("direction", query.sort.direction);
  for (const [key, value] of Object.entries(query.facets)) if (value) params.set(`f.${key}`, value);
  return params;
}

export function exportHref(run: AuditRun, query: ExportQuery): string {
  return `/api/v1/audit-runs/${encodeURIComponent(run.id)}/export.csv?${exportParams(query).toString()}`;
}

// The handoff for the same rows, with the columns the page shows as each account's evidence.
export function handoffHref(run: AuditRun, query: ExportQuery, columns: readonly string[]): string {
  const params = exportParams(query);
  params.set("columns", columns.join(","));
  return `/api/v1/audit-runs/${encodeURIComponent(run.id)}/handoff.json?${params.toString()}`;
}

export const handoffNote =
  "A handoff is a list for the tool that acts on it. Prism365 itself changes nothing.";
const HANDOFF_NOTED = "prism365.handoff.noted";

// True the first time this browser exports a handoff. Without storage, every time.
function firstHandoff(): boolean {
  try {
    if (localStorage.getItem(HANDOFF_NOTED)) return false;
    localStorage.setItem(HANDOFF_NOTED, "1");
  } catch {
    // Private mode or blocked storage: show the note again rather than never.
  }
  return true;
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
  const [handoff, setHandoff] = useState({ busy: false, error: "", noted: false });
  // Only an audit of accounts an action tool works on offers a handoff, and only for a run whose
  // rows name accounts by their directory id.
  const handoffable =
    definition.handoff && hasPeople && run.rows.some((row) => Boolean(row.personId));
  const queryKey = JSON.stringify([view, facets, search, sort]);
  const limit = page.key === queryKey ? page.limit : PAGE_ROWS;
  const columns = definition.columns.filter((column) => shown.has(column.key));
  const facetColumns = definition.columns.filter(
    (column) => column.facet && (options[column.key]?.length ?? 0) > 1,
  );
  const activeView = definition.views.find((item) => item.id === view);
  async function exportForAction() {
    setHandoff((current) => ({ ...current, busy: true, error: "" }));
    try {
      await downloadHandoff(
        handoffHref(
          run,
          { view, search, sort, facets },
          columns.map((column) => column.key),
        ),
      );
      const first = firstHandoff();
      setHandoff((current) => ({ busy: false, error: "", noted: current.noted || first }));
    } catch (problem) {
      setHandoff((current) => ({
        ...current,
        busy: false,
        error: problem instanceof Error ? problem.message : "The handoff could not be exported.",
      }));
    }
  }
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
            {handoffable ? (
              <button
                className="button button--secondary"
                type="button"
                disabled={handoff.busy}
                onClick={() => void exportForAction()}
              >
                <ArrowUpRightIcon /> {handoff.busy ? "Exporting…" : "Export for action"}
              </button>
            ) : null}
          </div>
        </div>
        {handoff.noted ? <p className="audit-note handoff-note">{handoffNote}</p> : null}
        {handoff.error ? (
          <p className="form-error" role="alert">
            {handoff.error}
          </p>
        ) : null}
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
  const status = run.status === "running" ? "running" : run.status === "failed" ? "failed" : null;
  return `${when(run.startedAt)}${status ? ` (${status})` : ""}`;
}

function AuditDetail({
  session,
  entry,
  scope,
  onBack,
  onOpenPerson,
  onChanged,
}: {
  session: PrismSession;
  entry: AuditCatalogEntry;
  // The saved audit scope: domains the account audits are limited to, empty for all of them.
  // null when the saved scope could not be loaded.
  scope: readonly string[] | null;
  onBack: () => void;
  onOpenPerson: (userId: string) => void;
  onChanged: () => void;
}) {
  const { definition } = entry;
  const historyId = useId();
  const [history, setHistory] = useState<AuditRun[] | null>(null);
  const [runId, setRunId] = useState<string | null>(entry.runningId ?? entry.lastRun?.id ?? null);
  const [run, setRun] = useState<AuditRunDetail | null>(null);
  const [error, setError] = useState("");
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

  useEffect(() => {
    void loadHistory();
  }, [loadHistory]);

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
  }, [runId, loadHistory]);

  async function start() {
    setStarting(true);
    setError("");
    try {
      const started = await readJson(
        `/api/v1/audits/${encodeURIComponent(definition.id)}/runs`,
        auditStartSchema.parse,
        post(session),
      );
      setRunId(started.run.id);
      setRun({ ...started.run, rows: [] });
      void loadHistory();
      changed.current();
    } catch (problem) {
      setError(problem instanceof Error ? problem.message : "The audit could not be started.");
    } finally {
      setStarting(false);
    }
  }

  const running = run?.status === "running" || starting;
  const access = accessText(entry);
  const took = run ? duration(run) : null;

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

      {access ? <p className="audit-note">{access}</p> : null}
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
                {run.startedBy.name ? ` by ${run.startedBy.name}` : ""}
                {took ? `, took ${took}` : ""}.{" "}
                {run.complete ? null : (
                  <strong className="audit-partial">Partial result: see the notes below.</strong>
                )}
              </>
            )}
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
  const [error, setError] = useState("");
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

  useEffect(() => {
    if (live && !session.authenticated) return;
    void refresh();
    void loadSettings();
  }, [live, session.authenticated, refresh, loadSettings]);

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
    try {
      await readJson(
        `/api/v1/audits/${encodeURIComponent(auditId)}/runs`,
        auditStartSchema.parse,
        post(session),
      );
      await refresh();
    } catch (problem) {
      setError(problem instanceof Error ? problem.message : "The audit could not be started.");
    } finally {
      setBusy(null);
    }
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
      {catalog ? (
        <AuditCatalogList
          catalog={catalog}
          busy={busy}
          onOpen={onSelect}
          onRun={(id) => void run(id)}
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

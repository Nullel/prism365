import {
  type AuditCatalog,
  auditCatalogSchema,
  investigationEmailSchema,
  type PersonList,
  type PrismSession,
  personListSchema,
} from "@prism365/contracts";
import {
  type KeyboardEvent,
  type MouseEvent,
  type Ref,
  useEffect,
  useId,
  useMemo,
  useRef,
  useState,
} from "react";
import { runResult } from "./Audits.tsx";
import { readJson } from "./common.tsx";
import { ChecklistIcon, SearchIcon, UsersIcon } from "./icons.tsx";
import { lookUpPerson } from "./People.tsx";

// Quick search: Cmd+K (Ctrl+K elsewhere) opens one box that goes straight to a page, an audit or
// a person. Everything it searches is already in the browser: the sidebar's pages, the audit
// catalog and the last 30 days of people, each fetched once the first time the box opens. There is
// no search request, so nothing here reaches Microsoft or the server beyond those two reads and
// the same directory lookup the People search makes.

// ---------------------------------------------------------------------------------------------
// What can be found
// ---------------------------------------------------------------------------------------------

type IconComponent = typeof SearchIcon;

export type ResultGroupId = "directory" | "pages" | "audits" | "people";

export interface SearchItem {
  // Unique within its group: a page's view, an audit's id, an account's id or an email address.
  id: string;
  label: string;
  // Muted text after the label: an audit's latest result, a person's principal name.
  detail: string | null;
  // Other words that find the item, weighed below the label.
  keywords: readonly string[];
  icon?: IconComponent;
}

export interface ResultGroup {
  id: ResultGroupId;
  heading: string;
  items: SearchItem[];
}

export interface SearchSources {
  pages: readonly SearchItem[];
  audits: readonly SearchItem[];
  people: readonly SearchItem[];
}

export const MAX_RESULTS = 8;

const headings: Record<ResultGroupId, string> = {
  directory: "Directory",
  pages: "Pages",
  audits: "Audits",
  people: "People",
};

// Words people reach for that a page's label does not hold. Keyed by the page's view; a page
// without an entry is found by its label and its sidebar section alone.
export const pageSynonyms: Readonly<Record<string, readonly string[]>> = {
  overview: ["home", "dashboard", "summary", "start"],
  week: ["weekly", "digest", "report", "recent", "what changed"],
  people: ["users", "user", "person", "accounts", "staff", "employees", "lookup"],
  notifications: [
    "notifications",
    "service health",
    "incidents",
    "message center",
    "deadlines",
    "microsoft",
    "settings",
  ],
  audits: ["checks", "run an audit", "reports"],
  signins: ["sign in", "signin", "login", "logins", "authentication", "failed sign-ins"],
  changes: ["audit log", "who changed", "history", "changelog"],
  handoffs: ["export", "hand off", "receipts", "lifecycle", "runbook"],
  alerts: ["notify", "email", "webhook", "teams", "settings"],
  access: ["permissions", "consent", "scopes", "connection", "tenant", "graph", "settings"],
};

interface NavGroupLike {
  label: string;
  items: readonly { label: string; view: string; icon: IconComponent }[];
}

// Every sidebar entry becomes a page result, in sidebar order. The section it sits under
// ("Monitor", "Settings") finds it too.
export function pageItems(groups: readonly NavGroupLike[]): SearchItem[] {
  return groups.flatMap((group) =>
    group.items.map((item) => ({
      id: item.view,
      label: item.label,
      detail: null,
      keywords: [...(group.label ? [group.label] : []), ...(pageSynonyms[item.view] ?? [])],
      icon: item.icon,
    })),
  );
}

export function auditItems(catalog: AuditCatalog): SearchItem[] {
  return catalog.audits.map((entry) => ({
    id: entry.definition.id,
    label: entry.definition.title,
    detail: runResult(entry.lastRun, entry.definition),
    keywords: [entry.definition.category],
    icon: ChecklistIcon,
  }));
}

export function personItems(list: PersonList): SearchItem[] {
  return list.items.map((item) => {
    const named = item.displayName !== null && item.displayName !== "";
    return {
      id: item.userId,
      label: (named ? item.displayName : item.userPrincipalName) ?? item.userId,
      detail: named ? item.userPrincipalName : null,
      keywords: item.userPrincipalName ? [item.userPrincipalName] : [],
      icon: UsersIcon,
    };
  });
}

// A full work email is the one thing quick search can look up that it has not stored.
export function emailToLookUp(query: string): string | null {
  const parsed = investigationEmailSchema.safeParse(query);
  return parsed.success ? parsed.data : null;
}

function lookupItem(email: string): SearchItem {
  return {
    id: email,
    label: `Look up ${email} in the directory`,
    detail: null,
    keywords: [],
    icon: UsersIcon,
  };
}

// ---------------------------------------------------------------------------------------------
// Matching
// ---------------------------------------------------------------------------------------------

// Lower case with accents taken off, so "Zoë" is found by "zoe" and the other way round.
export function fold(text: string): string {
  return text.normalize("NFD").replace(/\p{M}/gu, "").toLowerCase();
}

const isBoundary = (char: string) => !/[\p{L}\p{N}]/u.test(char);

// 4 when the text starts with the token, 3 when a word in it does, 1 when it is only inside a
// word, 0 when it is not there.
function position(text: string, token: string): number {
  let at = text.indexOf(token);
  if (at < 0) return 0;
  if (at === 0) return 4;
  while (at >= 0) {
    if (isBoundary(text.charAt(at - 1))) return 3;
    at = text.indexOf(token, at + 1);
  }
  return 1;
}

const LABEL_WEIGHT = 3;
const KEYWORD_WEIGHT = 2;
const EXACT_LABEL_BONUS = 10;

// 0 when the item does not match. Every word typed has to be found, in the label or a keyword;
// a word counts for more at the start of the label or of a word, and more in the label than in a
// keyword, so a page named "Audits" beats a page that only mentions audits.
export function scoreItem(item: SearchItem, query: string): number {
  const tokens = fold(query).split(/\s+/).filter(Boolean);
  if (tokens.length === 0) return 0;
  const label = fold(item.label);
  const keywords = item.keywords.map(fold);
  let total = label === tokens.join(" ") ? EXACT_LABEL_BONUS : 0;
  for (const token of tokens) {
    let best = position(label, token) * LABEL_WEIGHT;
    for (const keyword of keywords)
      best = Math.max(best, position(keyword, token) * KEYWORD_WEIGHT);
    if (best === 0) return 0;
    total += best;
  }
  return total;
}

interface Ranked {
  group: ResultGroupId;
  item: SearchItem;
  score: number;
  order: number;
}

function rank(group: ResultGroupId, items: readonly SearchItem[], query: string): Ranked[] {
  const found: Ranked[] = [];
  items.forEach((item, order) => {
    const score = scoreItem(item, query);
    if (score > 0) found.push({ group, item, score, order });
  });
  return found;
}

// Where equal scores go: the lookup the person typed an email for, then pages, audits, people.
const groupOrder: Record<ResultGroupId, number> = { directory: 0, pages: 1, audits: 2, people: 3 };
const LOOKUP_SCORE = 1_000;

export interface SearchOptions {
  // A full work email to offer a lookup for when it matches nobody stored. null while the people
  // are still loading or when looking someone up is not possible.
  email?: string | null;
  limit?: number;
}

// The results for what was typed: the best few across all sources, put under small headings. The
// group holding the best match comes first, and each group keeps its own best-first order. With
// nothing typed, the pages.
export function searchResults(
  sources: SearchSources,
  query: string,
  options: SearchOptions = {},
): ResultGroup[] {
  const limit = options.limit ?? MAX_RESULTS;
  if (fold(query).trim() === "")
    return [{ id: "pages", heading: headings.pages, items: sources.pages.slice(0, limit) }];
  const people = rank("people", sources.people, query);
  const found = [
    ...rank("pages", sources.pages, query),
    ...rank("audits", sources.audits, query),
    ...people,
  ];
  if (options.email && people.length === 0)
    found.push({
      group: "directory",
      item: lookupItem(options.email),
      score: LOOKUP_SCORE,
      order: 0,
    });
  found.sort(
    (a, b) => b.score - a.score || groupOrder[a.group] - groupOrder[b.group] || a.order - b.order,
  );
  const groups: ResultGroup[] = [];
  for (const entry of found.slice(0, limit)) {
    let group = groups.find((candidate) => candidate.id === entry.group);
    if (!group) {
      group = { id: entry.group, heading: headings[entry.group], items: [] };
      groups.push(group);
    }
    group.items.push(entry.item);
  }
  return groups;
}

export interface FlatResult {
  group: ResultGroupId;
  item: SearchItem;
}

// The results in the order the arrow keys walk them.
export function flatten(groups: readonly ResultGroup[]): FlatResult[] {
  return groups.flatMap((group) => group.items.map((item) => ({ group: group.id, item })));
}

// ---------------------------------------------------------------------------------------------
// Loading what is searched
// ---------------------------------------------------------------------------------------------

export type Load<T> =
  | { state: "idle" }
  | { state: "loading" }
  | { state: "ready"; value: T }
  | { state: "failed" };

// Both are read exactly as the demo recording holds them (docs/DEMO.md), so the public demo
// answers them. Change one and the recording has to change with it.
export const auditsUrl = "/api/v1/audits";
export const peopleUrl = "/api/v1/people?days=30";

const loadAudits = () => readJson(auditsUrl, (value) => auditCatalogSchema.parse(value));
const loadPeople = () => readJson(peopleUrl, (value) => personListSchema.parse(value));

// Loads once when first wanted and keeps the result. A failure is not kept: the next time it is
// wanted, it tries again.
function useLoadOnce<T>(wanted: boolean, load: () => Promise<T>): Load<T> {
  const [loaded, setLoaded] = useState<Load<T>>({ state: "idle" });
  const started = useRef(false);
  const mounted = useRef(false);
  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
    };
  }, []);
  useEffect(() => {
    if (!wanted || started.current) return;
    started.current = true;
    setLoaded({ state: "loading" });
    load().then(
      (value) => {
        if (mounted.current) setLoaded({ state: "ready", value });
      },
      () => {
        started.current = false;
        if (mounted.current) setLoaded({ state: "failed" });
      },
    );
  }, [wanted, load]);
  return loaded;
}

// Plain lines under the results about a source that is not there: not signed in, still loading,
// or failed. The other sources keep working either way.
export function sourceNotes(
  query: string,
  canRead: boolean,
  audits: Load<unknown>,
  people: Load<unknown>,
): string[] {
  if (!canRead) return ["Sign in to search audits and people."];
  if (fold(query).trim() === "") return ["Type to search audits and people as well."];
  const notes: string[] = [];
  if (audits.state === "loading") notes.push("Loading audits…");
  if (audits.state === "failed") notes.push("Audits could not be loaded.");
  if (people.state === "loading") notes.push("Loading people…");
  if (people.state === "failed") notes.push("People could not be loaded.");
  return notes;
}

// ---------------------------------------------------------------------------------------------
// Keyboard
// ---------------------------------------------------------------------------------------------

// The part of a keyboard event the handlers read, so both the browser's events and React's fit
// and a test can pass a plain object.
export interface KeyEventLike {
  key: string;
  metaKey: boolean;
  ctrlKey: boolean;
  altKey: boolean;
  shiftKey: boolean;
  defaultPrevented: boolean;
  target: unknown;
  isComposing?: boolean;
  nativeEvent?: { isComposing?: boolean };
  preventDefault(): void;
}

const notTextInputs = new Set([
  "button",
  "checkbox",
  "color",
  "file",
  "image",
  "radio",
  "range",
  "reset",
  "submit",
]);

// Somewhere typing a "/" means a "/": a text box, a select, or anything editable.
export function isTextField(target: unknown): boolean {
  if (!target || typeof target !== "object") return false;
  const { tagName, type, isContentEditable } = target as {
    tagName?: unknown;
    type?: unknown;
    isContentEditable?: unknown;
  };
  if (isContentEditable === true) return true;
  const tag = typeof tagName === "string" ? tagName.toUpperCase() : "";
  if (tag === "TEXTAREA" || tag === "SELECT") return true;
  if (tag === "INPUT") return !(typeof type === "string" && notTextInputs.has(type.toLowerCase()));
  return false;
}

const composing = (event: KeyEventLike) =>
  (event.nativeEvent?.isComposing ?? event.isComposing) === true;

// Cmd+K and Ctrl+K open it from anywhere, even inside another text box, as other tools do. "/"
// opens it only when it would not be typed into something.
export function opensQuickSearch(event: KeyEventLike): boolean {
  if (event.defaultPrevented || composing(event)) return false;
  const key = event.key.toLowerCase();
  if (key === "k") return (event.metaKey || event.ctrlKey) && !event.altKey && !event.shiftKey;
  if (key === "/")
    return !event.metaKey && !event.ctrlKey && !event.altKey && !isTextField(event.target);
  return false;
}

// The page-wide listener. The shortcut is claimed even while the box is open, so the browser's
// own Ctrl+K (the address bar's search) never fires on top of it.
export function openFromKeyboard(
  event: KeyEventLike,
  open: boolean,
  setOpen: (open: boolean) => void,
): void {
  if (!opensQuickSearch(event)) return;
  event.preventDefault();
  if (!open) setOpen(true);
}

export interface DialogKeyActions {
  move(to: number): void;
  choose(index: number): void;
  close(): void;
}

// Keys inside the box. Up and Down walk the results and wrap, Enter opens the one marked, Escape
// closes, and Tab stays put: the input is the only control, so focus has nowhere else to go.
export function handleDialogKey(
  event: KeyEventLike,
  state: { count: number; active: number },
  actions: DialogKeyActions,
): void {
  const { count, active } = state;
  if (event.key === "Escape") {
    event.preventDefault();
    actions.close();
    return;
  }
  if (event.key === "Tab") {
    event.preventDefault();
    return;
  }
  if (event.metaKey || event.ctrlKey || event.altKey) return;
  if (event.key === "ArrowDown" || event.key === "ArrowUp") {
    event.preventDefault();
    if (count > 0) actions.move((active + (event.key === "ArrowDown" ? 1 : count - 1)) % count);
    return;
  }
  // Enter that confirms a word in an input method is not a request to open anything.
  if (event.key === "Enter" && !composing(event)) {
    event.preventDefault();
    if (count > 0) actions.choose(active);
  }
}

interface Focusable {
  focus(): void;
  isConnected?: boolean;
}

const canFocus = (value: unknown): value is Focusable =>
  typeof value === "object" && value !== null && typeof (value as Focusable).focus === "function";

// Remembers what had focus when the box opened and gives it back when it closes, unless that
// element has left the page in the meantime.
export function focusKeeper(active: () => unknown) {
  let saved: Focusable | null = null;
  return {
    remember() {
      const current = active();
      saved = canFocus(current) ? current : null;
    },
    restore() {
      const target = saved;
      saved = null;
      if (target && target.isConnected !== false) target.focus();
    },
  };
}

export interface PlatformInfo {
  platform?: string;
  userAgentData?: { platform?: string };
}

export function isMacPlatform(info: PlatformInfo | undefined): boolean {
  const name = info?.userAgentData?.platform || info?.platform || "";
  return /mac|iphone|ipad|ipod/i.test(name);
}

export function shortcutHint(mac: boolean): string {
  return mac ? "⌘K" : "Ctrl K";
}

// ---------------------------------------------------------------------------------------------
// The button
// ---------------------------------------------------------------------------------------------

// At the top of the sidebar, and in the top bar on a phone.
export function QuickSearchButton({ mac, onOpen }: { mac: boolean; onOpen: () => void }) {
  return (
    <button
      className="quick-search-open"
      type="button"
      onClick={onOpen}
      aria-haspopup="dialog"
      aria-keyshortcuts={mac ? "Meta+K" : "Control+K"}
    >
      <SearchIcon />
      <span className="quick-search-open__label">Search</span>
      <span className="quick-search-open__keys" aria-hidden="true">
        <kbd>{shortcutHint(mac)}</kbd>
      </span>
    </button>
  );
}

// ---------------------------------------------------------------------------------------------
// The box
// ---------------------------------------------------------------------------------------------

export interface QuickSearchViewProps {
  idBase: string;
  query: string;
  groups: readonly ResultGroup[];
  active: number;
  notes: readonly string[];
  // The one line shown when nothing matches; null otherwise.
  empty: string | null;
  // What a screen reader hears when the results change.
  status: string;
  working: boolean;
  error: string;
  inputRef?: Ref<HTMLInputElement>;
  onQuery(value: string): void;
  onKeyDown(event: KeyboardEvent<HTMLElement>): void;
  onPick(index: number): void;
  onHover(index: number): void;
  onBackdrop(): void;
}

export const optionId = (idBase: string, index: number) => `${idBase}-option-${index}`;

const placeholder = "Search pages, audits and people";

// The dialog as the ARIA combobox and listbox pattern: focus stays in the input, and the option
// it points at with aria-activedescendant is the one Enter opens.
export function QuickSearchView({
  idBase,
  query,
  groups,
  active,
  notes,
  empty,
  status,
  working,
  error,
  inputRef,
  onQuery,
  onKeyDown,
  onPick,
  onHover,
  onBackdrop,
}: QuickSearchViewProps) {
  const titleId = `${idBase}-title`;
  const listId = `${idBase}-list`;
  const count = groups.reduce((sum, group) => sum + group.items.length, 0);
  let index = 0;
  return (
    // biome-ignore lint/a11y/noStaticElementInteractions: closing on the backdrop is a pointer convenience; Escape does the same from the keyboard.
    <div
      className="quick-search"
      role="presentation"
      onMouseDown={(event: MouseEvent<HTMLElement>) => {
        if (event.target === event.currentTarget) onBackdrop();
      }}
    >
      <div
        className="quick-search__panel"
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        onKeyDown={onKeyDown}
        onMouseDown={(event: MouseEvent<HTMLElement>) => {
          // A click on the panel's padding or a result must not move focus out of the input.
          if (!isTextField(event.target)) event.preventDefault();
        }}
      >
        <h2 className="visually-hidden" id={titleId}>
          Search pages, audits and people
        </h2>
        <div className="quick-search__field">
          <SearchIcon />
          <input
            ref={inputRef}
            className="quick-search__input"
            type="text"
            role="combobox"
            aria-label={placeholder}
            aria-autocomplete="list"
            aria-expanded={count > 0}
            aria-controls={count > 0 ? listId : undefined}
            aria-activedescendant={count > 0 ? optionId(idBase, active) : undefined}
            placeholder={placeholder}
            autoComplete="off"
            autoCapitalize="off"
            autoCorrect="off"
            spellCheck={false}
            value={query}
            onChange={(event) => onQuery(event.target.value)}
          />
        </div>
        {count > 0 ? (
          <div className="quick-search__results" id={listId} role="listbox" aria-label="Results">
            {groups.map((group) => {
              const headingId = `${idBase}-${group.id}`;
              return (
                <div role="presentation" key={group.id}>
                  <div className="quick-search__heading" id={headingId} aria-hidden="true">
                    {group.heading}
                  </div>
                  {/* biome-ignore lint/a11y/useSemanticElements: options inside a listbox group need role="group"; a fieldset is not allowed there. */}
                  <div role="group" aria-labelledby={headingId}>
                    {group.items.map((item) => {
                      const at = index++;
                      const Icon = item.icon;
                      const marked = at === active;
                      return (
                        // biome-ignore lint/a11y/useKeyWithClickEvents: the keyboard picks with Enter in the input, which owns focus.
                        // biome-ignore lint/a11y/useFocusableInteractive: options are never focused; the input points at one with aria-activedescendant.
                        <div
                          role="option"
                          id={optionId(idBase, at)}
                          aria-selected={marked}
                          className={`quick-search__option${marked ? " quick-search__option--active" : ""}`}
                          key={item.id}
                          onClick={() => onPick(at)}
                          onMouseMove={() => onHover(at)}
                        >
                          {Icon ? <Icon /> : null}
                          <span className="quick-search__label">{item.label}</span>
                          {item.detail ? (
                            <span className="quick-search__detail">{item.detail}</span>
                          ) : null}
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        ) : null}
        {error ? (
          <p className="form-error quick-search__note" role="alert">
            {error}
          </p>
        ) : null}
        {working ? <p className="quick-search__note">Looking up…</p> : null}
        {empty ? <p className="quick-search__note">{empty}</p> : null}
        {notes.map((note) => (
          <p className="quick-search__note" key={note}>
            {note}
          </p>
        ))}
        <div className="visually-hidden" role="status" aria-live="polite">
          {status}
        </div>
      </div>
    </div>
  );
}

export function resultStatus(query: string, count: number): string {
  if (fold(query).trim() === "") return "";
  if (count === 0) return "No results.";
  return count === 1 ? "1 result." : `${count} results.`;
}

export function noMatchText(query: string): string {
  return `No page, audit or person matches “${query.trim()}”.`;
}

export interface QuickSearchProps {
  open: boolean;
  onOpenChange(open: boolean): void;
  session: PrismSession | null;
  pages: readonly SearchItem[];
  onNavigate(view: string): void;
  onOpenAudit(auditId: string): void;
  onOpenPerson(userId: string): void;
}

// Owns the box: what is typed, which result is marked, the listener for the shortcut, and the
// two reads. Mounted for as long as the app is, so what it loaded is kept between openings.
export function QuickSearch({
  open,
  onOpenChange,
  session,
  pages,
  onNavigate,
  onOpenAudit,
  onOpenPerson,
}: QuickSearchProps) {
  const idBase = useId();
  const input = useRef<HTMLInputElement>(null);
  const keeper = useMemo(
    () => focusKeeper(() => (typeof document === "undefined" ? null : document.activeElement)),
    [],
  );
  // Bumped when the box closes, so a lookup that comes back afterwards opens nothing.
  const generation = useRef(0);
  const [query, setQuery] = useState("");
  const [active, setActive] = useState(0);
  const [lookup, setLookup] = useState({ working: false, error: "" });

  // The local demo mode has no tenant: it can list the audits (none has run) but has no people,
  // and a directory lookup needs a Microsoft session.
  const canRead = Boolean(session && (session.mode === "demo" || session.authenticated));
  const canLookUp = session?.mode === "live" && session.authenticated;
  const audits = useLoadOnce(open && canRead, loadAudits);
  const people = useLoadOnce(open && canRead, loadPeople);

  useEffect(() => {
    const listener = (event: globalThis.KeyboardEvent) =>
      openFromKeyboard(event, open, onOpenChange);
    window.addEventListener("keydown", listener);
    return () => window.removeEventListener("keydown", listener);
  }, [open, onOpenChange]);

  useEffect(() => {
    if (!open) return;
    if (document.activeElement !== input.current) keeper.remember();
    input.current?.focus();
  }, [open, keeper]);

  const sources = useMemo<SearchSources>(
    () => ({
      pages,
      audits: audits.state === "ready" ? auditItems(audits.value) : [],
      people: people.state === "ready" ? personItems(people.value) : [],
    }),
    [pages, audits, people],
  );
  // Only offered once the stored people have been searched, so it never appears and then goes.
  const email =
    canLookUp && (people.state === "ready" || people.state === "failed")
      ? emailToLookUp(query)
      : null;
  const groups = useMemo(() => searchResults(sources, query, { email }), [sources, query, email]);
  const results = flatten(groups);
  const marked = Math.min(active, Math.max(0, results.length - 1));
  const notes = sourceNotes(query, canRead, audits, people);
  const stillLoading = audits.state === "loading" || people.state === "loading";

  useEffect(() => {
    if (!open) return;
    document.getElementById(optionId(idBase, marked))?.scrollIntoView({ block: "nearest" });
  }, [open, idBase, marked]);

  function close() {
    generation.current += 1;
    keeper.restore();
    setQuery("");
    setActive(0);
    setLookup({ working: false, error: "" });
    onOpenChange(false);
  }

  async function lookUp(address: string) {
    if (lookup.working) return;
    const token = generation.current;
    setLookup({ working: true, error: "" });
    try {
      const userId = await lookUpPerson(address);
      if (token !== generation.current) return;
      close();
      onOpenPerson(userId);
    } catch (error) {
      if (token !== generation.current) return;
      setLookup({
        working: false,
        error: error instanceof Error ? error.message : "The lookup failed.",
      });
    }
  }

  function pick(index: number) {
    const chosen = results[index];
    if (!chosen) return;
    if (chosen.group === "directory") {
      void lookUp(chosen.item.id);
      return;
    }
    close();
    if (chosen.group === "pages") onNavigate(chosen.item.id);
    else if (chosen.group === "audits") onOpenAudit(chosen.item.id);
    else onOpenPerson(chosen.item.id);
  }

  if (!open) return null;
  return (
    <QuickSearchView
      idBase={idBase}
      query={query}
      groups={groups}
      active={marked}
      notes={notes}
      empty={
        fold(query).trim() !== "" && results.length === 0 && !stillLoading
          ? noMatchText(query)
          : null
      }
      status={resultStatus(query, results.length)}
      working={lookup.working}
      error={lookup.error}
      inputRef={input}
      onQuery={(value) => {
        setQuery(value);
        setActive(0);
        setLookup((current) => ({ ...current, error: "" }));
      }}
      onKeyDown={(event) =>
        handleDialogKey(
          event,
          { count: results.length, active: marked },
          { move: setActive, choose: pick, close },
        )
      }
      onPick={pick}
      onHover={setActive}
      onBackdrop={close}
    />
  );
}

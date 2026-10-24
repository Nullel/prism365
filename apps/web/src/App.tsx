import {
  type AutomationLifecycle,
  type AutomationPreview,
  type AutomationWorkflowList,
  automationPreviewSchema,
  automationWorkflowListSchema,
  type DashboardSummary,
  dashboardSummarySchema,
  type Metric,
  type PrismSession,
  sessionSchema,
} from "@prism365/contracts";
import { useCallback, useEffect, useId, useRef, useState } from "react";
import { Notifications } from "./Attention.tsx";
import { Audits } from "./Audits.tsx";
import { DirectoryChanges } from "./Changes.tsx";
import { DemoBanner } from "./demo/DemoBanner.tsx";
import { localDemoDashboard } from "./demo.ts";
import { Handoffs } from "./Handoffs.tsx";
import {
  ActivityIcon,
  ArrowUpRightIcon,
  BellIcon,
  ChecklistIcon,
  ChevronIcon,
  GridIcon,
  LockIcon,
  SearchIcon,
  SettingsIcon,
  ShieldIcon,
  SparkleIcon,
  UsersIcon,
} from "./icons.tsx";
import { LiveOverview } from "./LiveOverview.tsx";
import { People } from "./People.tsx";
import { ScriptPackages } from "./ScriptPackages.tsx";
import { SetupChecklist } from "./Setup.tsx";
import { SigninAudit } from "./SigninAudit.tsx";
import { TenantConnection } from "./TenantConnection.tsx";
import { UserIntelligence } from "./UserIntelligence.tsx";
import { ThisWeek } from "./Week.tsx";

export type WorkspaceView =
  | "overview"
  | "week"
  | "notifications"
  | "people"
  | "audits"
  | "signins"
  | "changes"
  | "users"
  | "handoffs"
  | "lifecycle"
  | "access";

interface NavItem {
  label: string;
  icon: typeof GridIcon;
  view: WorkspaceView;
}

// One navigation, in the sidebar. Only views that work are listed.
const navGroups: { label: string; items: NavItem[] }[] = [
  {
    label: "",
    items: [
      { label: "Overview", icon: GridIcon, view: "overview" },
      { label: "Notifications", icon: BellIcon, view: "notifications" },
      { label: "People", icon: UsersIcon, view: "people" },
    ],
  },
  {
    label: "Audit",
    items: [
      { label: "Audits", icon: ChecklistIcon, view: "audits" },
      { label: "Sign-in audit", icon: ShieldIcon, view: "signins" },
      { label: "Directory changes", icon: ActivityIcon, view: "changes" },
    ],
  },
  {
    label: "Tools",
    items: [
      { label: "Live lookup", icon: SearchIcon, view: "users" },
      { label: "Handoffs", icon: ArrowUpRightIcon, view: "handoffs" },
      { label: "Lifecycle automation", icon: SettingsIcon, view: "lifecycle" },
    ],
  },
];
const accessItem: NavItem = { label: "Microsoft access", icon: LockIcon, view: "access" };

const pageCopy: Record<WorkspaceView, { title: string; description: string }> = {
  overview: { title: "Overview", description: "What needs attention across the tenant right now." },
  week: {
    title: "This week",
    description:
      "What changed in the last seven days: new findings, accounts gone quiet, licence and MFA drift, directory changes and Microsoft deadlines.",
  },
  notifications: {
    title: "Notifications",
    description:
      "New Microsoft incidents, Message center deadlines, and what changed since you looked.",
  },
  people: {
    title: "People",
    description: "Everyone who signs in: where from, on which device, and anything flagged.",
  },
  audits: {
    title: "Audits",
    description:
      "Checks you run against the tenant when you need them: inactive accounts, unused licenses, MFA gaps, guests and license capacity.",
  },
  signins: {
    title: "Sign-in audit",
    description:
      "Collected sign-ins, each account's usual places, and the sign-ins that break them.",
  },
  changes: {
    title: "Directory changes",
    description:
      "Who changed which account, group, role, app or policy, kept for as long as Prism365 runs.",
  },
  handoffs: {
    title: "Handoffs",
    description:
      "Lists from audits, packaged for the tool that acts on them, and the receipts it sends back.",
  },
  users: {
    title: "Live lookup",
    description: "Look one account up directly in Microsoft, without stored history.",
  },
  lifecycle: {
    title: "Lifecycle automation",
    description: "Review onboarding and offboarding runbooks before anything changes.",
  },
  access: {
    title: "Microsoft access",
    description: "What this installation may read, and the session it reads with.",
  },
};

function initials(name: string | null | undefined): string {
  const parts = (name ?? "").trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return "P";
  return `${parts[0]?.[0] ?? ""}${parts.length > 1 ? (parts.at(-1)?.[0] ?? "") : ""}`.toUpperCase();
}

async function loadDashboard(signal: AbortSignal): Promise<DashboardSummary> {
  const response = await fetch("/api/v1/dashboard/summary", { signal });
  if (!response.ok) {
    throw new Error(`Dashboard request failed with HTTP ${response.status}.`);
  }
  return dashboardSummarySchema.parse(await response.json());
}

async function loadAutomationWorkflows(signal: AbortSignal): Promise<AutomationWorkflowList> {
  const response = await fetch("/api/v1/automation/workflows", { signal });
  if (!response.ok) {
    throw new Error(`Automation catalog request failed with HTTP ${response.status}.`);
  }
  return automationWorkflowListSchema.parse(await response.json());
}

async function createLifecyclePreview(
  workflowId: string,
  targetEmail: string,
  signal: AbortSignal,
): Promise<AutomationPreview> {
  const response = await fetch("/api/v1/automation/previews", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ workflowId, targetEmail }),
    signal,
  });
  if (!response.ok) {
    const body = (await response.json().catch(() => null)) as { message?: string } | null;
    throw new Error(body?.message ?? `Preview request failed with HTTP ${response.status}.`);
  }
  return automationPreviewSchema.parse(await response.json());
}

function Trend({ metric }: { metric: Metric }) {
  const arrow =
    metric.trend.direction === "up" ? "↗" : metric.trend.direction === "down" ? "↘" : "→";
  const signedValue = metric.trend.value === 0 ? "No change" : `${arrow} ${metric.trend.value}`;

  return (
    <span className={`trend trend--${metric.trend.sentiment}`}>
      <span>{signedValue}</span>
      <span className="trend__label">{metric.trend.label}</span>
    </span>
  );
}

function MetricCard({ metric, index }: { metric: Metric; index: number }) {
  const paths = [
    "M2 28C12 25 17 12 29 15c12 3 15-8 27-5 8 2 12 9 22 5",
    "M2 9c10 4 14 17 25 15 11-1 14-12 25-10 10 2 14 12 26 11",
    "M2 25c9-1 12-12 22-10s13 12 24 8 16-15 30-10",
    "M2 20c10-8 18 4 28-2s17-4 24 2 13 2 24-7",
  ];

  return (
    <article className="metric-card">
      <div className="metric-card__topline">
        <span>{metric.label}</span>
        <span className="data-chip">Demo</span>
      </div>
      <div className="metric-card__body">
        <strong>{metric.displayValue}</strong>
        <svg
          className="sparkline"
          viewBox="0 0 80 36"
          role="img"
          aria-label={`${metric.label} sparkline`}
        >
          <path d={paths[index % paths.length]} />
        </svg>
      </div>
      <Trend metric={metric} />
    </article>
  );
}

function SignInChart({ data }: { data: DashboardSummary["signInTrend"] }) {
  const max = Math.max(...data.map((point) => point.safe + point.review + point.elevated));

  return (
    <div className="chart" role="img" aria-label="Synthetic sign-ins by day, grouped by outcome">
      <div className="chart__guide chart__guide--top">
        <span>5k</span>
      </div>
      <div className="chart__guide chart__guide--middle">
        <span>2.5k</span>
      </div>
      <div className="chart__guide chart__guide--bottom">
        <span>0</span>
      </div>
      <div className="chart__columns">
        {data.map((point) => {
          const total = point.safe + point.review + point.elevated;
          const safePercent = (point.safe / total) * 100;
          const reviewPercent = (point.review / total) * 100;
          const elevatedPercent = (point.elevated / total) * 100;
          const height = Math.max(16, (total / max) * 100);
          return (
            <div className="chart__column" key={point.label}>
              <div
                className="chart__bar"
                style={{ height: `${height}%` }}
                title={`${point.label}: ${total.toLocaleString()} synthetic sign-ins`}
              >
                <span
                  className="chart__segment chart__segment--elevated"
                  style={{ height: `${elevatedPercent}%` }}
                />
                <span
                  className="chart__segment chart__segment--review"
                  style={{ height: `${reviewPercent}%` }}
                />
                <span
                  className="chart__segment chart__segment--safe"
                  style={{ height: `${safePercent}%` }}
                />
              </div>
              <span className="chart__label">{point.label}</span>
              <span className="visually-hidden">
                {point.safe} protected, {point.review} review, {point.elevated} elevated
              </span>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function LifecycleAutomation() {
  const emailId = useId();
  const workflowId = useId();
  const [catalog, setCatalog] = useState<AutomationWorkflowList | null>(null);
  const [lifecycle, setLifecycle] = useState<AutomationLifecycle>("offboarding");
  const [selectedWorkflow, setSelectedWorkflow] = useState("");
  const [targetEmail, setTargetEmail] = useState("alex.rivera@example.invalid");
  const [preview, setPreview] = useState<AutomationPreview | null>(null);
  const [state, setState] = useState<"loading" | "ready" | "previewing" | "error">("loading");
  const [error, setError] = useState("");
  const previewRequest = useRef(0);
  const pendingPreview = useRef<AbortController | null>(null);

  useEffect(() => () => pendingPreview.current?.abort(), []);

  useEffect(() => {
    const controller = new AbortController();
    loadAutomationWorkflows(controller.signal)
      .then((result) => {
        setCatalog(result);
        const firstOffboarding = result.items.find((item) => item.lifecycle === "offboarding");
        setSelectedWorkflow(firstOffboarding?.id ?? result.items[0]?.id ?? "");
        setState("ready");
      })
      .catch((loadError: unknown) => {
        if (loadError instanceof DOMException && loadError.name === "AbortError") return;
        setError("Start the local Prism365 API to load the reviewed runbook library.");
        setState("error");
      });
    return () => controller.abort();
  }, []);

  const visibleWorkflows =
    catalog?.items.filter((workflow) => workflow.lifecycle === lifecycle) ?? [];

  function chooseLifecycle(nextLifecycle: AutomationLifecycle) {
    setLifecycle(nextLifecycle);
    setSelectedWorkflow(
      catalog?.items.find((workflow) => workflow.lifecycle === nextLifecycle)?.id ?? "",
    );
    invalidatePreview();
  }

  function invalidatePreview() {
    previewRequest.current += 1;
    pendingPreview.current?.abort();
    setPreview(null);
    setError("");
    if (catalog) setState("ready");
  }

  async function handlePreview(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!selectedWorkflow) {
      setError("No reviewed runbook is available for this workflow.");
      return;
    }

    setState("previewing");
    setError("");
    setPreview(null);
    const request = ++previewRequest.current;
    pendingPreview.current?.abort();
    const controller = new AbortController();
    pendingPreview.current = controller;
    try {
      const result = await createLifecyclePreview(selectedWorkflow, targetEmail, controller.signal);
      if (request !== previewRequest.current) return;
      setPreview(result);
      setState("ready");
    } catch (previewError) {
      if (controller.signal.aborted || request !== previewRequest.current) return;
      setError(previewError instanceof Error ? previewError.message : "Unable to create preview.");
      setState("error");
    }
  }

  return (
    <section className="automation" aria-labelledby="automation-title">
      <div className="automation__header">
        <div>
          <h2 id="automation-title">Preview a runbook</h2>
          <p>
            Choose a reviewed runbook and inspect every proposed change before Prism365 can ask for
            approval.
          </p>
        </div>
        <span className="automation__lock">
          <LockIcon /> Preview only
        </span>
      </div>

      <div className="automation__grid">
        <form className="automation-form" onSubmit={handlePreview}>
          <fieldset className="automation-toggle">
            <legend>Choose a lifecycle</legend>
            <button
              aria-pressed={lifecycle === "onboarding"}
              className={lifecycle === "onboarding" ? "is-active" : ""}
              onClick={() => chooseLifecycle("onboarding")}
              type="button"
            >
              <span>Onboard</span>
              <small>Create access from an approved profile</small>
            </button>
            <button
              aria-pressed={lifecycle === "offboarding"}
              className={lifecycle === "offboarding" ? "is-active" : ""}
              onClick={() => chooseLifecycle("offboarding")}
              type="button"
            >
              <span>Offboard</span>
              <small>Block access and preserve a recovery path</small>
            </button>
          </fieldset>

          <label className="field" htmlFor={emailId}>
            <span>Work email</span>
            <input
              id={emailId}
              maxLength={254}
              onChange={(event) => {
                setTargetEmail(event.target.value);
                invalidatePreview();
              }}
              placeholder="person@company.com"
              required
              type="email"
              value={targetEmail}
            />
            <small>Demo input only. No directory lookup occurs.</small>
          </label>

          <label className="field" htmlFor={workflowId}>
            <span>Reviewed runbook</span>
            <select
              disabled={state === "loading" || visibleWorkflows.length === 0}
              id={workflowId}
              onChange={(event) => {
                setSelectedWorkflow(event.target.value);
                invalidatePreview();
              }}
              value={selectedWorkflow}
            >
              {visibleWorkflows.map((workflow) => (
                <option key={workflow.id} value={workflow.id}>
                  {workflow.name} · v{workflow.version}
                </option>
              ))}
            </select>
          </label>

          {visibleWorkflows.find((workflow) => workflow.id === selectedWorkflow)?.runtime ===
          "powershell7" ? (
            <div className="runbook-note">
              <span>PowerShell 7 package</span>
              <p>
                This synthetic imported package is reviewed and version-locked. Production scripts
                will run outside the API in an isolated worker.
              </p>
            </div>
          ) : null}

          {error ? (
            <p className="form-error" role="alert">
              {error}
            </p>
          ) : null}
          <button
            className="button button--primary button--full"
            disabled={state === "loading" || state === "previewing" || !selectedWorkflow}
            type="submit"
          >
            {state === "previewing" ? "Building preview…" : "Preview every step"}
            <ChevronIcon />
          </button>
          <p className="automation-form__footnote">
            Import review is scaffolded in the package contract but intentionally locked until the
            sandbox runner and approval ledger are ready.
          </p>
        </form>

        <div className="automation-preview" aria-live="polite">
          {preview ? (
            <>
              <div className="automation-preview__heading">
                <div>
                  <span className="eyebrow">Proposed plan</span>
                  <h3>{preview.workflow.name}</h3>
                  <p>{preview.targetEmail}</p>
                </div>
                <span>{preview.steps.length} steps</span>
              </div>
              <ol className="automation-steps">
                {preview.steps.map((step) => (
                  <li key={step.id}>
                    <span className="automation-steps__number">{step.order}</span>
                    <div>
                      <strong>{step.label}</strong>
                      <small>{step.permission}</small>
                      <details className="step-details">
                        <summary>Impact and recovery</summary>
                        <p>{step.description}</p>
                        <p>{step.recovery}</p>
                        <p>
                          {step.reversible
                            ? "Recovery is possible where provider state permits it."
                            : "This step cannot be fully reversed."}
                        </p>
                      </details>
                    </div>
                    <span className={`risk-chip risk-chip--${step.risk}`}>
                      {step.risk === "readOnly"
                        ? "Check"
                        : step.risk === "manualReview"
                          ? "Review"
                          : step.risk}
                    </span>
                  </li>
                ))}
              </ol>
              <div className="automation-preview__blocked">
                <ShieldIcon />
                <p>
                  <strong>Execution remains locked.</strong>
                  {preview.blockedReasons.map((reason) => (
                    <span key={reason}>{reason}</span>
                  ))}
                </p>
              </div>
              <details className="preview-limitations">
                <summary>Preview limitations</summary>
                <ul>
                  {preview.limitations.map((item) => (
                    <li key={item}>{item}</li>
                  ))}
                </ul>
              </details>
            </>
          ) : (
            <div className="automation-empty">
              <span>
                <UsersIcon />
              </span>
              <h3>Your next run, clearly laid out.</h3>
              <p>
                Prism365 will keep the fast email-first flow, while showing the exact impact,
                permissions, approvals, and recovery path before anything changes.
              </p>
              <ul>
                <li>Protected-account checks</li>
                <li>Typed confirmation and approval</li>
                <li>Immutable result and recovery record</li>
              </ul>
            </div>
          )}
        </div>
      </div>
    </section>
  );
}

export const signInFailedNotice =
  "Microsoft sign-in did not complete. Check that your account is on this installation's operator list and that the app registration's redirect URI matches this address, then try again.";

// The callback redirects to "/?auth=failed" for every rejected sign-in. Read it once, remove it
// from the address bar, and show one sentence; the server log holds the category.
export function readAuthNotice(): string {
  if (typeof window === "undefined") return "";
  const url = new URL(window.location.href);
  if (url.searchParams.get("auth") !== "failed") return "";
  url.searchParams.delete("auth");
  window.history.replaceState(null, "", `${url.pathname}${url.search}${url.hash}`);
  return signInFailedNotice;
}

// The page in the address bar, so links can open a view, an account or an audit, and the back
// button works: #/audits/accounts.inactive, #/people/<id>. Other hashes (the skip link) are not
// routes and leave the page alone.
export interface Route {
  view: WorkspaceView;
  person: string | null;
  audit: string | null;
}

const workspaceViews: readonly WorkspaceView[] = [
  "overview",
  "notifications",
  "people",
  "audits",
  "signins",
  "changes",
  "users",
  "handoffs",
  "lifecycle",
  "access",
];

export function parseRoute(hash: string): Route | null {
  if (!hash.startsWith("#/")) return null;
  const [name = "", ...rest] = hash.slice(2).split("/");
  const view = workspaceViews.find((item) => item === name);
  if (!view) return null;
  let detail: string | null = null;
  try {
    detail = rest.length > 0 && rest.join("/") ? decodeURIComponent(rest.join("/")) : null;
  } catch {
    detail = null;
  }
  return {
    view,
    person: view === "people" ? detail : null,
    audit: view === "audits" ? detail : null,
  };
}

export function routeHash(route: Route): string {
  const detail =
    route.view === "people" ? route.person : route.view === "audits" ? route.audit : null;
  return `#/${route.view}${detail ? `/${encodeURIComponent(detail)}` : ""}`;
}

const initialRoute = (): Route =>
  (typeof window === "undefined" ? null : parseRoute(window.location.hash)) ?? {
    view: "overview",
    person: null,
    audit: null,
  };

export function App() {
  const [dashboard, setDashboard] = useState<DashboardSummary>(localDemoDashboard);
  const [view, setView] = useState<WorkspaceView>(() => initialRoute().view);
  const [person, setPerson] = useState<string | null>(() => initialRoute().person);
  const [audit, setAudit] = useState<string | null>(() => initialRoute().audit);
  // Keep the address bar in step with the page, and follow back and forward. A load without
  // a route replaces its entry rather than adding one.
  useEffect(() => {
    const next = routeHash({ view, person, audit });
    if (window.location.hash === next) return;
    const url = `${window.location.pathname}${window.location.search}${next}`;
    if (parseRoute(window.location.hash)) window.history.pushState(null, "", url);
    else window.history.replaceState(null, "", url);
  }, [view, person, audit]);
  useEffect(() => {
    const follow = () => {
      const route = parseRoute(window.location.hash);
      if (!route) return;
      setView(route.view);
      setPerson(route.person);
      setAudit(route.audit);
    };
    window.addEventListener("popstate", follow);
    window.addEventListener("hashchange", follow);
    return () => {
      window.removeEventListener("popstate", follow);
      window.removeEventListener("hashchange", follow);
    };
  }, []);
  const openPerson = (userId: string) => {
    setPerson(userId);
    setView("people");
  };
  const navigate = (next: WorkspaceView) => {
    if (next === "people") setPerson(null);
    if (next === "audits") setAudit(null);
    setView(next);
  };
  const [session, setSession] = useState<PrismSession | null>(null);
  const [sessionError, setSessionError] = useState("");
  const [authNotice, setAuthNotice] = useState(readAuthNotice);
  const [loggingOut, setLoggingOut] = useState(false);
  const [unread, setUnread] = useState(0);

  const refreshSession = useCallback(async (signal?: AbortSignal) => {
    setSessionError("");
    try {
      const response = await fetch("/api/v1/session", {
        ...(signal ? { signal } : {}),
        credentials: "same-origin",
        cache: "no-store",
      });
      if (!response.ok) throw new Error("Unable to check the Prism365 connection.");
      const data = sessionSchema.parse(await response.json());
      if (!signal?.aborted) {
        setSession(data);
        if (data.authenticated) setAuthNotice("");
      }
    } catch (error) {
      if (signal?.aborted) return;
      setSessionError(error instanceof Error ? error.message : "Unable to check the connection.");
    }
  }, []);

  useEffect(() => {
    const controller = new AbortController();
    void refreshSession(controller.signal);
    return () => controller.abort();
  }, [refreshSession]);

  async function logout() {
    if (!session?.csrfToken) {
      setSessionError("Refresh this page before signing out.");
      return;
    }
    setLoggingOut(true);
    setSessionError("");
    try {
      const response = await fetch("/api/v1/auth/logout", {
        method: "POST",
        credentials: "same-origin",
        headers: { "x-csrf-token": session.csrfToken },
      });
      if (!response.ok) throw new Error("Sign-out failed. Try again.");
      setSession(null);
      await refreshSession();
    } catch (error) {
      setSessionError(error instanceof Error ? error.message : "Sign-out failed.");
    } finally {
      setLoggingOut(false);
    }
  }

  useEffect(() => {
    if (session?.mode !== "demo") return;
    const controller = new AbortController();
    loadDashboard(controller.signal)
      .then(setDashboard)
      .catch(() => {
        // The bundled synthetic dashboard stays in place.
      });
    return () => controller.abort();
  }, [session?.mode]);

  const live = session?.mode === "live" && session.authenticated;
  const refreshUnread = useCallback(async () => {
    try {
      const response = await fetch("/api/v1/attention?filter=unread&limit=1", {
        credentials: "same-origin",
        cache: "no-store",
      });
      if (!response.ok) return;
      const body = (await response.json()) as { counts?: { unread?: unknown } };
      if (typeof body.counts?.unread === "number") setUnread(body.counts.unread);
    } catch {
      // The badge is a convenience; the Notifications view reports its own errors.
    }
  }, []);
  useEffect(() => {
    if (!live) return;
    void refreshUnread();
    const timer = setInterval(() => void refreshUnread(), 60_000);
    return () => clearInterval(timer);
  }, [live, refreshUnread]);

  const generatedTime = new Intl.DateTimeFormat(undefined, {
    hour: "numeric",
    minute: "2-digit",
  }).format(new Date(dashboard.generatedAt));

  const demo = session?.mode === "demo";
  const signedIn = Boolean(session && (demo || session.authenticated));
  // Until someone is signed in, the connection card is the page.
  const gated = !signedIn;
  const copy = pageCopy[view];
  const showHeader =
    !(view === "people" && person && signedIn) && !(view === "audits" && audit && signedIn);
  const connection = (
    <TenantConnection
      session={session}
      error={sessionError || authNotice}
      onRetry={() => {
        setAuthNotice("");
        void refreshSession();
      }}
      onLogout={logout}
      loggingOut={loggingOut}
    />
  );
  const accountState = !session ? "checking" : demo ? "demo" : session.authenticated ? "on" : "off";
  const accountLabel = {
    checking: "Checking connection",
    demo: "Demo data",
    on: "Connected to Microsoft",
    off: "Not signed in",
  }[accountState];
  const navButton = (item: NavItem) => {
    const Icon = item.icon;
    const active = view === item.view;
    return (
      <button
        className={`nav__item${active ? " nav__item--active" : ""}`}
        type="button"
        onClick={() => navigate(item.view)}
        key={item.view}
        aria-current={active ? "page" : undefined}
      >
        <Icon />
        <span>{item.label}</span>
        {item.view === "notifications" && unread > 0 ? (
          <span className="nav__badge">
            {unread > 99 ? "99+" : unread}
            <span className="visually-hidden"> unread</span>
          </span>
        ) : null}
      </button>
    );
  };

  return (
    <div className="app-shell">
      <a className="skip-link" href="#main-content">
        Skip to main content
      </a>
      <aside className="sidebar" aria-label="Primary navigation">
        <a className="brand" href="#main-content" aria-label="Prism365 home">
          <span className="brand__mark">
            <span />
          </span>
          <span>Prism365</span>
        </a>

        <nav className="nav" aria-label="Workspace">
          {navGroups.map((group) => (
            <div className="nav__section" key={group.label || "main"}>
              {group.label ? <span className="nav__label">{group.label}</span> : null}
              {group.items.map(navButton)}
            </div>
          ))}
        </nav>

        <div className="sidebar__bottom">
          <div className="nav__section">{navButton(accessItem)}</div>
          <div className="account">
            <span className="account__avatar" aria-hidden="true">
              {demo ? "D" : initials(session?.displayName)}
            </span>
            <span className="account__text">
              <strong>{demo ? "Demo workspace" : (session?.displayName ?? "Prism365")}</strong>
              <small>
                <span className={`status-dot status-dot--${accountState}`} />
                {accountLabel}
              </small>
            </span>
          </div>
          {session?.authenticated ? (
            <button
              className="account__signout"
              type="button"
              onClick={() => void logout()}
              disabled={loggingOut}
            >
              {loggingOut ? "Signing out…" : "Sign out"}
            </button>
          ) : null}
          <a
            className="sidebar__footer"
            href="https://github.com/Nullel/prism365"
            target="_blank"
            rel="noreferrer"
          >
            Prism365 pre-alpha
            <ArrowUpRightIcon />
          </a>
        </div>
      </aside>

      <main className="main" id="main-content">
        <div className="content">
          {import.meta.env.VITE_DEMO === "true" ? <DemoBanner /> : null}
          {showHeader ? (
            <header className="page-header">
              <h1>{copy.title}</h1>
              <p>
                {demo && view === "overview"
                  ? "Synthetic data. No tenant is connected."
                  : copy.description}
              </p>
            </header>
          ) : null}

          {demo ? connection : null}
          {!demo && (gated || view === "access") ? connection : null}
          {signedIn && !demo && sessionError && view !== "access" ? (
            <p className="form-error form-error--banner" role="alert">
              {sessionError}
            </p>
          ) : null}

          {signedIn && session ? (
            <>
              {view === "overview" && !demo ? (
                <LiveOverview session={session} onOpenPerson={openPerson} onNavigate={navigate} />
              ) : null}
              {view === "notifications" && !demo ? (
                <Notifications
                  session={session}
                  onNavigate={navigate}
                  onCountsChanged={() => void refreshUnread()}
                />
              ) : null}
              {view === "notifications" && demo ? (
                <p className="empty-data">
                  Demo mode has no notifications, because Prism365 does not invent Microsoft
                  incidents.
                </p>
              ) : null}
              {view === "people" ? (
                <People
                  key={`${session.mode}-${session.authenticated}`}
                  session={session}
                  selected={person}
                  onSelect={setPerson}
                />
              ) : null}
              {view === "audits" ? (
                <Audits
                  key={`${session.mode}-${session.authenticated}`}
                  session={session}
                  selected={audit}
                  onSelect={setAudit}
                  onOpenPerson={openPerson}
                />
              ) : null}
              {view === "signins" ? (
                <SigninAudit
                  key={`${session.mode}-${session.authenticated}`}
                  session={session}
                  onOpenPerson={openPerson}
                />
              ) : null}
              {view === "week" ? (
                <ThisWeek
                  key={`${session.mode}-${session.authenticated}`}
                  session={session}
                  onNavigate={navigate}
                  onOpenPerson={openPerson}
                />
              ) : null}
              {view === "changes" ? (
                <DirectoryChanges
                  key={`${session.mode}-${session.authenticated}`}
                  session={session}
                  onOpenPerson={openPerson}
                />
              ) : null}
              {view === "handoffs" ? (
                <Handoffs key={`${session.mode}-${session.authenticated}`} session={session} />
              ) : null}
              {view === "users" ? (
                <UserIntelligence
                  key={`${session.mode}-${session.authenticated}`}
                  session={session}
                />
              ) : null}
              {view === "lifecycle" && demo ? <LifecycleAutomation /> : null}
              {view === "lifecycle" ? <ScriptPackages session={session} /> : null}
              {view === "access" ? (
                <SetupChecklist
                  key={`${session.mode}-${session.authenticated}`}
                  session={session}
                />
              ) : null}
              {view === "access" ? (
                <section className="panel access-note" aria-label="Write access">
                  <ShieldIcon />
                  <p>
                    <strong>Read-only by design.</strong> Prism365 asks Microsoft only for read
                    permissions, and the Graph client refuses every write. Purview audit search
                    would be a separate, optional connector.
                  </p>
                </section>
              ) : null}
            </>
          ) : null}

          {view === "overview" && demo ? (
            <>
              <section className="metrics" aria-label="Tenant summary">
                {dashboard.metrics.map((metric, index) => (
                  <MetricCard metric={metric} index={index} key={metric.id} />
                ))}
              </section>

              <div className="dashboard-grid">
                <section className="panel panel--chart">
                  <div className="panel__header">
                    <div>
                      <span className="eyebrow">Authentication pulse</span>
                      <h2>Sign-in outcomes</h2>
                    </div>
                    <fieldset className="legend">
                      <legend className="visually-hidden">Chart legend</legend>
                      <span>
                        <i className="legend__safe" />
                        Protected
                      </span>
                      <span>
                        <i className="legend__review" />
                        Review
                      </span>
                      <span>
                        <i className="legend__elevated" />
                        Elevated
                      </span>
                    </fieldset>
                  </div>
                  <SignInChart data={dashboard.signInTrend} />
                  <div className="panel__footer">
                    <span>
                      {dashboard.timeWindow} · generated {generatedTime}
                    </span>
                    <span>
                      Counts are operational signals, not measures of employee performance.
                    </span>
                  </div>
                </section>

                <section className="insight-card">
                  <div className="insight-card__glow" />
                  <div className="insight-card__icon">
                    <SparkleIcon />
                  </div>
                  <span className="eyebrow">Evidence note</span>
                  <h2>A quieter window—with one pattern worth reviewing.</h2>
                  <p>
                    Synthetic elevated sign-ins decreased across the selected period. Two
                    legacy-client events remain open.
                  </p>
                  <div className="insight-card__evidence">
                    <span>Observed fact</span>
                    <strong>7 elevated events on Sunday</strong>
                    <small>
                      Demo sign-in events · selected window · coverage may be incomplete
                    </small>
                  </div>
                  <button className="text-button" type="button" disabled>
                    Open investigation <ArrowUpRightIcon />
                  </button>
                </section>
              </div>

              <div className="lower-grid">
                <section className="panel panel--activity">
                  <div className="panel__header">
                    <div>
                      <span className="eyebrow">Evidence stream</span>
                      <h2>Recent activity</h2>
                    </div>
                    <span className="muted-label">Synthetic only</span>
                  </div>
                  <div className="activity-list">
                    {dashboard.recentActivity.map((activity) => (
                      <article className="activity-row" key={activity.id}>
                        <span
                          className={`activity-row__marker activity-row__marker--${activity.severity}`}
                        />
                        <div>
                          <div className="activity-row__title">
                            <strong>{activity.title}</strong>
                            <span className={`severity severity--${activity.severity}`}>
                              {activity.severity}
                            </span>
                          </div>
                          <p>{activity.detail}</p>
                          <small>{activity.evidence}</small>
                        </div>
                        <time dateTime={activity.observedAt}>
                          {new Intl.DateTimeFormat(undefined, {
                            hour: "numeric",
                            minute: "2-digit",
                          }).format(new Date(activity.observedAt))}
                        </time>
                      </article>
                    ))}
                  </div>
                </section>

                <section className="panel panel--coverage">
                  <div className="panel__header">
                    <div>
                      <span className="eyebrow">Data contract</span>
                      <h2>Coverage & consent</h2>
                    </div>
                  </div>
                  <div className="coverage-score">
                    <strong>3</strong>
                    <span>
                      of 5 read packs
                      <br />
                      represented in demo
                    </span>
                  </div>
                  <div
                    className="coverage-bar"
                    aria-label="Three of five demo permission packs represented"
                    aria-valuemax={5}
                    aria-valuemin={0}
                    aria-valuenow={3}
                    role="progressbar"
                  >
                    <span />
                  </div>
                  <ul className="coverage-list">
                    <li>
                      <span className="check">✓</span>
                      <div>
                        <strong>Directory profiles</strong>
                        <small>User.Read.All</small>
                      </div>
                    </li>
                    <li>
                      <span className="check">✓</span>
                      <div>
                        <strong>Sign-in audit</strong>
                        <small>AuditLog.Read.All</small>
                      </div>
                    </li>
                    <li>
                      <span className="check">✓</span>
                      <div>
                        <strong>Usage reports</strong>
                        <small>Reports.Read.All</small>
                      </div>
                    </li>
                    <li>
                      <span className="empty-check" />
                      <div>
                        <strong>Purview audit</strong>
                        <small>Separate optional connector</small>
                      </div>
                    </li>
                  </ul>
                  <button
                    className="button button--secondary button--full"
                    onClick={() => navigate("access")}
                    type="button"
                  >
                    Review Microsoft access
                  </button>
                </section>
              </div>
            </>
          ) : null}
        </div>
      </main>
    </div>
  );
}

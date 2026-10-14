import {
  automationPreviewSchema,
  automationWorkflowListSchema,
  dashboardSummarySchema,
  sessionSchema,
  type AutomationLifecycle,
  type AutomationPreview,
  type AutomationWorkflowList,
  type DashboardSummary,
  type Metric,
  type PrismSession,
} from "@prism365/contracts";
import { useCallback, useEffect, useId, useRef, useState } from "react";
import {
  ActivityIcon,
  ArrowUpRightIcon,
  ChevronIcon,
  GridIcon,
  LockIcon,
  ReportIcon,
  SearchIcon,
  SettingsIcon,
  ShieldIcon,
  SparkleIcon,
  UsersIcon,
} from "./icons.tsx";
import { localDemoDashboard } from "./demo.ts";
import { TenantConnection } from "./TenantConnection.tsx";
import { UserIntelligence } from "./UserIntelligence.tsx";
import { ScriptPackages } from "./ScriptPackages.tsx";

type WorkspaceView = "overview" | "users" | "lifecycle";

const navSections = [
  {
    label: "Workspace",
    items: [
      { label: "Overview", icon: GridIcon, view: "overview" },
      { label: "User Intelligence", icon: UsersIcon, view: "users" },
      { label: "Lifecycle automation", icon: ActivityIcon, view: "lifecycle" },
      { label: "Security", icon: ShieldIcon },
      { label: "Audit stream", icon: ActivityIcon },
    ],
  },
  {
    label: "Explore",
    items: [
      { label: "Reports", icon: ReportIcon },
      { label: "Configuration", icon: SettingsIcon },
    ],
  },
] as const;

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

function PermissionDialog({ onClose }: { onClose: () => void }) {
  const titleId = useId();
  const dialog = useRef<HTMLDialogElement>(null);

  useEffect(() => {
    const trigger = document.activeElement;
    dialog.current?.showModal();
    return () => {
      dialog.current?.close();
      if (trigger instanceof HTMLElement) trigger.focus();
    };
  }, []);

  return (
    <dialog
      ref={dialog}
      aria-labelledby={titleId}
      className="permission-dialog"
      onCancel={(event) => {
        event.preventDefault();
        onClose();
      }}
    >
      <div className="permission-dialog__header">
        <div>
          <span className="eyebrow">Access model</span>
          <h2 id={titleId}>Permission by purpose</h2>
        </div>
        <button className="icon-button" onClick={onClose} type="button" aria-label="Close dialog">
          ×
        </button>
      </div>
      <p className="dialog-intro">
        A tenant administrator enables each read-only feature separately. Microsoft also checks the
        signed-in administrator’s role, license, and service availability.
      </p>
      <div className="permission-list">
        {[
          ["User inventory", "User.Read.All", "Selected directory profile fields"],
          ["Sign-in audit", "AuditLog.Read.All", "Security and directory audit evidence"],
          ["Usage reports", "Reports.Read.All", "Tenant-level service usage trends"],
        ].map(([label, permission, detail]) => (
          <div className="permission-row" key={permission}>
            <span className="permission-row__icon">
              <LockIcon />
            </span>
            <div>
              <strong>{label}</strong>
              <span>{detail}</span>
            </div>
            <code>{permission}</code>
          </div>
        ))}
      </div>
      <div className="dialog-note">
        <ShieldIcon />
        <p>
          <strong>Write access is disabled.</strong> Purview audit search is a separate, optional
          connector.
        </p>
      </div>
      <button className="button button--primary button--full" onClick={onClose} type="button">
        Understood
      </button>
    </dialog>
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
          <span className="eyebrow">Lifecycle automation</span>
          <h2 id="automation-title">One clear path from email to completed checklist.</h2>
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

export function App() {
  const [dashboard, setDashboard] = useState<DashboardSummary>(localDemoDashboard);
  const [apiState, setApiState] = useState<"loading" | "connected" | "local">("loading");
  const [permissionDialogOpen, setPermissionDialogOpen] = useState(false);
  const [view, setView] = useState<WorkspaceView>("users");
  const [session, setSession] = useState<PrismSession | null>(null);
  const [sessionError, setSessionError] = useState("");
  const [loggingOut, setLoggingOut] = useState(false);

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
      if (!signal?.aborted) setSession(data);
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
      .then((data) => {
        setDashboard(data);
        setApiState("connected");
      })
      .catch((error: unknown) => {
        if (error instanceof DOMException && error.name === "AbortError") return;
        setApiState("local");
      });
    return () => controller.abort();
  }, [session?.mode]);

  const generatedTime = new Intl.DateTimeFormat(undefined, {
    hour: "numeric",
    minute: "2-digit",
  }).format(new Date(dashboard.generatedAt));

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

        <div className="workspace-switcher">
          <span className="workspace-avatar">P</span>
          <span>
            <strong>{session?.mode === "demo" ? "Northstar Labs" : "Your organization"}</strong>
            <small>{session?.mode === "demo" ? "Demo workspace" : "Microsoft 365 workspace"}</small>
          </span>
        </div>

        <nav className="nav">
          {navSections.map((section) => (
            <div className="nav__section" key={section.label}>
              <span className="nav__label">{section.label}</span>
              {section.items.map((item) => {
                const Icon = item.icon;
                return "view" in item ? (
                  <button
                    className={`nav__item${view === item.view ? " nav__item--active" : ""}`}
                    type="button"
                    onClick={() => setView(item.view)}
                    key={item.label}
                    aria-current={view === item.view ? "page" : undefined}
                  >
                    <Icon />
                    <span>{item.label}</span>
                  </button>
                ) : (
                  <span
                    className="nav__item nav__item--muted"
                    key={item.label}
                    title="Planned for a future milestone"
                  >
                    <Icon />
                    <span>{item.label}</span>
                    <small>Soon</small>
                  </span>
                );
              })}
            </div>
          ))}
        </nav>

        <div className="sidebar__principle">
          <LockIcon />
          <p>
            <strong>Privacy by default</strong>
            <span>Read-only. Minimal consent. No worker scoring.</span>
          </p>
        </div>
        <a
          className="sidebar__footer"
          href="https://github.com/nlj3/prism365"
          target="_blank"
          rel="noreferrer"
        >
          <span className="open-source-dot" />
          Open-source pre-alpha
          <ArrowUpRightIcon />
        </a>
      </aside>

      <main className="main" id="main-content">
        <header className="topbar">
          <div className="mobile-brand">
            <span className="brand__mark">
              <span />
            </span>
            Prism365
          </div>
          <button
            className="search"
            type="button"
            aria-label="Search is planned for a future milestone"
            disabled
          >
            <SearchIcon />
            <span>Search reports, users, or settings</span>
            <kbd>⌘ K</kbd>
          </button>
          <div className="topbar__actions">
            <span className={`api-status api-status--${apiState}`}>
              <span />
              {session?.mode === "live"
                ? session.authenticated
                  ? "Microsoft connected"
                  : "Live setup"
                : apiState === "connected"
                  ? "Demo API"
                  : apiState === "loading"
                    ? "Checking API"
                    : "Local preview"}
            </span>
            <span className="avatar-button" aria-hidden="true">
              {session?.displayName?.slice(0, 2).toUpperCase() ?? "P"}
            </span>
          </div>
        </header>

        <div className="content">
          <section className="page-heading">
            <div>
              <span className="eyebrow">Good morning, administrator</span>
              <h1>
                {view === "users"
                  ? "One account. Clear evidence."
                  : view === "lifecycle"
                    ? "Make the next change clear."
                    : "Your tenant, clearly understood."}
              </h1>
              <p>
                {view === "users"
                  ? "Directory details, license inventory, and sign-in context in one focused investigation."
                  : view === "lifecycle"
                    ? "Review onboarding and offboarding steps before taking action."
                    : "Security posture, service signals, and evidence—without the noise."}
              </p>
            </div>
            <div className="page-heading__actions">
              <button
                className="button button--secondary"
                onClick={() => setPermissionDialogOpen(true)}
                type="button"
              >
                <LockIcon /> Access model
              </button>
            </div>
          </section>

          <TenantConnection
            session={session}
            error={sessionError}
            onRetry={() => {
              void refreshSession();
            }}
            onLogout={logout}
            loggingOut={loggingOut}
          />

          <nav className="workspace-tabs" aria-label="Workspace views">
            {(
              [
                ["users", "User Intelligence"],
                ["overview", "Overview"],
                ["lifecycle", "Lifecycle automation"],
              ] as const
            ).map(([key, label]) => (
              <button
                key={key}
                type="button"
                aria-current={view === key ? "page" : undefined}
                className={view === key ? "is-active" : ""}
                onClick={() => setView(key)}
              >
                {label}
              </button>
            ))}
          </nav>

          {view === "users" && session ? (
            <UserIntelligence key={`${session.mode}-${session.authenticated}`} session={session} />
          ) : null}
          {view === "lifecycle" && session?.mode === "demo" ? <LifecycleAutomation /> : null}
          {view === "lifecycle" && session ? <ScriptPackages session={session} /> : null}
          {view === "overview" && session?.mode === "live" ? (
            <section className="live-placeholder">
              <GridIcon />
              <h2>Tenant-wide reporting is next.</h2>
              <p>
                The first live source is available in User Intelligence. This overview will appear
                when tenant-wide collection and reporting are connected.
              </p>
              <button
                className="button button--primary"
                type="button"
                onClick={() => setView("users")}
              >
                Open User Intelligence
              </button>
            </section>
          ) : null}
          {view === "overview" && session?.mode === "demo" ? (
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
                    onClick={() => setPermissionDialogOpen(true)}
                    type="button"
                  >
                    Review permission design
                  </button>
                </section>
              </div>
            </>
          ) : null}

          <footer className="page-footer">
            <span>Prism365 pre-alpha · built in the open</span>
            <span>
              {session?.mode === "live"
                ? "Live evidence · administration writes disabled · private installation"
                : "Demo mode · writes disabled · no tenant data stored"}
            </span>
          </footer>
        </div>
      </main>
      {permissionDialogOpen ? (
        <PermissionDialog onClose={() => setPermissionDialogOpen(false)} />
      ) : null}
    </div>
  );
}

import { dashboardSummarySchema, type DashboardSummary, type Metric } from "@prism365/contracts";
import { useEffect, useId, useState } from "react";
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

const navSections = [
  {
    label: "Workspace",
    items: [
      { label: "Overview", icon: GridIcon, active: true },
      { label: "Identity", icon: UsersIcon },
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

  useEffect(() => {
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === "Escape") onClose();
    };
    window.addEventListener("keydown", closeOnEscape);
    return () => window.removeEventListener("keydown", closeOnEscape);
  }, [onClose]);

  return (
    <div className="modal-backdrop">
      <section
        aria-labelledby={titleId}
        aria-modal="true"
        className="permission-dialog"
        onMouseDown={(event) => event.stopPropagation()}
        role="dialog"
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
          Nothing is connected in this demo. A tenant administrator enables each read-only feature
          separately.
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
      </section>
    </div>
  );
}

export function App() {
  const [dashboard, setDashboard] = useState<DashboardSummary>(localDemoDashboard);
  const [apiState, setApiState] = useState<"loading" | "connected" | "local">("loading");
  const [permissionDialogOpen, setPermissionDialogOpen] = useState(false);

  useEffect(() => {
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
  }, []);

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

        <button
          className="workspace-switcher"
          type="button"
          aria-label="Current workspace: Northstar Labs demo"
        >
          <span className="workspace-avatar">N</span>
          <span>
            <strong>Northstar Labs</strong>
            <small>Demo workspace</small>
          </span>
          <span className="workspace-switcher__chevrons">
            ⌃<br />⌄
          </span>
        </button>

        <nav className="nav">
          {navSections.map((section) => (
            <div className="nav__section" key={section.label}>
              <span className="nav__label">{section.label}</span>
              {section.items.map((item) => {
                const Icon = item.icon;
                return "active" in item && item.active ? (
                  <a
                    className="nav__item nav__item--active"
                    href="#main-content"
                    key={item.label}
                    aria-current="page"
                  >
                    <Icon />
                    <span>{item.label}</span>
                  </a>
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
          href="https://github.com/nlj3"
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
              {apiState === "connected"
                ? "Demo API"
                : apiState === "loading"
                  ? "Checking API"
                  : "Local preview"}
            </span>
            <button className="avatar-button" type="button" aria-label="Demo administrator profile">
              NL
            </button>
          </div>
        </header>

        <div className="content">
          <section className="page-heading">
            <div>
              <span className="eyebrow">Good morning, administrator</span>
              <h1>Your tenant, clearly understood.</h1>
              <p>Security posture, service signals, and evidence—without the noise.</p>
            </div>
            <div className="page-heading__actions">
              <button
                className="button button--secondary"
                onClick={() => setPermissionDialogOpen(true)}
                type="button"
              >
                <LockIcon /> Access model
              </button>
              <button
                className="button button--primary"
                type="button"
                disabled
                title="Tenant onboarding arrives in Milestone 1"
              >
                Connect tenant <ChevronIcon />
              </button>
            </div>
          </section>

          <section className="demo-banner" aria-label="Demo data notice">
            <div className="demo-banner__icon">
              <SparkleIcon />
            </div>
            <div>
              <strong>Exploring with synthetic data</strong>
              <span>
                No Microsoft 365 tenant is connected. Every identity, event, and metric shown here
                is fictional.
              </span>
            </div>
            <button onClick={() => setPermissionDialogOpen(true)} type="button">
              See how access works <ChevronIcon />
            </button>
          </section>

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
                <span>Counts are operational signals, not measures of employee performance.</span>
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
                Synthetic elevated sign-ins decreased across the selected period. Two legacy-client
                events remain open.
              </p>
              <div className="insight-card__evidence">
                <span>Observed fact</span>
                <strong>7 elevated events on Sunday</strong>
                <small>Demo sign-in events · selected window · coverage may be incomplete</small>
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

          <footer className="page-footer">
            <span>Prism365 pre-alpha · built in the open</span>
            <span>Demo mode · writes disabled · no tenant data stored</span>
          </footer>
        </div>
      </main>
      {permissionDialogOpen ? (
        <PermissionDialog onClose={() => setPermissionDialogOpen(false)} />
      ) : null}
    </div>
  );
}

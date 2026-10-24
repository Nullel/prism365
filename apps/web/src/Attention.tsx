import {
  type AttentionItem,
  type AttentionList,
  type AttentionPreferences,
  type AttentionReviewState,
  type AttentionSettings,
  attentionListSchema,
  attentionSettingsSchema,
  type PrismSession,
  type SystemAlert,
} from "@prism365/contracts";
import { useCallback, useEffect, useState } from "react";
import { day, post, readJson, when } from "./common.tsx";

// Notification center: Admin attention on the Overview, the Notifications inbox, and its settings.
// Microsoft's status and labels are shown as Microsoft's; the administrator's own state (read,
// reviewed, done, muted) is shown separately and never changes what Microsoft said.

const classTitles: Record<string, string> = {
  incident: "Service incidents",
  advisory: "Service advisories",
  "message.action": "Message center: action required",
  "message.retirement": "Message center: retirements",
  "message.change": "Message center: major changes",
};

const reviewLabels: Record<AttentionReviewState, string> = {
  new: "Not reviewed",
  reviewed: "Reviewed",
  done: "Done",
};

export function useAttention(
  session: PrismSession,
  filter: "open" | "unread" | "all" | "muted" = "open",
) {
  const [list, setList] = useState<AttentionList | null>(null);
  const [error, setError] = useState("");
  const load = useCallback(
    async (signal?: AbortSignal) => {
      if (session.mode !== "live" || !session.authenticated) return;
      try {
        setList(
          await readJson(
            `/api/v1/attention?filter=${filter}&limit=200`,
            (value) => attentionListSchema.parse(value),
            signal ? { signal } : {},
          ),
        );
        setError("");
      } catch (loadError) {
        if (signal?.aborted) return;
        setError(loadError instanceof Error ? loadError.message : "Notifications could not load.");
      }
    },
    [session.mode, session.authenticated, filter],
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
  return { list, error, reload: load };
}

function AttentionCard({
  item,
  session,
  onChanged,
}: {
  item: AttentionItem;
  session: PrismSession;
  onChanged: () => void;
}) {
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  async function act(path: string, body?: unknown) {
    setBusy(true);
    setError("");
    try {
      await readJson(`/api/v1/attention/${item.id}/${path}`, (value) => value, post(session, body));
      onChanged();
    } catch (actionError) {
      setError(actionError instanceof Error ? actionError.message : "The change did not save.");
    } finally {
      setBusy(false);
    }
  }
  const overdue = item.provider.deadline && Date.parse(item.provider.deadline) < Date.now();
  return (
    <article className={`attention${item.unread ? " attention--unread" : ""}`}>
      <div className="attention__topline">
        {item.unread ? (
          <span className="attention__dot">
            <span className="visually-hidden">Unread</span>
          </span>
        ) : null}
        <span className={`finding-severity finding-severity--${item.severity}`}>
          {item.severity}
        </span>
        {item.services.slice(0, 3).map((service) => (
          <span key={service} className="attention__service">
            {service}
          </span>
        ))}
        {item.services.length > 3 ? (
          <span className="person-muted">+{item.services.length - 3} more</span>
        ) : null}
      </div>
      <h3>{item.title}</h3>
      <div className="attention__labels">
        <span className="person-muted">Microsoft:</span>
        {item.provider.status ? <span className="person-tag">{item.provider.status}</span> : null}
        {item.provider.labels.map((label) => (
          <span key={label} className="person-tag">
            {label}
          </span>
        ))}
        {item.provider.deadline ? (
          <span className={`person-tag${overdue ? " person-tag--alert" : ""}`}>
            Action required by {day(item.provider.deadline)}
          </span>
        ) : null}
        {!item.provider.active ? <span className="person-tag person-tag--ok">Closed</span> : null}
      </div>
      {item.localLabels.length ? (
        <div className="attention__labels">
          <span className="person-muted">Prism365:</span>
          {item.localLabels.map((label) => (
            <span key={label} className="person-tag attention__local">
              {label}
            </span>
          ))}
        </div>
      ) : null}
      {item.changeNote ? (
        <p className="attention__change">
          <strong>Changed {when(item.changedAt)}:</strong> {item.changeNote}
        </p>
      ) : null}
      <p className="attention__summary">{item.provider.impact ?? item.summary}</p>
      <p className="incident__meta">
        {item.ref} · Microsoft updated {when(item.provider.updatedAt)} · You:{" "}
        {reviewLabels[item.review.state]}
        {item.muted ? ", muted" : ""}
        {item.link ? (
          <>
            {" · "}
            <a href={item.link} target="_blank" rel="noreferrer">
              Open in Microsoft 365 admin center
            </a>
          </>
        ) : null}
      </p>
      {error ? (
        <p className="form-error" role="alert">
          {error}
        </p>
      ) : null}
      <div className="finding__actions">
        {item.review.state !== "reviewed" ? (
          <button
            className="button button--secondary"
            type="button"
            disabled={busy}
            onClick={() => void act("review", { state: "reviewed" })}
          >
            Mark reviewed
          </button>
        ) : null}
        {item.review.state !== "done" ? (
          <button
            className="button button--secondary"
            type="button"
            disabled={busy}
            onClick={() => void act("review", { state: "done" })}
          >
            Mark done
          </button>
        ) : (
          <button
            className="button button--secondary"
            type="button"
            disabled={busy}
            onClick={() => void act("review", { state: "new" })}
          >
            Reopen
          </button>
        )}
        <button
          className="button button--secondary"
          type="button"
          disabled={busy}
          onClick={() => void act("mute", { muted: !item.muted })}
        >
          {item.muted ? "Unmute" : "Mute"}
        </button>
        {item.unread ? (
          <button
            className="link-button"
            type="button"
            disabled={busy}
            onClick={() => void act("read")}
          >
            Mark read
          </button>
        ) : null}
      </div>
    </article>
  );
}

function SystemAlertRow({
  alert,
  onNavigate,
}: {
  alert: SystemAlert;
  onNavigate: (view: "signins") => void;
}) {
  return (
    <li className="system-alert">
      <span className={`finding-severity finding-severity--${alert.severity}`}>
        {alert.severity}
      </span>
      <div>
        <strong>{alert.title}</strong>
        <p>
          {alert.detail}
          {alert.observedAt ? ` Last check ${when(alert.observedAt)}.` : ""}
        </p>
      </div>
      {alert.action ? (
        alert.action.href.startsWith("view:") ? (
          <button
            className="button button--secondary"
            type="button"
            onClick={() => onNavigate("signins")}
          >
            {alert.action.label}
          </button>
        ) : (
          <a className="button button--secondary" href={alert.action.href}>
            {alert.action.label}
          </a>
        )
      ) : null}
    </li>
  );
}

// Overview section: open items grouped by kind, plus problems that stop Prism365 from checking.
// Only categories whose connector exists and is set up appear.
export function AdminAttention({
  session,
  onNavigate,
}: {
  session: PrismSession;
  onNavigate: (view: "signins" | "notifications") => void;
}) {
  const { list, error, reload } = useAttention(session, "open");
  const groups = new Map<string, AttentionItem[]>();
  for (const item of list?.items ?? []) {
    // Incidents have their own panel above; advisories and Message center live here.
    if (item.class === "incident" || item.class === "advisory") continue;
    groups.set(item.class, [...(groups.get(item.class) ?? []), item]);
  }
  const unset = (list?.sources ?? []).filter((source) => source.state === "not_enabled");
  return (
    <section className="panel admin-attention" aria-labelledby="attention-title">
      <div className="overview-list__header">
        <h2 id="attention-title">Admin attention</h2>
        <button className="link-button" type="button" onClick={() => onNavigate("notifications")}>
          All notifications
        </button>
      </div>
      {error ? <p className="form-error">{error}</p> : null}
      {!list ? (
        <p className="empty-data">Loading…</p>
      ) : (
        <>
          {list.system.length > 0 ? (
            <div className="attention-group">
              <h3>Prism365 monitoring</h3>
              <ul className="system-alerts">
                {list.system.map((alert) => (
                  <SystemAlertRow
                    key={alert.id}
                    alert={alert}
                    onNavigate={() => onNavigate("signins")}
                  />
                ))}
              </ul>
            </div>
          ) : null}
          {[...groups].map(([itemClass, items]) => (
            <div className="attention-group" key={itemClass}>
              <h3>
                {classTitles[itemClass] ?? "Other items"} ({items.length})
              </h3>
              <ul className="attention-list">
                {items.slice(0, 5).map((item) => (
                  <li key={item.id}>
                    <AttentionCard item={item} session={session} onChanged={() => void reload()} />
                  </li>
                ))}
              </ul>
              {items.length > 5 ? (
                <button
                  className="link-button"
                  type="button"
                  onClick={() => onNavigate("notifications")}
                >
                  {items.length - 5} more in Notifications
                </button>
              ) : null}
            </div>
          ))}
          {list.system.length === 0 && groups.size === 0 ? (
            <p className="empty-data">Nothing needs your attention from the connected sources.</p>
          ) : null}
          {unset.length > 0 ? (
            <p className="signin-coverage">
              Not set up: {unset.map((source) => source.title).join(", ")}. Enable them in
              ENABLED_READ_FEATURES to add their items here.
            </p>
          ) : null}
        </>
      )}
    </section>
  );
}

const filters = [
  ["unread", "Unread"],
  ["open", "Open"],
  ["all", "All"],
  ["muted", "Muted"],
] as const;

export function Notifications({
  session,
  onNavigate,
  onCountsChanged,
}: {
  session: PrismSession;
  onNavigate: (view: "signins") => void;
  onCountsChanged: () => void;
}) {
  const [tab, setTab] = useState<"inbox" | "settings">("inbox");
  const [filter, setFilter] = useState<(typeof filters)[number][0]>("unread");
  const { list, error, reload } = useAttention(session, filter);
  const refresh = () => {
    void reload();
    onCountsChanged();
  };
  async function readAll() {
    await readJson("/api/v1/attention/read-all", (value) => value, post(session));
    refresh();
  }
  return (
    <div className="notifications">
      <fieldset className="signin-range">
        <legend className="visually-hidden">Notifications view</legend>
        {(
          [
            ["inbox", "Inbox"],
            ["settings", "Settings"],
          ] as const
        ).map(([value, label]) => (
          <button
            key={value}
            type="button"
            aria-pressed={tab === value}
            className={tab === value ? "is-active" : ""}
            onClick={() => setTab(value)}
          >
            {label}
          </button>
        ))}
      </fieldset>

      {tab === "settings" ? (
        <NotificationSettings session={session} />
      ) : (
        <section className="panel notifications__inbox" aria-label="Notifications">
          <div className="overview-list__header">
            <fieldset className="signin-range">
              <legend className="visually-hidden">Show</legend>
              {filters.map(([value, label]) => (
                <button
                  key={value}
                  type="button"
                  aria-pressed={filter === value}
                  className={filter === value ? "is-active" : ""}
                  onClick={() => setFilter(value)}
                >
                  {label}
                  {value === "unread" && list ? ` (${list.counts.unread})` : ""}
                </button>
              ))}
            </fieldset>
            {list && list.counts.unread > 0 ? (
              <button className="link-button" type="button" onClick={() => void readAll()}>
                Mark all read
              </button>
            ) : null}
          </div>
          {error ? <p className="form-error">{error}</p> : null}
          {list && list.system.length > 0 ? (
            <ul className="system-alerts">
              {list.system.map((alert) => (
                <SystemAlertRow
                  key={alert.id}
                  alert={alert}
                  onNavigate={() => onNavigate("signins")}
                />
              ))}
            </ul>
          ) : null}
          {!list ? (
            <p className="empty-data">Loading…</p>
          ) : list.items.length === 0 ? (
            <p className="empty-data">
              {filter === "unread"
                ? "No unread notifications. New incidents and meaningful changes appear here."
                : "Nothing here."}
            </p>
          ) : (
            <>
              <ul className="attention-list">
                {list.items.map((item) => (
                  <li key={item.id}>
                    <AttentionCard item={item} session={session} onChanged={refresh} />
                  </li>
                ))}
              </ul>
              {list.items.length >= 200 ? (
                <p className="signin-coverage">
                  Showing the first 200. Mark some done to see more.
                </p>
              ) : null}
            </>
          )}
        </section>
      )}
    </div>
  );
}

const notifyLabels: [keyof AttentionPreferences["notify"], string, string][] = [
  [
    "incident",
    "Service incidents",
    "A Microsoft incident appears, changes meaningfully, or is resolved.",
  ],
  ["advisory", "Service advisories", "The same, for advisories (usually limited impact)."],
  ["message.action", "Message center: action required", "A post with a Microsoft deadline."],
  ["message.retirement", "Message center: retirements", "A post Microsoft tags as a retirement."],
  ["message.change", "Message center: major changes", "Major planned changes without a deadline."],
];

function NotificationSettings({ session }: { session: PrismSession }) {
  const [settings, setSettings] = useState<AttentionSettings | null>(null);
  const [draft, setDraft] = useState<AttentionPreferences | null>(null);
  const [state, setState] = useState<"idle" | "saving" | "saved">("idle");
  const [error, setError] = useState("");
  useEffect(() => {
    const controller = new AbortController();
    readJson("/api/v1/settings/attention", (value) => attentionSettingsSchema.parse(value), {
      signal: controller.signal,
    }).then(
      (value) => {
        setSettings(value);
        setDraft(value.preferences);
      },
      (loadError: unknown) => {
        if (!controller.signal.aborted)
          setError(loadError instanceof Error ? loadError.message : "Settings could not load.");
      },
    );
    return () => controller.abort();
  }, []);
  if (!settings || !draft)
    return <section className="panel notification-settings">{error || "Loading…"}</section>;
  const services = [...new Set([...settings.services, ...draft.services])].sort();
  async function save() {
    if (!draft) return;
    setState("saving");
    setError("");
    try {
      const saved = await readJson(
        "/api/v1/settings/attention",
        (value) => attentionSettingsSchema.parse(value),
        post(session, draft),
      );
      setSettings(saved);
      setDraft(saved.preferences);
      setState("saved");
    } catch (saveError) {
      setError(saveError instanceof Error ? saveError.message : "Settings did not save.");
      setState("idle");
    }
  }
  return (
    <section className="panel notification-settings" aria-label="Notification settings">
      <fieldset>
        <legend>Notify me about</legend>
        {notifyLabels.map(([key, label, detail]) => (
          <label key={key} className="check-row">
            <input
              type="checkbox"
              checked={draft.notify[key]}
              onChange={(event) =>
                setDraft({ ...draft, notify: { ...draft.notify, [key]: event.target.checked } })
              }
            />
            <span>
              <strong>{label}</strong>
              <small>{detail}</small>
            </span>
          </label>
        ))}
      </fieldset>
      <fieldset>
        <legend>Important services</legend>
        <p className="person-muted">
          With none selected, every service can notify. With some selected, only those do, and their
          items carry an "Important service" label from Prism365.
        </p>
        {services.length === 0 ? (
          <p className="empty-data">
            Microsoft's service list appears here after the first service-health check.
          </p>
        ) : (
          <div className="service-grid">
            {services.map((service) => (
              <label key={service} className="check-row check-row--compact">
                <input
                  type="checkbox"
                  checked={draft.services.includes(service)}
                  onChange={(event) =>
                    setDraft({
                      ...draft,
                      services: event.target.checked
                        ? [...draft.services, service]
                        : draft.services.filter((item) => item !== service),
                    })
                  }
                />
                <span>{service}</span>
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
        <button
          className="button button--primary"
          type="button"
          disabled={state === "saving"}
          onClick={() => void save()}
        >
          {state === "saving" ? "Saving…" : "Save settings"}
        </button>
        {state === "saved" ? <span className="person-muted">Saved.</span> : null}
      </div>
    </section>
  );
}

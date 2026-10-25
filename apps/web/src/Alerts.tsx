import {
  type AlertChoices,
  type AlertOverview,
  type AlertRule,
  type AlertRuleInput,
  alertOverviewSchema,
  alertRuleListSchema,
  type DeliveryError,
  type DeliveryRecord,
  type Destination,
  type DestinationCreate,
  type DestinationCreated,
  type DestinationKind,
  destinationCreatedSchema,
  directoryAlertRuleIds,
  type PrismSession,
  type Severity,
} from "@prism365/contracts";
import { type FormEvent, useCallback, useEffect, useId, useState } from "react";
import { CopyButton } from "./CopyButton.tsx";
import { number, post, RelativeTime, readJson, SkeletonRows } from "./common.tsx";

// Settings, Alerts: where Prism365 sends word of something that needs attention, and when. A
// destination is an email list on the operator's own relay or a webhook (Teams, Slack, or any
// HTTPS endpoint); a rule says what to watch and where to send it. Prism365 never sends mail
// through Microsoft. A webhook's address is shown once, as typed, and never again: afterwards
// only its host.

const PATH = "/api/v1/alerts";

export const kindLabels: Record<DestinationKind, string> = {
  email: "Email",
  teams: "Teams workflow",
  slack: "Slack",
  generic: "Webhook",
};

export const errorText: Record<DeliveryError, string> = {
  invalid_url: "The address is not a valid HTTPS address.",
  port_not_allowed: "The address uses a port the server does not allow.",
  dns_failed: "The host name could not be looked up.",
  blocked_address: "The host resolves to a private or reserved network, so nothing was sent.",
  connection_failed: "The destination could not be reached.",
  tls_failed: "The destination's certificate could not be verified.",
  timeout: "The destination did not answer in time.",
  redirect: "The destination answered with a redirect, which Prism365 does not follow.",
  rejected: "The destination refused the message.",
  rate_limited: "The destination asked Prism365 to slow down.",
  server_error: "The destination had an error of its own.",
  secret_unreadable:
    "The saved address cannot be read with this server's key. Delete the destination and add it again.",
  email_not_configured: "Email is not set up on this server.",
  recipient_not_allowed: "A recipient is outside the domains this server allows.",
  smtp_auth_failed: "The mail relay refused the sign-in.",
  smtp_rejected: "The mail relay refused the message.",
  too_large: "The message was too large for the destination.",
  destination_deleted: "The destination was deleted before the message went out.",
  ledger_failed: "The access ledger could not be written, so nothing was sent.",
  queue_full: "Too many messages were waiting, so this one was not sent.",
  dropped:
    "More findings arrived at once than Prism365 holds, so these were left out, the lowest severity first.",
  outcome_unknown:
    "The mail relay stopped answering after it took the message, so it may have arrived. It was not sent again.",
  interrupted: "Prism365 restarted before the message went out.",
  error: "Something went wrong while sending.",
};

const severityLabels: Record<Severity, string> = {
  info: "Info",
  low: "Low",
  medium: "Medium",
  high: "High",
  critical: "Critical",
};

export const directoryLabels: Record<(typeof directoryAlertRuleIds)[number], string> = {
  "role.privileged-assigned": "A privileged role was assigned",
  "app.credential-added": "An app was given a new secret or certificate",
  "app.consent-granted": "Consent was granted to an app",
  "policy.conditional-access-changed": "A Conditional Access policy changed",
  "domain.federation-changed": "A domain's federation settings changed",
  "auth-method.removed-by-admin": "An administrator removed sign-in methods",
};

const quietChoices = [
  { minutes: 15, label: "15 minutes" },
  { minutes: 60, label: "1 hour" },
  { minutes: 240, label: "4 hours" },
  { minutes: 1_440, label: "1 day" },
  { minutes: 10_080, label: "7 days" },
];

export function destinationWhere(destination: Destination): string {
  if (destination.kind === "email") return destination.recipients.join(", ");
  return destination.host ?? "";
}

export function ruleSummary(rule: AlertRule, choices: AlertChoices): string {
  if (rule.kind === "finding") {
    const modules = rule.modules
      .map((id) => choices.modules.find((module) => module.id === id)?.title ?? id)
      .join(", ");
    return `New or escalated findings at ${severityLabels[rule.minimumSeverity].toLowerCase()} or above${modules ? ` from ${modules}` : ""}.`;
  }
  if (rule.kind === "directory")
    return `Directory changes: ${rule.ruleIds.map((id) => directoryLabels[id].toLowerCase()).join("; ")}.`;
  if (rule.kind === "collection") {
    const parts = [
      rule.stoppedHours !== null
        ? `collection stopped for ${rule.stoppedHours} hour${rule.stoppedHours === 1 ? "" : "s"}`
        : null,
      rule.gaps ? "a module could not read part of its source" : null,
    ].filter(Boolean);
    return `When ${parts.join(", or when ")}.`;
  }
  const audit = choices.audits.find((item) => item.id === rule.auditId);
  const tile = audit?.tiles.find((item) => item.id === rule.tileId)?.label ?? rule.tileId;
  const title = audit?.title ?? rule.auditId;
  if (rule.condition === "change")
    return `${title}: ${tile} moves by ${number.format(rule.value)} or more.`;
  return `${title}: ${tile} goes ${rule.condition} ${number.format(rule.value)}.`;
}

function quietLabel(minutes: number): string {
  return quietChoices.find((choice) => choice.minutes === minutes)?.label ?? `${minutes} minutes`;
}

function deliverySource(record: DeliveryRecord, rules: readonly AlertRule[]): string {
  if (record.source === "test") return "Test";
  if (record.source === "export") return "Scheduled export";
  if (record.source === "summary") return "Weekly summary";
  return rules.find((rule) => rule.id === record.ruleId)?.name ?? "A deleted rule";
}

export function deliveryStatus(record: DeliveryRecord): string {
  if (record.status === "sent") return "Sent";
  if (record.status === "pending")
    return record.attempts === 0 ? "Waiting to send" : "Will try again";
  return "Failed";
}

// The signing secret of a webhook just created. It is shown this once.
export function CreatedSecret({
  created,
  onDone,
}: {
  created: DestinationCreated;
  onDone: () => void;
}) {
  if (!created.signingSecret) return null;
  return (
    <div className="alert-secret" role="status">
      <p className="alert-secret__text">
        {created.destination.label} was added. Its signing secret is below. Copy it now: Prism365
        shows it only once. The receiver uses it to check that a request came from Prism365.
      </p>
      <div className="alert-secret__value">
        <code>{created.signingSecret}</code>
        <CopyButton text={created.signingSecret} label="Copy the signing secret" />
      </div>
      <button className="button button--secondary" type="button" onClick={onDone}>
        I have copied it
      </button>
    </div>
  );
}

export function DestinationForm({
  settings,
  busy,
  onCreate,
}: {
  settings: AlertOverview["settings"];
  busy: boolean;
  onCreate: (input: DestinationCreate) => void;
}) {
  const id = useId();
  const [kind, setKind] = useState<DestinationKind>("teams");
  const [label, setLabel] = useState("");
  const [address, setAddress] = useState("");
  const [details, setDetails] = useState(false);
  const submit = (event: FormEvent) => {
    event.preventDefault();
    onCreate(
      kind === "email"
        ? {
            kind,
            label,
            recipients: address
              .split(/[,;\s]+/)
              .map((part) => part.trim())
              .filter(Boolean),
            includeDetails: details,
          }
        : { kind, label, url: address.trim(), includeDetails: details },
    );
  };
  const emailOff = !settings.email.configured;
  return (
    <form className="alert-form" onSubmit={submit} aria-labelledby={`${id}-title`}>
      <h3 className="alert-form__title" id={`${id}-title`}>
        Add a destination
      </h3>
      <label className="field">
        <span>Kind</span>
        <select value={kind} onChange={(event) => setKind(event.target.value as DestinationKind)}>
          <option value="teams">Teams workflow</option>
          <option value="slack">Slack</option>
          <option value="generic">Webhook (Prism365 JSON)</option>
          <option value="email" disabled={emailOff}>
            Email{emailOff ? " (not set up on this server)" : ""}
          </option>
        </select>
      </label>
      <label className="field">
        <span>Name</span>
        <input
          value={label}
          maxLength={80}
          required
          placeholder="Security channel"
          onChange={(event) => setLabel(event.target.value)}
        />
      </label>
      <label className="field">
        <span>{kind === "email" ? "Recipients" : "Webhook address"}</span>
        <input
          value={address}
          required
          autoComplete="off"
          spellCheck={false}
          placeholder={
            kind === "email" ? "security@example.invalid" : "https://hooks.example/your-webhook"
          }
          onChange={(event) => setAddress(event.target.value)}
        />
        <small>
          {kind === "email"
            ? `Up to 10 addresses, separated by commas, in ${settings.email.domains.join(", ") || "the allowed domains"}.`
            : "Paste the full HTTPS address. It carries the webhook's own token, so after saving Prism365 shows only its host."}
        </small>
      </label>
      <label className="check-row">
        <input
          type="checkbox"
          checked={details}
          onChange={(event) => setDetails(event.target.checked)}
        />
        <span>
          <strong>Include details</strong>
          <small>
            Adds the evidence fields, such as names and sign-in details. Without it, an alert holds
            only its title, severity, counts, what it is about and a link back. IP addresses are cut
            to their network unless this is email with details on.
          </small>
        </span>
      </label>
      <button className="button button--primary" type="submit" disabled={busy}>
        Add destination
      </button>
    </form>
  );
}

export function DestinationList({
  destinations,
  busy,
  onTest,
  onDelete,
}: {
  destinations: readonly Destination[];
  busy: boolean;
  onTest: (id: string) => void;
  onDelete: (id: string) => void;
}) {
  if (destinations.length === 0)
    return <p className="empty-data">No destinations yet. Add one below.</p>;
  return (
    <ul className="alert-list">
      {destinations.map((destination) => (
        <li key={destination.id}>
          <span className="alert-list__main">
            <strong>{destination.label}</strong>
            <span className="person-muted">
              {kindLabels[destination.kind]}, {destinationWhere(destination)}
              {destination.includeDetails ? ", with details" : ""}
            </span>
            {destination.readable ? null : (
              <span className="form-error">{errorText.secret_unreadable}</span>
            )}
            {destination.recipientsAllowed ? null : (
              <span className="form-error">
                A recipient is refused by this server now, so nothing is sent here. Delete the
                destination and add it again with allowed addresses.
              </span>
            )}
          </span>
          <button
            className="button button--secondary"
            type="button"
            disabled={busy}
            onClick={() => onTest(destination.id)}
          >
            Send test
          </button>
          <button
            className="button button--secondary"
            type="button"
            disabled={busy}
            aria-label={`Delete ${destination.label}`}
            onClick={() => onDelete(destination.id)}
          >
            Delete
          </button>
        </li>
      ))}
    </ul>
  );
}

type RuleKind = AlertRuleInput["kind"];

export function RuleForm({
  destinations,
  choices,
  busy,
  onCreate,
}: {
  destinations: readonly Destination[];
  choices: AlertChoices;
  busy: boolean;
  onCreate: (input: AlertRuleInput) => void;
}) {
  const id = useId();
  const [kind, setKind] = useState<RuleKind>("finding");
  const [name, setName] = useState("");
  const [targets, setTargets] = useState<string[]>([]);
  const [quiet, setQuiet] = useState(60);
  const [severity, setSeverity] = useState<Severity>("high");
  const [changes, setChanges] = useState<string[]>([]);
  const [stoppedHours, setStoppedHours] = useState("6");
  const [gaps, setGaps] = useState(true);
  const [auditId, setAuditId] = useState(choices.audits[0]?.id ?? "");
  const [tileId, setTileId] = useState("");
  const [condition, setCondition] = useState<"above" | "below" | "change">("above");
  const [value, setValue] = useState("10");
  const audit = choices.audits.find((item) => item.id === auditId);
  const toggle = (list: string[], item: string, on: boolean) =>
    on ? [...list, item] : list.filter((entry) => entry !== item);
  const submit = (event: FormEvent) => {
    event.preventDefault();
    const common = { name, destinations: targets, quietMinutes: quiet };
    if (kind === "finding") onCreate({ kind, ...common, minimumSeverity: severity, modules: [] });
    else if (kind === "directory")
      onCreate({
        kind,
        ...common,
        ruleIds: changes as (typeof directoryAlertRuleIds)[number][],
      });
    else if (kind === "collection")
      onCreate({
        kind,
        ...common,
        stoppedHours: stoppedHours === "" ? null : Number(stoppedHours),
        gaps,
      });
    else
      onCreate({
        kind,
        ...common,
        auditId,
        tileId: tileId || audit?.tiles[0]?.id || "",
        condition,
        value: Number(value),
      });
  };
  if (destinations.length === 0)
    return <p className="empty-data">Add a destination first, then a rule that sends to it.</p>;
  return (
    <form className="alert-form" onSubmit={submit} aria-labelledby={`${id}-title`}>
      <h3 className="alert-form__title" id={`${id}-title`}>
        Add a rule
      </h3>
      <label className="field">
        <span>Watch for</span>
        <select value={kind} onChange={(event) => setKind(event.target.value as RuleKind)}>
          <option value="finding">New or escalated findings</option>
          <option value="directory">Directory changes</option>
          <option value="collection">Collection problems</option>
          <option value="audit">An audit number</option>
        </select>
      </label>
      <label className="field">
        <span>Name</span>
        <input
          value={name}
          maxLength={80}
          required
          placeholder="High findings to the security channel"
          onChange={(event) => setName(event.target.value)}
        />
      </label>
      {kind === "finding" ? (
        <label className="field">
          <span>At or above</span>
          <select
            value={severity}
            onChange={(event) => setSeverity(event.target.value as Severity)}
          >
            {(["critical", "high", "medium", "low", "info"] as const).map((level) => (
              <option key={level} value={level}>
                {severityLabels[level]}
              </option>
            ))}
          </select>
        </label>
      ) : null}
      {kind === "directory" ? (
        <fieldset className="alert-choices">
          <legend>Changes</legend>
          {directoryAlertRuleIds.map((ruleId) => (
            <label className="check-row check-row--compact" key={ruleId}>
              <input
                type="checkbox"
                checked={changes.includes(ruleId)}
                onChange={(event) => setChanges(toggle(changes, ruleId, event.target.checked))}
              />
              <span>{directoryLabels[ruleId]}</span>
            </label>
          ))}
        </fieldset>
      ) : null}
      {kind === "collection" ? (
        <>
          <label className="field">
            <span>Collection stopped for at least (hours)</span>
            <input
              type="number"
              min={1}
              max={168}
              value={stoppedHours}
              onChange={(event) => setStoppedHours(event.target.value)}
            />
            <small>
              Only while collection is meant to be on. Leave empty to watch only for gaps.
            </small>
          </label>
          <label className="check-row">
            <input
              type="checkbox"
              checked={gaps}
              onChange={(event) => setGaps(event.target.checked)}
            />
            <span>
              <strong>Gaps</strong>
              <small>A module could not read part of its source.</small>
            </span>
          </label>
        </>
      ) : null}
      {kind === "audit" ? (
        <>
          <label className="field">
            <span>Audit</span>
            <select
              value={auditId}
              onChange={(event) => {
                setAuditId(event.target.value);
                setTileId("");
              }}
            >
              {choices.audits.map((item) => (
                <option key={item.id} value={item.id}>
                  {item.title}
                </option>
              ))}
            </select>
          </label>
          {audit && audit.tiles.length > 0 ? (
            <label className="field">
              <span>Number</span>
              <select value={tileId} onChange={(event) => setTileId(event.target.value)}>
                {audit.tiles.map((tile) => (
                  <option key={tile.id} value={tile.id}>
                    {tile.label}
                  </option>
                ))}
              </select>
            </label>
          ) : (
            <p className="empty-data">Run this audit once to choose one of its numbers.</p>
          )}
          <label className="field">
            <span>When it</span>
            <select
              value={condition}
              onChange={(event) => setCondition(event.target.value as typeof condition)}
            >
              <option value="above">Goes above</option>
              <option value="below">Goes below</option>
              <option value="change">Moves by at least</option>
            </select>
            <small>Compared with the previous run of the same audit version and scope.</small>
          </label>
          <label className="field">
            <span>Value</span>
            <input
              type="number"
              min={0}
              value={value}
              onChange={(event) => setValue(event.target.value)}
            />
          </label>
        </>
      ) : null}
      <fieldset className="alert-choices">
        <legend>Send to</legend>
        {destinations.map((destination) => (
          <label className="check-row check-row--compact" key={destination.id}>
            <input
              type="checkbox"
              checked={targets.includes(destination.id)}
              onChange={(event) =>
                setTargets(toggle(targets, destination.id, event.target.checked))
              }
            />
            <span>{destination.label}</span>
          </label>
        ))}
      </fieldset>
      <label className="field">
        <span>Quiet period</span>
        <select value={quiet} onChange={(event) => setQuiet(Number(event.target.value))}>
          {quietChoices.map((choice) => (
            <option key={choice.minutes} value={choice.minutes}>
              {choice.label}
            </option>
          ))}
        </select>
        <small>At most one alert per rule and subject in this time.</small>
      </label>
      <button
        className="button button--primary"
        type="submit"
        disabled={
          busy ||
          targets.length === 0 ||
          (kind === "audit" && !audit?.tiles.length) ||
          (kind === "directory" && changes.length === 0) ||
          (kind === "collection" && stoppedHours === "" && !gaps)
        }
      >
        Add rule
      </button>
    </form>
  );
}

export function RuleList({
  rules,
  destinations,
  choices,
  busy,
  onToggle,
  onDelete,
}: {
  rules: readonly AlertRule[];
  destinations: readonly Destination[];
  choices: AlertChoices;
  busy: boolean;
  onToggle: (id: string, enabled: boolean) => void;
  onDelete: (id: string) => void;
}) {
  if (rules.length === 0)
    return (
      <p className="empty-data">
        No rules yet. Add one below to hear about findings or audit numbers.
      </p>
    );
  return (
    <ul className="alert-list">
      {rules.map((rule) => (
        <li key={rule.id}>
          <label className="check-row alert-list__main">
            <input
              type="checkbox"
              checked={rule.enabled}
              disabled={busy}
              aria-label={`${rule.enabled ? "Turn off" : "Turn on"} ${rule.name}`}
              onChange={(event) => onToggle(rule.id, event.target.checked)}
            />
            <span>
              <strong>{rule.name}</strong>
              <small>{ruleSummary(rule, choices)}</small>
              <small>
                To{" "}
                {rule.destinations
                  .map(
                    (id) =>
                      destinations.find((item) => item.id === id)?.label ?? "a deleted destination",
                  )
                  .join(", ")}
                . Quiet for {quietLabel(rule.quietMinutes)} after each alert.
              </small>
            </span>
          </label>
          <button
            className="button button--secondary"
            type="button"
            disabled={busy}
            aria-label={`Delete ${rule.name}`}
            onClick={() => onDelete(rule.id)}
          >
            Delete
          </button>
        </li>
      ))}
    </ul>
  );
}

const ruleIdPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// Stored rules this version keeps but never runs, so a newer or damaged setting cannot widen an
// alert. They can be deleted, not edited.
export function OtherRuleList({
  rules,
  busy,
  onDelete,
}: {
  rules: AlertOverview["otherRules"];
  busy: boolean;
  onDelete: (id: string) => void;
}) {
  if (rules.length === 0) return null;
  return (
    <ul className="alert-list">
      {rules.map((rule, index) => (
        <li key={rule.id ?? `other-${index}`}>
          <span className="alert-list__main">
            <strong>{rule.name ?? "A rule without a name"}</strong>
            <small className="person-muted">
              {rule.newer
                ? "Made by a newer version of Prism365, so this version keeps it but does not run it."
                : "This version cannot read it, so it keeps it but does not run it."}
            </small>
          </span>
          {rule.id && ruleIdPattern.test(rule.id) ? (
            <button
              className="button button--secondary"
              type="button"
              disabled={busy}
              aria-label={`Delete ${rule.name ?? "this rule"}`}
              onClick={() => onDelete(rule.id as string)}
            >
              Delete
            </button>
          ) : null}
        </li>
      ))}
    </ul>
  );
}

export function DeliveryTable({
  deliveries,
  destinations,
  rules,
}: {
  deliveries: readonly DeliveryRecord[];
  destinations: readonly Destination[];
  rules: readonly AlertRule[];
}) {
  if (deliveries.length === 0)
    return (
      <p className="empty-data">
        Nothing has been sent yet. Deliveries appear here once a rule or an export fires.
      </p>
    );
  return (
    <table className="audit-table alert-deliveries">
      <thead>
        <tr>
          <th scope="col">When</th>
          <th scope="col">What</th>
          <th scope="col">To</th>
          <th scope="col">Items</th>
          <th scope="col">Status</th>
        </tr>
      </thead>
      <tbody>
        {deliveries.map((record) => (
          <tr key={record.id}>
            <td>
              <RelativeTime value={record.createdAt} />
            </td>
            <td>{deliverySource(record, rules)}</td>
            <td>
              {destinations.find((item) => item.id === record.destinationId)?.label ??
                "A deleted destination"}
            </td>
            <td>{number.format(record.items)}</td>
            <td>
              <span
                className={
                  record.status === "failed"
                    ? "person-tag person-tag--alert"
                    : record.status === "sent"
                      ? "person-tag person-tag--ok"
                      : "person-tag"
                }
              >
                {deliveryStatus(record)}
              </span>
              {record.lastError ? (
                <small className="alert-deliveries__why">
                  {errorText[record.lastError]}
                  {record.attempts > 1 ? ` ${record.attempts} attempts.` : ""}
                </small>
              ) : null}
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function DestinationSelect({
  label,
  value,
  destinations,
  busy,
  onChange,
}: {
  label: string;
  value: string | null;
  destinations: readonly Destination[];
  busy: boolean;
  onChange: (destinationId: string | null) => void;
}) {
  return (
    <label className="field alert-schedule">
      <span>{label}</span>
      <select
        value={value ?? ""}
        disabled={busy}
        onChange={(event) => onChange(event.target.value || null)}
      >
        <option value="">Not sent</option>
        {destinations.map((destination) => (
          <option key={destination.id} value={destination.id}>
            {destination.label}
          </option>
        ))}
      </select>
    </label>
  );
}

// Where the weekly summary and each audit's scheduled runs go.
export function ScheduledPanel({
  overview,
  busy,
  onExport,
  onWeekly,
}: {
  overview: AlertOverview;
  busy: boolean;
  onExport: (auditId: string, destinationId: string | null) => void;
  onWeekly: (destinationId: string | null) => void;
}) {
  const id = useId();
  const { destinations, exports, weekly, choices } = overview;
  return (
    <section className="panel alert-panel" aria-labelledby={id}>
      <h2 id={id}>Weekly summary and scheduled exports</h2>
      {destinations.length === 0 ? (
        <p className="empty-data">Add a destination first.</p>
      ) : (
        <>
          <p className="empty-data">
            The weekly summary goes out every Monday at 07:00 ({weekly.timeZone}) with the counts
            from This week and no names.
          </p>
          <DestinationSelect
            label="Weekly summary"
            value={weekly.destinationId}
            destinations={destinations}
            busy={busy}
            onChange={onWeekly}
          />
          <p className="empty-data">
            After each scheduled run of an audit, its result can go to a destination. Email with
            details on gets the CSV attached, up to 5 MB; every other destination gets the counts
            and a link back.
          </p>
          {choices.audits.map((audit) => (
            <DestinationSelect
              key={audit.id}
              label={audit.title}
              value={exports.find((entry) => entry.auditId === audit.id)?.destinationId ?? null}
              destinations={destinations}
              busy={busy}
              onChange={(destinationId) => onExport(audit.id, destinationId)}
            />
          ))}
        </>
      )}
    </section>
  );
}

export function AlertsView({
  overview,
  busy,
  error,
  created,
  onCreateDestination,
  onCreatedDone,
  onTest,
  onDeleteDestination,
  onCreateRule,
  onToggleRule,
  onDeleteRule,
  onExport,
  onWeekly,
  resetKey = 0,
}: {
  overview: AlertOverview;
  busy: boolean;
  error: string;
  created: DestinationCreated | null;
  onCreateDestination: (input: DestinationCreate) => void;
  onCreatedDone: () => void;
  onTest: (id: string) => void;
  onDeleteDestination: (id: string) => void;
  onCreateRule: (input: AlertRuleInput) => void;
  onToggleRule: (id: string, enabled: boolean) => void;
  onDeleteRule: (id: string) => void;
  onExport: (auditId: string, destinationId: string | null) => void;
  onWeekly: (destinationId: string | null) => void;
  // Changed after a save, so the forms start empty again and a typed webhook address does not
  // stay on the page.
  resetKey?: number;
}) {
  const ids = useId();
  const { settings, destinations, rules, deliveries, choices } = overview;
  return (
    <div className="alerts-page">
      {error ? (
        <p className="form-error form-error--banner" role="alert">
          {error}
        </p>
      ) : null}
      {settings.temporaryKey ? (
        <p className="empty-data">
          This server has no STATE_ENCRYPTION_KEY, so saved webhook addresses are encrypted with a
          key held in memory and stop working after a restart.
        </p>
      ) : null}
      <section className="panel alert-panel" aria-labelledby={`${ids}-destinations`}>
        <h2 id={`${ids}-destinations`}>Destinations</h2>
        <p className="empty-data">
          Email goes through this server's own mail relay. Webhooks go to Teams, Slack or any HTTPS
          address outside your private network. Prism365 never sends through Microsoft.
        </p>
        {created ? <CreatedSecret created={created} onDone={onCreatedDone} /> : null}
        <DestinationList
          destinations={destinations}
          busy={busy}
          onTest={onTest}
          onDelete={onDeleteDestination}
        />
        <DestinationForm
          key={`destination-${resetKey}`}
          settings={settings}
          busy={busy}
          onCreate={onCreateDestination}
        />
      </section>
      <section className="panel alert-panel" aria-labelledby={`${ids}-rules`}>
        <h2 id={`${ids}-rules`}>Rules</h2>
        <RuleList
          rules={rules}
          destinations={destinations}
          choices={choices}
          busy={busy}
          onToggle={onToggleRule}
          onDelete={onDeleteRule}
        />
        <OtherRuleList rules={overview.otherRules} busy={busy} onDelete={onDeleteRule} />
        <RuleForm
          key={`rule-${resetKey}`}
          destinations={destinations}
          choices={choices}
          busy={busy}
          onCreate={onCreateRule}
        />
      </section>
      <ScheduledPanel overview={overview} busy={busy} onExport={onExport} onWeekly={onWeekly} />
      <section className="panel alert-panel" aria-labelledby={`${ids}-deliveries`}>
        <h2 id={`${ids}-deliveries`}>Recent deliveries</h2>
        <p className="empty-data">
          What went where, and how it went. Prism365 keeps no copy of the messages themselves.
        </p>
        <DeliveryTable deliveries={deliveries} destinations={destinations} rules={rules} />
      </section>
    </div>
  );
}

function request(session: PrismSession, method: "POST" | "DELETE", body?: unknown): RequestInit {
  const init = post(session, body);
  return { ...init, method };
}

export function Alerts({ session }: { session: PrismSession }) {
  // The public demo replays a live-mode server, so its session says "live"; the build flag is
  // what marks it.
  const demo = session.mode === "demo" || import.meta.env.VITE_DEMO === "true";
  const [overview, setOverview] = useState<AlertOverview | null>(null);
  const [state, setState] = useState<"loading" | "ready" | "unavailable">("loading");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [created, setCreated] = useState<DestinationCreated | null>(null);
  const [resetKey, setResetKey] = useState(0);

  const load = useCallback(async (signal?: AbortSignal) => {
    try {
      const value = await readJson(PATH, (body) => alertOverviewSchema.parse(body), {
        ...(signal ? { signal } : {}),
      });
      setOverview(value);
      setState("ready");
    } catch {
      if (!signal?.aborted) setState("unavailable");
    }
  }, []);

  useEffect(() => {
    if (demo) return;
    const controller = new AbortController();
    void load(controller.signal);
    return () => controller.abort();
  }, [load, demo]);

  const act = async (work: () => Promise<unknown>) => {
    setBusy(true);
    setError("");
    try {
      await work();
      setResetKey((value) => value + 1);
      await load();
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : "The request failed.");
    } finally {
      setBusy(false);
    }
  };

  if (demo)
    return (
      <p className="empty-data">
        The demo sends no alerts. Alerts go out through your own mail relay or an HTTPS webhook, and
        the demo has neither. On a live installation this page sets up those destinations, the rules
        that use them and scheduled exports, and lists recent deliveries.
      </p>
    );
  if (state === "loading") return <SkeletonRows rows={5} label="Loading alert settings" />;
  if (state === "unavailable" || !overview)
    return (
      <p className="empty-data">
        Alerts need live mode with module storage. They could not be read just now.
      </p>
    );
  return (
    <AlertsView
      overview={overview}
      busy={busy}
      error={error}
      created={created}
      resetKey={resetKey}
      onCreatedDone={() => setCreated(null)}
      onCreateDestination={(input) =>
        void act(async () => {
          setCreated(
            await readJson(
              `${PATH}/destinations`,
              (body) => destinationCreatedSchema.parse(body),
              post(session, input),
            ),
          );
        })
      }
      onTest={(id) =>
        void act(() => readJson(`${PATH}/destinations/${id}/test`, (body) => body, post(session)))
      }
      onDeleteDestination={(id) =>
        void act(() =>
          readJson(`${PATH}/destinations/${id}`, (body) => body, request(session, "DELETE")),
        )
      }
      onCreateRule={(input) =>
        void act(() =>
          readJson(
            `${PATH}/rules`,
            (body) => alertRuleListSchema.parse(body),
            post(session, input),
          ),
        )
      }
      onToggleRule={(id, enabled) =>
        void act(() =>
          readJson(
            `${PATH}/rules/${id}`,
            (body) => alertRuleListSchema.parse(body),
            post(session, { enabled }),
          ),
        )
      }
      onDeleteRule={(id) =>
        void act(() =>
          readJson(
            `${PATH}/rules/${id}`,
            (body) => alertRuleListSchema.parse(body),
            request(session, "DELETE"),
          ),
        )
      }
      onExport={(auditId, destinationId) =>
        void act(() =>
          readJson(
            `${PATH}/exports`,
            (body) => alertOverviewSchema.parse(body),
            post(session, { auditId, destinationId }),
          ),
        )
      }
      onWeekly={(destinationId) =>
        void act(() =>
          readJson(
            `${PATH}/weekly`,
            (body) => alertOverviewSchema.parse(body),
            post(session, { destinationId }),
          ),
        )
      }
    />
  );
}

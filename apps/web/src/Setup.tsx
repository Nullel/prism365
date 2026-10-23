import {
  type PrismSession,
  type SetupCheck,
  type SetupStatus,
  setupStatusSchema,
} from "@prism365/contracts";
import { useCallback, useEffect, useState } from "react";
import { readJson } from "./common.tsx";

// The first-run checklist on the Microsoft access page: what this installation has set up and
// what it has not, each with its fix (docs/SETUP.md). "Check with Microsoft" adds one small test
// read per granted permission, which the server records in the access ledger first. Hidden in
// demo mode and in the portfolio demo, which have no installation to check.

const stateLabels: Record<SetupCheck["state"], string> = {
  ok: "Done",
  warn: "Needs attention",
  fail: "Not working",
  unknown: "Not known",
};

export function SetupChecklist({ session }: { session: PrismSession }) {
  if (session.mode !== "live" || import.meta.env.VITE_DEMO === "true") return null;
  return <LiveSetupChecklist session={session} />;
}

function LiveSetupChecklist({ session }: { session: PrismSession }) {
  const [status, setStatus] = useState<SetupStatus | null>(null);
  const [error, setError] = useState("");
  const [probing, setProbing] = useState(false);
  const load = useCallback(async (probe: boolean, signal?: AbortSignal) => {
    try {
      setStatus(
        await readJson(
          `/api/v1/setup/status${probe ? "?probe=1" : ""}`,
          (value) => setupStatusSchema.parse(value),
          signal ? { signal } : {},
        ),
      );
      setError("");
    } catch (loadError) {
      if (signal?.aborted) return;
      setError(loadError instanceof Error ? loadError.message : "The checklist could not load.");
    }
  }, []);
  useEffect(() => {
    const controller = new AbortController();
    void load(false, controller.signal);
    return () => controller.abort();
  }, [load]);
  const probe = useCallback(async () => {
    setProbing(true);
    try {
      await load(true);
    } finally {
      setProbing(false);
    }
  }, [load]);
  return (
    <SetupChecklistView
      session={session}
      status={status}
      error={error}
      probing={probing}
      onProbe={() => void probe()}
    />
  );
}

export function SetupChecklistView({
  session,
  status,
  error,
  probing,
  onProbe,
}: {
  session: PrismSession;
  status: SetupStatus | null;
  error: string;
  probing: boolean;
  onProbe: () => void;
}) {
  const checks = status?.checks ?? [];
  const done = checks.filter((check) => check.state === "ok").length;
  const open = checks.length - done;
  return (
    <section className="panel setup-checklist" aria-labelledby="setup-title">
      <div className="panel__header">
        <div>
          <h2 id="setup-title">Setup checklist</h2>
          <p className="setup-checklist__intro">
            {checks.length > 0
              ? `${done} of ${checks.length} done${open > 0 ? `; ${open === 1 ? "1 needs" : `${open} need`} a look` : ""}.`
              : "What this installation has set up, and what is left."}
          </p>
        </div>
        {checks.length > 0 ? (
          <button
            className="button button--secondary"
            type="button"
            onClick={onProbe}
            disabled={probing}
          >
            {probing ? "Checking with Microsoft…" : "Check with Microsoft"}
          </button>
        ) : null}
      </div>
      {error ? (
        <p className="form-error" role="alert">
          {error}
        </p>
      ) : null}
      {!status && !error ? <p className="empty-data">Checking this installation…</p> : null}
      {status && checks.length === 0 ? (
        status.configured ? (
          <p className="setup-checklist__intro">
            {session.authenticated
              ? "Reload the page to see the checklist."
              : "Sign in with an administrator account on this installation's list to see the full checklist."}
          </p>
        ) : (
          <ul className="setup-checklist__items">
            <CheckItem
              check={{
                id: "configuration",
                state: "fail",
                title: "Configuration is incomplete",
                detail: `Missing: ${status.missingConfiguration.join(", ") || "unknown settings"}.`,
                fix: "Run npm run setup, or set these in .env and restart Prism365 (docs/SETUP.md).",
                grantFeature: null,
              }}
            />
          </ul>
        )
      ) : null}
      {checks.length > 0 ? (
        <ul className="setup-checklist__items">
          {checks.map((check) => (
            <CheckItem key={check.id} check={check} />
          ))}
        </ul>
      ) : null}
      {checks.length > 0 ? (
        <p className="setup-checklist__note" role="status">
          {status?.probed
            ? "Checked with Microsoft just now: one small read per granted permission, recorded in the access ledger."
            : "Check with Microsoft makes one small read per granted permission and records it in the access ledger."}
        </p>
      ) : null}
    </section>
  );
}

function CheckItem({ check }: { check: SetupCheck }) {
  return (
    <li className="setup-check">
      <span className={`setup-state setup-state--${check.state}`}>{stateLabels[check.state]}</span>
      <div>
        <strong>{check.title}</strong>
        {check.detail ? <p>{check.detail}</p> : null}
        {check.fix ? <p className="setup-check__fix">{check.fix}</p> : null}
        {check.grantFeature ? (
          <a href={`/api/v1/auth/login?feature=${encodeURIComponent(check.grantFeature)}`}>
            Grant access
          </a>
        ) : null}
      </div>
    </li>
  );
}

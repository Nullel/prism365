import {
  type CanaryAccount,
  type CanaryChange,
  type CanaryList,
  canaryListSchema,
  type PrismSession,
} from "@prism365/contracts";
import { useCallback, useEffect, useId, useState } from "react";
import { day, post, readJson } from "./common.tsx";

// Canary accounts: decoy accounts nobody should ever sign in to. A switch on the person's page
// marks or unmarks one, and a small panel on the sign-in audit page lists them. Any sign-in
// attempt by a canary, successful or not, becomes a critical finding in the ordinary findings list.

const PATH = "/api/v1/settings/canaries";

type Loaded =
  | { state: "loading" }
  | { state: "unavailable" }
  | { state: "ready"; list: CanaryList };

// The list, and the one write that changes it. A list that cannot be read (the demo, or an
// installation without the sign-in module) leaves the controls out rather than showing broken ones.
function useCanaries(session: PrismSession) {
  const [loaded, setLoaded] = useState<Loaded>({ state: "loading" });
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    const controller = new AbortController();
    readJson(PATH, (value) => canaryListSchema.parse(value), { signal: controller.signal }).then(
      (list) => setLoaded({ state: "ready", list }),
      () => {
        if (!controller.signal.aborted) setLoaded({ state: "unavailable" });
      },
    );
    return () => controller.abort();
  }, []);

  const change = useCallback(
    async (update: CanaryChange): Promise<boolean> => {
      setBusy(true);
      setError("");
      try {
        const list = await readJson(
          PATH,
          (value) => canaryListSchema.parse(value),
          post(session, update),
        );
        setLoaded({ state: "ready", list });
        return true;
      } catch (changeError) {
        setError(changeError instanceof Error ? changeError.message : "The request failed.");
        return false;
      } finally {
        setBusy(false);
      }
    },
    [session],
  );

  return { loaded, busy, error, change };
}

export function CanaryToggleView({
  name,
  active,
  confirming,
  busy,
  error,
  onChange,
  onConfirm,
  onCancel,
}: {
  name: string;
  active: boolean;
  confirming: boolean;
  busy: boolean;
  error: string;
  onChange: (on: boolean) => void;
  onConfirm: () => void;
  onCancel: () => void;
}) {
  return (
    <div className="canary-toggle">
      <label className="check-row">
        <input
          type="checkbox"
          checked={active}
          disabled={busy}
          onChange={(event) => onChange(event.target.checked)}
        />
        <span>
          <strong>Canary account</strong>
          <small>
            A decoy account nobody should ever sign in to. Any sign-in attempt on it, successful or
            not, opens a critical finding.
          </small>
        </span>
      </label>
      {confirming ? (
        <fieldset className="canary-confirm">
          <legend className="visually-hidden">Confirm canary account</legend>
          <p>
            Mark {name} as a canary account? Choose this only for an account nobody uses. If a real
            person signs in to it, Prism365 opens a critical finding.
          </p>
          <div className="canary-confirm__actions">
            <button
              className="button button--primary"
              type="button"
              disabled={busy}
              onClick={onConfirm}
            >
              Mark as canary
            </button>
            <button
              className="button button--secondary"
              type="button"
              disabled={busy}
              onClick={onCancel}
            >
              Cancel
            </button>
          </div>
        </fieldset>
      ) : null}
      {error ? (
        <p className="form-error" role="alert">
          {error}
        </p>
      ) : null}
    </div>
  );
}

export function CanaryToggle({
  session,
  userId,
  name,
}: {
  session: PrismSession;
  userId: string;
  name: string;
}) {
  const { loaded, busy, error, change } = useCanaries(session);
  const [confirming, setConfirming] = useState(false);
  if (loaded.state === "unavailable") return null;
  const active = loaded.state === "ready" && loaded.list.accounts.some((a) => a.userId === userId);
  return (
    <CanaryToggleView
      name={name}
      active={active}
      confirming={confirming}
      busy={busy || loaded.state === "loading"}
      error={error}
      // Turning it on waits for a confirmation; turning it off is safe to do at once.
      onChange={(on) => (on ? setConfirming(true) : void change({ remove: [userId] }))}
      onConfirm={() =>
        void change({ add: [userId] }).then((saved) => {
          if (saved) setConfirming(false);
        })
      }
      onCancel={() => setConfirming(false)}
    />
  );
}

export function CanaryPanelView({
  accounts,
  busy,
  error,
  onRemove,
  onOpenPerson,
}: {
  accounts: readonly CanaryAccount[];
  busy: boolean;
  error: string;
  onRemove: (userId: string) => void;
  onOpenPerson?: (userId: string) => void;
}) {
  const titleId = useId();
  return (
    <section className="panel canary-panel" aria-labelledby={titleId}>
      <h3 id={titleId}>Canary accounts</h3>
      {accounts.length === 0 ? (
        <p className="empty-data">
          No canary accounts yet. Turn on Canary account on a person's page to add one.
        </p>
      ) : (
        <>
          <p className="empty-data">
            Any sign-in attempt on these accounts, successful or not, opens a critical finding.
          </p>
          <ul className="canary-list">
            {accounts.map((account) => (
              <li key={account.userId}>
                {onOpenPerson ? (
                  <button
                    className="canary-list__name finding__subject finding__subject--link"
                    type="button"
                    onClick={() => onOpenPerson(account.userId)}
                  >
                    {account.principal}
                  </button>
                ) : (
                  <span className="canary-list__name">{account.principal}</span>
                )}
                <span className="person-muted">Added {day(account.addedAt)}</span>
                <button
                  className="button button--secondary"
                  type="button"
                  disabled={busy}
                  aria-label={`Stop treating ${account.principal} as a canary`}
                  onClick={() => onRemove(account.userId)}
                >
                  Remove
                </button>
              </li>
            ))}
          </ul>
        </>
      )}
      {error ? (
        <p className="form-error" role="alert">
          {error}
        </p>
      ) : null}
    </section>
  );
}

export function CanaryPanel({
  session,
  onOpenPerson,
}: {
  session: PrismSession;
  onOpenPerson?: (userId: string) => void;
}) {
  const { loaded, busy, error, change } = useCanaries(session);
  return (
    <CanaryPanelView
      accounts={loaded.state === "ready" ? loaded.list.accounts : []}
      busy={busy || loaded.state !== "ready"}
      error={loaded.state === "unavailable" ? "The canary list could not be read just now." : error}
      onRemove={(userId) => void change({ remove: [userId] })}
      {...(onOpenPerson ? { onOpenPerson } : {})}
    />
  );
}

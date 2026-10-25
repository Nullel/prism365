import {
  type CanaryAccount,
  type CanaryChange,
  type CanaryList,
  canaryListSchema,
  type PrismSession,
} from "@prism365/contracts";
import { useCallback, useEffect, useId, useState } from "react";
import { day, post, readJson } from "./common.tsx";

// Canary accounts: decoy accounts nobody should ever sign in to. An item in the Actions menu on
// the person's page marks or unmarks one, and a small panel on the sign-in audit page lists them.
// Any sign-in attempt by a canary, successful or not, becomes a critical finding in the ordinary
// findings list.

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

// The one-line explanation of what a canary is. It is the menu item's second line, so the person
// choosing it reads what it does before they choose.
export const canaryExplanation =
  "A decoy account nobody should ever sign in to. Any sign-in attempt on it, successful or not, opens a critical finding.";

// What follows a choice in the Actions menu: the confirmation before an account is marked, and
// any failure from marking or unmarking. Nothing is drawn when there is neither.
export function CanaryStepView({
  name,
  confirming,
  busy,
  error,
  onConfirm,
  onCancel,
}: {
  name: string;
  confirming: boolean;
  busy: boolean;
  error: string;
  onConfirm: () => void;
  onCancel: () => void;
}) {
  if (!confirming && !error) return null;
  return (
    <div className="canary-toggle">
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

export interface PersonCanary {
  // False when the list cannot be read (the demo, or an installation without the sign-in
  // module): the menu leaves the item out rather than showing a broken one.
  available: boolean;
  active: boolean;
  busy: boolean;
  error: string;
  confirming: boolean;
  onMark: () => void;
  onConfirm: () => void;
  onCancel: () => void;
  onUnmark: () => void;
}

// A person's canary state, shared by the status chip, the menu item and the confirmation step.
export function usePersonCanary(session: PrismSession, userId: string): PersonCanary {
  const { loaded, busy, error, change } = useCanaries(session);
  const [confirming, setConfirming] = useState(false);
  return {
    available: loaded.state !== "unavailable",
    active: loaded.state === "ready" && loaded.list.accounts.some((a) => a.userId === userId),
    busy: busy || loaded.state === "loading",
    error,
    confirming,
    // Turning it on waits for a confirmation; turning it off is safe to do at once.
    onMark: () => setConfirming(true),
    onConfirm: () =>
      void change({ add: [userId] }).then((saved) => {
        if (saved) setConfirming(false);
      }),
    onCancel: () => setConfirming(false),
    onUnmark: () => void change({ remove: [userId] }),
  };
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
          No canary accounts yet. Open a person's page and choose Mark as canary in its Actions menu
          to add one.
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

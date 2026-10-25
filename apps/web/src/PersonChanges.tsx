import {
  type PersonChange,
  type PersonChangePage,
  type PrismSession,
  personChangePageSchema,
} from "@prism365/contracts";
import { useCallback, useEffect, useState } from "react";
import { ChangeRow, startHistory } from "./Changes.tsx";
import { readJson, SkeletonRows } from "./common.tsx";

// Directory changes made to or by one person, newest first, on their page. Read from what
// collection stored; opening the panel writes its own access-ledger entry on the server.

const PAGE = 20;

function involvement(change: PersonChange): string {
  if (change.involvement.actor && change.involvement.target) return "Made by and to this person";
  return change.involvement.actor ? "Made by this person" : "Made to this person";
}

export function PersonChangeList({
  page,
  error,
  busy,
  onMore,
}: {
  // null while the first page loads.
  page: PersonChangePage | null;
  error: string;
  busy: boolean;
  onMore: () => void;
}) {
  return (
    <section
      className="panel signin-table person-grid person-changes"
      aria-label="Directory changes"
    >
      <h3>Directory changes</h3>
      {error ? (
        <p className="empty-data">{error}</p>
      ) : !page ? (
        <SkeletonRows rows={4} label="Loading changes" />
      ) : page.items.length === 0 ? (
        <p className="empty-data">
          No directory change involving this person has been collected. {startHistory}
        </p>
      ) : (
        <>
          <ol className="change-list">
            {page.items.map((change) => (
              <ChangeRow
                key={change.id}
                change={change}
                note={<span className="person-muted change-row__note">{involvement(change)}</span>}
              />
            ))}
          </ol>
          {page.next ? (
            <button
              className="button button--secondary person-more"
              type="button"
              disabled={busy}
              onClick={onMore}
            >
              {busy ? "Loading…" : "Show older changes"}
            </button>
          ) : null}
        </>
      )}
    </section>
  );
}

export function PersonChanges({
  session,
  userId,
  onLoaded,
}: {
  session: PrismSession;
  userId: string;
  // Told what has loaded each time the list changes, for the count on the page's Changes tab.
  onLoaded?: (page: PersonChangePage) => void;
}) {
  const path = `/api/v1/people/${encodeURIComponent(userId)}/changes`;
  const [page, setPage] = useState<PersonChangePage | null>(null);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const live = session.mode === "live" && session.authenticated;

  const load = useCallback(
    async (before: string | null, signal?: AbortSignal) => {
      const query = new URLSearchParams({ limit: String(PAGE) });
      if (before) query.set("before", before);
      const next = await readJson(
        `${path}?${query}`,
        (value) => personChangePageSchema.parse(value),
        signal ? { signal } : {},
      );
      setPage((current) =>
        before && current ? { items: [...current.items, ...next.items], next: next.next } : next,
      );
      setError("");
    },
    [path],
  );

  useEffect(() => {
    if (!live) return;
    const controller = new AbortController();
    load(null, controller.signal).catch((loadError: unknown) => {
      if (controller.signal.aborted) return;
      setError(
        loadError instanceof Error ? loadError.message : "Directory changes could not be loaded.",
      );
    });
    return () => controller.abort();
  }, [live, load]);

  useEffect(() => {
    if (page) onLoaded?.(page);
  }, [page, onLoaded]);

  if (!live) return null;
  return (
    <PersonChangeList
      page={page}
      error={error}
      busy={busy}
      onMore={() => {
        if (!page?.next) return;
        setBusy(true);
        load(page.next)
          .catch((loadError: unknown) =>
            setError(
              loadError instanceof Error ? loadError.message : "Older changes could not be loaded.",
            ),
          )
          .finally(() => setBusy(false));
      }}
    />
  );
}

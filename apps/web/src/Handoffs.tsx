import {
  actionLabel,
  type HandoffList,
  type HandoffRecord,
  handoffListSchema,
  MAX_RECEIPT_BYTES,
  type PrismSession,
  type ReceiptImport,
  type ReceiptItemPage,
  type ReceiptStatus,
  type ReceiptSummary,
  receiptImportSchema,
  receiptItemPageSchema,
} from "@prism365/contracts";
import { type ReactNode, useCallback, useEffect, useId, useState } from "react";
import { number, post, RelativeTime, readJson, SkeletonRows, when } from "./common.tsx";

// Handoffs: lists Prism365 exported from audits for the tool that acts on them, and the receipts
// that tool sent back. Prism365 never calls the tool. A receipt is shown as the tool reported it,
// labelled as its report, because Prism365 has not checked it against the tenant.

export const FORMATS_URL = "https://github.com/Nullel/prism365/blob/main/docs/handoff.md";
const ITEM_PAGE = 200;

export type ImportState =
  | { state: "idle" }
  | { state: "reading"; name: string }
  | { state: "done"; name: string; result: ReceiptImport }
  | { state: "failed"; name: string; message: string };

export function statusTagClass(status: ReceiptStatus): string {
  if (status === "done") return "person-tag person-tag--ok";
  if (status === "failed") return "person-tag person-tag--alert";
  return "person-tag";
}

export function importMessage(result: ReceiptImport): string {
  const { receipt } = result;
  if (!result.imported)
    return `This receipt was already imported on ${when(receipt.importedAt)}. Nothing changed.`;
  return `Imported the receipt from ${receipt.tool.name}: ${number.format(receipt.done)} done, ${number.format(receipt.failed)} failed, ${number.format(receipt.skipped)} skipped.`;
}

// A chosen file as the body to send: JSON, and no larger than the server reads.
export async function readReceiptFile(file: {
  size: number;
  text(): Promise<string>;
}): Promise<unknown> {
  if (file.size > MAX_RECEIPT_BYTES)
    throw new Error(
      "A receipt can be at most 4 MB. Ask the tool to split it into several receipts.",
    );
  const text = await file.text();
  try {
    return JSON.parse(text) as unknown;
  } catch {
    throw new Error("This file is not JSON. Choose the receipt file the action tool wrote.");
  }
}

// Saves a handoff through fetch rather than a plain link, so a refusal (a deleted run, an audit
// with no accounts) shows as a message on the page instead of a saved error file.
export async function downloadHandoff(url: string, target: Document = document): Promise<void> {
  const response = await fetch(url, { credentials: "same-origin", cache: "no-store" });
  if (!response.ok) {
    const body = (await response.json().catch(() => null)) as { message?: string } | null;
    throw new Error(body?.message ?? `Request failed with HTTP ${response.status}.`);
  }
  const name =
    /filename="([^"]+)"/.exec(response.headers.get("Content-Disposition") ?? "")?.[1] ??
    "prism365-handoff.json";
  const link = target.createElement("a");
  link.href = URL.createObjectURL(await response.blob());
  link.download = name;
  target.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(link.href), 1_000);
}

function viewText(handoff: HandoffRecord): string {
  const view = handoff.view?.label ?? "Every row";
  return handoff.filtered ? `${view}, filtered` : view;
}

// What a handoff did not cover: the domains its run was limited to, and a partial run.
export function HandoffLimits({ handoff }: { handoff: HandoffRecord }) {
  const scope = handoff.scope ?? [];
  return (
    <>
      {scope.length > 0 ? (
        <span className="person-tag handoff-match" title={scope.join(", ")}>
          Limited to {number.format(scope.length)} {scope.length === 1 ? "domain" : "domains"}
        </span>
      ) : null}
      {handoff.complete ? null : (
        <span
          className="person-tag person-tag--alert handoff-match"
          title="Microsoft returned less than everything to this run, so the list may miss accounts."
        >
          Partial result
        </span>
      )}
    </>
  );
}

// The handoff a receipt answers, by its audit and when it was exported.
export function answers(receipt: ReceiptSummary, handoffs: readonly HandoffRecord[]): ReactNode {
  if (!receipt.handoff) return "No handoff";
  const known = handoffs.find((handoff) => handoff.id === receipt.handoff?.id);
  return (
    <>
      {known?.auditTitle ?? receipt.handoff.auditId}, exported{" "}
      <RelativeTime value={receipt.handoff.generatedAt} />
    </>
  );
}

function HandoffTable({ handoffs }: { handoffs: readonly HandoffRecord[] }) {
  if (handoffs.length === 0)
    return (
      <p className="empty-data">
        Nothing exported yet. Open an audit's results and choose Export for action.
      </p>
    );
  return (
    <table>
      <caption className="visually-hidden">Handoffs exported from audits</caption>
      <thead>
        <tr>
          <th scope="col">Exported</th>
          <th scope="col">Audit</th>
          <th scope="col">View</th>
          <th scope="col" className="num">
            Accounts
          </th>
          <th scope="col">By</th>
        </tr>
      </thead>
      <tbody>
        {handoffs.map((handoff) => (
          <tr key={handoff.id}>
            <th scope="row" className="nowrap">
              <RelativeTime value={handoff.generatedAt} />
            </th>
            <td className="person-text">
              {handoff.auditTitle}
              <HandoffLimits handoff={handoff} />
            </td>
            <td className="person-text">{viewText(handoff)}</td>
            <td className="num">{number.format(handoff.items)}</td>
            <td className="person-text">{handoff.generatedBy.name ?? "Unknown"}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function ReceiptTable({
  list,
  selected,
  onSelect,
}: {
  list: HandoffList;
  selected: string | null;
  onSelect: (id: string | null) => void;
}) {
  if (list.receipts.length === 0)
    return <p className="empty-data">No receipts yet. Import one when the action tool sends it.</p>;
  return (
    <table>
      <caption className="visually-hidden">Receipts sent back by action tools</caption>
      <thead>
        <tr>
          <th scope="col">Tool</th>
          <th scope="col">Performed</th>
          <th scope="col" className="num">
            Done
          </th>
          <th scope="col" className="num">
            Failed
          </th>
          <th scope="col" className="num">
            Skipped
          </th>
          <th scope="col">Answers</th>
          <th scope="col">
            <span className="visually-hidden">Items</span>
          </th>
        </tr>
      </thead>
      <tbody>
        {list.receipts.map((receipt) => (
          <tr key={receipt.id}>
            <th scope="row" className="person-text">
              {receipt.tool.name} <span className="muted-label">{receipt.tool.version}</span>
            </th>
            <td className="nowrap">
              <RelativeTime value={receipt.performedAt} />
              <span className="muted-label"> by {receipt.performedBy}</span>
            </td>
            <td className="num">{number.format(receipt.done)}</td>
            <td className="num">{number.format(receipt.failed)}</td>
            <td className="num">{number.format(receipt.skipped)}</td>
            <td className="person-text">{answers(receipt, list.handoffs)}</td>
            <td>
              <button
                className="link-button"
                type="button"
                aria-expanded={selected === receipt.id}
                onClick={() => onSelect(selected === receipt.id ? null : receipt.id)}
              >
                {selected === receipt.id ? "Hide items" : "Show items"}
              </button>
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

export function ReceiptItemsTable({
  page,
  onMore,
  loading,
}: {
  page: ReceiptItemPage;
  onMore: (() => void) | null;
  loading: boolean;
}) {
  const { receipt } = page;
  return (
    <>
      <p className="audit-note">
        As reported by {receipt.tool.name} {receipt.tool.version}, run by {receipt.performedBy} on{" "}
        {when(receipt.performedAt)}. Prism365 has not checked these against the tenant.
      </p>
      <table>
        <caption className="visually-hidden">Actions this receipt reports</caption>
        <thead>
          <tr>
            <th scope="col">Account</th>
            <th scope="col">Action</th>
            <th scope="col">Status</th>
            <th scope="col">Detail</th>
          </tr>
        </thead>
        <tbody>
          {page.items.map((item) => (
            <tr key={item.position}>
              <th scope="row" className="person-text">
                {item.userPrincipalName ?? item.userId}
                {item.userPrincipalName && item.userId ? (
                  <span className="muted-label"> {item.userId}</span>
                ) : null}
              </th>
              <td className="person-text">{actionLabel(item.action)}</td>
              <td className="person-text">
                <span className={statusTagClass(item.status)}>{item.status}</span>
              </td>
              <td className="person-text">{item.detail || "None"}</td>
            </tr>
          ))}
        </tbody>
      </table>
      {onMore ? (
        <div className="audit-more-rows">
          <button
            className="button button--secondary"
            type="button"
            disabled={loading}
            onClick={onMore}
          >
            {loading ? "Loading…" : "Show more"}
          </button>
          <span className="muted-label">
            Showing {number.format(page.items.length)} of {number.format(receipt.items)}
          </span>
        </div>
      ) : null}
    </>
  );
}

function ReceiptDetail({ id }: { id: string }) {
  const [page, setPage] = useState<ReceiptItemPage | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const load = useCallback(
    async (after: number | null) => {
      setLoading(true);
      try {
        const query = new URLSearchParams({ limit: String(ITEM_PAGE) });
        if (after !== null) query.set("after", String(after));
        const next = await readJson(
          `/api/v1/receipts/${encodeURIComponent(id)}?${query}`,
          receiptItemPageSchema.parse,
        );
        setPage((current) =>
          after !== null && current ? { ...next, items: [...current.items, ...next.items] } : next,
        );
        setError("");
      } catch (problem) {
        setError(problem instanceof Error ? problem.message : "The receipt could not be loaded.");
      } finally {
        setLoading(false);
      }
    },
    [id],
  );

  useEffect(() => {
    void load(null);
  }, [load]);

  if (error)
    return (
      <p className="form-error" role="alert">
        {error}
      </p>
    );
  if (!page) return <SkeletonRows rows={4} label="Loading the receipt" />;
  const next = page.next;
  return (
    <ReceiptItemsTable
      page={page}
      loading={loading}
      onMore={next === null ? null : () => void load(next)}
    />
  );
}

export function HandoffsView({
  list,
  error,
  importState,
  selected,
  onImport,
  onSelect,
  detail,
}: {
  list: HandoffList | null;
  error: string;
  importState: ImportState;
  selected: string | null;
  onImport: (file: File) => void;
  onSelect: (id: string | null) => void;
  // The selected receipt's items; the page passes its loader.
  detail?: ReactNode;
}) {
  const uploadId = useId();
  const busy = importState.state === "reading";
  return (
    <div className="audits handoffs">
      <section className="panel handoff-intro" aria-label="How handoffs work">
        <p>
          Prism365 only reads. When an audit finds accounts that need something done, Export for
          action on its results saves them as a handoff: a file for the tool that acts on them, such
          as an offboarding tool. That tool sends back a receipt saying what it did to whom, and you
          import the receipt here.
        </p>
        <p>
          Receipts are reports from that tool. Prism365 keeps and shows them as reported, and does
          not check them against the tenant. Both file formats are described in{" "}
          <a href={FORMATS_URL} target="_blank" rel="noreferrer">
            docs/handoff.md
          </a>
          .
        </p>
      </section>

      {error ? (
        <p className="form-error form-error--banner" role="alert">
          {error}
        </p>
      ) : null}

      <section className="panel handoff-import" aria-label="Import a receipt">
        <label className="field" htmlFor={uploadId}>
          <span>Import receipt (.json, at most 4 MB)</span>
          <input
            id={uploadId}
            type="file"
            accept=".json,application/json"
            disabled={busy}
            onChange={(event) => {
              const file = event.target.files?.[0];
              // Cleared, so choosing the same file again imports it again.
              event.target.value = "";
              if (file) onImport(file);
            }}
          />
        </label>
        <div aria-live="polite">
          {importState.state === "reading" ? (
            <p className="muted-label">Importing {importState.name}…</p>
          ) : importState.state === "done" ? (
            <p className="handoff-result">{importMessage(importState.result)}</p>
          ) : importState.state === "failed" ? (
            <p className="form-error" role="alert">
              {importState.name}: {importState.message}
            </p>
          ) : null}
        </div>
      </section>

      <section className="panel signin-table person-grid" aria-label="Exported handoffs">
        <h3>Exported</h3>
        {list ? (
          <HandoffTable handoffs={list.handoffs} />
        ) : error ? null : (
          <SkeletonRows rows={3} label="Loading handoffs" />
        )}
      </section>

      <section className="panel signin-table person-grid" aria-label="Receipts">
        <h3>Receipts</h3>
        {list ? (
          <ReceiptTable list={list} selected={selected} onSelect={onSelect} />
        ) : error ? null : (
          <SkeletonRows rows={3} label="Loading receipts" />
        )}
      </section>

      {selected && detail ? (
        <section className="panel signin-table person-grid" aria-label="Receipt items">
          <h3>Receipt items</h3>
          {detail}
        </section>
      ) : null}
    </div>
  );
}

export function Handoffs({ session }: { session: PrismSession }) {
  const [list, setList] = useState<HandoffList | null>(null);
  const [error, setError] = useState("");
  const [importState, setImportState] = useState<ImportState>({ state: "idle" });
  const [selected, setSelected] = useState<string | null>(null);
  const live = session.mode === "live";

  const load = useCallback(async () => {
    try {
      setList(await readJson("/api/v1/handoffs", handoffListSchema.parse));
      setError("");
    } catch (problem) {
      setError(problem instanceof Error ? problem.message : "Handoffs could not be loaded.");
    }
  }, []);

  useEffect(() => {
    if (!live || !session.authenticated) return;
    void load();
  }, [live, session.authenticated, load]);

  async function importFile(file: File) {
    setImportState({ state: "reading", name: file.name });
    try {
      const body = await readReceiptFile(file);
      const result = await readJson(
        "/api/v1/receipts",
        receiptImportSchema.parse,
        post(session, body),
      );
      setImportState({ state: "done", name: file.name, result });
      await load();
    } catch (problem) {
      setImportState({
        state: "failed",
        name: file.name,
        message: problem instanceof Error ? problem.message : "The receipt could not be imported.",
      });
    }
  }

  if (!live)
    return <p className="empty-data">Handoffs need a connected tenant. Demo mode has none.</p>;
  if (!session.authenticated)
    return <p className="empty-data">Sign in to export handoffs and import receipts.</p>;
  return (
    <HandoffsView
      list={list}
      error={error}
      importState={importState}
      selected={selected}
      onImport={(file) => void importFile(file)}
      onSelect={setSelected}
      detail={selected ? <ReceiptDetail key={selected} id={selected} /> : null}
    />
  );
}

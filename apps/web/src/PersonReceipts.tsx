import {
  actionLabel,
  type PersonReceiptList,
  type PrismSession,
  personReceiptListSchema,
} from "@prism365/contracts";
import { useEffect, useId, useState } from "react";
import { RelativeTime, readJson } from "./common.tsx";
import { statusTagClass } from "./Handoffs.tsx";

// What other tools reported doing to this person, from the receipts imported on the Handoffs
// page. These are the tools' own reports: Prism365 shows them as sent and never as something it
// checked. Nothing is shown when no receipt names the person.

export function PersonReceiptsPanel({
  list,
  error,
}: {
  list: PersonReceiptList | null;
  error: string;
}) {
  const titleId = useId();
  if (!error && (!list || list.items.length === 0)) return null;
  return (
    <section className="panel signin-table person-grid" aria-labelledby={titleId}>
      <h3 id={titleId}>Reported by other tools</h3>
      {error ? (
        <p className="form-error" role="alert">
          {error}
        </p>
      ) : (
        <>
          <p className="audit-note">
            These are reports from the tools that acted, as they sent them. Prism365 has not checked
            them against the tenant.
          </p>
          <table>
            <caption className="visually-hidden">
              Actions other tools reported for this person
            </caption>
            <thead>
              <tr>
                <th scope="col">Action</th>
                <th scope="col">Status</th>
                <th scope="col">Tool</th>
                <th scope="col">When</th>
                <th scope="col">Detail</th>
              </tr>
            </thead>
            <tbody>
              {(list?.items ?? []).map((item) => (
                <tr key={`${item.receiptId}|${item.position}`}>
                  <th scope="row" className="person-text">
                    {actionLabel(item.action)}
                    {item.matchedBy === "userPrincipalName" ? (
                      <span
                        className="person-tag handoff-match"
                        title={`The tool named only ${item.userPrincipalName ?? "a sign-in name"}, not the account's id.`}
                      >
                        Matched by sign-in name
                      </span>
                    ) : null}
                  </th>
                  <td className="person-text">
                    <span className={statusTagClass(item.status)}>{item.status}</span>
                  </td>
                  <td className="person-text">
                    {item.tool.name} <span className="muted-label">{item.tool.version}</span>
                    <span className="muted-label"> run by {item.performedBy}</span>
                  </td>
                  <td className="nowrap">
                    <RelativeTime value={item.performedAt} />
                  </td>
                  <td className="person-text">{item.detail || "None"}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {list?.truncated ? <p className="muted-label">Showing the newest 100.</p> : null}
        </>
      )}
    </section>
  );
}

export function PersonReceipts({
  session,
  userId,
  principal = null,
}: {
  session: PrismSession;
  userId: string;
  // The principal name the page shows, for receipts that named only that.
  principal?: string | null;
}) {
  const [list, setList] = useState<PersonReceiptList | null>(null);
  const [error, setError] = useState("");
  const live = session.mode === "live" && session.authenticated;

  useEffect(() => {
    if (!live) return;
    const controller = new AbortController();
    const query = principal ? `?${new URLSearchParams({ principal })}` : "";
    readJson(
      `/api/v1/people/${encodeURIComponent(userId)}/receipts${query}`,
      personReceiptListSchema.parse,
      { signal: controller.signal },
    ).then(
      (value) => {
        setList(value);
        setError("");
      },
      (problem: unknown) => {
        if (controller.signal.aborted) return;
        setError(
          problem instanceof Error
            ? problem.message
            : "Reports from other tools could not be loaded.",
        );
      },
    );
    return () => controller.abort();
  }, [live, userId, principal]);

  if (!live) return null;
  return <PersonReceiptsPanel list={list} error={error} />;
}

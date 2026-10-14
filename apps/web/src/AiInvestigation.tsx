import {
  aiGrantSummarySchema,
  aiStatusSchema,
  governedAiAnalysisSchema,
  type PrismSession,
  type UserInvestigation,
} from "@prism365/contracts";
import { useEffect, useId, useRef, useState } from "react";
import { LockIcon, SparkleIcon } from "./icons.tsx";

type AiStatus = ReturnType<typeof aiStatusSchema.parse>;
type AiGrant = ReturnType<typeof aiGrantSummarySchema.parse>;
type AiAnalysis = ReturnType<typeof governedAiAnalysisSchema.parse>;
type InvestigationDataset = "directory" | "licenses" | "signIns";

const datasetLabels: Record<InvestigationDataset, string> = {
  directory: "Directory profile",
  licenses: "Assigned licenses",
  signIns: "Recent sign-ins",
};

export function availableAiDatasets(investigation: UserInvestigation): InvestigationDataset[] {
  const datasets: InvestigationDataset[] = [];
  if (investigation.profile.state === "available" && investigation.profile.data)
    datasets.push("directory");
  if (investigation.licenses.state === "available") datasets.push("licenses");
  if (investigation.signIns.state === "available") datasets.push("signIns");
  return datasets;
}

async function responseBody(response: Response): Promise<unknown> {
  const body: unknown = await response.json().catch(() => null);
  if (!response.ok) {
    const message =
      typeof body === "object" &&
      body !== null &&
      "message" in body &&
      typeof body.message === "string"
        ? body.message
        : `The request could not complete (HTTP ${response.status}).`;
    throw new Error(message);
  }
  return body;
}

export function AiInvestigation({
  investigation,
  session,
}: {
  investigation: UserInvestigation;
  session: PrismSession;
}) {
  const questionId = useId();
  const [status, setStatus] = useState<AiStatus | null>(null);
  const [selected, setSelected] = useState<InvestigationDataset[]>([]);
  const [grant, setGrant] = useState<AiGrant | null>(null);
  const [question, setQuestion] = useState("");
  const [analysis, setAnalysis] = useState<AiAnalysis | null>(null);
  const [busy, setBusy] = useState<"grant" | "revoke" | "analysis" | null>(null);
  const [error, setError] = useState("");
  const [now, setNow] = useState(() => Date.now());
  const grantId = useRef<string | null>(null);
  const mounted = useRef(false);
  const pendingAnalysis = useRef<AbortController | null>(null);
  const sources = availableAiDatasets(investigation);

  useEffect(() => {
    mounted.current = true;
    const controller = new AbortController();
    fetch("/api/v1/ai/status", {
      signal: controller.signal,
      cache: "no-store",
      credentials: "same-origin",
    })
      .then(responseBody)
      .then((body) => aiStatusSchema.parse(body))
      .then(setStatus)
      .catch((caught: unknown) => {
        if (controller.signal.aborted) return;
        setError(caught instanceof Error ? caught.message : "Unable to check the AI provider.");
      });
    const timer = window.setInterval(() => setNow(Date.now()), 30_000);
    return () => {
      mounted.current = false;
      controller.abort();
      pendingAnalysis.current?.abort();
      window.clearInterval(timer);
      const activeGrant = grantId.current;
      if (activeGrant && session.csrfToken) {
        void fetch(`/api/v1/ai/grants/${encodeURIComponent(activeGrant)}`, {
          method: "DELETE",
          credentials: "same-origin",
          keepalive: true,
          headers: { "x-csrf-token": session.csrfToken },
        }).catch(() => {});
      }
    };
  }, [session.csrfToken]);

  async function createGrant() {
    if (!session.csrfToken || selected.length === 0) return;
    setBusy("grant");
    setError("");
    try {
      const response = await fetch("/api/v1/ai/grants", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json", "x-csrf-token": session.csrfToken },
        body: JSON.stringify({ datasets: selected }),
      });
      const result = aiGrantSummarySchema.parse(await responseBody(response));
      if (!mounted.current) {
        void fetch(`/api/v1/ai/grants/${encodeURIComponent(result.grantId)}`, {
          method: "DELETE",
          credentials: "same-origin",
          keepalive: true,
          headers: { "x-csrf-token": session.csrfToken },
        }).catch(() => {});
        return;
      }
      grantId.current = result.grantId;
      setGrant(result);
      setNow(Date.now());
    } catch (caught) {
      if (mounted.current)
        setError(caught instanceof Error ? caught.message : "Unable to grant AI access.");
    } finally {
      if (mounted.current) setBusy(null);
    }
  }

  async function revokeGrant() {
    if (!grant || !session.csrfToken) return;
    pendingAnalysis.current?.abort();
    setBusy("revoke");
    setError("");
    try {
      const response = await fetch(`/api/v1/ai/grants/${encodeURIComponent(grant.grantId)}`, {
        method: "DELETE",
        credentials: "same-origin",
        headers: { "x-csrf-token": session.csrfToken },
      });
      if (!response.ok)
        throw new Error(
          "Access could not be revoked. Try again; the grant also expires automatically.",
        );
      grantId.current = null;
      setGrant(null);
      setAnalysis(null);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Unable to revoke AI access.");
    } finally {
      setBusy(null);
    }
  }

  async function askQuestion(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!grant || !session.csrfToken) return;
    pendingAnalysis.current?.abort();
    const controller = new AbortController();
    pendingAnalysis.current = controller;
    setBusy("analysis");
    setAnalysis(null);
    setError("");
    try {
      const response = await fetch("/api/v1/ai/investigations", {
        method: "POST",
        credentials: "same-origin",
        signal: controller.signal,
        headers: { "Content-Type": "application/json", "x-csrf-token": session.csrfToken },
        body: JSON.stringify({
          email: investigation.targetEmail,
          question: question.trim(),
          grantId: grant.grantId,
        }),
      });
      const result = governedAiAnalysisSchema.parse(await responseBody(response));
      if (result.providerDeployment !== "local")
        throw new Error(
          "The provider deployment changed. Review the installation’s AI configuration before continuing.",
        );
      if (!controller.signal.aborted && mounted.current && grantId.current === grant.grantId)
        setAnalysis(result);
    } catch (caught) {
      if (!controller.signal.aborted && mounted.current)
        setError(caught instanceof Error ? caught.message : "AI analysis could not complete.");
    } finally {
      if (!controller.signal.aborted && mounted.current) setBusy(null);
    }
  }

  const expired = grant ? new Date(grant.expiresAt).getTime() <= now : false;
  const canGrant =
    status?.available && session.mode === "live" && session.authenticated && session.csrfToken;

  return (
    <section className="ai-investigation" aria-labelledby="ai-investigation-title">
      <div className="ai-investigation__header">
        <div>
          <span className="eyebrow">Optional AI analysis</span>
          <h3 id="ai-investigation-title">Ask about the evidence you choose.</h3>
        </div>
        <SparkleIcon />
      </div>
      {!status ? (
        <p className="source-notice">
          {error ? "Provider status could not be verified." : "Checking the local AI provider…"}
        </p>
      ) : !status.available ? (
        <div className="source-state">
          <LockIcon />
          <div>
            <strong>AI setup is required</strong>
            <p>
              {status.reason ??
                "The operator must configure a local model and durable access records before AI analysis can run."}
            </p>
            <p>
              The first provider runs on the installation’s local model service. Microsoft tokens
              and imported scripts are never sent to the model.
            </p>
          </div>
        </div>
      ) : (
        <>
          <p className="source-notice">
            Provider: {status.provider ?? "Local model"}. Processing stays with the local provider
            configured for this installation. Direct identifiers and raw records are excluded.
          </p>
          <fieldset className="ai-datasets" disabled={!!grant || !!busy || !canGrant}>
            <legend>Choose the sources this model may receive</legend>
            {(["directory", "licenses", "signIns"] as const).map((dataset) => {
              const available = sources.includes(dataset) && status.datasets.includes(dataset);
              return (
                <label key={dataset}>
                  <input
                    type="checkbox"
                    checked={selected.includes(dataset)}
                    disabled={!available}
                    onChange={(event) =>
                      setSelected((current) =>
                        event.target.checked
                          ? [...current, dataset]
                          : current.filter((item) => item !== dataset),
                      )
                    }
                  />
                  <span>
                    {datasetLabels[dataset]}
                    {!available ? <small>Source is unavailable in this investigation</small> : null}
                  </span>
                </label>
              );
            })}
          </fieldset>
          {grant ? (
            <div className="ai-grant-status" role="status">
              <div>
                <strong>
                  {expired ? "AI access has expired." : "Selected-source access is active."}
                </strong>
                <span>
                  {grant.datasets
                    .map((dataset) => datasetLabels[dataset as InvestigationDataset] ?? dataset)
                    .join(", ")}
                </span>
                <span>
                  Expires{" "}
                  {new Intl.DateTimeFormat(undefined, {
                    dateStyle: "medium",
                    timeStyle: "short",
                  }).format(new Date(grant.expiresAt))}
                </span>
              </div>
              <button
                className="button button--secondary"
                type="button"
                onClick={revokeGrant}
                disabled={busy === "revoke"}
              >
                {busy === "revoke" ? "Revoking…" : "Revoke access"}
              </button>
            </div>
          ) : (
            <button
              className="button button--secondary"
              type="button"
              disabled={!canGrant || selected.length === 0 || !!busy}
              onClick={createGrant}
            >
              {busy === "grant" ? "Recording access…" : "Allow these sources for one hour"}
            </button>
          )}
          <form className="ai-question" onSubmit={askQuestion}>
            <label htmlFor={questionId}>Question about this account’s evidence</label>
            <textarea
              id={questionId}
              value={question}
              onChange={(event) => {
                setQuestion(event.target.value);
                setAnalysis(null);
              }}
              required
              minLength={10}
              maxLength={1000}
              disabled={!grant || expired || !!busy}
              placeholder="Which failed sign-in patterns warrant investigation, and what evidence is missing?"
            />
            <button
              className="button button--primary"
              type="submit"
              disabled={!grant || expired || !!busy}
            >
              <SparkleIcon />
              {busy === "analysis" ? "Analyzing selected evidence…" : "Analyze selected evidence"}
            </button>
          </form>
          <p className="source-notice">
            The grant expires automatically. You can revoke it immediately; closing this
            investigation also requests revocation. AI can explain evidence; it cannot run scripts
            or make tenant changes.
          </p>
        </>
      )}
      {error ? (
        <p className="form-error investigation-error" role="alert">
          {error}
        </p>
      ) : null}
      <div aria-live="polite" aria-busy={busy === "analysis"}>
        {analysis ? (
          <article className="ai-answer">
            <div className="ai-answer__meta">
              <span>AI-generated interpretation</span>
              <span>
                {analysis.evidenceCount} evidence records · {analysis.answer.model}
              </span>
            </div>
            <p className="ai-answer__summary">{analysis.answer.summary}</p>
            <h4>Evidence citations</h4>
            <ul className="ai-citations">
              {analysis.answer.citations.map((citation) => {
                const reference = analysis.citationReferences.find(
                  (item) => item.citationId === citation,
                );
                return (
                  <li key={citation}>
                    <code>{citation}</code>
                    <span>{reference?.evidenceId ?? "Evidence reference unavailable"}</span>
                  </li>
                );
              })}
            </ul>
            <h4>Limits of this interpretation</h4>
            <ul>
              {analysis.answer.limitations.map((item) => (
                <li key={item}>{item}</li>
              ))}
            </ul>
            <p>
              Generated{" "}
              {new Intl.DateTimeFormat(undefined, {
                dateStyle: "medium",
                timeStyle: "short",
              }).format(new Date(analysis.generatedAt))}
              . Review the cited evidence before acting.
            </p>
          </article>
        ) : null}
      </div>
    </section>
  );
}

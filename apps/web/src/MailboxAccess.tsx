import {
  isStaleExport,
  MAILBOX_ACCESS_STALE_DAYS,
  MAX_MAILBOX_ACCESS_BYTES,
  type MailboxAccessImport,
  type MailboxAccessImportResult,
  type MailboxAccessStatus,
  mailboxAccessClearSchema,
  mailboxAccessImportResultSchema,
  mailboxAccessStatusSchema,
  type PrismSession,
} from "@prism365/contracts";
import { useCallback, useEffect, useState } from "react";
import {
  number,
  post,
  RelativeTime,
  readJson,
  SeverityDot,
  SkeletonRows,
  when,
} from "./common.tsx";

// Mailbox access on the Microsoft access page (ADR-0012). Microsoft has no read-only API for Full
// Access, Send As or Send on Behalf, so an administrator runs a read-only script in their own
// Exchange Online session and imports the file here. This section explains that, offers the
// script, takes the file and shows what is stored. Prism365 never calls Exchange.

// Served with the web app (apps/web/public), so a self-hosted install offers its own copy.
export const SCRIPT_URL = "scripts/Export-MailboxAccess.ps1";
export const SECTION_ID = "mailbox-access";
// The SHA-256 of the script this version of Prism365 ships, also published in docs/SETUP.md. A
// test fails when the file, this value and the published one disagree.
export const SCRIPT_SHA256 = "ffc7745a13cf2a6d22a4515a01013270bd3c5c666c39c65da02b13c830c3f2c4";
// Where an administrator checks the hash and reads the script, away from this server: the
// release tag of the version they run, so an older install compares against its own hash.
export const RELEASE_TAGS_URL = "https://github.com/Nullel/prism365/tags";
// This build's version, from the root package.json (apps/web/vite.config.ts). null when the page
// was built some other way, as in tests.
export const APP_VERSION: string | null =
  typeof import.meta.env.VITE_PRISM365_VERSION === "string"
    ? import.meta.env.VITE_PRISM365_VERSION
    : null;

export const COMMANDS = [
  "Get-FileHash .\\Export-MailboxAccess.ps1 -Algorithm SHA256",
  "Install-Module ExchangeOnlineManagement -Scope CurrentUser",
  "Connect-ExchangeOnline",
  "Get-ManagementRole -Cmdlet Get-MailboxPermission",
  ".\\Export-MailboxAccess.ps1 -OutputPath .\\mailbox-access.json",
];

// The SHA-256 of the copy this server offers, as the browser worked it out.
export type ScriptHashState =
  | { state: "reading" }
  | { state: "ready"; sha256: string }
  | { state: "failed" };

// The SHA-256 of some bytes, in lower-case hex, with the browser's own Web Crypto.
export async function sha256Hex(bytes: ArrayBuffer, subtle: SubtleCrypto): Promise<string> {
  const digest = new Uint8Array(await subtle.digest("SHA-256", bytes));
  return Array.from(digest, (byte) => byte.toString(16).padStart(2, "0")).join("");
}

// What the page says about the copy it offers: its hash, and whether that is the hash this
// version shipped. The page can be wrong if the server itself was changed, so it always points to
// the published hash as the check to rely on.
function ScriptHash({ hash, version }: { hash: ScriptHashState; version: string | null }) {
  return (
    <div className="mailbox-access__hash">
      {hash.state === "reading" ? (
        <p className="mailbox-access__line">Working out the SHA-256 of this copy.</p>
      ) : hash.state === "failed" ? (
        <p className="mailbox-access__line">
          This browser could not work out the SHA-256 of this copy. Check the file with Get-FileHash
          as below.
        </p>
      ) : hash.sha256 === SCRIPT_SHA256 ? (
        <p className="mailbox-access__line">
          SHA-256 of this copy: <code>{hash.sha256}</code>. It matches the script this version of
          Prism365 shipped with.
        </p>
      ) : (
        <p className="mailbox-access__line" role="alert">
          <SeverityDot severity="high" label="Changed" /> SHA-256 of this copy:{" "}
          <code>{hash.sha256}</code>. It does not match the script this version of Prism365 shipped
          with (<code>{SCRIPT_SHA256}</code>). Do not run it; ask whoever runs this server what
          changed.
        </p>
      )}
      <p className="mailbox-access__line">
        The script runs with your Exchange session, so read it before you run it; it is short. Then
        run Get-FileHash on the file you downloaded and compare the result with the hash in
        docs/SETUP.md at the release tag of{" "}
        {version
          ? `the version this server runs, Prism365 ${version}`
          : "the Prism365 version you run"}
        , which you open from the{" "}
        <a href={RELEASE_TAGS_URL} target="_blank" rel="noreferrer">
          release tags on GitHub
        </a>
        . The script as released is in the same tag. If this server itself was changed, this page
        could show anything, its version included, so that comparison is the one to rely on.
      </p>
    </div>
  );
}

export type ExportImportState =
  | { state: "idle" }
  | { state: "reading"; name: string }
  | { state: "done"; name: string; result: MailboxAccessImportResult }
  | { state: "failed"; name: string; message: string };

// A chosen file as the body to send: JSON, and no larger than the server reads. Windows
// PowerShell 5.1 starts a UTF-8 file with a byte order mark, which JSON does not allow.
export async function readExportFile(file: {
  size: number;
  text(): Promise<string>;
}): Promise<unknown> {
  if (file.size > MAX_MAILBOX_ACCESS_BYTES)
    throw new Error(
      "An export can be at most 25 MB. Export shared mailboxes only, without -IncludeUserMailboxes.",
    );
  const text = (await file.text()).replace(/^\uFEFF/, "");
  try {
    return JSON.parse(text) as unknown;
  } catch {
    throw new Error("This file is not JSON. Choose the file Export-MailboxAccess.ps1 wrote.");
  }
}

export function importMessage(result: MailboxAccessImportResult): string {
  const stored = result.import;
  return `Imported the export of ${when(stored.exportedAt)}: ${number.format(stored.mailboxes)} ${stored.mailboxes === 1 ? "mailbox" : "mailboxes"} and ${number.format(stored.grants)} ${stored.grants === 1 ? "grant" : "grants"}.${result.replaced ? " It replaces the previous import." : ""}`;
}

function CurrentImport({ stored }: { stored: MailboxAccessImport }) {
  const stale = isStaleExport(stored.exportedAt);
  return (
    <div className="mailbox-access__current">
      <p className="mailbox-access__line">
        {stale ? (
          <>
            <SeverityDot
              severity="medium"
              label={`Older than ${MAILBOX_ACCESS_STALE_DAYS} days`}
            />{" "}
          </>
        ) : null}
        From the export of {when(stored.exportedAt)}, imported{" "}
        <RelativeTime value={stored.importedAt} />
        {stored.importedBy.name ? ` by ${stored.importedBy.name}` : ""}.
      </p>
      <p className="mailbox-access__line">
        {number.format(stored.mailboxes)} {stored.mailboxes === 1 ? "mailbox" : "mailboxes"}
        {stored.includeUserMailboxes ? ", user mailboxes included" : ", shared mailboxes only"},{" "}
        {number.format(stored.grants)} {stored.grants === 1 ? "grant" : "grants"}.
        {stored.unresolved > 0
          ? ` ${number.format(stored.unresolved)} ${stored.unresolved === 1 ? "grant names a trustee" : "grants name trustees"} the script could not resolve, such as a deleted account. Access through ${stored.unresolved === 1 ? "it" : "them"} is not shown on person pages.`
          : ""}
      </p>
      {stale ? (
        <p className="signin-coverage">
          Access may have changed since. Run the script again for a current export.
        </p>
      ) : null}
    </div>
  );
}

// The section as a function of what has loaded. It holds no state, so a test can draw it and
// press its buttons with an answer in hand.
export function MailboxAccessView({
  status,
  error,
  importState,
  confirming,
  clearing,
  readOnly,
  scriptHash,
  version = null,
  onImport,
  onClear,
  onConfirmClear,
  onCancelClear,
}: {
  // null while it loads.
  status: MailboxAccessStatus | null;
  error: string;
  importState: ExportImportState;
  confirming: boolean;
  clearing: boolean;
  // The public demo has no server to keep a file.
  readOnly: boolean;
  scriptHash: ScriptHashState;
  // The installed version, to find its release tag; null when the build does not say.
  version?: string | null;
  onImport: (file: File) => void;
  onClear: () => void;
  onConfirmClear: () => void;
  onCancelClear: () => void;
}) {
  const busy = importState.state === "reading" || clearing;
  const stored = status?.import ?? null;
  return (
    <section
      className="panel mailbox-access"
      id={SECTION_ID}
      aria-labelledby="mailbox-access-title"
    >
      <h2 id="mailbox-access-title">Mailbox access</h2>
      <p className="mailbox-access__intro">
        Microsoft offers no read-only way for Prism365 to see who has Full Access, Send As or Send
        on Behalf on a mailbox. An administrator can export them with a read-only script in their
        own Exchange Online session and import the file here, and person pages then show the shared
        mailboxes each account can use. Prism365 never holds an Exchange credential and never calls
        Exchange.
      </p>
      <p className="mailbox-access__intro">
        The page shows what the file says, with its date. Prism365 cannot check the file against
        Exchange, so import only files the script wrote, and keep them where only administrators can
        change them.
      </p>

      <h3>Run the export</h3>
      <p className="mailbox-access__intro">
        <a href={SCRIPT_URL} download="Export-MailboxAccess.ps1">
          Download Export-MailboxAccess.ps1
        </a>
        . It uses only Get- cmdlets and needs ExchangeOnlineManagement 3.0.0 or later. Install the
        module if you do not have it, connect to this tenant, check that your role may read mailbox
        permissions, then run the script in the same window.
      </p>
      <ScriptHash hash={scriptHash} version={version} />
      <pre className="mailbox-access__commands">
        <code>{COMMANDS.join("\n")}</code>
      </pre>
      <p className="signin-coverage">
        Global Reader includes View-Only Organization Management in Exchange Online. If the role
        check lists none of your roles, ask an Exchange administrator for one that does.
      </p>

      <h3>Import</h3>
      {readOnly ? (
        <p className="empty-data">
          The public demo cannot keep a file. Import it on your own installation.
        </p>
      ) : (
        <label className="field" htmlFor="mailbox-access-file">
          <span>Import export (.json, at most 25 MB)</span>
          <input
            id="mailbox-access-file"
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
      )}
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

      <h3>Current import</h3>
      {error ? (
        <p className="form-error" role="alert">
          {error}
        </p>
      ) : !status ? (
        <SkeletonRows rows={2} label="Loading the current import" />
      ) : !stored ? (
        <p className="empty-data">
          Nothing is imported. Person pages say so in their Shared mailboxes section.
        </p>
      ) : (
        <>
          <CurrentImport stored={stored} />
          {readOnly ? null : confirming ? (
            <div className="mailbox-access__confirm">
              <p>
                Clear the import? Person pages show no shared mailbox access until another export is
                imported.
              </p>
              <button
                className="button button--primary"
                type="button"
                onClick={onConfirmClear}
                disabled={busy}
              >
                {clearing ? "Clearing…" : "Clear the import"}
              </button>
              <button
                className="button button--secondary"
                type="button"
                onClick={onCancelClear}
                disabled={clearing}
              >
                Keep it
              </button>
            </div>
          ) : (
            <button
              className="button button--secondary"
              type="button"
              onClick={onClear}
              disabled={busy}
            >
              Clear
            </button>
          )}
        </>
      )}
    </section>
  );
}

const PATH = "/api/v1/mailbox-access";

export function MailboxAccess({ session }: { session: PrismSession }) {
  const [status, setStatus] = useState<MailboxAccessStatus | null>(null);
  const [error, setError] = useState("");
  const [importState, setImportState] = useState<ExportImportState>({ state: "idle" });
  const [confirming, setConfirming] = useState(false);
  const [clearing, setClearing] = useState(false);
  const [scriptHash, setScriptHash] = useState<ScriptHashState>({ state: "reading" });
  const live = session.mode === "live" && session.authenticated;
  // The public demo replays a live-mode server, so its session says "live"; the build flag is
  // what marks it.
  const readOnly = import.meta.env.VITE_DEMO === "true";

  const load = useCallback(async () => {
    try {
      setStatus(await readJson(PATH, mailboxAccessStatusSchema.parse));
      setError("");
    } catch (problem) {
      setError(problem instanceof Error ? problem.message : "The import could not be loaded.");
    }
  }, []);

  useEffect(() => {
    if (live) void load();
  }, [live, load]);

  // The hash of the very file the download link serves, worked out here rather than taken from
  // the server's word. Web Crypto exists only on HTTPS or localhost.
  useEffect(() => {
    if (!live) return;
    const controller = new AbortController();
    const subtle = globalThis.crypto?.subtle;
    if (!subtle) {
      setScriptHash({ state: "failed" });
      return;
    }
    fetch(SCRIPT_URL, { credentials: "same-origin", cache: "no-store", signal: controller.signal })
      .then(async (response) => {
        if (!response.ok) throw new Error("The script could not be read.");
        setScriptHash({
          state: "ready",
          sha256: await sha256Hex(await response.arrayBuffer(), subtle),
        });
      })
      .catch(() => {
        if (!controller.signal.aborted) setScriptHash({ state: "failed" });
      });
    return () => controller.abort();
  }, [live]);

  async function importFile(file: File) {
    setImportState({ state: "reading", name: file.name });
    setConfirming(false);
    try {
      const body = await readExportFile(file);
      const result = await readJson(
        `${PATH}/import`,
        mailboxAccessImportResultSchema.parse,
        post(session, body),
      );
      setImportState({ state: "done", name: file.name, result });
      await load();
    } catch (problem) {
      setImportState({
        state: "failed",
        name: file.name,
        message: problem instanceof Error ? problem.message : "The export could not be imported.",
      });
    }
  }

  async function clear() {
    setClearing(true);
    try {
      await readJson(PATH, mailboxAccessClearSchema.parse, { ...post(session), method: "DELETE" });
      setImportState({ state: "idle" });
      await load();
    } catch (problem) {
      setError(problem instanceof Error ? problem.message : "The import could not be cleared.");
    } finally {
      setClearing(false);
      setConfirming(false);
    }
  }

  if (!live) return null;
  return (
    <MailboxAccessView
      status={status}
      error={error}
      importState={importState}
      confirming={confirming}
      clearing={clearing}
      readOnly={readOnly}
      scriptHash={scriptHash}
      version={APP_VERSION}
      onImport={(file) => void importFile(file)}
      onClear={() => setConfirming(true)}
      onConfirmClear={() => void clear()}
      onCancelClear={() => setConfirming(false)}
    />
  );
}

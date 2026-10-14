import {
  automationManifestSchema,
  automationPackageListSchema,
  automationPackageSummarySchema,
  type PrismSession,
} from "@prism365/contracts";
import { useCallback, useEffect, useId, useRef, useState } from "react";
import { LockIcon, ReportIcon } from "./icons.tsx";

type PackageList = ReturnType<typeof automationPackageListSchema.parse>;
type PackageSummary = ReturnType<typeof automationPackageSummarySchema.parse>;
type UploadDetails = {
  name: string;
  version: string;
  lifecycle: string;
  fileCount: number;
  bytes: number;
};

export function PackageMetadataList({
  items,
  canRemove,
  busy,
  onRequestRemoval,
}: {
  items: PackageSummary[];
  canRemove: boolean;
  busy: boolean;
  onRequestRemoval: (item: PackageSummary) => void;
}) {
  return (
    <ul className="package-list">
      {items.map((item) => (
        <li key={item.id}>
          <div>
            <strong>{item.name}</strong>
            <span>
              v{item.version} · {item.lifecycle} · {item.stepCount} declared steps
            </span>
            <small>
              Imported{" "}
              {new Intl.DateTimeFormat(undefined, {
                dateStyle: "medium",
                timeStyle: "short",
              }).format(new Date(item.importedAt))}
            </small>
            <small>
              Expires{" "}
              {new Intl.DateTimeFormat(undefined, {
                dateStyle: "medium",
                timeStyle: "short",
              }).format(
                new Date(new Date(item.importedAt).getTime() + 30 * 24 * 60 * 60 * 1000),
              )}{" "}
              · 30-day retention
            </small>
            <details>
              <summary>Immutable package digest</summary>
              <code>{item.digest}</code>
            </details>
          </div>
          <div className="package-list__actions">
            <span className="package-status">Unreviewed · execution disabled</span>
            {canRemove ? (
              <button
                className="button button--secondary"
                type="button"
                disabled={busy}
                aria-label={`Remove private package ${item.name} version ${item.version}`}
                onClick={() => onRequestRemoval(item)}
              >
                Remove package
              </button>
            ) : null}
          </div>
        </li>
      ))}
    </ul>
  );
}

export function PackageRemovalDialog({
  item,
  busy,
  error,
  onCancel,
  onConfirm,
}: {
  item: PackageSummary;
  busy: boolean;
  error: string;
  onCancel: () => void;
  onConfirm: () => void;
}) {
  const titleId = useId();
  const dialog = useRef<HTMLDialogElement>(null);
  useEffect(() => {
    const trigger = document.activeElement;
    dialog.current?.showModal();
    return () => {
      dialog.current?.close();
      if (trigger instanceof HTMLElement && trigger.isConnected) trigger.focus();
      else document.getElementById("script-packages-title")?.focus();
    };
  }, []);
  return (
    <dialog
      ref={dialog}
      className="permission-dialog package-removal-dialog"
      aria-labelledby={titleId}
      onCancel={(event) => {
        event.preventDefault();
        if (!busy) onCancel();
      }}
    >
      <span className="eyebrow">Private library</span>
      <h2 id={titleId}>Remove this stored package?</h2>
      <p className="dialog-intro">
        Prism365 will delete the private stored source for{" "}
        <strong>
          {item.name} v{item.version}
        </strong>
        . This cannot be undone through the library; keep your own copy. Removing a package makes no
        Microsoft 365 account changes.
      </p>
      {error ? (
        <p className="form-error investigation-error" role="alert">
          {error}
        </p>
      ) : null}
      <div className="package-removal-dialog__actions">
        <button
          className="button button--secondary"
          type="button"
          disabled={busy}
          onClick={onCancel}
        >
          Cancel
        </button>
        <button className="button button--danger" type="button" disabled={busy} onClick={onConfirm}>
          {busy ? "Deleting stored source…" : "Delete stored package"}
        </button>
      </div>
    </dialog>
  );
}

async function readResponse(response: Response): Promise<unknown> {
  const body: unknown = await response.json().catch(() => null);
  if (!response.ok) {
    const message =
      typeof body === "object" &&
      body !== null &&
      "message" in body &&
      typeof body.message === "string"
        ? body.message
        : `Package request could not complete (HTTP ${response.status}).`;
    throw new Error(message);
  }
  return body;
}

export function ScriptPackages({ session }: { session: PrismSession }) {
  const uploadId = useId();
  const [catalog, setCatalog] = useState<PackageList | null>(null);
  const [details, setDetails] = useState<UploadDetails | null>(null);
  const [imported, setImported] = useState<PackageSummary | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [loadError, setLoadError] = useState("");
  const [removalTarget, setRemovalTarget] = useState<PackageSummary | null>(null);
  const [removing, setRemoving] = useState(false);
  const [removalError, setRemovalError] = useState("");
  const [removedName, setRemovedName] = useState("");
  const uploadBody = useRef<string | null>(null);
  const uploadInput = useRef<HTMLInputElement>(null);
  const selectionNumber = useRef(0);
  const mounted = useRef(false);

  const loadCatalog = useCallback(async (signal?: AbortSignal) => {
    setLoadError("");
    try {
      const response = await fetch("/api/v1/automation/packages", {
        ...(signal ? { signal } : {}),
        credentials: "same-origin",
        cache: "no-store",
      });
      const result = automationPackageListSchema.parse(await readResponse(response));
      if (!signal?.aborted && mounted.current) setCatalog(result);
    } catch (caught) {
      if (!signal?.aborted && mounted.current)
        setLoadError(
          caught instanceof Error ? caught.message : "Unable to load the private package library.",
        );
    }
  }, []);

  useEffect(() => {
    mounted.current = true;
    const controller = new AbortController();
    void loadCatalog(controller.signal);
    return () => {
      mounted.current = false;
      controller.abort();
      uploadBody.current = null;
      selectionNumber.current += 1;
    };
  }, [loadCatalog]);

  async function chooseFile(file: File | undefined) {
    const selection = ++selectionNumber.current;
    uploadBody.current = null;
    setDetails(null);
    setImported(null);
    setError("");
    if (!file) return;
    if (!file.name.toLowerCase().endsWith(".json")) {
      setError("Choose a JSON automation package containing its manifest and script files.");
      return;
    }
    if (file.size > 1_048_576) {
      setError("Choose a package smaller than 1 MB for this first import version.");
      return;
    }
    try {
      const text = await file.text();
      const parsed: unknown = JSON.parse(text);
      if (
        typeof parsed !== "object" ||
        parsed === null ||
        !("manifest" in parsed) ||
        !("files" in parsed) ||
        !Array.isArray(parsed.files)
      ) {
        throw new Error("The package needs a manifest and a files array.");
      }
      const manifest = automationManifestSchema.parse(parsed.manifest);
      if (
        parsed.files.length === 0 ||
        parsed.files.some(
          (item: unknown) =>
            typeof item !== "object" ||
            item === null ||
            !("path" in item) ||
            !("content" in item) ||
            typeof item.path !== "string" ||
            typeof item.content !== "string",
        )
      ) {
        throw new Error("Each package file needs a path and its content.");
      }
      if (selection !== selectionNumber.current || !mounted.current) return;
      uploadBody.current = text;
      setDetails({
        name: manifest.name,
        version: manifest.version,
        lifecycle: manifest.lifecycle,
        fileCount: parsed.files.length,
        bytes: file.size,
      });
    } catch {
      if (selection === selectionNumber.current && mounted.current)
        setError(
          "This file is not a valid automation package. Check the manifest and files against the package format.",
        );
    }
  }

  async function importPackage(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!uploadBody.current || !session.csrfToken || !catalog?.importEnabled) return;
    setBusy(true);
    setError("");
    try {
      const response = await fetch("/api/v1/automation/packages", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json", "x-csrf-token": session.csrfToken },
        body: uploadBody.current,
      });
      const result = automationPackageSummarySchema.parse(await readResponse(response));
      if (!mounted.current) return;
      setImported(result);
      uploadBody.current = null;
      setDetails(null);
      if (uploadInput.current) uploadInput.current.value = "";
      await loadCatalog();
    } catch (caught) {
      if (mounted.current)
        setError(caught instanceof Error ? caught.message : "The package could not be imported.");
    } finally {
      if (mounted.current) setBusy(false);
    }
  }

  async function removePackage() {
    if (!removalTarget || !session.csrfToken || !catalog?.importEnabled) return;
    setRemoving(true);
    setRemovalError("");
    try {
      const response = await fetch(
        `/api/v1/automation/packages/${encodeURIComponent(removalTarget.id)}`,
        {
          method: "DELETE",
          credentials: "same-origin",
          headers: { "x-csrf-token": session.csrfToken },
        },
      );
      if (!response.ok) await readResponse(response);
      if (!mounted.current) return;
      setRemovedName(`${removalTarget.name} v${removalTarget.version}`);
      setCatalog((current) =>
        current
          ? { ...current, items: current.items.filter((item) => item.id !== removalTarget.id) }
          : null,
      );
      if (imported?.id === removalTarget.id) setImported(null);
      setRemovalTarget(null);
      await loadCatalog();
    } catch (caught) {
      if (mounted.current)
        setRemovalError(
          caught instanceof Error ? caught.message : "The private package could not be removed.",
        );
    } finally {
      if (mounted.current) setRemoving(false);
    }
  }

  const canImport =
    catalog?.importEnabled &&
    session.mode === "live" &&
    session.authenticated &&
    !!session.csrfToken;

  return (
    <section className="script-packages" aria-labelledby="script-packages-title">
      <div className="script-packages__header">
        <div>
          <span className="eyebrow">Private runbook library</span>
          <h2 id="script-packages-title" tabIndex={-1}>
            Bring your organization’s scripts into Prism365.
          </h2>
          <p>
            Import a package for private storage and review. Its code stays in your installation and
            is never sent to an AI provider.
          </p>
        </div>
        <span className="automation__lock">
          <LockIcon />
          Execution disabled
        </span>
      </div>
      <div className="source-state">
        <LockIcon />
        <div>
          <strong>Imports enter the library as unreviewed.</strong>
          <p>
            Uploading stores the package and calculates its digest. Running it requires a separately
            reviewed version, an isolated worker, approval, and an action history.
          </p>
          <p>
            Private source packages expire after 30 days. You can remove them earlier from this
            library.
          </p>
        </div>
      </div>
      {loadError ? (
        <p className="form-error investigation-error" role="alert">
          {loadError}
        </p>
      ) : null}
      {catalog && !canImport ? (
        <p className="source-notice">
          Private imports require a live administrator session and the installation’s durable
          database.{" "}
          {session.mode === "demo"
            ? "This demonstration cannot store scripts."
            : "Complete sign-in and storage setup to import organization scripts."}
        </p>
      ) : null}
      {!catalog && !loadError ? (
        <p className="source-notice" role="status">
          Checking private package storage…
        </p>
      ) : null}
      <form className="script-import" onSubmit={importPackage}>
        <label className="field" htmlFor={uploadId}>
          <span>Organization-owned automation package (.json)</span>
          <input
            ref={uploadInput}
            id={uploadId}
            type="file"
            accept=".json,application/json"
            disabled={!canImport || busy || removing}
            onChange={(event) => {
              void chooseFile(event.target.files?.[0]);
            }}
          />
          <small>
            Use only scripts you own or have permission to use. Keep passwords, tokens, and other
            secrets out of package files.
          </small>
        </label>
        <details className="package-format">
          <summary>What goes in the package?</summary>
          <p>
            A JSON object with a <code>manifest</code> describing the name, version, lifecycle,
            capabilities, entrypoint, and recovery steps; plus <code>files</code> containing each
            relative <code>path</code> and script <code>content</code>. The entrypoint must refer to
            one of those files.
          </p>
          <p>
            Prism365 validates the package on the server without running its code. Script source is
            not displayed in this library.
          </p>
        </details>
        {details ? (
          <div className="package-selection" role="status">
            <ReportIcon />
            <div>
              <strong>
                {details.name} · v{details.version}
              </strong>
              <span>
                {details.lifecycle} · {details.fileCount} files ·{" "}
                {(details.bytes / 1024).toFixed(1)} KB
              </span>
            </div>
          </div>
        ) : null}
        {error ? (
          <p className="form-error investigation-error" role="alert">
            {error}
          </p>
        ) : null}
        <button
          className="button button--primary"
          type="submit"
          disabled={!canImport || !details || busy || removing}
        >
          {busy ? "Importing privately…" : "Import package for review"}
        </button>
        {imported ? (
          <p className="import-success" role="status">
            {imported.name} v{imported.version} was imported privately. Status: unreviewed;
            execution remains disabled.
          </p>
        ) : null}
      </form>
      <div className="package-library">
        <div className="package-library__heading">
          <h3>Private packages</h3>
          <button
            className="button button--secondary"
            type="button"
            disabled={busy || removing}
            onClick={() => {
              void loadCatalog();
            }}
          >
            Refresh library
          </button>
        </div>
        {catalog?.items.length === 0 ? (
          <p className="empty-data">No private script packages have been imported.</p>
        ) : null}
        {removedName ? (
          <p className="import-success" role="status">
            {removedName} was removed from the private library. Its stored source cannot be
            recovered through Prism365.
          </p>
        ) : null}
        <PackageMetadataList
          items={catalog?.items ?? []}
          canRemove={!!canImport}
          busy={busy || removing}
          onRequestRemoval={(item) => {
            setRemovalError("");
            setRemovalTarget(item);
          }}
        />
      </div>
      {removalTarget ? (
        <PackageRemovalDialog
          item={removalTarget}
          busy={removing}
          error={removalError}
          onCancel={() => setRemovalTarget(null)}
          onConfirm={() => {
            void removePackage();
          }}
        />
      ) : null}
    </section>
  );
}

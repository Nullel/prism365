import { createHash, randomUUID } from "node:crypto";
import {
  type AutomationLifecycle,
  type AutomationManifest,
  automationManifestSchema,
} from "@prism365/contracts";

export interface AutomationPackageFile {
  readonly path: string;
  readonly content: string;
}

export interface AutomationPackageRecord {
  readonly id: string;
  readonly tenantId: string;
  readonly importedByActorId: string;
  readonly importedAt: string;
  readonly digest: string;
  readonly status: "unreviewed";
  readonly manifest: AutomationManifest;
  readonly files: readonly AutomationPackageFile[];
}

export interface AutomationPackageSummary {
  id: string;
  name: string;
  version: string;
  lifecycle: AutomationLifecycle;
  digest: string;
  status: "unreviewed";
  importedAt: string;
  stepCount: number;
}

const maxFiles = 50;
const maxBytes = 500_000;
const encoder = new TextEncoder();

/** Inspect organization-owned source as data. Never execute, extract, or approve it. */
export function inspectAutomationPackage(
  input: unknown,
  context: { tenantId: string; actorId: string; now?: Date },
): AutomationPackageRecord {
  if (!context.tenantId || !context.actorId)
    throw new Error("An authenticated tenant and importer are required.");
  if (!isRecord(input) || !hasExactKeys(input, ["manifest", "files"]))
    throw new Error("An automation import must contain only manifest and files.");
  const parsedManifest = automationManifestSchema.safeParse(input.manifest);
  if (!parsedManifest.success) throw new Error("The automation package manifest is invalid.");
  const manifest = parsedManifest.data;
  if (manifest.runtime !== "powershell7")
    throw new Error("Only PowerShell 7 script packages can be imported.");
  if (manifest.allowedHosts.length !== 0)
    throw new Error(
      "Imported packages must use capability-broker mode with no direct network hosts.",
    );
  validateFilePath(manifest.entrypoint);
  if (!manifest.entrypoint.endsWith(".ps1"))
    throw new Error("The package entrypoint must be a .ps1 file.");
  if (!Array.isArray(input.files) || input.files.length === 0 || input.files.length > maxFiles)
    throw new Error("An automation package must contain between 1 and 50 files.");
  const paths = new Set<string>();
  const files: AutomationPackageFile[] = [];
  let byteCount = encoder.encode(canonicalize(manifest)).byteLength;
  for (const value of input.files) {
    if (
      !isRecord(value) ||
      !hasExactKeys(value, ["path", "content"]) ||
      typeof value.path !== "string" ||
      typeof value.content !== "string"
    )
      throw new Error("Every package file must contain only a path and text content.");
    validateFilePath(value.path);
    const foldedPath = value.path.toLowerCase();
    if (paths.has(foldedPath))
      throw new Error("Package file paths must be unique, including case differences.");
    paths.add(foldedPath);
    // Precheck prevents encoding a very large submitted string solely to reject it.
    if (value.content.length > maxBytes || value.content.includes("\0"))
      throw new Error("A script file is too large or contains binary content.");
    byteCount += encoder.encode(value.path).byteLength + encoder.encode(value.content).byteLength;
    if (byteCount > maxBytes)
      throw new Error("An automation package exceeds the 500000-byte limit.");
    files.push({ path: value.path, content: value.content });
  }
  if (!files.some((file) => file.path === manifest.entrypoint))
    throw new Error("The declared package entrypoint is missing.");
  // Stable ordering makes equivalent JSON imports produce the same content identity.
  files.sort((left, right) => (left.path < right.path ? -1 : left.path > right.path ? 1 : 0));
  const serialized = canonicalize({ manifest, files });
  if (encoder.encode(serialized).byteLength > maxBytes)
    throw new Error("An automation package exceeds the 500000-byte limit.");
  const importedAt = (context.now ?? new Date()).toISOString();
  const record: AutomationPackageRecord = {
    id: randomUUID(),
    tenantId: context.tenantId,
    importedByActorId: context.actorId,
    importedAt,
    digest: createHash("sha256").update(serialized, "utf8").digest("hex"),
    status: "unreviewed",
    manifest,
    files,
  };
  return freezeRecursively(record);
}

/** Public metadata only: never return proprietary script source through this view. */
export function summarizeAutomationPackage(
  record: AutomationPackageRecord,
): AutomationPackageSummary {
  return {
    id: record.id,
    name: record.manifest.name,
    version: record.manifest.version,
    lifecycle: record.manifest.lifecycle,
    digest: record.digest,
    status: record.status,
    importedAt: record.importedAt,
    stepCount: record.manifest.steps.length,
  };
}

function validateFilePath(path: string): void {
  if (
    path.length > 200 ||
    !/^[A-Za-z0-9][A-Za-z0-9._/-]*$/.test(path) ||
    path.includes("..") ||
    path.split("/").some((segment) => !segment || segment === ".")
  ) {
    throw new Error("Package files must use safe, fixed relative paths.");
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function hasExactKeys(value: Record<string, unknown>, keys: readonly string[]): boolean {
  const received = Object.keys(value);
  return received.length === keys.length && keys.every((key) => Object.hasOwn(value, key));
}

function canonicalize(value: unknown): string {
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(canonicalize).join(",")}]`;
  const entries = Object.entries(value).sort(([left], [right]) =>
    left < right ? -1 : left > right ? 1 : 0,
  );
  return `{${entries.map(([key, child]) => `${JSON.stringify(key)}:${canonicalize(child)}`).join(",")}}`;
}

function freezeRecursively<T>(value: T): T {
  if (typeof value === "object" && value !== null) {
    for (const child of Object.values(value)) freezeRecursively(child);
    Object.freeze(value);
  }
  return value;
}

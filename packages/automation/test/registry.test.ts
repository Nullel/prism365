import { describe, expect, it } from "vitest";
import { inspectAutomationPackage, summarizeAutomationPackage } from "../src/registry.ts";

const context = {
  tenantId: "synthetic-tenant",
  actorId: "synthetic-actor",
  now: new Date("2026-10-13T12:00:00.000Z"),
};
const manifest = {
  schemaVersion: 1,
  id: "example.offboarding",
  name: "Example organization offboarding",
  version: "1.0.0",
  description: "Synthetic organization-owned source fixture.",
  lifecycle: "offboarding",
  runtime: "powershell7",
  entrypoint: "scripts/offboard.ps1",
  publisher: "Example organization",
  parameters: [],
  capabilities: ["identity.lookup"],
  allowedHosts: [],
  timeoutSeconds: 60,
  steps: [
    {
      id: "check-target",
      label: "Check the target",
      description: "Read account state.",
      capability: "identity.lookup",
      risk: "readOnly",
      reversible: true,
      recovery: "No changes.",
    },
  ],
};
const source = "# Synthetic fixture: never execute\nparam([string] $TargetEmail)\n";
const input = { manifest, files: [{ path: manifest.entrypoint, content: source }] };

describe("automation package inspection", () => {
  it("returns an immutable tenant-bound unreviewed package without execution", () => {
    const record = inspectAutomationPackage(input, context);
    expect(record).toMatchObject({
      tenantId: context.tenantId,
      importedByActorId: context.actorId,
      status: "unreviewed",
      importedAt: context.now.toISOString(),
    });
    expect(record.digest).toMatch(/^[a-f0-9]{64}$/);
    expect(Object.isFrozen(record)).toBe(true);
    expect(Object.isFrozen(record.manifest.steps)).toBe(true);
    expect(Object.isFrozen(record.files[0])).toBe(true);
    expect(() => {
      (record.files[0] as { content: string }).content = "changed";
    }).toThrow();
  });

  it("exposes a metadata-only summary without script source or importer identity", () => {
    const record = inspectAutomationPackage(input, context);
    const summary = summarizeAutomationPackage(record);
    expect(summary).toMatchObject({
      id: record.id,
      name: manifest.name,
      version: "1.0.0",
      lifecycle: "offboarding",
      digest: record.digest,
      status: "unreviewed",
      stepCount: 1,
    });
    expect(summary).not.toHaveProperty("files");
    expect(summary).not.toHaveProperty("manifest");
    expect(summary).not.toHaveProperty("importedByActorId");
    expect(JSON.stringify(summary)).not.toContain(source);
  });

  it.each([
    "../offboard.ps1",
    "/offboard.ps1",
    "scripts/../offboard.ps1",
    "scripts//offboard.ps1",
    "scripts/./offboard.ps1",
    "scripts\\offboard.ps1",
    "scripts/offboard.ps1\n",
    "C:/offboard.ps1",
    "scripts/offboard.ps1/",
    "scripts/$offboard.ps1",
  ])('rejects unsafe file path "%s"', (path) => {
    expect(() =>
      inspectAutomationPackage(
        { ...input, files: [...input.files, { path, content: "synthetic" }] },
        context,
      ),
    ).toThrow("safe, fixed relative paths");
  });

  it("rejects duplicate and case-colliding file paths", () => {
    for (const path of [manifest.entrypoint, "SCRIPTS/OFFBOARD.PS1"]) {
      expect(() =>
        inspectAutomationPackage(
          { ...input, files: [...input.files, { path, content: "different" }] },
          context,
        ),
      ).toThrow("unique");
    }
  });

  it("requires the fixed PowerShell entrypoint to exist", () => {
    expect(() =>
      inspectAutomationPackage(
        { ...input, files: [{ path: "scripts/other.ps1", content: source }] },
        context,
      ),
    ).toThrow("entrypoint is missing");
    expect(() =>
      inspectAutomationPackage(
        { ...input, manifest: { ...manifest, entrypoint: "scripts/offboard.exe" } },
        context,
      ),
    ).toThrow(".ps1");
    expect(() =>
      inspectAutomationPackage(
        { ...input, manifest: { ...manifest, runtime: "prism365-native" } },
        context,
      ),
    ).toThrow("Only PowerShell 7");
  });

  it("rejects direct network hosts and undeclared capabilities", () => {
    expect(() =>
      inspectAutomationPackage(
        { ...input, manifest: { ...manifest, allowedHosts: ["example.invalid"] } },
        context,
      ),
    ).toThrow("no direct network hosts");
    expect(() =>
      inspectAutomationPackage(
        { ...input, manifest: { ...manifest, capabilities: ["groups.remove"] } },
        context,
      ),
    ).toThrow("manifest is invalid");
  });

  it("rejects untrusted approval and ownership fields on the import envelope", () => {
    for (const extra of [
      { status: "approved" },
      { digest: "fake" },
      { tenantId: "other-tenant" },
      { importedByActorId: "other-actor" },
    ]) {
      expect(() => inspectAutomationPackage({ ...input, ...extra }, context)).toThrow(
        "only manifest and files",
      );
    }
  });

  it("rejects file envelope overrides and binary content", () => {
    expect(() =>
      inspectAutomationPackage(
        { ...input, files: [{ ...input.files[0], executable: true }] },
        context,
      ),
    ).toThrow("only a path and text content");
    expect(() =>
      inspectAutomationPackage(
        { ...input, files: [{ path: manifest.entrypoint, content: "a\0b" }] },
        context,
      ),
    ).toThrow("binary content");
  });

  it("bounds file count, aggregate UTF-8 bytes and canonical content bytes", () => {
    expect(() =>
      inspectAutomationPackage(
        {
          ...input,
          files: Array.from({ length: 51 }, (_, index) => ({
            path: `scripts/file${index}.ps1`,
            content: "x",
          })),
        },
        context,
      ),
    ).toThrow("50 files");
    expect(() =>
      inspectAutomationPackage(
        { ...input, files: [{ path: manifest.entrypoint, content: "é".repeat(250_000) }] },
        context,
      ),
    ).toThrow("500000-byte");
    expect(() =>
      inspectAutomationPackage(
        { ...input, files: [{ path: manifest.entrypoint, content: "\\".repeat(260_000) }] },
        context,
      ),
    ).toThrow("500000-byte");
  });

  it("binds every content or manifest change to a different digest", () => {
    const first = inspectAutomationPackage(input, context);
    const changedScript = inspectAutomationPackage(
      { ...input, files: [{ path: manifest.entrypoint, content: `${source}# changed\n` }] },
      context,
    );
    const changedManifest = inspectAutomationPackage(
      { ...input, manifest: { ...manifest, version: "1.0.1" } },
      context,
    );
    expect(changedScript.digest).not.toBe(first.digest);
    expect(changedManifest.digest).not.toBe(first.digest);
    expect(changedScript.status).toBe("unreviewed");
    expect(changedManifest.status).toBe("unreviewed");
  });

  it("canonicalizes JSON key and file ordering without normalizing script bytes", () => {
    const files = [
      ...input.files,
      { path: "support/readme.txt", content: "synthetic supporting text" },
    ];
    const first = inspectAutomationPackage({ manifest, files }, context);
    const reorderedManifest = Object.fromEntries(Object.entries(manifest).reverse());
    const second = inspectAutomationPackage(
      { manifest: reorderedManifest, files: [...files].reverse() },
      { ...context, actorId: "another-actor" },
    );
    expect(second.digest).toBe(first.digest);
    expect(second.id).not.toBe(first.id);
    const changedNewlines = inspectAutomationPackage(
      {
        ...input,
        files: [{ path: manifest.entrypoint, content: source.replaceAll("\n", "\r\n") }],
      },
      context,
    );
    expect(changedNewlines.digest).not.toBe(inspectAutomationPackage(input, context).digest);
  });
});

import {
  automationPackageSummarySchema,
  sessionSchema,
  userInvestigationSchema,
} from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";
import { AiInvestigation, availableAiDatasets } from "../src/AiInvestigation.tsx";
import {
  PackageMetadataList,
  PackageRemovalDialog,
  ScriptPackages,
} from "../src/ScriptPackages.tsx";

const investigation = userInvestigationSchema.parse({
  mode: "demo",
  generatedAt: "2026-10-13T00:00:00Z",
  targetEmail: "example@example.invalid",
  profile: {
    state: "available",
    data: {
      id: "synthetic-profile",
      displayName: "Example user",
      userPrincipalName: "example@example.invalid",
      mail: null,
      jobTitle: null,
      department: null,
      userType: "Member",
    },
  },
  licenses: { state: "missingConsent", data: [] },
  signIns: {
    state: "missingRole",
    data: [],
    window: { start: "2026-10-12T00:00:00Z", end: "2026-10-13T00:00:00Z" },
    complete: false,
  },
  limitations: [],
});
const demoSession = sessionSchema.parse({
  mode: "demo",
  configured: false,
  authenticated: false,
  displayName: null,
  csrfToken: null,
  enabledFeatures: [],
  missingConfiguration: [],
});
const privatePackage = {
  ...automationPackageSummarySchema.parse({
    id: "00000000-0000-4000-8000-000000000017",
    name: "Example private offboarding",
    version: "1.0.0",
    lifecycle: "offboarding",
    digest: "a".repeat(64),
    status: "unreviewed",
    importedAt: "2026-10-13T00:00:00Z",
    stepCount: 3,
  }),
  source: "synthetic-source-canary",
};

describe("AI evidence selection and private import interface", () => {
  it("excludes sources that the administrator could not access", () => {
    expect(availableAiDatasets(investigation)).toEqual(["directory"]);
    expect(
      availableAiDatasets({
        ...investigation,
        profile: { state: "providerUnavailable", data: null },
      }),
    ).toEqual([]);
  });

  it("does not present a grant or model answer until provider status is known", () => {
    const html = renderToStaticMarkup(
      createElement(AiInvestigation, { investigation, session: demoSession }),
    );
    expect(html).toContain("Checking the local AI provider");
    expect(html).not.toContain("AI-generated interpretation");
    expect(html).not.toContain("checked");
    expect(html).not.toContain("Allow these sources");
  });

  it("keeps package upload disabled before storage readiness and offers no execution", () => {
    const html = renderToStaticMarkup(createElement(ScriptPackages, { session: demoSession }));
    expect(html).toContain("Execution disabled");
    expect(html).toContain("unreviewed");
    expect(html).toContain('type="file"');
    expect(html).toContain('accept=".json,application/json"');
    expect(html).toContain("disabled");
    expect(html).not.toContain("Execute package");
  });

  it("shows expiry and metadata without exposing source or removing a package on render", () => {
    const remove = vi.fn();
    const html = renderToStaticMarkup(
      createElement(PackageMetadataList, {
        items: [privatePackage],
        canRemove: true,
        busy: false,
        onRequestRemoval: remove,
      }),
    );
    expect(html).toContain("30-day retention");
    expect(html).toContain("Unreviewed · execution disabled");
    expect(html).toContain("Remove private package Example private offboarding version 1.0.0");
    expect(html).not.toContain("synthetic-source-canary");
    expect(remove).not.toHaveBeenCalled();
  });

  it("hides package removal when private storage access is unavailable", () => {
    const html = renderToStaticMarkup(
      createElement(PackageMetadataList, {
        items: [privatePackage],
        canRemove: false,
        busy: false,
        onRequestRemoval: () => {},
      }),
    );
    expect(html).not.toContain("Remove package");
    expect(html).toContain(privatePackage.digest);
  });

  it("requires a separate named confirmation with a cancel option before source deletion", () => {
    const confirm = vi.fn();
    const html = renderToStaticMarkup(
      createElement(PackageRemovalDialog, {
        item: privatePackage,
        busy: false,
        error: "",
        onCancel: () => {},
        onConfirm: confirm,
      }),
    );
    expect(html).toContain("<dialog");
    expect(html).toContain("Example private offboarding v1.0.0");
    expect(html).toContain("Cancel");
    expect(html).toContain("Delete stored package");
    expect(confirm).not.toHaveBeenCalled();
  });
});

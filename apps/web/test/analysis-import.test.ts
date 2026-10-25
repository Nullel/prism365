import { sessionSchema, userInvestigationSchema } from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { AiInvestigation, availableAiDatasets } from "../src/AiInvestigation.tsx";

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

describe("AI evidence selection", () => {
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
});

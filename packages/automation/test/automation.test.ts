import { describe, expect, it } from "vitest";
import { createAutomationPreview, listAutomationWorkflows } from "../src/index.ts";

describe("lifecycle automation", () => {
  it("lists only reviewed demo workflows while import and execution remain disabled", () => {
    const catalog = listAutomationWorkflows();

    expect(catalog).toMatchObject({
      mode: "demo",
      importEnabled: false,
      executionEnabled: false,
    });
    expect(catalog.items).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ lifecycle: "onboarding", trust: "approved" }),
        expect.objectContaining({
          lifecycle: "offboarding",
          runtime: "powershell7",
          origin: "imported",
          trust: "approved",
        }),
      ]),
    );
  });

  it("creates a non-executable offboarding preview with exact permissions and recovery text", () => {
    const preview = createAutomationPreview(
      {
        workflowId: "prism365.secure-offboarding",
        targetEmail: "alex.rivera@example.invalid",
      },
      new Date("2026-10-13T18:00:00.000Z"),
    );

    expect(preview.executionAllowed).toBe(false);
    expect(preview.confirmationPhrase).toBe("OFFBOARDING alex.rivera@example.invalid");
    expect(preview.requiredPermissions).toContain("User.RevokeSessions.All");
    expect(preview.steps).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          capability: "identity.user.disable",
          risk: "high",
          reversible: true,
        }),
        expect.objectContaining({
          capability: "mailbox.handoff",
          status: "manualReview",
        }),
      ]),
    );
    expect(preview.blockedReasons).not.toHaveLength(0);
  });

  it("rejects an unknown workflow", () => {
    expect(() =>
      createAutomationPreview({
        workflowId: "unknown",
        targetEmail: "alex.rivera@example.invalid",
      }),
    ).toThrow("Unknown lifecycle workflow");
  });
});

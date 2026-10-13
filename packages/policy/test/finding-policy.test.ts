import { describe, expect, it } from "vitest";
import { evaluateFinding } from "../src/index.ts";

const evidence = {
  hasSource: true,
  hasTimeWindow: true,
  hasLimitations: true,
} as const;

describe("finding policy", () => {
  it("allows an evidence-backed security finding", () => {
    expect(
      evaluateFinding({
        title: "Legacy authentication observed",
        summary: "Two synthetic sign-ins used a legacy client in the selected window.",
        purpose: "security",
        subject: "tenant",
        ...evidence,
      }),
    ).toEqual({ allowed: true, warnings: [] });
  });

  it("blocks workforce performance purposes", () => {
    expect(
      evaluateFinding({
        title: "Worker score",
        summary: "Rank people using activity counts.",
        purpose: "workforce-performance",
        subject: "person",
        ...evidence,
      }),
    ).toMatchObject({ allowed: false });
  });

  it("blocks productivity claims hidden in display language", () => {
    expect(
      evaluateFinding({
        title: "Low performers",
        summary: "People with fewer meetings.",
        purpose: "adoption",
        subject: "person",
        ...evidence,
      }),
    ).toMatchObject({ allowed: false });
  });

  it("requires source and time-window evidence", () => {
    expect(
      evaluateFinding({
        title: "Configuration changed",
        summary: "A synthetic setting changed.",
        purpose: "operations",
        subject: "resource",
        ...evidence,
        hasSource: false,
      }),
    ).toMatchObject({ allowed: false });
  });
});

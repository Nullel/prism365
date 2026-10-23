import { describe, expect, it } from "vitest";
import { graphSignInSchema, normalizeSignIn } from "../src/source.ts";

const wellFormed = (value: string) =>
  !/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(value);

describe("sign-in text from Microsoft", () => {
  it("never stores half an emoji or a NUL, whatever length it is cut to", () => {
    // An emoji (two UTF-16 units) straddling the cut point of every field limit used: 80 (risk
    // event types), 100 (device), 200, 300 (names), 1,000 (failure reason).
    const pads = [80, 100, 200, 300, 320, 1_000].flatMap((limit) =>
      [-2, -1, 0, 1].map((offset) => limit + offset),
    );
    for (const pad of pads) {
      const long = `${"a".repeat(pad)}\u{1F600}${"b".repeat(1_200)}\u0000`;
      const parsed = graphSignInSchema.parse({
        id: `id-\u0000${pad}`,
        createdDateTime: "2026-10-22T12:00:00Z",
        userDisplayName: long,
        deviceDetail: { browser: long, operatingSystem: long },
        status: { errorCode: 0, failureReason: long },
        riskEventTypes_v2: [long],
      });
      const event = normalizeSignIn(parsed);
      for (const value of [
        event.id,
        event.displayName ?? "",
        event.deviceBrowser ?? "",
        event.deviceOs ?? "",
        event.failureReason ?? "",
        ...event.riskEventTypes,
      ]) {
        expect(wellFormed(value), `${pad}: ${value.slice(-4)}`).toBe(true);
        expect(value.includes("\u0000")).toBe(false);
      }
    }
  });
});

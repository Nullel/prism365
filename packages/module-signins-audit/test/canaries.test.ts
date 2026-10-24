import { describe, expect, it } from "vitest";
import {
  CANARY_RULE,
  canaryDay,
  canaryDedupeKey,
  canaryFinding,
  canaryIndex,
  matchCanary,
  parseCanaryList,
  type SigninEvent,
} from "../src/index.ts";

// Synthetic accounts and documentation-range addresses only.
const decoy = { userId: "decoy-1", principal: "Decoy.Account@example.invalid" };
const index = canaryIndex([decoy]);
const at = new Date("2026-10-22T23:59:30Z");

function event(overrides: Partial<SigninEvent> = {}): SigninEvent {
  return {
    id: "event-1",
    createdAt: at,
    userId: "decoy-1",
    upn: "decoy.account@example.invalid",
    displayName: "Decoy Account",
    appId: null,
    appName: "Office 365 Exchange Online",
    resourceName: null,
    clientApp: "Browser",
    ip: "203.0.113.20",
    country: "NL",
    state: "Noord-Holland",
    city: "Amsterdam",
    lat: 52.37,
    lon: 4.9,
    deviceId: null,
    deviceOs: "Linux",
    deviceBrowser: "Chrome",
    isCompliant: false,
    isManaged: false,
    trustType: null,
    errorCode: 0,
    failureReason: null,
    caStatus: "success",
    riskLevel: null,
    riskLevelAggregated: null,
    riskState: null,
    riskEventTypes: [],
    isInteractive: true,
    ...overrides,
  };
}

describe("canary matching", () => {
  it("matches a sign-in by user id, whatever its result", () => {
    expect(matchCanary(event(), index)).toEqual(decoy);
    expect(matchCanary(event({ errorCode: 50126 }), index)).toEqual(decoy);
    expect(matchCanary(event({ errorCode: 53003 }), index)).toEqual(decoy);
    expect(matchCanary(event({ errorCode: -1 }), index)).toEqual(decoy);
  });

  it("does not match an account that is not a canary", () => {
    expect(matchCanary(event({ userId: "user-2", upn: "someone@example.invalid" }), index)).toBe(
      null,
    );
    expect(matchCanary(event(), canaryIndex([]))).toBe(null);
  });

  it("matches by principal name when the sign-in has no user id, in any letter case or form", () => {
    const noId = { userId: "" };
    expect(matchCanary(event({ ...noId, upn: "decoy.account@example.invalid" }), index)).toEqual(
      decoy,
    );
    expect(matchCanary(event({ ...noId, upn: "DECOY.ACCOUNT@EXAMPLE.INVALID" }), index)).toEqual(
      decoy,
    );
    // Fullwidth letters fold to the same name under NFKC.
    expect(
      matchCanary(event({ ...noId, upn: "ＤＥＣＯＹ.account@example.invalid" }), index),
    ).toEqual(decoy);
    expect(matchCanary(event({ ...noId, upn: "other@example.invalid" }), index)).toBe(null);
    expect(matchCanary(event({ ...noId, upn: null }), index)).toBe(null);
  });

  it("trusts the user id when there is one: a different id with the canary's name is not a canary", () => {
    expect(matchCanary(event({ userId: "user-2" }), index)).toBe(null);
  });
});

describe("canary finding", () => {
  it("opens one finding per account and UTC day", () => {
    expect(canaryDay(new Date("2026-10-22T23:59:59.999Z"))).toBe("2026-10-22");
    expect(canaryDay(new Date("2026-10-23T00:00:00.000Z"))).toBe("2026-10-23");
    const key = canaryDedupeKey(decoy, at);
    expect(key).toBe("canary:decoy-1:2026-10-22");
    // Another attempt the same day is the same finding; the next day is a new one.
    expect(canaryDedupeKey(decoy, new Date("2026-10-22T00:00:01Z"))).toBe(key);
    expect(canaryDedupeKey(decoy, new Date("2026-10-23T00:00:01Z"))).not.toBe(key);
    expect(canaryDedupeKey({ ...decoy, userId: "decoy-2" }, at)).not.toBe(key);
  });

  it("is critical for a successful sign-in, with the sign-in as evidence", () => {
    const finding = canaryFinding(event(), decoy);
    expect(finding).toMatchObject({
      ruleId: CANARY_RULE,
      purpose: "security",
      severity: "critical",
      dedupeKey: "canary:decoy-1:2026-10-22",
      subject: { kind: "person", id: "decoy-1", label: "Decoy Account" },
      observedAt: at,
    });
    expect(CANARY_RULE).toBe("signin.canary-used");
    expect(finding.summary).toBe(
      "A successful sign-in on Decoy Account, a canary account. It came from Amsterdam, Noord-Holland, NL. The app was Office 365 Exchange Online. Nobody should be using this account.",
    );
    expect(finding.evidence).toEqual([
      {
        source: "Microsoft Entra sign-in logs",
        ref: "event-1",
        observedAt: "2026-10-22T23:59:30.000Z",
        fields: {
          result: "success",
          errorCode: 0,
          failureReason: null,
          ip: "203.0.113.20",
          location: "Amsterdam, Noord-Holland, NL",
          app: "Office 365 Exchange Online",
          client: "Browser",
        },
      },
    ]);
    expect(finding.limitations).toContain(
      "Canaries are accounts you marked; Prism365 cannot tell a mistaken sign-in from an attack.",
    );
  });

  it("is critical for a failed attempt too, and carries the error code", () => {
    const finding = canaryFinding(
      event({ errorCode: 50126, failureReason: "Invalid username or password." }),
      decoy,
    );
    expect(finding.severity).toBe("critical");
    expect(finding.summary).toContain("failed");
    expect(finding.summary).toContain("50126");
    expect(finding.evidence[0]?.fields).toMatchObject({
      result: "failure",
      errorCode: 50126,
      failureReason: "Invalid username or password.",
    });
    expect(canaryFinding(event({ errorCode: -1 }), decoy).evidence[0]?.fields).toMatchObject({
      result: "unknown",
    });
  });

  it("names the account from the list when the sign-in carried no name", () => {
    const finding = canaryFinding(event({ userId: "", upn: null, displayName: null }), decoy);
    expect(finding.subject).toMatchObject({
      id: "decoy-1",
      label: "Decoy.Account@example.invalid",
    });
    expect(finding.dedupeKey).toBe("canary:decoy-1:2026-10-22");
  });

  it("stays inside the finding limits for odd provider values", () => {
    const finding = canaryFinding(
      event({ displayName: "x".repeat(400), failureReason: "y".repeat(2_000), errorCode: 50126 }),
      decoy,
    );
    expect(finding.subject.label?.length).toBeLessThanOrEqual(200);
    expect(finding.summary.length).toBeLessThanOrEqual(1_000);
    expect(String(finding.evidence[0]?.fields.failureReason).length).toBeLessThanOrEqual(300);
    expect(finding.title.length).toBeLessThanOrEqual(160);
  });
});

describe("canary list parsing", () => {
  const entry = (userId: string) => ({
    userId,
    principal: `${userId}@example.invalid`,
    addedAt: "2026-10-22T12:00:00.000Z",
    addedBy: "admin-1",
  });

  it("reads a stored list, as an object or as JSON text", () => {
    expect(parseCanaryList({ accounts: [entry("a")] }).accounts).toEqual([entry("a")]);
    expect(parseCanaryList(JSON.stringify({ accounts: [entry("a")] })).accounts).toEqual([
      entry("a"),
    ]);
  });

  it("reads nothing as an empty list, and keeps the good entries of a damaged one", () => {
    expect(parseCanaryList(null).accounts).toEqual([]);
    expect(parseCanaryList("not json").accounts).toEqual([]);
    expect(parseCanaryList({ accounts: "no" }).accounts).toEqual([]);
    const damaged = { accounts: [entry("a"), { userId: "b" }, 7, entry("c"), entry("a")] };
    expect(parseCanaryList(damaged).accounts.map((item) => item.userId)).toEqual(["a", "c"]);
  });

  it("keeps at most 100", () => {
    const many = { accounts: Array.from({ length: 150 }, (_value, i) => entry(`user-${i}`)) };
    expect(parseCanaryList(many).accounts).toHaveLength(100);
  });
});

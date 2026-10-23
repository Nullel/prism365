import { describe, expect, it } from "vitest";
import {
  defaultSigninsAuditConfig,
  type FailureRecord,
  failurePatterns,
  findingKey,
  graphSignInSchema,
  haversineKm,
  inNetworks,
  maskIp,
  normalizeSignIn,
  type SigninEvent,
  type SuccessHistory,
  signinsAuditConfigFromEnv,
  successReasons,
  totalScore,
} from "../src/index.ts";

// Synthetic accounts, documentation-range addresses (RFC 5737 / RFC 3849).
const config = defaultSigninsAuditConfig;
const at = new Date("2026-10-22T12:00:00Z");

function event(overrides: Partial<SigninEvent> = {}): SigninEvent {
  return {
    id: "event-1",
    createdAt: at,
    userId: "user-1",
    upn: "synthetic.user@example.invalid",
    displayName: "Synthetic User",
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

function history(overrides: Partial<SuccessHistory> = {}): SuccessHistory {
  return {
    hasBaseline: true,
    trusted: { ip: false, country: false, state: false, device: false },
    previous: null,
    recentFailures: 0,
    recentSprayFailures: 0,
    hourShare: null,
    ...overrides,
  };
}

const rules = (reasons: { rule: string }[]) =>
  reasons.map((reason) => reason.rule).filter((rule) => rule !== "context.location");

describe("success rules", () => {
  it("flags a new address, country and state on an unknown device", () => {
    const reasons = successReasons(event(), history(), config, false);
    expect(rules(reasons)).toEqual(["new.ip", "new.country", "new.state"]);
    expect(totalScore(reasons)).toBe(60);
  });

  it("stays quiet on a known device, a trusted network, or an account with no history", () => {
    expect(
      rules(
        successReasons(
          event(),
          history({ trusted: { ip: false, country: false, state: false, device: true } }),
          config,
          false,
        ),
      ),
    ).toEqual([]);
    expect(
      rules(
        successReasons(
          event(),
          history(),
          { ...config, trustedNetworks: ["203.0.113.0/24"] },
          false,
        ),
      ),
    ).toEqual([]);
    expect(rules(successReasons(event(), history({ hasBaseline: false }), config, false))).toEqual(
      [],
    );
  });

  it("flags impossible travel but not geolocation jitter between nearby places", () => {
    const fromChicago = history({
      trusted: { ip: true, country: true, state: true, device: false },
      previous: { createdAt: new Date(at.getTime() - 2 * 3_600_000), lat: 41.88, lon: -87.63 },
    });
    expect(rules(successReasons(event(), fromChicago, config, false))).toEqual([
      "travel.impossible",
    ]);
    const jitter = history({
      trusted: { ip: true, country: true, state: true, device: false },
      previous: { createdAt: new Date(at.getTime() - 60_000), lat: 52.08, lon: 4.3 },
    });
    expect(haversineKm(52.08, 4.3, 52.37, 4.9)).toBeLessThan(config.travelMinimumKm);
    expect(rules(successReasons(event(), jitter, config, false))).toEqual([]);
  });

  it("applies only baseline-free rules while learning history", () => {
    const reasons = successReasons(
      event({ riskLevel: "high", clientApp: "IMAP" }),
      history({ recentFailures: 6 }),
      config,
      true,
    );
    expect(rules(reasons)).toEqual(["legacy.auth", "risk.entra", "success.after_failures"]);
  });

  it("flags dormant accounts, unexpected countries, and success after failures", () => {
    const quiet = history({
      trusted: { ip: true, country: true, state: true, device: true },
      previous: { createdAt: new Date(at.getTime() - 70 * 86_400_000), lat: 52.37, lon: 4.9 },
      recentSprayFailures: 1,
    });
    expect(
      rules(successReasons(event(), quiet, { ...config, expectedCountries: ["US", "CA"] }, false)),
    ).toEqual(["dormant.revival", "country.unexpected", "success.after_failures"]);
  });

  it("never lets an unusual hour or legacy authentication raise a finding alone", () => {
    const familiar = history({
      trusted: { ip: true, country: true, state: true, device: true },
      hourShare: 0,
    });
    expect(rules(successReasons(event(), familiar, config, false))).toEqual([]);
    const legacy = successReasons(event({ clientApp: "IMAP" }), familiar, config, false);
    expect(rules(legacy)).toEqual(["legacy.auth"]);
    expect(totalScore(legacy)).toBeLessThan(config.minimumScore);
    const unfamiliar = successReasons(
      event(),
      history({ trusted: { ip: false, country: true, state: true, device: false }, hourShare: 0 }),
      config,
      false,
    );
    expect(rules(unfamiliar)).toEqual(["new.ip", "hour.unusual"]);
  });
});

function failure(overrides: Partial<FailureRecord> & { minute: number }): FailureRecord {
  const { minute, ...rest } = overrides;
  return {
    id: `f-${minute}-${rest.userId ?? "user-2"}-${rest.errorCode ?? 50126}`,
    createdAt: new Date(at.getTime() + minute * 60_000),
    userId: "user-2",
    ip: "203.0.113.9",
    country: "XA",
    errorCode: 50126,
    ...rest,
  };
}

describe("failure patterns", () => {
  it("finds a burst from one address and against one account at the threshold, not below", () => {
    const four = [0, 1, 2, 3].map((minute) => failure({ minute }));
    expect(failurePatterns(four, four, config)).toEqual([]);
    const five = [...four, failure({ minute: 4 })];
    expect(
      failurePatterns(five, five, config)
        .map((pattern) => pattern.rule)
        .sort(),
    ).toEqual(["failed_burst_account", "failed_burst_ip"]);
  });

  it("finds a password spray across accounts", () => {
    const spray = ["a", "b", "c", "d", "e"].map((user, index) =>
      failure({ minute: index * 10, userId: `user-${user}`, ip: "198.51.100.7" }),
    );
    const found = failurePatterns(spray, spray, config);
    expect(found.find((pattern) => pattern.rule === "password_spray")).toMatchObject({
      key: "198.51.100.7",
      accounts: 5,
    });
  });

  it("ignores interrupts that belong to normal sign-ins and trusted networks", () => {
    const interrupts = Array.from({ length: 10 }, (_v, minute) =>
      failure({ minute, errorCode: 50074 }),
    );
    expect(failurePatterns(interrupts, interrupts, config)).toEqual([]);
    const office = Array.from({ length: 10 }, (_v, minute) => failure({ minute, ip: "192.0.2.5" }));
    const fromOffice = failurePatterns(office, office, {
      ...config,
      trustedNetworks: ["192.0.2.0/24"],
    });
    expect(fromOffice.map((pattern) => pattern.rule)).toEqual(["failed_burst_account"]);
  });

  it("finds repeated MFA denials for one account", () => {
    const denials = [0, 2, 4].map((minute) => failure({ minute, errorCode: 500121 }));
    expect(failurePatterns(denials, denials, config).map((pattern) => pattern.rule)).toEqual([
      "mfa_denials",
    ]);
  });
});

describe("source handling", () => {
  it("masks addresses to their network and matches trusted ranges", () => {
    expect(maskIp("203.0.113.47")).toBe("203.0.113.0/24");
    expect(maskIp("2001:db8:abcd:12::1")).toBe("2001:db8:abcd::/48");
    expect(maskIp("not an ip")).toBeNull();
    expect(inNetworks("192.0.2.200", ["192.0.2.0/24"])).toBe(true);
    expect(inNetworks("192.0.3.1", ["192.0.2.0/24"])).toBe(false);
    expect(inNetworks("2001:db8::5", ["2001:db8::/32"])).toBe(true);
    expect(inNetworks("2001:db9::5", ["2001:db8::/32"])).toBe(false);
  });

  it("treats hidden risk as absent, missing status as unknown, and odd optional fields as null", () => {
    const record = graphSignInSchema.parse({
      id: "event-x",
      createdDateTime: "2026-10-22T12:00:00Z",
      riskLevelDuringSignIn: "hidden",
      deviceDetail: { isCompliant: "yes", operatingSystem: 7 },
      location: "somewhere",
      extraField: "dropped",
    });
    const normalized = normalizeSignIn(record);
    expect(normalized).toMatchObject({
      riskLevel: null,
      errorCode: -1,
      isCompliant: null,
      deviceOs: null,
      country: null,
    });
    expect("extraField" in record).toBe(false);
  });

  it("reads operator settings from the environment and rejects bad values", () => {
    expect(
      signinsAuditConfigFromEnv({
        SIGNINS_EXPECTED_COUNTRIES: "us, ca",
        SIGNINS_TRUSTED_NETWORKS: "192.0.2.0/24",
        SIGNINS_TIME_ZONE: "America/Chicago",
      }),
    ).toMatchObject({
      expectedCountries: ["US", "CA"],
      trustedNetworks: ["192.0.2.0/24"],
      timeZone: "America/Chicago",
    });
    expect(() => signinsAuditConfigFromEnv({ SIGNINS_TIME_ZONE: "Mars/Base" })).toThrow("IANA");
    expect(() => signinsAuditConfigFromEnv({ SIGNINS_TRUSTED_NETWORKS: "192.0.2.0" })).toThrow();
  });
});

describe("finding keys", () => {
  it("keeps short keys and hashes long ones to a stable 200-character-safe value", () => {
    expect(findingKey("user-1|NL|Noord-Holland")).toBe("user-1|NL|Noord-Holland");
    const long = "x".repeat(400);
    expect(findingKey(long)).toMatch(/^sha256:[0-9a-f]{64}$/);
    expect(findingKey(long)).toBe(findingKey(long));
  });
});

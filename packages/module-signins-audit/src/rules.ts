import type { SigninsAuditConfig } from "./config.ts";
import {
  credentialFailureCodes,
  inNetworks,
  isLegacyClient,
  mfaFailureCodes,
  type SigninEvent,
} from "./source.ts";

// Detection rules. Pure: each takes an event plus the history the collector loaded for it and
// returns reasons with scores. Nothing here reads storage or calls Microsoft, so every rule is
// tested directly. Scores add up per sign-in; a finding needs config.minimumScore.

export const RULES_VERSION = 1;

export interface Reason {
  rule: string;
  score: number;
  detail: string;
}

export interface SuccessHistory {
  // The account has at least one trusted baseline value. Without any, the sign-in only teaches
  // the baseline: a brand-new account must not raise "new country, new state, new address".
  hasBaseline: boolean;
  trusted: { ip: boolean; country: boolean; state: boolean; device: boolean };
  // Previous successful sign-in for this account (any location), if one exists.
  previous: { createdAt: Date; lat: number | null; lon: number | null } | null;
  // Credential failures for this account in the success-after-failures window before this event,
  // and whether any came from an address that was password-spraying.
  recentFailures: number;
  recentSprayFailures: number;
  // Share of this account's successful sign-ins that fell in this event's hour, or null when there
  // is under 30 days of history.
  hourShare: number | null;
}

const substantiveRules: ReadonlySet<string> = new Set([
  "new.ip",
  "new.country",
  "new.state",
  "travel.impossible",
  "dormant.revival",
  "risk.entra",
  "success.after_failures",
]);

export function haversineKm(lat1: number, lon1: number, lat2: number, lon2: number): number {
  const radius = 6_371;
  const toRadians = (value: number) => (value * Math.PI) / 180;
  const dLat = toRadians(lat2 - lat1);
  const dLon = toRadians(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(toRadians(lat1)) * Math.cos(toRadians(lat2)) * Math.sin(dLon / 2) ** 2;
  return radius * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// Rules for a successful sign-in. `learning` is true during the first-run history load: only
// rules that do not depend on the account's baseline apply then (Microsoft's risk rating,
// legacy authentication, unexpected country, success after a failure burst).
export function successReasons(
  event: SigninEvent,
  history: SuccessHistory,
  config: SigninsAuditConfig,
  learning: boolean,
): Reason[] {
  const reasons: Reason[] = [];
  const onTrustedNetwork = inNetworks(event.ip, config.trustedNetworks);
  const knownDevice = history.trusted.device;
  const location = [event.city, event.state, event.country].filter(Boolean).join(", ");

  if (!learning && history.hasBaseline && !onTrustedNetwork && !knownDevice) {
    if (event.ip && !history.trusted.ip)
      reasons.push({
        rule: "new.ip",
        score: 20,
        detail: "First successful sign-in from this address.",
      });
    if (event.country && !history.trusted.country)
      reasons.push({
        rule: "new.country",
        score: 15,
        detail: `First successful sign-in from ${event.country}.`,
      });
    if (event.state && !history.trusted.state)
      reasons.push({
        rule: "new.state",
        score: 25,
        detail: `First successful sign-in from ${event.state}.`,
      });
  }

  const previous = history.previous;
  if (
    !learning &&
    !onTrustedNetwork &&
    !knownDevice &&
    previous &&
    previous.lat !== null &&
    previous.lon !== null &&
    event.lat !== null &&
    event.lon !== null
  ) {
    const hours = (event.createdAt.getTime() - previous.createdAt.getTime()) / 3_600_000;
    const km = haversineKm(previous.lat, previous.lon, event.lat, event.lon);
    // A minimum distance keeps geolocation jitter between nearby cities from reading as travel.
    if (hours > 0 && hours <= config.travelMaxHours && km >= config.travelMinimumKm) {
      const speed = km / hours;
      if (speed > config.travelKmh)
        reasons.push({
          rule: "travel.impossible",
          score: 70,
          detail: `${Math.round(km)} km from the previous sign-in in ${hours.toFixed(1)} h (${Math.round(speed)} km/h).`,
        });
    }
  }

  if (
    !learning &&
    previous &&
    event.createdAt.getTime() - previous.createdAt.getTime() >= config.dormantDays * 86_400_000
  )
    reasons.push({
      rule: "dormant.revival",
      score: 25,
      detail: `First successful sign-in in ${Math.floor((event.createdAt.getTime() - previous.createdAt.getTime()) / 86_400_000)} days.`,
    });

  if (
    !onTrustedNetwork &&
    config.expectedCountries.length > 0 &&
    event.country &&
    !config.expectedCountries.includes(event.country)
  )
    reasons.push({
      rule: "country.unexpected",
      score: 20,
      detail: `${event.country} is not one of the organisation's expected countries.`,
    });

  if (isLegacyClient(event.clientApp))
    reasons.push({
      rule: "legacy.auth",
      score: 30,
      detail: `Legacy authentication client (${event.clientApp}).`,
    });

  if (event.riskLevel === "medium" || event.riskLevel === "high")
    reasons.push({
      rule: "risk.entra",
      score: event.riskLevel === "high" ? 70 : 50,
      detail: `Microsoft Entra ID Protection rated this sign-in ${event.riskLevel} risk.`,
    });

  if (history.recentFailures >= config.bruteForceFailures || history.recentSprayFailures > 0)
    reasons.push({
      rule: "success.after_failures",
      score: 40,
      detail:
        history.recentSprayFailures > 0
          ? "Successful sign-in shortly after failed attempts from an address trying many accounts."
          : `Successful sign-in after ${history.recentFailures} failed attempts in ${config.successAfterFailuresMinutes} minutes.`,
    });

  // Never enough on its own, and never stacked onto legacy authentication or an unexpected
  // country alone: it only adds weight to signals about this sign-in itself.
  const substantive = reasons.some((reason) => substantiveRules.has(reason.rule));
  if (!learning && history.hourShare !== null && history.hourShare < 0.02 && substantive)
    reasons.push({
      rule: "hour.unusual",
      score: 10,
      detail: "Outside this account's usual sign-in hours.",
    });

  if (location && reasons.length > 0 && !reasons.some((reason) => reason.detail.includes(location)))
    reasons.push({ rule: "context.location", score: 0, detail: `Reported location: ${location}.` });
  return reasons;
}

export function totalScore(reasons: readonly Reason[]): number {
  return Math.min(
    1_000,
    reasons.reduce((total, reason) => total + reason.score, 0),
  );
}

export function severityFor(score: number): "medium" | "high" | "critical" {
  if (score >= 90) return "critical";
  if (score >= 70) return "high";
  return "medium";
}

export interface FailureRecord {
  id: string;
  createdAt: Date;
  userId: string;
  ip: string | null;
  country: string | null;
  errorCode: number;
}

export interface FailurePattern {
  rule: "failed_burst_ip" | "failed_burst_account" | "password_spray" | "mfa_denials";
  key: string;
  count: number;
  accounts: number;
  start: Date;
  end: Date;
  eventIds: string[];
  score: number;
}

const within = (record: FailureRecord, end: Date, minutes: number) =>
  record.createdAt.getTime() <= end.getTime() &&
  record.createdAt.getTime() >= end.getTime() - minutes * 60_000;

// Failure patterns that end at one of the newly collected events. `window` holds every failure
// in the lookback the collector loaded, including older ones, so a burst that straddles two
// polls is still counted once it crosses the threshold.
export function failurePatterns(
  window: readonly FailureRecord[],
  fresh: readonly FailureRecord[],
  config: SigninsAuditConfig,
): FailurePattern[] {
  const found = new Map<string, FailurePattern>();
  const keep = (pattern: FailurePattern) => {
    const id = `${pattern.rule}|${pattern.key}`;
    const current = found.get(id);
    if (!current || pattern.count > current.count) found.set(id, pattern);
  };
  const credential = window.filter((record) => credentialFailureCodes.has(record.errorCode));
  const mfa = window.filter((record) => mfaFailureCodes.has(record.errorCode));

  for (const event of fresh) {
    if (credentialFailureCodes.has(event.errorCode)) {
      if (event.ip && !inNetworks(event.ip, config.trustedNetworks)) {
        const byIp = credential.filter(
          (record) =>
            record.ip === event.ip && within(record, event.createdAt, config.bruteForceMinutes),
        );
        if (byIp.length >= config.bruteForceFailures)
          keep({
            rule: "failed_burst_ip",
            key: event.ip,
            count: byIp.length,
            accounts: new Set(byIp.map((record) => record.userId)).size,
            start: byIp[0]?.createdAt ?? event.createdAt,
            end: event.createdAt,
            eventIds: byIp.slice(-10).map((record) => record.id),
            score: 50 + Math.min(30, (byIp.length - config.bruteForceFailures) * 5),
          });
        const sprayed = credential.filter(
          (record) =>
            record.ip === event.ip && within(record, event.createdAt, config.sprayMinutes),
        );
        const accounts = new Set(sprayed.map((record) => record.userId).filter(Boolean));
        if (accounts.size >= config.sprayAccounts)
          keep({
            rule: "password_spray",
            key: event.ip,
            count: sprayed.length,
            accounts: accounts.size,
            start: sprayed[0]?.createdAt ?? event.createdAt,
            end: event.createdAt,
            eventIds: sprayed.slice(-10).map((record) => record.id),
            score: 60 + Math.min(30, (accounts.size - config.sprayAccounts) * 5),
          });
      }
      if (event.userId) {
        const byUser = credential.filter(
          (record) =>
            record.userId === event.userId &&
            within(record, event.createdAt, config.bruteForceMinutes),
        );
        if (byUser.length >= config.bruteForceFailures)
          keep({
            rule: "failed_burst_account",
            key: event.userId,
            count: byUser.length,
            accounts: 1,
            start: byUser[0]?.createdAt ?? event.createdAt,
            end: event.createdAt,
            eventIds: byUser.slice(-10).map((record) => record.id),
            score: 40 + Math.min(30, (byUser.length - config.bruteForceFailures) * 5),
          });
      }
    }
    if (mfaFailureCodes.has(event.errorCode) && event.userId) {
      const denials = mfa.filter(
        (record) =>
          record.userId === event.userId && within(record, event.createdAt, config.mfaMinutes),
      );
      if (denials.length >= config.mfaDenials)
        keep({
          rule: "mfa_denials",
          key: event.userId,
          count: denials.length,
          accounts: 1,
          start: denials[0]?.createdAt ?? event.createdAt,
          end: event.createdAt,
          eventIds: denials.slice(-10).map((record) => record.id),
          score: 50,
        });
    }
  }
  return [...found.values()];
}

// The spraying addresses inside a window, for success.after_failures.
export function sprayingAddresses(
  window: readonly FailureRecord[],
  config: SigninsAuditConfig,
): Set<string> {
  const byIp = new Map<string, Set<string>>();
  for (const record of window) {
    if (!record.ip || !credentialFailureCodes.has(record.errorCode)) continue;
    const accounts = byIp.get(record.ip) ?? new Set<string>();
    accounts.add(record.userId);
    byIp.set(record.ip, accounts);
  }
  return new Set(
    [...byIp].filter(([, accounts]) => accounts.size >= config.sprayAccounts).map(([ip]) => ip),
  );
}

import { z } from "zod";

// Operator settings for the sign-in audit module. Defaults follow the values the earlier Entra
// Guard prototype ran with, plus the guards that prototype lacked (minimum travel distance,
// learning only from confirmed history).

const cidrPattern =
  /^(?:(?:\d{1,3}\.){3}\d{1,3}\/(?:[0-9]|[12]\d|3[0-2])|[0-9a-f:]+\/(?:\d|[1-9]\d|1[01]\d|12[0-8]))$/i;

export const signinsAuditConfigSchema = z.object({
  // ISO 3166 alpha-2 codes where the organisation normally works. Empty disables the rule.
  expectedCountries: z
    .array(z.string().regex(/^[A-Z]{2}$/))
    .max(250)
    .default([]),
  // Office or VPN egress networks. Novelty and location rules are not applied to them.
  trustedNetworks: z.array(z.string().regex(cidrPattern)).max(100).default([]),
  // IANA zone used for the "unusual hour" histogram.
  timeZone: z.string().min(1).max(64).default("UTC"),
  // Sign-in records can arrive late; each poll re-reads this far behind its cursor.
  lateArrivalMinutes: z.number().int().min(5).max(240).default(60),
  // History loaded on first run. Microsoft keeps sign-ins 30 days with Entra ID P1 or P2.
  backfillDays: z.number().int().min(0).max(30).default(30),
  // Records Graph may return for one request before the window is split. It also bounds how
  // long one request runs, so the run budget is checked often: Microsoft served about eight
  // sign-ins a second in the first live tenant, so 500 is roughly a minute.
  windowRecordLimit: z.number().int().min(2).max(10_000).default(500),
  // Time a run may spend reading before it saves its place and continues on the next run. No
  // Graph page starts after it. The scheduler stops a run at 240 s and one request attempt may
  // take 90 s (throttled retries can add up to 30 s of waiting), so a page started at the last
  // moment usually still finishes in time; when it does not, the next run sees the window it
  // was reading and treats it as failed (see `trying` in cursor.ts).
  runBudgetSeconds: z.number().int().min(10).max(200).default(135),
  bruteForceFailures: z.number().int().min(2).max(100).default(5),
  bruteForceMinutes: z.number().int().min(1).max(240).default(15),
  sprayAccounts: z.number().int().min(2).max(100).default(5),
  sprayMinutes: z.number().int().min(5).max(240).default(60),
  mfaDenials: z.number().int().min(2).max(50).default(3),
  mfaMinutes: z.number().int().min(1).max(120).default(15),
  successAfterFailuresMinutes: z.number().int().min(5).max(240).default(30),
  travelKmh: z.number().min(100).max(5_000).default(900),
  travelMaxHours: z.number().min(1).max(48).default(6),
  travelMinimumKm: z.number().min(50).max(5_000).default(300),
  dormantDays: z.number().int().min(14).max(365).default(60),
  minimumScore: z.number().int().min(10).max(200).default(40),
});

export type SigninsAuditConfig = z.infer<typeof signinsAuditConfigSchema>;

export const defaultSigninsAuditConfig: SigninsAuditConfig = signinsAuditConfigSchema.parse({});

function list(value: string | undefined): string[] {
  return (value ?? "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

// Reads SIGNINS_EXPECTED_COUNTRIES, SIGNINS_TRUSTED_NETWORKS and SIGNINS_TIME_ZONE. Invalid
// values fail startup rather than silently disabling a rule.
export function signinsAuditConfigFromEnv(
  env: Record<string, string | undefined>,
): SigninsAuditConfig {
  const timeZone = env.SIGNINS_TIME_ZONE?.trim();
  if (timeZone) {
    try {
      new Intl.DateTimeFormat("en-US", { timeZone });
    } catch {
      throw new Error("SIGNINS_TIME_ZONE must be an IANA time zone such as America/Chicago.");
    }
  }
  const parsed = signinsAuditConfigSchema.safeParse({
    expectedCountries: list(env.SIGNINS_EXPECTED_COUNTRIES).map((code) => code.toUpperCase()),
    trustedNetworks: list(env.SIGNINS_TRUSTED_NETWORKS),
    ...(timeZone ? { timeZone } : {}),
  });
  if (!parsed.success)
    throw new Error(
      "SIGNINS_EXPECTED_COUNTRIES must be two-letter codes and SIGNINS_TRUSTED_NETWORKS must be CIDR ranges.",
    );
  return parsed.data;
}

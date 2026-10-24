import { isIP } from "node:net";
import { cleanText } from "@prism365/contracts";
import { z } from "zod";

// Microsoft Graph v1.0, verified 2026-10-22:
// https://learn.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/resources/signin?view=graph-rest-1.0
// - Delegated and application access both need AuditLog.Read.All; delegated callers also need
//   Global Reader, Reports Reader, Security Administrator, Security Operator or Security Reader.
// - The tenant needs Microsoft Entra ID P1 or P2. Retention is 30 days on P1/P2.
// - v1.0 returns interactive user sign-ins (and successful federated sign-ins).
// - Supported query options are $top (max 1000), $skiptoken and $filter. There is no $select
//   or $orderby; results arrive newest first, so the collector sorts locally.
// - Risk fields read "hidden" without Entra ID P2.

export const signInContext = {
  feature: "audit.signins.read",
  permission: "AuditLog.Read.All",
} as const;

// Optional fields tolerate unexpected shapes (they become null) and long values are truncated.
// Identity and time stay strict; a record that fails them is left out and counted by the
// collector, never allowed to fail the page it arrived on.
const text = (max: number) =>
  z
    .string()
    .transform((value) => cleanText(value.slice(0, max)))
    .nullable()
    .optional()
    .catch(null);
const flag = () => z.boolean().nullable().optional().catch(null);

// Unknown fields are dropped here: the module never stores more of a sign-in than it uses.
export const graphSignInSchema = z.object({
  id: z
    .string()
    .max(200)
    .transform((value) => cleanText(value))
    .pipe(z.string().min(1)),
  createdDateTime: z.string().datetime({ offset: true }),
  userId: text(200),
  userPrincipalName: text(320),
  userDisplayName: text(300),
  appId: text(200),
  appDisplayName: text(300),
  resourceDisplayName: text(300),
  ipAddress: text(64),
  clientAppUsed: text(80),
  conditionalAccessStatus: text(40),
  isInteractive: flag(),
  riskLevelDuringSignIn: text(20),
  riskLevelAggregated: text(20),
  riskState: text(40),
  riskEventTypes_v2: z
    .array(z.string().transform((value) => cleanText(value.slice(0, 80))))
    .transform((values) => values.slice(0, 30))
    .nullable()
    .optional()
    .catch(null),
  status: z
    .object({
      errorCode: z.number().int(),
      failureReason: text(1_000),
    })
    .nullable()
    .optional()
    .catch(null),
  deviceDetail: z
    .object({
      deviceId: text(200),
      operatingSystem: text(100),
      browser: text(100),
      isCompliant: flag(),
      isManaged: flag(),
      trustType: text(60),
    })
    .nullable()
    .optional()
    .catch(null),
  location: z
    .object({
      city: text(120),
      state: text(120),
      countryOrRegion: text(8),
      geoCoordinates: z
        .object({
          latitude: z.number().min(-90).max(90).nullable().optional().catch(null),
          longitude: z.number().min(-180).max(180).nullable().optional().catch(null),
        })
        .nullable()
        .optional()
        .catch(null),
    })
    .nullable()
    .optional()
    .catch(null),
});

export type GraphSignIn = z.infer<typeof graphSignInSchema>;

export interface SigninEvent {
  id: string;
  createdAt: Date;
  userId: string;
  upn: string | null;
  displayName: string | null;
  appId: string | null;
  appName: string | null;
  resourceName: string | null;
  clientApp: string | null;
  ip: string | null;
  country: string | null;
  state: string | null;
  city: string | null;
  lat: number | null;
  lon: number | null;
  deviceId: string | null;
  deviceOs: string | null;
  deviceBrowser: string | null;
  isCompliant: boolean | null;
  isManaged: boolean | null;
  trustType: string | null;
  errorCode: number;
  failureReason: string | null;
  caStatus: string | null;
  riskLevel: string | null;
  riskLevelAggregated: string | null;
  riskState: string | null;
  riskEventTypes: string[];
  isInteractive: boolean | null;
}

const blank = (value: string | null | undefined) => {
  const trimmed = value?.trim();
  return trimmed ? trimmed : null;
};

// "hidden" means the tenant has no Entra ID P2; it is absence of data, not a risk level.
const risk = (value: string | null | undefined) => {
  const level = blank(value)?.toLowerCase() ?? null;
  return level === "hidden" || level === "unknownfuturevalue" ? null : level;
};

export function normalizeSignIn(record: GraphSignIn): SigninEvent {
  const ip = blank(record.ipAddress);
  const lat = record.location?.geoCoordinates?.latitude;
  const lon = record.location?.geoCoordinates?.longitude;
  return {
    id: record.id,
    createdAt: new Date(record.createdDateTime),
    userId: blank(record.userId) ?? "",
    upn: blank(record.userPrincipalName)?.toLowerCase() ?? null,
    displayName: blank(record.userDisplayName),
    appId: blank(record.appId),
    appName: blank(record.appDisplayName),
    resourceName: blank(record.resourceDisplayName),
    clientApp: blank(record.clientAppUsed),
    ip: ip && isIP(ip) ? ip.toLowerCase() : null,
    country: blank(record.location?.countryOrRegion)?.toUpperCase() ?? null,
    state: blank(record.location?.state),
    city: blank(record.location?.city),
    lat: typeof lat === "number" && typeof lon === "number" ? lat : null,
    lon: typeof lat === "number" && typeof lon === "number" ? lon : null,
    deviceId: blank(record.deviceDetail?.deviceId),
    deviceOs: blank(record.deviceDetail?.operatingSystem),
    deviceBrowser: blank(record.deviceDetail?.browser),
    isCompliant: record.deviceDetail?.isCompliant ?? null,
    isManaged: record.deviceDetail?.isManaged ?? null,
    trustType: blank(record.deviceDetail?.trustType),
    // Unknown status is -1: neither a success that teaches the baseline nor a counted failure.
    errorCode: record.status?.errorCode ?? -1,
    failureReason: blank(record.status?.failureReason),
    caStatus: blank(record.conditionalAccessStatus),
    riskLevel: risk(record.riskLevelDuringSignIn),
    riskLevelAggregated: risk(record.riskLevelAggregated),
    riskState: blank(record.riskState),
    riskEventTypes: (record.riskEventTypes_v2 ?? []).filter(Boolean),
    isInteractive: record.isInteractive ?? null,
  };
}

// Documented credential failures (https://learn.microsoft.com/en-us/entra/identity-platform/reference-error-codes):
// 50126 invalid username or password, 50053 account locked or sign-in from a known-malicious
// address, 50034 account not found, 50056 invalid or null password. Interrupts that belong to a
// normal sign-in (50074, 50076, 50079, 50125, 50140, 50158, 50072, 50097) are not failures.
export const credentialFailureCodes: ReadonlySet<number> = new Set([50126, 50053, 50034, 50056]);

// 500121 ("authentication failed during strong authentication request") is not in Microsoft's
// published code reference; Microsoft Q&A answers describe it as a failed or denied MFA prompt.
// Rules that use it are marked experimental.
export const mfaFailureCodes: ReadonlySet<number> = new Set([500121]);

export const conditionalAccessBlockCode = 53003;

export const legacyClients: ReadonlySet<string> = new Set(
  ["exchange activesync", "imap", "imap4", "mapi", "smtp", "pop", "pop3", "other clients"].map(
    (value) => value.toLowerCase(),
  ),
);

export function isLegacyClient(clientApp: string | null): boolean {
  return clientApp ? legacyClients.has(clientApp.trim().toLowerCase()) : false;
}

function ipv4Parts(ip: string): number[] | null {
  const parts = ip.split(".").map(Number);
  return parts.length === 4 &&
    parts.every((part) => Number.isInteger(part) && part >= 0 && part <= 255)
    ? parts
    : null;
}

function ipv6Groups(ip: string): number[] | null {
  if (isIP(ip) !== 6) return null;
  const [head = "", tail = ""] = ip.split("::");
  const parse = (value: string) => (value ? value.split(":") : []);
  const left = parse(head);
  const right = ip.includes("::") ? parse(tail) : [];
  if (left.some((group) => group.includes(".")) || right.some((group) => group.includes(".")))
    return null;
  const missing = 8 - left.length - right.length;
  const groups = [
    ...left,
    ...Array.from({ length: ip.includes("::") ? missing : 0 }, () => "0"),
    ...right,
  ];
  return groups.length === 8 ? groups.map((group) => Number.parseInt(group, 16)) : null;
}

// Network label for views that are not a single-account investigation.
export function maskIp(ip: string | null): string | null {
  if (!ip) return null;
  const v4 = ipv4Parts(ip);
  if (v4) return `${v4[0]}.${v4[1]}.${v4[2]}.0/24`;
  const v6 = ipv6Groups(ip);
  if (v6)
    return `${v6
      .slice(0, 3)
      .map((group) => group.toString(16))
      .join(":")}::/48`;
  return null;
}

export function inNetworks(ip: string | null, networks: readonly string[]): boolean {
  if (!ip || networks.length === 0) return false;
  const v4 = ipv4Parts(ip);
  const v6 = v4 ? null : ipv6Groups(ip);
  for (const network of networks) {
    const [base = "", bitsText = ""] = network.split("/");
    const bits = Number(bitsText);
    if (v4) {
      const range = ipv4Parts(base);
      if (!range || !Number.isInteger(bits) || bits < 0 || bits > 32) continue;
      const toNumber = (parts: number[]) => parts.reduce((total, part) => total * 256 + part, 0);
      const block = 2 ** (32 - bits);
      if (Math.floor(toNumber(v4) / block) === Math.floor(toNumber(range) / block)) return true;
    } else if (v6) {
      const range = ipv6Groups(base);
      if (!range || !Number.isInteger(bits) || bits < 0 || bits > 128) continue;
      let remaining = bits;
      let match = true;
      for (let index = 0; index < 8 && remaining > 0; index += 1) {
        const take = Math.min(16, remaining);
        const shift = 16 - take;
        if ((v6[index] ?? 0) >> shift !== (range[index] ?? 0) >> shift) {
          match = false;
          break;
        }
        remaining -= take;
      }
      if (match) return true;
    }
  }
  return false;
}

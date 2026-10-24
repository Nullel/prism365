import { createPrivateKey, X509Certificate } from "node:crypto";
import { readFile } from "node:fs/promises";
import {
  type DirectoryChangesConfig,
  directoryChangesConfigFromEnv,
} from "@prism365/module-directory-changes";
import { type SigninsAuditConfig, signinsAuditConfigFromEnv } from "@prism365/module-signins-audit";
import { type DeliveryConfig, deliveryConfigFromEnv } from "./delivery-smtp.ts";

export const supportedReadFeatures = {
  "directory.users.read": "User.Read.All",
  "user.licenses.read": "LicenseAssignment.Read.All",
  "tenant.licenses.read": "LicenseAssignment.Read.All",
  "audit.signins.read": "AuditLog.Read.All",
  "reports.authentication-methods.read": "AuditLog.Read.All",
  "service.health.read": "ServiceHealth.Read.All",
  "service.messages.read": "ServiceMessage.Read.All",
  "audit.directory.read": "AuditLog.Read.All",
  "identity.conditional-access.read": "Policy.Read.All",
  "identity.risk.read": "IdentityRiskyUser.Read.All",
  "devices.intune.read": "DeviceManagementManagedDevices.Read.All",
  "directory.applications.read": "Application.Read.All",
  "directory.permission-grants.read": "Directory.Read.All",
  "reports.usage.read": "Reports.Read.All",
  "reports.settings.read": "ReportSettings.Read.All",
} as const;
export type SupportedReadFeature = keyof typeof supportedReadFeatures;

export interface IdentityConfig {
  tenantId: string;
  clientId: string;
  origin: string;
  redirectUri: string;
  allowedActorIds: string[];
  enabledFeatures: SupportedReadFeature[];
  credential:
    | { clientSecret: string }
    | { clientCertificate: { thumbprintSha256: string; privateKey: string } };
}

export interface AppConfig {
  mode: "demo" | "live";
  origin: string;
  production: boolean;
  identity: IdentityConfig | null;
  missingConfiguration: string[];
  enabledFeatures: SupportedReadFeature[];
  databaseUrl?: string;
  encryptionKey?: string;
  aiModel?: string;
  // PGlite data directory for module data when no PostgreSQL URL is configured. Unset means
  // module data lives in memory and is gone when the process stops.
  moduleDataDir?: string;
  // Sign-in audit settings (SIGNINS_EXPECTED_COUNTRIES, SIGNINS_TRUSTED_NETWORKS, SIGNINS_TIME_ZONE).
  signinsAudit?: SigninsAuditConfig;
  // Directory change feed settings (DIRECTORY_CHANGES_LATE_MINUTES, DIRECTORY_CHANGES_RETENTION_DAYS).
  directoryChanges?: DirectoryChangesConfig;
  // Reverse proxies in front of the API whose X-Forwarded-For entry is trusted for the client
  // address (TRUST_PROXY_HOPS). 0 trusts none. The Docker install runs behind one proxy.
  trustProxyHops?: number;
  // The configured public certificate's fingerprint and expiry, for the setup checklist. Public
  // facts only; the private key stays in identity.credential.
  certificate?: { thumbprintSha256: string; expiresAt: string };
  // How long handoffs and receipts are kept (HANDOFF_RETENTION_DAYS). Unset means the default.
  handoffRetentionDays?: number;
  // Alert delivery: the mail relay (SMTP_URL, SMTP_PASSWORD_FILE, SMTP_FROM, ALERT_EMAIL_DOMAINS)
  // and the extra ports webhooks may use (WEBHOOK_ALLOWED_PORTS). Live mode only.
  delivery?: DeliveryConfig;
}

// Handoffs and receipts are kept three years unless the operator chooses otherwise: long enough
// to answer what was done to an account at the last access review, no longer by default.
export const DEFAULT_HANDOFF_RETENTION_DAYS = 1_095;

export function isUuid(value: string) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);
}

export async function loadConfig(env: NodeJS.ProcessEnv = process.env): Promise<AppConfig> {
  const mode = env.DATA_MODE ?? "demo";
  if (mode !== "demo" && mode !== "live") throw new Error("DATA_MODE must be demo or live.");
  const production = env.NODE_ENV === "production";
  const origin = env.PUBLIC_ORIGIN ?? "http://localhost:5173";
  const parsedOrigin = new URL(origin);
  if (parsedOrigin.origin !== origin || parsedOrigin.username || parsedOrigin.password) {
    throw new Error("PUBLIC_ORIGIN must be an exact origin without path or credentials.");
  }
  if (
    parsedOrigin.protocol !== "https:" &&
    !(
      parsedOrigin.protocol === "http:" &&
      (!production || mode === "demo") &&
      ["localhost", "127.0.0.1"].includes(parsedOrigin.hostname)
    )
  ) {
    throw new Error("PUBLIC_ORIGIN requires HTTPS except for local development.");
  }
  const featureNames = (env.ENABLED_READ_FEATURES ?? "directory.users.read")
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);
  if (featureNames.some((feature) => !(feature in supportedReadFeatures))) {
    throw new Error("ENABLED_READ_FEATURES contains an unsupported feature.");
  }
  const enabledFeatures = [...new Set(featureNames)] as SupportedReadFeature[];
  const base: AppConfig = {
    mode,
    origin,
    production,
    identity: null,
    missingConfiguration: [],
    enabledFeatures,
  };
  if (mode === "demo") return base;

  const missing: string[] = [];
  for (const key of ["ENTRA_TENANT_ID", "ENTRA_CLIENT_ID", "ENTRA_ADMIN_OBJECT_IDS"] as const) {
    if (!env[key]) missing.push(key);
  }
  if (env.MICROSOFT_CLOUD && env.MICROSOFT_CLOUD !== "worldwide") {
    throw new Error("Live authentication currently supports the Worldwide Microsoft cloud only.");
  }
  const tenantId = env.ENTRA_TENANT_ID ?? "";
  const clientId = env.ENTRA_CLIENT_ID ?? "";
  const allowedActorIds = (env.ENTRA_ADMIN_OBJECT_IDS ?? "")
    .split(",")
    .map((value) => value.trim().toLowerCase())
    .filter(Boolean);
  if (
    (tenantId && !isUuid(tenantId)) ||
    (clientId && !isUuid(clientId)) ||
    allowedActorIds.some((value) => !isUuid(value))
  ) {
    throw new Error("Entra tenant, client, and administrator object IDs must be UUIDs.");
  }
  if (allowedActorIds.length === 0 && !missing.includes("ENTRA_ADMIN_OBJECT_IDS"))
    missing.push("ENTRA_ADMIN_OBJECT_IDS");
  let credential: IdentityConfig["credential"] | null = null;
  if (env.ENTRA_CERTIFICATE_PATH && env.ENTRA_PRIVATE_KEY_PATH) {
    try {
      const certificate = new X509Certificate(await readFile(env.ENTRA_CERTIFICATE_PATH));
      const privateKey = await readFile(env.ENTRA_PRIVATE_KEY_PATH, "utf8");
      if (!certificate.checkPrivateKey(createPrivateKey(privateKey)))
        throw new Error("Certificate mismatch");
      const thumbprintSha256 = certificate.fingerprint256.replaceAll(":", "");
      credential = { clientCertificate: { thumbprintSha256, privateKey } };
      base.certificate = {
        thumbprintSha256,
        expiresAt: new Date(certificate.validTo).toISOString(),
      };
    } catch {
      throw new Error("Unable to load the matching Entra certificate and private key.");
    }
  } else if (env.ENTRA_CLIENT_SECRET && !production) {
    credential = { clientSecret: env.ENTRA_CLIENT_SECRET };
  } else {
    missing.push(
      production
        ? "ENTRA_CERTIFICATE_PATH / ENTRA_PRIVATE_KEY_PATH"
        : "ENTRA_CLIENT_SECRET (development) or ENTRA_CERTIFICATE_PATH / ENTRA_PRIVATE_KEY_PATH",
    );
  }
  if (env.DATABASE_URL && env.STATE_ENCRYPTION_KEY) {
    if (!/^[a-f0-9]{64}$/i.test(env.STATE_ENCRYPTION_KEY))
      throw new Error("STATE_ENCRYPTION_KEY must be 64 hex characters.");
    base.databaseUrl = env.DATABASE_URL;
    base.encryptionKey = env.STATE_ENCRYPTION_KEY;
  } else if (production) {
    missing.push("DATABASE_URL / STATE_ENCRYPTION_KEY");
  }
  if (env.AI_LOCAL_MODEL) base.aiModel = env.AI_LOCAL_MODEL;
  base.signinsAudit = signinsAuditConfigFromEnv(env);
  base.directoryChanges = directoryChangesConfigFromEnv(env);
  base.delivery = await deliveryConfigFromEnv(env);
  if (env.TRUST_PROXY_HOPS) {
    if (!/^[0-3]$/.test(env.TRUST_PROXY_HOPS))
      throw new Error("TRUST_PROXY_HOPS must be a whole number from 0 to 3.");
    base.trustProxyHops = Number(env.TRUST_PROXY_HOPS);
  }
  if (env.HANDOFF_RETENTION_DAYS) {
    const days = /^\d{1,4}$/.test(env.HANDOFF_RETENTION_DAYS)
      ? Number(env.HANDOFF_RETENTION_DAYS)
      : Number.NaN;
    if (!(days >= 30 && days <= 3_650))
      throw new Error("HANDOFF_RETENTION_DAYS must be a whole number of days from 30 to 3650.");
    base.handoffRetentionDays = days;
  }
  if (env.PRISM_DATA_DIR) {
    if (/[\0\n\r]/.test(env.PRISM_DATA_DIR)) throw new Error("PRISM_DATA_DIR is not a valid path.");
    base.moduleDataDir = env.PRISM_DATA_DIR;
  }
  base.missingConfiguration = missing;
  if (missing.length === 0 && credential)
    base.identity = {
      tenantId: tenantId.toLowerCase(),
      clientId: clientId.toLowerCase(),
      origin,
      redirectUri: `${origin}/api/v1/auth/callback`,
      allowedActorIds,
      enabledFeatures,
      credential,
    };
  return base;
}

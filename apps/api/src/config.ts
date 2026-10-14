import { createPrivateKey, X509Certificate } from "node:crypto";
import { readFile } from "node:fs/promises";

export const supportedReadFeatures = {
  "directory.users.read": "User.Read.All",
  "user.licenses.read": "LicenseAssignment.Read.All",
  "audit.signins.read": "AuditLog.Read.All",
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
}

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
      credential = {
        clientCertificate: {
          thumbprintSha256: certificate.fingerprint256.replaceAll(":", ""),
          privateKey,
        },
      };
    } catch {
      throw new Error("Unable to load the matching Entra certificate and private key.");
    }
  } else if (env.ENTRA_CLIENT_SECRET && !production) {
    credential = { clientSecret: env.ENTRA_CLIENT_SECRET };
  } else {
    missing.push("ENTRA_CERTIFICATE_PATH / ENTRA_PRIVATE_KEY_PATH");
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

import { resolve } from "node:path";
import { parseArgs } from "node:util";
import {
  isUuid,
  type SupportedReadFeature,
  supportedReadFeatures,
} from "../../apps/api/src/config.ts";

export const DEFAULT_ORIGIN = "https://localhost:8443";
export const DEV_ORIGIN = "http://localhost:5173";
export const CALLBACK_PATH = "/api/v1/auth/callback";

export interface SetupOptions {
  tenant: string;
  name: string;
  origin: string;
  redirectUris: string[];
  features: SupportedReadFeature[];
  secretsDir: string;
  envFile: string;
  appId: string | null;
  yes: boolean;
  dryRun: boolean;
  rotateCertificate: boolean;
  prunePermissions: boolean;
  // Keep what an existing registration has beyond what Prism365 asks for, after reading the plan.
  keepOtherPermissions: boolean;
  keepOtherRedirectUris: boolean;
  keepOtherCertificates: boolean;
}

export class UsageError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "UsageError";
  }
}

export const usage = `Usage: npm run setup -- --tenant <tenant ID or domain> [options]

Registers Prism365 in your Microsoft Entra tenant, creates its certificate and writes .env.
It shows its plan and asks before changing anything.

  --tenant <id or domain>     The tenant to set up (required)
  --name <display name>       Name of the app registration (default Prism365)
  --origin <origin>           Where Prism365 runs (default ${DEFAULT_ORIGIN})
  --dev                       Also allow sign-in from the development server (${DEV_ORIGIN})
  --features <list>           Comma-separated read packs, or "all" (default: the core packs)
  --app-id <app ID>           Use this registration when several match or only its name does
  --secrets-dir <dir>         Where the certificate goes (default ./secrets)
  --env-file <file>           The settings file to write (default .env)
  --rotate-certificate        Make a new certificate even if one exists
  --prune-permissions         Remove delegated permissions Prism365 does not use
  --keep-other-permissions    Keep them instead (the consent link grants them too)
  --keep-other-redirect-uris  Keep redirect URIs outside --origin
  --keep-other-certificates   Keep certificates this setup did not make
  --dry-run                   Sign in, read and show the plan, then stop without changing anything
  --yes                       Apply the plan without asking
  --help                      Show this text`;

// The core packs: what the People page, the audits, the sign-in audit, the change feed, service
// health and Message center read. Each is a read permission most administrators already expect a
// reporting tool to have. The consent-batch packs (Conditional Access, risk, devices, app
// consents, usage reports) are asked for only when named, or with "all".
export const CORE_FEATURES = [
  "directory.users.read",
  "user.licenses.read",
  "tenant.licenses.read",
  "audit.signins.read",
  "audit.directory.read",
  "reports.authentication-methods.read",
  "service.health.read",
  "service.messages.read",
] as const satisfies readonly SupportedReadFeature[];

// directory.permission-grants.read is Directory.Read.All, which reads the whole directory, so even
// "all" leaves it out unless it is named too.
export const NAMED_ONLY_FEATURES: readonly SupportedReadFeature[] = [
  "directory.permission-grants.read",
];

// In the API's order, so the same list always reads the same way.
export function defaultFeatures(): SupportedReadFeature[] {
  const core = new Set<string>(CORE_FEATURES);
  return (Object.keys(supportedReadFeatures) as SupportedReadFeature[]).filter((feature) =>
    core.has(feature),
  );
}

const DOMAIN = /^(?=.{1,253}$)(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$/i;

export function parseTenant(value: string): string {
  const tenant = value.trim().toLowerCase();
  if (!isUuid(tenant) && !DOMAIN.test(tenant))
    throw new UsageError(
      "--tenant must be the tenant ID (a GUID) or one of its domains, such as contoso.onmicrosoft.com.",
    );
  return tenant;
}

// The same rule the server applies to PUBLIC_ORIGIN: an exact origin, HTTPS except on this
// computer.
export function parseOrigin(value: string): string {
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new UsageError("--origin must be an address such as https://prism365.example.");
  }
  if (url.origin !== value || url.username || url.password)
    throw new UsageError(
      "--origin must be an exact origin, with no path, such as https://prism365.example.",
    );
  if (
    url.protocol !== "https:" &&
    !(url.protocol === "http:" && ["localhost", "127.0.0.1"].includes(url.hostname))
  )
    throw new UsageError("--origin must use HTTPS, except for localhost.");
  return value;
}

export function parseFeatures(value: string): SupportedReadFeature[] {
  const names = value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
  const unknown = names.filter((name) => name !== "all" && !(name in supportedReadFeatures));
  if (unknown.length > 0)
    throw new UsageError(
      `Unknown read pack: ${unknown.join(", ")}. Choose from all, ${Object.keys(supportedReadFeatures).join(", ")}.`,
    );
  if (names.length === 0) throw new UsageError("--features needs at least one read pack.");
  const wanted = new Set(names);
  // In the API's order, so the same list always reads the same way.
  return (Object.keys(supportedReadFeatures) as SupportedReadFeature[]).filter(
    (feature) =>
      wanted.has(feature) || (wanted.has("all") && !NAMED_ONLY_FEATURES.includes(feature)),
  );
}

export function parseOptions(argv: string[], cwd: string): SetupOptions | "help" {
  let parsed: ReturnType<typeof read>;
  try {
    parsed = read(argv);
  } catch (error) {
    throw new UsageError(error instanceof Error ? error.message : "The options are not valid.");
  }
  const { values } = parsed;
  if (values.help) return "help";
  if (!values.tenant) throw new UsageError("--tenant is required.");
  const name = (values.name ?? "Prism365").trim();
  if (!name || name.length > 120 || /\p{Cc}/u.test(name))
    throw new UsageError("--name must be 1 to 120 characters.");
  const origin = parseOrigin(values.origin ?? DEFAULT_ORIGIN);
  const redirectUris = [`${origin}${CALLBACK_PATH}`];
  if (values.dev && origin !== DEV_ORIGIN) redirectUris.push(`${DEV_ORIGIN}${CALLBACK_PATH}`);
  if (values["prune-permissions"] && values["keep-other-permissions"])
    throw new UsageError("Choose --prune-permissions or --keep-other-permissions, not both.");
  const appId = values["app-id"]?.trim().toLowerCase() ?? null;
  if (appId !== null && !isUuid(appId))
    throw new UsageError("--app-id must be the registration's application (client) ID.");
  return {
    tenant: parseTenant(values.tenant),
    name,
    origin,
    redirectUris,
    features: values.features ? parseFeatures(values.features) : defaultFeatures(),
    secretsDir: resolve(cwd, values["secrets-dir"] ?? "secrets"),
    envFile: resolve(cwd, values["env-file"] ?? ".env"),
    appId,
    yes: Boolean(values.yes),
    dryRun: Boolean(values["dry-run"]),
    rotateCertificate: Boolean(values["rotate-certificate"]),
    prunePermissions: Boolean(values["prune-permissions"]),
    keepOtherPermissions: Boolean(values["keep-other-permissions"]),
    keepOtherRedirectUris: Boolean(values["keep-other-redirect-uris"]),
    keepOtherCertificates: Boolean(values["keep-other-certificates"]),
  };
}

function read(argv: string[]) {
  return parseArgs({
    args: argv,
    strict: true,
    allowPositionals: false,
    options: {
      tenant: { type: "string" },
      name: { type: "string" },
      origin: { type: "string" },
      dev: { type: "boolean" },
      features: { type: "string" },
      "app-id": { type: "string" },
      "secrets-dir": { type: "string" },
      "env-file": { type: "string" },
      "rotate-certificate": { type: "boolean" },
      "prune-permissions": { type: "boolean" },
      "keep-other-permissions": { type: "boolean" },
      "keep-other-redirect-uris": { type: "boolean" },
      "keep-other-certificates": { type: "boolean" },
      "dry-run": { type: "boolean" },
      yes: { type: "boolean" },
      help: { type: "boolean" },
    },
  });
}

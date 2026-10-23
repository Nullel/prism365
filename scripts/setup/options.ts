import { resolve } from "node:path";
import { parseArgs } from "node:util";
import { builtInAudits } from "@prism365/audits";
import {
  isUuid,
  type SupportedReadFeature,
  supportedReadFeatures,
} from "../../apps/api/src/config.ts";
import { createModuleRegistry } from "../../apps/api/src/modules.ts";

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

  --tenant <id or domain>   The tenant to set up (required)
  --name <display name>     Name of the app registration (default Prism365)
  --origin <origin>         Where Prism365 runs (default ${DEFAULT_ORIGIN})
  --dev                     Also allow sign-in from the development server (${DEV_ORIGIN})
  --features <list>         Comma-separated read packs (default: what the audits and modules use)
  --app-id <app ID>         Use this registration when several could match
  --secrets-dir <dir>       Where the certificate goes (default ./secrets)
  --env-file <file>         The settings file to write (default .env)
  --rotate-certificate      Make a new certificate even if one exists
  --prune-permissions       Remove permissions on the registration that Prism365 does not use
  --dry-run                 Sign in, read and show the plan, then stop without changing anything
  --yes                     Apply the plan without asking
  --help                    Show this text`;

// What the People page and investigations read (profiles, licence details, sign-ins), plus every
// pack a built-in audit or module uses, in the order the API lists them.
export function defaultFeatures(): SupportedReadFeature[] {
  const wanted = new Set<string>([
    "directory.users.read",
    "user.licenses.read",
    "audit.signins.read",
  ]);
  for (const audit of builtInAudits)
    for (const pack of [...audit.packs, ...audit.optionalPacks]) wanted.add(pack);
  for (const module of createModuleRegistry().list())
    for (const pack of module.packs) wanted.add(pack);
  return (Object.keys(supportedReadFeatures) as SupportedReadFeature[]).filter((feature) =>
    wanted.has(feature),
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
  const unknown = names.filter((name) => !(name in supportedReadFeatures));
  if (unknown.length > 0)
    throw new UsageError(
      `Unknown read pack: ${unknown.join(", ")}. Choose from ${Object.keys(supportedReadFeatures).join(", ")}.`,
    );
  if (names.length === 0) throw new UsageError("--features needs at least one read pack.");
  return [...new Set(names)] as SupportedReadFeature[];
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
      "dry-run": { type: "boolean" },
      yes: { type: "boolean" },
      help: { type: "boolean" },
    },
  });
}

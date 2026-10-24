import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach } from "vitest";
import type { SetupOptions } from "../options.ts";
import { runSetup, type SetupDependencies } from "../setup.ts";
import { type FakeTenant, TENANT_ID } from "./fake-graph.ts";

// Shared by the setup's tests: a temporary workspace per test, the options a run gets, and a run
// against a fake tenant with no network, no prompts and no waiting.

export const envExample = join(import.meta.dirname, "../../../.env.example");
const directories: string[] = [];

afterEach(async () => {
  await Promise.all(directories.splice(0).map((dir) => rm(dir, { recursive: true, force: true })));
});

export async function workspace() {
  const dir = await mkdtemp(join(tmpdir(), "prism365-setup-"));
  directories.push(dir);
  return dir;
}

export function optionsFor(dir: string, overrides: Partial<SetupOptions> = {}): SetupOptions {
  return {
    tenant: "contoso.example",
    name: "Prism365",
    origin: "https://localhost:8443",
    redirectUris: ["https://localhost:8443/api/v1/auth/callback"],
    features: ["directory.users.read", "user.licenses.read", "audit.signins.read"],
    secretsDir: join(dir, "secrets"),
    envFile: join(dir, ".env"),
    appId: null,
    yes: true,
    dryRun: false,
    rotateCertificate: false,
    prunePermissions: false,
    keepOtherPermissions: false,
    keepOtherRedirectUris: false,
    keepOtherCertificates: false,
    keepOtherOwners: false,
    ...overrides,
  };
}

export function run(
  tenant: FakeTenant,
  options: SetupOptions,
  dir: string,
  output: string[] = [],
  extra: Partial<SetupDependencies> = {},
) {
  return runSetup(options, {
    fetch: tenant.fetch,
    signIn: async () => ({ tenantId: TENANT_ID, getToken: async () => "synthetic-token" }),
    confirm: async () => true,
    print: (text) => output.push(text),
    envExample,
    cwd: dir,
    sleep: async () => {},
    // No ACL tools by default: the tests run the same on every system.
    acl: { read: async () => null, apply: async () => false },
    ...extra,
  });
}

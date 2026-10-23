import {
  createPrivateKey,
  createPublicKey,
  generateKeyPairSync,
  X509Certificate,
} from "node:crypto";
import {
  chmod,
  link,
  lstat,
  mkdir,
  mkdtemp,
  readdir,
  readFile,
  rename,
  rm,
  stat,
  symlink,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { supportedReadFeatures } from "../../../apps/api/src/config.ts";
import { modeDecision, parseGetfacl, setfaclCommand } from "../acl.ts";
import {
  CERTIFICATE_DAYS,
  certificateFiles,
  certificateRequest,
  commonNameFor,
  inspectSecrets,
  interruptedRotation,
  keyProblem,
  newCredential,
  promotePending,
  selfSignedCertificate,
  writeCredential,
  writePending,
} from "../certificate.ts";
import {
  adminObjectIds,
  EnvFileError,
  MANAGED_KEYS,
  type ManagedValues,
  mergeEnv,
  readEnv,
  writeEnvFile,
} from "../env-file.ts";
import { explainGraphError, SetupGraph, SetupGraphError, terminalText } from "../graph.ts";
import {
  hasCertificate,
  KeyCredentialSafetyError,
  mergeKeyCredentials,
  type NewKeyCredential,
  verifyKeyCredentials,
} from "../key-credentials.ts";
import { defaultFeatures, parseFeatures, parseOptions, UsageError } from "../options.ts";
import {
  GRAPH_RESOURCE_APP_ID,
  graphResourceAccess,
  mergeRequiredResourceAccess,
  requestedPermissions,
  UnknownPermissionError,
} from "../permissions.ts";

const directories: string[] = [];
afterEach(async () => {
  await Promise.all(directories.splice(0).map((dir) => rm(dir, { recursive: true, force: true })));
});
async function workspace() {
  const dir = await mkdtemp(join(tmpdir(), "prism365-setup-unit-"));
  directories.push(dir);
  return dir;
}
const mode = async (path: string) => (await stat(path)).mode & 0o777;

// Graph's service principal as the fake publishes it: invented IDs, real permission names.
const published = [
  { id: "00000000-0000-4000-8000-00000000000a", value: "openid", isEnabled: true },
  { id: "00000000-0000-4000-8000-00000000000b", value: "profile", isEnabled: true },
  { id: "00000000-0000-4000-8000-00000000000c", value: "offline_access", isEnabled: true },
  { id: "00000000-0000-4000-8000-00000000000d", value: "User.Read", isEnabled: true },
  { id: "00000000-0000-4000-8000-00000000000e", value: "User.Read.All", isEnabled: true },
  { id: "00000000-0000-4000-8000-00000000000f", value: "AuditLog.Read.All", isEnabled: true },
  { id: "00000000-0000-4000-8000-000000000010", value: "Policy.Read.All", isEnabled: false },
];

describe("permissions", () => {
  it("lists each permission once with the packs it serves", () => {
    expect(
      requestedPermissions([
        "audit.signins.read",
        "directory.users.read",
        "reports.authentication-methods.read",
      ]),
    ).toEqual([
      { name: "openid", features: [] },
      { name: "profile", features: [] },
      { name: "offline_access", features: [] },
      { name: "User.Read", features: [] },
      {
        name: "AuditLog.Read.All",
        features: ["audit.signins.read", "reports.authentication-methods.read"],
      },
      { name: "User.Read.All", features: ["directory.users.read"] },
    ]);
  });

  it("maps permission names to the scope IDs Graph publishes", () => {
    const names = requestedPermissions(["directory.users.read", "audit.signins.read"]).map(
      (permission) => permission.name,
    );
    expect(graphResourceAccess(names, published)).toEqual({
      resourceAppId: GRAPH_RESOURCE_APP_ID,
      resourceAccess: ["a", "b", "c", "d", "e", "f"].map((suffix) => ({
        id: `00000000-0000-4000-8000-00000000000${suffix}`,
        type: "Scope",
      })),
    });
  });

  it("fails clearly on names Graph does not publish or has disabled", () => {
    expect(() =>
      graphResourceAccess(["User.Read", "Nope.Read.All", "Policy.Read.All"], published),
    ).toThrow(UnknownPermissionError);
    try {
      graphResourceAccess(["Nope.Read.All", "Policy.Read.All"], published);
    } catch (error) {
      expect((error as UnknownPermissionError).names).toEqual(["Nope.Read.All", "Policy.Read.All"]);
      expect((error as Error).message).toContain("Nope.Read.All, Policy.Read.All");
    }
  });

  const wanted = {
    resourceAppId: GRAPH_RESOURCE_APP_ID,
    resourceAccess: [
      { id: "A", type: "Scope" },
      { id: "B", type: "Scope" },
    ],
  };
  const existing = [
    {
      resourceAppId: GRAPH_RESOURCE_APP_ID,
      resourceAccess: [
        { id: "b", type: "Scope" },
        { id: "X", type: "Scope" },
        { id: "A", type: "Role" },
      ],
    },
    { resourceAppId: "other", resourceAccess: [{ id: "Y", type: "Role" }] },
  ];

  it("adds what is missing and keeps every permission already on the registration", () => {
    const merged = mergeRequiredResourceAccess(existing, wanted, false);
    expect(merged.value).toEqual([
      {
        resourceAppId: GRAPH_RESOURCE_APP_ID,
        resourceAccess: [...(existing[0]?.resourceAccess ?? []), { id: "A", type: "Scope" }],
      },
      existing[1],
    ]);
    expect(merged.added).toEqual([{ id: "A", type: "Scope" }]);
    expect(merged.removed).toEqual([]);
    expect(mergeRequiredResourceAccess([], wanted, false).value).toEqual([wanted]);
  });

  it("removes the rest only when pruning", () => {
    const merged = mergeRequiredResourceAccess(existing, wanted, true);
    expect(merged.value).toEqual([wanted]);
    expect(merged.added).toEqual([{ id: "A", type: "Scope" }]);
    expect(merged.removed).toEqual([
      { resourceAppId: GRAPH_RESOURCE_APP_ID, access: { id: "X", type: "Scope" } },
      { resourceAppId: GRAPH_RESOURCE_APP_ID, access: { id: "A", type: "Role" } },
      { resourceAppId: "other", access: { id: "Y", type: "Role" } },
    ]);
  });
});

describe(".env merge", () => {
  const values: ManagedValues = {
    DATA_MODE: "live",
    ENTRA_TENANT_ID: "11111111-1111-4111-8111-111111111111",
    ENTRA_CLIENT_ID: "33333333-3333-4333-8333-333333333333",
    ENTRA_ADMIN_OBJECT_IDS: "22222222-2222-4222-8222-222222222222",
    ENABLED_READ_FEATURES: "directory.users.read",
    PUBLIC_ORIGIN: "https://localhost:8443",
    ENTRA_CERTIFICATE_PATH: "/srv/prism365/secrets/entra-public.pem",
    ENTRA_PRIVATE_KEY_PATH: "/srv/prism365/secrets/entra-private.pem",
  };

  it("changes managed keys in place and keeps unknown keys, comments and order", () => {
    const source = [
      "# My notes about this install",
      "HOST=127.0.0.1",
      "DATA_MODE=demo # was the demo",
      "UNKNOWN_SETTING=keep me",
      "",
      "# ENTRA_TENANT_ID=",
      'export PUBLIC_ORIGIN="http://localhost:5173"',
      "SIGNINS_TIME_ZONE=America/Chicago",
      "DATA_MODE=demo",
      "",
    ].join("\n");
    const merged = mergeEnv(source, values);
    expect(merged.text.split("\n")).toEqual([
      "# My notes about this install",
      "HOST=127.0.0.1",
      "# was the demo",
      "DATA_MODE=live",
      "UNKNOWN_SETTING=keep me",
      "",
      "# ENTRA_TENANT_ID=",
      `ENTRA_TENANT_ID=${values.ENTRA_TENANT_ID}`,
      "PUBLIC_ORIGIN=https://localhost:8443",
      "SIGNINS_TIME_ZONE=America/Chicago",
      "DATA_MODE=live",
      "",
      "# Set by npm run setup.",
      `ENTRA_CLIENT_ID=${values.ENTRA_CLIENT_ID}`,
      `ENTRA_ADMIN_OBJECT_IDS=${values.ENTRA_ADMIN_OBJECT_IDS}`,
      "ENABLED_READ_FEATURES=directory.users.read",
      `ENTRA_CERTIFICATE_PATH=${values.ENTRA_CERTIFICATE_PATH}`,
      `ENTRA_PRIVATE_KEY_PATH=${values.ENTRA_PRIVATE_KEY_PATH}`,
      "",
    ]);
    expect(merged.changes.map((change) => change.key)).toEqual([...MANAGED_KEYS]);
    expect(merged.changes[0]).toEqual({ key: "DATA_MODE", before: "demo", after: "live" });
    expect(readEnv(merged.text).get("UNKNOWN_SETTING")).toBe("keep me");
    expect(mergeEnv(merged.text, values)).toEqual({ text: merged.text, changes: [] });
  });

  it("refuses a value that runs over several lines instead of editing inside it", () => {
    const text = 'NOTE="first line\nENTRA_CLIENT_ID=inside-a-quoted-value\nlast line"\nOTHER=1\n';
    expect(() => mergeEnv(text, values)).toThrow(EnvFileError);
    expect(() => mergeEnv(text, values)).toThrow("NOTE");
  });

  it("writes managed keys as plain KEY=value, keeping an inline comment on its own line", () => {
    const merged = mergeEnv(
      [
        "ENTRA_CLIENT_ID = old-id",
        "export DATA_MODE=demo",
        `PUBLIC_ORIGIN="${values.PUBLIC_ORIGIN}"`,
        "ENTRA_TENANT_ID=old # the pilot tenant",
        "",
      ].join("\n"),
      values,
    );
    expect(merged.text.split("\n").slice(0, 5)).toEqual([
      `ENTRA_CLIENT_ID=${values.ENTRA_CLIENT_ID}`,
      "DATA_MODE=live",
      `PUBLIC_ORIGIN=${values.PUBLIC_ORIGIN}`,
      "# the pilot tenant",
      `ENTRA_TENANT_ID=${values.ENTRA_TENANT_ID}`,
    ]);
    // What scripts/docker-live.sh reads: everything after "KEY=" on a line that starts with it.
    const sed = (key: string) =>
      merged.text
        .split("\n")
        .filter((line) => line.startsWith(`${key}=`))
        .map((line) => line.slice(key.length + 1))
        .at(-1);
    for (const key of MANAGED_KEYS) expect(sed(key), key).toBe(values[key]);
  });

  it("keeps Windows line endings and a file without a final newline", () => {
    const merged = mergeEnv("HOST=127.0.0.1\r\nDATA_MODE=demo", values);
    expect(merged.text).toMatch(
      /^HOST=127\.0\.0\.1\r\nDATA_MODE=live\r\n\r\n# Set by npm run setup\.\r\n/,
    );
    expect(merged.text).not.toMatch(/[^\r]\n/);
  });

  it("adds the signed-in administrator once and keeps the others", () => {
    const admin = "22222222-2222-4222-8222-222222222222";
    expect(adminObjectIds(undefined, admin)).toBe(admin);
    expect(
      adminObjectIds(`AAAAAAAA-0000-4000-8000-000000000001, ${admin.toUpperCase()}`, admin),
    ).toBe(`aaaaaaaa-0000-4000-8000-000000000001,${admin}`);
    expect(adminObjectIds("aaaaaaaa-0000-4000-8000-000000000001", admin)).toBe(
      `aaaaaaaa-0000-4000-8000-000000000001,${admin}`,
    );
  });

  it("refuses values the file could not hold", () => {
    expect(() =>
      mergeEnv("", { ...values, ENTRA_CERTIFICATE_PATH: "/srv/my#certs/entra-public.pem" }),
    ).toThrow("ENTRA_CERTIFICATE_PATH");
    expect(() =>
      mergeEnv("", { ...values, PUBLIC_ORIGIN: "https://x.example\nDATA_MODE=demo" }),
    ).toThrow("PUBLIC_ORIGIN");
  });

  it("writes the file owner-only in one step, through a symbolic link", async () => {
    const dir = await workspace();
    const file = join(dir, ".env");
    await writeFile(file, "OLD=1\n");
    await chmod(file, 0o644);
    await writeEnvFile(file, "NEW=1\n");
    expect(await readFile(file, "utf8")).toBe("NEW=1\n");
    expect(await mode(file)).toBe(0o600);
    expect(await readdir(dir)).toEqual([".env"]);

    await mkdir(join(dir, "real"));
    await writeFile(join(dir, "real/env"), "A=1\n");
    await symlink(join(dir, "real/env"), join(dir, "linked.env"));
    await writeEnvFile(join(dir, "linked.env"), "B=2\n");
    expect(await readFile(join(dir, "real/env"), "utf8")).toBe("B=2\n");
    expect((await lstat(join(dir, "linked.env"))).isSymbolicLink()).toBe(true);
  });
});

describe("certificate credentials", () => {
  const certificate = (pem: string) => new X509Certificate(pem);
  const keyOf = (pem: string) => certificate(pem).raw.toString("base64");
  const added = (key: string): NewKeyCredential => ({
    type: "AsymmetricX509Cert",
    usage: "Verify",
    key,
    displayName: "CN=Prism365 contoso.example",
    startDateTime: "2026-10-22T00:00:00.000Z",
    endDateTime: "2027-10-22T00:00:00.000Z",
  });
  const pair = newCredential(certificateRequest("contoso.example", new Date()));
  const other = newCredential(certificateRequest("other.example", new Date()));
  const existing = {
    keyId: "44444444-4444-4444-8444-444444444444",
    customKeyIdentifier: "C0FFEE",
    displayName: "CN=Other",
    type: "AsymmetricX509Cert",
    usage: "Verify",
    key: keyOf(other.certificatePem),
    endDateTime: "2027-01-01T00:00:00Z",
  };

  it("sends existing certificates back unchanged, followed by the new one", () => {
    const merged = mergeKeyCredentials([existing], added(keyOf(pair.certificatePem)));
    expect(merged).toHaveLength(2);
    expect(merged[0]).toBe(existing);
    expect(merged[1]).toEqual(added(keyOf(pair.certificatePem)));
  });

  it("refuses when an existing key value is missing", () => {
    for (const key of [null, undefined, "", "not base64!"]) {
      const partial = { ...existing, key };
      expect(() => mergeKeyCredentials([partial], added(keyOf(pair.certificatePem)))).toThrow(
        KeyCredentialSafetyError,
      );
    }
    expect(() =>
      mergeKeyCredentials([{ key: existing.key }], added(keyOf(pair.certificatePem))),
    ).toThrow("readable key ID");
  });

  it("catches a key that the update dropped, or a new key that never arrived", () => {
    const newKey = keyOf(pair.certificatePem);
    const after = [existing, { ...added(newKey), keyId: "55555555-5555-4555-8555-555555555555" }];
    expect(() => verifyKeyCredentials([existing], after, newKey)).not.toThrow();
    expect(() => verifyKeyCredentials([existing], after.slice(1), newKey)).toThrow(existing.keyId);
    expect(() => verifyKeyCredentials([existing], [existing], newKey)).toThrow(
      "the new certificate is not on the registration",
    );
  });

  it("recognises a certificate already on the registration by its bytes", () => {
    expect(hasCertificate([existing], keyOf(other.certificatePem))).toBe(true);
    expect(hasCertificate([existing], keyOf(pair.certificatePem))).toBe(false);
  });

  it("builds a valid self-signed RSA certificate for a year at most", () => {
    const now = new Date("2026-10-22T12:34:56.789Z");
    const made = newCredential(certificateRequest("contoso.example", now));
    const parsed = certificate(made.certificatePem);
    expect(parsed.subject).toBe("CN=Prism365 contoso.example");
    expect(parsed.issuer).toBe(parsed.subject);
    expect(parsed.validFromDate.toISOString()).toBe("2026-10-22T12:34:56.000Z");
    expect(parsed.validToDate.getTime() - parsed.validFromDate.getTime()).toBe(
      CERTIFICATE_DAYS * 86_400_000,
    );
    expect(parsed.ca).toBe(false);
    expect(parsed.publicKey.asymmetricKeyType).toBe("rsa");
    expect(parsed.publicKey.asymmetricKeyDetails?.modulusLength).toBe(2048);
    expect(parsed.checkPrivateKey(createPrivateKey(made.privateKeyPem))).toBe(true);
    expect(parsed.verify(createPublicKey(made.privateKeyPem))).toBe(true);
    expect(made.privateKeyPem).toMatch(/^-----BEGIN PRIVATE KEY-----\n/);
  });

  it("switches to GeneralizedTime from 2050 and bounds the common name", () => {
    const parsed = certificate(
      newCredential(certificateRequest("contoso.example", new Date("2049-12-01T00:00:00Z")))
        .certificatePem,
    );
    expect(parsed.validToDate.toISOString()).toBe("2050-12-01T00:00:00.000Z");
    expect(commonNameFor("a".repeat(200))).toHaveLength(64);
  });

  it("refuses keys that are not RSA", () => {
    const { privateKey } = generateKeyPairSync("ec", { namedCurve: "P-256" });
    expect(() =>
      selfSignedCertificate(privateKey, certificateRequest("contoso.example", new Date())),
    ).toThrow("RSA");
  });

  it("writes a new pair owner-only, never overwrites a key, and leaves an existing folder's mode alone", async () => {
    const dir = await workspace();
    const files = certificateFiles(join(dir, "secrets"));
    await writeCredential(files, pair);
    expect(await mode(files.directory)).toBe(0o700);
    expect(await mode(files.privateKey)).toBe(0o600);
    expect(await mode(files.certificate)).toBe(0o644);
    expect((await inspectSecrets(files)).credential.state).toBe("pair");

    await expect(writeCredential(files, other)).rejects.toThrow("--rotate-certificate");
    expect(await readFile(files.privateKey, "utf8")).toBe(pair.privateKeyPem);
    expect(await readFile(files.certificate, "utf8")).toBe(pair.certificatePem);
    expect(await readdir(files.directory)).toEqual(["entra-private.pem", "entra-public.pem"]);

    // An existing folder may carry an access list; only an explicit plan step changes it.
    const shared = certificateFiles(join(dir, "shared"));
    await mkdir(shared.directory, { mode: 0o750 });
    await chmod(shared.directory, 0o750);
    await writeCredential(shared, pair, { privateKey: 0o640, certificate: 0o644 });
    expect(await mode(shared.directory)).toBe(0o750);
    expect(await mode(shared.privateKey)).toBe(0o640);
  });

  it("rotates through .new files, keeping the old pair as .previous, and can finish an interrupted move", async () => {
    const dir = await workspace();
    const files = certificateFiles(join(dir, "secrets"));
    await writeCredential(files, pair);
    await writePending(files, other);
    // Until it is moved, the working pair is untouched and the new one is only pending.
    expect(await readFile(files.privateKey, "utf8")).toBe(pair.privateKeyPem);
    let inspected = await inspectSecrets(files);
    expect(inspected.credential.state).toBe("pair");
    expect(inspected.pending?.certificate.raw.equals(certificate(other.certificatePem).raw)).toBe(
      true,
    );
    expect(await mode(files.pendingKey)).toBe(0o600);

    // Interrupted after the key moved but before the certificate did.
    await link(files.privateKey, files.previousKey);
    await rename(files.pendingKey, files.privateKey);
    inspected = await inspectSecrets(files);
    expect(inspected.credential.state).toBe("mismatch");
    expect(interruptedRotation(inspected)).toBe(true);

    await promotePending(files);
    inspected = await inspectSecrets(files);
    expect(inspected.credential.state).toBe("pair");
    expect(await readFile(files.privateKey, "utf8")).toBe(other.privateKeyPem);
    expect(await readFile(files.certificate, "utf8")).toBe(other.certificatePem);
    expect(await readFile(files.previousKey, "utf8")).toBe(pair.privateKeyPem);
    expect(await readFile(files.previousCertificate, "utf8")).toBe(pair.certificatePem);
    expect(inspected.pendingFiles).toBe(false);
    expect(inspected.previousFiles).toBe(true);
    expect(await mode(files.previousKey)).toBe(0o600);
    // Every name ends in .pem, so .gitignore, .dockerignore and the Dockerfile's guard see them.
    expect((await readdir(files.directory)).every((name) => name.endsWith(".pem"))).toBe(true);
  });

  it("names the key problems MSAL and Microsoft would reject", () => {
    expect(keyProblem(createPrivateKey(pair.privateKeyPem))).toBeNull();
    expect(keyProblem(generateKeyPairSync("ec", { namedCurve: "P-256" }).privateKey)).toContain(
      "EC key",
    );
    expect(keyProblem(generateKeyPairSync("rsa", { modulusLength: 1024 }).privateKey)).toContain(
      "1024-bit",
    );
  });

  it("tells apart every state the secrets directory can be in", async () => {
    const dir = await workspace();
    const files = certificateFiles(join(dir, "secrets"));
    expect(await inspectSecrets(files)).toEqual({
      directoryMode: null,
      credential: { state: "none" },
      pending: null,
      pendingFiles: false,
      previous: null,
      previousFiles: false,
    });
    await mkdir(files.directory, { mode: 0o700 });
    await writeFile(files.certificate, pair.certificatePem);
    expect((await inspectSecrets(files)).credential.state).toBe("certificateOnly");
    await writeFile(files.privateKey, other.privateKeyPem, { mode: 0o644 });
    expect((await inspectSecrets(files)).credential.state).toBe("mismatch");
    await writeFile(files.privateKey, pair.privateKeyPem);
    const inspected = await inspectSecrets(files);
    expect(inspected.credential).toMatchObject({ state: "pair", privateKeyMode: 0o644 });
    await rm(files.certificate);
    expect((await inspectSecrets(files)).credential.state).toBe("keyOnly");
    await writeFile(files.privateKey, "not a key");
    expect((await inspectSecrets(files)).credential).toEqual({
      state: "unreadable",
      file: "privateKey",
    });
  });
});

describe("access control lists", () => {
  it("reads only what is more than the mode bits from getfacl", () => {
    expect(
      parseGetfacl(
        [
          "user::rw-",
          "user:1000:r--",
          "group::---",
          "mask::r--",
          "other::---",
          "default:user::rwx",
          "default:user:1000:r--  #effective:r--",
          "default:group::---",
        ].join("\n"),
      ),
    ).toEqual({ entries: ["user:1000:r--", "mask::r--"], defaults: ["user:1000:r--"] });
    expect(parseGetfacl("user::rw-\ngroup::r--\nother::---\n")).toEqual({
      entries: [],
      defaults: [],
    });
  });

  it("never tightens what an access list or a group may depend on", () => {
    const acl = { entries: ["user:1000:r--", "mask::r--"], defaults: [] };
    const plain = { entries: [], defaults: [] };
    expect(modeDecision(0o600, null)).toBe("fine");
    expect(modeDecision(0o640, acl)).toBe("keep-acl");
    expect(modeDecision(0o640, null)).toBe("keep-group");
    expect(modeDecision(0o640, plain)).toBe("tighten");
    expect(modeDecision(0o644, null)).toBe("tighten");
    expect(setfaclCommand("/srv/prism 365/secrets/entra-private.pem", acl.entries)).toBe(
      "setfacl -m user:1000:r-- '/srv/prism 365/secrets/entra-private.pem'",
    );
  });
});

describe("options", () => {
  const cwd = "/srv/prism365";

  it("defaults to the core packs, and asks for the consent-batch packs only when named", () => {
    expect(defaultFeatures()).toEqual([
      "directory.users.read",
      "user.licenses.read",
      "tenant.licenses.read",
      "audit.signins.read",
      "reports.authentication-methods.read",
      "service.health.read",
      "service.messages.read",
      "audit.directory.read",
    ]);
    for (const feature of defaultFeatures()) expect(feature in supportedReadFeatures).toBe(true);
    const all = parseFeatures("all");
    expect(all).toEqual(
      Object.keys(supportedReadFeatures).filter((id) => id !== "directory.permission-grants.read"),
    );
    expect(parseFeatures("all,directory.permission-grants.read")).toEqual(
      Object.keys(supportedReadFeatures),
    );
    expect(parseFeatures("identity.risk.read,directory.users.read")).toEqual([
      "directory.users.read",
      "identity.risk.read",
    ]);
  });

  it("reads the options and resolves paths from the working directory", () => {
    const options = parseOptions(
      [
        "--tenant",
        "Contoso.Example",
        "--dev",
        "--features",
        "audit.signins.read,audit.signins.read",
      ],
      cwd,
    );
    expect(options).toMatchObject({
      tenant: "contoso.example",
      name: "Prism365",
      origin: "https://localhost:8443",
      redirectUris: [
        "https://localhost:8443/api/v1/auth/callback",
        "http://localhost:5173/api/v1/auth/callback",
      ],
      features: ["audit.signins.read"],
      secretsDir: "/srv/prism365/secrets",
      envFile: "/srv/prism365/.env",
      yes: false,
      dryRun: false,
    });
    expect(parseOptions(["--help"], cwd)).toBe("help");
  });

  it("rejects what Prism365 could not use", () => {
    for (const argv of [
      [],
      ["--tenant", "common"],
      ["--tenant", "contoso.example", "--features", "mail.read"],
      ["--tenant", "contoso.example", "--origin", "http://prism365.example"],
      ["--tenant", "contoso.example", "--origin", "https://prism365.example/app"],
      ["--tenant", "contoso.example", "--app-id", "not-a-guid"],
      ["--tenant", "contoso.example", "--unknown"],
    ])
      expect(() => parseOptions(argv, cwd), argv.join(" ")).toThrow(UsageError);
  });
});

describe("the setup's Graph client", () => {
  it("reports Microsoft's code and request ID, never the provider's message or the token", async () => {
    const graph = new SetupGraph({
      fetch: async () =>
        Response.json(
          {
            error: {
              code: "Authorization_RequestDenied",
              message: "Insufficient privileges for alex@example.invalid",
            },
          },
          { status: 403, headers: { "request-id": "abc-123" } },
        ),
      getToken: async () => "synthetic-token",
    });
    const failure = await graph.patch("applications/x", {}).catch((error: unknown) => error);
    expect(failure).toBeInstanceOf(SetupGraphError);
    const message = explainGraphError("Updating the app registration", failure as SetupGraphError);
    expect(message).toContain("HTTP 403 (Authorization_RequestDenied)");
    expect(message).toContain("Application Administrator");
    expect(message).toContain("abc-123");
    expect(message).not.toContain("alex@example.invalid");
    expect(message).not.toContain("synthetic-token");
  });

  it("waits out throttling and refuses hosts other than Graph", async () => {
    const answers = [
      new Response("{}", { status: 429, headers: { "Retry-After": "2" } }),
      Response.json({ ok: true }),
    ];
    const waits: number[] = [];
    const graph = new SetupGraph({
      fetch: async () => answers.shift() ?? new Response(null, { status: 500 }),
      getToken: async () => "synthetic-token",
      sleep: async (milliseconds) => {
        waits.push(milliseconds);
      },
    });
    const { z } = await import("zod");
    expect(await graph.get("me", z.object({ ok: z.boolean() }))).toEqual({ ok: true });
    expect(waits).toEqual([2_000]);
    await expect(graph.get("https://evil.example/v1.0/me", z.unknown())).rejects.toThrow(
      "Microsoft Graph v1.0",
    );
  });

  it("strips control characters from provider text bound for the terminal", () => {
    expect(terminalText("Sign in\u001b[2J at ‮example")).toBe("Sign in [2J at  example");
  });
});

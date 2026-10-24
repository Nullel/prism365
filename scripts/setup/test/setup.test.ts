import { createPrivateKey, X509Certificate } from "node:crypto";
import { readdir, readFile, stat, unlink, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { certificateRequest, newCredential } from "../certificate.ts";
import { KeyCredentialSafetyError } from "../key-credentials.ts";
import { GRAPH_RESOURCE_APP_ID } from "../permissions.ts";
import { SETUP_TAG } from "../plan.ts";
import { CLI_TOOLS_REMINDER, runSetup, SetupError } from "../setup.ts";
import { ADMIN_ID, error, FakeTenant, OTHER_API, scopeIds, TENANT_ID } from "./fake-graph.ts";
import { envExample, optionsFor, run, workspace } from "./harness.ts";

const mode = async (path: string) => (await stat(path)).mode & 0o777;
const graphScopes = (tenant: FakeTenant) =>
  tenant.applications[0]?.requiredResourceAccess.find(
    (entry) => entry.resourceAppId === GRAPH_RESOURCE_APP_ID,
  )?.resourceAccess ?? [];
const scope = (name: string) => ({ id: scopeIds.get(name), type: "Scope" });

describe("npm run setup against a fake tenant", () => {
  it("creates the registration, service principal, certificate and .env on a first run", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const output: string[] = [];
    expect(await run(tenant, optionsFor(dir), dir, output)).toBe("applied");

    expect(tenant.writes.map((request) => `${request.method} ${request.path}`)).toEqual([
      "POST applications",
      "POST servicePrincipals",
      `PATCH applications/${tenant.applications[0]?.id}`,
      `PATCH applications/${tenant.applications[0]?.id}`,
    ]);
    const application = tenant.applications[0];
    expect(tenant.applications).toHaveLength(1);
    // Created with the label and a marker for this run only; the marker comes off at the end.
    const created = tenant.writes[0]?.body as { tags: string[] };
    expect(created.tags[0]).toBe(SETUP_TAG);
    expect(created.tags[1]).toMatch(/^prism365-setup-run-[0-9a-f]{16}$/);
    expect(tenant.writes[3]?.body).toEqual({ tags: [SETUP_TAG] });
    expect(application?.tags).toEqual([SETUP_TAG]);
    expect(application?.signInAudience).toBe("AzureADMyOrg");
    expect(application?.web).toEqual({
      redirectUris: ["https://localhost:8443/api/v1/auth/callback"],
      implicitGrantSettings: { enableIdTokenIssuance: false, enableAccessTokenIssuance: false },
    });
    expect(graphScopes(tenant)).toEqual(
      [
        "openid",
        "profile",
        "offline_access",
        "User.Read",
        "User.Read.All",
        "LicenseAssignment.Read.All",
        "AuditLog.Read.All",
      ].map(scope),
    );
    expect(tenant.servicePrincipals.map((principal) => principal.appId)).toEqual([
      application?.appId,
    ]);

    const secrets = join(dir, "secrets");
    expect(await mode(secrets)).toBe(0o700);
    expect(await mode(join(secrets, "entra-private.pem"))).toBe(0o600);
    expect(await mode(join(secrets, "entra-public.pem"))).toBe(0o644);
    const certificate = new X509Certificate(await readFile(join(secrets, "entra-public.pem")));
    expect(certificate.subject).toBe("CN=Prism365 contoso.example");
    expect(
      certificate.checkPrivateKey(
        createPrivateKey(await readFile(join(secrets, "entra-private.pem"))),
      ),
    ).toBe(true);
    expect(application?.keyCredentials).toEqual([
      expect.objectContaining({
        type: "AsymmetricX509Cert",
        usage: "Verify",
        key: certificate.raw.toString("base64"),
        displayName: "CN=Prism365 contoso.example",
      }),
    ]);

    const env = await readFile(join(dir, ".env"), "utf8");
    expect(await mode(join(dir, ".env"))).toBe(0o600);
    expect(env).toContain(
      "# Copy to .env. Keep credentials and tenant-specific values out of Git.",
    );
    expect(env).toMatch(/^DATA_MODE=live$/m);
    expect(env).toMatch(new RegExp(`^ENTRA_TENANT_ID=${TENANT_ID}$`, "m"));
    expect(env).toMatch(new RegExp(`^ENTRA_CLIENT_ID=${application?.appId}$`, "m"));
    expect(env).toMatch(new RegExp(`^ENTRA_ADMIN_OBJECT_IDS=${ADMIN_ID}$`, "m"));
    expect(env).toMatch(
      /^ENABLED_READ_FEATURES=directory\.users\.read,user\.licenses\.read,audit\.signins\.read$/m,
    );
    expect(env).toContain(`ENTRA_PRIVATE_KEY_PATH=${join(secrets, "entra-private.pem")}`);
    expect(env).not.toMatch(/^ENTRA_CLIENT_SECRET=/m);

    const printed = output.join("\n");
    expect(printed).toContain(
      `https://login.microsoftonline.com/${TENANT_ID}/adminconsent?client_id=${application?.appId}`,
    );
    expect(printed).toContain("scripts/docker-live.sh init");
    expect(printed).toContain("Prism365 itself only reads.");
    expect(printed).toContain(CLI_TOOLS_REMINDER);
    expect(printed).not.toContain("synthetic-token");
    expect(tenant.tokens.every((header) => header === "Bearer synthetic-token")).toBe(true);
  });

  it("changes nothing on a second run", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    await run(tenant, optionsFor(dir), dir);
    const files = ["secrets/entra-private.pem", "secrets/entra-public.pem", ".env"];
    const before = await Promise.all(
      files.map(async (file) => ({
        text: await readFile(join(dir, file), "utf8"),
        modified: (await stat(join(dir, file))).mtimeMs,
      })),
    );
    const keys = structuredClone(tenant.applications[0]?.keyCredentials);
    tenant.requests.length = 0;

    const output: string[] = [];
    expect(await run(tenant, optionsFor(dir), dir, output)).toBe("unchanged");
    expect(tenant.writes).toEqual([]);
    expect(tenant.applications).toHaveLength(1);
    expect(tenant.servicePrincipals).toHaveLength(1);
    expect(tenant.applications[0]?.keyCredentials).toEqual(keys);
    for (const [index, file] of files.entries()) {
      expect(await readFile(join(dir, file), "utf8")).toBe(before[index]?.text);
      expect((await stat(join(dir, file))).mtimeMs).toBe(before[index]?.modified);
    }
    expect(output.join("\n")).toContain("already in place");
  });

  it("updates the registration named by --app-id, keeping what you chose to keep", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const otherCertificate = new X509Certificate(
      newCredential(certificateRequest("other.example", new Date())).certificatePem,
    );
    const existingKey = {
      keyId: "44444444-4444-4444-8444-444444444444",
      customKeyIdentifier: "ABCDEF",
      displayName: "CN=Someone else",
      type: "AsymmetricX509Cert",
      usage: "Verify",
      key: otherCertificate.raw.toString("base64"),
      startDateTime: "2026-01-01T00:00:00Z",
      endDateTime: "2026-12-31T00:00:00Z",
    };
    const otherPermissions = {
      resourceAppId: OTHER_API,
      resourceAccess: [{ id: "55555555-5555-4555-8555-555555555555", type: "Scope" }],
    };
    const seeded = tenant.application({
      displayName: "Renamed by the operator",
      tags: ["their-tag", SETUP_TAG],
      web: {
        redirectUris: ["https://prism365.example/api/v1/auth/callback"],
        logoutUrl: "https://prism365.example/signed-out",
        implicitGrantSettings: { enableIdTokenIssuance: false, enableAccessTokenIssuance: false },
      },
      requiredResourceAccess: [
        {
          resourceAppId: GRAPH_RESOURCE_APP_ID,
          resourceAccess: [scope("Mail.Read"), scope("User.Read")] as {
            id: string;
            type: string;
          }[],
        },
        otherPermissions,
      ],
      keyCredentials: [existingKey],
    });
    tenant.servicePrincipal(seeded.appId);

    const options = optionsFor(dir, {
      appId: seeded.appId,
      features: ["audit.signins.read"],
      keepOtherPermissions: true,
      keepOtherRedirectUris: true,
      keepOtherCertificates: true,
    });
    const output: string[] = [];
    expect(await run(tenant, options, dir, output)).toBe("applied");
    expect(output[0]).toContain("The setup checked that you are its only owner");
    expect(output[0]).toContain("its owners can change it at any time");
    expect(output[0]).toContain(
      "You chose to keep them, so the admin consent link grants them too.",
    );
    expect(tenant.applications).toHaveLength(1);
    expect(tenant.writes.map((request) => request.method)).toEqual(["PATCH", "PATCH"]);
    const update = tenant.writes[0]?.body as Record<string, unknown>;
    expect(Object.keys(update).sort()).toEqual(["requiredResourceAccess", "web"]);
    expect(seeded.web).toEqual({
      redirectUris: [
        "https://prism365.example/api/v1/auth/callback",
        "https://localhost:8443/api/v1/auth/callback",
      ],
      logoutUrl: "https://prism365.example/signed-out",
      implicitGrantSettings: { enableIdTokenIssuance: false, enableAccessTokenIssuance: false },
    });
    expect(seeded.displayName).toBe("Renamed by the operator");
    expect(seeded.tags).toEqual(["their-tag", SETUP_TAG]);
    expect(seeded.requiredResourceAccess).toEqual([
      {
        resourceAppId: GRAPH_RESOURCE_APP_ID,
        resourceAccess: [
          "Mail.Read",
          "User.Read",
          "openid",
          "profile",
          "offline_access",
          "AuditLog.Read.All",
        ].map(scope),
      },
      otherPermissions,
    ]);
    // The other certificate goes back exactly as Microsoft returned it, ahead of the new one.
    expect(seeded.keyCredentials[0]).toEqual(existingKey);
    expect(seeded.keyCredentials).toHaveLength(2);
    expect(tenant.servicePrincipals).toHaveLength(1);
  });

  it("removes permissions it did not add only with --prune-permissions, and says so first", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const seeded = tenant.application({
      tags: [SETUP_TAG],
      requiredResourceAccess: [
        {
          resourceAppId: GRAPH_RESOURCE_APP_ID,
          resourceAccess: [scope("Mail.Read"), scope("User.Read")] as {
            id: string;
            type: string;
          }[],
        },
        {
          resourceAppId: OTHER_API,
          resourceAccess: [{ id: "55555555-5555-4555-8555-555555555555", type: "Scope" }],
        },
      ],
    });
    const output: string[] = [];
    await run(
      tenant,
      optionsFor(dir, {
        appId: seeded.appId,
        features: ["audit.signins.read"],
        prunePermissions: true,
      }),
      dir,
      output,
    );
    expect(output[0]).toContain(
      `remove the requests for Mail.Read and a delegated permission of another API (${OTHER_API}) (--prune-permissions; this removes the request only, and consent already granted stays until it is revoked in the Entra admin center)`,
    );
    expect(seeded.requiredResourceAccess).toEqual([
      {
        resourceAppId: GRAPH_RESOURCE_APP_ID,
        resourceAccess: [
          "openid",
          "profile",
          "offline_access",
          "User.Read",
          "AuditLog.Read.All",
        ].map(scope),
      },
    ]);
  });

  it("makes no request that changes anything and writes no file on --dry-run", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const output: string[] = [];
    expect(await run(tenant, optionsFor(dir, { dryRun: true, yes: false }), dir, output)).toBe(
      "dry-run",
    );
    expect(tenant.requests.length).toBeGreaterThan(0);
    expect(tenant.writes).toEqual([]);
    expect(await readdir(dir)).toEqual([]);
    const plan = output[0] ?? "";
    expect(plan).toContain(
      'App registration: create "Prism365", a single-tenant web app with implicit grant off, labelled prism365-setup.',
    );
    expect(plan).toContain("  openid, profile, offline_access and User.Read for sign-in");
    expect(plan).toContain("  User.Read.All for directory.users.read");
    expect(plan).toContain("  LicenseAssignment.Read.All for user.licenses.read");
    expect(plan).toContain("Certificate: create an RSA 2048 key");
    expect(plan).toContain(".env: create it from .env.example, readable only by you, and set:");
    expect(plan).toContain("\n  DATA_MODE from demo to live\n");
    expect(plan).toContain("\n  ENTRA_CLIENT_ID to the new registration's application ID\n");
    expect(output.at(-1)).toContain("Dry run: nothing was changed.");
  });

  it("finishes the job on a re-run after failing midway", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    tenant.intercept = (request) =>
      request.method === "PATCH" ? error(500, "Service_InternalServerError") : null;
    const failure = await run(tenant, optionsFor(dir), dir).catch((caught: unknown) => caught);
    expect(failure).toBeInstanceOf(SetupError);
    expect((failure as Error).message).toContain(
      "Adding the certificate to the registration failed: Microsoft Graph answered HTTP 500 (Service_InternalServerError).",
    );
    expect((failure as Error).message).not.toContain("Provider text");
    expect(tenant.applications[0]?.keyCredentials).toEqual([]);
    const privateKey = await readFile(join(dir, "secrets/entra-private.pem"), "utf8");
    const certificate = new X509Certificate(await readFile(join(dir, "secrets/entra-public.pem")));
    // .env names the registration as soon as it exists, so the next run carries on with it.
    expect(await readFile(join(dir, ".env"), "utf8")).toContain(
      `ENTRA_CLIENT_ID=${tenant.applications[0]?.appId}`,
    );

    tenant.intercept = null;
    tenant.requests.length = 0;
    const output: string[] = [];
    expect(await run(tenant, optionsFor(dir), dir, output)).toBe("applied");
    expect(output[0]).toContain("the one .env uses");
    // The run label the failed run left comes off, and the certificate goes on.
    expect(tenant.writes.map((request) => Object.keys(request.body as object))).toEqual([
      ["tags"],
      ["keyCredentials"],
    ]);
    expect(tenant.applications[0]?.tags).toEqual([SETUP_TAG]);
    expect(tenant.applications).toHaveLength(1);
    expect(tenant.servicePrincipals).toHaveLength(1);
    expect(await readFile(join(dir, "secrets/entra-private.pem"), "utf8")).toBe(privateKey);
    expect(tenant.applications[0]?.keyCredentials).toEqual([
      expect.objectContaining({ key: certificate.raw.toString("base64") }),
    ]);
    expect(await readFile(join(dir, ".env"), "utf8")).toContain(
      `ENTRA_CLIENT_ID=${tenant.applications[0]?.appId}`,
    );
  });

  it("makes a certificate for a key left without one, never replacing the key", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    await run(tenant, optionsFor(dir), dir);
    const privateKey = await readFile(join(dir, "secrets/entra-private.pem"), "utf8");
    await unlink(join(dir, "secrets/entra-public.pem"));
    expect(await run(tenant, optionsFor(dir), dir)).toBe("applied");
    expect(await readFile(join(dir, "secrets/entra-private.pem"), "utf8")).toBe(privateKey);
    const certificate = new X509Certificate(await readFile(join(dir, "secrets/entra-public.pem")));
    expect(certificate.checkPrivateKey(createPrivateKey(privateKey))).toBe(true);
    expect(tenant.applications[0]?.keyCredentials).toHaveLength(2);
  });

  it("keeps the old certificate on the registration when rotating", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    await run(tenant, optionsFor(dir), dir);
    const [old] = tenant.applications[0]?.keyCredentials ?? [];
    const privateKey = await readFile(join(dir, "secrets/entra-private.pem"), "utf8");
    await run(tenant, optionsFor(dir, { rotateCertificate: true }), dir);
    expect(await readFile(join(dir, "secrets/entra-private.pem"), "utf8")).not.toBe(privateKey);
    expect(await mode(join(dir, "secrets/entra-private.pem"))).toBe(0o600);
    expect(tenant.applications[0]?.keyCredentials).toHaveLength(2);
    expect(tenant.applications[0]?.keyCredentials[0]).toEqual(old);
  });

  it("refuses to replace a private key that does not match its certificate", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    await run(tenant, optionsFor(dir), dir);
    const other = newCredential(certificateRequest("other.example", new Date()));
    await writeFile(join(dir, "secrets/entra-public.pem"), other.certificatePem);
    const privateKey = await readFile(join(dir, "secrets/entra-private.pem"), "utf8");
    tenant.requests.length = 0;
    const output: string[] = [];
    await expect(run(tenant, optionsFor(dir), dir, output)).rejects.toThrow(SetupError);
    expect(output[0]).toContain("do not belong together");
    expect(tenant.writes).toEqual([]);
    expect(await readFile(join(dir, "secrets/entra-private.pem"), "utf8")).toBe(privateKey);
  });

  it("never searches: registrations others tagged or named Prism365 change nothing", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    // Someone allowed to register apps tags sixty of their own, and names one Prism365.
    for (let index = 0; index < 60; index += 1)
      tenant.application({
        tags: [SETUP_TAG],
        displayName: index === 0 ? "Prism365" : `junk ${index}`,
        owners: ["99999999-0000-4000-8000-000000000009"],
      });
    const output: string[] = [];
    expect(await run(tenant, optionsFor(dir), dir, output)).toBe("applied");
    // The only search is for this run's own marker, never the public label or a name.
    const searches = tenant.requests.filter(
      (request) => request.method === "GET" && request.path === "applications",
    );
    for (const search of searches)
      expect(search.query.get("$filter")).toMatch(
        /^tags\/any\(t:t eq 'prism365-setup-run-[0-9a-f]{16}'\)$/,
      );
    const mine = tenant.applications[60];
    expect(
      tenant.requests.filter(
        (request) => request.path.endsWith("/owners") && !request.path.includes(mine?.id ?? "-"),
      ),
    ).toEqual([]);
    expect(tenant.applications).toHaveLength(61);
    expect(tenant.applications.slice(0, 60).every((item) => item.keyCredentials.length === 0)).toBe(
      true,
    );
  });

  it("stops before any change when a certificate's key cannot be read in full", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const hidden = "77777777-7777-4777-8777-777777777777";
    const target = tenant.application({
      tags: [SETUP_TAG],
      keyCredentials: [{ keyId: hidden, type: "AsymmetricX509Cert", usage: "Verify", key: "AAAA" }],
      hiddenKeys: [hidden],
    });
    const output: string[] = [];
    await expect(
      run(tenant, optionsFor(dir, { appId: target.appId }), dir, output),
    ).rejects.toThrow(SetupError);
    expect(output[0]).toContain(`did not return the full key of ${hidden}`);
    expect(tenant.writes).toEqual([]);
  });

  it("reports a certificate the update lost", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const kept = new X509Certificate(
      newCredential(certificateRequest("other.example", new Date())).certificatePem,
    );
    const keyId = "88888888-8888-4888-8888-888888888888";
    const target = tenant.application({
      tags: [SETUP_TAG],
      keyCredentials: [
        { keyId, type: "AsymmetricX509Cert", usage: "Verify", key: kept.raw.toString("base64") },
      ],
    });
    tenant.dropOnPatch = [keyId];
    const failure = await run(
      tenant,
      optionsFor(dir, { appId: target.appId, keepOtherCertificates: true }),
      dir,
    ).catch((caught: unknown) => caught);
    expect(failure).toBeInstanceOf(KeyCredentialSafetyError);
    expect((failure as Error).message).toContain(keyId);
  });

  it("refuses a sign-in to a different tenant than the one named", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const options = optionsFor(dir, { tenant: "abababab-abab-4bab-8bab-abababababab" });
    await expect(run(tenant, options, dir)).rejects.toThrow("not abababab");
    expect(tenant.requests).toEqual([]);
  });

  it("asks before applying and changes nothing when the answer is no", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const output: string[] = [];
    const outcome = await runSetup(optionsFor(dir, { yes: false }), {
      fetch: tenant.fetch,
      signIn: async () => ({ tenantId: TENANT_ID, getToken: async () => "synthetic-token" }),
      confirm: async () => false,
      print: (text) => output.push(text),
      envExample,
      cwd: dir,
      sleep: async () => {},
    });
    expect(outcome).toBe("declined");
    expect(tenant.writes).toEqual([]);
    expect(await readdir(dir)).toEqual([]);
  });
});

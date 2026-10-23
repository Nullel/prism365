import { createPrivateKey, generateKeyPairSync, X509Certificate } from "node:crypto";
import { chmod, link, mkdir, readdir, readFile, rename, stat, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import type { AccessList } from "../acl.ts";
import {
  certificateFiles,
  certificateRequest,
  newCredential,
  writePending,
} from "../certificate.ts";
import { GRAPH_RESOURCE_APP_ID } from "../permissions.ts";
import { SETUP_TAG } from "../plan.ts";
import { SetupError } from "../setup.ts";
import { ADMIN_ID, error, FakeTenant, OTHER_API, roleIds, scopeIds } from "./fake-graph.ts";
import { optionsFor, run, workspace } from "./harness.ts";

// Regressions for the adversarial review of the setup (H1, M1, M2, L1, L2, L5, L6). Each test
// reproduces what the review showed and fails against the code it was written for.

const OTHER_ADMIN = "77777777-7777-4777-8777-777777777777";
const mode = async (path: string) => (await stat(path)).mode & 0o777;
const scope = (name: string) => ({ id: scopeIds.get(name) ?? "", type: "Scope" });
const someoneElsesCertificate = () =>
  new X509Certificate(
    newCredential(certificateRequest("other.example", new Date())).certificatePem,
  );

describe("adopting an existing registration (H1)", () => {
  it("never touches a tagged registration someone else owns, and makes your own instead", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const planted = tenant.application({
      owners: [OTHER_ADMIN],
      tags: [SETUP_TAG],
      web: { redirectUris: ["https://collector.invalid/cb"] },
      requiredResourceAccess: [
        {
          resourceAppId: GRAPH_RESOURCE_APP_ID,
          resourceAccess: [{ id: roleIds.get("Directory.Read.All") ?? "", type: "Role" }],
        },
      ],
      passwordCredentials: [{ keyId: "88888888-8888-4888-8888-888888888888", hint: "abc" }],
    });
    tenant.servicePrincipal(planted.appId);
    const output: string[] = [];
    expect(await run(tenant, optionsFor(dir), dir, output)).toBe("applied");
    expect(tenant.writes.some((request) => request.path.includes(planted.id))).toBe(false);
    expect(planted.keyCredentials).toEqual([]);
    expect(tenant.applications).toHaveLength(2);
    const mine = tenant.applications[1];
    expect(mine?.owners).toEqual([ADMIN_ID]);
    expect(output[0]).toContain(`but you do not own it, so the setup leaves it alone`);
    const printed = output.join("\n");
    expect(printed).toContain(`client_id=${mine?.appId}`);
    expect(printed).not.toContain(`client_id=${planted.appId}`);
  });

  it("stops before any write on a registration you own that carries what Prism365 must not adopt", async () => {
    const cases: [string, (tenant: FakeTenant) => void, string][] = [
      [
        "a client secret",
        (tenant) =>
          tenant.application({
            tags: [SETUP_TAG],
            passwordCredentials: [{ keyId: "88888888-8888-4888-8888-888888888888" }],
          }),
        "has a client secret",
      ],
      [
        "mobile and desktop sign-in",
        (tenant) =>
          tenant.application({
            tags: [SETUP_TAG],
            publicClient: { redirectUris: ["msal://callback"] },
          }),
        "mobile and desktop applications",
      ],
      [
        "single-page sign-in",
        (tenant) =>
          tenant.application({
            tags: [SETUP_TAG],
            spa: { redirectUris: ["https://localhost:8443/spa"] },
          }),
        "single-page applications",
      ],
      [
        "public client flows",
        (tenant) => tenant.application({ tags: [SETUP_TAG], isFallbackPublicClient: true }),
        "public client flows",
      ],
      [
        "an application permission it asks for",
        (tenant) =>
          tenant.application({
            tags: [SETUP_TAG],
            requiredResourceAccess: [
              {
                resourceAppId: GRAPH_RESOURCE_APP_ID,
                resourceAccess: [{ id: roleIds.get("Directory.Read.All") ?? "", type: "Role" }],
              },
            ],
          }),
        "asks for application permissions: Directory.Read.All (application permission)",
      ],
      [
        "an application permission already granted",
        (tenant) => {
          const application = tenant.application({ tags: [SETUP_TAG] });
          tenant.servicePrincipal(application.appId, {
            appRoleAssignments: [{ id: "assignment-1", resourceDisplayName: "Microsoft Graph" }],
          });
        },
        "has already been granted an application permission (on Microsoft Graph)",
      ],
    ];
    for (const [name, seed, message] of cases) {
      const dir = await workspace();
      const tenant = new FakeTenant();
      seed(tenant);
      const output: string[] = [];
      // No flag makes these acceptable: they have to be removed in the Entra admin center.
      const options = optionsFor(dir, {
        prunePermissions: true,
        keepOtherRedirectUris: true,
        keepOtherCertificates: true,
      });
      await expect(run(tenant, options, dir, output), name).rejects.toThrow(SetupError);
      expect(output[0], name).toContain(message);
      expect(tenant.writes, name).toEqual([]);
      expect(await readdir(dir), name).toEqual([]);
    }
  });

  it("asks for --app-id when the only match is by name, and checks ownership even then", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const named = tenant.application();
    const output: string[] = [];
    await expect(run(tenant, optionsFor(dir), dir, output)).rejects.toThrow(SetupError);
    expect(output[0]).toContain(`run the setup again with --app-id ${named.appId}`);
    expect(tenant.writes).toEqual([]);

    expect(await run(tenant, optionsFor(dir, { appId: named.appId }), dir)).toBe("applied");
    expect(named.tags).toEqual([SETUP_TAG]);

    const other = tenant.application({ displayName: "Not mine", owners: [OTHER_ADMIN] });
    tenant.requests.length = 0;
    const refused: string[] = [];
    await expect(
      run(tenant, optionsFor(dir, { appId: other.appId }), dir, refused),
    ).rejects.toThrow(SetupError);
    expect(refused[0]).toContain("You are not an owner of");
    expect(tenant.writes).toEqual([]);
  });

  it("keeps other redirect URIs, delegated permissions and certificates only when asked", async () => {
    const seed = (tenant: FakeTenant) => {
      const application = tenant.application({
        tags: [SETUP_TAG],
        web: { redirectUris: ["https://prism365.example/api/v1/auth/callback"] },
        requiredResourceAccess: [
          { resourceAppId: GRAPH_RESOURCE_APP_ID, resourceAccess: [scope("Mail.Read")] },
          {
            resourceAppId: OTHER_API,
            resourceAccess: [{ id: "55555555-5555-4555-8555-555555555555", type: "Scope" }],
          },
        ],
        keyCredentials: [
          {
            keyId: "44444444-4444-4444-8444-444444444444",
            displayName: "CN=Someone else",
            type: "AsymmetricX509Cert",
            usage: "Verify",
            key: someoneElsesCertificate().raw.toString("base64"),
          },
        ],
      });
      tenant.servicePrincipal(application.appId);
      return application;
    };
    const dir = await workspace();
    const tenant = new FakeTenant();
    seed(tenant);
    const output: string[] = [];
    await expect(run(tenant, optionsFor(dir), dir, output)).rejects.toThrow(SetupError);
    for (const flag of [
      "--prune-permissions",
      "--keep-other-permissions",
      "--keep-other-redirect-uris",
      "--keep-other-certificates",
    ])
      expect(output[0]).toContain(flag);
    expect(output[0]).toContain("Mail.Read");
    expect(output[0]).toContain("https://prism365.example/api/v1/auth/callback");
    expect(output[0]).toContain("CN=Someone else");
    expect(output[0]).not.toContain("Nothing else on it changes");
    expect(tenant.writes).toEqual([]);

    const kept: string[] = [];
    const options = optionsFor(dir, {
      keepOtherPermissions: true,
      keepOtherRedirectUris: true,
      keepOtherCertificates: true,
    });
    expect(await run(tenant, options, dir, kept)).toBe("applied");
    expect(kept[0]).toContain("The update sends only web and requiredResourceAccess.");
  });
});

describe("rotating the certificate (M1)", () => {
  it("keeps the working pair when Microsoft refuses the new certificate", async () => {
    const dir = await workspace();
    const files = certificateFiles(join(dir, "secrets"));
    const tenant = new FakeTenant();
    await run(tenant, optionsFor(dir), dir);
    const workingKey = await readFile(files.privateKey, "utf8");
    const registered = tenant.applications[0]?.keyCredentials.map((item) => item.key);

    tenant.intercept = (request) =>
      request.method === "PATCH" ? error(403, "Authorization_RequestDenied") : null;
    await expect(run(tenant, optionsFor(dir, { rotateCertificate: true }), dir)).rejects.toThrow(
      SetupError,
    );
    expect(await readFile(files.privateKey, "utf8")).toBe(workingKey);
    const certificate = new X509Certificate(await readFile(files.certificate));
    expect(certificate.checkPrivateKey(createPrivateKey(workingKey))).toBe(true);
    expect(registered).toContain(certificate.raw.toString("base64"));
    expect(await readdir(files.directory)).not.toContain("entra-private.previous.pem");

    tenant.intercept = null;
    expect(await run(tenant, optionsFor(dir, { rotateCertificate: true }), dir)).toBe("applied");
    expect(await readFile(files.previousKey, "utf8")).toBe(workingKey);
    expect(await readFile(files.privateKey, "utf8")).not.toBe(workingKey);
    expect(await mode(files.previousKey)).toBe(0o600);
    expect(await mode(files.previousCertificate)).toBe(0o644);
    expect(tenant.applications[0]?.keyCredentials).toHaveLength(2);
    // A second run recognises both certificates as its own.
    expect(await run(tenant, optionsFor(dir), dir)).toBe("unchanged");

    const again: string[] = [];
    await expect(
      run(tenant, optionsFor(dir, { rotateCertificate: true }), dir, again),
    ).rejects.toThrow(SetupError);
    expect(again[0]).toContain("is still there from the last rotation");
  });

  it("finishes a rotation whose move was interrupted", async () => {
    const dir = await workspace();
    const files = certificateFiles(join(dir, "secrets"));
    const tenant = new FakeTenant();
    await run(tenant, optionsFor(dir), dir);
    const made = newCredential(certificateRequest("contoso.example", new Date()));
    await writePending(files, made);
    const application = tenant.applications[0];
    application?.keyCredentials.push({
      keyId: "99999999-9999-4999-8999-999999999999",
      type: "AsymmetricX509Cert",
      usage: "Verify",
      key: new X509Certificate(made.certificatePem).raw.toString("base64"),
    });
    // The key was moved, the certificate was not.
    await link(files.privateKey, files.previousKey);
    await rename(files.pendingKey, files.privateKey);
    tenant.requests.length = 0;

    const output: string[] = [];
    expect(await run(tenant, optionsFor(dir), dir, output)).toBe("applied");
    expect(output[0]).toContain("finish the earlier rotation");
    expect(await readFile(files.privateKey, "utf8")).toBe(made.privateKeyPem);
    expect(await readFile(files.certificate, "utf8")).toBe(made.certificatePem);
    expect(tenant.writes.filter((request) => request.method === "PATCH")).toEqual([]);
  });
});

describe("access control lists (M2)", () => {
  const entries = ["user:1000:r--", "mask::r--"];

  it("leaves a key with an access list alone, and gives a rotated key the same list", async () => {
    const dir = await workspace();
    const files = certificateFiles(join(dir, "secrets"));
    const tenant = new FakeTenant();
    await run(tenant, optionsFor(dir), dir);
    // What setfacl -m u:1000:r leaves behind: the mask shows in the group bits.
    await chmod(files.privateKey, 0o640);
    const applied: { path: string; entries: readonly string[] }[] = [];
    const acl = {
      read: async (path: string): Promise<AccessList | null> =>
        path === files.privateKey ? { entries, defaults: [] } : { entries: [], defaults: [] },
      apply: async (path: string, list: readonly string[]) => {
        applied.push({ path, entries: list });
        return true;
      },
    };
    const output: string[] = [];
    expect(await run(tenant, optionsFor(dir), dir, output, { acl })).toBe("unchanged");
    expect(await mode(files.privateKey)).toBe(0o640);
    expect(output[0]).not.toContain("Restrict");
    expect(output[0]).toContain("has an access control list (ACL)");

    const rotated: string[] = [];
    await run(tenant, optionsFor(dir, { rotateCertificate: true }), dir, rotated, { acl });
    expect(applied).toEqual([{ path: files.privateKey, entries: ["user:1000:r--"] }]);

    // Where the list cannot be copied, the exact command is printed instead.
    const dir2 = await workspace();
    const files2 = certificateFiles(join(dir2, "secrets"));
    const tenant2 = new FakeTenant();
    await run(tenant2, optionsFor(dir2), dir2);
    const failing = {
      read: async (path: string) => (path === files2.privateKey ? { entries, defaults: [] } : null),
      apply: async () => false,
    };
    const printed: string[] = [];
    await run(tenant2, optionsFor(dir2, { rotateCertificate: true }), dir2, printed, {
      acl: failing,
    });
    expect(printed.join("\n")).toContain(`setfacl -m user:1000:r-- '${files2.privateKey}'`);
  });

  it("does not tighten a group-readable key it cannot read the access list of", async () => {
    const dir = await workspace();
    const files = certificateFiles(join(dir, "secrets"));
    const tenant = new FakeTenant();
    await run(tenant, optionsFor(dir), dir);
    await chmod(files.privateKey, 0o640);
    const output: string[] = [];
    await run(tenant, optionsFor(dir), dir, output);
    expect(await mode(files.privateKey)).toBe(0o640);
    expect(output[0]).toContain("which may be an access control list's mask");
  });

  it("writes keys so a directory's default access list applies to them", async () => {
    const dir = await workspace();
    await mkdir(join(dir, "secrets"), { mode: 0o700 });
    const tenant = new FakeTenant();
    const acl = {
      read: async (path: string) =>
        path === join(dir, "secrets") ? { entries: [], defaults: ["user:1000:r--"] } : null,
      apply: async () => true,
    };
    await run(tenant, optionsFor(dir), dir, [], { acl });
    expect(await mode(join(dir, "secrets/entra-private.pem"))).toBe(0o640);
  });
});

describe("creating once (L1)", () => {
  it("never makes a second registration when Microsoft committed a create it answered with a 504", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    let first = true;
    tenant.intercept = (request) => {
      if (request.method === "POST" && request.path === "applications" && first) {
        first = false;
        tenant.application({
          ...(request.body as object),
          keyCredentials: [],
          owners: [ADMIN_ID],
        } as never);
        return error(504, "GatewayTimeout");
      }
      if (request.method === "POST" && request.path === "servicePrincipals" && first === false) {
        first = true;
        tenant.servicePrincipal((request.body as { appId: string }).appId);
        return error(503, "ServiceUnavailable");
      }
      return null;
    };
    expect(await run(tenant, optionsFor(dir), dir)).toBe("applied");
    expect(tenant.applications).toHaveLength(1);
    expect(tenant.servicePrincipals).toHaveLength(1);
    expect(tenant.writes.filter((request) => request.method === "POST")).toHaveLength(2);
    tenant.intercept = null;
    expect(await run(tenant, optionsFor(dir), dir)).toBe("unchanged");
  });
});

describe("the certificate list update (L2)", () => {
  it("stops when the registration's certificates changed after the plan was made", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const application = tenant.application({ tags: [SETUP_TAG] });
    tenant.servicePrincipal(application.appId);
    const output: string[] = [];
    const failure = await run(tenant, optionsFor(dir, { yes: false }), dir, output, {
      // While the plan is on screen, someone uploads a certificate in the portal.
      confirm: async () => {
        application.keyCredentials.push({
          keyId: "abababab-abab-4bab-8bab-abababababab",
          type: "AsymmetricX509Cert",
          usage: "Verify",
          key: someoneElsesCertificate().raw.toString("base64"),
        });
        return true;
      },
    }).catch((caught: unknown) => caught);
    expect(failure).toBeInstanceOf(SetupError);
    expect((failure as Error).message).toContain(
      "certificates changed while the setup was running",
    );
    expect(
      tenant.writes.filter(
        (request) => (request.body as { keyCredentials?: unknown }).keyCredentials,
      ),
    ).toEqual([]);
    expect(application.keyCredentials).toHaveLength(1);
  });

  it("says only what it verified after the update", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const output: string[] = [];
    await run(tenant, optionsFor(dir), dir, output);
    const printed = output.join("\n");
    expect(printed).toContain("checked that every certificate it had just before is still on it");
    expect(printed).not.toContain("Every certificate that was there before is still there");
  });
});

describe("the key in the secrets directory (L5)", () => {
  it("stops at the plan, before Microsoft is changed, when the key is not RSA 2048", async () => {
    for (const key of [
      generateKeyPairSync("ec", { namedCurve: "P-256" }).privateKey,
      generateKeyPairSync("rsa", { modulusLength: 1024 }).privateKey,
    ]) {
      const dir = await workspace();
      await mkdir(join(dir, "secrets"), { mode: 0o700 });
      await writeFile(
        join(dir, "secrets/entra-private.pem"),
        key.export({ type: "pkcs8", format: "pem" }),
        {
          mode: 0o600,
        },
      );
      const tenant = new FakeTenant();
      const output: string[] = [];
      await expect(run(tenant, optionsFor(dir), dir, output)).rejects.toThrow(SetupError);
      expect(output[0]).toContain("--rotate-certificate");
      expect(tenant.writes).toEqual([]);
    }
  });
});

describe("Windows (L6)", () => {
  it("plans no mode changes and says the folder's access list decides", async () => {
    const dir = await workspace();
    await mkdir(join(dir, "secrets"));
    await chmod(join(dir, "secrets"), 0o755);
    const tenant = new FakeTenant();
    const output: string[] = [];
    await run(tenant, optionsFor(dir, { dryRun: true }), dir, output, { platform: "win32" });
    expect(output[0]).not.toContain("Restrict");
    expect(output[0]).not.toContain("readable only by you");
    expect(output[0]).toContain("on Windows the folder's access control list decides");

    const unix: string[] = [];
    await run(tenant, optionsFor(dir, { dryRun: true }), dir, unix, { platform: "darwin" });
    expect(unix[0]).toContain("Restrict secrets so only you can open it (mode 700).");
  });
});

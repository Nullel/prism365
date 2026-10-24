import { X509Certificate } from "node:crypto";
import { readFile, rm, stat, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import type { AccessList } from "../acl.ts";
import { certificateFiles, certificateRequest, newCredential } from "../certificate.ts";
import { SETUP_TAG } from "../plan.ts";
import { SetupError } from "../setup.ts";
import { ADMIN_ID, error, FakeTenant, GRAPH_PRINCIPAL_ID } from "./fake-graph.ts";
import { optionsFor, run, workspace } from "./harness.ts";

// Regressions for the second adversarial review of the setup: the registration it works on is
// never found by searching (R-H1, R-L3), credentials and consent beyond the application's own
// (R-M1, R-L1), a create Microsoft is slow to show (R-L2), and the follow-ups (INFO).

const mode = async (path: string) => (await stat(path)).mode & 0o777;
const someoneElsesCertificate = () =>
  new X509Certificate(
    newCredential(certificateRequest("other.example", new Date())).certificatePem,
  );

describe("the registration to work on", () => {
  it("is the one .env names, or the one --app-id names, and never a missing one", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    await writeFile(join(dir, ".env"), "ENTRA_CLIENT_ID=12121212-1212-4212-8212-121212121212\n");
    const output: string[] = [];
    await expect(run(tenant, optionsFor(dir), dir, output)).rejects.toThrow(SetupError);
    expect(output[0]).toContain("names application 12121212-1212-4212-8212-121212121212");
    expect(tenant.writes).toEqual([]);

    const chosen = tenant.application();
    expect(await run(tenant, optionsFor(dir, { appId: chosen.appId }), dir)).toBe("applied");
    expect(await readFile(join(dir, ".env"), "utf8")).toContain(`ENTRA_CLIENT_ID=${chosen.appId}`);
    // From now on .env names it, and the setup works on it without --app-id.
    const again: string[] = [];
    expect(await run(tenant, optionsFor(dir), dir, again)).toBe("unchanged");
    expect(again[0]).toContain("the one .env uses");

    await writeFile(join(dir, ".env"), "ENTRA_CLIENT_ID=not-an-id\n");
    const invalid: string[] = [];
    await expect(run(tenant, optionsFor(dir), dir, invalid)).rejects.toThrow(SetupError);
    expect(invalid[0]).toContain('is "not-an-id", which is not an application ID');
  });
});

describe("credentials beyond the application's own (R-M1)", () => {
  it("stops on federated credentials and service principal secrets, whatever is kept", async () => {
    const cases: [string, (tenant: FakeTenant) => { appId: string }, string][] = [
      [
        "federated credential",
        (tenant) =>
          tenant.application({
            federatedIdentityCredentials: [{ id: "f1", issuer: "https://token.example" }],
          }),
        "has a federated credential",
      ],
      [
        "service principal secret",
        (tenant) => {
          const application = tenant.application();
          tenant.servicePrincipal(application.appId, { passwordCredentials: [{ keyId: "p1" }] });
          return application;
        },
        "or its service principal has a client secret",
      ],
    ];
    for (const [name, seed, message] of cases) {
      const dir = await workspace();
      const tenant = new FakeTenant();
      const target = seed(tenant);
      const output: string[] = [];
      const options = optionsFor(dir, {
        appId: target.appId,
        keepOtherCertificates: true,
        keepOtherOwners: true,
        keepOtherPermissions: true,
        keepOtherRedirectUris: true,
      });
      await expect(run(tenant, options, dir, output), name).rejects.toThrow(SetupError);
      expect(output[0], name).toContain(message);
      expect(tenant.writes, name).toEqual([]);
    }
  });

  it("stops on a service principal certificate it did not make, unless kept", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const application = tenant.application();
    tenant.servicePrincipal(application.appId, {
      keyCredentials: [
        {
          keyId: "abababab-abab-4bab-8bab-abababababab",
          displayName: "CN=Signing",
          key: someoneElsesCertificate().raw.toString("base64"),
        },
      ],
    });
    const output: string[] = [];
    await expect(
      run(tenant, optionsFor(dir, { appId: application.appId }), dir, output),
    ).rejects.toThrow(SetupError);
    expect(output[0]).toContain(
      "CN=Signing (abababab-abab-4bab-8bab-abababababab) on its service principal",
    );
    expect(tenant.writes).toEqual([]);
    const kept: string[] = [];
    const options = optionsFor(dir, { appId: application.appId, keepOtherCertificates: true });
    expect(await run(tenant, options, dir, kept)).toBe("applied");
    expect(kept[0]).not.toContain("no certificates other than");
  });

  it("says what it checked on both the registration and its service principal", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const application = tenant.application();
    tenant.servicePrincipal(application.appId);
    const output: string[] = [];
    await run(tenant, optionsFor(dir, { appId: application.appId }), dir, output);
    expect(output[0]).toContain(
      "it and its service principal have no client secrets or federated credentials",
    );
    expect(output[0]).toContain(
      "it and its service principal have no certificates other than the ones this setup made",
    );
  });
});

describe("consent already granted (R-L1)", () => {
  const grant = (scope: string, resourceId = GRAPH_PRINCIPAL_ID) => ({
    id: `grant-${scope.length}`,
    consentType: "AllPrincipals",
    resourceId,
    scope,
  });

  it("stops when consent was given beyond Prism365's permissions, and never deletes it", async () => {
    for (const [name, grants, message] of [
      ["a Graph scope", [grant("openid User.Read Mail.ReadWrite")], "consent for Mail.ReadWrite"],
      [
        "another API",
        [grant("user_impersonation", "77777777-0000-4000-8000-000000000007")],
        "consent for delegated permissions on another API",
      ],
    ] as const) {
      const dir = await workspace();
      const tenant = new FakeTenant();
      const application = tenant.application();
      tenant.servicePrincipal(application.appId, { grants: [...grants] });
      const output: string[] = [];
      const options = optionsFor(dir, { appId: application.appId, prunePermissions: true });
      await expect(run(tenant, options, dir, output), name).rejects.toThrow(SetupError);
      expect(output[0], name).toContain(message);
      expect(output[0], name).toContain("Revoke it under Enterprise applications");
      expect(tenant.writes, name).toEqual([]);
      expect(
        tenant.requests.some((request) => request.method === "DELETE"),
        name,
      ).toBe(false);
    }
  });

  it("accepts consent given to any of Prism365's own permissions", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const application = tenant.application();
    tenant.servicePrincipal(application.appId, {
      grants: [grant("openid profile offline_access User.Read User.Read.All Policy.Read.All")],
    });
    const output: string[] = [];
    expect(await run(tenant, optionsFor(dir, { appId: application.appId }), dir, output)).toBe(
      "applied",
    );
    expect(output[0]).toContain("no consent has been granted beyond Prism365's permissions");
  });
});

describe("a create Microsoft is slow to show (R-L2)", () => {
  // Microsoft commits the create but answers 504, and its searches miss the new registration for
  // the first `hidden` look-ups.
  function slowTenant(hidden: number, commit: boolean) {
    const tenant = new FakeTenant();
    let posts = 0;
    let misses = hidden;
    tenant.intercept = (request) => {
      if (request.method === "POST" && request.path === "applications") {
        posts += 1;
        if (posts === 1) {
          if (commit)
            tenant.application({
              ...(request.body as object),
              keyCredentials: [],
              owners: [ADMIN_ID],
            } as never);
          return error(504, "GatewayTimeout");
        }
      }
      if (request.method === "GET" && request.path === "applications" && misses > 0) {
        misses -= 1;
        return Response.json({ value: [] });
      }
      return null;
    };
    return { tenant, posts: () => posts };
  }

  it("keeps looking for this run's marker for the whole wait, not five seconds", async () => {
    const dir = await workspace();
    // Hidden from three look-ups: two, four and eight seconds after the create.
    const { tenant, posts } = slowTenant(3, true);
    const waits: number[] = [];
    expect(
      await run(tenant, optionsFor(dir), dir, [], {
        sleep: async (milliseconds) => {
          waits.push(milliseconds);
        },
      }),
    ).toBe("applied");
    expect(posts()).toBe(1);
    expect(tenant.applications).toHaveLength(1);
    expect(waits.slice(0, 4)).toEqual([2_000, 4_000, 8_000, 15_000]);
    // Every look-up was by this run's marker, never by the public label.
    const lookups = tenant.requests.filter(
      (request) => request.method === "GET" && request.path === "applications",
    );
    expect(
      lookups.every((request) => request.query.get("$filter")?.includes("prism365-setup-run-")),
    ).toBe(true);
    // The marker comes off once the run is done.
    expect(tenant.applications[0]?.tags).toEqual([SETUP_TAG]);
  });

  it("creates again only once the marker is confirmed absent after the whole wait", async () => {
    const dir = await workspace();
    const { tenant, posts } = slowTenant(0, false);
    expect(await run(tenant, optionsFor(dir), dir)).toBe("applied");
    expect(posts()).toBe(2);
    expect(tenant.applications).toHaveLength(1);
    const lookups = tenant.requests.filter(
      (request) => request.method === "GET" && request.path === "applications",
    );
    // Five look-ups (the whole wait) between the first create and the second.
    expect(lookups.length).toBeGreaterThanOrEqual(5);
  });

  it("names the extra registration when the first create shows up after all", async () => {
    const dir = await workspace();
    // Hidden through the whole wait: the setup creates again, then both appear.
    const { tenant, posts } = slowTenant(5, true);
    const output: string[] = [];
    expect(await run(tenant, optionsFor(dir), dir, output)).toBe("applied");
    expect(posts()).toBe(2);
    expect(tenant.applications).toHaveLength(2);
    const [first] = tenant.applications;
    expect(output.join("\n")).toContain(`made a second registration, ${first?.appId}`);
  });
});

describe("after a rotation (INFO)", () => {
  it("names the certificate it replaced once the .previous files are gone", async () => {
    const dir = await workspace();
    const files = certificateFiles(join(dir, "secrets"));
    const tenant = new FakeTenant();
    await run(tenant, optionsFor(dir), dir);
    await run(tenant, optionsFor(dir, { rotateCertificate: true }), dir);
    const record = JSON.parse(await readFile(files.record, "utf8"));
    const old = new X509Certificate(await readFile(files.previousCertificate));
    expect(record.replacedThumbprint).toBe(old.fingerprint256.replaceAll(":", ""));
    expect(await mode(files.record)).toBe(0o600);
    await rm(files.previousKey);
    await rm(files.previousCertificate);
    const output: string[] = [];
    await expect(run(tenant, optionsFor(dir), dir, output)).rejects.toThrow(SetupError);
    expect(output[0]).toContain("is the certificate the setup replaced in its last rotation");
    expect(output[0]).not.toContain("certificates this setup did not make");
  });

  it("tries again with the entries the .previous files carry, never the record's", async () => {
    const dir = await workspace();
    const files = certificateFiles(join(dir, "secrets"));
    const tenant = new FakeTenant();
    await run(tenant, optionsFor(dir), dir);
    const entries = ["user:1000:r--", "mask::r--"];
    // The old key has the list; after the rotation it is the .previous file, and the new key is
    // the one without it.
    let rotated = false;
    const read = async (path: string): Promise<AccessList | null> =>
      (!rotated && path === files.privateKey) || path === files.previousKey
        ? { entries, defaults: [] }
        : { entries: [], defaults: [] };
    const failing = { read, apply: async () => false };
    await run(tenant, optionsFor(dir, { rotateCertificate: true }), dir, [], { acl: failing });
    rotated = true;
    const record = JSON.parse(await readFile(files.record, "utf8"));
    // Only a hint: the record never holds the entries themselves.
    expect(record.aclPending).toBe(true);
    expect(JSON.stringify(record)).not.toContain("user:1000");

    const applied: string[][] = [];
    const working = {
      read,
      apply: async (_path: string, list: readonly string[]) => {
        applied.push([...list]);
        return true;
      },
    };
    const output: string[] = [];
    expect(await run(tenant, optionsFor(dir), dir, output, { acl: working })).toBe("applied");
    expect(output[0]).toContain("the access list entries the .previous files carry");
    expect(output[0]).toContain("user:1000:r--");
    expect(applied).toEqual([["user:1000:r--"]]);
    expect(JSON.parse(await readFile(files.record, "utf8")).aclPending).toBe(false);
  });
});

import { chmod, readFile, stat, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import type { AccessList } from "../acl.ts";
import { certificateFiles } from "../certificate.ts";
import { RUN_TAG_PREFIX, SETUP_TAG } from "../plan.ts";
import { SetupError } from "../setup.ts";
import { ADMIN_ID, error, FakeTenant } from "./fake-graph.ts";
import { optionsFor, run, workspace } from "./harness.ts";

// Regressions for the third adversarial review of the setup: a registration that carries the
// run's label is checked before it is used (R3-L1), group write and recorded ACL entries (R3-L2),
// and moving an install to another registration (R3-L3).

const ATTACKER = "99999999-0000-4000-8000-000000000009";
const mode = async (path: string) => (await stat(path)).mode & 0o777;

// A copy someone who can read applications makes of the run's label, with the administrator
// added as a second owner, a secret, and a redirect URI of theirs.
function copyOf(tenant: FakeTenant, marker: string) {
  return tenant.application({
    displayName: "Prism365",
    tags: [SETUP_TAG, marker],
    owners: [ATTACKER, ADMIN_ID],
    web: {
      redirectUris: ["https://collector.invalid/cb", "https://localhost:8443/api/v1/auth/callback"],
    },
    passwordCredentials: [{ keyId: "s", hint: "mal" }],
  });
}
const markerOf = (body: unknown) =>
  (body as { tags: string[] }).tags.find((tag) => tag.startsWith(RUN_TAG_PREFIX)) ?? "";

describe("a registration carrying this run's label (R3-L1)", () => {
  it("is used only when it passes every check, and is never handed a consent link otherwise", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    let real: { id: string } | null = null;
    let copy: { appId: string; keyCredentials: unknown[] } | null = null;
    tenant.intercept = (request) => {
      if (request.method === "POST" && request.path === "applications" && !real) {
        real = tenant.application({
          ...(request.body as object),
          owners: [ADMIN_ID],
        } as never);
        copy = copyOf(tenant, markerOf(request.body));
        return error(504, "GatewayTimeout");
      }
      // The real one is not visible to the label search yet; the copy is.
      if (request.method === "GET" && request.path === "applications" && real) {
        const hidden = (real as { id: string }).id;
        const marker = /eq '([^']+)'/.exec(request.query.get("$filter") ?? "")?.[1] ?? "";
        return Response.json({
          value: tenant.applications
            .filter((item) => item.id !== hidden && item.tags.includes(marker))
            .map((item) => ({ id: item.id, appId: item.appId })),
        });
      }
      return null;
    };
    const output: string[] = [];
    const failure = await run(tenant, optionsFor(dir), dir, output).catch((caught) => caught);
    expect(failure).toBeInstanceOf(SetupError);
    const message = (failure as Error).message;
    const copied = copy as unknown as { appId: string; keyCredentials: unknown[] };
    expect(message).toContain(`application ID ${copied.appId}`);
    expect(message).toContain("does not pass the setup's checks");
    expect(message).toContain("did not use it or print a consent link for it");
    expect(copied.keyCredentials).toEqual([]);
    expect(await readFile(join(dir, ".env"), "utf8").catch(() => "")).not.toContain(copied.appId);
    expect(output.join("\n")).not.toContain("adminconsent");
  });

  it("calls a late registration with the label someone else's when it fails the checks", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    let copied: string | null = null;
    tenant.intercept = (request) => {
      // After the create succeeded, someone copies its label onto a registration of theirs.
      if (request.method === "POST" && request.path === "servicePrincipals" && !copied) {
        const mine = tenant.applications.find(
          (item) => item.appId === (request.body as { appId: string }).appId,
        );
        const marker = mine?.tags.find((tag) => tag.startsWith(RUN_TAG_PREFIX)) ?? "";
        copied = copyOf(tenant, marker).appId;
      }
      return null;
    };
    const output: string[] = [];
    expect(await run(tenant, optionsFor(dir), dir, output)).toBe("applied");
    const printed = output.join("\n");
    expect(printed).toContain(
      `A registration that is not this run's, "Prism365" (application ID ${copied})`,
    );
    expect(printed).not.toContain(`made a second registration, ${copied}`);
    expect(printed).not.toContain(`client_id=${copied}`);
  });
});

describe("group write and access list entries (R3-L2)", () => {
  it("always removes group write, even where group bits are kept for an access list", async () => {
    const dir = await workspace();
    const files = certificateFiles(join(dir, "secrets"));
    const tenant = new FakeTenant();
    await run(tenant, optionsFor(dir), dir);
    await chmod(files.directory, 0o770);
    await run(tenant, optionsFor(dir), dir, [], {
      platform: "linux",
      acl: { read: async () => null, apply: async () => false },
    });
    expect(await mode(files.directory)).toBe(0o750);
    await chmod(files.privateKey, 0o660);
    await run(tenant, optionsFor(dir), dir, [], {
      platform: "linux",
      acl: {
        read: async () => ({ entries: ["user:1000:r--", "mask::rw-"], defaults: [] }),
        apply: async () => true,
      },
    });
    expect(await mode(files.privateKey)).toBe(0o640);
  });

  it("applies only entries the .previous files carry, whatever setup-state.json says", async () => {
    const dir = await workspace();
    const files = certificateFiles(join(dir, "secrets"));
    const tenant = new FakeTenant();
    await run(tenant, optionsFor(dir), dir);
    await run(tenant, optionsFor(dir, { rotateCertificate: true }), dir);
    // Someone who can write in the folder plants an entry of their own in the record.
    await writeFile(
      files.record,
      JSON.stringify({
        aclPending: true,
        pendingAcl: { privateKey: ["user:mallory:rwx"], certificate: [] },
      }),
    );
    const applied: string[] = [];
    const output: string[] = [];
    await run(tenant, optionsFor(dir), dir, output, {
      platform: "linux",
      acl: {
        read: async (): Promise<AccessList> => ({ entries: [], defaults: [] }),
        apply: async (path, entries) => {
          applied.push(`${path} ${entries.join(",")}`);
          return true;
        },
      },
    });
    expect(applied).toEqual([]);
    expect(output[0]).not.toContain("mallory");
    expect(JSON.parse(await readFile(files.record, "utf8")).aclPending).toBe(false);
  });

  it("lists the entries a rotation will give the new files before it applies them", async () => {
    const dir = await workspace();
    const files = certificateFiles(join(dir, "secrets"));
    const tenant = new FakeTenant();
    await run(tenant, optionsFor(dir), dir);
    const output: string[] = [];
    await run(tenant, optionsFor(dir, { rotateCertificate: true, dryRun: true }), dir, output, {
      platform: "linux",
      acl: {
        read: async (path) =>
          path === files.privateKey
            ? { entries: ["user:1000:r--", "mask::r--"], defaults: [] }
            : null,
        apply: async () => true,
      },
    });
    expect(output[0]).toContain(
      "After the move, give the new files the access list entries the old ones have: secrets/entra-private.pem user:1000:r--.",
    );
  });
});

describe("moving an install to another registration (R3-L3)", () => {
  it("writes the new registration only after every step worked, and remembers the old one", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    await run(tenant, optionsFor(dir), dir);
    const first = tenant.applications[0]?.appId ?? "";
    const other = tenant.application();
    tenant.intercept = (request) =>
      request.method === "PATCH" && (request.body as { keyCredentials?: unknown }).keyCredentials
        ? error(500, "Service_InternalServerError")
        : null;
    const options = optionsFor(dir, { appId: other.appId });
    const failed: string[] = [];
    await expect(run(tenant, options, dir, failed)).rejects.toThrow(SetupError);
    expect(failed[0]).toContain(`This install uses application ${first} now`);
    // The failed switch leaves the install on the registration it had.
    expect(await readFile(join(dir, ".env"), "utf8")).toContain(`ENTRA_CLIENT_ID=${first}`);

    tenant.intercept = null;
    const output: string[] = [];
    expect(await run(tenant, options, dir, output)).toBe("applied");
    expect(await readFile(join(dir, ".env"), "utf8")).toContain(`ENTRA_CLIENT_ID=${other.appId}`);
    const record = JSON.parse(
      await readFile(certificateFiles(join(dir, "secrets")).record, "utf8"),
    );
    expect(record.previousClientId).toBe(first);
    expect(output.join("\n")).toContain(
      `To switch back, run: npm run setup -- --tenant contoso.example --app-id ${first}`,
    );
  });

  it("still writes a registration the install never named before at once, so a failed run resumes on it", async () => {
    const dir = await workspace();
    const tenant = new FakeTenant();
    const chosen = tenant.application();
    tenant.intercept = (request) =>
      request.method === "PATCH" && (request.body as { keyCredentials?: unknown }).keyCredentials
        ? error(500, "Service_InternalServerError")
        : null;
    await expect(run(tenant, optionsFor(dir, { appId: chosen.appId }), dir)).rejects.toThrow(
      SetupError,
    );
    expect(await readFile(join(dir, ".env"), "utf8")).toContain(`ENTRA_CLIENT_ID=${chosen.appId}`);
  });
});

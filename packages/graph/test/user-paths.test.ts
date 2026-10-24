import { describe, expect, it } from "vitest";
import { GraphBoundaryError, GraphHttpClient } from "../src/client.ts";
import { investigateUser, UserInvestigationError } from "../src/investigation.ts";
import { lookupPersonId, readPersonDirectory } from "../src/person.ts";

// The user pack's paths, and the repository's own builders of them. Every value is invented.

const users = { feature: "directory.users.read", permission: "User.Read.All" } as const;
const userId = "44444444-4444-4444-8444-444444444444";

function recorder(respond: (url: URL) => Response = () => Response.json({ value: [] })) {
  const sent: string[] = [];
  const client = new GraphHttpClient({
    getAccessToken: async () => "synthetic-token",
    sleep: async () => {},
    maxRetries: 0,
    fetchImplementation: (async (input: RequestInfo | URL) => {
      const url = new URL(String(input));
      sent.push(url.pathname);
      return respond(url);
    }) as typeof fetch,
  });
  return { client, sent };
}

const allowed = async (path: string) =>
  recorder()
    .client.request(path, users)
    .then(
      () => true,
      (error: unknown) => {
        if (error instanceof GraphBoundaryError && error.reason === "endpoint") return false;
        throw error;
      },
    );

// The two forms the path builders write (person.ts, investigation.ts): the id percent-encoded, or
// the principal name with its quotes doubled, percent-encoded, and the quotes left as %27.
const byId = (id: string) => `users/${encodeURIComponent(id)}`;
const byName = (name: string) =>
  `users('${encodeURIComponent(name.replaceAll("'", "''")).replaceAll("'", "%27")}')`;

// Principal names in the characters Microsoft allows (letters, digits and ' . - _ ! # ^ ~), a guest,
// and the leading "$" the OData key form exists for.
const principals = [
  "ada@lab.example",
  "o'hara@lab.example",
  "o!hara@lab.example",
  "a~b@lab.example",
  "a^b@lab.example",
  "a_b-c.d@lab.example",
  "guest_partner.example#EXT#@lab.example",
  "$admin@lab.example",
];

describe("the user pack's paths", () => {
  it("allow a user, its manager and its memberships in both forms the builders write", async () => {
    for (const name of principals)
      for (const base of [byId(name), byName(name)])
        for (const path of [base, `${base}/manager`, `${base}/memberOf`])
          expect(await allowed(path), path).toBe(true);
    // encodeURIComponent leaves ! ~ * ' ( ) as they are, so the id class takes them too.
    for (const path of [
      "users/a*b!c~d'e(f)/memberOf",
      "users('a*b!c~d%40lab.example')/memberOf",
      `${byId(userId)}/memberOf`,
    ])
      expect(await allowed(path), path).toBe(true);
  });

  it("refuse an encoded slash, backslash or question mark in every form", async () => {
    for (const separator of ["%2F", "%2f", "%5C", "%5c", "%3F", "%3f"])
      for (const base of [`users/a${separator}b`, `users('a${separator}b')`])
        for (const path of [base, `${base}/manager`, `${base}/memberOf`])
          expect(await allowed(path), path).toBe(false);
    for (const path of [
      "users/x%2Fmanager",
      "users/x%2FtransitiveMemberOf",
      "users('x%2FtransitiveMemberOf')",
      "users('a%2F..%2Fx')/memberOf",
      "users/a%zzb",
      "users/a%2",
    ])
      expect(await allowed(path), path).toBe(false);
  });

  it("refuse characters the builders always encode", async () => {
    // The URL parser itself encodes a space, quotes, <, >, ^, `, { and }, so they never arrive raw.
    for (const character of ["+", ",", ";", "=", "$", "&", "[", "]", "|"])
      for (const path of [
        `users/a${character}b`,
        `users/a${character}b/manager`,
        `users/a${character}b/memberOf`,
        `users('a${character}b')`,
      ])
        expect(await allowed(path), path).toBe(false);
  });
});

describe("the repository's own user path builders", () => {
  it("reach Graph for every principal name Microsoft allows", async () => {
    for (const userPrincipalName of principals) {
      const { client, sent } = recorder(() => Response.json({ id: userId }));
      await expect(lookupPersonId({ userPrincipalName, client }), userPrincipalName).resolves.toBe(
        userId,
      );
      expect(sent, userPrincipalName).toHaveLength(1);
    }
  });

  it("read a person's profile, manager and licences by id", async () => {
    const { client, sent } = recorder((url) =>
      url.pathname.endsWith("/licenseDetails")
        ? Response.json({ value: [] })
        : Response.json({
            id: userId,
            displayName: "Synthetic Person",
            userPrincipalName: "person@lab.example",
          }),
    );
    await readPersonDirectory({
      key: { userId: "a.b_c@lab:d-e" },
      client,
      grantedScopes: ["User.Read.All", "LicenseAssignment.Read.All"],
    });
    expect(sent.some((path) => path.endsWith("/manager"))).toBe(true);
    expect(sent.length).toBeGreaterThanOrEqual(2);
  });

  it("answer not found for a name no account can hold, without asking Graph", async () => {
    // Microsoft does not allow "/", "\" or "?" in a principal name, and the client refuses a user
    // path that encodes them, so the builders do not send one.
    for (const userPrincipalName of ["a/b@lab.example", "a?b@lab.example", "a\\b@lab.example"]) {
      const { client, sent } = recorder(() => Response.json({ id: userId }));
      await expect(lookupPersonId({ userPrincipalName, client })).resolves.toBeNull();
      expect(sent, userPrincipalName).toHaveLength(0);
    }
    const directory = recorder(() => Response.json({ id: userId }));
    await expect(
      readPersonDirectory({
        key: { userPrincipalName: "a/b@lab.example" },
        client: directory.client,
        grantedScopes: ["User.Read.All"],
      }),
    ).resolves.toMatchObject({ state: "notFound", profile: null });
    expect(directory.sent).toHaveLength(0);
    for (const email of ["a/b@lab.example", "a?b@lab.example"]) {
      const { client, sent } = recorder(() => Response.json({ id: userId }));
      const error = await investigateUser({
        email,
        client,
        grantedScopes: ["User.Read.All"],
      }).catch((caught: unknown) => caught);
      expect(error, email).toBeInstanceOf(UserInvestigationError);
      expect(error, email).toMatchObject({ code: "user_not_found", status: 404 });
      expect(sent, email).toHaveLength(0);
    }
  });
});

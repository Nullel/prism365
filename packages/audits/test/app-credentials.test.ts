import { decodeOwnerList, queryAuditRows } from "@prism365/contracts";
import { GraphRetryDeferredError } from "@prism365/graph";
import { describe, expect, it } from "vitest";
import {
  analyseCredentials,
  appCredentialsAudit,
  auditDefinition,
  isInsecureRedirect,
  isLocalhostRedirect,
  MAX_APPLICATIONS,
  MAX_REQUESTS,
} from "../src/index.ts";
import {
  daysAgo,
  expectValidTable,
  fake,
  find,
  now,
  paged,
  type Route,
} from "./apps-devices-usage-support.ts";

// A synthetic tenant: invented applications, example.invalid hosts, made-up IDs. `now` is fixed by
// the support file, so every date here is a number of days from it.
const id = (n: number) => `b1c2d3e4-0000-4000-8000-${String(n).padStart(12, "0")}`;
const DAY = 86_400_000;
const inDays = (days: number) => new Date(now.getTime() + days * DAY).toISOString();

interface Cred {
  startDateTime: string | null;
  endDateTime: string | null;
}
// A credential that started `startedDaysAgo` days ago and ends `endsInDays` days from now, or never.
const cred = (startedDaysAgo: number, endsInDays: number | null): Cred => ({
  startDateTime: daysAgo(startedDaysAgo),
  endDateTime: endsInDays === null ? null : inDays(endsInDays),
});

const app = (n: number, over: Record<string, unknown> = {}): Record<string, unknown> => ({
  id: id(n),
  appId: id(1_000 + n),
  displayName: `Quillnet app ${n}`,
  createdDateTime: daysAgo(500),
  signInAudience: "AzureADMyOrg",
  passwordCredentials: [],
  keyCredentials: [],
  web: { redirectUris: [] },
  spa: { redirectUris: [] },
  publicClient: { redirectUris: [] },
  // Fields the audit does not ask for are tolerated.
  tags: ["x"],
  ...over,
});

const withSecret = (n: number, ends: number | null, started = 100, over = {}) =>
  app(n, { passwordCredentials: [cred(started, ends)], ...over });

const user = (n: number, extra: Record<string, unknown> = {}) => ({
  "@odata.type": "#microsoft.graph.user",
  id: id(5_000 + n),
  displayName: `Person ${n}`,
  userPrincipalName: `person${n}@example.invalid`,
  ...extra,
});

interface Tenant {
  apps: unknown[];
  principals?: unknown[];
  owners?: Record<string, unknown[]>;
}

const principal = (n: number, over: Record<string, unknown> = {}) => ({
  id: id(2_000 + n),
  appId: id(1_000 + n),
  displayName: `Quillnet app ${n}`,
  servicePrincipalType: "Application",
  passwordCredentials: [],
  keyCredentials: [],
  ...over,
});

function route(tenant: Tenant): Route {
  return (url) => {
    const path = url.pathname.replace("/v1.0/", "");
    if (path === "applications") return paged(url, tenant.apps);
    if (path === "servicePrincipals") return paged(url, tenant.principals ?? []);
    const owners = /^applications\/([^/]+)\/owners$/.exec(path);
    if (owners) return paged(url, tenant.owners?.[owners[1] ?? ""] ?? [user(1)]);
    return undefined;
  };
}

const granted = ["Application.Read.All"];
const run = (tenant: Tenant) =>
  appCredentialsAudit.run(fake(appCredentialsAudit, route(tenant), { granted }).context);
const note = (result: { notes: string[] }) => result.notes.join("\n");
const cellsOf = async (tenant: Tenant, n: number) => find((await run(tenant)).rows, id(n)).cells;

describe("app credentials: definition", () => {
  it("publishes a valid table with every tag filterable", async () => {
    await expectValidTable(
      appCredentialsAudit,
      route({
        apps: [
          withSecret(1, -5),
          withSecret(2, 12, 100, { web: { redirectUris: ["http://kiosk.example.invalid/in"] } }),
          withSecret(3, null, 900, {
            signInAudience: "AzureADMultipleOrgs",
            web: { redirectUris: ["https://localhost:5001/cb"] },
          }),
          app(4),
          app(5, {
            passwordCredentials: [1, 2, 3, 4, 5].map(() => cred(10, 400)),
          }),
        ],
        owners: { [id(2)]: [] },
      }),
      { granted },
    );
  });

  it("declares the application pack alone, with nothing optional, scoped or handed off", () => {
    expect(auditDefinition(appCredentialsAudit)).toMatchObject({
      id: "apps.credentials",
      title: "App credentials",
      permissions: ["Application.Read.All"],
      optionalPermissions: [],
      scoped: false,
      handoff: false,
    });
    const columns = auditDefinition(appCredentialsAudit).columns;
    const column = (key: string) => columns.find((item) => item.key === key);
    for (const key of ["ownerIds", "servicePrincipalId"])
      expect(column(key)?.internal, key).toBe(true);
    expect(column("appId")?.hidden).toBe(true);
    expect(column("status")?.facet).toBe(true);
    expect(column("audience")?.facet).toBe(true);
  });
});

describe("app credentials: status and expiry", () => {
  const statusOf = async (ends: number | null, started = 100) => {
    const cells = await cellsOf({ apps: [withSecret(1, ends, started)] }, 1);
    return cells;
  };

  it("counts 30 days as within 30 and 31 days as within 90", async () => {
    expect((await statusOf(30)).status).toBe("Expires within 30 days");
    expect((await statusOf(30)).expiresIn).toBe(30);
    expect((await statusOf(30.5)).status).toBe("Expires within 90 days");
    expect((await statusOf(31)).status).toBe("Expires within 90 days");
    expect((await statusOf(29)).status).toBe("Expires within 30 days");
  });

  it("counts 90 days as within 90 and 91 days as fine", async () => {
    expect((await statusOf(90)).status).toBe("Expires within 90 days");
    expect((await statusOf(90.5)).status).toBe("OK");
    expect((await statusOf(91)).status).toBe("OK");
    expect((await statusOf(89)).status).toBe("Expires within 90 days");
  });

  it("treats a credential as expired once its end date has passed, and not before", async () => {
    const justPast = withSecret(1, -1 / DAY);
    const exactly = withSecret(2, 0);
    const result = await run({ apps: [justPast, exactly] });
    expect(find(result.rows, id(1)).cells).toMatchObject({
      status: "Expired",
      expired: 1,
      secrets: 0,
      expiresIn: null,
      nextExpiry: null,
    });
    expect(find(result.rows, id(2)).cells).toMatchObject({
      status: "Expires within 30 days",
      expired: 0,
      secrets: 1,
      expiresIn: 0,
    });
  });

  it("does not count an expired credential anywhere but as expired", async () => {
    const mixed = app(1, {
      passwordCredentials: [cred(900, -10), cred(100, 200)],
      keyCredentials: [cred(800, -400)],
    });
    const result = await run({ apps: [mixed] });
    const row = find(result.rows, id(1));
    expect(row.cells).toMatchObject({
      secrets: 1,
      certificates: 0,
      expired: 2,
      expiresIn: 200,
      // Rotation leftovers beside a working credential are not a broken app.
      status: "Expired leftovers",
      // The oldest credential that is still active, not the expired one from 900 days ago.
      oldestAge: 100,
    });
    expect(row.cells.nextExpiry).toBe(inDays(200));
    expect(row.tags).toContain("expired-leftover");
    expect(row.tags).not.toContain("expired");
    expect(row.tags).not.toContain("expiring-90");
    expect(row.tags).not.toContain("old-secret");
  });

  it("takes the soonest end among the active credentials of both kinds", async () => {
    const both = app(1, {
      passwordCredentials: [cred(100, 80)],
      keyCredentials: [cred(100, 20)],
    });
    const cells = await cellsOf({ apps: [both] }, 1);
    expect(cells).toMatchObject({ secrets: 1, certificates: 1, expiresIn: 20 });
    expect(cells.nextExpiry).toBe(inDays(20));
  });

  it("tags the 30 day window inside the 90 day window", async () => {
    const result = await run({ apps: [withSecret(1, 12), withSecret(2, 60), withSecret(3, 200)] });
    expect(find(result.rows, id(1)).tags).toEqual(
      expect.arrayContaining(["expiring-30", "expiring-90"]),
    );
    const sixty = find(result.rows, id(2)).tags;
    expect(sixty).toContain("expiring-90");
    expect(sixty).not.toContain("expiring-30");
    expect(find(result.rows, id(3)).tags.some((tag) => tag.startsWith("expiring"))).toBe(false);
  });

  it("says no credentials only for an app with none, expired or not", async () => {
    const result = await run({ apps: [app(1), withSecret(2, -3)] });
    expect(find(result.rows, id(1))).toMatchObject({
      tags: expect.arrayContaining(["no-credentials"]),
      cells: { status: "No credentials", secrets: 0, certificates: 0, nextExpiry: null },
    });
    expect(find(result.rows, id(2)).tags).not.toContain("no-credentials");
  });

  it("leaves the next expiry empty when every active credential never ends", async () => {
    const cells = await cellsOf({ apps: [withSecret(1, null)] }, 1);
    expect(cells).toMatchObject({ nextExpiry: null, expiresIn: null, status: "OK" });
  });
});

describe("app credentials: never expires, old secrets and many secrets", () => {
  const tagsOf = async (application: Record<string, unknown>) =>
    find((await run({ apps: [application] })).rows, id(1)).tags;

  it("tags an active credential with no end date", async () => {
    expect(await tagsOf(withSecret(1, null, 10))).toContain("never-expires");
    expect(await tagsOf(withSecret(1, 400, 10))).not.toContain("never-expires");
    // An expired credential is not active, whatever else it is.
    expect(await tagsOf(withSecret(1, -5, 10))).not.toContain("never-expires");
  });

  it("calls an active secret old only after 730 days", async () => {
    expect(await tagsOf(withSecret(1, 100, 731))).toContain("old-secret");
    expect(await tagsOf(withSecret(1, 100, 730))).not.toContain("old-secret");
    expect(await tagsOf(withSecret(1, 100, 729))).not.toContain("old-secret");
  });

  it("calls a secret old whether or not it ends", async () => {
    expect(await tagsOf(withSecret(1, null, 1_000))).toContain("old-secret");
    expect(await tagsOf(withSecret(1, 300, 1_000))).toContain("old-secret");
  });

  it("calls a certificate old only when it never ends", async () => {
    const certificate = (ends: number | null) => app(1, { keyCredentials: [cred(1_000, ends)] });
    expect(await tagsOf(certificate(null))).toContain("old-secret");
    expect(await tagsOf(certificate(300))).not.toContain("old-secret");
  });

  it("ignores an expired secret when looking for old ones", async () => {
    expect(await tagsOf(withSecret(1, -2, 1_000))).not.toContain("old-secret");
  });

  it("flags five active secrets and not four, and does not count expired ones", async () => {
    const secrets = (count: number, ends = 300) =>
      app(1, { passwordCredentials: Array.from({ length: count }, () => cred(20, ends)) });
    expect(await tagsOf(secrets(5))).toContain("many-secrets");
    expect(await tagsOf(secrets(4))).not.toContain("many-secrets");
    expect(await tagsOf(secrets(5, -1))).not.toContain("many-secrets");
    // Certificates are not secrets.
    expect(
      await tagsOf(app(1, { keyCredentials: Array.from({ length: 6 }, () => cred(20, 300)) })),
    ).not.toContain("many-secrets");
  });

  it("reports the age of the oldest active credential in whole days", async () => {
    const cells = await cellsOf(
      { apps: [app(1, { passwordCredentials: [cred(40.5, 300), cred(12, 300)] })] },
      1,
    );
    expect(cells.oldestAge).toBe(40);
    expect((await cellsOf({ apps: [app(2)] }, 2)).oldestAge).toBeNull();
  });

  it("does not guess at a credential without a start date", async () => {
    const odd = app(1, {
      passwordCredentials: [{ startDateTime: null, endDateTime: inDays(10) }],
    });
    const row = find((await run({ apps: [odd] })).rows, id(1));
    expect(row.cells.oldestAge).toBeNull();
    expect(row.tags).not.toContain("old-secret");
  });
});

describe("app credentials: enterprise application credentials", () => {
  it("adds credentials from the enterprise application with the same application ID", async () => {
    const tenant: Tenant = {
      apps: [withSecret(1, 200)],
      principals: [
        principal(1, {
          passwordCredentials: [cred(800, -20)],
          keyCredentials: [cred(100, 25)],
        }),
      ],
    };
    const cells = await cellsOf(tenant, 1);
    expect(cells).toMatchObject({
      secrets: 1,
      certificates: 1,
      expired: 1,
      expiresIn: 25,
      status: "Expires within 30 days",
      servicePrincipalId: id(2_001),
      appId: id(1_001),
    });
  });

  it("matches on the application ID whatever its letter case, and not on the display name", async () => {
    const tenant: Tenant = {
      apps: [app(1), app(2)],
      principals: [
        principal(1, { appId: id(1_001).toUpperCase(), keyCredentials: [cred(10, 10)] }),
        principal(2, {
          appId: id(9_999),
          displayName: "Quillnet app 2",
          keyCredentials: [cred(10, 10)],
        }),
      ],
    };
    const result = await run(tenant);
    expect(find(result.rows, id(1)).cells.certificates).toBe(1);
    expect(find(result.rows, id(2)).cells.certificates).toBe(0);
    expect(find(result.rows, id(2)).cells.servicePrincipalId).toBeNull();
  });

  it("leaves out managed identities and other types", async () => {
    const tenant: Tenant = {
      apps: [app(1)],
      principals: [
        principal(1, { servicePrincipalType: "ManagedIdentity", keyCredentials: [cred(10, 10)] }),
      ],
    };
    const cells = await cellsOf(tenant, 1);
    expect(cells).toMatchObject({ certificates: 0, servicePrincipalId: null });
  });

  it("asks for the fields it uses, a page of 100 at a time", async () => {
    const ctx = fake(appCredentialsAudit, route({ apps: [app(1)], principals: [principal(1)] }), {
      granted,
    });
    await appCredentialsAudit.run(ctx.context);
    const list = ctx.requests.find((url) => url.pathname === "/v1.0/applications");
    expect(list?.searchParams.get("$top")).toBe("100");
    expect(list?.searchParams.get("$select")).toBe(
      "id,appId,displayName,createdDateTime,signInAudience,passwordCredentials,keyCredentials,web,spa,publicClient",
    );
    const principals = ctx.requests.find((url) => url.pathname === "/v1.0/servicePrincipals");
    expect(principals?.searchParams.get("$top")).toBe("100");
    expect(principals?.searchParams.get("$select")).toBe(
      "id,appId,displayName,servicePrincipalType,passwordCredentials,keyCredentials",
    );
  });
});

describe("app credentials: audience", () => {
  it("maps each sign-in audience, and flags every one but this tenant", async () => {
    const audiences: [string, string | null, boolean][] = [
      ["AzureADMyOrg", "This tenant", false],
      ["AzureADMultipleOrgs", "Any organisation", true],
      ["AzureADandPersonalMicrosoftAccount", "Any organisation or personal account", true],
      ["PersonalMicrosoftAccount", "Personal accounts", true],
      ["somethingNew", null, false],
    ];
    const result = await run({
      apps: audiences.map(([audience], index) => app(index + 1, { signInAudience: audience })),
    });
    audiences.forEach(([, label, multi], index) => {
      const row = find(result.rows, id(index + 1));
      expect(row.cells.audience).toBe(label);
      expect(row.tags.includes("multi-tenant")).toBe(multi);
    });
  });
});

describe("app credentials: redirect URIs", () => {
  const redirects = async (web: string[], spa: string[] = [], publicClient: string[] = []) => {
    const row = find(
      (
        await run({
          apps: [
            app(1, {
              web: { redirectUris: web },
              spa: { redirectUris: spa },
              publicClient: { redirectUris: publicClient },
            }),
          ],
        })
      ).rows,
      id(1),
    );
    return { tags: row.tags, listed: row.cells.redirects };
  };

  it("flags plain http to another host as insecure", async () => {
    const found = await redirects(["http://kiosk.example.invalid/signin"]);
    expect(found.tags).toContain("insecure-redirect");
    expect(found.tags).not.toContain("localhost-redirect");
    expect(found.listed).toEqual(["http://kiosk.example.invalid/signin"]);
  });

  it("does not call plain http to this computer insecure, but lists it as localhost", async () => {
    for (const uri of ["http://localhost:3000/cb", "http://127.0.0.1/cb", "http://[::1]:8080/cb"]) {
      const found = await redirects([uri]);
      expect(found.tags, uri).toContain("localhost-redirect");
      expect(found.tags, uri).not.toContain("insecure-redirect");
      expect(found.listed, uri).toEqual([uri]);
    }
  });

  it("flags a wildcard whatever the scheme", async () => {
    for (const uri of ["https://*.example.invalid/cb", "http://localhost:*/cb"]) {
      const found = await redirects([uri]);
      expect(found.tags, uri).toContain("insecure-redirect");
      expect(found.listed, uri).toEqual([uri]);
    }
  });

  it("leaves clean https and app schemes alone", async () => {
    const found = await redirects(
      ["https://app.example.invalid/cb", "https://localhostel.example.invalid/cb"],
      ["https://spa.example.invalid/"],
      ["msauth.com.example.invalid://auth", "urn:ietf:wg:oauth:2.0:oob"],
    );
    expect(found.tags.filter((tag) => tag.endsWith("redirect"))).toEqual([]);
    expect(found.listed).toEqual([]);
  });

  it("reads web, single-page and public client URIs, and lists at most ten", async () => {
    const many = Array.from({ length: 14 }, (_, n) => `http://h${n}.example.invalid/cb`);
    const found = await redirects(many.slice(0, 5), many.slice(5, 9), many.slice(9));
    expect(found.tags).toContain("insecure-redirect");
    expect(found.listed).toHaveLength(10);
    const spaOnly = await redirects([], ["http://spa.example.invalid/"]);
    expect(spaOnly.tags).toContain("insecure-redirect");
    const publicOnly = await redirects([], [], ["http://pub.example.invalid/"]);
    expect(publicOnly.tags).toContain("insecure-redirect");
  });

  it("classifies single addresses", () => {
    expect(isInsecureRedirect("http://a.example.invalid")).toBe(true);
    expect(isInsecureRedirect("HTTP://A.EXAMPLE.INVALID")).toBe(true);
    expect(isInsecureRedirect("https://a.example.invalid")).toBe(false);
    expect(isInsecureRedirect("http://localhost")).toBe(false);
    expect(isInsecureRedirect("http://localhost.example.invalid")).toBe(true);
    expect(isInsecureRedirect("https://*.example.invalid")).toBe(true);
    expect(isLocalhostRedirect("https://localhost:44300/cb")).toBe(true);
    expect(isLocalhostRedirect("https://localhost.example.invalid/cb")).toBe(false);
    expect(isLocalhostRedirect("not a uri")).toBe(false);
  });

  it("keeps an unusual value out of the way without failing the run", async () => {
    const result = await run({
      apps: [app(1, { web: { redirectUris: [42, null, "http://ok.example.invalid\u0000/x"] } })],
    });
    expect(result.complete).toBe(true);
    expect(find(result.rows, id(1)).cells.redirects).toEqual(["http://ok.example.invalid/x"]);
  });
});

describe("app credentials: owners", () => {
  it("reads the first page of owners and counts them", async () => {
    const tenant: Tenant = {
      apps: [withSecret(1, 200), withSecret(2, 200)],
      owners: { [id(1)]: [user(1), user(2), user(3)], [id(2)]: [] },
    };
    const ctx = fake(appCredentialsAudit, route(tenant), { granted });
    const result = await appCredentialsAudit.run(ctx.context);
    expect(find(result.rows, id(1)).cells).toMatchObject({
      owners: 3,
      ownerNames: ["Person 1", "Person 2", "Person 3"],
    });
    const owners = ctx.requests.find(
      (url) => url.pathname === `/v1.0/applications/${id(1)}/owners`,
    );
    expect(owners?.searchParams.get("$select")).toBe("id,displayName,userPrincipalName");
    // Only the first page is asked for.
    expect(owners?.searchParams.get("$skiptoken")).toBeNull();
  });

  it("tags an app whose owners were read and are none", async () => {
    const result = await run({
      apps: [withSecret(1, 200), withSecret(2, 200)],
      owners: { [id(1)]: [], [id(2)]: [user(1)] },
    });
    const none = find(result.rows, id(1));
    expect(none.tags).toContain("no-owner");
    expect(none.cells.owners).toBe(0);
    expect(find(result.rows, id(2)).tags).not.toContain("no-owner");
  });

  it("counts an owner that is not an account but does not name it", async () => {
    const result = await run({
      apps: [withSecret(1, 200)],
      owners: {
        [id(1)]: [
          user(1),
          {
            "@odata.type": "#microsoft.graph.servicePrincipal",
            id: id(7_001),
            displayName: "Robot owner",
          },
          { "@odata.type": "#microsoft.graph.group", id: id(7_002), displayName: "Some group" },
          // No type, but a sign-in name: only accounts have one.
          { id: id(7_003), displayName: "Untyped", userPrincipalName: "untyped@example.invalid" },
          // No type and no sign-in name: not claimed to be an account.
          { id: id(7_004), displayName: "Mystery" },
        ],
      },
    });
    const cells = find(result.rows, id(1)).cells;
    expect(cells.owners).toBe(5);
    expect(cells.ownerNames).toEqual(["Person 1", "Untyped"]);
    expect(decodeOwnerList(cells.ownerIds)).toEqual([
      { id: id(5_001), name: "Person 1", kind: "user" },
      { id: id(7_001), name: null, kind: "servicePrincipal" },
      { id: id(7_002), name: null, kind: "other" },
      { id: id(7_003), name: "Untyped", kind: "user" },
      { id: id(7_004), name: null, kind: "other" },
    ]);
  });

  it("names at most five owners", async () => {
    const result = await run({
      apps: [withSecret(1, 200)],
      owners: { [id(1)]: [7, 6, 5, 4, 3, 2, 1].map((n) => user(n)) },
    });
    const cells = find(result.rows, id(1)).cells;
    expect(cells.owners).toBe(7);
    expect(cells.ownerNames).toEqual(["Person 1", "Person 2", "Person 3", "Person 4", "Person 5"]);
  });

  it("reads owners only for apps with a credential or another flag", async () => {
    const tenant: Tenant = {
      apps: [
        app(1), // nothing: no credential, this tenant, no redirect
        withSecret(2, 200),
        app(3, { signInAudience: "AzureADMultipleOrgs" }),
        app(4, { web: { redirectUris: ["https://localhost/cb"] } }),
        app(5, { web: { redirectUris: ["http://x.example.invalid/cb"] } }),
        app(6, { passwordCredentials: [cred(900, -10)] }),
      ],
    };
    const ctx = fake(appCredentialsAudit, route(tenant), { granted });
    const result = await appCredentialsAudit.run(ctx.context);
    const asked = ctx.requests
      .map((url) => /^\/v1\.0\/applications\/([^/]+)\/owners$/.exec(url.pathname)?.[1])
      .filter((value): value is string => Boolean(value))
      .sort();
    expect(asked).toEqual([2, 3, 4, 5, 6].map(id));
    const quiet = find(result.rows, id(1));
    expect(quiet.cells.owners).toBeNull();
    // Never meant to be read, so not reported as unread.
    expect(quiet.tags).not.toContain("owners-not-read");
    expect(result.complete).toBe(true);
  });

  describe("under the request limit", () => {
    // 700 apps hold a secret, so every one wants owners. The lists cost 7 pages of applications
    // and 1 of service principals, which leaves the rest of the 600 for owners.
    const total = 700;
    const crowd = Array.from({ length: total }, (_, index) =>
      // The last ten have expired; they are read first whatever their position.
      withSecret(index + 1, index >= total - 10 ? -5 : 200),
    );

    it("marks the applications it did not get to, makes the run partial, and says so", async () => {
      const ctx = fake(appCredentialsAudit, route({ apps: crowd }), { granted });
      const result = await appCredentialsAudit.run(ctx.context);
      expect(ctx.requests).toHaveLength(MAX_REQUESTS);
      const unread = result.rows.filter((row) => row.tags.includes("owners-not-read"));
      expect(unread).toHaveLength(total - (MAX_REQUESTS - 8));
      for (const row of unread) {
        expect(row.cells.owners).toBeNull();
        expect(row.cells.ownerNames).toBeNull();
        expect(row.tags).not.toContain("no-owner");
      }
      expect(result.complete).toBe(false);
      expect(note(result)).toContain("before reaching the limit of 600 requests per run");
      expect(note(result)).toContain("108 applications show Owners not read");
    });

    it("reads the expired applications first", async () => {
      const result = await run({ apps: crowd });
      for (let n = total - 9; n <= total; n += 1)
        expect(find(result.rows, id(n)).cells.owners, `app ${n}`).toBe(1);
    });

    it("does not call a tile zero when owners were not all read", async () => {
      const result = await run({ apps: crowd, owners: {} });
      const owners = result.summary.find((item) => item.label === "No owner");
      expect(owners?.value).toBe(0);
      expect(owners?.tone).toBe("neutral");
    });
  });

  it("keeps what it read when Microsoft asks it to slow down", async () => {
    const tenant: Tenant = { apps: [withSecret(1, 200), withSecret(2, 201), withSecret(3, 202)] };
    const inner = route(tenant);
    const result = await appCredentialsAudit.run(
      fake(
        appCredentialsAudit,
        (url) =>
          url.pathname === `/v1.0/applications/${id(2)}/owners`
            ? new Response("{}", { status: 429 })
            : inner(url),
        { granted },
      ).context,
    );
    expect(result.complete).toBe(false);
    expect(find(result.rows, id(1)).cells.owners).toBe(1);
    expect(find(result.rows, id(2)).tags).toContain("owners-not-read");
    expect(find(result.rows, id(3)).tags).toContain("owners-not-read");
    expect(note(result)).toContain("slow down after the owners of 1 of 3 applications");
  });

  it("stops reading owners when Microsoft refuses, and keeps the credentials", async () => {
    const tenant: Tenant = { apps: [withSecret(1, 200), withSecret(2, 201)] };
    const inner = route(tenant);
    const result = await appCredentialsAudit.run(
      fake(
        appCredentialsAudit,
        (url) =>
          /\/owners$/.test(url.pathname) ? new Response("{}", { status: 403 }) : inner(url),
        { granted },
      ).context,
    );
    expect(result.complete).toBe(false);
    expect(note(result)).toContain("Microsoft refused to list owners");
    expect(find(result.rows, id(1)).cells).toMatchObject({ secrets: 1, owners: null });
  });

  it("tolerates a registration deleted after it was listed", async () => {
    const tenant: Tenant = { apps: [withSecret(1, 200), withSecret(2, 201)] };
    const inner = route(tenant);
    const result = await appCredentialsAudit.run(
      fake(
        appCredentialsAudit,
        (url) =>
          url.pathname === `/v1.0/applications/${id(1)}/owners`
            ? new Response("{}", { status: 404 })
            : inner(url),
        { granted },
      ).context,
    );
    expect(result.complete).toBe(true);
    expect(find(result.rows, id(1)).tags).not.toContain("owners-not-read");
    expect(find(result.rows, id(2)).cells.owners).toBe(1);
  });
});

describe("app credentials: reading limits", () => {
  it("stops at 2,000 applications and says the run is partial", async () => {
    const apps = Array.from({ length: MAX_APPLICATIONS + 100 }, (_, index) => app(index + 1));
    const ctx = fake(appCredentialsAudit, route({ apps }), { granted });
    const result = await appCredentialsAudit.run(ctx.context);
    expect(ctx.requests.filter((url) => url.pathname === "/v1.0/applications")).toHaveLength(20);
    expect(result.rows).toHaveLength(MAX_APPLICATIONS);
    expect(result.complete).toBe(false);
    expect(note(result)).toContain("Stopped after 2,000 applications");
  });

  it("is complete at exactly 2,000 applications", async () => {
    const apps = Array.from({ length: MAX_APPLICATIONS }, (_, index) => app(index + 1));
    const result = await run({ apps });
    expect(result.rows).toHaveLength(MAX_APPLICATIONS);
    expect(result.complete).toBe(true);
  });

  it("says when the service principals were cut short", async () => {
    const principals = Array.from({ length: 2_100 }, (_, index) =>
      principal(index + 1, { servicePrincipalType: "ManagedIdentity" }),
    );
    const result = await run({ apps: [app(1)], principals });
    expect(result.complete).toBe(false);
    expect(note(result)).toContain("Stopped after 2,000 service principals");
  });

  it("keeps the applications it read when Microsoft throttles a later page", async () => {
    const apps = Array.from({ length: 250 }, (_, index) => app(index + 1));
    const inner = route({ apps });
    const result = await appCredentialsAudit.run(
      fake(
        appCredentialsAudit,
        (url) =>
          url.pathname === "/v1.0/applications" && url.searchParams.get("$skiptoken") === "100"
            ? new Response("{}", { status: 429 })
            : inner(url),
        { granted },
      ).context,
    );
    expect(result.rows).toHaveLength(100);
    expect(result.complete).toBe(false);
    expect(note(result)).toContain("slow down after 100 applications");
  });

  it("fails when Microsoft throttles the very first read", async () => {
    await expect(
      appCredentialsAudit.run(
        fake(appCredentialsAudit, () => new Response("{}", { status: 429 }), { granted }).context,
      ),
    ).rejects.toBeInstanceOf(GraphRetryDeferredError);
  });

  it("leaves out a record it cannot read, marks the run partial, and keeps the rest", async () => {
    const result = await run({
      apps: [app(1), { id: "not a guid", displayName: "Broken" }, "nonsense", app(2)],
    });
    expect(result.rows.map((row) => row.key).sort()).toEqual([id(1), id(2)]);
    expect(result.complete).toBe(false);
    expect(note(result)).toContain("2 records could not be read");
  });

  it("does not guess at a credential with a date that is not one", async () => {
    const result = await run({
      apps: [
        app(1, {
          passwordCredentials: [
            { startDateTime: "yesterday-ish", endDateTime: null },
            cred(10, 100),
          ],
        }),
      ],
    });
    const cells = find(result.rows, id(1)).cells;
    // Not read as "never expires": the credential is left out and the run says so.
    expect(cells).toMatchObject({ secrets: 1, nextExpiry: inDays(100) });
    expect(find(result.rows, id(1)).tags).not.toContain("never-expires");
    expect(result.complete).toBe(false);
    expect(note(result)).toContain("1 credential");
  });

  it("lists the same application once", async () => {
    const result = await run({ apps: [app(1), app(1)] });
    expect(result.rows).toHaveLength(1);
  });
});

describe("app credentials: rows, tiles and notes", () => {
  const tenant: Tenant = {
    apps: [
      withSecret(1, -40, 400, { displayName: "Zebra Payroll" }),
      withSecret(2, 12, 350, { displayName: "badge sync" }),
      withSecret(3, 60),
      withSecret(4, null, 1_095, { displayName: "Legacy Importer" }),
      withSecret(5, 145, 80, {
        displayName: "Visitor Kiosk",
        web: { redirectUris: ["http://kiosk.example.invalid/signin"] },
      }),
      app(6, { displayName: "Partner Portal", signInAudience: "AzureADMultipleOrgs" }),
    ],
    owners: { [id(5)]: [] },
  };

  it("sets the cells, key and tags of a row", async () => {
    const result = await run(tenant);
    const row = find(result.rows, id(2));
    expect(row.personId).toBeNull();
    expect(row.cells).toMatchObject({
      app: "badge sync",
      audience: "This tenant",
      secrets: 1,
      certificates: 0,
      expired: 0,
      expiresIn: 12,
      oldestAge: 350,
      owners: 1,
      status: "Expires within 30 days",
      appId: id(1_002),
      servicePrincipalId: null,
    });
    expect(row.cells.created).toBe(daysAgo(500));
  });

  it("sorts rows by name, ignoring case", async () => {
    const result = await run(tenant);
    expect(result.rows.map((row) => row.cells.app)).toEqual([
      "badge sync",
      "Legacy Importer",
      "Partner Portal",
      "Quillnet app 3",
      "Visitor Kiosk",
      "Zebra Payroll",
    ]);
  });

  it("counts applications in each tile and points each at its view", async () => {
    const result = await run(tenant);
    const tiles = Object.fromEntries(result.summary.map((item) => [item.label, item]));
    expect(tiles["Expired credentials"]).toMatchObject({
      value: 1,
      tone: "critical",
      view: "expired",
    });
    expect(tiles["Expiring within 30 days"]).toMatchObject({
      value: 1,
      tone: "warning",
      view: "expiring-30",
    });
    expect(tiles["No owner"]).toMatchObject({ value: 1, tone: "warning", view: "no-owner" });
    expect(tiles["Secrets older than 2 years"]).toMatchObject({
      value: 1,
      tone: "warning",
      view: "old-secret",
    });
    expect(tiles["Insecure redirect URIs"]).toMatchObject({
      value: 1,
      tone: "warning",
      view: "insecure-redirect",
    });
    expect(tiles.Applications).toMatchObject({ value: 6, tone: "neutral", view: "all" });
  });

  it("shows good tiles on a tenant with nothing to flag", async () => {
    const result = await run({ apps: [withSecret(1, 300, 30)] });
    for (const label of [
      "Expired credentials",
      "Expiring within 30 days",
      "No owner",
      "Secrets older than 2 years",
      "Insecure redirect URIs",
    ]) {
      const tile = result.summary.find((item) => item.label === label);
      expect(tile, label).toMatchObject({ value: 0, tone: "good" });
    }
  });

  it("says how much it read, that values are never returned, and what active means", async () => {
    const result = await run({
      apps: [app(1), app(2)],
      principals: [principal(1), principal(9, { servicePrincipalType: "ManagedIdentity" })],
    });
    expect(note(result)).toContain("Read 2 applications and 1 enterprise application");
    expect(note(result)).toContain("Owners were read for 0 of 2 applications");
    expect(note(result)).toContain("limit of 600 requests per run");
    expect(note(result)).toContain("Microsoft never returns the value of a secret");
    expect(note(result)).toContain("not yet past its end date");
  });

  it("is complete when everything was read", async () => {
    expect((await run(tenant)).complete).toBe(true);
  });

  it("puts exactly the right rows in each view", async () => {
    const result = await run(tenant);
    const definition = auditDefinition(appCredentialsAudit);
    const keysIn = (view: string) =>
      queryAuditRows(definition, result.rows, { view })
        .map((item) => item.key)
        .sort();
    const only = (...numbers: number[]) => numbers.map(id).sort();
    expect(keysIn("all")).toEqual(only(1, 2, 3, 4, 5, 6));
    expect(keysIn("expired")).toEqual(only(1));
    expect(keysIn("expired-leftover")).toEqual([]);
    expect(keysIn("expiring-30")).toEqual(only(2));
    expect(keysIn("expiring-90")).toEqual(only(2, 3));
    expect(keysIn("no-owner")).toEqual(only(5));
    expect(keysIn("old-secret")).toEqual(only(4));
    expect(keysIn("never-expires")).toEqual(only(4));
    expect(keysIn("insecure-redirect")).toEqual(only(5));
    expect(keysIn("multi-tenant")).toEqual(only(6));
    expect(keysIn("no-credentials")).toEqual(only(6));
    expect(keysIn("many-secrets")).toEqual([]);
    expect(keysIn("owners-not-read")).toEqual([]);
    expect(keysIn("localhost-redirect")).toEqual([]);
  });
});

describe("app credentials: expired leftovers", () => {
  it("separates a broken app from a rotation leftover, in tag, status and tile", async () => {
    const result = await run({
      apps: [
        withSecret(1, -5),
        app(2, { passwordCredentials: [cred(500, -5), cred(30, 400)] }),
        app(3, { passwordCredentials: [cred(500, -5), cred(30, 10)] }),
      ],
    });
    expect(find(result.rows, id(1)).tags).toContain("expired");
    expect(find(result.rows, id(1)).tags).not.toContain("expired-leftover");
    expect(find(result.rows, id(2)).tags).toContain("expired-leftover");
    expect(find(result.rows, id(2)).tags).not.toContain("expired");
    expect(find(result.rows, id(2)).cells).toMatchObject({
      status: "Expired leftovers",
      expired: 1,
    });
    // The near expiry is the more urgent status; the leftover is still recorded.
    expect(find(result.rows, id(3)).cells.status).toBe("Expires within 30 days");
    expect(find(result.rows, id(3)).tags).toContain("expired-leftover");
    const tiles = Object.fromEntries(result.summary.map((item) => [item.label, item]));
    expect(tiles["Expired credentials"]).toMatchObject({ value: 1, tone: "critical" });
    expect(tiles["Expired credentials left in place"]).toMatchObject({
      value: 2,
      tone: "neutral",
      view: "expired-leftover",
    });
  });
});

describe("app credentials: far-future end dates", () => {
  const far = "2299-12-31T00:00:00Z";
  const tagsFor = async (application: Record<string, unknown>) => {
    const row = find((await run({ apps: [application] })).rows, id(1));
    return row;
  };

  it("treats a secret ending in the year 2299 like one with no end date", async () => {
    const secret = (ends: string | null) =>
      app(1, { passwordCredentials: [{ startDateTime: daysAgo(900), endDateTime: ends }] });
    const forever = await tagsFor(secret(far));
    const none = await tagsFor(secret(null));
    for (const row of [forever, none]) {
      expect(row.tags).toEqual(expect.arrayContaining(["never-expires", "old-secret"]));
      expect(row.tags.some((tag) => tag.startsWith("expiring"))).toBe(false);
      expect(row.cells).toMatchObject({
        nextExpiry: null,
        expiresIn: null,
        status: "OK",
        secrets: 1,
      });
    }
    expect(forever.tags).toEqual(none.tags);
  });

  it("does the same for a certificate, and ages a recent one by its start", async () => {
    const certificate = (started: number) =>
      app(1, { keyCredentials: [{ startDateTime: daysAgo(started), endDateTime: far }] });
    const old = await tagsFor(certificate(900));
    expect(old.tags).toEqual(expect.arrayContaining(["never-expires", "old-secret"]));
    expect(old.cells).toMatchObject({ certificates: 1, nextExpiry: null });
    const recent = await tagsFor(certificate(100));
    expect(recent.tags).toContain("never-expires");
    expect(recent.tags).not.toContain("old-secret");
  });

  it("does not stretch the rule to an end date a few years away", async () => {
    const row = await tagsFor(
      app(1, {
        passwordCredentials: [{ startDateTime: daysAgo(10), endDateTime: inDays(365 * 5) }],
      }),
    );
    expect(row.tags).not.toContain("never-expires");
    expect(row.cells.nextExpiry).toBe(inDays(365 * 5));
  });
});

describe("app credentials: SAML certificates on the enterprise application", () => {
  const thumb = (label: string) => `thumb-${label}`;
  const key = (usage: string, label: string, started: number, ends: number) => ({
    startDateTime: daysAgo(started),
    endDateTime: inDays(ends),
    type: "AsymmetricX509Cert",
    usage,
    customKeyIdentifier: thumb(label),
  });
  const keyPassword = (label: string, started: number, ends: number) => ({
    startDateTime: daysAgo(started),
    endDateTime: inDays(ends),
    customKeyIdentifier: thumb(label),
  });
  const saml = (extra: Record<string, unknown> = {}) =>
    principal(1, {
      keyCredentials: [
        key("Verify", "new", 30, 700),
        key("Sign", "new", 30, 700),
        key("Verify", "old", 900, -20),
        key("Sign", "old", 900, -20),
      ],
      passwordCredentials: [keyPassword("new", 30, 700), keyPassword("old", 900, -20)],
      ...extra,
    });

  it("counts each certificate once, ignores the private key's password, and ages nothing from it", async () => {
    const result = await run({ apps: [app(1)], principals: [saml()] });
    const row = find(result.rows, id(1));
    expect(row.cells).toMatchObject({
      certificates: 1,
      secrets: 0,
      expired: 1,
      oldestAge: 30,
      status: "Expired leftovers",
    });
    expect(row.tags).toContain("expired-leftover");
    expect(row.tags).not.toContain("old-secret");
    expect(row.tags).not.toContain("expired");
  });

  it("counts a password that pairs with nothing, and a Sign key with no Verify twin", async () => {
    const result = await run({
      apps: [app(1)],
      principals: [
        principal(1, {
          keyCredentials: [key("Sign", "lone", 40, 300)],
          passwordCredentials: [keyPassword("unrelated", 800, 300)],
        }),
      ],
    });
    const row = find(result.rows, id(1));
    expect(row.cells).toMatchObject({ certificates: 1, secrets: 1 });
    expect(row.tags).toContain("old-secret");
  });

  it("does not look at thumbprints on the registration itself", async () => {
    const result = await run({
      apps: [
        app(1, {
          keyCredentials: [key("Verify", "x", 10, 300), key("Sign", "x", 10, 300)],
        }),
      ],
    });
    expect(find(result.rows, id(1)).cells.certificates).toBe(2);
  });

  it("never puts a thumbprint in a row", async () => {
    const result = await run({ apps: [app(1)], principals: [saml()] });
    expect(JSON.stringify(result)).not.toContain("thumb-");
  });
});

describe("analyseCredentials", () => {
  const at = (days: number) => now.getTime() + days * DAY;
  const credential = (start: number | null, end: number | null) => ({
    start: start === null ? null : at(start),
    end: end === null ? null : at(end),
  });

  it("returns an empty, fine answer for no credentials", () => {
    expect(analyseCredentials([], [], now.getTime())).toMatchObject({
      secrets: 0,
      certificates: 0,
      expired: 0,
      nextExpiry: null,
      status: "No credentials",
      oldSecret: false,
      neverExpires: false,
    });
  });

  it("calls an app Expired only when no credential is left that works", () => {
    const broken = analyseCredentials(
      [credential(-100, -1)],
      [credential(-100, -5)],
      now.getTime(),
    );
    expect(broken).toMatchObject({ status: "Expired", broken: true, leftover: false });
    const rotated = analyseCredentials(
      [credential(-100, -1), credential(-5, 400)],
      [],
      now.getTime(),
    );
    expect(rotated).toMatchObject({ status: "Expired leftovers", broken: false, leftover: true });
    // A near expiry outranks a leftover in the status; the leftover tag stays.
    const near = analyseCredentials([credential(-100, 5), credential(-100, -1)], [], now.getTime());
    expect(near).toMatchObject({ status: "Expires within 30 days", expiresIn: 5, leftover: true });
  });
});

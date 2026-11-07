import { GraphHttpClient } from "@prism365/graph";
import { ModuleScheduler } from "@prism365/modules";
import {
  MemoryStateStore,
  type ModuleDatabase,
  PGliteModuleDatabase,
  type SqlClient,
} from "@prism365/storage";
import { afterAll, beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import { currentImport } from "../src/mailbox-access-store.ts";
import { createModuleRegistry } from "../src/modules.ts";

vi.setConfig({ testTimeout: 30_000 });

// Shared mailbox access from an imported Exchange export (ADR-0012). Synthetic tenant,
// administrator, accounts, mailboxes and groups only; every address ends in .invalid.
const tenantId = "11111111-1111-4111-8111-111111111111";
const otherTenant = "99999999-9999-4999-8999-999999999999";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";

const PERSON = "aaaaaaaa-0000-4000-8000-000000000001";
const OTHER_PERSON = "aaaaaaaa-0000-4000-8000-000000000002";
const LAB = "eeeeeeee-0000-4000-8000-000000000001";
const BILLING = "eeeeeeee-0000-4000-8000-000000000002";
// The person is in this group only through another group: transitiveMemberOf lists it, memberOf
// would not.
const NESTED_GROUP = "bbbbbbbb-0000-4000-8000-000000000003";
const OTHER_GROUP = "bbbbbbbb-0000-4000-8000-000000000009";
const SID = "S-1-5-21-1111111111-2222222222-3333333333-1105";

const DAY = 86_400_000;
const ago = (days: number) => new Date(Date.now() - days * DAY).toISOString();

const user = (id: string, name: string, address: string) => ({
  id,
  address,
  name,
  kind: "user",
});
const exportFile = (overrides: Record<string, unknown> = {}) => ({
  format: "prism365.mailbox-access",
  version: 1,
  tenantId,
  exportedAt: ago(2),
  includeUserMailboxes: false,
  mailboxes: [
    {
      id: LAB,
      address: "lab.results@example.invalid",
      displayName: "Lab results",
      type: "SharedMailbox",
    },
    {
      id: BILLING,
      address: "billing@example.invalid",
      displayName: "Billing",
      type: "SharedMailbox",
    },
  ],
  grants: [
    {
      mailboxId: LAB,
      right: "FullAccess",
      trustee: {
        id: NESTED_GROUP,
        address: "plant.staff@example.invalid",
        name: "All plant staff",
        kind: "group",
      },
    },
    {
      mailboxId: LAB,
      right: "SendAs",
      trustee: user(PERSON, "Pat Pell", "pat.pell@example.invalid"),
    },
    {
      mailboxId: LAB,
      right: "FullAccess",
      trustee: { id: null, address: null, name: SID, kind: "unresolved" },
    },
    {
      mailboxId: BILLING,
      right: "SendOnBehalf",
      trustee: user(PERSON, "Pat Pell", "pat.pell@example.invalid"),
    },
    {
      mailboxId: BILLING,
      right: "FullAccess",
      trustee: user(OTHER_PERSON, "Robin Reyes", "robin.reyes@example.invalid"),
    },
    {
      mailboxId: BILLING,
      right: "FullAccess",
      trustee: { id: OTHER_GROUP, address: null, name: "Finance", kind: "group" },
    },
  ],
  unresolved: 1,
  ...overrides,
});

// Switches on the module database the app sees, to show a failed grant keeps the old import and
// that the ledger entry comes before any write.
let failGrantBatchFrom: number | null = null;
let ledgerAtFirstWrite: number | null = null;
let currentStore: MemoryStateStore | null = null;
// Runs once while the person route waits for Microsoft to list the groups, as another
// administrator's import or clear would.
let duringGroupRead: (() => Promise<unknown>) | null = null;
function watched(database: ModuleDatabase): ModuleDatabase {
  return {
    kind: database.kind,
    durable: database.durable,
    tenant: (tenant, operation, options) =>
      database.tenant(
        tenant,
        (sql) =>
          operation({
            query: <T extends object>(text: string, params?: readonly unknown[]) => {
              if (/^\s*(INSERT|DELETE)[\s\S]*prism365_mailbox_access/.test(text)) {
                ledgerAtFirstWrite ??= currentStore?.audit.length ?? 0;
              }
              if (
                failGrantBatchFrom !== null &&
                text.includes("INSERT INTO prism365_mailbox_access_grant") &&
                Number(params?.[2]) >= failGrantBatchFrom
              )
                return Promise.reject(new Error("Synthetic storage failure."));
              return sql.query<T>(text, params);
            },
          } satisfies SqlClient),
        options,
      ),
    verify: () => database.verify(),
    close: () => database.close(),
  };
}

type Handler = (url: URL) => Response | undefined;
async function harness(
  options: {
    features?: readonly string[];
    scopes?: readonly string[];
    handler?: Handler;
    modules?: boolean;
  } = {},
) {
  const features = options.features ?? ["directory.users.read", "audit.signins.read"];
  const scopes = options.scopes ?? ["User.Read", "User.Read.All", "AuditLog.Read.All"];
  const requests: string[] = [];
  let ledgerAtFirstRequest: number | null = null;
  const store = new MemoryStateStore();
  const handler: Handler =
    options.handler ??
    ((url) => {
      const path = url.pathname.replace("/v1.0/", "");
      if (path === `users/${PERSON}/transitiveMemberOf`)
        return Response.json({
          value: [
            { "@odata.type": "#microsoft.graph.group", id: "bbbbbbbb-0000-4000-8000-000000000001" },
            { "@odata.type": "#microsoft.graph.group", id: NESTED_GROUP.toUpperCase() },
            { "@odata.type": "#microsoft.graph.directoryRole", id: "not-a-guid" },
          ],
        });
      if (path.endsWith("/transitiveMemberOf")) return Response.json({ value: [] });
      return undefined;
    });
  const graphFetch = (async (input: RequestInfo | URL) => {
    const url = new URL(String(input));
    requests.push(`${url.pathname}${url.search}`);
    ledgerAtFirstRequest ??= store.audit.length;
    if (duringGroupRead && url.pathname.endsWith("/transitiveMemberOf")) {
      const meanwhile = duringGroupRead;
      duringGroupRead = null;
      await meanwhile();
    }
    return handler(url) ?? new Response("{}", { status: 404 });
  }) as typeof fetch;
  const identityConfig: IdentityConfig = {
    tenantId,
    clientId: "33333333-3333-4333-8333-333333333333",
    origin,
    redirectUri: `${origin}/api/v1/auth/callback`,
    allowedActorIds: [actorId],
    enabledFeatures: features as IdentityConfig["enabledFeatures"],
    credential: { clientSecret: "synthetic-only" },
  };
  const config: AppConfig = {
    mode: "live",
    origin,
    identity: identityConfig,
    production: false,
    missingConfiguration: [],
    enabledFeatures: features as AppConfig["enabledFeatures"],
  };
  const provider: MicrosoftIdentityProvider = {
    authorizationUrl: async () => "https://login.microsoftonline.com/synthetic",
    redeem: async () => ({
      tenantId,
      actorId,
      displayName: "Synthetic admin",
      accountId: "synthetic-account",
      cache: "synthetic-cache",
      scopes: [...scopes],
    }),
    token: async (session) => ({ accessToken: "synthetic", cache: "c", scopes: session.scopes }),
  };
  const identity = new IdentityService(identityConfig, store, provider);
  const registry = createModuleRegistry(config);
  const database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  const seen = watched(database);
  const scheduler = new ModuleScheduler({
    enabledPacks: config.enabledFeatures,
    registry,
    database: seen,
    graphClient: (getAccessToken) => new GraphHttpClient({ getAccessToken }),
  });
  const app = createApp({
    config,
    store,
    identity,
    graphFetch,
    ...(options.modules === false ? {} : { modules: { registry, database: seen, scheduler } }),
  });
  // The API allows 60 requests a minute per session, so tests sign in afresh.
  let admin = { cookie: "", csrf: "" };
  const signIn = async () => {
    const login = await identity.startLogin(features[0] as "directory.users.read");
    const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
    admin = {
      cookie: `prism365_session=${loggedIn.handle}`,
      csrf: (await identity.session(loggedIn.handle))?.csrfToken ?? "",
    };
  };
  await signIn();
  const get = (url: string, signedIn = true) =>
    app.inject({ method: "GET", url, headers: signedIn ? { cookie: admin.cookie } : {} });
  const upload = (payload: unknown, headers: Record<string, string> = {}) =>
    app.inject({
      method: "POST",
      url: "/v1/mailbox-access/import",
      headers: {
        cookie: admin.cookie,
        origin,
        "x-csrf-token": admin.csrf,
        "content-type": "application/json",
        ...headers,
      },
      payload: typeof payload === "string" ? payload : JSON.stringify(payload),
    });
  const clear = (headers: Record<string, string> = {}) =>
    app.inject({
      method: "DELETE",
      url: "/v1/mailbox-access",
      headers: { cookie: admin.cookie, origin, "x-csrf-token": admin.csrf, ...headers },
    });
  const count = async (table: string, tenant = tenantId) =>
    (
      await database.tenant(tenant, (sql) =>
        sql.query<{ n: number }>(`SELECT count(*)::int AS n FROM ${table}`),
      )
    ).rows[0]?.n;
  const ledger = (type: string) => store.audit.filter((entry) => entry.event.type === type);
  return {
    app,
    store,
    database,
    signIn,
    get,
    upload,
    clear,
    count,
    ledger,
    requests,
    ledgerAtFirstRequest: () => ledgerAtFirstRequest,
    resetFirstRequest: () => {
      ledgerAtFirstRequest = null;
    },
  };
}
type Harness = Awaited<ReturnType<typeof harness>>;

describe("mailbox access import", () => {
  let h: Harness;
  beforeAll(async () => {
    h = await harness();
    currentStore = h.store;
  }, 60_000);
  afterAll(async () => {
    await h?.app.close();
  });
  beforeEach(async () => {
    await h.signIn();
    failGrantBatchFrom = null;
    ledgerAtFirstWrite = null;
    await h.clear();
  });

  it("imports an export and answers its dates, importer and counts", async () => {
    const file = exportFile();
    const response = await h.upload(file);
    expect(response.statusCode, response.body).toBe(201);
    const answer = response.json();
    expect(answer.replaced).toBe(false);
    expect(answer.import).toMatchObject({
      exportedAt: new Date(file.exportedAt).toISOString(),
      importedBy: { id: actorId, name: "Synthetic admin" },
      includeUserMailboxes: false,
      mailboxes: 2,
      grants: 6,
      unresolved: 1,
    });
    const status = await h.get("/v1/mailbox-access");
    expect(status.statusCode).toBe(200);
    expect(status.json().import).toEqual(answer.import);
    expect(await h.count("prism365_mailbox_access_grant")).toBe(6);
    // The ledger holds ids and counts, never a name or an address.
    const entry = h.ledger("mailbox_access.imported").at(-1)?.event;
    expect(entry).toMatchObject({
      importId: answer.import.id,
      mailboxes: 2,
      grants: 6,
      unresolved: 1,
    });
    expect(JSON.stringify(entry)).not.toMatch(/@|Pat|Lab|S-1-5/);
  });

  it("keeps ids in lower case and cleans the text it keeps", async () => {
    const file = exportFile({
      mailboxes: [
        {
          id: LAB.toUpperCase(),
          address: "lab.results@example.invalid",
          displayName: "Lab\u0000 results\n\u202Eetluser",
          type: "SharedMailbox",
        },
      ],
      grants: [
        {
          mailboxId: LAB.toUpperCase(),
          right: "SendAs",
          trustee: user(PERSON.toUpperCase(), "Pat\tPell", "pat.pell@example.invalid"),
        },
      ],
      unresolved: 0,
    });
    expect((await h.upload(file)).statusCode).toBe(201);
    const rows = await h.database.tenant(tenantId, (sql) =>
      sql.query<{
        mailbox_id: string;
        mailbox_name: string;
        trustee_id: string;
        trustee_name: string;
      }>(
        "SELECT mailbox_id, mailbox_name, trustee_id, trustee_name FROM prism365_mailbox_access_grant",
      ),
    );
    expect(rows.rows).toEqual([
      {
        mailbox_id: LAB,
        mailbox_name: "Lab results etluser",
        trustee_id: PERSON,
        trustee_name: "Pat Pell",
      },
    ]);
  });

  // Every refusal stores nothing, records nothing and says what is wrong in plain words.
  const refusals: [string, () => unknown, RegExp][] = [
    ["not a JSON object", () => [1, 2], /not a JSON object/],
    [
      "another format",
      () => exportFile({ format: "prism365.receipt" }),
      /not a Prism365 mailbox access export/,
    ],
    [
      "another version",
      () => exportFile({ version: 2 }),
      /version 2\. This Prism365 reads version 1/,
    ],
    [
      "another tenant",
      () => exportFile({ tenantId: otherTenant }),
      /from tenant 99999999-.*reads tenant 11111111-/,
    ],
    [
      "an export older than 90 days",
      () => exportFile({ exportedAt: ago(91) }),
      /more than 90 days ago/,
    ],
    [
      "an export dated in the future",
      () => exportFile({ exportedAt: new Date(Date.now() + 3_600_000).toISOString() }),
      /in the future/,
    ],
    [
      "a date without a time zone",
      () => exportFile({ exportedAt: "2026-11-01T10:00:00" }),
      /exportedAt is not a date and time with a time zone/,
    ],
    [
      "more than 20,000 mailboxes",
      () =>
        exportFile({
          mailboxes: Array.from({ length: 20_001 }, (_, index) => ({
            id: `eeeeeeee-0000-4000-8000-${String(index).padStart(12, "0")}`,
            address: "m@example.invalid",
            displayName: "M",
            type: "SharedMailbox",
          })),
          grants: [],
          unresolved: 0,
        }),
      /mailboxes has more than 20,000 entries/,
    ],
    [
      "a grant for a mailbox the file does not list",
      () => {
        const file = exportFile();
        file.grants[4] = {
          ...file.grants[4],
          mailboxId: OTHER_GROUP,
        } as (typeof file.grants)[number];
        return file;
      },
      /grants\[4\]\.mailboxId is not a mailbox in this file/,
    ],
    [
      "a display name over 256 characters",
      () => {
        const file = exportFile();
        file.mailboxes[1] = {
          ...file.mailboxes[1],
          displayName: "x".repeat(257),
        } as (typeof file.mailboxes)[number];
        return file;
      },
      /mailboxes\[1\]\.displayName is longer than 256 characters/,
    ],
    [
      "a trustee name over 512 characters",
      () => {
        const file = exportFile();
        file.grants[2] = {
          ...file.grants[2],
          trustee: { id: null, address: null, name: "S".repeat(513), kind: "unresolved" },
        } as (typeof file.grants)[number];
        return file;
      },
      /grants\[2\]\.trustee\.name is longer than 512 characters/,
    ],
    [
      "a mailbox id that is not a GUID",
      () => {
        const file = exportFile();
        file.mailboxes[0] = {
          ...file.mailboxes[0],
          id: "lab-results",
        } as (typeof file.mailboxes)[number];
        return file;
      },
      /mailboxes\[0\]\.id is not a GUID/,
    ],
    [
      "a trustee id that is not a GUID",
      () => {
        const file = exportFile();
        file.grants[1] = {
          ...file.grants[1],
          trustee: user("pat.pell@example.invalid", "Pat Pell", "pat.pell@example.invalid"),
        } as (typeof file.grants)[number];
        return file;
      },
      /grants\[1\]\.trustee\.id is not a GUID/,
    ],
    [
      "an unknown field",
      () => exportFile({ comment: "edited" }),
      /a field Prism365 does not read: "comment"/,
    ],
    [
      "an unresolved count that disagrees",
      () => exportFile({ unresolved: 0 }),
      /unresolved says 0, but 1 grant has an unresolved trustee/,
    ],
    [
      "a repeated grant",
      () => {
        const file = exportFile();
        file.grants.push(file.grants[1] as (typeof file.grants)[number]);
        return file;
      },
      /grants\[6\] repeats grants\[1\]/,
    ],
    [
      "an unresolved trustee with an id",
      () => {
        const file = exportFile();
        file.grants[2] = {
          ...file.grants[2],
          trustee: { id: OTHER_PERSON, address: null, name: SID, kind: "unresolved" },
        } as (typeof file.grants)[number];
        return file;
      },
      /grants\[2\]\.trustee\.id has an id, but the trustee is unresolved/,
    ],
    [
      "a user mailbox in a shared-only export",
      () => {
        const file = exportFile();
        file.mailboxes[1] = {
          ...file.mailboxes[1],
          type: "UserMailbox",
        } as (typeof file.mailboxes)[number];
        return file;
      },
      /mailboxes\[1\]\.type is UserMailbox, but includeUserMailboxes is false/,
    ],
    [
      "a right Prism365 does not read",
      () => {
        const file = exportFile();
        file.grants[0] = {
          ...file.grants[0],
          right: "ReadPermission",
        } as (typeof file.grants)[number];
        return file;
      },
      /grants\[0\]\.right should be FullAccess, SendAs or SendOnBehalf/,
    ],
  ];
  for (const [name, build, message] of refusals)
    it(`refuses ${name}, keeping the current import and recording nothing`, async () => {
      expect((await h.upload(exportFile())).statusCode).toBe(201);
      const before = (await h.get("/v1/mailbox-access")).json().import;
      const ledgered = h.ledger("mailbox_access.imported").length;
      const response = await h.upload(build());
      expect(response.statusCode, response.body).toBe(400);
      expect(response.json().message).toMatch(message);
      expect(h.ledger("mailbox_access.imported")).toHaveLength(ledgered);
      expect((await h.get("/v1/mailbox-access")).json().import).toEqual(before);
      expect(await h.count("prism365_mailbox_access_grant")).toBe(6);
    });

  it("refuses a body that is not JSON, and one over 25 MB, before reading it as an export", async () => {
    const ledgered = h.ledger("mailbox_access.imported").length;
    const broken = await h.upload("{ not json");
    expect(broken.statusCode).toBe(400);
    expect(broken.json()).toMatchObject({
      error: "invalid_export",
      message: "The export is not valid JSON.",
    });
    const large = await h.upload(`{"pad":"${"x".repeat(25 * 1024 * 1024)}"}`);
    expect(large.statusCode).toBe(413);
    expect(large.json().message).toMatch(/at most 25 MB/);
    expect(h.ledger("mailbox_access.imported")).toHaveLength(ledgered);
  });

  it("reads no large body for a request without a session", async () => {
    const response = await h.app.inject({
      method: "POST",
      url: "/v1/mailbox-access/import",
      headers: { origin, "content-type": "application/json" },
      payload: `{"pad":"${"x".repeat(2_000_000)}"}`,
    });
    expect(response.statusCode).toBe(401);
  });

  it("refuses an import without the CSRF token", async () => {
    const response = await h.upload(exportFile(), { "x-csrf-token": "wrong" });
    expect(response.statusCode).toBe(403);
    expect(await h.count("prism365_mailbox_access_import")).toBe(0);
  });

  it("replaces the previous import whole", async () => {
    const first = (await h.upload(exportFile())).json().import;
    const smaller = exportFile({
      mailboxes: [exportFile().mailboxes[1]],
      grants: [exportFile().grants[3]],
      unresolved: 0,
    });
    const response = await h.upload(smaller);
    expect(response.statusCode).toBe(201);
    expect(response.json().replaced).toBe(true);
    expect(response.json().import.id).not.toBe(first.id);
    expect(await h.count("prism365_mailbox_access_import")).toBe(1);
    expect(await h.count("prism365_mailbox_access_grant")).toBe(1);
  });

  it("keeps the old import intact when a grant fails to store midway", async () => {
    const first = (await h.upload(exportFile())).json().import;
    // 450 grants are three INSERT batches; the second one fails.
    const many = exportFile({
      grants: Array.from({ length: 450 }, (_, index) => ({
        mailboxId: LAB,
        right: "FullAccess",
        trustee: user(
          `cccccccc-0000-4000-8000-${String(index).padStart(12, "0")}`,
          `Person ${index}`,
          `p${index}@example.invalid`,
        ),
      })),
      unresolved: 0,
    });
    failGrantBatchFrom = 200;
    const response = await h.upload(many);
    failGrantBatchFrom = null;
    expect(response.statusCode).toBe(500);
    expect(response.json().message).toMatch(/previous import, if any, is unchanged/);
    expect((await h.get("/v1/mailbox-access")).json().import).toEqual(first);
    expect(await h.count("prism365_mailbox_access_grant")).toBe(6);
    // Without the failure the same file stores every grant.
    expect((await h.upload(many)).statusCode).toBe(201);
    expect(await h.count("prism365_mailbox_access_grant")).toBe(450);
  });

  it("records the import before writing, and changes nothing when the ledger fails", async () => {
    const first = (await h.upload(exportFile())).json().import;
    ledgerAtFirstWrite = null;
    const ledgered = h.store.audit.length;
    const second = await h.upload(exportFile({ exportedAt: ago(1) }));
    expect(second.statusCode).toBe(201);
    // The entry was in the ledger before the first delete or insert reached storage.
    expect(ledgerAtFirstWrite).toBe(ledgered + 1);
    const replacement = second.json().import;

    const append = h.store.appendAudit.bind(h.store);
    h.store.appendAudit = async () => {
      throw new Error("Synthetic ledger failure.");
    };
    ledgerAtFirstWrite = null;
    try {
      expect((await h.upload(exportFile())).statusCode).toBe(500);
    } finally {
      h.store.appendAudit = append;
    }
    expect(ledgerAtFirstWrite).toBeNull();
    expect((await h.get("/v1/mailbox-access")).json().import).toEqual(replacement);
    expect(replacement.id).not.toBe(first.id);
  });

  it("clears the import, recording it first, and says when there was nothing", async () => {
    expect((await h.upload(exportFile())).statusCode).toBe(201);
    ledgerAtFirstWrite = null;
    const ledgered = h.store.audit.length;
    const cleared = await h.clear();
    expect(cleared.statusCode).toBe(200);
    expect(cleared.json()).toEqual({ cleared: true });
    expect(ledgerAtFirstWrite).toBe(ledgered + 1);
    expect(h.ledger("mailbox_access.cleared").length).toBeGreaterThan(0);
    expect((await h.get("/v1/mailbox-access")).json()).toEqual({ import: null });
    expect(await h.count("prism365_mailbox_access_grant")).toBe(0);
    expect((await h.clear()).json()).toEqual({ cleared: false });
  });

  it("deletes nothing when the ledger cannot record the clear", async () => {
    expect((await h.upload(exportFile())).statusCode).toBe(201);
    const append = h.store.appendAudit.bind(h.store);
    h.store.appendAudit = async () => {
      throw new Error("Synthetic ledger failure.");
    };
    try {
      expect((await h.clear()).statusCode).toBe(500);
    } finally {
      h.store.appendAudit = append;
    }
    expect(await h.count("prism365_mailbox_access_grant")).toBe(6);
  });

  it("refuses a clear without the CSRF token", async () => {
    expect((await h.upload(exportFile())).statusCode).toBe(201);
    expect((await h.clear({ "x-csrf-token": "wrong" })).statusCode).toBe(403);
    expect(await h.count("prism365_mailbox_access_import")).toBe(1);
  });

  it("keeps each tenant's import to that tenant", async () => {
    expect((await h.upload(exportFile())).statusCode).toBe(201);
    const other = await h.database.tenant(otherTenant, (sql) => currentImport(sql, otherTenant));
    expect(other).toBeNull();
    expect(await h.count("prism365_mailbox_access_grant", otherTenant)).toBe(0);
  });
});

describe("a person's mailbox access", () => {
  let h: Harness;
  beforeAll(async () => {
    h = await harness();
    currentStore = h.store;
  }, 60_000);
  afterAll(async () => {
    await h?.app.close();
  });
  beforeEach(async () => {
    await h.signIn();
    await h.clear();
    h.resetFirstRequest();
    h.requests.length = 0;
  });

  it("shows one export's date with that export's grants when it is replaced mid-read", async () => {
    // An older export, then a newer one imported while the route waits for the groups.
    expect((await h.upload(exportFile({ exportedAt: ago(20) }))).statusCode).toBe(201);
    const newer = exportFile({
      exportedAt: ago(1),
      mailboxes: [exportFile().mailboxes[1]],
      grants: [exportFile().grants[3]],
      unresolved: 0,
    });
    duringGroupRead = async () => expect((await h.upload(newer)).statusCode).toBe(201);
    const answer = (await h.get(`/v1/people/${PERSON}/mailbox-access`)).json();
    expect(duringGroupRead).toBeNull();
    expect(answer.import.exportedAt).toBe(new Date(newer.exportedAt).toISOString());
    expect(answer.import.grants).toBe(1);
    expect(
      answer.grants.map((grant: { mailbox: { id: string }; right: string }) => [
        grant.mailbox.id,
        grant.right,
      ]),
    ).toEqual([[BILLING, "SendOnBehalf"]]);
  });

  it("says nothing is imported when the import is cleared mid-read", async () => {
    expect((await h.upload(exportFile())).statusCode).toBe(201);
    duringGroupRead = async () => expect((await h.clear()).json()).toEqual({ cleared: true });
    const answer = (await h.get(`/v1/people/${PERSON}/mailbox-access`)).json();
    expect(duringGroupRead).toBeNull();
    expect(answer.import).toBeNull();
    expect(answer.grants).toEqual([]);
  });

  it("asks Microsoft nothing and says so when nothing was imported", async () => {
    const response = await h.get(`/v1/people/${PERSON}/mailbox-access`);
    expect(response.statusCode).toBe(200);
    expect(response.json()).toEqual({
      import: null,
      groups: { status: "not_read" },
      grants: [],
      grantsTruncated: false,
      holders: [],
      holdersTruncated: false,
    });
    expect(h.requests).toEqual([]);
    expect(h.ledger("person.mailbox_access.viewed").length).toBeGreaterThan(0);
  });

  it("lists rights held directly and through a nested group, recording the view first", async () => {
    expect((await h.upload(exportFile())).statusCode).toBe(201);
    const ledgered = h.store.audit.length;
    const response = await h.get(`/v1/people/${PERSON.toUpperCase()}/mailbox-access`);
    expect(response.statusCode).toBe(200);
    const answer = response.json();
    expect(h.ledgerAtFirstRequest()).toBe(ledgered + 1);
    expect(h.ledger("person.mailbox_access.viewed").at(-1)?.event).toMatchObject({
      datasets: ["mailboxAccess", "directory"],
    });
    expect(h.requests).toEqual([`/v1.0/users/${PERSON}/transitiveMemberOf?$select=id&$top=100`]);
    expect(answer.groups).toMatchObject({ status: "ok", count: 2, truncated: false });
    expect(
      answer.grants.map((grant: { mailbox: { id: string }; right: string; via: unknown }) => [
        grant.mailbox.id,
        grant.right,
        grant.via,
      ]),
    ).toEqual([
      [BILLING, "SendOnBehalf", { kind: "direct" }],
      [LAB, "FullAccess", { kind: "group", id: NESTED_GROUP, name: "All plant staff" }],
      [LAB, "SendAs", { kind: "direct" }],
    ]);
    // A person is not a mailbox in this export, so nobody holds anything on theirs.
    expect(answer.holders).toEqual([]);
  });

  it("never matches a group grant for a group the person is not in, or to someone else", async () => {
    expect((await h.upload(exportFile())).statusCode).toBe(201);
    const answer = (await h.get(`/v1/people/${PERSON}/mailbox-access`)).json();
    const trustees = JSON.stringify(answer.grants);
    expect(trustees).not.toContain(OTHER_GROUP);
    expect(trustees).not.toContain("Finance");
    const other = (await h.get(`/v1/people/${OTHER_PERSON}/mailbox-access`)).json();
    expect(other.grants).toEqual([
      {
        mailbox: {
          id: BILLING,
          address: "billing@example.invalid",
          name: "Billing",
          type: "SharedMailbox",
        },
        right: "FullAccess",
        via: { kind: "direct" },
      },
    ]);
  });

  it("matches a group grant only when the export calls the trustee a group", async () => {
    const file = exportFile();
    file.grants[0] = {
      ...file.grants[0],
      trustee: { id: NESTED_GROUP, address: null, name: "All plant staff", kind: "other" },
    } as (typeof file.grants)[number];
    expect((await h.upload(file)).statusCode).toBe(201);
    const answer = (await h.get(`/v1/people/${PERSON}/mailbox-access`)).json();
    expect(
      answer.grants.some((grant: { via: { kind: string } }) => grant.via.kind === "group"),
    ).toBe(false);
  });

  it("lists who holds which right on a mailbox when the person is that mailbox", async () => {
    expect((await h.upload(exportFile())).statusCode).toBe(201);
    const answer = (await h.get(`/v1/people/${LAB}/mailbox-access`)).json();
    expect(answer.holders).toEqual([
      {
        right: "FullAccess",
        trustee: {
          id: NESTED_GROUP,
          address: "plant.staff@example.invalid",
          name: "All plant staff",
          kind: "group",
        },
      },
      { right: "FullAccess", trustee: { id: null, address: null, name: SID, kind: "unresolved" } },
      {
        right: "SendAs",
        trustee: {
          id: PERSON,
          address: "pat.pell@example.invalid",
          name: "Pat Pell",
          kind: "user",
        },
      },
    ]);
  });

  it("still answers direct grants when Microsoft refuses the group read", async () => {
    const refusing = await harness({
      handler: () =>
        Response.json({ error: { code: "Authorization_RequestDenied" } }, { status: 403 }),
    });
    currentStore = refusing.store;
    try {
      expect((await refusing.upload(exportFile())).statusCode).toBe(201);
      const answer = (await refusing.get(`/v1/people/${PERSON}/mailbox-access`)).json();
      expect(answer.groups.status).toBe("failed");
      expect(answer.groups.message).toMatch(/Microsoft refused this read/);
      expect(answer.grants.map((grant: { right: string }) => grant.right)).toEqual([
        "SendOnBehalf",
        "SendAs",
      ]);
    } finally {
      await refusing.app.close();
      currentStore = h.store;
    }
  });

  it("refuses an unknown account, and answers nothing without a session", async () => {
    expect((await h.get("/v1/people/pat.pell@example.invalid/mailbox-access")).statusCode).toBe(
      400,
    );
    expect((await h.get(`/v1/people/${PERSON}/mailbox-access`, false)).statusCode).toBe(401);
    expect((await h.get("/v1/mailbox-access", false)).statusCode).toBe(401);
  });
});

describe("a person's mailbox access without the users pack", () => {
  it("finds direct grants only, says why, and never asks Microsoft", async () => {
    const h = await harness({
      features: ["audit.signins.read"],
      scopes: ["User.Read", "AuditLog.Read.All"],
    });
    currentStore = h.store;
    try {
      expect((await h.upload(exportFile())).statusCode).toBe(201);
      const answer = (await h.get(`/v1/people/${PERSON}/mailbox-access`)).json();
      expect(answer.groups).toEqual({
        status: "not_granted",
        pack: "directory.users.read",
        permission: "User.Read.All",
        enabled: false,
      });
      expect(answer.grants.map((grant: { via: { kind: string } }) => grant.via.kind)).toEqual([
        "direct",
        "direct",
      ]);
      expect(h.requests).toEqual([]);
      expect(h.ledger("person.mailbox_access.viewed").at(-1)?.event).toMatchObject({
        datasets: ["mailboxAccess"],
      });
    } finally {
      await h.app.close();
    }
  });

  it("says the permission is missing when the pack is on but not granted", async () => {
    const h = await harness({ scopes: ["User.Read", "AuditLog.Read.All"] });
    currentStore = h.store;
    try {
      expect((await h.upload(exportFile())).statusCode).toBe(201);
      const answer = (await h.get(`/v1/people/${PERSON}/mailbox-access`)).json();
      expect(answer.groups).toMatchObject({ status: "not_granted", enabled: true });
      expect(h.requests).toEqual([]);
    } finally {
      await h.app.close();
    }
  });
});

describe("mailbox access without module storage", () => {
  it("answers that it is unavailable", async () => {
    const h = await harness({ modules: false });
    try {
      expect((await h.get("/v1/mailbox-access")).statusCode).toBe(409);
      expect((await h.upload(exportFile())).statusCode).toBe(409);
      expect((await h.get(`/v1/people/${PERSON}/mailbox-access`)).statusCode).toBe(409);
    } finally {
      await h.app.close();
    }
  });
});

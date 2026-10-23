import { randomUUID } from "node:crypto";
import {
  auditDefinition,
  builtInAudits,
  findAudit,
  finishRun,
  inactiveAccountsAudit,
  insertRun,
  licenseCapacityAudit,
} from "@prism365/audits";
import { type AuditRow, canonicalJson, handoffSchema, receiptSchema } from "@prism365/contracts";
import { GraphHttpClient } from "@prism365/graph";
import { ModuleScheduler } from "@prism365/modules";
import {
  MemoryStateStore,
  type ModuleDatabase,
  PGliteModuleDatabase,
  type SqlClient,
} from "@prism365/storage";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import { createApp } from "../src/app.ts";
import { AuditRunner } from "../src/audits.ts";
import { type AppConfig, type IdentityConfig, loadConfig } from "../src/config.ts";
import { insertReceipt } from "../src/handoff-store.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import { createModuleRegistry } from "../src/modules.ts";

// Synthetic tenant, administrators, accounts and tools only.
const tenantId = "11111111-1111-4111-8111-111111111111";
const otherTenant = "99999999-9999-4999-8999-999999999999";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
const features = ["directory.users.read", "audit.signins.read"] as const;
const identityConfig: IdentityConfig = {
  tenantId,
  clientId: "33333333-3333-4333-8333-333333333333",
  origin,
  redirectUri: `${origin}/api/v1/auth/callback`,
  allowedActorIds: [actorId],
  enabledFeatures: [...features],
  credential: { clientSecret: "synthetic-only" },
};
const config: AppConfig = {
  mode: "live",
  origin,
  identity: identityConfig,
  production: false,
  missingConfiguration: [],
  enabledFeatures: [...features],
};
const provider: MicrosoftIdentityProvider = {
  authorizationUrl: async () => "https://login.microsoftonline.com/synthetic",
  redeem: async () => ({
    tenantId,
    actorId,
    displayName: "Synthetic admin",
    accountId: "synthetic-account",
    cache: "synthetic-cache",
    scopes: ["User.Read", "User.Read.All", "AuditLog.Read.All"],
  }),
  token: async (session) => ({ accessToken: "synthetic", cache: "c", scopes: session.scopes }),
};

// One account per inactivity range, with the cells the inactive accounts audit shows.
const account = (n: number, days: number, status: string, tags: string[]): AuditRow => ({
  key: `user-${n}`,
  personId: `0000000${n}-aaaa-4aaa-8aaa-00000000000${n}`,
  tags: [...tags, status.toLowerCase()],
  cells: {
    person: { name: `Synthetic Person ${n}`, principal: `person${n}@example.invalid` },
    days,
    basis: "Last successful sign-in",
    lastSuccess: new Date(Date.now() - days * 86_400_000).toISOString(),
    status,
    type: "Member",
    domain: "example.invalid",
    licenses: n % 2 ? ["Example E3"] : [],
    licensed: n % 2 ? "Licensed" : "Unlicensed",
    created: "2024-01-01T00:00:00.000Z",
    source: "Cloud",
    department: n % 2 ? "Finance" : "Sales",
    jobTitle: "Synthetic",
  },
});
const rows: AuditRow[] = [
  account(1, 3, "Enabled", []),
  account(2, 40, "Enabled", ["d7", "d30"]),
  account(3, 95, "Disabled", ["d7", "d30", "d60", "d90"]),
  account(4, 200, "Enabled", ["d7", "d30", "d60", "d90", "d120", "d180"]),
  account(5, 400, "Disabled", ["d7", "d30", "d60", "d90", "d120", "d180", "d365", "never"]),
  // A row with no account behind it is never handed off.
  { ...account(6, 90, "Enabled", ["d7", "d30", "d60", "d90"]), personId: null },
];

let failItemsFrom: number | null = null;
// Makes the next look-up of a stored receipt answer "not found", as when another request stores
// the same id between this request's check and its insert.
let hideStoredReceiptOnce = false;
// Reads of the handoff and receipt tables, to show the ledger entry is written before them.
let handoffReads = 0;
// The module database the app sees: the real one, with switches that make storing receipt items
// fail from a given position, to show a failed item keeps nothing of its receipt, and that make
// the import lose a race.
function failingItems(database: ModuleDatabase): ModuleDatabase {
  return {
    kind: database.kind,
    durable: database.durable,
    tenant: (tenant, operation, options) =>
      database.tenant(
        tenant,
        (sql) =>
          operation({
            query: <T extends object>(text: string, params?: readonly unknown[]) => {
              if (/^\s*SELECT[\s\S]*FROM prism365_(handoff|receipt)/.test(text)) handoffReads += 1;
              if (
                failItemsFrom !== null &&
                text.includes("INSERT INTO prism365_receipt_item") &&
                Number(params?.[2]) >= failItemsFrom
              )
                return Promise.reject(new Error("Synthetic storage failure."));
              if (hideStoredReceiptOnce && /WHERE r\.tenant_id=\$1 AND r\.id=\$2/.test(text)) {
                hideStoredReceiptOnce = false;
                return Promise.resolve({ rows: [] as T[], rowCount: 0 });
              }
              return sql.query<T>(text, params);
            },
          } satisfies SqlClient),
        options,
      ),
    verify: () => database.verify(),
    close: () => database.close(),
  };
}

let app: ReturnType<typeof createApp>;
let store: MemoryStateStore;
let database: PGliteModuleDatabase;
let identity: IdentityService;
let admin: { cookie: string; csrf: string };
let inactiveRun: string;
let capacityRun: string;
let failedRun: string;
let scopedRun: string;
let devicesRun: string;
let usageRun: string;

// The API allows 60 requests a minute per session, so each test signs in afresh.
async function signIn() {
  const login = await identity.startLogin("directory.users.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
  return {
    cookie: `prism365_session=${loggedIn.handle}`,
    csrf: (await identity.session(loggedIn.handle))?.csrfToken ?? "",
  };
}

beforeEach(async () => {
  admin = await signIn();
});

beforeAll(async () => {
  store = new MemoryStateStore();
  identity = new IdentityService(identityConfig, store, provider);
  const registry = createModuleRegistry(config);
  database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  const seen = failingItems(database);
  const scheduler = new ModuleScheduler({
    enabledPacks: config.enabledFeatures,
    registry,
    database: seen,
    graphClient: (getAccessToken) => new GraphHttpClient({ getAccessToken }),
  });
  app = createApp({
    config,
    store,
    identity,
    audits: new AuditRunner({ database: seen }),
    modules: { registry, database: seen, scheduler },
  });
  const actor = { id: actorId, name: "Synthetic admin" };
  const stored = async (
    auditId: string,
    version: number,
    outcome: Parameters<typeof finishRun>[4],
    scope: string[] | null = null,
  ) => {
    const id = randomUUID();
    await database.tenant(tenantId, async (sql) => {
      await insertRun(sql, tenantId, {
        id,
        auditId,
        auditVersion: version,
        startedAt: new Date("2026-10-21T09:00:00Z"),
        startedBy: actor,
        scope,
      });
      await finishRun(sql, tenantId, id, new Date("2026-10-21T09:01:00Z"), outcome);
    });
    return id;
  };
  const inactive = inactiveAccountsAudit.definition;
  inactiveRun = await stored(inactive.id, inactive.version, {
    status: "succeeded",
    complete: true,
    summary: [],
    notes: [],
    rows,
  });
  capacityRun = await stored(
    licenseCapacityAudit.definition.id,
    licenseCapacityAudit.definition.version,
    { status: "succeeded", complete: true, summary: [], notes: [], rows: [] },
  );
  failedRun = await stored(inactive.id, inactive.version, {
    status: "failed",
    code: "x",
    message: "Synthetic failure.",
  });
  // A run limited to two domains, which Microsoft answered only in part.
  scopedRun = await stored(
    inactive.id,
    inactive.version,
    { status: "succeeded", complete: false, summary: [], notes: [], rows },
    ["example.invalid", "sales.example.invalid"],
  );
  // Devices with an owner: a person column, but the rows are devices, not accounts to act on.
  const devices = findAudit("devices.managed");
  const usage = findAudit("usage.m365-apps");
  if (!devices || !usage) throw new Error("The device and usage audits are missing.");
  devicesRun = await stored(devices.definition.id, devices.definition.version, {
    status: "succeeded",
    complete: true,
    summary: [],
    notes: [],
    rows: [{ ...account(1, 3, "Enabled", ["not-compliant"]), key: "device-1" }],
  });
  // The usage report names accounts only by sign-in name, so its rows carry no directory id.
  usageRun = await stored(usage.definition.id, usage.definition.version, {
    status: "succeeded",
    complete: true,
    summary: [],
    notes: [],
    rows: [{ ...account(2, 40, "Enabled", ["idle"]), personId: null }],
  });
}, 60_000);

afterAll(async () => {
  await app?.close();
});

const get = (url: string) => app.inject({ method: "GET", url, headers: { cookie: admin.cookie } });
const importReceipt = (payload: unknown, headers: Record<string, string> = {}) =>
  app.inject({
    method: "POST",
    url: "/v1/receipts",
    headers: {
      cookie: admin.cookie,
      origin,
      "x-csrf-token": admin.csrf,
      "content-type": "application/json",
      ...headers,
    },
    payload: typeof payload === "string" ? payload : JSON.stringify(payload),
  });
const count = async (table: string, tenant = tenantId) =>
  (
    await database.tenant(tenant, (sql) =>
      sql.query<{ n: number }>(`SELECT count(*)::int AS n FROM ${table}`),
    )
  ).rows[0]?.n;
const ledger = (type: string) => store.audit.filter((entry) => entry.event.type === type);

const receipt = (id: string, items: Record<string, unknown>[], extra = {}) => ({
  format: "prism365.receipt",
  version: 1,
  id,
  handoffId: null,
  tool: { name: "Example action tool", version: "1.0.0" },
  performedAt: "2026-10-22T14:30:00Z",
  performedBy: "Operator One",
  items,
  ...extra,
});
const item = (overrides: Record<string, unknown> = {}) => ({
  userId: "00000003-aaaa-4aaa-8aaa-000000000003",
  userPrincipalName: "person3@example.invalid",
  action: "account.sign-in-blocked",
  status: "done",
  detail: "",
  ...overrides,
});

// The principal names in a CSV export, one per data row: the column after the account's name.
function csvPrincipals(csv: string): string[] {
  return csv
    .replace(/^﻿/, "")
    .trim()
    .split("\r\n")
    .slice(1)
    .map((line) => /^"[^"]*","([^"]*)"/.exec(line)?.[1] ?? "");
}

describe("handoff routes", () => {
  it("require a session, and the import its CSRF token", async () => {
    for (const url of [
      `/v1/audit-runs/${inactiveRun}/handoff.json`,
      "/v1/handoffs",
      `/v1/receipts/${randomUUID()}`,
      "/v1/people/user-1/receipts",
    ])
      expect((await app.inject({ method: "GET", url })).statusCode, url).toBe(401);
    const anonymous = await app.inject({
      method: "POST",
      url: "/v1/receipts",
      headers: { "content-type": "application/json", origin },
      payload: JSON.stringify(receipt(randomUUID(), [item()])),
    });
    expect(anonymous.statusCode).toBe(401);
    // 2 MB that is not JSON: parsing it would answer 400, so a 401 shows it was never read.
    const unread = await app.inject({
      method: "POST",
      url: "/v1/receipts",
      headers: { "content-type": "application/json", origin },
      payload: `{${"x".repeat(2 * 1024 * 1024)}`,
    });
    expect(unread.statusCode).toBe(401);
    const forged = await importReceipt(receipt(randomUUID(), [item()]), { "x-csrf-token": "no" });
    expect(forged.statusCode).toBe(403);
    expect(ledger("receipt.imported")).toHaveLength(0);
  });
});

describe("handoff export", () => {
  it("selects the same accounts as the CSV export for the same parameters", async () => {
    for (const query of [
      "",
      "view=all",
      "view=d90&sort=days&direction=asc",
      "view=all&f.status=Disabled",
      "view=all&search=PERSON4",
      "view=all&sort=no-such-column",
    ]) {
      const csv = await get(`/v1/audit-runs/${inactiveRun}/export.csv?${query}`);
      const json = await get(`/v1/audit-runs/${inactiveRun}/handoff.json?${query}`);
      expect(csv.statusCode, query).toBe(200);
      expect(json.statusCode, query).toBe(200);
      const document = handoffSchema.parse(json.json());
      const principals = csvPrincipals(csv.body).filter(
        (name) => name !== "person6@example.invalid",
      );
      expect(
        document.items.map((entry) => entry.userPrincipalName),
        query,
      ).toEqual(principals);
      const byPrincipal = new Map(rows.map((row) => [row.cells.person, row.personId]));
      expect(
        document.items.map((entry) => entry.userId),
        query,
      ).toEqual(
        principals.map(
          (principal) =>
            [...byPrincipal].find(
              ([cell]) => (cell as { principal: string }).principal === principal,
            )?.[1],
        ),
      );
    }
  });

  it("records the export first, then the handoff, and sends it as a named file", async () => {
    const before = ledger("handoff.exported").length;
    const response = await get(
      `/v1/audit-runs/${inactiveRun}/handoff.json?view=d90&columns=person,days,status,jobTitle`,
    );
    expect(response.statusCode).toBe(200);
    expect(response.headers["content-disposition"]).toBe(
      'attachment; filename="prism365-handoff-accounts.inactive-2026-10-21.json"',
    );
    expect(response.headers["content-type"]).toBe("application/json; charset=utf-8");
    const document = handoffSchema.parse(response.json());
    expect(document).toMatchObject({
      format: "prism365.handoff",
      version: 1,
      tenantId,
      generatedBy: { actorId, name: "Synthetic admin" },
      source: { auditId: "accounts.inactive", runId: inactiveRun, view: { id: "d90" } },
    });
    expect(document.items.map((entry) => entry.userPrincipalName)).toEqual([
      "person5@example.invalid",
      "person4@example.invalid",
      "person3@example.invalid",
    ]);
    expect(document.items[2]).toMatchObject({
      reasons: ["7+ days", "30+ days", "60+ days", "90+ days"],
      evidence: { "Inactive for": 95, Status: "Disabled", "Job title": "Synthetic" },
    });
    const entries = ledger("handoff.exported");
    expect(entries).toHaveLength(before + 1);
    expect(entries.at(-1)?.event).toEqual({
      type: "handoff.exported",
      auditId: "accounts.inactive",
      runId: inactiveRun,
      handoffId: document.id,
      items: 3,
      occurredAt: document.generatedAt,
    });
    expect(document.source).toMatchObject({ scope: null, complete: true });
    const list = await get("/v1/handoffs");
    expect(list.json().handoffs[0]).toMatchObject({
      id: document.id,
      auditTitle: "Inactive accounts",
      view: { id: "d90", label: "90+ days" },
      filtered: false,
      scope: null,
      complete: true,
      items: 3,
      generatedBy: { id: actorId, name: "Synthetic admin" },
    });
  });

  it("says when its run was limited to some domains or partial, in the file and its name", async () => {
    const response = await get(`/v1/audit-runs/${scopedRun}/handoff.json?view=d90`);
    expect(response.statusCode).toBe(200);
    expect(response.headers["content-disposition"]).toBe(
      'attachment; filename="prism365-handoff-accounts.inactive-2026-10-21-limited-to-2-domains.json"',
    );
    // The CSV of the same run carries the same suffix.
    expect(
      (await get(`/v1/audit-runs/${scopedRun}/export.csv`)).headers["content-disposition"],
    ).toContain("-limited-to-2-domains.csv");
    const document = handoffSchema.parse(response.json());
    expect(document.source).toMatchObject({
      scope: ["example.invalid", "sales.example.invalid"],
      complete: false,
    });
    expect((await get("/v1/handoffs")).json().handoffs[0]).toMatchObject({
      id: document.id,
      scope: ["example.invalid", "sales.example.invalid"],
      complete: false,
    });
  });

  it("refuses a view the audit does not have, where the CSV export takes every row", async () => {
    const before = ledger("handoff.exported").length;
    const response = await get(`/v1/audit-runs/${inactiveRun}/handoff.json?view=no-such-view`);
    expect(response.statusCode).toBe(400);
    expect(response.json()).toMatchObject({ error: "unknown_view" });
    expect(ledger("handoff.exported")).toHaveLength(before);
    expect(
      (await get(`/v1/audit-runs/${inactiveRun}/export.csv?view=no-such-view`)).statusCode,
    ).toBe(200);
  });

  it("is offered only by audits of accounts an action tool works on", async () => {
    const flagged = builtInAudits
      .map((audit) => auditDefinition(audit))
      .filter((definition) => definition.handoff)
      .map((definition) => definition.id)
      .sort();
    expect(flagged).toEqual(
      [
        "accounts.inactive",
        "guests.review",
        "identity.risky-users",
        "licenses.reclaim",
        "mfa.registration",
        "usage.m365-apps",
      ].sort(),
    );
    for (const id of [
      "devices.managed",
      "apps.consents",
      "identity.conditional-access",
      "licenses.capacity",
    ])
      expect(
        builtInAudits.find((audit) => audit.definition.id === id),
        id,
      ).toBeDefined();
    // Every audit that offers it lists accounts.
    for (const audit of builtInAudits) {
      const definition = auditDefinition(audit);
      if (definition.handoff)
        expect(
          definition.columns.some((column) => column.kind === "person"),
          definition.id,
        ).toBe(true);
    }
  });

  it("stores nothing and sends nothing when the ledger cannot record the export", async () => {
    const stored = await count("prism365_handoff");
    const append = store.appendAudit.bind(store);
    store.appendAudit = async () => {
      throw new Error("Synthetic ledger failure.");
    };
    try {
      const response = await get(`/v1/audit-runs/${inactiveRun}/handoff.json`);
      expect(response.statusCode).toBe(500);
      expect(response.body).not.toContain("example.invalid");
    } finally {
      store.appendAudit = append;
    }
    expect(await count("prism365_handoff")).toBe(stored);
  });

  it("is refused for unknown and failed runs, audits not of accounts, and bad parameters", async () => {
    for (const runId of [randomUUID(), failedRun, capacityRun, devicesRun, usageRun]) {
      const response = await get(`/v1/audit-runs/${runId}/handoff.json`);
      expect(response.statusCode, runId).toBe(404);
    }
    // Device owners are people, but a device is not an account to act on.
    for (const runId of [capacityRun, devicesRun])
      expect((await get(`/v1/audit-runs/${runId}/handoff.json`)).json().error).toBe(
        "not_handed_off",
      );
    // A usage run names accounts only by sign-in name: nothing to hand off by id.
    expect((await get(`/v1/audit-runs/${usageRun}/handoff.json`)).json()).toMatchObject({
      error: "no_accounts",
      message: "This run names no accounts by their directory id, so it has nothing to hand off.",
    });
    for (const query of ["columns=person;drop", `search=${"x".repeat(201)}`, "view=a&view=b"])
      expect((await get(`/v1/audit-runs/${inactiveRun}/handoff.json?${query}`)).statusCode).toBe(
        400,
      );
    expect((await get("/v1/audit-runs/not-a-run/handoff.json")).statusCode).toBe(400);
  });
});

describe("receipt import", () => {
  it("keeps a valid receipt and answers a repeat with what it kept", async () => {
    const id = randomUUID();
    const body = receipt(id, [item(), item({ status: "failed", detail: "Synthetic refusal." })]);
    const first = await importReceipt(body);
    expect(first.statusCode).toBe(201);
    expect(first.json()).toMatchObject({
      imported: true,
      receipt: {
        id,
        handoff: null,
        tool: { name: "Example action tool", version: "1.0.0" },
        performedBy: "Operator One",
        items: 2,
        done: 1,
        failed: 1,
        skipped: 0,
      },
    });
    // The same receipt with its keys in another order and different spacing.
    const reordered = JSON.stringify(Object.fromEntries(Object.entries(body).reverse()), null, 4);
    const again = await importReceipt(reordered);
    expect(again.statusCode).toBe(200);
    expect(again.json()).toEqual({ ...first.json(), imported: false });
    const ids = (await get("/v1/handoffs"))
      .json()
      .receipts.map((entry: { id: string }) => entry.id);
    expect(ids.filter((value: string) => value === id)).toHaveLength(1);
  });

  it("refuses a different receipt under an id already kept", async () => {
    const id = randomUUID();
    expect((await importReceipt(receipt(id, [item()]))).statusCode).toBe(201);
    const changed = await importReceipt(receipt(id, [item({ status: "skipped" })]));
    expect(changed.statusCode).toBe(409);
    expect(changed.json().error).toBe("receipt_conflict");
    const page = await get(`/v1/receipts/${id}`);
    expect(page.json().items[0].status).toBe("done");
  });

  it("links a receipt to a handoff of this tenant and refuses any other", async () => {
    const exported = handoffSchema.parse(
      (await get(`/v1/audit-runs/${inactiveRun}/handoff.json?view=d90`)).json(),
    );
    const linked = await importReceipt(receipt(randomUUID(), [item()], { handoffId: exported.id }));
    expect(linked.statusCode).toBe(201);
    expect(linked.json().receipt.handoff).toMatchObject({
      id: exported.id,
      auditId: "accounts.inactive",
      items: 3,
    });
    // A handoff id of another tenant is as unknown here as one never exported.
    const foreign = randomUUID();
    await database.tenant(otherTenant, (sql) =>
      sql.query(
        `INSERT INTO prism365_handoff (tenant_id, id, audit_id, audit_title, run_id, complete,
           items, generated_at, generated_by)
         VALUES ($1,$2,'accounts.inactive','Inactive accounts',$3,true,0,now(),$4)`,
        [otherTenant, foreign, randomUUID(), actorId],
      ),
    );
    for (const handoffId of [randomUUID(), foreign]) {
      const before = await count("prism365_receipt");
      const ledgered = ledger("receipt.imported").length;
      const response = await importReceipt(receipt(randomUUID(), [item()], { handoffId }));
      expect(response.statusCode).toBe(400);
      expect(response.json().error).toBe("unknown_handoff");
      expect(await count("prism365_receipt")).toBe(before);
      // Refused by the check before the ledger, not by the foreign key after it.
      expect(ledger("receipt.imported")).toHaveLength(ledgered);
    }
  });

  it("answers invalid receipts with the first problem in plain words", async () => {
    const before = await count("prism365_receipt");
    const cases: [unknown, string][] = [
      [receipt(randomUUID(), [item({ action: "Disable" })]), "items[0].action is not an action"],
      [receipt(randomUUID(), [item({ userId: null, userPrincipalName: null })]), "needs a userId"],
      [{ ...receipt(randomUUID(), [item()]), extra: true }, 'does not read: "extra"'],
      [receipt(randomUUID(), []), "items is empty"],
      // Past year 9999 once stored, which would break every later list of receipts.
      [
        receipt(randomUUID(), [item()], { performedAt: "9999-12-31T23:59:59-14:00" }),
        "between the years 2000 and 2099",
      ],
      // An offset PostgreSQL refuses would fail only once stored, on every retry.
      [
        receipt(randomUUID(), [item()], { performedAt: "2026-10-22T14:30:00+16:00" }),
        "time zone offset beyond 15:59",
      ],
      [
        receipt(randomUUID(), [item()], { performedAt: "2026-10-22T14:30:00-23:59" }),
        "time zone offset beyond 15:59",
      ],
      // Dated ahead of the clock, it would sort first for as long as it is kept.
      [
        receipt(randomUUID(), [item()], {
          performedAt: new Date(Date.now() + 2 * 86_400_000).toISOString(),
        }),
        "more than a day in the future",
      ],
      [
        receipt(randomUUID(), [item()], { tool: { name: "Prism365", version: "1" } }),
        "cannot be Prism365",
      ],
      [receipt(randomUUID(), [item({ action: "prism365.audit.done" })]), "kept for Prism365"],
      ["not json", "not valid JSON"],
      ["", "not valid JSON"],
    ];
    const ledgered = ledger("receipt.imported").length;
    for (const [payload, message] of cases) {
      const response = await importReceipt(payload);
      expect(response.statusCode, JSON.stringify(payload).slice(0, 120)).toBe(400);
      expect(response.json().error).toBe("invalid_receipt");
      expect(response.json().message).toContain(message);
    }
    expect(await count("prism365_receipt")).toBe(before);
    expect(ledger("receipt.imported")).toHaveLength(ledgered);
  });

  it("keeps the time it was performed in UTC, as one point in time", async () => {
    const id = randomUUID();
    const response = await importReceipt(
      receipt(id, [item()], { performedAt: "2026-10-22T14:30:00+15:59" }),
    );
    expect(response.statusCode).toBe(201);
    expect(response.json().receipt.performedAt).toBe("2026-10-21T22:31:00.000Z");
  });

  it("reads at most 4 MB, and only for a signed-in administrator", async () => {
    const huge = JSON.stringify({ padding: "x".repeat(4 * 1024 * 1024) });
    const response = await importReceipt(huge);
    expect(response.statusCode).toBe(413);
    expect(response.json().message).toContain("at most 4 MB");
    // Just under the limit is read, and then judged on its content.
    const large = receipt(
      randomUUID(),
      Array.from({ length: 5_000 }, (_, index) =>
        item({
          userId: null,
          userPrincipalName: `bulk${index}@example.invalid`,
          detail: "y".repeat(500),
        }),
      ),
    );
    expect(JSON.stringify(large).length).toBeGreaterThan(1_000_000);
    expect(JSON.stringify(large).length).toBeLessThan(4 * 1024 * 1024);
    expect((await importReceipt(large)).statusCode).toBe(201);
  }, 30_000);

  it("stores the strings it was sent cleaned", async () => {
    const id = randomUUID();
    const response = await importReceipt(
      receipt(id, [item({ detail: "Blocked\u0000 by \u202Epolicy\u202C\uD800" })], {
        tool: { name: "Example\u0000 action \u2066tool\u2069", version: "1.0.0" },
      }),
    );
    expect(response.statusCode).toBe(201);
    expect(response.json().receipt.tool.name).toBe("Example action tool");
    const page = await get(`/v1/receipts/${id}`);
    expect(page.json().items[0].detail).toBe("Blocked by policy\uFFFD");
  });

  it("digests the receipt as received, so a change in cleaning never refuses a true retry", async () => {
    const id = randomUUID();
    // Two bodies that clean to the same receipt are still two receipts.
    expect((await importReceipt(receipt(id, [item({ detail: "a\u0000b" })]))).statusCode).toBe(201);
    expect((await importReceipt(receipt(id, [item({ detail: "ab" })]))).statusCode).toBe(409);
    const stored = await database.tenant(tenantId, (sql) =>
      sql.query<{ digest: string; digest_version: number }>(
        "SELECT digest, digest_version FROM prism365_receipt WHERE id=$1",
        [id],
      ),
    );
    const { createHash } = await import("node:crypto");
    expect(stored.rows[0]).toEqual({
      digest: createHash("sha256")
        .update(canonicalJson(receipt(id, [item({ detail: "a\u0000b" })])))
        .digest("hex"),
      digest_version: 1,
    });
  });

  it("answers a receipt kept under another digest version as kept, never as a conflict", async () => {
    const id = randomUUID();
    expect((await importReceipt(receipt(id, [item()]))).statusCode).toBe(201);
    // As a later Prism365 would have stored it, with a digest this one cannot compare.
    await database.tenant(tenantId, (sql) =>
      sql.query("UPDATE prism365_receipt SET digest=$2, digest_version=2 WHERE id=$1", [
        id,
        "f".repeat(64),
      ]),
    );
    const again = await importReceipt(receipt(id, [item()]));
    expect(again.statusCode).toBe(200);
    expect(again.json().imported).toBe(false);
  });

  it("answers an import that lost the race for its id from what the winner kept", async () => {
    const id = randomUUID();
    expect((await importReceipt(receipt(id, [item()]))).statusCode).toBe(201);
    const ledgered = ledger("receipt.imported").length;
    // The check finds nothing, as if the winner stored the id a moment later: the insert is
    // then the one that finds it.
    hideStoredReceiptOnce = true;
    const same = await importReceipt(receipt(id, [item()]));
    expect(same.statusCode).toBe(200);
    expect(same.json().imported).toBe(false);
    hideStoredReceiptOnce = true;
    const different = await importReceipt(receipt(id, [item({ status: "failed" })]));
    expect(different.statusCode).toBe(409);
    hideStoredReceiptOnce = false;
    // Both went past the check, so both were recorded before the insert found the winner.
    expect(ledger("receipt.imported")).toHaveLength(ledgered + 2);
    expect(await count("prism365_receipt_item")).toBeGreaterThan(0);
    const kept = await get(`/v1/receipts/${id}`);
    expect(kept.json().items.map((entry: { status: string }) => entry.status)).toEqual(["done"]);
  });

  it("records the import first, and keeps nothing when the ledger fails", async () => {
    const id = randomUUID();
    const before = await count("prism365_receipt");
    const append = store.appendAudit.bind(store);
    store.appendAudit = async () => {
      throw new Error("Synthetic ledger failure.");
    };
    try {
      expect((await importReceipt(receipt(id, [item()]))).statusCode).toBe(500);
    } finally {
      store.appendAudit = append;
    }
    expect(await count("prism365_receipt")).toBe(before);
    expect((await importReceipt(receipt(id, [item(), item()]))).statusCode).toBe(201);
    // Ids and counts only: nothing the tool wrote in its own words.
    const entry = ledger("receipt.imported").at(-1)?.event;
    expect(entry).toEqual({
      type: "receipt.imported",
      receiptId: id,
      handoffId: null,
      items: 2,
      done: 2,
      failed: 0,
      skipped: 0,
      occurredAt: entry?.occurredAt,
    });
    expect(JSON.stringify(entry)).not.toContain("Example action tool");
  });

  it("keeps nothing of a receipt when one of its items fails to store", async () => {
    const id = randomUUID();
    const items = Array.from({ length: 300 }, (_, index) =>
      item({ userId: null, userPrincipalName: `batch${index}@example.invalid` }),
    );
    const receipts = await count("prism365_receipt");
    const stored = await count("prism365_receipt_item");
    failItemsFrom = 250;
    try {
      const failed = await importReceipt(receipt(id, items));
      expect(failed.statusCode).toBe(500);
      expect(failed.json().message).toContain("kept once");
    } finally {
      failItemsFrom = null;
    }
    expect(await count("prism365_receipt")).toBe(receipts);
    expect(await count("prism365_receipt_item")).toBe(stored);
    // Nothing of it was kept, so the same receipt now imports as new.
    const retried = await importReceipt(receipt(id, items));
    expect(retried.statusCode).toBe(201);
    expect(await count("prism365_receipt_item")).toBe((stored ?? 0) + 300);
  });
});

describe("receipt reads", () => {
  it("pages a receipt's items in its order", async () => {
    const id = randomUUID();
    const items = Array.from({ length: 5 }, (_, index) =>
      item({ userId: null, userPrincipalName: `page${index}@example.invalid` }),
    );
    expect((await importReceipt(receipt(id, items))).statusCode).toBe(201);
    const first = (await get(`/v1/receipts/${id}?limit=2`)).json();
    expect(first.items.map((entry: { position: number }) => entry.position)).toEqual([0, 1]);
    expect(first.next).toBe(1);
    const last = (await get(`/v1/receipts/${id}?limit=10&after=3`)).json();
    expect(last.items.map((entry: { position: number }) => entry.position)).toEqual([4]);
    expect(last.next).toBeNull();
    expect(ledger("receipt.viewed").at(-1)?.event).toMatchObject({ receiptId: id });
    expect((await get(`/v1/receipts/${randomUUID()}`)).statusCode).toBe(404);
    expect((await get(`/v1/receipts/${id}?limit=501`)).statusCode).toBe(400);
  });

  it("shows a person what was reported by id, or by principal name in any case", async () => {
    const userId = "0000000a-bbbb-4bbb-8bbb-00000000000a";
    const id = randomUUID();
    const body = receipt(
      id,
      [
        item({ userId, userPrincipalName: null, action: "license.removed" }),
        item({ userId: null, userPrincipalName: "Casey.Example@EXAMPLE.invalid" }),
        // Another account that happens to carry the same name is not this person.
        item({
          userId: "0000000b-bbbb-4bbb-8bbb-00000000000b",
          userPrincipalName: "casey.example@example.invalid",
        }),
        item({ userId: null, userPrincipalName: "someone.else@example.invalid" }),
      ],
      { performedAt: "2026-10-22T16:00:00Z" },
    );
    expect((await importReceipt(body)).statusCode).toBe(201);
    const older = randomUUID();
    expect(
      (
        await importReceipt(
          receipt(older, [item({ userId: userId.toUpperCase(), status: "skipped" })], {
            performedAt: "2026-10-20T08:00:00Z",
          }),
        )
      ).statusCode,
    ).toBe(201);

    const byId = (await get(`/v1/people/${userId}/receipts`)).json();
    expect(byId.items.map((entry: { receiptId: string }) => entry.receiptId)).toEqual([id, older]);
    expect(byId.items[0]).toMatchObject({
      action: "license.removed",
      matchedBy: "userId",
      tool: { name: "Example action tool" },
    });

    const both = (
      await get(
        `/v1/people/${userId}/receipts?principal=${encodeURIComponent("casey.EXAMPLE@example.INVALID")}`,
      )
    ).json();
    expect(
      both.items.map((entry: { position: number; matchedBy: string; receiptId: string }) => [
        entry.receiptId === id ? "new" : "old",
        entry.position,
        entry.matchedBy,
      ]),
    ).toEqual([
      ["new", 0, "userId"],
      ["new", 1, "userPrincipalName"],
      ["old", 0, "userId"],
    ]);
    expect(both.truncated).toBe(false);
    expect(ledger("person.receipts.viewed").at(-1)?.event).toMatchObject({
      datasets: ["receipts"],
    });
    expect((await get("/v1/people/has%20space/receipts")).statusCode).toBe(400);
  });

  it("records each read of the lists before it reads them", async () => {
    for (const [url, type] of [
      ["/v1/handoffs", "handoffs.viewed"],
      ["/v1/people/0000000d-dddd-4ddd-8ddd-00000000000d/receipts", "person.receipts.viewed"],
    ] as const) {
      const before = handoffReads;
      let readsAtLedger = -1;
      const append = store.appendAudit.bind(store);
      store.appendAudit = async (...args) => {
        readsAtLedger = handoffReads;
        return append(...args);
      };
      try {
        expect((await get(url)).statusCode, url).toBe(200);
      } finally {
        store.appendAudit = append;
      }
      expect(readsAtLedger, url).toBe(before);
      expect(handoffReads, url).toBeGreaterThan(before);
      expect(ledger(type).at(-1)?.event, url).toMatchObject({ type });
      // Without its ledger entry, nothing is read or sent.
      const reads = handoffReads;
      store.appendAudit = async () => {
        throw new Error("Synthetic ledger failure.");
      };
      try {
        const refused = await get(url);
        expect(refused.statusCode, url).toBe(500);
        expect(refused.body, url).not.toContain("items");
      } finally {
        store.appendAudit = append;
      }
      expect(handoffReads, url).toBe(reads);
    }
  });

  it("finds a person's receipts whatever the case of the id in the address", async () => {
    const userId = "0000000e-eeee-4eee-8eee-00000000000e";
    const id = randomUUID();
    expect(
      (await importReceipt(receipt(id, [item({ userId, userPrincipalName: null })]))).statusCode,
    ).toBe(201);
    const upper = (await get(`/v1/people/${userId.toUpperCase()}/receipts`)).json();
    expect(upper.items.map((entry: { receiptId: string }) => entry.receiptId)).toEqual([id]);
  });

  it("shows the newest 100 and says there were more", async () => {
    const userId = "0000000f-ffff-4fff-8fff-00000000000f";
    const items = Array.from({ length: 101 }, (_, index) =>
      item({ userId, userPrincipalName: null, detail: `Step ${index}` }),
    );
    expect((await importReceipt(receipt(randomUUID(), items))).statusCode).toBe(201);
    const listed = (await get(`/v1/people/${userId}/receipts`)).json();
    expect(listed.items).toHaveLength(100);
    expect(listed.truncated).toBe(true);
    expect(listed.items[0].detail).toBe("Step 0");
  });

  it("matches principal names by their normalised form, not the database's lower()", async () => {
    const userId = "00000010-1010-4010-8010-000000000010";
    const id = randomUUID();
    // The Kelvin sign: NFKC makes it K, and lower case k.
    expect(
      (
        await importReceipt(
          receipt(id, [item({ userId: null, userPrincipalName: "\u212Aelvin@example.invalid" })]),
        )
      ).statusCode,
    ).toBe(201);
    const found = (
      await get(
        `/v1/people/${userId}/receipts?principal=${encodeURIComponent("kelvin@example.invalid")}`,
      )
    ).json();
    expect(found.items.map((entry: { receiptId: string }) => entry.receiptId)).toEqual([id]);
    // Kept as the tool wrote it.
    expect(found.items[0].userPrincipalName).toBe("\u212Aelvin@example.invalid");
  });

  it("never shows another tenant's receipts", async () => {
    const userId = "0000000c-cccc-4ccc-8ccc-00000000000c";
    const foreign = receiptSchema.parse(
      receipt(randomUUID(), [
        item({ userId, userPrincipalName: "shared.name@example.invalid", detail: "Other tenant" }),
      ]),
    );
    await database.tenant(otherTenant, (sql) =>
      insertReceipt(sql, otherTenant, foreign, {
        digest: "0".repeat(64),
        digestVersion: 1,
        importedAt: new Date(),
        importedBy: actorId,
      }),
    );
    const seen = (
      await get(`/v1/people/${userId}/receipts?principal=shared.name@example.invalid`)
    ).json();
    expect(seen.items).toEqual([]);
    expect((await get(`/v1/receipts/${foreign.id}`)).statusCode).toBe(404);
    // Row security hides them even from a query with no tenant predicate.
    const unfiltered = await database.tenant(tenantId, (sql) =>
      sql.query<{ detail: string }>("SELECT detail FROM prism365_receipt_item"),
    );
    expect(unfiltered.rows.some((row) => row.detail === "Other tenant")).toBe(false);
    expect(await count("prism365_receipt_item", otherTenant)).toBe(1);
    // The same receipt id is free in this tenant: ids are keyed by tenant.
    expect((await importReceipt(receipt(foreign.id, [item()]))).statusCode).toBe(201);
  });
});

describe("retention", () => {
  const years = (count: number) => new Date(Date.now() - count * 365 * 86_400_000).toISOString();

  it("deletes handoffs and receipts past their retention when the next one is stored", async () => {
    const oldHandoff = randomUUID();
    const answered = randomUUID();
    const oldReceipt = randomUUID();
    const recentReceipt = randomUUID();
    await database.tenant(tenantId, async (sql) => {
      for (const id of [oldHandoff, answered])
        await sql.query(
          `INSERT INTO prism365_handoff (tenant_id, id, audit_id, audit_title, run_id, complete,
             items, generated_at, generated_by)
           VALUES ($1,$2,'accounts.inactive','Inactive accounts',$3,true,1,$4,$5)`,
          [tenantId, id, randomUUID(), years(4), actorId],
        );
      for (const [id, handoffId, at] of [
        [oldReceipt, null, years(4)],
        // A recent receipt answering an old handoff keeps that handoff.
        [recentReceipt, answered, years(1)],
      ] as const) {
        await sql.query(
          `INSERT INTO prism365_receipt (tenant_id, id, handoff_id, tool_name, tool_version,
             performed_at, performed_by, items, done_count, failed_count, skipped_count,
             imported_at, imported_by, digest, digest_version)
           VALUES ($1,$2,$3,'Example action tool','1',$4,'Operator',1,1,0,0,$4,$5,$6,1)`,
          [tenantId, id, handoffId, at, actorId, "a".repeat(64)],
        );
        await sql.query(
          `INSERT INTO prism365_receipt_item (tenant_id, receipt_id, position, user_id, action,
             status, detail) VALUES ($1,$2,0,'user-old','account.disabled','done','')`,
          [tenantId, id],
        );
      }
    });
    const ids = async (table: string, column = "id") =>
      (
        await database.tenant(tenantId, (sql) =>
          sql.query<{ id: string }>(`SELECT ${column} AS id FROM ${table}`),
        )
      ).rows.map((row) => row.id);
    expect(await ids("prism365_receipt")).toContain(oldReceipt);
    expect((await importReceipt(receipt(randomUUID(), [item()]))).statusCode).toBe(201);
    expect(await ids("prism365_receipt")).not.toContain(oldReceipt);
    expect(await ids("prism365_receipt")).toContain(recentReceipt);
    expect(await ids("prism365_receipt_item", "receipt_id")).not.toContain(oldReceipt);
    expect(await ids("prism365_handoff")).not.toContain(oldHandoff);
    expect(await ids("prism365_handoff")).toContain(answered);
    // An export cleans up too.
    await database.tenant(tenantId, (sql) =>
      sql.query(
        `INSERT INTO prism365_handoff (tenant_id, id, audit_id, audit_title, run_id, complete,
           items, generated_at, generated_by)
         VALUES ($1,$2,'accounts.inactive','Inactive accounts',$3,true,1,$4,$5)`,
        [tenantId, oldHandoff, randomUUID(), years(4), actorId],
      ),
    );
    expect((await get(`/v1/audit-runs/${inactiveRun}/handoff.json`)).statusCode).toBe(200);
    expect(await ids("prism365_handoff")).not.toContain(oldHandoff);
  });

  it("is three years unless HANDOFF_RETENTION_DAYS says otherwise, from 30 days to 10 years", async () => {
    const live = { DATA_MODE: "live" };
    expect((await loadConfig(live)).handoffRetentionDays).toBeUndefined();
    expect((await loadConfig({ ...live, HANDOFF_RETENTION_DAYS: "30" })).handoffRetentionDays).toBe(
      30,
    );
    expect(
      (await loadConfig({ ...live, HANDOFF_RETENTION_DAYS: "3650" })).handoffRetentionDays,
    ).toBe(3650);
    for (const value of ["29", "3651", "0", "-5", "90.5", "1e3", "a year"])
      await expect(loadConfig({ ...live, HANDOFF_RETENTION_DAYS: value }), value).rejects.toThrow(
        "HANDOFF_RETENTION_DAYS",
      );
  });
});

import { randomUUID } from "node:crypto";
import { finishRun, insertRun } from "@prism365/audits";
import type { AuditRow, MailboxAccessExport, PersonSummary, Receipt } from "@prism365/contracts";
import { GraphHttpClient } from "@prism365/graph";
import { ModuleScheduler } from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase, type SqlClient } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import { createApp } from "../src/app.ts";
import type { AppConfig, IdentityConfig } from "../src/config.ts";
import { insertReceipt } from "../src/handoff-store.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../src/identity.ts";
import { replaceImport } from "../src/mailbox-access-store.ts";
import { createModuleRegistry } from "../src/modules.ts";

vi.setConfig({ testTimeout: 60_000 });

// The Overview's "Where they show up" answer. Synthetic tenant, administrator and accounts only;
// every address ends in .invalid.
const tenantId = "11111111-1111-4111-8111-111111111111";
const actorId = "22222222-2222-4222-8222-222222222222";
const origin = "http://localhost:5173";
const person = "aaaaaaaa-0000-4000-8000-000000000001";
const quiet = "aaaaaaaa-0000-4000-8000-000000000002";
const MAILBOX = "eeeeeeee-0000-4000-8000-000000000001";
const principal = "pat.pell@example.invalid";
const url = (id: string) => `/v1/people/${id}/summary`;
const features = ["directory.users.read", "audit.signins.read", "audit.directory.read"] as const;
const DAY = 86_400_000;

const row = (
  key: string,
  personId: string | null,
  cells: AuditRow["cells"],
  tags: string[] = [],
): AuditRow => ({ key, personId, tags, cells });
const who = { name: "Pat Pell", principal };

async function harness(withModules = true) {
  const store = new MemoryStateStore();
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
  const requests: string[] = [];
  const graphFetch = (async (input: RequestInfo | URL) => {
    requests.push(String(input));
    return new Response("{}", { status: 404 });
  }) as typeof fetch;
  const identity = new IdentityService(identityConfig, store, provider);
  const login = await identity.startLogin("directory.users.read");
  const loggedIn = await identity.finishLogin("synthetic-code", login.state, login.state);
  const cookie = `prism365_session=${loggedIn.handle}`;
  const registry = createModuleRegistry(config);
  const database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  const scheduler = new ModuleScheduler({
    enabledPacks: config.enabledFeatures,
    registry,
    database,
    graphClient: (getAccessToken) => new GraphHttpClient({ getAccessToken }),
  });
  const app = createApp({
    config,
    store,
    identity,
    graphFetch,
    ...(withModules ? { modules: { registry, database, scheduler } } : {}),
  });
  const get = (path: string, signedIn = true) =>
    app.inject({ method: "GET", url: path, headers: signedIn ? { cookie } : {} });
  let runs = 0;
  const storeRun = async (
    auditId: string,
    rows: AuditRow[],
    outcome: { complete?: boolean; finished?: string } = {},
  ) => {
    runs += 1;
    const id = `00000000-0000-4000-8000-${String(runs).padStart(12, "0")}`;
    await database.tenant(tenantId, async (sql) => {
      await insertRun(sql, tenantId, {
        id,
        auditId,
        auditVersion: 1,
        startedAt: new Date("2026-10-20T10:00:00Z"),
        startedBy: { id: actorId, name: null },
      });
      await finishRun(sql, tenantId, id, new Date(outcome.finished ?? "2026-10-20T10:05:00Z"), {
        status: "succeeded",
        complete: outcome.complete ?? true,
        summary: [],
        notes: [],
        rows,
      });
    });
  };
  const sql = <T>(work: (client: SqlClient) => Promise<T>) => database.tenant(tenantId, work);
  return { app, get, store, requests, storeRun, sql, database };
}
type Harness = Awaited<ReturnType<typeof harness>>;
const ledger = (h: Harness) => h.store.audit.map((entry) => entry.event as Record<string, unknown>);
async function summary(h: Harness, id = person): Promise<PersonSummary> {
  const response = await h.get(url(id));
  expect(response.statusCode).toBe(200);
  return response.json();
}

describe("person summary API", () => {
  let h: Harness;
  beforeAll(async () => {
    h = await harness();
  }, 60_000);
  afterAll(async () => {
    await h?.app.close();
  });

  it("needs a session and a usable account id, and says plainly when there is no module storage", async () => {
    expect((await h.get(url(person), false)).statusCode).toBe(401);
    expect((await h.get(url("..."))).statusCode).toBe(400);
    expect((await h.get(url("has%20space"))).statusCode).toBe(400);
    expect(ledger(h).filter((event) => event.type === "person.summary.viewed")).toEqual([]);
    const bare = await harness(false);
    const response = await bare.get(url(person));
    expect(response.statusCode).toBe(409);
    expect(response.json().message).toMatch(/module storage/);
    await bare.app.close();
  });

  it("writes the ledger entry first, with dataset names only, and asks Microsoft nothing", async () => {
    let at: number | null = null;
    const original = h.database.tenant.bind(h.database);
    h.database.tenant = ((...args: Parameters<typeof original>) => {
      at ??= h.store.audit.length;
      return original(...args);
    }) as typeof original;
    const before = h.store.audit.length;
    await summary(h);
    h.database.tenant = original;
    expect(at).toBeGreaterThan(before);
    const entry = ledger(h).find((event) => event.type === "person.summary.viewed");
    expect(entry?.datasets).toEqual([
      "auditResults",
      "signIns",
      "findings",
      "mailboxAccess",
      "directoryAudit",
      "receipts",
    ]);
    expect(JSON.stringify(entry)).not.toContain(person);
    expect(h.requests).toEqual([]);
  });

  it("says not run for every source before anything has run", async () => {
    const empty = await summary(h);
    expect(
      empty.audits.every((audit) => audit.status === "no_run" || audit.status === "no_name"),
    ).toBe(true);
    expect(empty.combinedRisks).toEqual({ status: "no_run" });
    expect(empty.owned).toEqual({ status: "no_run" });
    expect(empty.sharedMailboxes).toEqual({ status: "no_import" });
    expect(empty.signinFindings).toEqual({ status: "ok", open: 0, highest: null, more: false });
    expect(empty.changes).toMatchObject({ status: "ok", count: 0, days: 30 });
    expect(empty.handoff).toEqual({ status: "ok", latest: null });
  });

  it("gives every source its state: rows, not in this run, a stale run's date, no sign-in name", async () => {
    await h.storeRun(
      "identity.risky-users",
      [row("r1", person, { level: "High", state: "At risk", detail: "x", updated: null })],
      { finished: "2026-08-01T10:00:00Z", complete: false },
    );
    await h.storeRun("mfa.registration", [row("q", quiet, { registered: true })]);
    await h.storeRun("accounts.activity", [
      row("a1", person, { person: who, verdict: "Active", lastSuccess: null }),
    ]);
    await h.storeRun("usage.m365-apps", [
      row(principal, null, { person: who, lastActivity: null, days: 12, apps: ["Teams"] }),
      row("hashed", null, {
        person: { name: null, principal: "0123456789abcdef0123456789abcdef" },
      }),
    ]);
    const result = await summary(h);
    const by = new Map(result.audits.map((audit) => [audit.auditId, audit]));
    // A stale, incomplete run still shows its own date and completeness.
    expect(by.get("identity.risky-users")).toMatchObject({
      status: "ok",
      finishedAt: "2026-08-01T10:00:00.000Z",
      complete: false,
      matchedBy: "account",
    });
    expect(by.get("mfa.registration")).toMatchObject({ status: "not_in_run", complete: true });
    expect(by.get("passwords")).toBeUndefined();
    expect(by.get("accounts.passwords")).toMatchObject({ status: "no_run" });
    // The sign-in name comes from the person's own account cell, so the app usage row is found.
    expect(by.get("usage.m365-apps")).toMatchObject({ status: "ok", matchedBy: "sign-in name" });
    // No stored name for this account, and a mailbox run exists: it cannot be matched.
    await h.storeRun("mail.mailbox-storage", [
      row(principal, null, { mailbox: principal, used: 1 }),
    ]);
    const other = await summary(h, "aaaaaaaa-0000-4000-8000-0000000000cc");
    expect(other.audits.find((a) => a.auditId === "mail.mailbox-storage")).toMatchObject({
      status: "no_name",
    });
    expect(other.audits.find((a) => a.auditId === "usage.m365-apps")).toMatchObject({
      status: "no_name",
    });
  });

  it("matches a mailbox by sign-in name from stored sign-ins and never on a hash", async () => {
    await h.sql((client) =>
      client.query(
        `INSERT INTO mod_signins_audit_event (tenant_id, id, created_at, user_id, upn, error_code,
           risk_event_types, backfilled, processed, collected_at)
         VALUES ($1,'e1',$2,$3,$4,0,'[]'::jsonb,false,true,$2)`,
        [tenantId, "2026-10-20T10:00:00.000Z", quiet, "Lee.Park@Example.invalid"],
      ),
    );
    await h.storeRun("mail.mailbox-storage", [
      row("m1", null, { mailbox: "lee.park@example.invalid", used: 2 }),
      row("m2", null, { mailbox: "0123456789abcdef0123456789abcdef", used: 3 }),
    ]);
    const result = await summary(h, quiet);
    const mailbox = result.audits.find((audit) => audit.auditId === "mail.mailbox-storage");
    expect(mailbox).toMatchObject({ status: "ok", matchedBy: "sign-in name" });
    expect(mailbox?.status === "ok" && mailbox.rows.map((r) => r.key)).toEqual(["m1"]);
  });

  it("lists the groups whose hidden owner ids hold the person", async () => {
    await h.storeRun("groups.teams", [
      row(
        "g1",
        null,
        { group: "Plant team", kind: "Team", status: "One owner", ownerIds: [person, quiet] },
        ["team", "one-owner"],
      ),
      row("g2", null, { group: "Other", kind: "Security group", status: "OK", ownerIds: [quiet] }),
      row("g3", null, {
        group: "Unread",
        kind: "Security group",
        status: "Owners not read",
        ownerIds: null,
      }),
    ]);
    const result = await summary(h, person.toUpperCase());
    expect(result.owned).toMatchObject({ status: "ok", more: false, complete: true });
    expect(result.owned.status === "ok" && result.owned.entries).toEqual([
      {
        id: "g1",
        name: "Plant team",
        kind: "Team",
        status: "One owner",
        tags: ["team", "one-owner"],
      },
    ]);
  });

  it("reads combined risks, open sign-in findings and the highest severity", async () => {
    await h.storeRun("identity.combined-risks", [
      row("c1", person, { reasons: "Synthetic reason." }, ["privileged-no-mfa"]),
    ]);
    await h.sql(async (client) => {
      for (const [severity, state] of [
        ["medium", "open"],
        ["high", "acknowledged"],
        ["critical", "resolved"],
      ] as const) {
        const id = randomUUID();
        await client.query(
          `INSERT INTO prism365_finding (id, tenant_id, module, rule_id, rule_version, dedupe_key,
             purpose, severity, score, subject_kind, subject_id, subject_label, title, summary,
             evidence, limitations, window_start, window_end, state, occurrences, first_seen, last_seen)
           VALUES ($1,$2,'signins.audit','signins.unfamiliar_signin',1,$7,'security',$3,50,'person',$4,'Pat','T','S',
             '[]'::jsonb,'[]'::jsonb,$5,$5,$6,1,$5,$5)`,
          [id, tenantId, severity, person, "2026-10-20T10:00:00Z", state, `key-${id}`],
        );
      }
    });
    const result = await summary(h);
    expect(result.combinedRisks).toMatchObject({ status: "ok" });
    expect(result.signinFindings).toEqual({ status: "ok", open: 2, highest: "high", more: false });
  });

  it("counts direct shared mailbox grants and who holds the person's own mailbox", async () => {
    const file = {
      format: "prism365.mailbox-access",
      version: 1,
      tenantId,
      exportedAt: new Date(Date.now() - 2 * DAY).toISOString(),
      includeUserMailboxes: false,
      mailboxes: [
        { id: MAILBOX, address: "lab@example.invalid", displayName: "Lab", type: "SharedMailbox" },
        { id: person, address: principal, displayName: "Pat", type: "UserMailbox" },
      ],
      grants: [
        {
          mailboxId: MAILBOX,
          right: "FullAccess",
          trustee: { id: person, address: principal, name: "Pat", kind: "user" },
        },
        {
          mailboxId: MAILBOX,
          right: "SendAs",
          trustee: { id: person, address: principal, name: "Pat", kind: "user" },
        },
        // A grant to a group is not expanded here.
        {
          mailboxId: MAILBOX,
          right: "FullAccess",
          trustee: {
            id: "bbbbbbbb-0000-4000-8000-000000000003",
            address: null,
            name: "Staff",
            kind: "group",
          },
        },
        {
          mailboxId: person,
          right: "FullAccess",
          trustee: { id: quiet, address: null, name: "Lee", kind: "user" },
        },
        {
          mailboxId: person,
          right: "SendAs",
          trustee: { id: quiet, address: null, name: "Lee", kind: "user" },
        },
      ],
      unresolved: 0,
    } as MailboxAccessExport;
    await h.sql((client) =>
      replaceImport(client, tenantId, file, {
        id: randomUUID(),
        importedBy: { id: actorId, name: null },
      }),
    );
    const result = await summary(h);
    expect(result.sharedMailboxes).toMatchObject({
      status: "ok",
      mailboxes: 1,
      holders: 1,
      includesUserMailboxes: false,
    });
    const none = await summary(h, "aaaaaaaa-0000-4000-8000-0000000000dd");
    expect(none.sharedMailboxes).toMatchObject({ status: "ok", mailboxes: 0, holders: null });
  });

  it("counts directory changes in the last 30 days and shows the latest handoff", async () => {
    await h.sql(async (client) => {
      for (const [id, daysAgo] of [
        ["c-new", 2],
        ["c-old", 45],
      ] as const) {
        const at = new Date(Date.now() - daysAgo * DAY).toISOString();
        await client.query(
          `INSERT INTO mod_directory_changes_event (tenant_id, id, activity_at, category, activity,
             operation_type, result, initiated_by_kind, initiated_by_id, initiated_by_name,
             initiated_by_principal, targets, details, search_text, processed, collected_at)
           VALUES ($1,$2,$3,'UserManagement','Update user','Update','success','user',$4,'Admin',null,'[]'::jsonb,'[]'::jsonb,'',true,$3)`,
          [tenantId, id, at, actorId],
        );
        await client.query(
          `INSERT INTO mod_directory_changes_subject (tenant_id, event_id, subject_id, subject_kind, subject_type, activity_at)
           VALUES ($1,$2,$3,'target','User',$4)`,
          [tenantId, id, person, at],
        );
      }
      const receipt = {
        id: randomUUID(),
        handoffId: null,
        tool: { name: "Example action tool", version: "1.0.0" },
        performedAt: "2026-10-21T10:00:00.000Z",
        performedBy: "Jordan Vale",
        items: [
          {
            userId: person,
            userPrincipalName: null,
            action: "account.sign-in-blocked",
            status: "done",
            detail: "Done.",
          },
        ],
      } as unknown as Receipt;
      await insertReceipt(client, tenantId, receipt, {
        digest: "a".repeat(64),
        digestVersion: 1,
        importedBy: actorId,
      });
    });
    const result = await summary(h);
    expect(result.changes).toEqual({ status: "ok", days: 30, count: 1, more: false });
    expect(result.handoff).toEqual({
      status: "ok",
      latest: {
        state: "done",
        action: "account.sign-in-blocked",
        tool: "Example action tool",
        performedAt: "2026-10-21T10:00:00.000Z",
      },
    });
  });

  it("keeps a failing section from taking the rest down", async () => {
    const original = h.database.tenant.bind(h.database);
    h.database.tenant = ((
      tenant: string,
      work: (sql: SqlClient) => Promise<unknown>,
      options?: never,
    ) =>
      original(
        tenant,
        (sql) =>
          work({
            query: ((text: string, params?: readonly unknown[]) =>
              text.includes("prism365_receipt")
                ? Promise.reject(new Error("Synthetic storage failure."))
                : sql.query(text, params)) as SqlClient["query"],
          } as SqlClient),
        options,
      )) as typeof original;
    const result = await summary(h);
    h.database.tenant = original;
    expect(result.handoff.status).toBe("failed");
    expect(JSON.stringify(result)).not.toContain("Synthetic storage failure");
    expect(result.owned.status).toBe("ok");
  });
});

// A sign-in name is reused when its owner is renamed and the old name goes to a new hire. These
// tests walk that sequence: Ana (person) was ana.old and is now ana.new; Ben (quiet) was hired
// with ana.old afterwards.
describe("person summary, sign-in names that change hands", () => {
  const OLD = "ana.old@example.invalid";
  const NEW = "ana.new@example.invalid";
  const seed = (h: Harness, user: string, upn: string, at: string) =>
    h.sql((client) =>
      client.query(
        `INSERT INTO mod_signins_audit_event (tenant_id, id, created_at, user_id, upn, error_code,
           risk_event_types, backfilled, processed, collected_at)
         VALUES ($1,$2,$3,$4,$5,0,'[]'::jsonb,false,true,$3)`,
        [tenantId, `${user}-${at}`, at, user, upn],
      ),
    );
  const mailboxRows = [
    row("old", null, { mailbox: OLD, used: 1 }),
    row("new", null, { mailbox: NEW, used: 2 }),
  ];
  const mailboxOf = (result: PersonSummary) =>
    result.audits.find((a) => a.auditId === "mail.mailbox-storage");
  const keysOf = (result: PersonSummary) => {
    const audit = mailboxOf(result);
    return audit?.status === "ok" ? audit.rows.map((r) => r.key) : audit?.status;
  };
  let h: Harness;
  const fresh = async () => {
    h = await harness();
    return h;
  };
  afterAll(async () => {
    await h?.app.close();
  });

  it("uses the name in the account's own newest rows, so each person gets their own mailbox", async () => {
    await fresh();
    await seed(h, person, OLD, "2026-09-01T10:00:00.000Z");
    await seed(h, quiet, OLD, "2026-10-01T10:00:00.000Z");
    await h.storeRun("accounts.activity", [
      row("a", person, { person: { name: "Ana", principal: NEW }, verdict: "Active" }),
      row("b", quiet, { person: { name: "Ben", principal: OLD }, verdict: "Active" }),
    ]);
    await h.storeRun("mail.mailbox-storage", mailboxRows);
    await h.sql((client) =>
      insertReceipt(
        client,
        tenantId,
        {
          id: randomUUID(),
          handoffId: null,
          tool: { name: "Example action tool", version: "1.0.0" },
          performedAt: "2026-10-21T10:00:00.000Z",
          performedBy: "Jordan Vale",
          items: [
            {
              userId: null,
              userPrincipalName: OLD,
              action: "license.removed",
              status: "done",
              detail: "Done.",
            },
          ],
        } as unknown as Receipt,
        { digest: "b".repeat(64), digestVersion: 1, importedBy: actorId },
      ),
    );
    const ana = await summary(h, person);
    const ben = await summary(h, quiet);
    expect(keysOf(ana)).toEqual(["new"]);
    expect(keysOf(ben)).toEqual(["old"]);
    // The receipt names only the old sign-in name, which is Ben's now.
    expect(ana.handoff).toEqual({ status: "ok", latest: null });
    expect(ben.handoff.status === "ok" && ben.handoff.latest?.action).toBe("license.removed");
  });

  it("does not use a stored sign-in name that another account has signed in with since", async () => {
    await fresh();
    await seed(h, person, OLD, "2026-09-01T10:00:00.000Z");
    await seed(h, quiet, OLD, "2026-10-01T10:00:00.000Z");
    await h.storeRun("mail.mailbox-storage", mailboxRows);
    // Ana has no rows of her own to say her new name, and Ben has taken the old one.
    expect(keysOf(await summary(h, person))).toBe("no_name");
    expect(keysOf(await summary(h, quiet))).toEqual(["old"]);
  });

  it("drops a name that a row in a run gives to a different account", async () => {
    await fresh();
    await seed(h, person, OLD, "2026-10-01T10:00:00.000Z");
    await h.storeRun("accounts.activity", [
      row("b", quiet, { person: { name: "Ben", principal: OLD }, verdict: "Active" }),
    ]);
    await h.storeRun("mail.mailbox-storage", mailboxRows);
    expect(keysOf(await summary(h, person))).toBe("no_name");
  });

  it("says it could not be matched when the account's own rows disagree", async () => {
    await fresh();
    await h.storeRun("accounts.activity", [
      row("a", person, { person: { name: "Ana", principal: NEW }, verdict: "Active" }),
    ]);
    await h.storeRun("mfa.registration", [
      row("a", person, { person: { name: "Ana", principal: OLD }, registered: true }),
    ]);
    await h.storeRun("mail.mailbox-storage", mailboxRows);
    expect(keysOf(await summary(h, person))).toBe("no_name");
  });

  it("takes the newest run's name when an older run still has the old one", async () => {
    await fresh();
    await h.storeRun(
      "mfa.registration",
      [row("a", person, { person: { name: "Ana", principal: OLD }, registered: true })],
      { finished: "2026-10-01T10:00:00Z" },
    );
    await h.storeRun("accounts.activity", [
      row("a", person, { person: { name: "Ana", principal: NEW }, verdict: "Active" }),
    ]);
    await h.storeRun("mail.mailbox-storage", mailboxRows);
    expect(keysOf(await summary(h, person))).toEqual(["new"]);
  });

  it("ignores a hashed name in the account's rows instead of matching it", async () => {
    const hashed = "0123456789abcdef0123456789abcdef";
    await fresh();
    await h.storeRun("accounts.activity", [
      row("a", person, { person: { name: "Ana", principal: hashed }, verdict: "Active" }),
    ]);
    await h.storeRun("mail.mailbox-storage", [row("hash", null, { mailbox: hashed, used: 1 })]);
    // No usable name, so the mailbox cannot be matched. Taking the hash as a name would find the
    // run (not in it), not refuse.
    expect(keysOf(await summary(h, person))).toBe("no_name");
  });
});

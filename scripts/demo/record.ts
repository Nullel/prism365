import { randomUUID } from "node:crypto";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { canonicalJson } from "@prism365/contracts";
import { GraphHttpClient } from "@prism365/graph";
import {
  DIRECTORY_CHANGES_MODULE as CHANGES,
  changeCoverage,
} from "@prism365/module-directory-changes";
import {
  MODULE_ID as SIGNINS,
  signinsAuditConfigSchema,
  summarizeSignins,
} from "@prism365/module-signins-audit";
import { latestRuns, ModuleScheduler } from "@prism365/modules";
import { MemoryStateStore, PGliteModuleDatabase } from "@prism365/storage";
import { createApp } from "../../apps/api/src/app.ts";
import { AuditRunner } from "../../apps/api/src/audits.ts";
import type { AppConfig, IdentityConfig } from "../../apps/api/src/config.ts";
import { IdentityService, type MicrosoftIdentityProvider } from "../../apps/api/src/identity.ts";
import { createModuleRegistry } from "../../apps/api/src/modules.ts";
import { type Fixtures, normalizeKey } from "../../apps/web/src/demo/replay.ts";
import { demoMailboxAccessExport } from "./mailbox-access.ts";
import {
  assertSynthetic,
  buildSyntheticTenant,
  createSyntheticGraph,
  guid,
  HOME_FAILURE_IP,
  SYNTHETIC_ACTOR_ID,
  SYNTHETIC_TENANT_ID,
} from "./tenant.ts";

// Records the portfolio demo.
//
// The real API runs in this process against the synthetic Microsoft tenant in tenant.ts: the
// real collectors read it, the real rules judge it, the real audits run on it. Then every
// request the web app makes is replayed against that API and its answer is saved, so the static
// demo shows exactly what Prism365 produced. Nothing touches Microsoft, Docker, or any running
// server: the API is never listened on, requests go straight into it.

// The words the demo's Quick search example looks groups, roles and devices up by: the company's
// own name, which the invented groups, policies and devices carry.
const DEMO_SEARCH = "larkspur";

const HOUR = 3_600_000;
const DAY = 24 * HOUR;
const root = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const output = resolve(root, "apps/web/demo/fixtures.json");
const started = Date.now();
const log = (message: string) =>
  console.log(`[${((Date.now() - started) / 1_000).toFixed(1).padStart(6)}s] ${message}`);

const features = [
  "directory.users.read",
  "user.licenses.read",
  "tenant.licenses.read",
  "audit.signins.read",
  "reports.authentication-methods.read",
  "service.health.read",
  "service.messages.read",
  "audit.directory.read",
  "identity.conditional-access.read",
  "identity.policies.read",
  "identity.risk.read",
  "devices.intune.read",
  "directory.applications.read",
  "directory.permission-grants.read",
  "directory.roles.read",
  "directory.groups.read",
  "directory.domains.read",
  "reports.usage.read",
  "reports.settings.read",
] as const;
const scopes = [
  "User.Read",
  "User.Read.All",
  "LicenseAssignment.Read.All",
  "AuditLog.Read.All",
  "ServiceHealth.Read.All",
  "ServiceMessage.Read.All",
  "Policy.Read.All",
  "IdentityRiskyUser.Read.All",
  "DeviceManagementManagedDevices.Read.All",
  "Application.Read.All",
  "Directory.Read.All",
  "Reports.Read.All",
  "ReportSettings.Read.All",
  "RoleManagement.Read.Directory",
  "Group.Read.All",
  "Domain.Read.All",
];
const origin = "http://localhost:5173";
const tenantId = SYNTHETIC_TENANT_ID;
const actorId = SYNTHETIC_ACTOR_ID;

const identityConfig: IdentityConfig = {
  tenantId,
  clientId: guid("client"),
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
  // An organisation in the United States, working in Central time.
  signinsAudit: signinsAuditConfigSchema.parse({
    expectedCountries: ["US"],
    timeZone: "America/Chicago",
  }),
};

// The administrator's Microsoft account, granting every enabled read. No sign-in leaves the process.
const provider: MicrosoftIdentityProvider = {
  authorizationUrl: async () => "https://login.microsoftonline.com/synthetic",
  redeem: async () => ({
    tenantId,
    actorId,
    displayName: "Demo administrator",
    accountId: "synthetic-account",
    cache: "synthetic-cache",
    scopes,
  }),
  token: async (session) => ({ accessToken: "synthetic", cache: "c", scopes: session.scopes }),
};

interface Login {
  cookie: string;
  csrf: string;
}

async function main() {
  // The synthetic tenant is anchored here: its newest sign-ins are minutes old.
  const anchor = new Date();
  const tenant = buildSyntheticTenant(anchor);
  const graph = createSyntheticGraph({ now: anchor });

  const store = new MemoryStateStore();
  const identity = new IdentityService(identityConfig, store, provider);
  const registry = createModuleRegistry(config);
  // An on-disk PGlite in a temporary directory, as a real install with PRISM_DATA_DIR uses, so the
  // recorded screens describe durable storage. It is deleted when the recording ends.
  const dataDir = await mkdtemp(join(tmpdir(), "prism365-demo-"));
  const database = await PGliteModuleDatabase.open({
    dataDir,
    scopes: registry.migrationScopes(),
  });
  // Collection starts three hours "ago", so the first run loads the history and a second run,
  // after the clock catches up, follows the newest sign-ins the way a running install would.
  let clock = new Date(anchor.getTime() - 3 * HOUR);
  const scheduler = new ModuleScheduler({
    enabledPacks: config.enabledFeatures,
    registry,
    database,
    now: () => clock,
    graphClient: (getAccessToken) =>
      new GraphHttpClient({
        getAccessToken,
        fetchImplementation: graph,
        requestTimeoutMilliseconds: 90_000,
        maxResponseBytes: 8_388_608,
        maxRetries: 3,
        maxRetryWaitMilliseconds: 30_000,
      }),
  });
  // Audits run on their own clock, against the synthetic tenant as it stood at that time, so the
  // recording holds runs from earlier days and every audit's trend has points. Both go back to
  // the present once the audits have run.
  let auditClock: Date | null = null;
  let auditGraph = graph;
  const auditRunner = new AuditRunner({ database, now: () => auditClock ?? new Date() });
  // The full review stamps its times on the audits' clock, and does not wait out Microsoft's
  // spacing: the synthetic tenant has no limits to keep.
  const auditReview = { now: () => auditClock ?? new Date(), pause: async () => {} };
  const app = createApp({
    config,
    store,
    identity,
    graphFetch: ((input, init) => auditGraph(input, init)) as typeof fetch,
    audits: auditRunner,
    auditReview,
    modules: { registry, database, scheduler },
  });
  log("API ready on an in-memory database.");

  // The API limits one session to 60 requests a minute, so recording signs in again regularly.
  async function signIn(): Promise<Login> {
    const login = await identity.startLogin("directory.users.read");
    const done = await identity.finishLogin("synthetic-code", login.state, login.state);
    return {
      cookie: `prism365_session=${done.handle}`,
      csrf: (await identity.session(done.handle))?.csrfToken ?? "",
    };
  }
  const first = await signIn();
  let current = first;
  let used = 0;
  async function call(method: "GET" | "POST", url: string, payload?: unknown) {
    if (used >= 45) {
      current = await signIn();
      used = 0;
    }
    used += 1;
    return app.inject({
      method,
      url,
      headers: {
        cookie: current.cookie,
        ...(method === "POST"
          ? {
              origin,
              "x-csrf-token": current.csrf,
              ...(payload === undefined ? {} : { "content-type": "application/json" }),
            }
          : {}),
      },
      ...(payload === undefined ? {} : { payload: JSON.stringify(payload) }),
    });
  }
  async function getJson<T>(url: string): Promise<T> {
    const response = await call("GET", url);
    if (response.statusCode !== 200)
      throw new Error(`GET ${url} answered ${response.statusCode}: ${response.body.slice(0, 200)}`);
    return response.json<T>();
  }

  // 1. Collect. The first tick loads the history; the second reads what arrived since. The
  // training account is marked as a canary first, so every sign-in it has is read as one.
  const marked = await call("POST", "/api/v1/settings/canaries", {
    add: [tenant.stories.canary ?? ""],
  });
  if (marked.statusCode !== 200) throw new Error(`The canary was not marked: ${marked.body}`);
  const started1 = await app.inject({
    method: "POST",
    url: "/api/v1/collection/start",
    headers: { cookie: first.cookie, origin, "x-csrf-token": first.csrf },
  });
  if (started1.statusCode !== 200) throw new Error(`Collection did not start: ${started1.body}`);
  const status = async () => {
    const runs = await database.tenant(tenantId, (sql) => latestRuns(sql, tenantId));
    const summary = await database.tenant(tenantId, (sql) =>
      summarizeSignins(sql, tenantId, new Date(), 7),
    );
    const changes = await database.tenant(tenantId, (sql) => changeCoverage(sql, tenantId));
    const run = (module: string) => runs.find((candidate) => candidate.module === module);
    return {
      run,
      phase: summary.history.phase,
      changes: changes.phase,
      events: summary.coverage.events,
    };
  };
  const collected = ["signins.audit", "service.health", "service.messages", CHANGES];
  const converged = async (at: Date, phase: string) => {
    const now = await status();
    return (
      now.phase === phase &&
      now.changes === phase &&
      collected.every((module) => {
        const run = now.run(module);
        return run?.status === "succeeded" && Date.parse(run.startedAt) === at.getTime();
      })
    );
  };
  async function collectUntil(at: Date, phase: string) {
    for (let attempt = 1; attempt <= 12; attempt += 1) {
      await scheduler.tick();
      if (await converged(at, phase)) return;
    }
    const now = await status();
    throw new Error(
      `Collection did not converge: phase ${now.phase}, changes ${now.changes}, runs ${JSON.stringify(
        collected.map((module) => now.run(module)),
      )}`,
    );
  }
  await collectUntil(clock, "live");
  log(`History loaded: ${(await status()).events} sign-ins stored.`);
  clock = anchor;
  await collectUntil(clock, "live");
  log("Newest sign-ins read; service health, Message center and directory changes checked.");

  // 2. The stories must have been found by the real rules.
  const findings = (
    await getJson<{ items: Finding[] }>(`/api/v1/findings?module=${SIGNINS}&limit=100`)
  ).items;
  const find = (ruleId: string, subject: string) =>
    findings.find((item) => item.ruleId === ruleId && item.subject.id === subject);
  const expected = {
    travel: find("signins.unfamiliar_signin", tenant.stories.traveler ?? ""),
    spray: findings.find((item) => item.ruleId === "signins.password_spray"),
    mfa: find("signins.mfa_denials", tenant.stories.mfaTarget ?? ""),
    afterFailures: find("signins.unfamiliar_signin", tenant.stories.forgetful ?? ""),
    burstAccount: find("signins.failed_burst_account", tenant.stories.forgetful ?? ""),
    burstAddress: findings.find(
      (item) => item.ruleId === "signins.failed_burst_ip" && item.summary.includes(HOME_FAILURE_IP),
    ),
    canary: find("signin.canary-used", tenant.stories.canary ?? ""),
  };
  for (const [name, finding] of Object.entries(expected))
    if (!finding) throw new Error(`The rules did not flag the planted story "${name}".`);
  if (expected.travel?.severity !== "critical")
    throw new Error(`Impossible travel was rated ${expected.travel?.severity}, not critical.`);
  if (expected.canary?.severity !== "critical")
    throw new Error(`The canary was rated ${expected.canary?.severity}, not critical.`);
  log(`Rules flagged ${findings.length} findings, including all seven planted stories.`);
  const changeFindings = (
    await getJson<{ items: Finding[] }>(`/api/v1/findings?module=${CHANGES}&limit=100`)
  ).items;
  for (const ruleId of [
    "role.privileged-assigned",
    "app.credential-added",
    "policy.conditional-access-changed",
    "app.consent-granted",
  ])
    if (!changeFindings.some((item) => item.ruleId === ruleId))
      throw new Error(`The change rules did not flag the planted ${ruleId} story.`);
  log(`Change rules flagged ${changeFindings.length} directory changes.`);

  // 3. Leave the findings in a lived-in state: one acknowledged, two resolved, the rest open.
  const setState = async (finding: Finding | undefined, state: string) => {
    if (!finding) return;
    const response = await call("POST", `/api/v1/findings/${finding.id}/state`, { state });
    if (response.statusCode !== 200)
      throw new Error(`Finding state change failed: ${response.body}`);
  };
  await setState(expected.spray, "acknowledged");
  await setState(expected.burstAccount, "resolved");
  await setState(expected.burstAddress, "resolved");

  // 4. Run a full review: five weeks ago, two weeks ago, a week ago and now, so each audit has a
  // trend and the This week page has a result from before even its longest range to compare with.
  // The review runs the audits in its own order, Combined risks last, and finishes before the next
  // round starts.
  for (const daysAgo of [35, 14, 7, 0]) {
    auditClock = daysAgo === 0 ? null : new Date(anchor.getTime() - daysAgo * DAY);
    auditGraph = auditClock ? createSyntheticGraph({ now: auditClock, quietFrom: anchor }) : graph;
    const response = await call("POST", "/api/v1/audits/review");
    if (response.statusCode !== 202)
      throw new Error(`The full review did not start: ${response.body}`);
    type Review = {
      state: string;
      audits: { auditId: string; state: string; reason: string | null }[];
    };
    let review = response.json<{ review: Review }>().review;
    for (let poll = 0; ["running", "stopping"].includes(review.state) && poll < 2_400; poll += 1) {
      await new Promise((done) => setTimeout(done, 250));
      review = (await getJson<{ review: Review }>("/api/v1/audits/review")).review;
    }
    const problems = review.audits.filter((item) => item.state !== "done");
    if (review.state !== "finished" || problems.length > 0)
      throw new Error(
        `The full review ${review.state}: ${problems.map((item) => `${item.auditId} ${item.state} ${item.reason ?? ""}`).join("; ")}`,
      );
    log(`Full review${daysAgo ? ` (${daysAgo} days ago)` : ""}: ${review.audits.length} audits.`);
  }
  auditClock = null;
  auditGraph = graph;
  // Two audits run on their own. Saved after the last runs, so neither is due while recording.
  const scheduled = await call("POST", "/api/v1/settings/audit-schedule", {
    audits: { "accounts.inactive": "weekly", "mfa.registration": "daily" },
  });
  if (scheduled.statusCode !== 200)
    throw new Error(`The audit schedule was not saved: ${scheduled.body}`);
  // The full review runs on its own every Monday morning. Saved after the last review, so it is
  // not due while recording.
  const reviewScheduled = await call("POST", "/api/v1/settings/audit-review", {
    frequency: "weekly",
    weekday: 0,
    hour: 6,
  });
  if (reviewScheduled.statusCode !== 200)
    throw new Error(`The full review schedule was not saved: ${reviewScheduled.body}`);

  // 5. Record every request the web app makes.
  const responses: Fixtures["responses"] = {};
  async function record<T = unknown>(url: string, allowed: number[] = [200]): Promise<T> {
    const response = await call("GET", url);
    if (!allowed.includes(response.statusCode))
      throw new Error(`GET ${url} answered ${response.statusCode}: ${response.body.slice(0, 200)}`);
    const body = response.json<T>();
    responses[normalizeKey("GET", url)] = { status: response.statusCode, body };
    return body;
  }

  const session = await record<{ displayName: string | null; csrfToken: string | null }>(
    "/api/v1/session",
  );
  // The recording is of a signed-in administrator, whatever login produced it.
  session.displayName = "Demo administrator";
  session.csrfToken = "demo";
  await record("/api/v1/modules");
  await record("/api/v1/dashboard/summary", [200, 409]);
  for (const state of ["open", "acknowledged", "resolved"])
    await record(`/api/v1/findings?module=${SIGNINS}&state=${state}&limit=100`);
  for (const days of [7, 30]) await record(`/api/v1/modules/${SIGNINS}/summary?days=${days}`);
  for (const days of [7, 30]) await record(`/api/v1/people?days=${days}`);
  for (const state of ["open", "acknowledged", "resolved"])
    await record(`/api/v1/findings?module=${CHANGES}&state=${state}&limit=100`);
  // Every page of each range the change feed offers; the demo works the filters out from them.
  // Page positions are "<time>|<id>", which the replay leaves alone when it moves the dates.
  async function recordPages(path: string) {
    let next: string | null = null;
    for (let page = 0; page < 50; page += 1) {
      const url: string = next ? `${path}&before=${encodeURIComponent(next)}` : path;
      next = (await record<{ next: string | null }>(url)).next;
      if (!next) return;
    }
    throw new Error(`${path} has more than 50 pages.`);
  }
  for (const days of [7, 30, 90, 365]) await recordPages(`/api/v1/changes?days=${days}&limit=200`);
  for (const filter of ["open", "unread", "all", "muted"])
    await record(`/api/v1/attention?filter=${filter}&limit=200`);
  await record("/api/v1/attention?filter=unread&limit=1");
  await record("/api/v1/settings/attention");
  await record("/api/v1/settings/audits");
  await record("/api/v1/settings/audit-schedule");
  await record("/api/v1/settings/audit-review");
  await record("/api/v1/audits/review");
  await record("/api/v1/settings/canaries");
  await record("/api/v1/modules/service.health/status");
  // The This week page, for each range it offers (the demo has no other).
  for (const days of [1, 7, 14, 30]) await record(`/api/v1/digest?days=${days}`);
  await record("/api/v1/ai/status");

  // A handoff from the inactive accounts audit, and a receipt from an invented action tool for
  // some of its accounts, so the Handoffs page and those people's pages have something to show.
  // The export is stored as the page would receive it; the replay builds others from it.
  const inactive = (
    await getJson<{ runs: { id: string; status: string }[] }>(
      "/api/v1/audits/accounts.inactive/runs",
    )
  ).runs.find((run) => run.status === "succeeded");
  if (!inactive) throw new Error("The inactive accounts audit has no result to hand off.");
  const exportUrl = `/api/v1/audit-runs/${inactive.id}/handoff.json?view=d30&sort=days&direction=desc`;
  const exported = await call("GET", exportUrl);
  if (exported.statusCode !== 200)
    throw new Error(
      `The handoff export answered ${exported.statusCode}: ${exported.body.slice(0, 200)}`,
    );
  const handoff = exported.json<{
    id: string;
    items: { userId: string; userPrincipalName: string | null }[];
  }>();
  responses[normalizeKey("GET", exportUrl)] = { status: 200, body: handoff };
  const [released, blocked, kept, refused] = handoff.items;
  if (!released || !blocked || !kept || !refused)
    throw new Error("The inactive accounts audit handed off fewer than four accounts.");
  const reported = (
    account: { userId: string; userPrincipalName: string | null },
    action: string,
    status: "done" | "failed" | "skipped",
    detail: string,
  ) => ({
    userId: account.userId,
    userPrincipalName: account.userPrincipalName,
    action,
    status,
    detail,
  });
  const receipt = {
    format: "prism365.receipt",
    version: 1,
    id: randomUUID(),
    handoffId: handoff.id,
    tool: { name: "Example action tool", version: "1.4.0" },
    performedAt: new Date(anchor.getTime() - 40 * 60_000).toISOString(),
    performedBy: "Jordan Vale",
    items: [
      reported(released, "account.sign-in-blocked", "done", "Sign-in blocked after review."),
      reported(released, "license.removed", "done", "Released the license for reuse."),
      reported(blocked, "account.sign-in-blocked", "done", "Sign-in blocked after review."),
      reported(
        kept,
        "account.sign-in-blocked",
        "skipped",
        "The manager asked to keep it until the end of the month.",
      ),
      reported(
        refused,
        "account.sign-in-blocked",
        "failed",
        "The tool's account lacked the role to change this account.",
      ),
    ],
  };
  const imported = await call("POST", "/api/v1/receipts", receipt);
  if (imported.statusCode !== 201)
    throw new Error(
      `The receipt import answered ${imported.statusCode}: ${imported.body.slice(0, 200)}`,
    );
  await record("/api/v1/handoffs");
  await record(`/api/v1/receipts/${receipt.id}?limit=200`);
  log(`Handed off ${handoff.items.length} inactive accounts and imported a receipt for four.`);

  // A mailbox access export from two days ago, imported as an administrator would import the file
  // Export-MailboxAccess.ps1 wrote (scripts/demo/mailbox-access.ts), so the Microsoft access page
  // and the Access tab's Shared mailboxes section have something to show.
  const mailboxExport = demoMailboxAccessExport(
    tenant,
    new Date(anchor.getTime() - 2 * DAY),
    tenantId,
  );
  const mailboxImport = await call("POST", "/api/v1/mailbox-access/import", mailboxExport);
  if (mailboxImport.statusCode !== 201)
    throw new Error(
      `The mailbox access import answered ${mailboxImport.statusCode}: ${mailboxImport.body.slice(0, 200)}`,
    );
  await record("/api/v1/mailbox-access");
  log(
    `Imported mailbox access for ${mailboxExport.mailboxes.length} shared mailboxes and ${mailboxExport.grants.length} grants.`,
  );

  const audits = await record<{ audits: { definition: { id: string } }[] }>("/api/v1/audits");
  const accounts: { id: string; principal: string }[] = [];
  // The names the Mailbox storage table shows. It carries no account ids, so each name is looked
  // up when clicked, and the lookup has to be on the recording.
  const mailboxNames = new Set<string>();
  // The groups, roles, policies, devices and apps the audit tables list, which open a panel.
  const entities = new Map<string, { kind: string; id: string }>();
  const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  const addEntity = (kind: string, id: unknown) => {
    if (typeof id !== "string" || !GUID.test(id)) return;
    entities.set(`${kind}/${id.toLowerCase()}`, { kind, id: id.toLowerCase() });
  };
  const entityOfAudit: Record<
    string,
    (row: { key: string; cells: Record<string, unknown> }) => void
  > = {
    "groups.teams": (row) => addEntity("group", row.key),
    "directory.privileged-roles": (row) => addEntity("role", row.cells.roleId),
    "identity.conditional-access": (row) => addEntity("policy", row.key),
    "devices.managed": (row) => addEntity("device", row.key),
    "apps.consents": (row) => addEntity("app", row.key),
    "apps.credentials": (row) => addEntity("app", row.cells.servicePrincipalId),
  };
  for (const entry of audits.audits) {
    const id = entry.definition.id;
    await record(`/api/v1/audits/${id}/trend?days=365`);
    const runs = await record<{ runs: { id: string }[] }>(`/api/v1/audits/${id}/runs`);
    for (const run of runs.runs) {
      const detail = await record<{
        rows: {
          key: string;
          personId: string | null;
          cells: { person?: { principal: string | null }; mailbox?: unknown } & Record<
            string,
            unknown
          >;
        }[];
      }>(`/api/v1/audit-runs/${run.id}`);
      for (const row of detail.rows) entityOfAudit[id]?.(row);
      if (id === "mail.mailbox-storage")
        for (const row of detail.rows)
          if (typeof row.cells.mailbox === "string") mailboxNames.add(row.cells.mailbox);
      // The accounts audit lists everyone in the directory, so its newest run is the list of
      // person pages.
      if (id === "accounts.inactive" && run.id === runs.runs[0]?.id)
        for (const row of detail.rows) {
          const principal = row.cells.person?.principal;
          if (row.personId && principal) accounts.push({ id: row.personId, principal });
        }
    }
  }
  if (accounts.length === 0) throw new Error("The accounts audit listed no one.");

  for (const { id, principal } of accounts) {
    await record(`/api/v1/people/${encodeURIComponent(id)}`);
    await record(`/api/v1/people/${encodeURIComponent(id)}/directory`);
    await record(`/api/v1/people/${encodeURIComponent(id)}/access`);
    // The Overview's "Where they show up" panel, read from the stored runs recorded above.
    await record(`/api/v1/people/${encodeURIComponent(id)}/summary`);
    await record(`/api/v1/people/${encodeURIComponent(id)}/mailbox-access`);
    // The first page of all sign-ins; the demo works the "successful" and "not successful"
    // filters out from it. A later page would be asked for by a position that moves with the
    // dates, so every account has to fit on this one.
    const page = await record<{ next: string | null }>(
      `/api/v1/people/${encodeURIComponent(id)}/signins?result=all&limit=100`,
    );
    if (page.next !== null) throw new Error(`${principal} has more sign-ins than one page.`);
    await recordPages(`/api/v1/people/${encodeURIComponent(id)}/changes?limit=20`);
    await record(`/api/v1/people/lookup?email=${encodeURIComponent(principal)}`);
    // Recorded without the principal name: the replay answers the page's request, which names
    // it, from this one, and the demo receipt names every account by its id.
    await record(`/api/v1/people/${encodeURIComponent(id)}/receipts`);
  }
  // What each person's Access tab and Overview name that has a panel too: the groups they are in
  // and own, the roles they hold, the policies that reach them and the apps they own.
  for (const { id } of accounts) {
    const key = (suffix: string) =>
      normalizeKey("GET", `/api/v1/people/${encodeURIComponent(id)}/${suffix}`);
    const access = responses[key("access")]?.body as
      | {
          memberships?: { entries?: { kind: string; id: string; roleTemplateId: string | null }[] };
          owned?: {
            entries?: { kind: string; id: string }[];
            details?: { kind: string; entityId?: string }[];
          };
          conditionalAccess?: { policies?: { id: string }[] };
        }
      | undefined;
    for (const entry of access?.memberships?.entries ?? [])
      if (entry.kind === "role") addEntity("role", entry.roleTemplateId);
      else if (entry.kind !== "administrative-unit" && entry.kind !== "hidden")
        addEntity("group", entry.id);
    for (const entry of access?.owned?.entries ?? [])
      if (entry.kind === "group" || entry.kind === "team") addEntity("group", entry.id);
      else if (entry.kind === "service-principal") addEntity("app", entry.id);
    for (const detail of access?.owned?.details ?? [])
      if (detail.kind === "app") addEntity("app", detail.entityId);
    for (const policy of access?.conditionalAccess?.policies ?? []) addEntity("policy", policy.id);
  }
  // Each panel, and the one search the demo's Quick search example makes.
  for (const { kind, id } of entities.values()) await record(`/api/v1/entities/${kind}/${id}`);
  await record(`/api/v1/entities/search?q=${encodeURIComponent(DEMO_SEARCH)}`);
  log(`Recorded ${entities.size} entity panels.`);
  // A mailbox that is not an account in the accounts audit (a shared mailbox, say) is looked up
  // too, and a name no account has is recorded as the not-found answer the table then shows.
  const looked = new Set(accounts.map((account) => account.principal));
  for (const name of mailboxNames)
    if (!looked.has(name))
      await record(`/api/v1/people/lookup?email=${encodeURIComponent(name)}`, [200, 404]);
  log(`Recorded ${Object.keys(responses).length} responses for ${accounts.length} accounts.`);

  // 6. Check, write and format.
  // The receipt's body as sent, so the replay can tell the same receipt from a different one.
  const fixtures: Fixtures = {
    recordedAt: new Date().toISOString(),
    responses,
    receipts: { [receipt.id]: canonicalJson(receipt) },
  };
  const text = JSON.stringify(fixtures);
  assertSynthetic(text, tenant.addresses);
  await mkdir(dirname(output), { recursive: true });
  await writeFile(output, text);
  log(`Wrote ${output} (${(text.length / 1_048_576).toFixed(1)} MB).`);

  await app.close();
  await rm(dataDir, { recursive: true, force: true });
}

interface Finding {
  id: string;
  ruleId: string;
  severity: string;
  summary: string;
  subject: { id: string };
}

main().then(
  () => process.exit(0),
  (error: unknown) => {
    console.error(error);
    process.exit(1);
  },
);

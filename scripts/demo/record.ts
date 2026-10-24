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
  const app = createApp({
    config,
    store,
    identity,
    graphFetch: ((input, init) => auditGraph(input, init)) as typeof fetch,
    audits: auditRunner,
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

  // 1. Collect. The first tick loads the history; the second reads what arrived since.
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
  };
  for (const [name, finding] of Object.entries(expected))
    if (!finding) throw new Error(`The rules did not flag the planted story "${name}".`);
  if (expected.travel?.severity !== "critical")
    throw new Error(`Impossible travel was rated ${expected.travel?.severity}, not critical.`);
  log(`Rules flagged ${findings.length} findings, including all six planted stories.`);
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

  // 4. Run every audit: five weeks ago, two weeks ago, a week ago and now, so each has a trend and
  // the This week page has a result from before even its longest range to compare with.
  const catalog = await getJson<{ audits: { definition: { id: string } }[] }>("/api/v1/audits");
  for (const daysAgo of [35, 14, 7, 0]) {
    auditClock = daysAgo === 0 ? null : new Date(anchor.getTime() - daysAgo * DAY);
    auditGraph = auditClock ? createSyntheticGraph({ now: auditClock, quietFrom: anchor }) : graph;
    for (const entry of catalog.audits) {
      const id = entry.definition.id;
      const response = await call("POST", `/api/v1/audits/${id}/runs`);
      if (response.statusCode !== 202)
        throw new Error(`Audit ${id} did not start: ${response.body}`);
      const runId = (response.json() as { run: { id: string } }).run.id;
      const read = () =>
        getJson<{ status: string; rowCount: number; error: unknown }>(
          `/api/v1/audit-runs/${runId}`,
        );
      let detail = await read();
      for (let poll = 0; detail.status === "running" && poll < 400; poll += 1) {
        await new Promise((done) => setTimeout(done, 100));
        detail = await read();
      }
      if (detail.status !== "succeeded")
        throw new Error(`Audit ${id} ${detail.status}: ${JSON.stringify(detail.error)}`);
      log(`Audit ${id}${daysAgo ? ` (${daysAgo} days ago)` : ""}: ${detail.rowCount} rows.`);
    }
  }
  auditClock = null;
  auditGraph = graph;
  // Two audits run on their own. Saved after the last runs, so neither is due while recording.
  const scheduled = await call("POST", "/api/v1/settings/audit-schedule", {
    audits: { "accounts.inactive": "weekly", "mfa.registration": "daily" },
  });
  if (scheduled.statusCode !== 200)
    throw new Error(`The audit schedule was not saved: ${scheduled.body}`);

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
  await record("/api/v1/modules/service.health/status");
  // The This week page, for each range it offers (the demo has no other).
  for (const days of [1, 7, 14, 30]) await record(`/api/v1/digest?days=${days}`);
  await record("/api/v1/automation/workflows");
  await record("/api/v1/automation/packages");
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

  const audits = await record<{ audits: { definition: { id: string } }[] }>("/api/v1/audits");
  const accounts: { id: string; principal: string }[] = [];
  for (const entry of audits.audits) {
    const id = entry.definition.id;
    await record(`/api/v1/audits/${id}/trend?days=365`);
    const runs = await record<{ runs: { id: string }[] }>(`/api/v1/audits/${id}/runs`);
    for (const run of runs.runs) {
      const detail = await record<{
        rows: { personId: string | null; cells: { person?: { principal: string | null } } }[];
      }>(`/api/v1/audit-runs/${run.id}`);
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
    // The first page of all sign-ins; the demo works the "successful" and "not successful"
    // filters out from it. A later page would be asked for by a position that moves with the
    // dates, so every account has to fit on this one.
    const page = await record<{ next: string | null }>(
      `/api/v1/people/${encodeURIComponent(id)}/signins?result=all&limit=100`,
    );
    if (page.next !== null) throw new Error(`${principal} has more sign-ins than one page.`);
    await recordPages(`/api/v1/people/${encodeURIComponent(id)}/changes?limit=20`);
    await record(`/api/v1/people/lookup?email=${encodeURIComponent(principal)}`);
    await record(`/api/v1/investigations/user?email=${encodeURIComponent(principal)}`);
    // Recorded without the principal name: the replay answers the page's request, which names
    // it, from this one, and the demo receipt names every account by its id.
    await record(`/api/v1/people/${encodeURIComponent(id)}/receipts`);
  }
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

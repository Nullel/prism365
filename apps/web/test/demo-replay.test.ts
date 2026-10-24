import {
  type AttentionItem,
  type AuditRow,
  auditCsv,
  auditDefinitionSchema,
  queryAuditRows,
} from "@prism365/contracts";
import { describe, expect, it, vi } from "vitest";
import {
  assertSynthetic,
  buildSyntheticTenant,
  createSyntheticGraph,
} from "../../../scripts/demo/tenant.ts";
import {
  attentionCounts,
  createDemoState,
  type DemoState,
  type Fixtures,
  handleRequest,
  installReplay,
  matchFixture,
  normalizeKey,
  SIMULATED_RUN_MS,
  shiftString,
  shiftTimestamps,
  unavailableBody,
  wholeDaysSince,
} from "../src/demo/replay.ts";

// Invented data only: a tiny recording shaped like the real one.
const recordedAt = "2026-10-20T10:00:00.000Z";
const now = (iso: string) => () => new Date(iso);
const byteOrderMark = String.fromCharCode(0xfeff);

const definition = auditDefinitionSchema.parse({
  id: "accounts.inactive",
  version: 1,
  title: "Inactive accounts",
  description: "Every account and how long it has gone without signing in.",
  category: "Accounts",
  subject: { one: "account", many: "accounts" },
  columns: [
    { key: "person", label: "Account", kind: "person" },
    { key: "days", label: "Inactive for", kind: "days" },
    { key: "status", label: "Status", kind: "enum", facet: true },
    { key: "lastSuccess", label: "Last successful sign-in", kind: "date" },
  ],
  views: [
    { id: "all", label: "All accounts", tags: [] },
    { id: "d30", label: "30+ days", tags: ["d30"] },
  ],
  defaultView: "all",
  defaultSort: { column: "days", direction: "desc" },
  permissions: ["User.Read.All"],
  optionalPermissions: [],
  requirements: [],
});

const rows: AuditRow[] = [
  {
    key: "u-1",
    personId: "u-1",
    tags: ["d30", "enabled"],
    cells: {
      person: { name: "Avery Example", principal: "avery@larkspur.example" },
      days: 41,
      status: "Enabled",
      lastSuccess: "2026-09-09T08:30:00.000Z",
    },
  },
  {
    key: "u-2",
    personId: "u-2",
    tags: ["disabled"],
    cells: {
      person: { name: "Blake, Example", principal: "blake@larkspur.example" },
      days: 3,
      status: "Disabled",
      lastSuccess: "2026-10-17T08:30:00.000Z",
    },
  },
  {
    key: "u-3",
    personId: "u-3",
    tags: ["d30", "enabled"],
    cells: {
      person: { name: "Casey Example", principal: "casey@larkspur.example" },
      days: 90,
      status: "Enabled",
      lastSuccess: null,
    },
  },
];

const recordedRun = {
  id: "11111111-1111-4111-8111-111111111111",
  auditId: "accounts.inactive",
  auditVersion: 1,
  status: "succeeded",
  complete: true,
  startedAt: "2026-10-20T09:59:00.000Z",
  finishedAt: "2026-10-20T09:59:03.000Z",
  startedBy: { id: "22222222-2222-4222-8222-222222222222", name: "Demo administrator" },
  error: null,
  summary: [{ label: "Accounts", value: 3, tone: "neutral", view: "all" }],
  notes: ["A note."],
  rowCount: 3,
};

function item(
  overrides: Partial<AttentionItem> & Pick<AttentionItem, "id" | "ref">,
): AttentionItem {
  return {
    module: "service.health",
    kind: "service.incident",
    class: "incident",
    title: "Some users cannot send mail",
    summary: "Users may be unable to send mail.",
    services: ["Exchange Online"],
    provider: {
      status: "Investigating",
      labels: ["Incident"],
      deadline: null,
      updatedAt: "2026-10-20T09:00:00.000Z",
      active: true,
      impact: null,
      workaround: null,
    },
    localLabels: [],
    severity: "high",
    link: null,
    revision: 1,
    changeNote: null,
    firstSeen: "2026-10-20T08:00:00.000Z",
    lastSeen: "2026-10-20T09:00:00.000Z",
    changedAt: "2026-10-20T08:00:00.000Z",
    unread: true,
    review: { state: "new", changedAt: null },
    muted: false,
    ...overrides,
  };
}

const incident = item({ id: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa", ref: "EX000001" });
const action = item({
  id: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
  ref: "MC000002",
  module: "service.messages",
  kind: "service.message",
  class: "message.action",
  severity: "medium",
  provider: {
    status: null,
    labels: ["Action required"],
    deadline: "2026-10-30T00:00:00.000Z",
    updatedAt: "2026-10-19T09:00:00.000Z",
    active: true,
    impact: null,
    workaround: null,
  },
});
const change = item({
  id: "cccccccc-cccc-4ccc-8ccc-cccccccccccc",
  ref: "MC000003",
  class: "message.change",
  severity: "low",
  unread: false,
});

const person = {
  history: null,
  findings: [
    {
      id: "dddddddd-dddd-4ddd-8ddd-dddddddddddd",
      module: "signins.audit",
      ruleId: "signins.mfa_denials",
      ruleVersion: 1,
      purpose: "security",
      severity: "medium",
      score: 50,
      subject: { kind: "person", id: "u-1", label: "Avery Example" },
      title: "Repeated failed MFA prompts for one account",
      summary: "5 failed MFA prompts.",
      evidence: [],
      limitations: [],
      windowStart: "2026-10-18T10:00:00.000Z",
      windowEnd: "2026-10-18T10:10:00.000Z",
      state: "open",
      occurrences: 1,
      firstSeen: "2026-10-18T10:10:00.000Z",
      lastSeen: "2026-10-18T10:10:00.000Z",
      stateChangedAt: null,
    },
  ],
};

const signinPage = {
  items: [
    { id: "s-1", result: "success", errorCode: 0 },
    { id: "s-2", result: "failure", errorCode: 50126 },
    { id: "s-3", result: "interrupted", errorCode: 50074 },
  ],
  next: null,
};

const fixtures: Fixtures = {
  recordedAt,
  responses: {
    "GET /api/v1/session": {
      status: 200,
      body: { mode: "live", displayName: "Demo administrator" },
    },
    "GET /api/v1/ai/status": { status: 200, body: { available: false } },
    "GET /api/v1/dashboard/summary": { status: 409, body: { error: "dashboard_not_collected" } },
    "GET /api/v1/attention?filter=all&limit=200": {
      status: 200,
      body: {
        mode: "live",
        items: [incident, action, change],
        system: [],
        sources: [
          { module: "service.health", title: "Service health", state: "available", missing: [] },
        ],
        counts: { unread: 2, active: 3 },
      },
    },
    "GET /api/v1/settings/attention": {
      status: 200,
      body: {
        preferences: {
          services: [],
          notify: {
            incident: true,
            advisory: true,
            "message.action": true,
            "message.retirement": true,
            "message.change": false,
          },
        },
        services: ["Exchange Online"],
      },
    },
    "GET /api/v1/modules": {
      status: 200,
      body: {
        mode: "live",
        collection: {
          available: true,
          state: "collecting",
          credentialKind: "delegated",
          boundAt: "2026-10-20T07:00:00.000Z",
          lastStopReason: null,
          durableStorage: false,
          reason: null,
          resumes: false,
        },
        modules: [],
      },
    },
    "GET /api/v1/findings?limit=100&module=signins.audit&state=open": {
      status: 200,
      body: { mode: "live", items: person.findings },
    },
    "GET /api/v1/findings?limit=100&module=signins.audit&state=acknowledged": {
      status: 200,
      body: { mode: "live", items: [] },
    },
    "GET /api/v1/people/u-1": { status: 200, body: person },
    "GET /api/v1/people/u-1/signins?limit=100&result=all": { status: 200, body: signinPage },
    "GET /api/v1/audits": {
      status: 200,
      body: {
        mode: "live",
        audits: [
          {
            definition,
            availability: { state: "available", missing: [], grantFeature: null, enable: [] },
            lastRun: recordedRun,
            runningId: null,
          },
        ],
      },
    },
    "GET /api/v1/audits/accounts.inactive/runs": { status: 200, body: { runs: [recordedRun] } },
    [`GET /api/v1/audit-runs/${recordedRun.id}`]: { status: 200, body: { ...recordedRun, rows } },
  },
};

// "Now" is three days and two hours after the recording.
const later = "2026-10-23T12:00:00.000Z";
const start = (clock = later) => {
  const state = createDemoState(fixtures, now(clock));
  const request = (method: string, url: string, body?: unknown) =>
    handleRequest(state, { method, url, ...(body === undefined ? {} : { body }) });
  return { state, request };
};

describe("keys and fallback", () => {
  it("sorts query parameters and ignores their order and the host", () => {
    const key = "GET /api/v1/people?days=30&search=al";
    expect(normalizeKey("get", "/api/v1/people?search=al&days=30")).toBe(key);
    expect(normalizeKey("GET", "https://demo.example/api/v1/people?days=30&search=al")).toBe(key);
    expect(normalizeKey("POST", "/api/v1/collection/start")).toBe("POST /api/v1/collection/start");
    expect(normalizeKey("GET", "/api/v1/people/lookup?email=a%40b.example")).toBe(
      "GET /api/v1/people/lookup?email=a%40b.example",
    );
  });

  it("answers an exact key, then the same path without its query, then says it is not available", () => {
    const responses = {
      "GET /api/v1/ai/status": { status: 200, body: { available: false } },
      "GET /api/v1/people?days=30": { status: 200, body: { days: 30 } },
    };
    expect(matchFixture(responses, "GET", "/api/v1/people?days=30")?.body).toEqual({ days: 30 });
    expect(matchFixture(responses, "GET", "/api/v1/ai/status?cache=no")?.body).toEqual({
      available: false,
    });
    expect(matchFixture(responses, "GET", "/api/v1/people?days=7")).toBeNull();

    const { request } = start();
    expect(request("GET", "/api/v1/ai/status?x=1")).toMatchObject({ status: 200 });
    expect(request("GET", "/api/v1/something/new")).toMatchObject({
      status: 404,
      body: unavailableBody,
    });
    expect(request("POST", "/api/v1/ai/grants", {})).toMatchObject({
      status: 404,
      body: unavailableBody,
    });
    expect(request("GET", "/somewhere/else")).toMatchObject({ status: 404 });
    expect(unavailableBody).toEqual({
      error: "demo",
      message: "This is not available in the demo.",
    });
    // Recorded errors are replayed as recorded.
    expect(request("GET", "/api/v1/dashboard/summary")).toMatchObject({ status: 409 });
  });

  it("works out the other sign-in filters from the page of all sign-ins", () => {
    const { request } = start();
    const ids = (url: string) =>
      (request("GET", url).body as { items: { id: string }[] }).items.map((entry) => entry.id);
    expect(ids("/api/v1/people/u-1/signins?result=all&limit=100")).toEqual(["s-1", "s-2", "s-3"]);
    expect(ids("/api/v1/people/u-1/signins?limit=100&result=success")).toEqual(["s-1"]);
    expect(ids("/api/v1/people/u-1/signins?limit=100&result=failed")).toEqual(["s-2", "s-3"]);
    expect(
      request("GET", "/api/v1/people/u-1/signins?result=all&limit=100&before=x|y").status,
    ).toBe(404);
  });
});

describe("moving the recording to today", () => {
  it("counts whole days and never goes negative", () => {
    expect(wholeDaysSince(recordedAt, new Date("2026-10-23T09:59:59.000Z"))).toBe(2);
    expect(wholeDaysSince(recordedAt, new Date("2026-10-23T10:00:00.000Z"))).toBe(3);
    expect(wholeDaysSince(recordedAt, new Date("2026-10-20T11:00:00.000Z"))).toBe(0);
    expect(wholeDaysSince(recordedAt, new Date("2026-10-01T00:00:00.000Z"))).toBe(0);
    expect(wholeDaysSince("not a date", new Date())).toBe(0);
  });

  it("shifts timestamps and dates, keeping the time of day and the offset", () => {
    expect(shiftString("2026-10-20T10:00:00.000Z", 3)).toBe("2026-10-23T10:00:00.000Z");
    expect(shiftString("2026-10-30T23:30:00+02:00", 3)).toBe("2026-11-02T23:30:00+02:00");
    expect(shiftString("2026-10-30T23:30:00Z", 0)).toBe("2026-10-30T23:30:00Z");
    expect(shiftString("2026-02-27", 2)).toBe("2026-03-01");
    expect(shiftString("2028-02-27", 2)).toBe("2028-02-29");
    // Anything that is not a whole timestamp or date stays as it is.
    for (const other of [
      "EX900417",
      "11111111-1111-4111-8111-111111111111",
      "Last seen 2026-10-20",
      "2026-10-20T10:00:00.000Z|abc",
      "",
    ])
      expect(shiftString(other, 3)).toBe(other);
  });

  it("shifts every nested value and leaves keys, numbers and the original alone", () => {
    const original = {
      "2026-10-20": "key",
      at: "2026-10-20T10:00:00.000Z",
      list: [{ day: "2026-10-20", n: 5, flag: true, none: null }],
    };
    const shifted = shiftTimestamps(original, 2);
    expect(shifted).toEqual({
      "2026-10-20": "key",
      at: "2026-10-22T10:00:00.000Z",
      list: [{ day: "2026-10-22", n: 5, flag: true, none: null }],
    });
    expect(original.at).toBe("2026-10-20T10:00:00.000Z");
  });

  it("serves the recording already shifted", () => {
    const { request } = start();
    const run = request("GET", `/api/v1/audit-runs/${recordedRun.id}`).body as {
      startedAt: string;
      rows: { cells: { lastSuccess: string | null } }[];
    };
    expect(run.startedAt).toBe("2026-10-23T09:59:00.000Z");
    expect(run.rows[0]?.cells.lastSuccess).toBe("2026-09-12T08:30:00.000Z");
    const modules = request("GET", "/api/v1/modules").body as { collection: { boundAt: string } };
    expect(modules.collection.boundAt).toBe("2026-10-23T07:00:00.000Z");
  });
});

describe("notifications", () => {
  const list = (request: ReturnType<typeof start>["request"], filter: string) =>
    request("GET", `/api/v1/attention?filter=${filter}&limit=200`).body as {
      items: AttentionItem[];
      counts: { unread: number; active: number };
    };

  it("lists by filter, most serious and unread first, with counts", () => {
    const { request } = start();
    expect(list(request, "all").items.map((entry) => entry.ref)).toEqual([
      "EX000001",
      "MC000002",
      "MC000003",
    ]);
    expect(list(request, "unread").items.map((entry) => entry.ref)).toEqual([
      "EX000001",
      "MC000002",
    ]);
    expect(list(request, "muted").items).toEqual([]);
    expect(list(request, "open").counts).toEqual({ unread: 2, active: 3 });
    expect(request("GET", "/api/v1/attention?filter=nonsense").status).toBe(400);
  });

  it("marks an item reviewed: read, Microsoft's words untouched, counts and filters follow", () => {
    const { request } = start();
    const reviewed = request("POST", `/api/v1/attention/${incident.id}/review`, {
      state: "reviewed",
    });
    expect(reviewed.status).toBe(200);
    expect(reviewed.body).toMatchObject({
      id: incident.id,
      unread: false,
      review: { state: "reviewed", changedAt: later },
      provider: { status: "Investigating", active: true },
    });
    expect(list(request, "unread").items.map((entry) => entry.ref)).toEqual(["MC000002"]);
    expect(list(request, "unread").counts.unread).toBe(1);

    request("POST", `/api/v1/attention/${incident.id}/review`, { state: "done" });
    expect(list(request, "open").items.map((entry) => entry.ref)).toEqual(["MC000002", "MC000003"]);
    expect(list(request, "open").counts.active).toBe(2);

    // Reopening does not make it unread again.
    request("POST", `/api/v1/attention/${incident.id}/review`, { state: "new" });
    expect(list(request, "unread").items.map((entry) => entry.ref)).toEqual(["MC000002"]);
  });

  it("rejects a bad review and an unknown item", () => {
    const { request } = start();
    expect(
      request("POST", `/api/v1/attention/${incident.id}/review`, { state: "archived" }).status,
    ).toBe(400);
    expect(
      request("POST", "/api/v1/attention/eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee/review", {
        state: "done",
      }),
    ).toMatchObject({
      status: 404,
    });
  });

  it("mutes, reads one and reads all", () => {
    const { request } = start();
    expect(
      request("POST", `/api/v1/attention/${action.id}/mute`, { muted: true }).body,
    ).toMatchObject({
      muted: true,
      unread: false,
    });
    expect(list(request, "muted").items.map((entry) => entry.ref)).toEqual(["MC000002"]);
    expect(list(request, "open").items.map((entry) => entry.ref)).toEqual(["EX000001", "MC000003"]);
    expect(request("POST", `/api/v1/attention/${action.id}/mute`, { muted: "yes" }).status).toBe(
      400,
    );

    expect(request("POST", `/api/v1/attention/${incident.id}/read`).body).toEqual({ updated: 1 });
    expect(request("POST", `/api/v1/attention/${incident.id}/read`).body).toEqual({ updated: 0 });
    expect(request("POST", "/api/v1/attention/read-all").body).toEqual({ updated: 0 });
    expect(attentionCounts(start().state.attention)).toEqual({ unread: 2, active: 3 });
  });

  it("adds the alert the server raises when collection is stopped, and saves settings", () => {
    const { request } = start();
    request("POST", "/api/v1/collection/stop");
    expect(list(request, "open")).toMatchObject({ mode: "live" });
    const alerts = (
      request("GET", "/api/v1/attention?filter=open&limit=200").body as {
        system: { id: string }[];
      }
    ).system;
    expect(alerts.map((alert) => alert.id)).toEqual(["collection.stopped"]);
    request("POST", "/api/v1/collection/start");
    expect(
      (request("GET", "/api/v1/attention?filter=open&limit=200").body as { system: unknown[] })
        .system,
    ).toEqual([]);

    const saved = request("POST", "/api/v1/settings/attention", {
      services: ["Exchange Online"],
      notify: { incident: true, advisory: false },
    });
    expect(saved.body).toMatchObject({
      preferences: {
        services: ["Exchange Online"],
        notify: { advisory: false, "message.action": true },
      },
      services: ["Exchange Online"],
    });
    expect(request("GET", "/api/v1/settings/attention").body).toEqual(saved.body);
    expect(request("POST", "/api/v1/settings/attention", { services: "Exchange" }).status).toBe(
      400,
    );
  });
});

describe("collection and findings", () => {
  it("starts and stops collection in the modules response", () => {
    const { request } = start();
    const collection = () =>
      (request("GET", "/api/v1/modules").body as { collection: Record<string, unknown> })
        .collection;
    expect(collection()).toMatchObject({ state: "collecting", credentialKind: "delegated" });

    const stopped = request("POST", "/api/v1/collection/stop");
    expect(stopped.body).toMatchObject({ state: "stopped", credentialKind: null, boundAt: null });
    expect(collection()).toMatchObject({ state: "stopped", lastStopReason: "operator_stopped" });

    request("POST", "/api/v1/collection/start");
    expect(collection()).toMatchObject({
      state: "collecting",
      credentialKind: "delegated",
      boundAt: later,
    });
  });

  it("changes a finding's state everywhere it is shown", () => {
    const { request } = start();
    const id = person.findings[0]?.id ?? "";
    const titles = (url: string) =>
      (request("GET", url).body as { items: { id: string }[] }).items.map((entry) => entry.id);
    expect(titles("/api/v1/findings?module=signins.audit&state=open&limit=100")).toEqual([id]);

    const changed = request("POST", `/api/v1/findings/${id}/state`, { state: "acknowledged" });
    expect(changed.body).toMatchObject({ id, state: "acknowledged", stateChangedAt: later });
    expect(titles("/api/v1/findings?module=signins.audit&state=open&limit=100")).toEqual([]);
    expect(titles("/api/v1/findings?module=signins.audit&state=acknowledged&limit=100")).toEqual([
      id,
    ]);
    expect(
      (request("GET", "/api/v1/people/u-1").body as { findings: { state: string }[] }).findings[0]
        ?.state,
    ).toBe("acknowledged");

    expect(request("POST", `/api/v1/findings/${id}/state`, { state: "archived" }).status).toBe(400);
    expect(
      request("POST", "/api/v1/findings/eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee/state", {
        state: "open",
      }).status,
    ).toBe(404);
  });

  it("makes sign-out a no-op success", () => {
    expect(start().request("POST", "/api/v1/auth/logout")).toMatchObject({
      status: 200,
      body: { signedOut: true },
    });
  });
});

describe("running an audit again", () => {
  it("shows a running copy, then the recorded result two seconds later", () => {
    let clock = Date.parse(later);
    const state: DemoState = createDemoState(fixtures, () => new Date(clock));
    const request = (method: string, url: string) => handleRequest(state, { method, url });

    const started = request("POST", "/api/v1/audits/accounts.inactive/runs");
    expect(started.status).toBe(202);
    const run = (
      started.body as { run: { id: string; status: string; startedAt: string }; started: boolean }
    ).run;
    expect((started.body as { started: boolean }).started).toBe(true);
    expect(run.id).not.toBe(recordedRun.id);
    expect(run).toMatchObject({ status: "running", startedAt: later });

    const catalogEntry = () =>
      (
        request("GET", "/api/v1/audits").body as {
          audits: { lastRun: { id: string; status: string }; runningId: string | null }[];
        }
      ).audits[0];
    expect(catalogEntry()).toMatchObject({
      runningId: run.id,
      lastRun: { id: run.id, status: "running" },
    });
    expect(request("GET", `/api/v1/audit-runs/${run.id}`).body).toMatchObject({
      status: "running",
      rows: [],
    });
    // A second request joins the run in progress.
    const joined = request("POST", "/api/v1/audits/accounts.inactive/runs");
    expect(joined).toMatchObject({ status: 200, body: { started: false, run: { id: run.id } } });

    clock += SIMULATED_RUN_MS;
    const finished = request("GET", `/api/v1/audit-runs/${run.id}`).body as {
      status: string;
      finishedAt: string;
      startedAt: string;
      rowCount: number;
      rows: AuditRow[];
      summary: unknown[];
    };
    expect(finished).toMatchObject({
      status: "succeeded",
      startedAt: later,
      finishedAt: "2026-10-23T12:00:02.000Z",
      rowCount: 3,
    });
    expect(finished.rows.map((row) => row.key)).toEqual(["u-1", "u-2", "u-3"]);
    expect(finished.summary).toHaveLength(1);

    expect(catalogEntry()).toMatchObject({
      runningId: null,
      lastRun: { id: run.id, status: "succeeded" },
    });
    const history = (
      request("GET", "/api/v1/audits/accounts.inactive/runs").body as { runs: { id: string }[] }
    ).runs;
    expect(history.map((entry) => entry.id)).toEqual([run.id, recordedRun.id]);

    // Another run after the first has finished is a new one, listed first.
    const second = request("POST", "/api/v1/audits/accounts.inactive/runs");
    expect((second.body as { started: boolean }).started).toBe(true);
    expect(
      (request("GET", "/api/v1/audits/accounts.inactive/runs").body as { runs: unknown[] }).runs,
    ).toHaveLength(3);
    expect(request("POST", "/api/v1/audits/not.an-audit/runs").status).toBe(404);
  });
});

describe("CSV export", () => {
  // The rows as the demo serves them: three days later than they were recorded.
  const shifted = shiftTimestamps(rows, 3);
  const url = (id: string, query = "") => `/api/v1/audit-runs/${id}/export.csv${query}`;

  it("builds the file the server would, from the page's view, filters and sort", () => {
    const { request } = start();
    const query = "?view=d30&sort=days&direction=asc&f.status=Enabled&search=example";
    const response = request("GET", url(recordedRun.id, query));
    expect(response.status).toBe(200);
    expect(response.contentType).toBe("text/csv; charset=utf-8");
    expect(response.headers?.["Content-Disposition"]).toBe(
      'attachment; filename="prism365-accounts.inactive-2026-10-23.csv"',
    );
    const expected = auditCsv(
      definition,
      queryAuditRows(definition, shifted, {
        view: "d30",
        facets: { status: "Enabled" },
        search: "example",
        sort: { column: "days", direction: "asc" },
      }),
    );
    expect(response.body).toBe(expected);
    const text = response.body as string;
    expect(text.startsWith(byteOrderMark)).toBe(true);
    expect(text.split("\r\n").filter(Boolean)).toHaveLength(3);
    expect(text.indexOf("Avery Example")).toBeLessThan(text.indexOf("Casey Example"));
    expect(text).not.toContain("Blake");
  });

  it("exports everything in the default order without a query", () => {
    const { request } = start();
    const text = request("GET", url(recordedRun.id)).body as string;
    expect(text).toBe(auditCsv(definition, queryAuditRows(definition, shifted, { view: null })));
    expect(text.split("\r\n").filter(Boolean)).toHaveLength(4);
    // A name with a comma is quoted.
    expect(text).toContain('"Blake, Example"');
  });

  it("exports a run started in the demo once it has finished, and refuses it before", () => {
    let clock = Date.parse(later);
    const state = createDemoState(fixtures, () => new Date(clock));
    const request = (method: string, target: string) =>
      handleRequest(state, { method, url: target });
    const id = (
      request("POST", "/api/v1/audits/accounts.inactive/runs").body as { run: { id: string } }
    ).run.id;
    expect(request("GET", url(id)).status).toBe(404);
    clock += SIMULATED_RUN_MS;
    const finished = request("GET", url(id));
    expect(finished.status).toBe(200);
    expect(finished.headers?.["Content-Disposition"]).toContain(
      "prism365-accounts.inactive-2026-10-23",
    );
    expect(request("GET", url("ffffffff-ffff-4fff-8fff-ffffffffffff")).status).toBe(404);
  });
});

describe("installing the shim", () => {
  const target = () => {
    const real = vi.fn(async () => new Response("from the network", { status: 200 }));
    return {
      real,
      target: {
        fetch: real as unknown as typeof fetch,
        location: { origin: "https://demo.example" },
      },
    };
  };

  it("answers /api/v1 from the recording and sends everything else on", async () => {
    const { real, target: page } = target();
    const replay = installReplay(fixtures, { target: page, now: now(later) });

    const session = await page.fetch("/api/v1/session", { cache: "no-store" });
    expect(session.status).toBe(200);
    expect(await session.json()).toMatchObject({ displayName: "Demo administrator" });
    const missing = await page.fetch("https://demo.example/api/v1/nothing-here");
    expect(missing.status).toBe(404);
    expect(await missing.json()).toEqual(unavailableBody);
    const reviewed = await page.fetch(`/api/v1/attention/${incident.id}/review`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ state: "done" }),
    });
    expect((await reviewed.json()).review.state).toBe("done");
    const csv = await page.fetch(`/api/v1/audit-runs/${recordedRun.id}/export.csv?view=all`);
    expect(csv.headers.get("Content-Type")).toBe("text/csv; charset=utf-8");
    // Reading a response as text drops the byte-order mark, so look at the bytes.
    expect([...new Uint8Array(await csv.arrayBuffer()).slice(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);

    expect(real).not.toHaveBeenCalled();
    await page.fetch("/assets/app.js");
    await page.fetch("https://elsewhere.example/api/v1/session");
    expect(real).toHaveBeenCalledTimes(2);

    const controller = new AbortController();
    controller.abort();
    await expect(
      page.fetch("/api/v1/session", { signal: controller.signal }),
    ).rejects.toMatchObject({
      name: "AbortError",
    });

    replay.restore();
    expect(page.fetch).toBe(real);
  });
});

describe("the synthetic tenant behind the recording", () => {
  const anchor = new Date("2026-10-22T18:00:00.000Z");
  const day = 86_400_000;

  it("has the same people whenever it is built, with only the dates moving", () => {
    const first = buildSyntheticTenant(anchor);
    const later = buildSyntheticTenant(new Date(anchor.getTime() + 5 * day + 3_600_000));
    const people = (tenant: typeof first) =>
      tenant.users.map((user) => [
        user.id,
        user.userPrincipalName,
        user.displayName,
        user.userType,
      ]);
    expect(first.users).toHaveLength(53);
    expect(people(later)).toEqual(people(first));
    // Two invented domains in one tenant: the company and its research subsidiary.
    expect(
      first.users.every((user) => /@larkspur(-research)?\.example$/.test(user.userPrincipalName)),
    ).toBe(true);
    expect(
      first.users.some((user) => user.userPrincipalName.endsWith("@larkspur-research.example")),
    ).toBe(true);
    // Nobody has more sign-ins than fit on one page of their own.
    const counts = new Map<string, number>();
    for (const event of first.signIns)
      counts.set(event.userId, (counts.get(event.userId) ?? 0) + 1);
    expect(Math.max(...counts.values())).toBeLessThan(100);
    expect(Date.parse(first.signIns.at(-1)?.createdDateTime ?? "")).toBeLessThan(anchor.getTime());
  });

  it("answers the Microsoft Graph requests the server makes", async () => {
    const graph = createSyntheticGraph({ now: anchor });
    const get = async (path: string) => {
      const response = await graph(new URL(`https://graph.microsoft.com/v1.0/${path}`));
      return { status: response.status, body: (await response.json()) as Record<string, unknown> };
    };
    const tenant = buildSyntheticTenant(anchor);
    const traveler = tenant.stories.traveler ?? "";

    const filter = encodeURIComponent(
      `createdDateTime ge ${new Date(anchor.getTime() - 10 * day).toISOString()} and createdDateTime le ${anchor.toISOString()}`,
    );
    const page = await get(`auditLogs/signIns?$filter=${filter}&$top=5`);
    const times = (page.body.value as { createdDateTime: string }[]).map(
      (item) => item.createdDateTime,
    );
    expect(times).toHaveLength(5);
    expect([...times].sort().reverse()).toEqual(times);
    const next = String(page.body["@odata.nextLink"]);
    expect(next).toContain("graph.microsoft.com/v1.0/auditLogs/signIns");
    const second = await get(next.replace("https://graph.microsoft.com/v1.0/", ""));
    expect((second.body.value as unknown[]).length).toBe(5);
    expect(second.body.value).not.toEqual(page.body.value);

    const mine = await get(
      `auditLogs/signIns?$filter=${encodeURIComponent(`userId eq '${traveler}'`)}&$top=100`,
    );
    expect(
      (mine.body.value as { userId: string }[]).every((item) => item.userId === traveler),
    ).toBe(true);
    expect((mine.body.value as unknown[]).length).toBeGreaterThan(10);

    const byName = await get(
      "users('admin%40larkspur.example')?$select=displayName,accountEnabled",
    );
    expect(byName.body).toMatchObject({ displayName: "Demo administrator", accountEnabled: true });
    expect(Object.keys(byName.body).sort()).toEqual(["accountEnabled", "displayName", "id"]);
    expect((await get(`users/${traveler}/manager`)).status).toBe(200);
    expect((await get("users/00000000-0000-4000-8000-000000000000")).status).toBe(404);
    expect(((await get("subscribedSkus")).body.value as unknown[]).length).toBe(5);
    expect(
      (await get("admin/serviceAnnouncement/issues?$filter=isResolved eq false")).body.value,
    ).toHaveLength(1);
    expect((await get("security/incidents")).status).toBe(404);
    const elsewhere = await graph(new URL("https://example.org/v1.0/users"));
    expect(elsewhere.status).toBe(404);
  });

  it("refuses a recording that mentions anything outside the invented company", () => {
    const known = new Set(["admin@larkspur.example"]);
    const clean = JSON.stringify({
      mail: "admin@larkspur.example",
      guest: "ada@harborline.example",
      ip: ["192.0.2.14", "2001:db8:4f1a::9c"],
      link: "https://admin.microsoft.com/#/servicehealth/:/alerts/EX900417",
      at: "2026-10-22T10:00:00.000Z",
    });
    expect(() => assertSynthetic(clean, known)).not.toThrow();
    for (const bad of [
      "someone@gmail.com",
      "ada.lovelace@larkspur.example",
      "https://tracker.example.org/a",
      "8.8.8.8",
      "2606:4700:4700::1111",
    ])
      expect(() => assertSynthetic(JSON.stringify({ value: bad }), known), bad).toThrow(
        /outside the synthetic tenant/,
      );
  });
});

describe("directory changes", () => {
  const change = (id: string, activityAt: string, extra: Record<string, unknown> = {}) => ({
    id,
    activityAt,
    activity: "Update user",
    category: "UserManagement",
    operationType: "Update",
    result: "success",
    resultReason: null,
    loggedByService: "Core Directory",
    correlationId: null,
    initiatedBy: { kind: "user", id: "u-9", name: "Avery Example", principal: null },
    targets: [{ id: "u-1", name: "Blake Example", type: "User", principal: null, properties: [] }],
    details: [],
    ...extra,
  });
  const base = {
    days: 7,
    categories: [],
    coverage: { from: null, through: null, lastCollected: null, phase: "live" },
  };
  const changes: Fixtures = {
    recordedAt,
    responses: {
      [normalizeKey("GET", "/api/v1/changes?days=7&limit=200")]: {
        status: 200,
        body: {
          ...base,
          items: [
            change("c-3", "2026-10-20T09:00:00.000Z", {
              targets: [
                { id: "u-1", name: "Blake Example", type: "User", principal: null, properties: [] },
                {
                  id: "u-2",
                  name: "İlker ΟΔΥΣΣΕΥΣ",
                  type: "User",
                  principal: null,
                  properties: [],
                },
              ],
            }),
            change("c-2", "2026-10-19T09:00:00.000Z", {
              category: "RoleManagement",
              activity: "Add member to role",
            }),
          ],
          next: "2026-10-19T09:00:00.000Z|c-2",
        },
      },
      [normalizeKey(
        "GET",
        `/api/v1/changes?days=7&limit=200&before=${encodeURIComponent("2026-10-19T09:00:00.000Z|c-2")}`,
      )]: {
        status: 200,
        body: {
          ...base,
          items: [change("c-1", "2026-10-18T09:00:00.000Z", { result: "timeout" })],
          next: null,
        },
      },
    },
  };

  it("filters and pages the recorded range the way the server does", () => {
    const state = createDemoState(changes, now(recordedAt));
    const get = (url: string) => handleRequest(state, { method: "GET", url });
    const ids = (url: string) =>
      (get(url).body as { items: { id: string }[] }).items.map((entry) => entry.id);
    expect(ids("/api/v1/changes?days=7&limit=100")).toEqual(["c-3", "c-2", "c-1"]);
    expect(ids("/api/v1/changes?days=7&limit=100&category=rolemanagement")).toEqual(["c-2"]);
    expect(ids("/api/v1/changes?days=7&limit=100&result=failure")).toEqual(["c-1"]);
    expect(ids("/api/v1/changes?days=7&limit=100&search=BLAKE")).toEqual(["c-3", "c-2", "c-1"]);
    expect(ids("/api/v1/changes?days=7&limit=100&search=role")).toEqual(["c-2"]);
    // Folded the way the server folds it: dotted capitals and final sigma.
    expect(ids("/api/v1/changes?days=7&limit=100&search=ilker")).toEqual(["c-3"]);
    expect(ids(`/api/v1/changes?days=7&limit=100&search=${encodeURIComponent("ΟΔΥΣ")}`)).toEqual([
      "c-3",
    ]);
    const first = get("/api/v1/changes?days=7&limit=2").body as { next: string | null };
    expect(first.next).toBe("2026-10-19T09:00:00.000Z|c-2");
    expect(
      ids(`/api/v1/changes?days=7&limit=2&before=${encodeURIComponent(first.next ?? "")}`),
    ).toEqual(["c-1"]);
    expect(get("/api/v1/changes?days=7&before=nowhere").status).toBe(400);
    // A range nobody recorded is not available.
    expect(get("/api/v1/changes?days=30&limit=100").status).toBe(404);
  });
});

describe("canary accounts in the demo", () => {
  const person = (id: string, principal: string) => ({
    [`GET /api/v1/people/${id}/directory`]: {
      status: 200,
      body: { state: "available", message: null, profile: { userPrincipalName: principal } },
    },
  });
  const canaries: Fixtures = {
    recordedAt,
    responses: {
      "GET /api/v1/settings/canaries": {
        status: 200,
        body: {
          accounts: [
            {
              userId: "u-1",
              principal: "decoy.one@larkspur.example",
              addedAt: "2026-10-20T09:00:00.000Z",
              addedBy: "demo-administrator",
            },
          ],
        },
      },
      ...person("u-2", "decoy.two@larkspur.example"),
      "GET /api/v1/people/u-3": {
        status: 200,
        body: {
          history: { userPrincipalName: "decoy.three@larkspur.example" },
          findings: [],
        },
      },
    },
  };
  const send = (state: DemoState, method: string, body?: unknown) =>
    handleRequest(state, { method, url: "/api/v1/settings/canaries", body });
  const ids = (state: DemoState) =>
    ((send(state, "GET").body as { accounts: { userId: string }[] }).accounts ?? []).map(
      (account) => account.userId,
    );

  it("answers the recorded list, moved to today", () => {
    const state = createDemoState(canaries, now("2026-10-22T10:00:00.000Z"));
    const body = send(state, "GET").body as { accounts: { addedAt: string }[] };
    expect(ids(state)).toEqual(["u-1"]);
    expect(body.accounts[0]?.addedAt).toBe("2026-10-22T09:00:00.000Z");
  });

  it("marks and unmarks accounts for this page only", () => {
    const state = createDemoState(canaries, now("2026-10-22T10:00:00.000Z"));
    // An account the recording opened gets its recorded name; marking it twice changes nothing.
    expect(send(state, "POST", { add: ["u-2", "u-3", "u-2"] }).status).toBe(200);
    expect(ids(state)).toEqual(["u-1", "u-2", "u-3"]);
    const added = (send(state, "GET").body as { accounts: { principal: string }[] }).accounts;
    expect(added.map((account) => account.principal)).toEqual([
      "decoy.one@larkspur.example",
      "decoy.two@larkspur.example",
      "decoy.three@larkspur.example",
    ]);
    expect(send(state, "POST", { remove: ["u-1", "u-9"] }).status).toBe(200);
    expect(ids(state)).toEqual(["u-2", "u-3"]);
    // A fresh page starts from the recording again.
    expect(ids(createDemoState(canaries, now("2026-10-22T10:00:00.000Z")))).toEqual(["u-1"]);
  });

  it("refuses an account the recording never opened, and a body that names nothing", () => {
    const state = createDemoState(canaries, now("2026-10-22T10:00:00.000Z"));
    expect(send(state, "POST", { add: ["u-404"] }).status).toBe(404);
    expect(send(state, "POST", {}).status).toBe(400);
    expect(ids(state)).toEqual(["u-1"]);
  });

  it("is unavailable when the recording has no list", () => {
    const state = createDemoState({ recordedAt, responses: {} }, now(recordedAt));
    expect(send(state, "GET").status).toBe(404);
    expect(send(state, "POST", { add: ["u-2"] }).status).toBe(404);
  });
});

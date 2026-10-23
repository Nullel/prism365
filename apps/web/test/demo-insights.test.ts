import { auditScheduleSettingsSchema, auditTrendSchema } from "@prism365/contracts";
import { describe, expect, it } from "vitest";
import { createDemoState, type Fixtures, handleRequest } from "../src/demo/replay.ts";

// Audit trends and schedules in the portfolio demo: trends replay from the recording, schedule
// changes are kept for the page. Invented data only.

const recordedAt = "2026-10-20T10:00:00.000Z";
const trend = {
  auditId: "accounts.inactive",
  days: 365,
  points: [
    {
      runId: "00000000-0000-4000-8000-000000000001",
      finishedAt: "2026-10-06T10:00:00.000Z",
      version: 3,
      complete: true,
      scopeLimited: false,
      trigger: "manual",
      tiles: { accounts: 40, "inactive-30-days": 6 },
      notComparable: [],
    },
    {
      runId: "00000000-0000-4000-8000-000000000002",
      finishedAt: "2026-10-20T10:00:00.000Z",
      version: 3,
      complete: true,
      scopeLimited: false,
      trigger: "manual",
      tiles: { accounts: 40, "inactive-30-days": 5 },
      notComparable: [],
    },
  ],
  tiles: [
    { id: "accounts", label: "Accounts", tone: "neutral" },
    { id: "inactive-30-days", label: "Inactive 30+ days", tone: "warning" },
  ],
  truncated: false,
};
const schedule = {
  audits: {
    "accounts.inactive": "weekly",
    "mfa.registration": "daily",
    "guests.review": "off",
  },
  collecting: true,
};
const collection = {
  available: true,
  state: "collecting",
  credentialKind: "delegated",
  boundAt: "2026-10-20T07:00:00.000Z",
  lastStopReason: null,
  durableStorage: false,
  reason: null,
  resumes: false,
};
const fixtures: Fixtures = {
  recordedAt,
  responses: {
    "GET /api/v1/modules": { status: 200, body: { mode: "live", collection, modules: [] } },
    "GET /api/v1/settings/audit-schedule": { status: 200, body: schedule },
    "GET /api/v1/audits/accounts.inactive/trend?days=365": { status: 200, body: trend },
  },
};

const start = () => {
  const state = createDemoState(fixtures, () => new Date("2026-10-20T12:00:00.000Z"));
  const request = (method: string, url: string, body?: unknown) =>
    handleRequest(state, { method, url, ...(body === undefined ? {} : { body }) });
  return { state, request };
};

describe("audit trends in the demo", () => {
  it("replays the recorded trend the page asks for", () => {
    const { request } = start();
    const answer = request("GET", "/api/v1/audits/accounts.inactive/trend?days=365");
    expect(answer.status).toBe(200);
    expect(auditTrendSchema.parse(answer.body).points).toHaveLength(2);
  });
});

describe("audit schedules in the demo", () => {
  it("answers the recorded schedule", () => {
    const { request } = start();
    const answer = request("GET", "/api/v1/settings/audit-schedule");
    expect(answer.status).toBe(200);
    expect(auditScheduleSettingsSchema.parse(answer.body)).toEqual(schedule);
  });

  it("saves a change for the page, changing only the audits named and dropping unknown ones", () => {
    const { request } = start();
    const saved = request("POST", "/api/v1/settings/audit-schedule", {
      audits: { "guests.review": "daily", "retired.audit": "weekly" },
    });
    expect(saved.status).toBe(200);
    expect(saved.body).toEqual({
      audits: {
        "accounts.inactive": "weekly",
        "mfa.registration": "daily",
        "guests.review": "daily",
      },
      collecting: true,
    });
    expect(request("GET", "/api/v1/settings/audit-schedule").body).toEqual(saved.body);
  });

  it("refuses what the server refuses", () => {
    const { request } = start();
    for (const body of [{}, { audits: { "accounts.inactive": "hourly" } }, { audits: "daily" }])
      expect(request("POST", "/api/v1/settings/audit-schedule", body).status).toBe(400);
    expect(request("GET", "/api/v1/settings/audit-schedule").body).toEqual(schedule);
  });

  it("says collection is off once the page stops it", () => {
    const { request } = start();
    request("POST", "/api/v1/collection/stop");
    expect(request("GET", "/api/v1/settings/audit-schedule").body).toMatchObject({
      collecting: false,
    });
  });
});

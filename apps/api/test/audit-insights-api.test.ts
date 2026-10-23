import { randomUUID } from "node:crypto";
import { finishRun, inactiveAccountsAudit, insertRun } from "@prism365/audits";
import { auditScheduleSettingsSchema, auditTrendSchema } from "@prism365/contracts";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { type AuditHarness, actorId, auditHarness, tenantId } from "./audit-harness.ts";

// Audit trend and schedule routes. Synthetic tenant only.

const DAY = 86_400_000;

// A succeeded run stored directly, finished `daysAgo` days ago.
async function storedRun(
  h: AuditHarness,
  options: { daysAgo: number; version?: number; scope?: string[]; complete?: boolean },
) {
  const id = randomUUID();
  const finished = new Date(Date.now() - options.daysAgo * DAY);
  await h.database.tenant(tenantId, async (sql) => {
    await insertRun(sql, tenantId, {
      id,
      auditId: "accounts.inactive",
      auditVersion: options.version ?? 1,
      startedAt: new Date(finished.getTime() - 60_000),
      startedBy: { id: actorId, name: null },
      scope: options.scope ?? null,
    });
    await finishRun(sql, tenantId, id, finished, {
      status: "succeeded",
      complete: options.complete ?? true,
      summary: [
        { label: "Accounts", value: 10 + options.daysAgo, tone: "neutral", view: "all" },
        { label: "Inactive 30+ days", value: options.daysAgo, tone: "warning", view: "d30" },
      ],
      notes: [],
      rows: [],
    });
  });
  return id;
}

describe("audit trend route", () => {
  let h: AuditHarness;
  beforeAll(async () => {
    h = await auditHarness();
  }, 60_000);
  afterAll(async () => {
    await h?.app.close();
  });

  it("needs a session and a known audit", async () => {
    expect((await h.get("/v1/audits/accounts.inactive/trend", null)).statusCode).toBe(401);
    for (const id of ["nope.audit", "not-an-id", "accounts.inactive.x"]) {
      const response = await h.get(`/v1/audits/${id}/trend`);
      expect(response.statusCode, id).toBe(404);
      expect(response.json()).toMatchObject({ error: "unknown_audit" });
    }
  });

  it("refuses a window outside 1 to 3,650 days", async () => {
    for (const days of ["0", "3651", "-5", "abc", "1.5", ""]) {
      const response = await h.get(`/v1/audits/accounts.inactive/trend?days=${days}`);
      expect(response.statusCode, days).toBe(400);
      expect(response.json()).toMatchObject({ error: "invalid_request" });
    }
    const repeated = await h.get("/v1/audits/accounts.inactive/trend?days=30&days=60");
    expect(repeated.statusCode).toBe(400);
  });

  it("is empty before any run, and has a point for each run after, oldest first", async () => {
    const empty = auditTrendSchema.parse(
      (await h.get("/v1/audits/accounts.inactive/trend")).json(),
    );
    expect(empty).toEqual({
      auditId: "accounts.inactive",
      days: 365,
      points: [],
      tiles: [],
      truncated: false,
    });
    const first = await h.finish("accounts.inactive");
    const second = await h.finish("accounts.inactive");
    const response = await h.get("/v1/audits/accounts.inactive/trend");
    expect(response.statusCode).toBe(200);
    const trend = auditTrendSchema.parse(response.json());
    expect(trend.points.map((point) => point.runId)).toEqual([first, second]);
    expect(trend.points[1]).toMatchObject({
      version: inactiveAccountsAudit.definition.version,
      complete: true,
      scopeLimited: false,
      trigger: "manual",
      notComparable: [],
    });
    expect(trend.points[1]?.tiles).toMatchObject({ accounts: 2, "never-signed-in": 0 });
    expect(trend.tiles[0]).toEqual({ id: "accounts", label: "Accounts", tone: "neutral" });
    // Reading stored counts is not an access-ledger event; the runs were.
    expect(h.store.audit.filter((entry) => entry.event.type === "audit.run.viewed")).toEqual([]);
  });

  it("flags points that do not compare with the newest one, and keeps to the window", async () => {
    const fresh = await auditHarness();
    try {
      const old = await storedRun(fresh, { daysAgo: 400 });
      const version = await storedRun(fresh, { daysAgo: 30, version: 2 });
      const scoped = await storedRun(fresh, { daysAgo: 20, scope: ["a.example"] });
      const partial = await storedRun(fresh, { daysAgo: 10, complete: false });
      const newest = await storedRun(fresh, { daysAgo: 1 });
      const trend = auditTrendSchema.parse(
        (await fresh.get("/v1/audits/accounts.inactive/trend")).json(),
      );
      expect(trend.points.map((point) => [point.runId, point.notComparable])).toEqual([
        [version, ["version"]],
        [scoped, ["scope"]],
        [partial, ["incomplete"]],
        [newest, []],
      ]);
      expect(trend.points.map((point) => point.scopeLimited)).toEqual([false, true, false, false]);
      expect(trend.points.map((point) => point.tiles["inactive-30-days"])).toEqual([30, 20, 10, 1]);
      const year = auditTrendSchema.parse(
        (await fresh.get("/v1/audits/accounts.inactive/trend?days=500")).json(),
      );
      expect(year.days).toBe(500);
      expect(year.points[0]?.runId).toBe(old);
      const week = auditTrendSchema.parse(
        (await fresh.get("/v1/audits/accounts.inactive/trend?days=7")).json(),
      );
      expect(week.points.map((point) => point.runId)).toEqual([newest]);
      // Another audit has its own trend.
      const other = auditTrendSchema.parse(
        (await fresh.get("/v1/audits/mfa.registration/trend")).json(),
      );
      expect(other.points).toEqual([]);
    } finally {
      await fresh.app.close();
    }
  }, 60_000);
});

describe("audit schedule settings", () => {
  let h: AuditHarness;
  beforeAll(async () => {
    h = await auditHarness();
  }, 60_000);
  afterAll(async () => {
    await h?.app.close();
  });
  const settings = async () =>
    auditScheduleSettingsSchema.parse((await h.get("/v1/settings/audit-schedule")).json());
  const changes = () =>
    h.store.audit.filter((entry) => entry.event.type === "audits.schedule.changed");

  it("needs a session, and the CSRF header to save", async () => {
    expect((await h.get("/v1/settings/audit-schedule", null)).statusCode).toBe(401);
    const body = { audits: { "accounts.inactive": "daily" } };
    expect((await h.post("/v1/settings/audit-schedule", body, null)).statusCode).toBe(401);
    expect((await h.post("/v1/settings/audit-schedule", body, h.who, false)).statusCode).toBe(403);
    expect((await settings()).audits["accounts.inactive"]).toBe("off");
    expect(changes()).toEqual([]);
  });

  it("lists every audit as off, and says collection is off", async () => {
    const current = await settings();
    expect(Object.values(current.audits).every((frequency) => frequency === "off")).toBe(true);
    expect(Object.keys(current.audits)).toContain("licenses.capacity");
    expect(current.collecting).toBe(false);
  });

  it("refuses anything but a list of audits with off, daily or weekly", async () => {
    const bodies: unknown[] = [
      undefined,
      {},
      { audits: "daily" },
      { audits: [] },
      { audits: { "accounts.inactive": "hourly" } },
      { audits: { "accounts.inactive": true } },
      { audits: { "Not an id": "daily" } },
      { audits: Object.fromEntries(Array.from({ length: 101 }, (_, i) => [`a.b${i}`, "daily"])) },
    ];
    for (const body of bodies) {
      const response = await h.post("/v1/settings/audit-schedule", body);
      expect(response.statusCode, JSON.stringify(body)).toBe(400);
      expect(response.json(), JSON.stringify(body)).toMatchObject({ error: "invalid_request" });
    }
    // A "__proto__" key is refused by the JSON parser before the route sees it.
    const prototype = await h.post(
      "/v1/settings/audit-schedule",
      JSON.parse('{"audits":{"__proto__":"daily"}}'),
    );
    expect(prototype.statusCode).toBe(400);
    expect(changes()).toEqual([]);
    expect(Object.values((await settings()).audits).every((value) => value === "off")).toBe(true);
  });

  it("saves the schedule, drops unknown audits, and records only a count first", async () => {
    const response = await h.post("/v1/settings/audit-schedule", {
      audits: {
        "accounts.inactive": "daily",
        "mfa.registration": "weekly",
        "licenses.reclaim": "off",
        "retired.audit": "daily",
      },
    });
    expect(response.statusCode).toBe(200);
    const saved = auditScheduleSettingsSchema.parse(response.json());
    expect(saved.audits).toMatchObject({
      "accounts.inactive": "daily",
      "mfa.registration": "weekly",
      "licenses.reclaim": "off",
      "guests.review": "off",
    });
    expect(saved.audits).not.toHaveProperty("retired.audit");
    expect(await settings()).toEqual(saved);
    expect(changes()).toHaveLength(1);
    expect(changes()[0]).toEqual({
      tenantId,
      actorId,
      event: { type: "audits.schedule.changed", scheduled: 2, occurredAt: expect.any(String) },
    });
  });

  it("does not change the schedule when the change cannot be recorded", async () => {
    const before = await settings();
    const append = h.store.appendAudit.bind(h.store);
    h.store.appendAudit = async (tenant, actor, event) => {
      if (event.type === "audits.schedule.changed") throw new Error("ledger unavailable");
      return append(tenant, actor, event);
    };
    const failed = await h.post("/v1/settings/audit-schedule", { audits: {} });
    h.store.appendAudit = append;
    expect(failed.statusCode).toBe(500);
    expect(await settings()).toEqual(before);
  });

  it("says when collection is on", async () => {
    const started = await h.post("/v1/collection/start");
    expect(started.statusCode).toBe(200);
    expect((await settings()).collecting).toBe(true);
    await h.post("/v1/collection/stop");
    expect((await settings()).collecting).toBe(false);
  });

  it("shows a manual run as manual", async () => {
    const id = await h.finish("licenses.reclaim");
    const runs = (await h.get("/v1/audits/licenses.reclaim/runs")).json().runs;
    expect(runs.find((run: { id: string }) => run.id === id)?.trigger).toBe("manual");
  });
});

describe("audit trend and schedule routes without module storage", () => {
  it("answer 409 like the other audit routes", async () => {
    const h = await auditHarness({ withModules: false });
    try {
      for (const url of ["/v1/audits/accounts.inactive/trend", "/v1/settings/audit-schedule"]) {
        const response = await h.get(url);
        expect(response.statusCode, url).toBe(409);
        expect(response.json()).toMatchObject({ error: "audits_unavailable" });
      }
      const saved = await h.post("/v1/settings/audit-schedule", { audits: {} });
      expect(saved.statusCode).toBe(409);
    } finally {
      await h.app.close();
    }
  }, 60_000);
});

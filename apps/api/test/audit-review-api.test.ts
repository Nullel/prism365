import { auditReviewResponseSchema } from "@prism365/contracts";
import { afterAll, describe, expect, it } from "vitest";
import { type AuditHarness, auditHarness } from "./audit-harness.ts";

// The full review routes on the real runner and a synthetic tenant. Only two read packs are
// enabled here, so most audits are skipped with the reason, which is part of what is tested.

const harnesses: AuditHarness[] = [];
async function harness(options: Parameters<typeof auditHarness>[0] = {}) {
  const h = await auditHarness(options);
  harnesses.push(h);
  return h;
}
afterAll(async () => {
  for (const h of harnesses) await h.app.close();
});

async function finished(h: AuditHarness) {
  for (let turn = 0; turn < 600; turn += 1) {
    const review = auditReviewResponseSchema.parse(
      (await h.get("/v1/audits/review")).json(),
    ).review;
    if (review && review.state !== "running" && review.state !== "stopping") return review;
    await new Promise((resolve) => setTimeout(resolve, 25));
  }
  throw new Error("The review never ended.");
}
const types = (h: AuditHarness) => h.store.audit.map((entry) => String(entry.event.type));

// Each test builds a whole API on a fresh database, which can take longer than the default five
// seconds while the full suite runs in parallel; the other API suites allow a minute too.
describe("the review routes", { timeout: 60_000 }, () => {
  it("answer nobody without a session, and refuse a start or stop without the CSRF token", async () => {
    const h = await harness();
    for (const [method, url] of [
      ["GET", "/v1/audits/review"],
      ["GET", "/v1/settings/audit-review"],
    ] as const)
      expect((await h.app.inject({ method, url })).statusCode, url).toBe(401);
    for (const url of [
      "/v1/audits/review",
      "/v1/audits/review/stop",
      "/v1/settings/audit-review",
    ]) {
      expect((await h.post(url, {}, null)).statusCode, `${url} without a session`).toBe(401);
      expect((await h.post(url, {}, h.who, false)).statusCode, `${url} without a token`).toBe(403);
    }
    expect(types(h)).not.toContain("audit.review.started");
  });

  it("say there is no review before one has run", async () => {
    const h = await harness();
    const response = await h.get("/v1/audits/review");
    expect(response.statusCode).toBe(200);
    expect(response.json()).toEqual({ review: null });
    const stop = await h.post("/v1/audits/review/stop");
    expect(stop.statusCode).toBe(409);
    expect(stop.json().error).toBe("no_review_running");
  });

  it("answer 409 where audits are not available", async () => {
    const h = await harness({ withModules: false });
    expect((await h.get("/v1/audits/review")).statusCode).toBe(409);
    expect((await h.post("/v1/audits/review")).statusCode).toBe(409);
    expect((await h.get("/v1/settings/audit-review")).statusCode).toBe(409);
  });

  it("run the audits this session can run, skip the rest with the reason, and end with Combined risks", async () => {
    const h = await harness();
    const started = await h.post("/v1/audits/review");
    expect(started.statusCode).toBe(202);
    const first = auditReviewResponseSchema.parse({ review: started.json().review }).review;
    expect(first?.audits).toHaveLength(17);
    expect(first?.audits.at(-1)?.auditId).toBe("identity.combined-risks");
    const review = await finished(h);
    expect(review.state).toBe("finished");
    expect(review.trigger).toBe("manual");
    expect(review.startedBy.name).toBe("Synthetic admin");
    expect(review.finishedAt).not.toBeNull();
    const skipped = review.audits.filter((item) => item.state === "skipped");
    expect(skipped.length).toBeGreaterThan(0);
    for (const item of skipped) expect(item.reason, item.auditId).toContain("has not enabled");
    const ran = review.audits.filter((item) => item.state !== "skipped").map((a) => a.auditId);
    expect(ran).toContain("accounts.inactive");
    expect(ran.at(-1)).toBe("identity.combined-risks");
    for (const item of review.audits.filter((entry) => entry.state === "done"))
      expect(item.runId, item.auditId).not.toBeNull();
    // The runs are ordinary runs, each with its own ledger entry naming the review.
    const runEntries = h.store.audit.filter((entry) => entry.event.type === "audit.run.started");
    expect(runEntries).toHaveLength(ran.length);
    expect(runEntries.every((entry) => entry.event.review === true)).toBe(true);
  });

  it("write the review's ledger entry before any audit reads, holding counts and no names", async () => {
    const h = await harness();
    await h.post("/v1/audits/review");
    await finished(h);
    const events = types(h).filter((type) => type.startsWith("audit."));
    expect(events[0]).toBe("audit.review.started");
    expect(events.indexOf("audit.review.started")).toBeLessThan(
      events.indexOf("audit.run.started"),
    );
    const entry = h.store.audit.find((item) => item.event.type === "audit.review.started")?.event;
    expect(entry).toMatchObject({ trigger: "manual" });
    expect(typeof entry?.audits).toBe("number");
    expect(typeof entry?.skipped).toBe("number");
    expect(JSON.stringify(entry)).not.toContain("Synthetic admin");
    // The first Graph read came after the ledger entry.
    expect(h.reads[0]?.ledgerStarts).toBeGreaterThan(0);
  });

  it("refuse a second review while one is running, and stop the first when asked", async () => {
    const h = await harness();
    const gate = h.holdUsers();
    const started = await h.post("/v1/audits/review");
    expect(started.statusCode).toBe(202);
    for (let turn = 0; turn < 400 && gate.waiting() === 0; turn += 1)
      await new Promise((resolve) => setTimeout(resolve, 10));
    expect(gate.waiting()).toBeGreaterThan(0);
    const again = await h.post("/v1/audits/review");
    expect(again.statusCode).toBe(409);
    expect(again.json().error).toBe("review_running");
    expect(types(h).filter((type) => type === "audit.review.started")).toHaveLength(1);

    const stop = await h.post("/v1/audits/review/stop");
    expect(stop.statusCode).toBe(202);
    expect(types(h).filter((type) => type === "audit.review.stopped")).toHaveLength(1);
    const events = types(h);
    expect(events.indexOf("audit.review.stopped")).toBeGreaterThan(
      events.indexOf("audit.review.started"),
    );
    gate.release();
    const review = await finished(h);
    expect(review.state).toBe("stopped");
    expect(review.audits.find((item) => item.auditId === "identity.combined-risks")?.state).toBe(
      "skipped",
    );
    expect(review.audits.find((item) => item.state === "failed")?.reason).toContain("stopped");
    // A new review can start once it has ended.
    expect((await h.post("/v1/audits/review")).statusCode).toBe(202);
    await finished(h);
  });

  it("refuse to start when nothing can run, and write no ledger entry", async () => {
    const h = await harness({ scopes: ["User.Read"] });
    const response = await h.post("/v1/audits/review");
    expect(response.statusCode).toBe(403);
    expect(response.json().error).toBe("nothing_to_run");
    expect(types(h)).not.toContain("audit.review.started");
    expect(h.reads).toEqual([]);
  });

  // Tenant settings has its own read pack and also reads a stored run. It still reads Microsoft,
  // so a review with only it and Combined risks available has something to run.
  it("start when only Tenant settings and Combined risks can run, and refuse when only Combined risks can", async () => {
    const only = await harness({
      scopes: ["User.Read", "Policy.Read.All"],
      features: ["identity.policies.read"],
    });
    const started = await only.post("/v1/audits/review");
    expect(started.statusCode).toBe(202);
    const review = await finished(only);
    const ran = review.audits.filter((item) => item.state !== "skipped").map((a) => a.auditId);
    expect(ran).toEqual(["tenant.settings", "identity.combined-risks"]);
    expect(types(only)).toContain("audit.review.started");

    // Nothing granted that any pack-reading audit needs: Combined risks alone does not start one.
    const none = await harness({ scopes: ["User.Read"], features: ["identity.policies.read"] });
    const refused = await none.post("/v1/audits/review");
    expect(refused.statusCode).toBe(403);
    expect(refused.json().error).toBe("nothing_to_run");
    expect(types(none)).not.toContain("audit.review.started");
  });
});

describe("the review schedule routes", { timeout: 60_000 }, () => {
  it("start off, save a weekly day and hour in the configured time zone, and ledger the change first", async () => {
    const h = await harness();
    const off = await h.get("/v1/settings/audit-review");
    expect(off.json()).toMatchObject({ frequency: "off", timeZone: "UTC", collecting: false });
    const saved = await h.post("/v1/settings/audit-review", {
      frequency: "weekly",
      weekday: 4,
      hour: 21,
    });
    expect(saved.statusCode).toBe(200);
    expect(saved.json()).toMatchObject({
      frequency: "weekly",
      weekday: 4,
      hour: 21,
      by: "Synthetic admin",
    });
    expect((await h.get("/v1/settings/audit-review")).json()).toMatchObject({
      frequency: "weekly",
      weekday: 4,
      hour: 21,
    });
    expect(
      h.store.audit.find((entry) => entry.event.type === "audits.review.schedule.changed")?.event,
    ).toMatchObject({
      frequency: "weekly",
      weekday: 4,
      hour: 21,
    });
    const turnedOff = await h.post("/v1/settings/audit-review", {
      frequency: "off",
      weekday: 4,
      hour: 21,
    });
    expect(turnedOff.json()).toMatchObject({ frequency: "off", by: null });
  });

  it("refuse a day or hour that does not exist, and a frequency that is not offered", async () => {
    const h = await harness();
    for (const body of [
      { frequency: "weekly", weekday: 7, hour: 3 },
      { frequency: "weekly", weekday: 1, hour: 24 },
      { frequency: "daily", weekday: 1, hour: 3 },
      { frequency: "weekly", weekday: 1.5, hour: 3 },
      {},
    ]) {
      const response = await h.post("/v1/settings/audit-review", body);
      expect(response.statusCode, JSON.stringify(body)).toBe(400);
    }
    expect(types(h)).not.toContain("audits.review.schedule.changed");
  });
});

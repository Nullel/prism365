import { MAX_REPORT_DOWNLOADS, REPORT_WINDOW_MILLISECONDS } from "@prism365/audits";
import { describe, expect, it } from "vitest";
import { ReportBudget } from "../src/report-budget.ts";

// The report-download budget on a clock the test turns: a pause moves the clock by what it was
// asked to wait.
const MINUTE = 60_000;

function budget(limit?: number) {
  let now = 1_000_000;
  const waits: number[] = [];
  const instance = new ReportBudget({
    now: () => now,
    pause: async (milliseconds) => {
      waits.push(milliseconds);
      now += milliseconds;
    },
    ...(limit === undefined ? {} : { limit }),
  });
  return {
    instance,
    waits,
    advance: (milliseconds: number) => {
      now += milliseconds;
    },
    at: () => now,
  };
}
const signal = new AbortController().signal;

describe("the report budget", () => {
  it("lets audits start while their downloads fit the limit", async () => {
    const b = budget();
    await b.instance.waitFor(7, signal);
    b.instance.record(7);
    await b.instance.waitFor(5, signal);
    expect(b.waits).toEqual([]);
    expect(b.instance.used()).toBe(7);
  });

  it("waits until the oldest audit leaves the ten minutes when the next would pass the limit", async () => {
    const b = budget();
    b.instance.record(7);
    b.advance(MINUTE);
    b.instance.record(5);
    const endedFirst = b.at() - MINUTE;
    await b.instance.waitFor(2, signal);
    expect(b.waits).toEqual([endedFirst + REPORT_WINDOW_MILLISECONDS - (endedFirst + MINUTE)]);
    expect(b.at()).toBe(endedFirst + REPORT_WINDOW_MILLISECONDS);
    expect(b.instance.used()).toBe(5);
  });

  it("waits for as many audits as it must, one after the other", async () => {
    const b = budget();
    b.instance.record(6);
    b.advance(MINUTE);
    b.instance.record(6);
    await b.instance.waitFor(12, signal);
    expect(b.waits).toHaveLength(2);
    expect(b.instance.used()).toBe(0);
  });

  it("counts exactly the limit as allowed and one more as too many", async () => {
    const b = budget();
    b.instance.record(MAX_REPORT_DOWNLOADS - 1);
    await b.instance.waitFor(1, signal);
    expect(b.waits).toEqual([]);
    await b.instance.waitFor(2, signal);
    expect(b.waits).toHaveLength(1);
  });

  it("does not wait for an audit that alone is over the limit more than it must", async () => {
    const b = budget(5);
    await b.instance.waitFor(9, signal);
    expect(b.waits).toEqual([]);
    b.instance.record(9);
    expect(b.instance.used()).toBe(5);
  });

  it("forgets what the wait freed even when the clock does not move", async () => {
    let now = 5_000;
    let pauses = 0;
    const instance = new ReportBudget({
      now: () => now,
      pause: async () => {
        pauses += 1;
      },
    });
    instance.record(12);
    now += 1;
    await instance.waitFor(3, signal);
    expect(pauses).toBe(1);
  });

  it("stops waiting when the review is stopped", async () => {
    const controller = new AbortController();
    const instance = new ReportBudget({
      now: () => 0,
      pause: async (_milliseconds, aborted) => {
        controller.abort(new Error("stopped"));
        aborted.throwIfAborted();
      },
    });
    instance.record(12);
    await expect(instance.waitFor(1, controller.signal)).rejects.toThrow("stopped");
  });

  it("charges nothing for an audit that downloads no reports", () => {
    const b = budget();
    b.instance.record(0);
    expect(b.instance.used()).toBe(0);
  });
});

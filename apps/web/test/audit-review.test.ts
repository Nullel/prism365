import type { AuditReview, AuditReviewScheduleSettings } from "@prism365/contracts";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import {
  AuditReviewPanel,
  reviewCounts,
  reviewIsActive,
  reviewStatus,
} from "../src/AuditReview.tsx";
import { findAll, hostTree, textOf } from "./markup.ts";

// The full review panel in each state it can be in, over synthetic reviews.

const NOW = Date.parse("2026-10-26T12:00:00Z");
const titles: Record<string, string> = {
  "accounts.inactive": "Inactive accounts",
  "accounts.passwords": "Passwords",
  "accounts.activity": "Account activity",
  "mail.mailbox-storage": "Mailbox storage",
  "identity.combined-risks": "Combined risks",
};
const titleOf = (id: string) => titles[id] ?? id;
const IDS = Object.keys(titles);

type State = AuditReview["audits"][number]["state"];
function review(
  states: State[],
  over: Partial<AuditReview> = {},
  reasons: Record<number, string> = {},
): AuditReview {
  return {
    id: "66666666-6666-4666-8666-666666666666",
    state: "running",
    trigger: "manual",
    startedAt: "2026-10-26T11:50:00.000Z",
    finishedAt: null,
    startedBy: { id: "22222222-2222-4222-8222-222222222222", name: "Synthetic admin" },
    audits: states.map((state, index) => ({
      auditId: IDS[index] ?? `extra.audit-${index}`,
      state,
      reason: reasons[index] ?? null,
      runId: null,
      startedAt: null,
      finishedAt: null,
    })),
    ...over,
  };
}

const schedule = (
  over: Partial<AuditReviewScheduleSettings> = {},
): AuditReviewScheduleSettings => ({
  frequency: "off",
  weekday: 0,
  hour: 6,
  timeZone: "America/Toronto",
  collecting: true,
  by: null,
  ...over,
});

function panel(props: Partial<Parameters<typeof AuditReviewPanel>[0]> = {}) {
  const calls: string[] = [];
  const element = createElement(AuditReviewPanel, {
    review: null,
    titleOf,
    busy: false,
    error: "",
    canRun: true,
    now: NOW,
    onRun: () => calls.push("run"),
    onStop: () => calls.push("stop"),
    onSchedule: (next) => calls.push(`schedule ${JSON.stringify(next)}`),
    ...props,
  });
  return { html: renderToStaticMarkup(element), tree: hostTree(element), calls };
}
const button = (tree: ReturnType<typeof hostTree>, label: string) =>
  findAll(tree, (host) => host.tag === "button" && textOf(host) === label)[0];

describe("the full review panel", () => {
  it("offers Run full review when no review has run", () => {
    const { html, tree, calls } = panel();
    expect(html).toContain("Full review");
    expect(html).toContain("No full review has run yet.");
    expect(html).not.toContain("<progress");
    const run = button(tree, "Run full review");
    expect(run?.props.disabled).toBe(false);
    (run?.props.onClick as (() => void) | undefined)?.();
    expect(calls).toEqual(["run"]);
    expect(button(tree, "Stop")).toBeUndefined();
  });

  it("shows n of m, the audit running, and Stop while a review runs", () => {
    const running = review(["done", "done", "running", "waiting", "waiting"]);
    const { html, tree, calls } = panel({ review: running });
    expect(html).toContain("2 of 5 audits finished or skipped. Running Account activity.");
    expect(html).toContain('<progress class="audit-review__bar" max="5" value="2"');
    expect(button(tree, "Run full review")).toBeUndefined();
    (button(tree, "Stop")?.props.onClick as (() => void) | undefined)?.();
    expect(calls).toEqual(["stop"]);
  });

  it("counts a skipped audit as ended, and names a wait between audits", () => {
    const { html } = panel({
      review: review(["done", "skipped", "waiting", "waiting", "waiting"]),
    });
    expect(html).toContain("2 of 5 audits finished or skipped.");
    expect(html).toContain("Waiting a moment to stay inside Microsoft&#x27;s limits.");
  });

  it("lists the audits that failed, by name, while it runs", () => {
    const { html } = panel({
      review: review(
        ["done", "failed", "running", "waiting", "waiting"],
        {},
        { 1: "Microsoft refused." },
      ),
    });
    expect(html).toContain("1 audit failed: Passwords.");
    expect(html).toContain("Microsoft refused.");
    const two = reviewStatus(review(["failed", "failed", "done", "done", "done"]), titleOf);
    expect(two.failures).toBe("2 audits failed: Inactive accounts and Passwords.");
  });

  it("disables Stop once stopping, and says what is ending", () => {
    const stopping = review(["done", "running", "waiting", "waiting", "waiting"], {
      state: "stopping",
    });
    const { html, tree } = panel({ review: stopping });
    expect(html).toContain("Stopping. 1 of 5 audits finished or skipped, and Passwords is ending.");
    expect(button(tree, "Stop")?.props.disabled).toBe(true);
    expect(reviewIsActive(stopping)).toBe(true);
  });

  it("shows when the last review finished, with the audits skipped and why", () => {
    const done = review(
      ["done", "done", "done", "skipped", "done"],
      {
        state: "finished",
        finishedAt: "2026-10-26T11:55:00.000Z",
      },
      { 3: "This installation has not enabled Reports.Read.All." },
    );
    const { html, tree } = panel({ review: done });
    expect(html).toContain("Finished: 4 audits done, 1 skipped.");
    expect(html).toContain("Last full review finished <time");
    expect(html).toContain("5 min ago");
    expect(html).toContain("This installation has not enabled Reports.Read.All.");
    expect(html).not.toContain("<progress");
    expect(button(tree, "Run full review")).toBeDefined();
  });

  it("says a review was stopped, and that a failure was seen", () => {
    const stopped = review(["done", "failed", "skipped", "skipped", "skipped"], {
      state: "stopped",
      finishedAt: "2026-10-26T11:58:00.000Z",
    });
    const { html } = panel({ review: stopped });
    expect(html).toContain("Stopped: 1 audit done, 3 skipped.");
    expect(html).toContain("Last full review stopped");
    expect(html).toContain("1 audit failed: Passwords.");
  });

  it("says a review a schedule started was started on its schedule", () => {
    const { html } = panel({
      review: review(["done", "done", "done", "done", "done"], {
        state: "finished",
        trigger: "schedule",
        finishedAt: "2026-10-26T11:55:00.000Z",
      }),
    });
    expect(html).toContain("started on its schedule");
  });

  it("keeps Run disabled, and says why, where a review cannot start", () => {
    const { html, tree } = panel({ canRun: false, note: "Demo mode cannot run a full review." });
    expect(button(tree, "Run full review")?.props.disabled).toBe(true);
    expect(html).toContain("Demo mode cannot run a full review.");
  });

  it("disables Run while the request is in flight, and shows a refused start", () => {
    const { html, tree } = panel({ busy: true, error: "A full review is already running." });
    expect(button(tree, "Run full review")?.props.disabled).toBe(true);
    expect(html).toContain('role="alert"');
    expect(html).toContain("A full review is already running.");
  });

  it("counts the audits by state", () => {
    expect(reviewCounts(review(["done", "failed", "failed", "skipped", "waiting"]))).toEqual({
      waiting: 1,
      running: 0,
      done: 1,
      failed: 2,
      skipped: 1,
    });
  });
});

describe("the weekly schedule control", () => {
  it("offers Off and Weekly, with the pressed one marked, and no day or hour while off", () => {
    const { html } = panel({ schedule: schedule() });
    expect(html).toContain("Run a full review automatically");
    expect(html).toContain('aria-pressed="true">Off<');
    expect(html).toContain('aria-pressed="false">Weekly<');
    expect(html).not.toContain("<select");
    expect(html).toContain("A full review runs only when someone starts it.");
  });

  it("shows the day and hour in the configured time zone once weekly, and says who set it", () => {
    const { html } = panel({
      schedule: schedule({ frequency: "weekly", weekday: 2, hour: 21, by: "Synthetic admin" }),
    });
    expect(html).toContain('<option value="2" selected="">Wednesday</option>');
    expect(html).toContain('<option value="21" selected="">21:00</option>');
    expect(html).toContain(
      "Runs every Wednesday at 21:00 (America/Toronto) while collection is on",
    );
    expect(html).toContain("Set by Synthetic admin.");
    expect(html.match(/<option/g)).toHaveLength(7 + 24);
  });

  it("warns that nothing runs while collection is off", () => {
    const { html } = panel({ schedule: schedule({ frequency: "weekly", collecting: false }) });
    expect(html).toContain("Collection is off, so the review will not run on its own");
    expect(panel({ schedule: schedule({ frequency: "weekly" }) }).html).not.toContain(
      "Collection is off",
    );
  });

  it("saves the whole schedule when one part changes", () => {
    const { tree, calls } = panel({
      schedule: schedule({ frequency: "weekly", weekday: 1, hour: 8 }),
    });
    (button(tree, "Off")?.props.onClick as (() => void) | undefined)?.();
    const selects = findAll(tree, (host) => host.tag === "select");
    (selects[0]?.props.onChange as ((event: unknown) => void) | undefined)?.({
      target: { value: "4" },
    });
    (selects[1]?.props.onChange as ((event: unknown) => void) | undefined)?.({
      target: { value: "17" },
    });
    expect(calls).toEqual([
      'schedule {"frequency":"off","weekday":1,"hour":8}',
      'schedule {"frequency":"weekly","weekday":4,"hour":8}',
      'schedule {"frequency":"weekly","weekday":1,"hour":17}',
    ]);
  });

  it("does nothing when the pressed choice is pressed again", () => {
    const { tree, calls } = panel({ schedule: schedule() });
    (button(tree, "Off")?.props.onClick as (() => void) | undefined)?.();
    expect(calls).toEqual([]);
  });

  it("disables the control while saving and shows a refused save", () => {
    const { html } = panel({
      schedule: schedule({ frequency: "weekly" }),
      scheduleSaving: true,
      scheduleError: "Choose off or weekly.",
    });
    expect(html).toContain("<fieldset");
    expect(html).toContain("disabled");
    expect(html).toContain("Saving…");
    expect(html).toContain("Choose off or weekly.");
  });

  it("is left out when the schedule has not loaded", () => {
    expect(panel({ schedule: null }).html).not.toContain("automatically");
  });
});

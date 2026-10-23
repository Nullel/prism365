import { GraphBoundaryError, GraphHttpClient } from "@prism365/graph";
import { scopeGraph } from "@prism365/modules";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  abortablePause,
  createPacer,
  type Pause,
  REQUEST_SPACING_MILLISECONDS,
  readPaced,
} from "../src/index.ts";

// Synthetic data only: paths and ids are invented.
const request = {
  feature: "identity.risk.read",
  permission: "IdentityRiskyUser.Read.All",
} as const;
const base = "https://graph.microsoft.com/v1.0/identityProtection/riskyUsers";

function graph(route: (url: URL) => unknown) {
  const urls: URL[] = [];
  const client = new GraphHttpClient({
    getAccessToken: async () => "synthetic",
    fetchImplementation: async (input) => {
      const url = new URL(String(input));
      urls.push(url);
      return Response.json(route(url));
    },
  });
  return { graph: scopeGraph(client, ["identity.risk.read"]), urls };
}

const pagesOf = (url: URL) => {
  const page = Number(url.searchParams.get("$skiptoken") ?? "0");
  return {
    value: [{ id: `user-${page}` }],
    "@odata.nextLink": `${base}?$skiptoken=${page + 1}`,
  };
};

afterEach(() => {
  vi.useRealTimers();
});

describe("the request pacer", () => {
  it("sends the first request at once and waits before every later one", async () => {
    const waits: number[] = [];
    const pause: Pause = async (milliseconds) => {
      waits.push(milliseconds);
    };
    const pacer = createPacer(pause);
    const signal = new AbortController().signal;
    await pacer.beforeRequest(signal);
    expect(waits).toEqual([]);
    await pacer.beforeRequest(signal);
    await pacer.beforeRequest(signal);
    expect(waits).toEqual([REQUEST_SPACING_MILLISECONDS, REQUEST_SPACING_MILLISECONDS]);
  });

  it("waits longer than the one second Microsoft allows between requests", () => {
    expect(REQUEST_SPACING_MILLISECONDS).toBeGreaterThan(1_000);
  });

  it("does not send a request once the audit's time limit has passed", async () => {
    const controller = new AbortController();
    const pacer = createPacer(async () => {});
    controller.abort(new Error("time limit"));
    await expect(pacer.beforeRequest(controller.signal)).rejects.toThrow("time limit");
  });
});

describe("the abortable pause", () => {
  it("resolves when the time is up", async () => {
    vi.useFakeTimers();
    let done = false;
    const waiting = abortablePause(1_100, new AbortController().signal).then(() => {
      done = true;
    });
    await vi.advanceTimersByTimeAsync(1_099);
    expect(done).toBe(false);
    await vi.advanceTimersByTimeAsync(1);
    await waiting;
    expect(done).toBe(true);
  });

  it("stops waiting as soon as the audit is stopped", async () => {
    vi.useFakeTimers();
    const controller = new AbortController();
    const waiting = abortablePause(60_000, controller.signal);
    const outcome = expect(waiting).rejects.toThrow("stopped");
    controller.abort(new Error("stopped"));
    await outcome;
    expect(vi.getTimerCount()).toBe(0);
  });

  it("does not start waiting when the audit is already stopped", async () => {
    vi.useFakeTimers();
    const controller = new AbortController();
    controller.abort(new Error("already stopped"));
    await expect(abortablePause(1_000, controller.signal)).rejects.toThrow("already stopped");
    expect(vi.getTimerCount()).toBe(0);
  });
});

describe("a paced read", () => {
  it("waits before every page after the first, in order with the requests", async () => {
    const trace: string[] = [];
    const { graph: scoped, urls } = graph((url) => {
      trace.push(`request ${url.searchParams.get("$skiptoken") ?? "first"}`);
      const page = Number(url.searchParams.get("$skiptoken") ?? "0");
      return page < 2 ? pagesOf(url) : { value: [{ id: "user-2" }] };
    });
    const pause: Pause = async (milliseconds) => {
      trace.push(`pause ${milliseconds}`);
    };
    const result = await readPaced(scoped, `${base}?$top=500`, request, {
      pacer: createPacer(pause),
      signal: new AbortController().signal,
      maxPages: 10,
      maxRecords: 100,
    });
    expect(result.complete).toBe(true);
    expect(result.items).toHaveLength(3);
    expect(urls[0]?.searchParams.get("$top")).toBe("500");
    expect(trace).toEqual([
      "request first",
      `pause ${REQUEST_SPACING_MILLISECONDS}`,
      "request 1",
      `pause ${REQUEST_SPACING_MILLISECONDS}`,
      "request 2",
    ]);
  });

  it("says the read is partial when it stops at the page limit", async () => {
    const { graph: scoped, urls } = graph(pagesOf);
    const result = await readPaced(scoped, base, request, {
      pacer: createPacer(async () => {}),
      signal: new AbortController().signal,
      maxPages: 3,
      maxRecords: 100,
    });
    expect(result.complete).toBe(false);
    expect(result.items).toHaveLength(3);
    expect(urls).toHaveLength(3);
  });

  it("says the read is partial when it stops at the record limit", async () => {
    const { graph: scoped } = graph(() => ({
      value: [{ id: "a" }, { id: "b" }],
      "@odata.nextLink": `${base}?$skiptoken=1`,
    }));
    const result = await readPaced(scoped, base, request, {
      pacer: createPacer(async () => {}),
      signal: new AbortController().signal,
      maxPages: 10,
      maxRecords: 2,
    });
    expect(result.complete).toBe(false);
    expect(result.items).toHaveLength(2);
  });

  it("is complete when the last page exactly fills the record limit", async () => {
    const { graph: scoped } = graph(() => ({ value: [{ id: "a" }, { id: "b" }] }));
    const result = await readPaced(scoped, base, request, {
      pacer: createPacer(async () => {}),
      signal: new AbortController().signal,
      maxPages: 10,
      maxRecords: 2,
    });
    expect(result.complete).toBe(true);
  });

  it("refuses a continuation link that leaves the pack's endpoints", async () => {
    const { graph: scoped, urls } = graph(() => ({
      value: [{ id: "a" }],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=1",
    }));
    await expect(
      readPaced(scoped, base, request, {
        pacer: createPacer(async () => {}),
        signal: new AbortController().signal,
        maxPages: 10,
        maxRecords: 100,
      }),
    ).rejects.toBeInstanceOf(GraphBoundaryError);
    expect(urls).toHaveLength(1);
  });

  it("refuses a continuation link that repeats", async () => {
    const { graph: scoped } = graph(() => ({
      value: [{ id: "a" }],
      "@odata.nextLink": base,
    }));
    await expect(
      readPaced(scoped, base, request, {
        pacer: createPacer(async () => {}),
        signal: new AbortController().signal,
        maxPages: 10,
        maxRecords: 100,
      }),
    ).rejects.toMatchObject({ reason: "pagination" });
  });

  it("stops between pages when the audit is stopped", async () => {
    const controller = new AbortController();
    const { graph: scoped, urls } = graph(pagesOf);
    await expect(
      readPaced(scoped, base, request, {
        pacer: createPacer(async () => {
          controller.abort(new Error("time limit"));
        }),
        signal: controller.signal,
        maxPages: 10,
        maxRecords: 100,
      }),
    ).rejects.toThrow("time limit");
    expect(urls).toHaveLength(1);
  });
});

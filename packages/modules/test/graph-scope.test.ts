import { GraphHttpClient } from "@prism365/graph";
import { describe, expect, it } from "vitest";
import { ModuleBoundaryError, scopeGraph } from "../src/graph-scope.ts";

// The usage report reader through a module's scoped Graph. Every value is invented.

const usage = { feature: "reports.usage.read", permission: "Reports.Read.All" } as const;
const REPORT = "reports/getOffice365ActiveUserDetail(period='D7')";
const LINK = "https://reports.office.com/data/download/synthetic-scope-token";

function client(onDownload: (init: RequestInit | undefined) => Promise<Response> | Response) {
  const urls: string[] = [];
  const graph = new GraphHttpClient({
    getAccessToken: async () => "synthetic",
    requestTimeoutMilliseconds: 60_000,
    fetchImplementation: (async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = new URL(String(input));
      urls.push(url.href);
      if (url.hostname === "graph.microsoft.com")
        return new Response(null, { status: 302, headers: { Location: LINK } });
      return onDownload(init);
    }) as typeof fetch,
  });
  return { graph, urls };
}

const csv = () =>
  new Response("\uFEFFUser Principal Name,Has Teams License\r\nada@lab.example,True\r\n", {
    headers: { "Content-Type": "application/octet-stream" },
  });

describe("scopeGraph getReport", () => {
  it("reads a report for a module that declared the usage reports pack", async () => {
    const { graph, urls } = client(csv);
    const guarded: string[] = [];
    const scoped = scopeGraph(graph, ["reports.usage.read"], () => guarded.push("checked"));
    await expect(scoped.getReport(REPORT, usage)).resolves.toMatchObject({
      rows: [{ "User Principal Name": "ada@lab.example", "Has Teams License": "True" }],
    });
    expect(guarded).toEqual(["checked"]);
    expect(urls).toHaveLength(2);
  });

  it("refuses the pack when the module did not declare it, before any request", async () => {
    const { graph, urls } = client(csv);
    const scoped = scopeGraph(graph, ["directory.users.read"]);
    await expect(scoped.getReport(REPORT, usage)).rejects.toBeInstanceOf(ModuleBoundaryError);
    expect(urls).toHaveLength(0);
  });

  it("stops before any request when the guard refuses", async () => {
    const { graph, urls } = client(csv);
    const scoped = scopeGraph(graph, ["reports.usage.read"], () => {
      throw new Error("run has ended");
    });
    await expect(scoped.getReport(REPORT, usage)).rejects.toThrow("run has ended");
    expect(urls).toHaveLength(0);
  });

  it("forwards the run's signal, so an ended run stops the download in flight", async () => {
    let downloadAborted = false;
    const { graph, urls } = client(
      (init) =>
        new Promise<Response>((_resolve, reject) => {
          init?.signal?.addEventListener("abort", () => {
            downloadAborted = true;
            reject(new DOMException("Aborted", "AbortError"));
          });
        }),
    );
    const ended = new AbortController();
    const scoped = scopeGraph(graph, ["reports.usage.read"], () => {}, ended.signal);
    const pending = scoped.getReport(REPORT, usage, { maxBytes: 1_048_576 });
    setTimeout(() => ended.abort(new Error("collector run ended")), 20);
    await expect(pending).rejects.toThrow("collector run ended");
    expect(downloadAborted).toBe(true);
    expect(urls).toHaveLength(2);
  });

  it("passes the byte cap through to the client", async () => {
    const { graph } = client(csv);
    const scoped = scopeGraph(graph, ["reports.usage.read"]);
    await expect(scoped.getReport(REPORT, usage, { maxBytes: 64 })).rejects.toThrow("maxBytes");
  });
});

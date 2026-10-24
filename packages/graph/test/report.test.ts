import { describe, expect, it, vi } from "vitest";
import {
  GraphBoundaryError,
  GraphHttpClient,
  GraphRequestError,
  GraphRetryDeferredError,
} from "../src/client.ts";
import {
  type GraphFeatureId,
  type GraphPermissionPack,
  type PermissionPack,
  permissionPacks,
} from "../src/permissions.ts";

// The usage report reader (ADR-0010) against a fake fetch. Every value is invented: the token,
// the download links and the .example addresses.

const usage = { feature: "reports.usage.read", permission: "Reports.Read.All" } as const;
const graphPacks = permissionPacks.filter(
  (pack: PermissionPack): pack is GraphPermissionPack => pack.source === "microsoftGraph",
);
const REPORT = "reports/getOffice365ActiveUserDetail(period='D30')";
const TOKEN = "synthetic-graph-token";
const SECRET = "SyntheticDownloadSecret_4f2a";
const LINK = `https://reports.office.com/data/download/${SECRET}`;
const CSV = "\uFEFFReport Refresh Date,User Principal Name\r\n2026-10-20,ada@lab.example\r\n";

interface Sent {
  url: string;
  headers: Headers;
  init: RequestInit | undefined;
}

type Answer = (url: URL, init: RequestInit | undefined) => Response | Promise<Response>;

// Graph answers with `graph`; any other host with `download`. Every request is recorded with its
// headers, so tests can check exactly what each host was sent.
function fakeFetch(graph: Answer, download: Answer = () => csvResponse(CSV)) {
  const sent: Sent[] = [];
  const fetchImplementation = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = new URL(String(input));
    sent.push({ url: url.href, headers: new Headers(init?.headers), init });
    return url.hostname === "graph.microsoft.com" ? graph(url, init) : download(url, init);
  });
  return { sent, fetchImplementation: fetchImplementation as typeof fetch };
}

const redirectTo = (location: string, status = 302) =>
  new Response(null, { status, headers: { Location: location } });
const csvResponse = (body: BodyInit, headers: Record<string, string> = {}) =>
  new Response(body, {
    status: 200,
    headers: { "Content-Type": "application/octet-stream", ...headers },
  });

function reader(
  fetchImplementation: typeof fetch,
  options: Partial<ConstructorParameters<typeof GraphHttpClient>[0]> = {},
) {
  return new GraphHttpClient({
    getAccessToken: async () => TOKEN,
    fetchImplementation,
    sleep: async () => undefined,
    ...options,
  });
}

async function failure(promise: Promise<unknown>): Promise<unknown> {
  return promise.then(
    () => {
      throw new Error("expected the report read to fail");
    },
    (error: unknown) => error,
  );
}

// A body far over any cap, one chunk a pull, that says how much was read and whether the reader
// let go. It ends after `chunks`, so a missing cap shows as a wrong answer (one header far over
// the field limit) instead of a read that never stops.
function largeBody(chunkBytes = 65_536, chunks = 64) {
  const state = { pulls: 0, cancelled: false };
  const stream = new ReadableStream<Uint8Array>({
    pull(controller) {
      state.pulls += 1;
      if (state.pulls > chunks) controller.close();
      else controller.enqueue(new Uint8Array(chunkBytes).fill(0x61));
    },
    cancel() {
      state.cancelled = true;
    },
  });
  return { stream, state };
}

// A fetch that only ends when its signal aborts, recording that it did.
const stalled = (aborted: { value: boolean }) => (_url: URL, init: RequestInit | undefined) =>
  new Promise<Response>((_resolve, reject) => {
    init?.signal?.addEventListener("abort", () => {
      aborted.value = true;
      reject(new DOMException("Aborted", "AbortError"));
    });
  });

describe("getReport: the Graph request and the download", () => {
  it("sends the token to Graph and only Accept to the download host", async () => {
    const { sent, fetchImplementation } = fakeFetch(() => redirectTo(LINK));
    const table = await reader(fetchImplementation).getReport(REPORT, usage);
    expect(table).toEqual({
      header: ["Report Refresh Date", "User Principal Name"],
      rows: [{ "Report Refresh Date": "2026-10-20", "User Principal Name": "ada@lab.example" }],
      skippedRows: 0,
      truncated: false,
    });
    expect(sent.map((request) => request.url)).toEqual([
      `https://graph.microsoft.com/v1.0/${REPORT}`,
      LINK,
    ]);
    const [graph, download] = sent;
    expect(graph?.headers.get("Authorization")).toBe(`Bearer ${TOKEN}`);
    expect(graph?.init?.redirect).toBe("manual");
    // The download carries Accept and nothing else, refuses redirects and sends no cookies.
    expect([...(download?.headers.keys() ?? [])]).toEqual(["accept"]);
    for (const [, value] of download?.headers ?? []) expect(value).not.toContain(TOKEN);
    expect(download?.init?.redirect).toBe("error");
    expect(download?.init?.credentials).toBe("omit");
    expect(download?.init?.method).toBe("GET");
    expect(download?.init?.signal).toBeInstanceOf(AbortSignal);
    expect(Object.keys(download?.init ?? {}).sort()).toEqual([
      "credentials",
      "headers",
      "method",
      "redirect",
      "signal",
    ]);
  });

  it("accepts the default port written out", async () => {
    const { sent, fetchImplementation } = fakeFetch(() =>
      redirectTo(`https://reports.office.com:443/data/download/${SECRET}`),
    );
    await reader(fetchImplementation).getReport(REPORT, usage);
    expect(sent[1]?.url).toBe(LINK);
  });

  it("refuses a download link outside the allowed location before any download", async () => {
    const refused: [string, string][] = [
      [`https://reports.example/data/download/${SECRET}`, "reports.example"],
      [`https://reports.office.com.example/data/download/${SECRET}`, "reports.office.com.example"],
      [`https://eu.reports.office.com/data/download/${SECRET}`, "eu.reports.office.com"],
      [`https://reports.office.com./data/download/${SECRET}`, "reports.office.com."],
      [`https://graph.microsoft.com/data/download/${SECRET}`, "graph.microsoft.com"],
      [`http://reports.office.com/data/download/${SECRET}`, "reports.office.com"],
      [`https://user:pass@reports.office.com/data/download/${SECRET}`, "reports.office.com"],
      [`https://user@reports.office.com/data/download/${SECRET}`, "reports.office.com"],
      [`https://reports.office.com:8443/data/download/${SECRET}`, "reports.office.com"],
      [`https://reports.office.com/data/other/${SECRET}`, "reports.office.com"],
      [`https://reports.office.com/download/${SECRET}`, "reports.office.com"],
      [`https://reports.office.com/data/DOWNLOAD/${SECRET}`, "reports.office.com"],
      ["https://reports.office.com/data/download/", "reports.office.com"],
      [`https://reports.office.com/data/download/../${SECRET}`, "reports.office.com"],
      [`https://reports.office.com/data/download/${SECRET}#part`, "reports.office.com"],
      [`/data/download/${SECRET}`, "none"],
      [`//reports.office.com/data/download/${SECRET}`, "none"],
      [`data/download/${SECRET}`, "none"],
      [`ftp://reports.office.com/data/download/${SECRET}`, "reports.office.com"],
      ["", "none"],
      [`https://reports.office.com/data/download/${"x".repeat(17_000)}`, "none"],
    ];
    for (const [location, host] of refused) {
      const { sent, fetchImplementation } = fakeFetch(() => redirectTo(location));
      const error = await failure(reader(fetchImplementation).getReport(REPORT, usage));
      expect(error, location).toBeInstanceOf(GraphBoundaryError);
      expect(error, location).toMatchObject({ reason: "report_host" });
      // The message names the host, so an operator can report it, and never the token.
      expect(String(error), location).toContain(`(host: ${host})`);
      expect(String(error), location).not.toContain(SECRET);
      expect(sent, location).toHaveLength(1);
    }
  });

  it("refuses a redirect with no link and any redirect other than 302", async () => {
    for (const status of [301, 303, 307, 308]) {
      const { sent, fetchImplementation } = fakeFetch(() => redirectTo(LINK, status));
      await expect(reader(fetchImplementation).getReport(REPORT, usage)).rejects.toMatchObject({
        reason: "redirect",
      });
      expect(sent).toHaveLength(1);
    }
    const { sent, fetchImplementation } = fakeFetch(() => new Response(null, { status: 302 }));
    await expect(reader(fetchImplementation).getReport(REPORT, usage)).rejects.toMatchObject({
      reason: "redirect",
    });
    expect(sent).toHaveLength(1);
  });

  it("follows no second redirect from the download host", async () => {
    // A real fetch with redirect "error" rejects on a redirect. This fake does the same, and
    // follows the redirect to a second host when asked to, so a weaker option would show.
    const hops: string[] = [];
    const following = fakeFetch(
      () => redirectTo(LINK),
      (_url, init) => {
        if (init?.redirect === "error") throw new TypeError("fetch failed");
        if (init?.redirect === "manual") return redirectTo("https://elsewhere.example/report.csv");
        hops.push("https://elsewhere.example/report.csv");
        return csvResponse(CSV);
      },
    );
    const error = await failure(reader(following.fetchImplementation).getReport(REPORT, usage));
    expect(error).toMatchObject({ name: "GraphBoundaryError", reason: "download" });
    expect(String(error)).not.toContain(SECRET);
    expect(following.sent).toHaveLength(2);
    expect(hops).toEqual([]);
    // A 3xx the fetch hands back anyway is a failure too, and nothing else is requested.
    const answered = fakeFetch(
      () => redirectTo(LINK),
      () => redirectTo("https://elsewhere.example/report.csv"),
    );
    const handedBack = await failure(reader(answered.fetchImplementation).getReport(REPORT, usage));
    expect(handedBack).toBeInstanceOf(GraphRequestError);
    expect(handedBack).toMatchObject({ status: 302 });
    expect(answered.sent).toHaveLength(2);
  });

  it("uses a download link once: a failed download is not retried", async () => {
    for (const status of [404, 429, 500, 503]) {
      const { sent, fetchImplementation } = fakeFetch(
        () => redirectTo(LINK),
        () =>
          new Response("<html>provider page with ada@lab.example</html>", {
            status,
            headers: { "Retry-After": "1", "request-id": "download-request-1" },
          }),
      );
      const error = await failure(reader(fetchImplementation).getReport(REPORT, usage));
      expect(error).toBeInstanceOf(GraphRequestError);
      expect(error).not.toBeInstanceOf(GraphRetryDeferredError);
      expect(error).toMatchObject({ status, requestId: undefined });
      expect(String(error)).not.toContain("ada@lab.example");
      expect(sent).toHaveLength(2);
    }
  });

  it("maps a download network failure to a fixed error without the link", async () => {
    const { sent, fetchImplementation } = fakeFetch(
      () => redirectTo(LINK),
      () => {
        throw new TypeError(`fetch failed for ${LINK}`);
      },
    );
    const error = await failure(reader(fetchImplementation).getReport(REPORT, usage));
    expect(error).toMatchObject({ name: "GraphBoundaryError", reason: "download" });
    expect(String(error)).not.toContain(SECRET);
    expect((error as Error).cause).toBeUndefined();
    expect(sent).toHaveLength(2);
  });

  it("stops a large download mid-stream at the byte cap", async () => {
    const large = largeBody();
    const { fetchImplementation } = fakeFetch(
      () => redirectTo(LINK),
      () => csvResponse(large.stream),
    );
    const error = await failure(
      reader(fetchImplementation).getReport(REPORT, usage, { maxBytes: 1_048_576 }),
    );
    expect(error).toMatchObject({ name: "GraphBoundaryError", reason: "size" });
    expect(large.state.pulls).toBeLessThanOrEqual(1_048_576 / 65_536 + 3);
    expect(large.state.cancelled).toBe(true);
  });

  it("caps a download at 16 MiB unless told otherwise, and at most 32 MiB", async () => {
    const large = largeBody(1_048_576, 40);
    const { fetchImplementation } = fakeFetch(
      () => redirectTo(LINK),
      () => csvResponse(large.stream),
    );
    await expect(reader(fetchImplementation).getReport(REPORT, usage)).rejects.toMatchObject({
      reason: "size",
    });
    expect(large.state.pulls).toBeGreaterThanOrEqual(16);
    expect(large.state.pulls).toBeLessThanOrEqual(19);
    const unused = fakeFetch(() => redirectTo(LINK));
    for (const maxBytes of [32 * 1_048_576 + 1, 0, 1_023, 1.5])
      await expect(
        reader(unused.fetchImplementation).getReport(REPORT, usage, { maxBytes }),
      ).rejects.toThrow("maxBytes");
    expect(unused.sent).toHaveLength(0);
  });

  it("refuses a download whose declared length is over the cap before reading it", async () => {
    const large = largeBody();
    const { fetchImplementation } = fakeFetch(
      () => redirectTo(LINK),
      () => csvResponse(large.stream, { "Content-Length": String(2 * 1_048_576) }),
    );
    await expect(
      reader(fetchImplementation).getReport(REPORT, usage, { maxBytes: 1_048_576 }),
    ).rejects.toMatchObject({ reason: "size" });
    expect(large.state.cancelled).toBe(true);
    expect(large.state.pulls).toBeLessThanOrEqual(1);
  });

  it("refuses a download that is not a file: JSON or an HTML page", async () => {
    for (const type of ["application/json", "text/html", "text/plain"]) {
      const { fetchImplementation } = fakeFetch(
        () => redirectTo(LINK),
        () => csvResponse(CSV, { "Content-Type": type }),
      );
      await expect(reader(fetchImplementation).getReport(REPORT, usage)).rejects.toMatchObject({
        reason: "content_type",
      });
    }
    const untyped = fakeFetch(
      () => redirectTo(LINK),
      () => new Response(new Blob([CSV]).stream(), { status: 200 }),
    );
    await expect(
      reader(untyped.fetchImplementation).getReport(REPORT, usage),
    ).rejects.toMatchObject({ reason: "content_type" });
    const csvType = fakeFetch(
      () => redirectTo(LINK),
      () => csvResponse(CSV, { "Content-Type": "text/csv; charset=utf-8" }),
    );
    await expect(
      reader(csvType.fetchImplementation).getReport(REPORT, usage),
    ).resolves.toMatchObject({ rows: [{ "User Principal Name": "ada@lab.example" }] });
  });

  it("reports a download that is not valid CSV with a fixed reason", async () => {
    const { fetchImplementation } = fakeFetch(
      () => redirectTo(LINK),
      () => csvResponse("Name,Name\nada@lab.example,x\n"),
    );
    const error = await failure(reader(fetchImplementation).getReport(REPORT, usage));
    expect(error).toMatchObject({ name: "GraphBoundaryError", reason: "csv" });
    expect(String(error)).toContain("duplicate_header");
    expect(String(error)).not.toContain("ada@lab.example");
  });

  it("parses a CSV that Graph answers itself, with the same cap", async () => {
    const { sent, fetchImplementation } = fakeFetch(() => csvResponse(CSV));
    await expect(reader(fetchImplementation).getReport(REPORT, usage)).resolves.toMatchObject({
      rows: [{ "User Principal Name": "ada@lab.example" }],
    });
    expect(sent).toHaveLength(1);
    const body = largeBody();
    const large = fakeFetch(() => csvResponse(body.stream));
    await expect(
      reader(large.fetchImplementation).getReport(REPORT, usage, { maxBytes: 1_048_576 }),
    ).rejects.toMatchObject({ reason: "size" });
    expect(body.state.cancelled).toBe(true);
    // JSON from Graph is not a report download; request() is the JSON reader.
    const json = fakeFetch(() => Response.json({ value: [] }));
    await expect(reader(json.fetchImplementation).getReport(REPORT, usage)).rejects.toMatchObject({
      reason: "content_type",
    });
    // Any other success status carries no report.
    const empty = fakeFetch(() => new Response(null, { status: 204 }));
    await expect(reader(empty.fetchImplementation).getReport(REPORT, usage)).rejects.toMatchObject({
      reason: "shape",
    });
  });
});

describe("getReport: time limits, cancellation and throttling", () => {
  it("ends a stalled Graph request within the timeout, before any download", async () => {
    const aborted = { value: false };
    const { sent, fetchImplementation } = fakeFetch(stalled(aborted));
    await expect(
      reader(fetchImplementation, { requestTimeoutMilliseconds: 100 }).getReport(REPORT, usage),
    ).rejects.toMatchObject({ reason: "timeout" });
    expect(aborted.value).toBe(true);
    expect(sent).toHaveLength(1);
  });

  it("ends a stalled download, and a download body that stops arriving, within the timeout", async () => {
    const aborted = { value: false };
    const stall = fakeFetch(() => redirectTo(LINK), stalled(aborted));
    await expect(
      reader(stall.fetchImplementation, { requestTimeoutMilliseconds: 100 }).getReport(
        REPORT,
        usage,
      ),
    ).rejects.toMatchObject({ reason: "timeout" });
    expect(aborted.value).toBe(true);
    expect(stall.sent).toHaveLength(2);

    let downloadSignal: AbortSignal | undefined;
    const trickle = fakeFetch(
      () => redirectTo(LINK),
      (_url, init) => {
        downloadSignal = init?.signal ?? undefined;
        return csvResponse(
          new ReadableStream<Uint8Array>({
            start(controller) {
              controller.enqueue(new TextEncoder().encode("Name\n"));
            },
          }),
        );
      },
    );
    await expect(
      reader(trickle.fetchImplementation, { requestTimeoutMilliseconds: 100 }).getReport(
        REPORT,
        usage,
      ),
    ).rejects.toMatchObject({ reason: "timeout" });
    expect(downloadSignal?.aborted).toBe(true);
  });

  it("stops the Graph request when the caller's signal aborts", async () => {
    const aborted = { value: false };
    const controller = new AbortController();
    const { sent, fetchImplementation } = fakeFetch(stalled(aborted));
    const pending = reader(fetchImplementation, { requestTimeoutMilliseconds: 60_000 }).getReport(
      REPORT,
      usage,
      { signal: controller.signal },
    );
    setTimeout(() => controller.abort(new Error("audit stopped")), 20);
    await expect(pending).rejects.toThrow("audit stopped");
    expect(aborted.value).toBe(true);
    expect(sent).toHaveLength(1);
  });

  it("stops the download when the caller's signal aborts", async () => {
    const aborted = { value: false };
    const controller = new AbortController();
    const { sent, fetchImplementation } = fakeFetch(() => redirectTo(LINK), stalled(aborted));
    const started = Date.now();
    const pending = reader(fetchImplementation, { requestTimeoutMilliseconds: 60_000 }).getReport(
      REPORT,
      usage,
      { signal: controller.signal },
    );
    setTimeout(() => controller.abort(new Error("audit stopped")), 20);
    await expect(pending).rejects.toThrow("audit stopped");
    expect(aborted.value).toBe(true);
    expect(sent).toHaveLength(2);
    expect(Date.now() - started).toBeLessThan(5_000);
  });

  it("sends nothing when the caller's signal has already aborted", async () => {
    const controller = new AbortController();
    controller.abort(new Error("already stopped"));
    const { sent, fetchImplementation } = fakeFetch(() => redirectTo(LINK));
    await expect(
      reader(fetchImplementation).getReport(REPORT, usage, { signal: controller.signal }),
    ).rejects.toThrow("already stopped");
    expect(sent).toHaveLength(0);
  });

  it("does not start the download when the signal aborts while Graph answers", async () => {
    const controller = new AbortController();
    const { sent, fetchImplementation } = fakeFetch(() => {
      controller.abort(new Error("stopped between requests"));
      return redirectTo(LINK);
    });
    await expect(
      reader(fetchImplementation).getReport(REPORT, usage, { signal: controller.signal }),
    ).rejects.toThrow("stopped between requests");
    expect(sent).toHaveLength(1);
  });

  it("retries a throttled Graph request after Retry-After within the budget", async () => {
    const sleep = vi.fn(async () => undefined);
    let graphCalls = 0;
    const { sent, fetchImplementation } = fakeFetch(() => {
      graphCalls += 1;
      return graphCalls === 1
        ? new Response(null, { status: 429, headers: { "Retry-After": "2" } })
        : redirectTo(LINK);
    });
    await expect(
      reader(fetchImplementation, { sleep }).getReport(REPORT, usage),
    ).resolves.toMatchObject({ rows: [{ "User Principal Name": "ada@lab.example" }] });
    expect(sleep).toHaveBeenCalledWith(2_000);
    expect(sent.map((request) => new URL(request.url).hostname)).toEqual([
      "graph.microsoft.com",
      "graph.microsoft.com",
      "reports.office.com",
    ]);
  });

  it("defers a Retry-After beyond the budget without waiting or downloading", async () => {
    const sleep = vi.fn(async () => undefined);
    const { sent, fetchImplementation } = fakeFetch(
      () => new Response(null, { status: 429, headers: { "Retry-After": "600" } }),
    );
    const error = await failure(reader(fetchImplementation, { sleep }).getReport(REPORT, usage));
    expect(error).toBeInstanceOf(GraphRetryDeferredError);
    expect(error).toMatchObject({ retryAfterMilliseconds: 600_000 });
    expect(sleep).not.toHaveBeenCalled();
    expect(sent).toHaveLength(1);
  });

  it("keeps a 403 from Graph a GraphRequestError with no provider text", async () => {
    const { sent, fetchImplementation } = fakeFetch(
      () =>
        new Response('{"error":{"message":"ada@lab.example lacks Reports Reader"}}', {
          status: 403,
          headers: { "Content-Type": "application/json", "request-id": "graph-request-1" },
        }),
    );
    const error = await failure(reader(fetchImplementation).getReport(REPORT, usage));
    expect(error).toBeInstanceOf(GraphRequestError);
    expect(error).toMatchObject({ status: 403, requestId: "graph-request-1" });
    expect(String(error)).not.toContain("ada@lab.example");
    expect(sent).toHaveLength(1);
  });
});

describe("getReport: the checks before anything is sent", () => {
  it("refuses every pack but the usage reports pack", async () => {
    const others = graphPacks.filter((pack) => pack.id !== "reports.usage.read");
    expect(others.length).toBeGreaterThan(10);
    for (const pack of others) {
      const { sent, fetchImplementation } = fakeFetch(() => redirectTo(LINK));
      const context = { feature: pack.id, permission: pack.permission } as never;
      await expect(
        reader(fetchImplementation).getReport(REPORT, context),
        pack.id,
      ).rejects.toMatchObject({ name: "GraphBoundaryError", reason: "report" });
      expect(sent, pack.id).toHaveLength(0);
    }
  });

  it("runs the normal checks: permission, cloud, host, path, query options", async () => {
    const { sent, fetchImplementation } = fakeFetch(() => redirectTo(LINK));
    const client = reader(fetchImplementation);
    const highHigh = reader(fetchImplementation, { cloud: "usGovernmentL4" });
    const wrongPermission = { feature: "reports.usage.read", permission: "User.Read.All" };
    const cases: [GraphHttpClient, string, unknown, string][] = [
      [client, REPORT, wrongPermission, "pack"],
      [highHigh, REPORT, usage, "cloud"],
      [client, `https://reports.office.com/v1.0/${REPORT}`, usage, "host"],
      [client, `https://graph.microsoft.com/beta/${REPORT}`, usage, "host"],
      [client, "reports/getOffice365ActiveUserCounts(period='D30')", usage, "endpoint"],
      [client, `${REPORT}?$expand=x`, usage, "endpoint"],
      [client, `${REPORT}?$format=application/json`, usage, "endpoint"],
      [client, `${REPORT}?$top=10`, usage, "endpoint"],
    ];
    for (const [candidate, path, context, reason] of cases)
      await expect(candidate.getReport(path, context as never), path).rejects.toMatchObject({
        name: "GraphBoundaryError",
        reason,
      });
    expect(sent).toHaveLength(0);
  });
});

// One allowed path for each Graph pack. The record type makes a new pack fail to compile here
// until it has one, so the redirect check below always covers every pack.
const samplePaths: Record<GraphFeatureId, string> = {
  profile: "me",
  "directory.users.read": "users",
  "user.licenses.read": "users/u-1/licenseDetails",
  "tenant.licenses.read": "subscribedSkus",
  "service.health.read": "admin/serviceAnnouncement/issues",
  "service.messages.read": "admin/serviceAnnouncement/messages",
  "audit.signins.read": "auditLogs/signIns",
  "audit.directory.read": "auditLogs/directoryAudits",
  "reports.authentication-methods.read": "reports/authenticationMethods/userRegistrationDetails",
  "identity.conditional-access.read": "identity/conditionalAccess/policies",
  "directory.applications.read": "servicePrincipals",
  "directory.permission-grants.read": "oauth2PermissionGrants",
  "identity.risk.read": "identityProtection/riskyUsers",
  "reports.usage.read": REPORT,
  "reports.settings.read": "admin/reportSettings",
  "security.score.read": "security/secureScores",
  "security.incidents.read": "security/incidents",
  "devices.intune.read": "deviceManagement/managedDevices",
  "purview.entra.read": "security/auditLog/queries",
};

describe("request() and getCollection() still refuse every redirect", () => {
  it("refuses a 302 to the report download host for every pack, including usage reports", async () => {
    expect(Object.keys(samplePaths).sort()).toEqual(graphPacks.map((pack) => pack.id).sort());
    for (const pack of graphPacks) {
      const context = { feature: pack.id, permission: pack.permission } as never;
      for (const status of [301, 302, 303, 307, 308]) {
        const { sent, fetchImplementation } = fakeFetch(() => redirectTo(LINK, status));
        await expect(
          reader(fetchImplementation).request(samplePaths[pack.id], context),
          `${pack.id} ${status}`,
        ).rejects.toMatchObject({ reason: "redirect" });
        expect(sent, `${pack.id} ${status}`).toHaveLength(1);
      }
      const { sent, fetchImplementation } = fakeFetch(() => redirectTo(LINK));
      await expect(
        reader(fetchImplementation).getCollection(samplePaths[pack.id], context),
        pack.id,
      ).rejects.toMatchObject({ reason: "redirect" });
      expect(sent, pack.id).toHaveLength(1);
    }
    // The JSON form of the apps report, the one usage report request() reads.
    const { sent, fetchImplementation } = fakeFetch(() => redirectTo(LINK));
    await expect(
      reader(fetchImplementation).request(
        "reports/getM365AppUserDetail(period='D90')?$format=application/json",
        usage,
      ),
    ).rejects.toMatchObject({ reason: "redirect" });
    expect(sent).toHaveLength(1);
  });
});

describe("the usage report and membership paths", () => {
  const allowed = async (path: string, feature: string, permission: string) => {
    const { fetchImplementation } = fakeFetch(() => Response.json({ value: [] }));
    return reader(fetchImplementation)
      .request(path, { feature, permission } as never)
      .then(
        () => true,
        (error: unknown) => {
          if (error instanceof GraphBoundaryError && error.reason === "endpoint") return false;
          throw error;
        },
      );
  };
  const report = (path: string) =>
    allowed(path, ...(["reports.usage.read", "Reports.Read.All"] as const));
  const users = (path: string) =>
    allowed(path, ...(["directory.users.read", "User.Read.All"] as const));

  const reports = [
    "getM365AppUserDetail",
    "getOffice365ActiveUserDetail",
    "getEmailActivityUserDetail",
    "getTeamsUserActivityUserDetail",
    "getOneDriveActivityUserDetail",
    "getOneDriveUsageAccountDetail",
    "getSharePointActivityUserDetail",
    "getMailboxUsageDetail",
  ];

  it("allows each listed report for each period", async () => {
    for (const name of reports)
      for (const period of ["D7", "D30", "D90", "D180"])
        expect(await report(`reports/${name}(period='${period}')`), `${name} ${period}`).toBe(true);
  });

  it("refuses the date= form, other periods, other reports and trailing segments", async () => {
    for (const name of reports)
      for (const path of [
        `reports/${name}(date=2026-10-01)`,
        `reports/${name}(date='2026-10-01')`,
        `reports/${name}(period='D1')`,
        `reports/${name}(period='D365')`,
        `reports/${name}(period='D07')`,
        `reports/${name}(period=D7)`,
        `reports/${name}(period="D7")`,
        `reports/${name}(period=%27D7%27)`,
        `reports/${name}(period='D7',date='2026-10-01')`,
        `reports/${name}(period='D7')/extra`,
        `reports/${name}(period='D7')/$value`,
        `reports/${name}`,
        `reports/x/${name}(period='D7')`,
      ])
        expect(await report(path), path).toBe(false);
    for (const name of [
      "getOffice365ActiveUserCounts",
      "getM365AppUserCounts",
      "getEmailActivityCounts",
      "getMailboxUsageStorage",
      "getTeamsDeviceUsageUserDetail",
      "getSharePointSiteUsageDetail",
      "getOneDriveUsageFileCounts",
      "getYammerActivityUserDetail",
      "getSkypeForBusinessActivityUserDetail",
    ])
      expect(await report(`reports/${name}(period='D7')`), name).toBe(false);
  });

  it("allows a user's direct memberships by id or by principal name", async () => {
    for (const path of [
      "users/6f1c2d8e-4b7a-4e35-9a0c-3d5b7e91c204/memberOf",
      "users/ada%40lab.example/memberOf",
      "users('ada%40lab.example')/memberOf",
      "users('o''hara%40lab.example')/memberOf",
      "users('guest_partner.example%23EXT%23%40lab.example')/memberOf",
      "users/6f1c2d8e-4b7a-4e35-9a0c-3d5b7e91c204/memberOf?$select=id,displayName",
      "users/6f1c2d8e-4b7a-4e35-9a0c-3d5b7e91c204/memberOf/",
    ])
      expect(await users(path), path).toBe(true);
  });

  it("refuses transitive memberships, deeper paths and encoded separators", async () => {
    for (const path of [
      "users/u-1/transitiveMemberOf",
      "users('ada%40lab.example')/transitiveMemberOf",
      "users/u-1/memberOf/extra",
      "users/u-1/memberOf/microsoft.graph.group",
      "users/u-1/memberOf/$count",
      "users/u-1/manager/memberOf",
      "users/u%2F..%2Fgroups/memberOf",
      "users/u%2f/memberOf",
      "users/u%5C/memberOf",
      "users/u%3F/memberOf",
      "users/u%zz/memberOf",
      "users/u%2/memberOf",
      "users/u+1/memberOf",
      "users/u!1/memberOf",
      "groups/g-1/memberOf",
      "users/u-1/ownedObjects",
    ])
      expect(await users(path), path).toBe(false);
    // The paths the user pack already had stay as they were.
    for (const path of ["users", "users/u-1", "users('ada%40lab.example')", "users/u-1/manager"])
      expect(await users(path), path).toBe(true);
  });
});

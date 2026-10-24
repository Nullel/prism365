import { z } from "zod";
import { getMicrosoftCloudProfile, type MicrosoftCloud } from "./clouds.ts";
import { CsvFormatError, type CsvTable, parseCsv } from "./csv.ts";
import { type GraphFeatureId, type GraphPermission, getPermissionPack } from "./permissions.ts";

type FetchImplementation = (input: RequestInfo | URL, init?: RequestInit) => Promise<Response>;
type Sleep = (milliseconds: number) => Promise<void>;

export interface GraphRequestContext {
  feature: GraphFeatureId;
  permission: GraphPermission;
}

export interface GraphClientOptions {
  getAccessToken: () => Promise<string>;
  fetchImplementation?: FetchImplementation;
  sleep?: Sleep;
  maxRetries?: number;
  cloud?: MicrosoftCloud;
  requestTimeoutMilliseconds?: number;
  maxResponseBytes?: number;
  maxRetryWaitMilliseconds?: number;
}

export interface CollectionOptions {
  maxPages?: number;
  maxRecords?: number;
  // Epoch milliseconds after which no further page is requested.
  deadline?: number;
  // Stops paging, and the request in flight, when aborted.
  signal?: AbortSignal;
}

export interface ReportOptions {
  // Stops the Graph request and the download, including one in flight, when aborted.
  signal?: AbortSignal;
  // The report's byte cap: 16 MiB unless set, at most 32 MiB.
  maxBytes?: number;
}

export interface CollectionResult<T> {
  items: T[];
  // false when more records exist than were read: a record or page cap, or the deadline.
  complete: boolean;
  // Set when paging stopped because the caller's deadline passed.
  stopped?: "deadline";
}

export class GraphRequestError extends Error {
  readonly status: number;
  readonly requestId: string | undefined;
  constructor(status: number, requestId?: string) {
    super(`Microsoft Graph request failed with HTTP ${status}.`);
    this.name = "GraphRequestError";
    this.status = status;
    this.requestId = requestId && /^[\w-]{1,128}$/.test(requestId) ? requestId : undefined;
  }
}

export class GraphRetryDeferredError extends GraphRequestError {
  readonly retryAfterMilliseconds: number;
  constructor(status: number, delay: number, requestId?: string) {
    super(status, requestId);
    this.name = "GraphRetryDeferredError";
    this.retryAfterMilliseconds = delay;
  }
}

// Fixed reason codes, safe to log: they never carry a tenant value or provider message.
export type GraphBoundaryReason =
  | "configuration"
  | "pack"
  | "cloud"
  | "method"
  | "endpoint"
  | "credential"
  | "redirect"
  | "timeout"
  | "shape"
  | "retry_budget"
  | "pagination"
  | "continuation"
  | "item"
  | "page_limit"
  | "host"
  | "content_type"
  | "size"
  | "empty"
  | "json"
  | "report"
  | "report_host"
  | "download"
  | "csv";

export class GraphBoundaryError extends Error {
  readonly reason: GraphBoundaryReason;
  constructor(message: string, reason: GraphBoundaryReason = "configuration") {
    super(message);
    this.name = "GraphBoundaryError";
    this.reason = reason;
  }
}

// A usage report download link refused by getReport (ADR-0010). It carries the link's host so an
// audit can show it and the log can record it, and an operator can report a host Microsoft moved
// to. Never the path or query: they carry the pre-authenticated token. A parsed hostname is
// printable ASCII (the URL parser refuses control and space characters in hosts and converts other
// names to punycode), and it is cut to the longest DNS name. `problem` is fixed text from the check.
export class GraphReportHostError extends GraphBoundaryError {
  readonly host: string | null;
  // Whether the host is a listed download host, so the problem is the link's form, not its host.
  readonly knownHost: boolean;
  readonly problem: string;
  constructor(problem: string, host: string | null, knownHost: boolean) {
    const allowed = reportDownloadHosts.map((name) => `https://${name}${REPORT_DOWNLOAD_PATH}`);
    super(
      `Microsoft's usage report download link was refused because ${problem} (host: ${host ?? "none"}). Prism365 downloads reports only from ${allowed.join(" or ")}. If Microsoft has moved report downloads, report this host so the list can be reviewed.`,
      "report_host",
    );
    this.name = "GraphReportHostError";
    this.host = host === null ? null : host.slice(0, 253);
    this.knownHost = knownHost;
    this.problem = problem;
  }
}

// A usage report download that the download host answered with anything but 200. It is not a
// GraphRequestError: a 401 or 403 here means the short-lived link failed, not that the session or a
// role is missing, and it keeps only the status, never the link or the host's text.
export class GraphReportDownloadError extends Error {
  readonly status: number;
  constructor(status: number) {
    super(`Microsoft's usage report download failed with HTTP ${status}.`);
    this.name = "GraphReportDownloadError";
    this.status = status;
  }
}

const collectionSchema = z.object({
  value: z.array(z.unknown()).max(1_000),
  "@odata.nextLink": z.string().max(16_384).optional(),
});

function boundedInteger(value: number, min: number, max: number, label: string): number {
  if (!Number.isInteger(value) || value < min || value > max) {
    throw new GraphBoundaryError(`${label} must be an integer between ${min} and ${max}.`);
  }
  return value;
}

// Paths and declared features must agree even if the supplied token has broader consent.
// Matching ignores case: Graph treats path segments case-insensitively and its continuation
// links do not always keep the casing of the request (auditLogs/signIns comes back as signins).
const allowedPaths: Record<GraphFeatureId, RegExp> = {
  profile: /^\/v1\.0\/me\/?$/i,
  // The user list, one user, one user's manager, or one user's direct memberships (groups,
  // directory roles and administrative units; not transitiveMemberOf), in the /users/{id} or the
  // users('upn') form. The key takes only what the path builders (person.ts, investigation.ts)
  // produce from an id or a principal name: what encodeURIComponent leaves as it is (letters,
  // digits and - _ . ! ~ * ' ( )), the : and @ of raw ids, and well-formed percent escapes other
  // than an encoded slash, backslash or question mark, which no id or principal name holds.
  "directory.users.read":
    /^\/v1\.0\/users(?:(?:\/(?:[A-Za-z0-9._:@!~*'()-]|%(?!2f|5c|3f)[0-9a-f]{2})+|\((?:[A-Za-z0-9._:@!~*'()-]|%(?!2f|5c|3f)[0-9a-f]{2})+\))(?:\/manager|\/memberOf)?)?\/?$/i,
  "user.licenses.read": /^\/v1\.0\/users(?:\/[^/]+|\([^/]+\))\/licenseDetails\/?$/i,
  "tenant.licenses.read": /^\/v1\.0\/subscribedSkus\/?$/i,
  "service.health.read":
    /^\/v1\.0\/admin\/serviceAnnouncement\/(?:healthOverviews|issues)(?:\/[^/]+)?\/?$/i,
  "service.messages.read": /^\/v1\.0\/admin\/serviceAnnouncement\/messages(?:\/[^/]+)?\/?$/i,
  "audit.signins.read": /^\/v1\.0\/auditLogs\/signIns(?:\/[^/]+)?\/?$/i,
  "audit.directory.read": /^\/v1\.0\/auditLogs\/directoryAudits(?:\/[A-Za-z0-9._:-]+)?\/?$/i,
  "reports.authentication-methods.read":
    /^\/v1\.0\/reports\/authenticationMethods\/userRegistrationDetails(?:\/[^/]+)?\/?$/i,
  "identity.conditional-access.read":
    /^\/v1\.0\/identity\/conditionalAccess\/(?:policies|namedLocations)(?:\/[A-Za-z0-9._:-]+)?\/?$/i,
  // Enterprise applications, one of them, or the application permissions granted to one.
  "directory.applications.read":
    /^\/v1\.0\/servicePrincipals(?:\/[A-Za-z0-9._:-]+(?:\/appRoleAssignments)?)?\/?$/i,
  "directory.permission-grants.read": /^\/v1\.0\/oauth2PermissionGrants\/?$/i,
  // The four role lists, never one item or a navigation: definitions, active assignments, and the
  // Privileged Identity Management instances (active and eligible) behind them.
  "directory.roles.read":
    /^\/v1\.0\/roleManagement\/directory\/(?:roleDefinitions|roleAssignments|roleAssignmentScheduleInstances|roleEligibilityScheduleInstances)\/?$/i,
  // The group list, or one group's owners by its object ID (a GUID). Members are left out.
  "directory.groups.read":
    /^\/v1\.0\/groups(?:\/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\/owners)?\/?$/i,
  "directory.domains.read": /^\/v1\.0\/domains\/?$/i,
  "identity.risk.read": /^\/v1\.0\/identityProtection\/riskyUsers(?:\/[^/]+)?\/?$/i,
  // The per-user usage reports and the per-team Teams activity report, by period only (no date=
  // form). getM365AppUserDetail also answers in JSON through request(); every one of them answers
  // with a redirect to a CSV download, which only getReport() follows (ADR-0010).
  "reports.usage.read":
    /^\/v1\.0\/reports\/get(?:M365AppUserDetail|Office365ActiveUserDetail|EmailActivityUserDetail|TeamsUserActivityUserDetail|OneDriveActivityUserDetail|OneDriveUsageAccountDetail|SharePointActivityUserDetail|MailboxUsageDetail|TeamsTeamActivityDetail)\(period='D(?:7|30|90|180)'\)\/?$/i,
  "reports.settings.read": /^\/v1\.0\/admin\/reportSettings\/?$/i,
  "security.score.read":
    /^\/v1\.0\/security\/(?:secureScores|secureScoreControlProfiles)(?:\/[^/]+)?\/?$/i,
  "security.incidents.read": /^\/v1\.0\/security\/incidents(?:\/[^/]+)?\/?$/i,
  "devices.intune.read": /^\/v1\.0\/deviceManagement\/managedDevices(?:\/[^/]+)?\/?$/i,
  "purview.entra.read": /^\/v1\.0\/security\/auditLog\/queries(?:\/[^/]+(?:\/records)?)?\/?$/i,
};

// Usage report downloads (ADR-0010). Microsoft answers a usage report with 302 Found and a
// short-lived, pre-authenticated link; following that link is the one redirect this client ever
// makes, and only to a listed host, over HTTPS on the default port, under this path.
const REPORT_FEATURE: GraphFeatureId = "reports.usage.read";
const reportDownloadHosts: readonly string[] = ["reports.office.com"];
const REPORT_DOWNLOAD_PATH = "/data/download/";
const REPORT_DEFAULT_BYTES = 16 * 1_048_576;
const REPORT_MAX_BYTES = 32 * 1_048_576;
// Microsoft serves the download as application/octet-stream; text/csv is accepted in case a report
// is ever answered directly. Anything else, JSON or an HTML error page, is refused.
const reportContentTypes: ReadonlySet<string> = new Set(["application/octet-stream", "text/csv"]);
// What Microsoft's own SDK sends for these stream-returning functions: the file, or a JSON error.
const REPORT_ACCEPT = "application/octet-stream, application/json";
const DOWNLOAD_ACCEPT = "application/octet-stream, text/csv";

export class GraphHttpClient {
  readonly #getAccessToken: () => Promise<string>;
  readonly #fetch: FetchImplementation;
  readonly #sleep: Sleep;
  readonly #maxRetries: number;
  readonly #timeout: number;
  readonly #maxResponseBytes: number;
  readonly #retryBudget: number;
  readonly #baseUrl: URL;
  readonly #cloud: MicrosoftCloud;

  constructor(options: GraphClientOptions) {
    this.#getAccessToken = options.getAccessToken;
    this.#fetch = options.fetchImplementation ?? globalThis.fetch.bind(globalThis);
    this.#sleep = options.sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
    this.#maxRetries = boundedInteger(options.maxRetries ?? 2, 0, 5, "maxRetries");
    this.#timeout = boundedInteger(
      options.requestTimeoutMilliseconds ?? 10_000,
      100,
      // Sign-in log queries can take well over 30 seconds; background readers need the room.
      120_000,
      "requestTimeoutMilliseconds",
    );
    this.#maxResponseBytes = boundedInteger(
      options.maxResponseBytes ?? 1_048_576,
      128,
      8_388_608,
      "maxResponseBytes",
    );
    this.#retryBudget = boundedInteger(
      options.maxRetryWaitMilliseconds ?? 5_000,
      0,
      30_000,
      "maxRetryWaitMilliseconds",
    );
    this.#cloud = options.cloud ?? "worldwide";
    this.#baseUrl = new URL(getMicrosoftCloudProfile(this.#cloud).graphBaseUrl);
  }

  request(
    pathOrNextLink: string,
    context: GraphRequestContext,
    init?: RequestInit,
  ): Promise<unknown>;
  request<T>(
    pathOrNextLink: string,
    context: GraphRequestContext,
    init: RequestInit,
    schema: z.ZodType<T>,
  ): Promise<T>;
  async request(
    pathOrNextLink: string,
    context: GraphRequestContext,
    init: RequestInit = {},
    schema?: z.ZodType,
  ): Promise<unknown> {
    const method = (init.method ?? "GET").toUpperCase();
    const target = this.#checkedTarget(pathOrNextLink, context, method, init.body);
    // A caller's signal (an audit's time limit, a collector run ending) stops the request and
    // any retry; the request's own timeout still applies.
    const { signal: external, ...requestInit } = init;
    const value = await this.#exchange(
      target,
      method,
      requestInit,
      external,
      "application/json",
      async (response) => {
        if (response.status >= 300) {
          await response.body?.cancel();
          throw new GraphBoundaryError(
            "Graph redirects are disabled; this resource needs a reviewed download adapter.",
            "redirect",
          );
        }
        if (response.status === 204 || method === "HEAD") {
          await response.body?.cancel();
          return undefined;
        }
        return readBoundedJson(response, this.#maxResponseBytes);
      },
    );
    if (!schema) return value;
    const parsed = schema.safeParse(value);
    if (!parsed.success)
      throw new GraphBoundaryError("Microsoft Graph returned an invalid response shape.", "shape");
    return parsed.data;
  }

  // Reads one Microsoft 365 usage report (ADR-0010): the one place this client follows a
  // redirect. The Graph request passes every check request() applies and carries the token;
  // Microsoft answers 302 with a short-lived, pre-authenticated link, which must be an HTTPS URL
  // on a listed host under /data/download/, and the download is sent with no token, no cookies,
  // no header but Accept, no redirects and one attempt. A CSV answered by Graph itself (200) is
  // read the same way. Either is capped in bytes while it streams, then parsed by parseCsv.
  async getReport(
    path: string,
    context: GraphRequestContext,
    options: ReportOptions = {},
  ): Promise<CsvTable> {
    if (context.feature !== REPORT_FEATURE)
      throw new GraphBoundaryError(
        "Report downloads are only for the Microsoft 365 usage reports pack.",
        "report",
      );
    const maxBytes = boundedInteger(
      options.maxBytes ?? REPORT_DEFAULT_BYTES,
      1_024,
      REPORT_MAX_BYTES,
      "maxBytes",
    );
    const target = this.#checkedTarget(path, context, "GET", undefined);
    // A report is named by its function and period alone. A query option could change what
    // Microsoft answers (JSON instead of a download, for one report), and none is needed.
    if (target.search)
      throw new GraphBoundaryError("Usage report requests take no query options.", "endpoint");
    const { signal } = options;
    const answer = await this.#exchange(
      target,
      "GET",
      {},
      signal,
      REPORT_ACCEPT,
      async (response): Promise<{ download: URL } | { text: string }> => {
        if (response.status === 302) {
          const location = response.headers.get("Location");
          await response.body?.cancel();
          if (location === null)
            throw new GraphBoundaryError(
              "Microsoft Graph answered a usage report with a redirect but no download link.",
              "redirect",
            );
          return { download: reportDownloadUrl(location) };
        }
        if (response.status !== 200) {
          await response.body?.cancel();
          throw response.status >= 300
            ? new GraphBoundaryError(
                "Microsoft Graph answered a usage report with a redirect other than 302.",
                "redirect",
              )
            : new GraphBoundaryError(
                "Microsoft Graph answered a usage report with no report in it.",
                "shape",
              );
        }
        return { text: await readBoundedReport(response, maxBytes) };
      },
    );
    const text =
      "text" in answer ? answer.text : await this.#download(answer.download, maxBytes, signal);
    try {
      return parseCsv(text);
    } catch (error) {
      if (error instanceof CsvFormatError)
        throw new GraphBoundaryError(
          `Microsoft's usage report could not be read as CSV (${error.problem}).`,
          "csv",
        );
      throw error;
    }
  }

  getCollection(
    path: string,
    context: GraphRequestContext,
    options?: CollectionOptions,
  ): Promise<CollectionResult<unknown>>;
  getCollection<T>(
    path: string,
    context: GraphRequestContext,
    options: CollectionOptions & { schema: z.ZodType<T> },
  ): Promise<CollectionResult<T>>;
  async getCollection(
    path: string,
    context: GraphRequestContext,
    options: CollectionOptions & { schema?: z.ZodType } = {},
  ): Promise<CollectionResult<unknown>> {
    const maxPages = boundedInteger(options.maxPages ?? 100, 1, 100, "maxPages");
    const maxRecords = boundedInteger(options.maxRecords ?? 10_000, 1, 10_000, "maxRecords");
    const items: unknown[] = [];
    const seen = new Set<string>();
    let next: string | undefined = path;
    for (let pageNumber = 0; next && pageNumber < maxPages; pageNumber += 1) {
      options.signal?.throwIfAborted();
      // A page is not started past the caller's deadline, so one slow collection cannot carry a
      // run past the point where it can still save its place. What was read is returned, marked
      // as stopped, so the caller can keep it.
      if (pageNumber > 0 && options.deadline !== undefined && Date.now() > options.deadline)
        return { items, complete: false, stopped: "deadline" };
      const target = this.#resolveTarget(next).href;
      if (seen.has(target))
        throw new GraphBoundaryError(
          "Graph pagination repeated a continuation link.",
          "pagination",
        );
      seen.add(target);
      const page: z.infer<typeof collectionSchema> = await this.request(
        next,
        context,
        options.signal ? { signal: options.signal } : {},
        collectionSchema,
      );
      const continuationLink = page["@odata.nextLink"];
      if (continuationLink) {
        const continuation = this.#resolveTarget(continuationLink);
        if (!allowedPaths[context.feature].test(continuation.pathname))
          throw new GraphBoundaryError(
            "Graph continuation escaped its declared feature pack.",
            "continuation",
          );
      }
      for (const raw of page.value) {
        if (items.length >= maxRecords) return { items, complete: false };
        if (options.schema) {
          const parsed = options.schema.safeParse(raw);
          if (!parsed.success)
            throw new GraphBoundaryError(
              "Microsoft Graph returned an invalid collection item.",
              "item",
            );
          items.push(parsed.data);
        } else items.push(raw);
      }
      next = continuationLink;
      if (items.length >= maxRecords) return { items, complete: !next };
    }
    return { items, complete: !next };
  }

  getAll(
    path: string,
    context: GraphRequestContext,
    options?: { maxPages?: number; maxRecords?: number },
  ): Promise<unknown[]>;
  getAll<T>(
    path: string,
    context: GraphRequestContext,
    options: { schema: z.ZodType<T>; maxPages?: number; maxRecords?: number },
  ): Promise<T[]>;
  async getAll(
    path: string,
    context: GraphRequestContext,
    options: { schema?: z.ZodType; maxPages?: number; maxRecords?: number } = {},
  ): Promise<unknown[]> {
    const result = await this.getCollection(path, context, options);
    if (!result.complete)
      throw new GraphBoundaryError(
        "Graph pagination exceeded its page or record safety limit.",
        "page_limit",
      );
    return result.items;
  }

  // The checks every Graph request passes before anything is sent: the declared pack and its
  // permission, the cloud, read-only methods, the configured host, the pack's paths and $expand.
  #checkedTarget(
    pathOrNextLink: string,
    context: GraphRequestContext,
    method: string,
    body: RequestInit["body"],
  ): URL {
    const pack = getPermissionPack(context.feature);
    if (pack.source !== "microsoftGraph" || pack.permission !== context.permission) {
      throw new GraphBoundaryError(
        "The Graph request permission does not match its declared feature pack.",
        "pack",
      );
    }
    if (!(pack.clouds as readonly MicrosoftCloud[]).includes(this.#cloud)) {
      throw new GraphBoundaryError(
        `Feature ${context.feature} is not supported in the selected Microsoft cloud (${this.#cloud}).`,
        "cloud",
      );
    }
    if ((method !== "GET" && method !== "HEAD") || body != null) {
      throw new GraphBoundaryError(
        `Microsoft Graph ${method} requests are disabled by the read-only boundary.`,
        "method",
      );
    }
    const target = this.#resolveTarget(pathOrNextLink);
    if (!allowedPaths[context.feature].test(target.pathname)) {
      throw new GraphBoundaryError(
        "Graph endpoint does not belong to its declared feature pack.",
        "endpoint",
      );
    }
    // $expand pulls related objects into the response, so it could reach data the pack's paths
    // deliberately leave out (an application's owners, a user's group memberships). Nothing in
    // Prism365 uses it; it is refused rather than allowlisted per pack.
    // Graph also accepts query options without the "$" prefix on some endpoints.
    const expands = (key: string) => key.trim().toLowerCase().replace(/^\$/, "") === "expand";
    if ([...target.searchParams.keys()].some(expands)) {
      throw new GraphBoundaryError(
        "Graph $expand is not allowed by the read boundary.",
        "endpoint",
      );
    }
    return target;
  }

  // Sends a checked Graph request with the bearer token and redirect "manual", so a redirect is
  // seen rather than followed. `answer` reads a 2xx or 3xx response inside the request's time
  // budget; any other status is cancelled here, and 429 or 503 is retried after Retry-After
  // within the retry budget, or becomes a GraphRequestError with only its status and request ID.
  async #exchange<T>(
    target: URL,
    method: string,
    init: Omit<RequestInit, "signal">,
    external: AbortSignal | null | undefined,
    accept: string,
    answer: (response: Response) => Promise<T>,
  ): Promise<T> {
    let waited = 0;
    for (let attempt = 0; attempt <= this.#maxRetries; attempt += 1) {
      external?.throwIfAborted();
      const controller = new AbortController();
      const forward = () => controller.abort(external?.reason);
      external?.addEventListener("abort", forward, { once: true });
      let timer: ReturnType<typeof setTimeout> | undefined;
      const operation = async (): Promise<
        { answered: true; value: T } | { answered: false; response: Response }
      > => {
        const token = await this.#getAccessToken();
        if (!token || controller.signal.aborted)
          throw new GraphBoundaryError(
            "Graph credentials are unavailable or the request expired.",
            "credential",
          );
        const headers = new Headers(init.headers);
        headers.set("Accept", accept);
        headers.set("Authorization", `Bearer ${token}`);
        const response = await this.#fetch(target, {
          ...init,
          method,
          headers,
          redirect: "manual",
          signal: controller.signal,
        });
        if (response.status < 400) return { answered: true, value: await answer(response) };
        await response.body?.cancel();
        return { answered: false, response };
      };
      let result: Awaited<ReturnType<typeof operation>>;
      try {
        result = await Promise.race([
          operation(),
          new Promise<never>((_resolve, reject) => {
            timer = setTimeout(() => {
              controller.abort();
              reject(new GraphBoundaryError("Graph request exceeded its time budget.", "timeout"));
            }, this.#timeout);
          }),
        ]);
      } catch (error) {
        if (external?.aborted) throw external.reason;
        throw error;
      } finally {
        if (timer !== undefined) clearTimeout(timer);
        external?.removeEventListener("abort", forward);
      }
      if (result.answered) return result.value;
      const { response } = result;
      const requestId = response.headers.get("request-id") ?? undefined;
      if (response.status === 429 || response.status === 503) {
        const delay = retryDelayMilliseconds(response.headers.get("Retry-After"), attempt);
        if (attempt >= this.#maxRetries || delay > this.#retryBudget - waited)
          throw new GraphRetryDeferredError(response.status, delay, requestId);
        waited += delay;
        await this.#sleep(delay);
        continue;
      }
      throw new GraphRequestError(response.status, requestId);
    }
    throw new GraphBoundaryError("Graph retry budget exhausted.", "retry_budget");
  }

  // The report download: a GET to the checked link with no token, no cookies and no header but
  // Accept, refusing any redirect, within the client's request timeout and the caller's signal.
  // The link is pre-authenticated and short-lived, so it is used once and never retried, and it
  // never goes into an error or a log: whoever holds it can read the report. Any answer but 200 is
  // a GraphReportDownloadError, a partial (206) or empty (204) one included.
  async #download(url: URL, maxBytes: number, external: AbortSignal | undefined): Promise<string> {
    external?.throwIfAborted();
    const controller = new AbortController();
    const forward = () => controller.abort(external?.reason);
    external?.addEventListener("abort", forward, { once: true });
    let timer: ReturnType<typeof setTimeout> | undefined;
    const operation = async () => {
      try {
        const response = await this.#fetch(url, {
          method: "GET",
          headers: new Headers({ Accept: DOWNLOAD_ACCEPT }),
          redirect: "error",
          credentials: "omit",
          signal: controller.signal,
        });
        if (response.status !== 200) {
          await response.body?.cancel();
          throw new GraphReportDownloadError(response.status);
        }
        return await readBoundedReport(response, maxBytes);
      } catch (error) {
        if (
          controller.signal.aborted ||
          error instanceof GraphBoundaryError ||
          error instanceof GraphReportDownloadError
        )
          throw error;
        // A network failure, or a redirect that fetch refused. The underlying error may carry
        // the link, so only a fixed message leaves here.
        throw new GraphBoundaryError(
          "The usage report download failed. Prism365 follows no redirect from the download host.",
          "download",
        );
      }
    };
    try {
      return await Promise.race([
        operation(),
        new Promise<never>((_resolve, reject) => {
          timer = setTimeout(() => {
            controller.abort();
            reject(
              new GraphBoundaryError(
                "The usage report download exceeded its time budget.",
                "timeout",
              ),
            );
          }, this.#timeout);
        }),
      ]);
    } catch (error) {
      if (external?.aborted) throw external.reason;
      throw error;
    } finally {
      if (timer !== undefined) clearTimeout(timer);
      external?.removeEventListener("abort", forward);
    }
  }

  #resolveTarget(pathOrNextLink: string): URL {
    const target = new URL(pathOrNextLink.replace(/^\//, ""), `${this.#baseUrl.href}/`);
    if (
      target.origin !== this.#baseUrl.origin ||
      !target.pathname.startsWith("/v1.0/") ||
      target.username ||
      target.password ||
      target.hash
    )
      throw new GraphBoundaryError("Graph URL escaped the configured v1.0 host boundary.", "host");
    return target;
  }
}

async function readBoundedJson(response: Response, maxBytes: number): Promise<unknown> {
  if (
    response.headers.get("Content-Type")?.split(";", 1)[0]?.trim().toLowerCase() !==
    "application/json"
  ) {
    await response.body?.cancel();
    throw new GraphBoundaryError(
      "Microsoft Graph returned an unsupported response content type.",
      "content_type",
    );
  }
  const tooLarge = "Microsoft Graph response exceeded the size budget.";
  const length = response.headers.get("Content-Length");
  if (length && Number(length) > maxBytes) {
    await response.body?.cancel();
    throw new GraphBoundaryError(tooLarge, "size");
  }
  if (!response.body)
    throw new GraphBoundaryError("Microsoft Graph returned an empty JSON response.", "empty");
  const bytes = await readBoundedBytes(response.body, maxBytes, tooLarge);
  try {
    return z.json().parse(JSON.parse(new TextDecoder().decode(bytes)));
  } catch {
    throw new GraphBoundaryError("Microsoft Graph returned invalid JSON.", "json");
  }
}

// A usage report body, as text for parseCsv. The byte cap is checked against Content-Length
// first and again while the body streams, so an oversized report stops partway.
async function readBoundedReport(response: Response, maxBytes: number): Promise<string> {
  const type = response.headers.get("Content-Type")?.split(";", 1)[0]?.trim().toLowerCase();
  if (type === undefined || !reportContentTypes.has(type)) {
    await response.body?.cancel();
    throw new GraphBoundaryError(
      "Microsoft returned a usage report with an unsupported content type.",
      "content_type",
    );
  }
  const tooLarge = "The usage report exceeded its size limit.";
  const length = response.headers.get("Content-Length");
  if (length && Number(length) > maxBytes) {
    await response.body?.cancel();
    throw new GraphBoundaryError(tooLarge, "size");
  }
  const bytes = response.body
    ? await readBoundedBytes(response.body, maxBytes, tooLarge)
    : new Uint8Array(0);
  // ignoreBOM keeps a leading byte-order mark in the text, so parseCsv strips exactly one. Bytes
  // that are not UTF-8 become U+FFFD rather than failing the report.
  return new TextDecoder("utf-8", { ignoreBOM: true }).decode(bytes);
}

async function readBoundedBytes(
  body: ReadableStream<Uint8Array>,
  maxBytes: number,
  tooLarge: string,
): Promise<Uint8Array> {
  const reader = body.getReader();
  const chunks: Uint8Array[] = [];
  let size = 0;
  try {
    for (;;) {
      const chunk = await reader.read();
      if (chunk.done) break;
      size += chunk.value.byteLength;
      if (size > maxBytes) throw new GraphBoundaryError(tooLarge, "size");
      chunks.push(chunk.value);
    }
  } finally {
    await reader.cancel();
    reader.releaseLock();
  }
  const bytes = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes;
}

// Checks the Location of a usage report redirect. A refusal is a GraphReportHostError, which names
// the host and nothing else of the link.
function reportDownloadUrl(location: string): URL {
  let url: URL | undefined;
  try {
    url = location.length <= 16_384 ? new URL(location) : undefined;
  } catch {
    url = undefined;
  }
  const problem = downloadLinkProblem(url);
  if (url && problem === null) return url;
  const host = url?.hostname ? url.hostname : null;
  throw new GraphReportHostError(
    problem ?? "it is not an allowed link",
    host,
    host !== null && reportDownloadHosts.includes(host),
  );
}

// Why a parsed link is not an allowed report download, or null when it is. `new URL` without a
// base refuses relative and protocol-relative links, so those arrive here as undefined.
function downloadLinkProblem(url: URL | undefined): string | null {
  if (!url) return "it is not an absolute URL";
  if (url.protocol !== "https:") return "it is not HTTPS";
  if (!reportDownloadHosts.includes(url.hostname)) return "its host is not a report download host";
  if (url.port !== "") return "it names a port";
  if (url.username !== "" || url.password !== "") return "it carries a user name or password";
  if (url.hash !== "") return "it has a fragment";
  if (
    !url.pathname.startsWith(REPORT_DOWNLOAD_PATH) ||
    url.pathname.length === REPORT_DOWNLOAD_PATH.length
  )
    return `its path is not under ${REPORT_DOWNLOAD_PATH}`;
  return null;
}

function retryDelayMilliseconds(retryAfter: string | null, attempt: number): number {
  if (retryAfter) {
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds) && seconds >= 0) return seconds * 1_000;
    const date = Date.parse(retryAfter);
    if (!Number.isNaN(date)) return Math.max(0, date - Date.now());
  }
  return Math.min(500 * 2 ** attempt, 30_000);
}

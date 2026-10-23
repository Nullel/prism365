import { z } from "zod";
import { getMicrosoftCloudProfile, type MicrosoftCloud } from "./clouds.ts";
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
  | "json";

export class GraphBoundaryError extends Error {
  readonly reason: GraphBoundaryReason;
  constructor(message: string, reason: GraphBoundaryReason = "configuration") {
    super(message);
    this.name = "GraphBoundaryError";
    this.reason = reason;
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
  // The user list, one user, or one user's manager.
  "directory.users.read": /^\/v1\.0\/users(?:(?:\/[^/]+|\([^/]+\))(?:\/manager)?)?\/?$/i,
  "user.licenses.read": /^\/v1\.0\/users(?:\/[^/]+|\([^/]+\))\/licenseDetails\/?$/i,
  "tenant.licenses.read": /^\/v1\.0\/subscribedSkus\/?$/i,
  "service.health.read":
    /^\/v1\.0\/admin\/serviceAnnouncement\/(?:healthOverviews|issues)(?:\/[^/]+)?\/?$/i,
  "service.messages.read": /^\/v1\.0\/admin\/serviceAnnouncement\/messages(?:\/[^/]+)?\/?$/i,
  "audit.signins.read": /^\/v1\.0\/auditLogs\/signIns(?:\/[^/]+)?\/?$/i,
  "audit.directory.read": /^\/v1\.0\/auditLogs\/directoryAudits(?:\/[^/]+)?\/?$/i,
  "reports.authentication-methods.read":
    /^\/v1\.0\/reports\/authenticationMethods\/userRegistrationDetails(?:\/[^/]+)?\/?$/i,
  "identity.conditional-access.read":
    /^\/v1\.0\/identity\/conditionalAccess\/(?:policies|namedLocations)(?:\/[^/]+)?\/?$/i,
  // Enterprise applications, one of them, or the application permissions granted to one.
  "directory.applications.read":
    /^\/v1\.0\/servicePrincipals(?:\/[^/]+(?:\/appRoleAssignments)?)?\/?$/i,
  "directory.permission-grants.read": /^\/v1\.0\/oauth2PermissionGrants\/?$/i,
  "identity.risk.read": /^\/v1\.0\/identityProtection\/riskyUsers(?:\/[^/]+)?\/?$/i,
  "reports.usage.read": /^\/v1\.0\/reports\/get[A-Za-z0-9]+\([^/]+\)\/?$/i,
  "reports.settings.read": /^\/v1\.0\/admin\/reportSettings\/?$/i,
  "security.score.read":
    /^\/v1\.0\/security\/(?:secureScores|secureScoreControlProfiles)(?:\/[^/]+)?\/?$/i,
  "security.incidents.read": /^\/v1\.0\/security\/incidents(?:\/[^/]+)?\/?$/i,
  "devices.intune.read": /^\/v1\.0\/deviceManagement\/managedDevices(?:\/[^/]+)?\/?$/i,
  "purview.entra.read": /^\/v1\.0\/security\/auditLog\/queries(?:\/[^/]+(?:\/records)?)?\/?$/i,
};

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
    const method = (init.method ?? "GET").toUpperCase();
    if ((method !== "GET" && method !== "HEAD") || init.body != null) {
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
    // A caller's signal (an audit's time limit, a collector run ending) stops the request and
    // any retry; the request's own timeout still applies.
    const { signal: external, ...requestInit } = init;
    let waited = 0;
    for (let attempt = 0; attempt <= this.#maxRetries; attempt += 1) {
      external?.throwIfAborted();
      const controller = new AbortController();
      const forward = () => controller.abort(external?.reason);
      external?.addEventListener("abort", forward, { once: true });
      let timer: ReturnType<typeof setTimeout> | undefined;
      const operation = async () => {
        const token = await this.#getAccessToken();
        if (!token || controller.signal.aborted)
          throw new GraphBoundaryError(
            "Graph credentials are unavailable or the request expired.",
            "credential",
          );
        const headers = new Headers(requestInit.headers);
        headers.set("Accept", "application/json");
        headers.set("Authorization", `Bearer ${token}`);
        const response = await this.#fetch(target, {
          ...requestInit,
          method,
          headers,
          redirect: "manual",
          signal: controller.signal,
        });
        if (response.status >= 300 && response.status < 400) {
          await response.body?.cancel();
          throw new GraphBoundaryError(
            "Graph redirects are disabled; this resource needs a reviewed download adapter.",
            "redirect",
          );
        }
        if (!response.ok || response.status === 204 || method === "HEAD") {
          await response.body?.cancel();
          return { response, value: undefined };
        }
        return { response, value: await readBoundedJson(response, this.#maxResponseBytes) };
      };
      let result: { response: Response; value: unknown };
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
      const { response, value } = result;
      if (response.ok) {
        if (!schema) return value;
        const parsed = schema.safeParse(value);
        if (!parsed.success)
          throw new GraphBoundaryError(
            "Microsoft Graph returned an invalid response shape.",
            "shape",
          );
        return parsed.data;
      }
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
  const length = response.headers.get("Content-Length");
  if (length && Number(length) > maxBytes) {
    await response.body?.cancel();
    throw new GraphBoundaryError("Microsoft Graph response exceeded the size budget.", "size");
  }
  if (!response.body)
    throw new GraphBoundaryError("Microsoft Graph returned an empty JSON response.", "empty");
  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let size = 0;
  try {
    for (;;) {
      const chunk = await reader.read();
      if (chunk.done) break;
      size += chunk.value.byteLength;
      if (size > maxBytes)
        throw new GraphBoundaryError("Microsoft Graph response exceeded the size budget.", "size");
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
  try {
    return z.json().parse(JSON.parse(new TextDecoder().decode(bytes)));
  } catch {
    throw new GraphBoundaryError("Microsoft Graph returned invalid JSON.", "json");
  }
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

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

export class GraphBoundaryError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "GraphBoundaryError";
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
const allowedPaths: Record<GraphFeatureId, RegExp> = {
  profile: /^\/v1\.0\/me\/?$/,
  "directory.users.read": /^\/v1\.0\/users(?:\/[^/]+|\([^/]+\))?\/?$/,
  "user.licenses.read": /^\/v1\.0\/users(?:\/[^/]+|\([^/]+\))\/licenseDetails\/?$/,
  "tenant.licenses.read": /^\/v1\.0\/subscribedSkus\/?$/,
  "service.health.read":
    /^\/v1\.0\/admin\/serviceAnnouncement\/(?:healthOverviews|issues)(?:\/[^/]+)?\/?$/,
  "audit.signins.read": /^\/v1\.0\/auditLogs\/(?:signIns|directoryAudits)(?:\/[^/]+)?\/?$/,
  "identity.conditional-access.read":
    /^\/v1\.0\/identity\/conditionalAccess\/policies(?:\/[^/]+)?\/?$/,
  "identity.risk.read": /^\/v1\.0\/identityProtection\/riskyUsers(?:\/[^/]+)?\/?$/,
  "reports.usage.read": /^\/v1\.0\/reports\/get[A-Za-z0-9]+\([^/]+\)\/?$/,
  "reports.settings.read": /^\/v1\.0\/admin\/reportSettings\/?$/,
  "security.score.read":
    /^\/v1\.0\/security\/(?:secureScores|secureScoreControlProfiles)(?:\/[^/]+)?\/?$/,
  "security.incidents.read": /^\/v1\.0\/security\/incidents(?:\/[^/]+)?\/?$/,
  "devices.intune.read": /^\/v1\.0\/deviceManagement\/managedDevices(?:\/[^/]+)?\/?$/,
  "purview.entra.read": /^\/v1\.0\/security\/auditLog\/queries(?:\/[^/]+(?:\/records)?)?\/?$/,
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
      30_000,
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
      );
    }
    if (!(pack.clouds as readonly MicrosoftCloud[]).includes(this.#cloud)) {
      throw new GraphBoundaryError(
        `Feature ${context.feature} is not supported in the selected Microsoft cloud (${this.#cloud}).`,
      );
    }
    const method = (init.method ?? "GET").toUpperCase();
    if ((method !== "GET" && method !== "HEAD") || init.body != null) {
      throw new GraphBoundaryError(
        `Microsoft Graph ${method} requests are disabled by the read-only boundary.`,
      );
    }
    const target = this.#resolveTarget(pathOrNextLink);
    if (!allowedPaths[context.feature].test(target.pathname)) {
      throw new GraphBoundaryError("Graph endpoint does not belong to its declared feature pack.");
    }
    let waited = 0;
    for (let attempt = 0; attempt <= this.#maxRetries; attempt += 1) {
      const controller = new AbortController();
      let timer: ReturnType<typeof setTimeout> | undefined;
      const operation = async () => {
        const token = await this.#getAccessToken();
        if (!token || controller.signal.aborted)
          throw new GraphBoundaryError("Graph credentials are unavailable or the request expired.");
        const headers = new Headers(init.headers);
        headers.set("Accept", "application/json");
        headers.set("Authorization", `Bearer ${token}`);
        const response = await this.#fetch(target, {
          ...init,
          method,
          headers,
          redirect: "manual",
          signal: controller.signal,
        });
        if (response.status >= 300 && response.status < 400) {
          await response.body?.cancel();
          throw new GraphBoundaryError(
            "Graph redirects are disabled; this resource needs a reviewed download adapter.",
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
              reject(new GraphBoundaryError("Graph request exceeded its time budget."));
            }, this.#timeout);
          }),
        ]);
      } finally {
        if (timer !== undefined) clearTimeout(timer);
      }
      const { response, value } = result;
      if (response.ok) {
        if (!schema) return value;
        const parsed = schema.safeParse(value);
        if (!parsed.success)
          throw new GraphBoundaryError("Microsoft Graph returned an invalid response shape.");
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
    throw new GraphBoundaryError("Graph retry budget exhausted.");
  }

  getCollection(
    path: string,
    context: GraphRequestContext,
    options?: { maxPages?: number; maxRecords?: number },
  ): Promise<{ items: unknown[]; complete: boolean }>;
  getCollection<T>(
    path: string,
    context: GraphRequestContext,
    options: { schema: z.ZodType<T>; maxPages?: number; maxRecords?: number },
  ): Promise<{ items: T[]; complete: boolean }>;
  async getCollection(
    path: string,
    context: GraphRequestContext,
    options: { schema?: z.ZodType; maxPages?: number; maxRecords?: number } = {},
  ): Promise<{ items: unknown[]; complete: boolean }> {
    const maxPages = boundedInteger(options.maxPages ?? 100, 1, 100, "maxPages");
    const maxRecords = boundedInteger(options.maxRecords ?? 10_000, 1, 10_000, "maxRecords");
    const items: unknown[] = [];
    const seen = new Set<string>();
    let next: string | undefined = path;
    for (let pageNumber = 0; next && pageNumber < maxPages; pageNumber += 1) {
      const target = this.#resolveTarget(next).href;
      if (seen.has(target))
        throw new GraphBoundaryError("Graph pagination repeated a continuation link.");
      seen.add(target);
      const page: z.infer<typeof collectionSchema> = await this.request(
        next,
        context,
        {},
        collectionSchema,
      );
      const continuationLink = page["@odata.nextLink"];
      if (continuationLink) {
        const continuation = this.#resolveTarget(continuationLink);
        if (!allowedPaths[context.feature].test(continuation.pathname))
          throw new GraphBoundaryError("Graph continuation escaped its declared feature pack.");
      }
      for (const raw of page.value) {
        if (items.length >= maxRecords) return { items, complete: false };
        if (options.schema) {
          const parsed = options.schema.safeParse(raw);
          if (!parsed.success)
            throw new GraphBoundaryError("Microsoft Graph returned an invalid collection item.");
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
      throw new GraphBoundaryError("Graph pagination exceeded its page or record safety limit.");
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
      throw new GraphBoundaryError("Graph URL escaped the configured v1.0 host boundary.");
    return target;
  }
}

async function readBoundedJson(response: Response, maxBytes: number): Promise<unknown> {
  if (
    response.headers.get("Content-Type")?.split(";", 1)[0]?.trim().toLowerCase() !==
    "application/json"
  ) {
    await response.body?.cancel();
    throw new GraphBoundaryError("Microsoft Graph returned an unsupported response content type.");
  }
  const length = response.headers.get("Content-Length");
  if (length && Number(length) > maxBytes) {
    await response.body?.cancel();
    throw new GraphBoundaryError("Microsoft Graph response exceeded the size budget.");
  }
  if (!response.body)
    throw new GraphBoundaryError("Microsoft Graph returned an empty JSON response.");
  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let size = 0;
  try {
    for (;;) {
      const chunk = await reader.read();
      if (chunk.done) break;
      size += chunk.value.byteLength;
      if (size > maxBytes)
        throw new GraphBoundaryError("Microsoft Graph response exceeded the size budget.");
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
    throw new GraphBoundaryError("Microsoft Graph returned invalid JSON.");
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

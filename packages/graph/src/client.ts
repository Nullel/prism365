import type { GraphPermission, PermissionPackId } from "./permissions.ts";

const GLOBAL_GRAPH_ORIGIN = "https://graph.microsoft.com";
const GLOBAL_GRAPH_ROOT = `${GLOBAL_GRAPH_ORIGIN}/v1.0`;

type FetchImplementation = (input: RequestInfo | URL, init?: RequestInit) => Promise<Response>;
type Sleep = (milliseconds: number) => Promise<void>;

export interface GraphRequestContext {
  feature: PermissionPackId;
  permission: GraphPermission;
}

export interface GraphClientOptions {
  getAccessToken: () => Promise<string>;
  fetchImplementation?: FetchImplementation;
  sleep?: Sleep;
  maxRetries?: number;
  allowMutations?: boolean;
  baseUrl?: string;
}

interface GraphCollection<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

export class GraphRequestError extends Error {
  readonly status: number;
  readonly requestId: string | undefined;

  constructor(status: number, requestId?: string) {
    super(`Microsoft Graph request failed with HTTP ${status}.`);
    this.name = "GraphRequestError";
    this.status = status;
    this.requestId = requestId;
  }
}

export class GraphHttpClient {
  readonly #getAccessToken: () => Promise<string>;
  readonly #fetch: FetchImplementation;
  readonly #sleep: Sleep;
  readonly #maxRetries: number;
  readonly #allowMutations: boolean;
  readonly #baseUrl: URL;

  constructor(options: GraphClientOptions) {
    this.#getAccessToken = options.getAccessToken;
    this.#fetch = options.fetchImplementation ?? globalThis.fetch.bind(globalThis);
    this.#sleep =
      options.sleep ??
      ((milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds)));
    this.#maxRetries = options.maxRetries ?? 3;
    this.#allowMutations = options.allowMutations ?? false;
    this.#baseUrl = new URL(options.baseUrl ?? GLOBAL_GRAPH_ROOT);

    if (
      this.#baseUrl.protocol !== "https:" ||
      this.#baseUrl.origin !== GLOBAL_GRAPH_ORIGIN ||
      this.#baseUrl.pathname.replace(/\/$/, "") !== "/v1.0"
    ) {
      throw new Error("Only the Microsoft Graph global v1.0 endpoint is enabled in this scaffold.");
    }
  }

  async request<T>(
    pathOrNextLink: string,
    context: GraphRequestContext,
    init: RequestInit = {},
  ): Promise<T> {
    if (!context.feature || !context.permission) {
      throw new Error("Every Graph request must declare its feature and permission.");
    }

    const method = (init.method ?? "GET").toUpperCase();
    if (!this.#allowMutations && method !== "GET" && method !== "HEAD") {
      throw new Error(`Microsoft Graph ${method} requests are disabled by the read-only boundary.`);
    }

    const target = this.#resolveTarget(pathOrNextLink);
    const token = await this.#getAccessToken();

    for (let attempt = 0; attempt <= this.#maxRetries; attempt += 1) {
      const headers = new Headers(init.headers);
      headers.set("Accept", "application/json");
      headers.set("Authorization", `Bearer ${token}`);

      const response = await this.#fetch(target, { ...init, method, headers });
      if (response.ok) {
        if (response.status === 204) {
          return undefined as T;
        }
        return (await response.json()) as T;
      }

      if ((response.status === 429 || response.status === 503) && attempt < this.#maxRetries) {
        await this.#sleep(retryDelayMilliseconds(response.headers.get("Retry-After"), attempt));
        continue;
      }

      throw new GraphRequestError(
        response.status,
        response.headers.get("request-id") ??
          response.headers.get("client-request-id") ??
          undefined,
      );
    }

    throw new Error("Unreachable retry state.");
  }

  async getAll<T>(
    path: string,
    context: GraphRequestContext,
    options: { maxPages?: number } = {},
  ): Promise<T[]> {
    const maxPages = options.maxPages ?? 100;
    const values: T[] = [];
    let next: string | undefined = path;

    for (let pageNumber = 0; next && pageNumber < maxPages; pageNumber += 1) {
      const page: GraphCollection<T> = await this.request<GraphCollection<T>>(next, context);
      values.push(...page.value);
      next = page["@odata.nextLink"];
    }

    if (next) {
      throw new Error(`Graph pagination exceeded the safety limit of ${maxPages} pages.`);
    }

    return values;
  }

  #resolveTarget(pathOrNextLink: string): URL {
    const target = pathOrNextLink.startsWith("https://")
      ? new URL(pathOrNextLink)
      : new URL(
          pathOrNextLink.replace(/^\//, ""),
          `${this.#baseUrl.toString().replace(/\/$/, "")}/`,
        );

    if (target.origin !== this.#baseUrl.origin || !target.pathname.startsWith("/v1.0/")) {
      throw new Error("Graph URL escaped the configured v1.0 host boundary.");
    }

    return target;
  }
}

function retryDelayMilliseconds(retryAfter: string | null, attempt: number): number {
  if (retryAfter) {
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds) && seconds >= 0) {
      return Math.min(seconds * 1_000, 60_000);
    }

    const date = Date.parse(retryAfter);
    if (!Number.isNaN(date)) {
      return Math.max(0, Math.min(date - Date.now(), 60_000));
    }
  }

  return Math.min(500 * 2 ** attempt, 30_000);
}

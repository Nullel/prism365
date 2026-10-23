import { z } from "zod";

// The setup's own small Microsoft Graph client. It is the only code in the repository that
// writes to Microsoft, so it lives here, apart from the runtime's read-only GraphHttpClient,
// which it neither uses nor changes. It talks to graph.microsoft.com/v1.0 only, never logs a
// token, and turns failures into Microsoft's error code plus a plain explanation, never the
// provider's own message or body.

const GRAPH_BASE = "https://graph.microsoft.com/v1.0/";
const MAX_RESPONSE_BYTES = 8_388_608;
const MAX_RETRIES = 3;
const MAX_RETRY_WAIT_MILLISECONDS = 30_000;

export type Fetch = (input: string | URL, init?: RequestInit) => Promise<Response>;

export interface SetupGraphOptions {
  fetch: Fetch;
  getToken: () => Promise<string>;
  sleep?: (milliseconds: number) => Promise<void>;
  timeoutMilliseconds?: number;
}

export class SetupGraphError extends Error {
  readonly status: number;
  readonly code: string | null;
  readonly requestId: string | null;
  constructor(status: number, code: string | null, requestId: string | null) {
    super(`Microsoft Graph answered HTTP ${status}${code ? ` (${code})` : ""}.`);
    this.name = "SetupGraphError";
    this.status = status;
    this.code = code;
    this.requestId = requestId;
  }
}

const errorBodySchema = z.object({
  error: z.object({ code: z.string().regex(/^[A-Za-z0-9_.-]{1,100}$/) }),
});

async function errorFrom(response: Response): Promise<SetupGraphError> {
  let code: string | null = null;
  try {
    const parsed = errorBodySchema.safeParse(JSON.parse(await response.text()));
    if (parsed.success) code = parsed.data.error.code;
  } catch {
    // An unreadable error body still has its status.
  }
  const requestId = response.headers.get("request-id");
  return new SetupGraphError(
    response.status,
    code,
    requestId && /^[\w-]{1,128}$/.test(requestId) ? requestId : null,
  );
}

function retryDelay(response: Response, attempt: number): number {
  const seconds = Number(response.headers.get("Retry-After"));
  return Number.isFinite(seconds) && seconds >= 0 ? seconds * 1_000 : 1_000 * 2 ** attempt;
}

export class SetupGraph {
  readonly #fetch: Fetch;
  readonly #getToken: () => Promise<string>;
  readonly #sleep: (milliseconds: number) => Promise<void>;
  readonly #timeout: number;

  constructor(options: SetupGraphOptions) {
    this.#fetch = options.fetch;
    this.#getToken = options.getToken;
    this.#sleep =
      options.sleep ?? ((milliseconds) => new Promise((done) => setTimeout(done, milliseconds)));
    this.#timeout = options.timeoutMilliseconds ?? 60_000;
  }

  async #send(method: "GET" | "POST" | "PATCH", path: string, body?: unknown) {
    const url = new URL(path, GRAPH_BASE);
    if (url.origin !== "https://graph.microsoft.com" || !url.pathname.startsWith("/v1.0/"))
      throw new Error("The setup only talks to Microsoft Graph v1.0.");
    for (let attempt = 0; ; attempt += 1) {
      const headers = new Headers({
        Accept: "application/json",
        Authorization: `Bearer ${await this.#getToken()}`,
      });
      if (body !== undefined) headers.set("Content-Type", "application/json");
      const response = await this.#fetch(url, {
        method,
        headers,
        redirect: "error",
        signal: AbortSignal.timeout(this.#timeout),
        ...(body === undefined ? {} : { body: JSON.stringify(body) }),
      });
      if ([429, 503, 504].includes(response.status) && attempt < MAX_RETRIES) {
        const delay = retryDelay(response, attempt);
        await response.body?.cancel();
        if (delay <= MAX_RETRY_WAIT_MILLISECONDS) {
          await this.#sleep(delay);
          continue;
        }
      }
      if (!response.ok) throw await errorFrom(response);
      if (response.status === 204) return undefined;
      const bytes = await response.arrayBuffer();
      if (bytes.byteLength > MAX_RESPONSE_BYTES)
        throw new Error("Microsoft Graph sent a response larger than the setup accepts.");
      if (bytes.byteLength === 0) return undefined;
      try {
        return JSON.parse(new TextDecoder().decode(bytes)) as unknown;
      } catch {
        throw new Error("Microsoft Graph sent a response that is not JSON.");
      }
    }
  }

  async get<T>(path: string, schema: z.ZodType<T>): Promise<T> {
    return parse(schema, await this.#send("GET", path));
  }

  async post<T>(path: string, body: unknown, schema: z.ZodType<T>): Promise<T> {
    return parse(schema, await this.#send("POST", path, body));
  }

  async patch(path: string, body: unknown): Promise<void> {
    await this.#send("PATCH", path, body);
  }
}

function parse<T>(schema: z.ZodType<T>, value: unknown): T {
  const parsed = schema.safeParse(value);
  if (!parsed.success)
    throw new Error("Microsoft Graph sent a response in a shape the setup does not expect.");
  return parsed.data;
}

// A plain sentence for a Graph failure during one step of the setup.
export function explainGraphError(step: string, error: SetupGraphError): string {
  const reason =
    error.code === "Authorization_RequestDenied" || error.status === 403
      ? "The signed-in account is not allowed to do this. Creating and changing app registrations needs a role such as Application Administrator or Cloud Application Administrator."
      : error.status === 401
        ? "Microsoft did not accept the sign-in. Run the setup again to sign in."
        : error.code === "Request_ResourceNotFound" || error.status === 404
          ? "Microsoft could not find the object. If it was created moments ago, wait a minute and run the setup again; it carries on from where it stopped."
          : error.code === "Directory_QuotaExceeded"
            ? "The tenant has reached its limit of directory objects."
            : error.status === 429 || error.status === 503 || error.status === 504
              ? "Microsoft is busy. Wait a few minutes and run the setup again; it carries on from where it stopped."
              : error.status === 400
                ? "Microsoft rejected the request."
                : "Microsoft could not complete the request. Run the setup again; it carries on from where it stopped.";
  return `${step} failed: Microsoft Graph answered HTTP ${error.status}${error.code ? ` (${error.code})` : ""}. ${reason}${error.requestId ? ` Microsoft request ID: ${error.requestId}.` : ""}`;
}

// Text from Microsoft (names, the device code message) printed to a terminal: control and
// direction characters are removed so it cannot move the cursor or disguise itself.
export function terminalText(value: string, limit = 300): string {
  return value
    .replace(/[\p{Cc}‎‏‪-‮⁦-⁩]/gu, " ")
    .slice(0, limit)
    .trim();
}

const guid = z
  .string()
  .regex(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i)
  .transform((value) => value.toLowerCase());

export const meSchema = z.object({
  id: guid,
  userPrincipalName: z.string().max(1_024).nullable().optional(),
});

const resourceAccessSchema = z.object({
  id: z.string().max(100),
  type: z.string().max(40),
});
const requiredResourceAccessSchema = z.object({
  resourceAppId: z.string().max(100),
  resourceAccess: z.array(resourceAccessSchema).max(1_000),
});

export const applicationSchema = z.object({
  id: guid,
  appId: guid,
  displayName: z.string().max(1_024).nullable().optional(),
  signInAudience: z.string().max(100).nullable().optional(),
  tags: z.array(z.string().max(1_024)).max(1_000).nullable().optional(),
  web: z
    .object({
      redirectUris: z.array(z.string().max(2_048)).max(1_000).nullable().optional(),
      implicitGrantSettings: z
        .object({
          enableIdTokenIssuance: z.boolean().nullable().optional(),
          enableAccessTokenIssuance: z.boolean().nullable().optional(),
        })
        .loose()
        .nullable()
        .optional(),
    })
    .loose()
    .nullable()
    .optional(),
  requiredResourceAccess: z.array(requiredResourceAccessSchema).max(1_000).nullable().optional(),
});
export type GraphApplication = z.infer<typeof applicationSchema>;

export const applicationListSchema = z.object({
  value: z.array(applicationSchema).max(100),
  "@odata.nextLink": z.string().optional(),
});

export const servicePrincipalListSchema = z.object({
  value: z.array(z.object({ id: guid })).max(100),
});
export const servicePrincipalSchema = z.object({ id: guid });

const publishedSchema = z.object({
  id: z.string().max(100),
  value: z.string().max(400).nullable(),
  isEnabled: z.boolean().nullable().optional(),
});
export const graphServicePrincipalSchema = z.object({
  id: guid,
  oauth2PermissionScopes: z.array(publishedSchema).max(5_000),
  appRoles: z.array(publishedSchema).max(5_000).nullable().optional(),
});
export type GraphServicePrincipal = z.infer<typeof graphServicePrincipalSchema>;

export const keyCredentialsResponseSchema = z.object({
  keyCredentials: z.array(z.unknown()).max(500),
});

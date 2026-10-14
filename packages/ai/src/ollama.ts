import type { AiProvider, AiProviderInput } from "./index.ts";

type FetchImplementation = (input: RequestInfo | URL, init?: RequestInit) => Promise<Response>;

export interface LocalOllamaProviderOptions {
  model: string;
  endpoint?: string;
  fetchImplementation?: FetchImplementation;
  timeoutMilliseconds?: number;
  maxResponseBytes?: number;
}

const outputSchema = {
  type: "object",
  properties: {
    summary: { type: "string" },
    citations: { type: "array", items: { type: "string" } },
    limitations: { type: "array", minItems: 1, items: { type: "string" } },
  },
  required: ["summary", "citations", "limitations"],
  additionalProperties: false,
} as const;

export class LocalOllamaProvider implements AiProvider {
  readonly descriptor;
  readonly #model: string;
  readonly #endpoint: URL;
  readonly #fetch: FetchImplementation;
  readonly #timeoutMilliseconds: number;
  readonly #maxResponseBytes: number;

  constructor(options: LocalOllamaProviderOptions) {
    if (!/^[A-Za-z0-9._:/-]{1,200}$/.test(options.model)) {
      throw new Error("The Ollama model name contains unsupported characters.");
    }

    this.#endpoint = new URL(options.endpoint ?? "http://127.0.0.1:11434/api/chat");
    const loopbackHosts = new Set(["127.0.0.1", "localhost", "[::1]"]);
    if (
      !loopbackHosts.has(this.#endpoint.hostname) ||
      !["http:", "https:"].includes(this.#endpoint.protocol) ||
      this.#endpoint.pathname !== "/api/chat" ||
      this.#endpoint.username ||
      this.#endpoint.password ||
      this.#endpoint.search ||
      this.#endpoint.hash
    ) {
      throw new Error("The local Ollama adapter only accepts an exact loopback /api/chat URL.");
    }

    this.#timeoutMilliseconds = options.timeoutMilliseconds ?? 60_000;
    if (
      !Number.isInteger(this.#timeoutMilliseconds) ||
      this.#timeoutMilliseconds < 1_000 ||
      this.#timeoutMilliseconds > 300_000
    ) {
      throw new Error("The Ollama timeout must be between 1000 and 300000 milliseconds.");
    }
    this.#maxResponseBytes = options.maxResponseBytes ?? 1_000_000;
    if (
      !Number.isInteger(this.#maxResponseBytes) ||
      this.#maxResponseBytes < 1_024 ||
      this.#maxResponseBytes > 2_000_000
    ) {
      throw new Error("The Ollama response limit must be between 1024 and 2000000 bytes.");
    }

    this.#model = options.model;
    this.#fetch = options.fetchImplementation ?? globalThis.fetch.bind(globalThis);
    this.descriptor = {
      id: `ollama-local:${this.#model}`,
      label: `Local Ollama (${this.#model})`,
      deployment: "local" as const,
    };
  }

  async analyze(input: AiProviderInput, options: { signal?: AbortSignal } = {}): Promise<unknown> {
    const signal = AbortSignal.any([
      AbortSignal.timeout(this.#timeoutMilliseconds),
      ...(options.signal ? [options.signal] : []),
    ]);
    signal.throwIfAborted();
    const body = JSON.stringify({
      model: this.#model,
      stream: false,
      think: false,
      format: outputSchema,
      options: { temperature: 0 },
      messages: [
        {
          role: "system",
          content: `${input.systemPolicy.join("\n")}\nReturn only JSON matching the supplied schema.`,
        },
        {
          role: "user",
          content: JSON.stringify({
            question: input.question,
            purpose: input.purpose,
            timeRange: input.timeRange,
            evidence: input.evidence,
          }),
        },
      ],
    });
    if (new TextEncoder().encode(body).byteLength > 1_000_000) {
      throw new Error("The local Ollama request exceeded the safety limit.");
    }
    let response: Response;
    try {
      response = await this.#fetch(this.#endpoint, {
        method: "POST",
        redirect: "error",
        headers: {
          Accept: "application/json",
          "Content-Type": "application/json",
        },
        body,
        signal,
      });
    } catch {
      throw new Error(
        signal.aborted
          ? "The local Ollama request was cancelled or timed out."
          : "The local Ollama request failed.",
      );
    }

    if (response.redirected || (response.status >= 300 && response.status < 400)) {
      await response.body?.cancel();
      throw new Error("The local Ollama provider attempted a redirect.");
    }

    if (!response.ok) {
      await response.body?.cancel();
      throw new Error(`The local Ollama provider failed with HTTP ${response.status}.`);
    }

    const responseText = await readBoundedResponse(response, this.#maxResponseBytes, signal);

    const payload = parseObject(responseText, "The local Ollama provider returned invalid JSON.");
    const message = payload.message;
    if (!isRecord(message) || typeof message.content !== "string") {
      throw new Error("The local Ollama provider response did not contain message content.");
    }
    if (typeof payload.model !== "string" || payload.model.length === 0) {
      throw new Error("The local Ollama provider response did not identify its model.");
    }
    if (payload.model !== this.#model) {
      throw new Error("The local Ollama provider returned a different model than configured.");
    }

    const result = parseObject(
      message.content,
      "The local Ollama model returned invalid structured output.",
    );
    return { ...result, model: payload.model };
  }
}

async function readBoundedResponse(
  response: Response,
  maxBytes: number,
  signal: AbortSignal,
): Promise<string> {
  const declaredLength = response.headers.get("content-length");
  if (declaredLength && Number(declaredLength) > maxBytes) {
    await response.body?.cancel();
    throw new Error("The local Ollama provider response exceeded the safety limit.");
  }
  if (!response.body) throw new Error("The local Ollama provider returned an empty response.");
  const reader = response.body.getReader();
  const cancelOnAbort = () => {
    void reader.cancel().catch(() => undefined);
  };
  signal.addEventListener("abort", cancelOnAbort, { once: true });
  const decoder = new TextDecoder();
  let size = 0;
  let serialized = "";
  try {
    while (true) {
      signal.throwIfAborted();
      const { value, done } = await reader.read();
      signal.throwIfAborted();
      if (done) return serialized + decoder.decode();
      size += value.byteLength;
      if (size > maxBytes)
        throw new Error("The local Ollama provider response exceeded the safety limit.");
      serialized += decoder.decode(value, { stream: true });
    }
  } catch (error) {
    await reader.cancel().catch(() => undefined);
    if (signal.aborted) throw new Error("The local Ollama response was cancelled or timed out.");
    if (
      error instanceof Error &&
      error.message === "The local Ollama provider response exceeded the safety limit."
    )
      throw error;
    throw new Error("The local Ollama response could not be read.");
  } finally {
    signal.removeEventListener("abort", cancelOnAbort);
    reader.releaseLock();
  }
}

function parseObject(serialized: string, errorMessage: string): Record<string, unknown> {
  try {
    const parsed: unknown = JSON.parse(serialized);
    if (isRecord(parsed)) {
      return parsed;
    }
  } catch {
    // The caller receives a stable error without provider-controlled response content.
  }
  throw new Error(errorMessage);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

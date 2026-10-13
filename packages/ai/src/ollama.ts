import type { AiProvider, AiProviderInput } from "./index.ts";

type FetchImplementation = (input: RequestInfo | URL, init?: RequestInit) => Promise<Response>;

export interface LocalOllamaProviderOptions {
  model: string;
  endpoint?: string;
  fetchImplementation?: FetchImplementation;
  timeoutMilliseconds?: number;
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
    if (this.#timeoutMilliseconds < 1_000 || this.#timeoutMilliseconds > 300_000) {
      throw new Error("The Ollama timeout must be between 1000 and 300000 milliseconds.");
    }

    this.#model = options.model;
    this.#fetch = options.fetchImplementation ?? globalThis.fetch.bind(globalThis);
    this.descriptor = {
      id: `ollama-local:${this.#model}`,
      label: `Local Ollama (${this.#model})`,
      deployment: "local" as const,
    };
  }

  async analyze(input: AiProviderInput): Promise<unknown> {
    const response = await this.#fetch(this.#endpoint, {
      method: "POST",
      headers: {
        Accept: "application/json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
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
      }),
      signal: AbortSignal.timeout(this.#timeoutMilliseconds),
    });

    if (!response.ok) {
      throw new Error(`The local Ollama provider failed with HTTP ${response.status}.`);
    }

    const responseText = await response.text();
    if (new TextEncoder().encode(responseText).byteLength > 1_000_000) {
      throw new Error("The local Ollama provider response exceeded the safety limit.");
    }

    const payload = parseObject(responseText, "The local Ollama provider returned invalid JSON.");
    const message = payload.message;
    if (!isRecord(message) || typeof message.content !== "string") {
      throw new Error("The local Ollama provider response did not contain message content.");
    }
    if (typeof payload.model !== "string" || payload.model.length === 0) {
      throw new Error("The local Ollama provider response did not identify its model.");
    }

    const result = parseObject(
      message.content,
      "The local Ollama model returned invalid structured output.",
    );
    return { ...result, model: payload.model };
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

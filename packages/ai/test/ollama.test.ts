import { describe, expect, it, vi } from "vitest";
import { AI_SYSTEM_POLICY, type AiProviderInput, LocalOllamaProvider } from "../src/index.ts";

const input: AiProviderInput = {
  question: "What needs review?",
  purpose: "securityInvestigation",
  timeRange: {
    start: "2026-10-01T00:00:00.000Z",
    end: "2026-10-13T00:00:00.000Z",
  },
  evidence: [
    {
      evidenceId: "evidence-synthetic-1",
      dataset: "signIns",
      observedAt: "2026-10-12T00:00:00.000Z",
      source: "synthetic",
      sensitivity: "aggregate",
      fields: { failures: 2 },
      limitations: ["Synthetic data."],
    },
  ],
  systemPolicy: AI_SYSTEM_POLICY,
};

describe("LocalOllamaProvider", () => {
  it("uses the loopback chat API with structured output and no tools", async () => {
    const fetchImplementation = vi.fn(async (_input: RequestInfo | URL, _init?: RequestInit) =>
      Response.json({
        model: "qwen3:8b",
        message: {
          role: "assistant",
          content: JSON.stringify({
            summary: "Two synthetic failures need review.",
            citations: ["evidence-synthetic-1"],
            limitations: ["Synthetic data only."],
          }),
        },
      }),
    );
    const provider = new LocalOllamaProvider({
      model: "qwen3:8b",
      fetchImplementation,
    });

    await expect(provider.analyze(input)).resolves.toEqual({
      summary: "Two synthetic failures need review.",
      citations: ["evidence-synthetic-1"],
      limitations: ["Synthetic data only."],
      model: "qwen3:8b",
    });

    const [target, init] = fetchImplementation.mock.calls[0] ?? [];
    expect(String(target)).toBe("http://127.0.0.1:11434/api/chat");
    const body = JSON.parse(String(init?.body));
    expect(body).toMatchObject({
      model: "qwen3:8b",
      stream: false,
      think: false,
      options: { temperature: 0 },
      format: { type: "object", additionalProperties: false },
    });
    expect(body).not.toHaveProperty("tools");
    expect(body.messages[1].content).toContain("evidence-synthetic-1");
  });

  it("rejects non-loopback endpoints", () => {
    expect(
      () =>
        new LocalOllamaProvider({
          model: "qwen3:8b",
          endpoint: "https://example.invalid/api/chat",
        }),
    ).toThrow("loopback");
  });

  it("does not expose a provider-controlled error body", async () => {
    const provider = new LocalOllamaProvider({
      model: "qwen3:8b",
      fetchImplementation: async () =>
        new Response('{"error":"sensitive provider detail"}', { status: 500 }),
    });

    const error = await provider.analyze(input).catch((caught: unknown) => caught);
    expect(String(error)).toContain("HTTP 500");
    expect(String(error)).not.toContain("sensitive provider detail");
  });
});

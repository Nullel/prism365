import type { AiAccessGrant, AiAnalysisRequest, AiEvidenceRecord } from "@prism365/contracts";
import { describe, expect, it, vi } from "vitest";
import {
  type AiAnalysisAuditEvent,
  AiAnalysisBroker,
  type AiProvider,
  type AiProviderInput,
} from "../src/index.ts";

const now = new Date("2026-10-13T12:00:00.000Z");

const request: AiAnalysisRequest = {
  tenantId: "tenant-synthetic",
  actorId: "actor-synthetic",
  question: "Which sign-in risks need review?",
  purpose: "securityInvestigation",
  datasets: ["signIns"],
  timeRange: {
    start: "2026-10-01T00:00:00.000Z",
    end: "2026-10-13T11:00:00.000Z",
  },
};

const grant: AiAccessGrant = {
  grantId: "grant-synthetic",
  tenantId: "tenant-synthetic",
  providerId: "provider-synthetic",
  datasets: ["signIns"],
  purposes: ["securityInvestigation"],
  allowExternalProcessing: false,
  allowDirectIdentifiers: false,
  allowRawRecords: false,
  maxLookbackDays: 30,
  maxEvidenceRecords: 100,
  approvedByActorId: "admin-synthetic",
  approvedAt: "2026-10-01T00:00:00.000Z",
  expiresAt: "2026-11-01T00:00:00.000Z",
};

const evidence: AiEvidenceRecord = {
  tenantId: "tenant-synthetic",
  dataset: "signIns",
  evidenceId: "evidence-synthetic-1",
  observedAt: "2026-10-12T08:00:00.000Z",
  source: "Microsoft Graph v1.0 signIns",
  sensitivity: "personLinked",
  fields: { riskLevel: "medium", status: "failure" },
  directIdentifiers: { userPrincipalName: "synthetic.user@example.invalid" },
  limitations: ["Synthetic record; selected time window only."],
  raw: false,
};

function createProvider(
  deployment: "local" | "tenantHosted" | "external" = "local",
  analyze = vi.fn(async (_input: AiProviderInput) => ({
    summary: "One synthetic medium-risk failure needs review.",
    citations: ["evidence-synthetic-1"],
    limitations: ["Only the selected time window was evaluated."],
    model: "synthetic-model",
  })),
): AiProvider {
  return {
    descriptor: {
      id: "provider-synthetic",
      label: "Synthetic provider",
      deployment,
    },
    analyze,
  };
}

function createBroker(options: { provider?: AiProvider; evidenceRecords?: unknown[] } = {}) {
  const provider = options.provider ?? createProvider();
  const read = vi.fn(async () => options.evidenceRecords ?? [evidence]);
  const append = vi.fn(async (_event: AiAnalysisAuditEvent) => undefined);
  return {
    broker: new AiAnalysisBroker({
      provider,
      evidenceReader: { read },
      auditSink: { append },
      clock: () => now,
    }),
    provider,
    read,
    append,
  };
}

describe("AiAnalysisBroker", () => {
  it("sends scoped evidence without direct identifiers by default", async () => {
    const { broker, provider, append } = createBroker();

    const result = await broker.analyze(request, grant);

    const input = vi.mocked(provider.analyze).mock.calls[0]?.[0];
    expect(input?.evidence[0]).not.toHaveProperty("directIdentifiers");
    expect(input?.evidence[0]?.fields).toEqual({ riskLevel: "medium", status: "failure" });
    expect(result.answer.citations).toEqual(["evidence-synthetic-1"]);
    expect(append).toHaveBeenCalledWith(
      expect.objectContaining({
        tenantId: "tenant-synthetic",
        evidenceCount: 1,
        directIdentifiersIncluded: false,
      }),
    );
    expect(JSON.stringify(append.mock.calls)).not.toContain(request.question);
  });

  it("includes separated direct identifiers only with explicit authorization", async () => {
    const provider = createProvider();
    const { broker } = createBroker({ provider });

    await broker.analyze(request, { ...grant, allowDirectIdentifiers: true });

    const input = vi.mocked(provider.analyze).mock.calls[0]?.[0];
    expect(input?.evidence[0]?.directIdentifiers).toEqual({
      userPrincipalName: "synthetic.user@example.invalid",
    });
  });

  it("blocks external processing unless the tenant grant explicitly allows it", async () => {
    const { broker, read } = createBroker({ provider: createProvider("external") });

    await expect(broker.analyze(request, grant)).rejects.toThrow("External AI processing");
    expect(read).not.toHaveBeenCalled();
  });

  it("blocks workforce-performance questions before reading evidence", async () => {
    const { broker, read } = createBroker();

    await expect(
      broker.analyze(
        { ...request, question: "Rank employee productivity from sign-in counts." },
        grant,
      ),
    ).rejects.toThrow("workforce-performance");
    expect(read).not.toHaveBeenCalled();
  });

  it("blocks datasets outside the exact grant snapshot", async () => {
    const { broker, read } = createBroker();

    await expect(
      broker.analyze({ ...request, datasets: ["signIns", "securityIncidents"] }, grant),
    ).rejects.toThrow("outside its access grant");
    expect(read).not.toHaveBeenCalled();
  });

  it("blocks analysis purposes outside the exact grant", async () => {
    const { broker, read } = createBroker();

    await expect(
      broker.analyze({ ...request, purpose: "reportExplanation" }, grant),
    ).rejects.toThrow("purpose is outside");
    expect(read).not.toHaveBeenCalled();
  });

  it("rejects cross-tenant evidence before calling the model", async () => {
    const provider = createProvider();
    const { broker } = createBroker({
      provider,
      evidenceRecords: [{ ...evidence, tenantId: "tenant-other" }],
    });

    await expect(broker.analyze(request, grant)).rejects.toThrow("cross-tenant");
    expect(provider.analyze).not.toHaveBeenCalled();
  });

  it("rejects raw records unless the grant explicitly allows them", async () => {
    const provider = createProvider();
    const { broker } = createBroker({
      provider,
      evidenceRecords: [{ ...evidence, raw: true }],
    });

    await expect(broker.analyze(request, grant)).rejects.toThrow("raw records");
    expect(provider.analyze).not.toHaveBeenCalled();
  });

  it("enforces a serialized payload limit before calling the model", async () => {
    const provider = createProvider();
    const read = vi.fn(async () => [
      {
        ...evidence,
        fields: { oversizedSyntheticField: "x".repeat(2_000) },
      },
    ]);
    const broker = new AiAnalysisBroker({
      provider,
      evidenceReader: { read },
      auditSink: { append: vi.fn(async () => undefined) },
      clock: () => now,
      hardPayloadBytes: 1_024,
    });

    await expect(broker.analyze(request, grant)).rejects.toThrow("payload safety limit");
    expect(provider.analyze).not.toHaveBeenCalled();
  });

  it("rejects citations that were not in the evidence bundle", async () => {
    const provider = createProvider(
      "local",
      vi.fn(async () => ({
        summary: "Unsupported assertion.",
        citations: ["evidence-invented"],
        limitations: ["Synthetic test."],
        model: "synthetic-model",
      })),
    );
    const { broker, append } = createBroker({ provider });

    await expect(broker.analyze(request, grant)).rejects.toThrow("not in its evidence bundle");
    expect(append).not.toHaveBeenCalled();
  });

  it("rejects generated workforce-performance claims", async () => {
    const provider = createProvider(
      "local",
      vi.fn(async () => ({
        summary: "This proves employee productivity is low.",
        citations: ["evidence-synthetic-1"],
        limitations: ["Synthetic test."],
        model: "synthetic-model",
      })),
    );
    const { broker } = createBroker({ provider });

    await expect(broker.analyze(request, grant)).rejects.toThrow("workforce-performance");
  });
});

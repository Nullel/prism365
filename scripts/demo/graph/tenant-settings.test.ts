import { type AuditResult, tenantSettingsAudit } from "@prism365/audits";
import { GraphHttpClient } from "@prism365/graph";
import { scopeGraph } from "@prism365/modules";
import { describe, expect, it } from "vitest";
import { createSyntheticGraph, SYNTHETIC_TENANT_ID } from "../tenant.ts";

// The synthetic tenant through the real Tenant settings audit, the way the recorder runs it: the
// demo must show the settings the story plants, so a change that quietly empties the page fails
// here and not in a recording.
const now = new Date("2026-10-22T12:00:00Z");

async function run(): Promise<{ result: AuditResult; paths: string[] }> {
  const paths: string[] = [];
  const graph = createSyntheticGraph({ now });
  const client = new GraphHttpClient({
    getAccessToken: async () => "synthetic",
    fetchImplementation: (async (input: Parameters<typeof fetch>[0], init?: RequestInit) => {
      paths.push(new URL(String(input)).pathname);
      return graph(input, init);
    }) as typeof fetch,
    maxResponseBytes: 8_388_608,
    maxRetries: 0,
    sleep: async () => {},
  });
  const signal = new AbortController().signal;
  const result = await tenantSettingsAudit.run({
    graph: scopeGraph(client, tenantSettingsAudit.packs, () => {}, signal),
    now,
    granted: new Set(["Policy.Read.All"]),
    signal,
    scope: null,
    tenantId: SYNTHETIC_TENANT_ID,
    stored: {
      async latestRun() {
        return null;
      },
      async openFindings() {
        return [];
      },
    },
  });
  return { result, paths };
}

const statusOf = (result: AuditResult, key: string) =>
  result.rows.find((row) => row.key === key)?.cells.status;

describe("the synthetic Tenant settings", () => {
  it("answers all five paths, completely", async () => {
    const { result, paths } = await run();
    expect(paths.sort()).toEqual([
      "/v1.0/policies/adminConsentRequestPolicy",
      "/v1.0/policies/authenticationMethodsPolicy",
      "/v1.0/policies/authorizationPolicy",
      "/v1.0/policies/identitySecurityDefaultsEnforcementPolicy",
      "/v1.0/policies/permissionGrantPolicies",
    ]);
    expect(result.complete).toBe(true);
    expect(result.rows).toHaveLength(22);
    expect(result.rows.filter((row) => row.cells.status === "Not read")).toEqual([]);
  });

  it("plants the settings the demo story needs, and meets the rest", async () => {
    const { result } = await run();
    const review = [
      "method-sms",
      "method-tap",
      "registration-campaign",
      "guest-invites",
      "guest-role",
      "create-apps",
      "admin-consent-workflow",
    ];
    for (const key of review) expect(statusOf(result, key), key).toBe("Review");
    for (const key of ["method-voice", "method-fido2", "method-authenticator", "policy-migration"])
      expect(statusOf(result, key), key).toBe("Meets");
    expect(statusOf(result, "email-verified-join")).toBe("Meets");
    expect(statusOf(result, "user-consent")).toBe("Meets");
    expect(result.rows.find((row) => row.key === "user-consent")?.cells.value).toContain(
      "verified publishers",
    );
  });

  it("has security defaults off, which waits for a Conditional Access run", async () => {
    const { result } = await run();
    const row = result.rows.find((item) => item.key === "security-defaults");
    expect(row?.cells.value).toBe("Off. Run Conditional Access policies first");
    expect(row?.cells.status).toBe("Review");
  });
});

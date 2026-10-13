export type FindingPurpose =
  | "security"
  | "operations"
  | "compliance"
  | "adoption"
  | "workforce-performance";
export type FindingSubject = "tenant" | "service" | "resource" | "person";

export interface FindingPolicyInput {
  title: string;
  summary: string;
  purpose: FindingPurpose;
  subject: FindingSubject;
  hasSource: boolean;
  hasTimeWindow: boolean;
  hasLimitations: boolean;
}

export type FindingPolicyDecision =
  | { allowed: true; warnings: string[] }
  | { allowed: false; reason: string };

const unsupportedWorkforceClaims = [
  /productiv(?:e|ity)/i,
  /employee\s+(?:performance|engagement|effort)/i,
  /(?:top|low)\s+performer/i,
  /lazy\s+(?:employee|worker|user)/i,
  /idle\s+(?:employee|worker)/i,
  /rank\s+(?:employees|workers|users)\b/i,
  /(?:worker|employee)\s+score/i,
  /(?:attendance|engagement)\s+score/i,
];

export function containsUnsupportedWorkforceClaim(text: string): boolean {
  return unsupportedWorkforceClaims.some((pattern) => pattern.test(text));
}

export function evaluateFinding(input: FindingPolicyInput): FindingPolicyDecision {
  if (input.purpose === "workforce-performance") {
    return {
      allowed: false,
      reason:
        "TenantLens does not use Microsoft 365 activity as evidence of workforce performance.",
    };
  }

  const language = `${input.title} ${input.summary}`;
  if (containsUnsupportedWorkforceClaim(language)) {
    return {
      allowed: false,
      reason: "Finding language makes an unsupported productivity or employee-performance claim.",
    };
  }

  if (!input.hasSource || !input.hasTimeWindow) {
    return {
      allowed: false,
      reason: "A finding must identify its evidence source and observation time window.",
    };
  }

  const warnings: string[] = [];
  if (!input.hasLimitations) {
    warnings.push("Add evidence limitations or coverage notes before presenting this finding.");
  }
  if (input.subject === "person" && input.purpose === "adoption") {
    warnings.push("Prefer tenant- or service-level aggregation for adoption reporting.");
  }

  return { allowed: true, warnings };
}

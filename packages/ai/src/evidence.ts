import type { AiDataset, AiEvidenceRecord } from "@prism365/contracts";

type Rule = "count" | "boolean" | readonly string[];

/** Explicit review boundary: new datasets do not authorize arbitrary source fields. */
const fieldRules: Record<AiDataset, Record<string, Rule>> = {
  findings: { severity: ["info", "low", "medium", "high", "critical"], affectedCount: "count" },
  directory: { accountEnabled: "boolean", userType: ["Member", "Guest"], userCount: "count" },
  licenses: {
    consumedUnits: "count",
    enabledUnits: "count",
    availableUnits: "count",
    licenseCount: "count",
    servicePlanCount: "count",
    complete: "boolean",
  },
  serviceHealth: {
    status: [
      "serviceOperational",
      "investigating",
      "serviceDegradation",
      "serviceInterruption",
      "restoringService",
      "extendedRecovery",
      "investigationSuspended",
      "resolved",
      "postIncidentReviewPublished",
    ],
    affectedServices: "count",
  },
  signIns: {
    riskLevel: ["none", "low", "medium", "high", "hidden", "unknownFutureValue"],
    status: ["success", "failure", "interrupted", "unknown"],
    failureCount: "count",
    successCount: "count",
    errorCode: "count",
    isInteractive: "boolean",
    complete: "boolean",
  },
  directoryAudit: {
    result: ["success", "failure", "timeout", "unknownFutureValue"],
    eventCount: "count",
  },
  usageReports: { activeUsers: "count", licensedUsers: "count", reportPeriodDays: "count" },
  conditionalAccess: {
    state: ["enabled", "disabled", "enabledForReportingButNotEnforced"],
    policyCount: "count",
  },
  identityRisk: {
    riskLevel: ["none", "low", "medium", "high", "hidden", "unknownFutureValue"],
    isDeleted: "boolean",
    isProcessing: "boolean",
    affectedCount: "count",
  },
  secureScore: { currentScore: "count", maxScore: "count" },
  securityIncidents: {
    severity: ["unknown", "informational", "low", "medium", "high", "unknownFutureValue"],
    status: ["active", "resolved", "redirected", "unknownFutureValue"],
    incidentCount: "count",
  },
  devices: {
    complianceState: [
      "unknown",
      "compliant",
      "noncompliant",
      "conflict",
      "error",
      "inGracePeriod",
      "configManager",
    ],
    deviceCount: "count",
    isEncrypted: "boolean",
  },
  configuration: { enabled: "boolean", configuredCount: "count", missingCount: "count" },
  purviewAudit: { eventCount: "count", partialResults: "boolean" },
  managementActivity: { eventCount: "count", partialResults: "boolean" },
};

export function minimizeAiFields(record: AiEvidenceRecord): AiEvidenceRecord["fields"] {
  const fields: AiEvidenceRecord["fields"] = {};
  for (const [key, rule] of Object.entries(fieldRules[record.dataset])) {
    const value = record.fields[key];
    if (rule === "count" && typeof value === "number" && Number.isFinite(value) && value >= 0)
      fields[key] = value;
    else if (rule === "boolean" && typeof value === "boolean") fields[key] = value;
    else if (Array.isArray(rule) && typeof value === "string" && rule.includes(value))
      fields[key] = value;
  }
  return fields;
}

/** Heuristic defense for operator-authored text; unknown names are not reliably detectable. */
export function redactAiQuestion(question: string, records: AiEvidenceRecord[]): string {
  let redacted = question;
  const identifiers = records.flatMap((record) => Object.values(record.directIdentifiers));
  for (const identifier of [...new Set(identifiers)].sort((a, b) => b.length - a.length)) {
    if (!identifier) continue;
    redacted = redacted.replaceAll(
      new RegExp(escapeRegExp(identifier), "gi"),
      "[identifier removed]",
    );
  }
  return redacted
    .replace(/\b[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi, "[email removed]")
    .replace(/\b(?:\d{1,3}\.){3}\d{1,3}\b/g, "[IP removed]")
    .replace(
      /\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b/gi,
      "[identifier removed]",
    );
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export const AI_EVIDENCE_LIMITATIONS = [
  "Only reviewed scalar fields from selected datasets are supplied; source text and unknown fields are omitted.",
  "Source records may be incomplete or delayed. Source limitations remain available in Prism365, not in this minimized bundle.",
  "Question redaction is heuristic: unknown personal names or other sensitive text typed by the operator may remain.",
] as const;

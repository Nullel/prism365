import { z } from "zod";
import { personHistorySchema } from "./people.ts";

export * from "./attention.ts";
export * from "./audits.ts";
export * from "./canaries.ts";
export * from "./changes.ts";
export * from "./delivery.ts";
export * from "./digest.ts";
export * from "./handoff.ts";
export * from "./insights.ts";
export * from "./investigation.ts";
export * from "./licenses.ts";
export * from "./people.ts";
export * from "./setup.ts";
export * from "./text.ts";

export const sessionSchema = z.object({
  mode: z.enum(["demo", "live"]),
  configured: z.boolean(),
  authenticated: z.boolean(),
  displayName: z.string().nullable(),
  csrfToken: z.string().nullable(),
  enabledFeatures: z.array(z.string()),
  missingConfiguration: z.array(z.string()),
});
export type PrismSession = z.infer<typeof sessionSchema>;

export const aiStatusSchema = z.object({
  available: z.boolean(),
  provider: z.string().nullable(),
  datasets: z.array(z.enum(["directory", "licenses", "signIns"])),
  externalProcessing: z.literal(false),
  reason: z.string().nullable(),
});
export const aiGrantSummarySchema = z.object({
  grantId: z.uuid(),
  datasets: z.array(z.enum(["directory", "licenses", "signIns"])),
  expiresAt: z.string().datetime(),
});
export const governedAiAnalysisSchema = z.object({
  requestId: z.uuid(),
  providerId: z.string(),
  providerDeployment: z.enum(["local", "tenantHosted", "external"]),
  generatedAt: z.string().datetime(),
  datasets: z.array(z.string()),
  evidenceCount: z.number().int().nonnegative(),
  citationReferences: z.array(z.object({ citationId: z.string(), evidenceId: z.string() })),
  answer: z.object({
    summary: z.string(),
    citations: z.array(z.string()),
    limitations: z.array(z.string()),
    model: z.string(),
  }),
});

export const severitySchema = z.enum(["info", "low", "medium", "high", "critical"]);
export type Severity = z.infer<typeof severitySchema>;

export const dataStateSchema = z.enum(["fresh", "stale", "partial", "unavailable", "demo"]);
export type DataState = z.infer<typeof dataStateSchema>;

export const trendSchema = z.object({
  direction: z.enum(["up", "down", "flat"]),
  value: z.number(),
  label: z.string().min(1),
  sentiment: z.enum(["positive", "negative", "neutral"]),
});
export type Trend = z.infer<typeof trendSchema>;

export const metricSchema = z.object({
  id: z.string().min(1),
  label: z.string().min(1),
  value: z.number(),
  displayValue: z.string().min(1),
  state: dataStateSchema,
  trend: trendSchema,
});
export type Metric = z.infer<typeof metricSchema>;

export const chartPointSchema = z.object({
  label: z.string().min(1),
  safe: z.number().nonnegative(),
  review: z.number().nonnegative(),
  elevated: z.number().nonnegative(),
});
export type ChartPoint = z.infer<typeof chartPointSchema>;

export const activityItemSchema = z.object({
  id: z.string().min(1),
  title: z.string().min(1),
  detail: z.string().min(1),
  category: z.enum(["identity", "sharing", "configuration", "system"]),
  severity: severitySchema,
  observedAt: z.string().datetime(),
  evidence: z.string().min(1),
});
export type ActivityItem = z.infer<typeof activityItemSchema>;

export const dashboardSummarySchema = z.object({
  mode: z.literal("demo"),
  generatedAt: z.string().datetime(),
  timeWindow: z.string().min(1),
  state: dataStateSchema,
  metrics: z.array(metricSchema),
  signInTrend: z.array(chartPointSchema),
  recentActivity: z.array(activityItemSchema),
});
export type DashboardSummary = z.infer<typeof dashboardSummarySchema>;

export const permissionPackSchema = z.object({
  id: z.string().min(1),
  label: z.string().min(1),
  source: z.enum(["microsoftGraph", "officeManagementActivity", "exchangeOnlinePowerShell"]),
  permission: z.string().min(1),
  mode: z.enum(["delegated", "application", "either"]),
  phase: z.enum(["mvp", "experimental", "future"]),
  defaultEnabled: z.boolean(),
  purpose: z.string().min(1),
  deniedBehavior: z.string().min(1),
  clouds: z.array(z.enum(["worldwide", "gcc", "usGovernmentL4", "usGovernmentL5", "china"])),
  prerequisites: z.array(z.string().min(1)),
  roleNotes: z.string().min(1),
  licenseNotes: z.string().min(1),
  documentationUrl: z.string().url(),
  relatedDocumentationUrls: z.array(z.string().url()).optional(),
  documentationLastVerified: z.string().date(),
});
export type PermissionPack = z.infer<typeof permissionPackSchema>;

export const capabilityStateSchema = z.enum([
  "available",
  "missingConsent",
  "missingRole",
  "missingLicense",
  "unsupportedCloud",
  "notEnabled",
  "notImplemented",
  "providerUnavailable",
  "unknown",
]);
export type CapabilityState = z.infer<typeof capabilityStateSchema>;

export const microsoftCloudProfileSchema = z.object({
  id: z.enum(["worldwide", "gcc", "usGovernmentL4", "usGovernmentL5", "china"]),
  label: z.string().min(1),
  authorityHost: z.string().url(),
  graphBaseUrl: z.string().url(),
  managementActivityBaseUrl: z.string().url().nullable(),
});
export type MicrosoftCloudProfile = z.infer<typeof microsoftCloudProfileSchema>;

export const capabilitySummarySchema = z.object({
  dataMode: z.enum(["demo", "live"]),
  writesEnabled: z.literal(false),
  tenantConnectionEnabled: z.boolean(),
  permissionPacks: z.array(permissionPackSchema),
  cloudProfiles: z.array(microsoftCloudProfileSchema),
  ai: z.object({
    analysisEnabled: z.boolean(),
    providerConfigured: z.boolean(),
    writeToolsEnabled: z.literal(false),
    supportedDeployments: z.array(z.enum(["local", "tenantHosted", "external"])),
  }),
});
export type CapabilitySummary = z.infer<typeof capabilitySummarySchema>;

export const aiDatasetSchema = z.enum([
  "findings",
  "directory",
  "licenses",
  "serviceHealth",
  "signIns",
  "directoryAudit",
  "usageReports",
  "conditionalAccess",
  "identityRisk",
  "secureScore",
  "securityIncidents",
  "devices",
  "configuration",
  "purviewAudit",
  "managementActivity",
]);
export type AiDataset = z.infer<typeof aiDatasetSchema>;

export const aiDatasets = aiDatasetSchema.options;

export const aiPurposeSchema = z.enum([
  "securityInvestigation",
  "operations",
  "compliance",
  "configurationReview",
  "reportExplanation",
]);
export type AiPurpose = z.infer<typeof aiPurposeSchema>;

export const aiProviderDeploymentSchema = z.enum(["local", "tenantHosted", "external"]);
export type AiProviderDeployment = z.infer<typeof aiProviderDeploymentSchema>;

const uniqueAiDatasetsSchema = z
  .array(aiDatasetSchema)
  .min(1)
  .refine((datasets) => new Set(datasets).size === datasets.length, {
    message: "AI datasets must be unique.",
  });

const uniqueAiPurposesSchema = z
  .array(aiPurposeSchema)
  .min(1)
  .refine((purposes) => new Set(purposes).size === purposes.length, {
    message: "AI purposes must be unique.",
  });

export const aiTimeRangeSchema = z
  .object({
    start: z.string().datetime(),
    end: z.string().datetime(),
  })
  .refine((range) => Date.parse(range.start) <= Date.parse(range.end), {
    message: "AI time range start must not be after its end.",
  });
export type AiTimeRange = z.infer<typeof aiTimeRangeSchema>;

export const aiAnalysisRequestSchema = z.object({
  tenantId: z.string().min(1),
  actorId: z.string().min(1),
  question: z.string().trim().min(1).max(4_000),
  purpose: aiPurposeSchema,
  datasets: uniqueAiDatasetsSchema,
  timeRange: aiTimeRangeSchema,
});
export type AiAnalysisRequest = z.infer<typeof aiAnalysisRequestSchema>;

export const aiAccessGrantSchema = z
  .object({
    grantId: z.string().min(1),
    tenantId: z.string().min(1),
    providerId: z.string().min(1),
    datasets: uniqueAiDatasetsSchema,
    purposes: uniqueAiPurposesSchema,
    allowExternalProcessing: z.boolean(),
    allowDirectIdentifiers: z.boolean(),
    allowRawRecords: z.boolean(),
    maxLookbackDays: z.number().int().positive().max(3_660),
    maxEvidenceRecords: z.number().int().positive().max(2_000),
    approvedByActorId: z.string().min(1),
    approvedAt: z.string().datetime(),
    expiresAt: z.string().datetime(),
  })
  .refine((grant) => Date.parse(grant.approvedAt) < Date.parse(grant.expiresAt), {
    message: "AI access grant expiration must be after approval.",
  });
export type AiAccessGrant = z.infer<typeof aiAccessGrantSchema>;

export const aiEvidenceSensitivitySchema = z.enum([
  "aggregate",
  "tenantMetadata",
  "personLinked",
  "sensitiveSecurity",
]);
export type AiEvidenceSensitivity = z.infer<typeof aiEvidenceSensitivitySchema>;

export const aiEvidenceRecordSchema = z.object({
  tenantId: z.string().min(1),
  dataset: aiDatasetSchema,
  evidenceId: z.string().min(1),
  observedAt: z.string().datetime(),
  source: z.string().min(1),
  sensitivity: aiEvidenceSensitivitySchema,
  fields: z.record(z.string(), z.json()),
  directIdentifiers: z.record(z.string(), z.string()),
  limitations: z.array(z.string().min(1)),
  raw: z.boolean(),
});
export type AiEvidenceRecord = z.infer<typeof aiEvidenceRecordSchema>;

export const aiProviderResultSchema = z.object({
  summary: z.string().trim().min(1).max(20_000),
  citations: z.array(z.string().min(1)),
  limitations: z.array(z.string().min(1)).min(1),
  model: z.string().min(1),
  providerRequestId: z.string().min(1).optional(),
});
export type AiProviderResult = z.infer<typeof aiProviderResultSchema>;

// Module runtime: findings, collector runs, and background collection status.
export const findingStateSchema = z.enum(["open", "acknowledged", "resolved"]);
export type FindingState = z.infer<typeof findingStateSchema>;

export const findingSchema = z.object({
  id: z.uuid(),
  module: z.string().min(1),
  ruleId: z.string().min(1),
  ruleVersion: z.number().int().positive(),
  purpose: z.enum(["security", "operations", "compliance", "adoption"]),
  severity: severitySchema,
  score: z.number().int().min(0).max(1_000),
  subject: z.object({
    kind: z.enum(["tenant", "service", "resource", "person"]),
    id: z.string().min(1),
    label: z.string().nullable(),
  }),
  title: z.string().min(1),
  summary: z.string().min(1),
  evidence: z.array(
    z.object({
      source: z.string().min(1),
      ref: z.string().min(1),
      observedAt: z.string().datetime({ offset: true }),
      fields: z.record(z.string(), z.union([z.string(), z.number(), z.boolean(), z.null()])),
    }),
  ),
  limitations: z.array(z.string().min(1)),
  windowStart: z.string().datetime({ offset: true }),
  windowEnd: z.string().datetime({ offset: true }),
  state: findingStateSchema,
  occurrences: z.number().int().positive(),
  firstSeen: z.string().datetime({ offset: true }),
  lastSeen: z.string().datetime({ offset: true }),
  stateChangedAt: z.string().datetime({ offset: true }).nullable(),
});
export type Finding = z.infer<typeof findingSchema>;

export const findingListSchema = z.object({
  mode: z.enum(["demo", "live"]),
  items: z.array(findingSchema),
});
export type FindingList = z.infer<typeof findingListSchema>;

// One account's stored record: its sign-in history (null when none is stored) and the findings
// about it.
export const personRecordSchema = z.object({
  history: personHistorySchema.nullable(),
  findings: z.array(findingSchema),
});
export type PersonRecord = z.infer<typeof personRecordSchema>;

export const findingStateRequestSchema = z.object({ state: findingStateSchema });

export const moduleRunSchema = z.object({
  status: z.enum(["running", "succeeded", "partial", "failed", "abandoned"]),
  credentialKind: z.enum(["delegated", "application"]),
  startedAt: z.string().datetime({ offset: true }),
  finishedAt: z.string().datetime({ offset: true }).nullable(),
  fetched: z.number().int().nonnegative(),
  inserted: z.number().int().nonnegative(),
  findings: z.number().int().nonnegative(),
  skipped: z.number().int().nonnegative(),
  errorCategory: z.string().nullable(),
});
export type ModuleRun = z.infer<typeof moduleRunSchema>;

export const collectionStatusSchema = z.object({
  available: z.boolean(),
  state: z.enum(["stopped", "collecting"]),
  credentialKind: z.enum(["delegated", "application"]).nullable(),
  boundAt: z.string().datetime({ offset: true }).nullable(),
  lastStopReason: z
    .enum(["operator_stopped", "signed_out", "credential_unavailable", "replaced", "shutdown"])
    .nullable(),
  durableStorage: z.boolean(),
  reason: z.string().nullable(),
  // An administrator chose to collect and nobody stopped it: collection resumes on that
  // administrator's next signed-in request.
  resumes: z.boolean().default(false),
});
export type CollectionStatus = z.infer<typeof collectionStatusSchema>;

export const moduleListSchema = z.object({
  mode: z.enum(["demo", "live"]),
  collection: collectionStatusSchema,
  modules: z.array(
    z.object({
      id: z.string().min(1),
      title: z.string().min(1),
      description: z.string().min(1),
      packs: z.array(z.string().min(1)),
      permissions: z.array(z.string().min(1)),
      missingPermissions: z.array(z.string().min(1)),
      collectors: z.array(
        z.object({
          id: z.string().min(1),
          title: z.string().min(1),
          cadenceSeconds: z.number().int().positive(),
          lastRun: moduleRunSchema.nullable(),
        }),
      ),
    }),
  ),
});
export type ModuleList = z.infer<typeof moduleListSchema>;

// Sign-in audit module: tenant-level aggregates only. Addresses are reduced to their network.
export const signinSummarySchema = z.object({
  days: z.number().int().min(1).max(90),
  totals: z.object({
    successes: z.number().int().nonnegative(),
    failures: z.number().int().nonnegative(),
    blockedByPolicy: z.number().int().nonnegative(),
    legacySuccesses: z.number().int().nonnegative(),
    accounts: z.number().int().nonnegative(),
  }),
  countries: z.array(
    z.object({
      country: z.string().min(1).max(8),
      successes: z.number().int().nonnegative(),
      failures: z.number().int().nonnegative(),
    }),
  ),
  failingNetworks: z.array(
    z.object({
      network: z.string().min(1).max(64),
      failures: z.number().int().nonnegative(),
      accounts: z.number().int().nonnegative(),
    }),
  ),
  coverage: z.object({
    events: z.number().int().nonnegative(),
    firstEvent: z.string().datetime({ offset: true }).nullable(),
    lastEvent: z.string().datetime({ offset: true }).nullable(),
    lastCollected: z.string().datetime({ offset: true }).nullable(),
  }),
  history: z.object({
    phase: z.enum(["none", "backfill", "live"]),
    from: z.string().datetime({ offset: true }).nullable(),
    through: z.string().datetime({ offset: true }).nullable(),
  }),
});
export type SigninSummary = z.infer<typeof signinSummarySchema>;

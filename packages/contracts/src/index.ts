import { z } from "zod";

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
  dataMode: z.literal("demo"),
  writesEnabled: z.literal(false),
  tenantConnectionEnabled: z.literal(false),
  permissionPacks: z.array(permissionPackSchema),
  cloudProfiles: z.array(microsoftCloudProfileSchema),
  ai: z.object({
    analysisEnabled: z.literal(false),
    providerConfigured: z.literal(false),
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

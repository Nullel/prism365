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
  permission: z.string().min(1),
  mode: z.enum(["delegated", "application", "either"]),
  phase: z.enum(["mvp", "experimental", "future"]),
  defaultEnabled: z.boolean(),
  purpose: z.string().min(1),
  deniedBehavior: z.string().min(1),
});
export type PermissionPack = z.infer<typeof permissionPackSchema>;

export const capabilitySummarySchema = z.object({
  dataMode: z.literal("demo"),
  writesEnabled: z.literal(false),
  tenantConnectionEnabled: z.literal(false),
  permissionPacks: z.array(permissionPackSchema),
});
export type CapabilitySummary = z.infer<typeof capabilitySummarySchema>;

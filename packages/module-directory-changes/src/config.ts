import { z } from "zod";

// Operator settings for the directory change feed. There are no environment variables for them
// yet; the defaults are what a deployment runs with, and tests narrow them.

export const directoryChangesConfigSchema = z.object({
  // Microsoft publishes some records late. Every run re-reads this far behind where it stopped,
  // and records it has already stored change nothing.
  lateArrivalMinutes: z.number().int().min(5).max(1_440).default(120),
  // Records one time window may hold. A window with more is split in half; one already at the
  // minimum size keeps what it read and is recorded as a gap.
  windowRecordLimit: z.number().int().min(2).max(10_000).default(5_000),
  // Time a run spends reading before it saves its place and asks to continue. No request starts
  // after it. The scheduler stops a run at 120 s and one request may take 90 s, so 25 s leaves
  // the last request room to finish.
  runBudgetSeconds: z.number().int().min(1).max(100).default(25),
});

export type DirectoryChangesConfig = z.infer<typeof directoryChangesConfigSchema>;

export const defaultDirectoryChangesConfig: DirectoryChangesConfig =
  directoryChangesConfigSchema.parse({});

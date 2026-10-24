import { z } from "zod";

// Operator settings for the directory change feed. Deployments set the late-arrival margin and
// the retention with DIRECTORY_CHANGES_LATE_MINUTES and DIRECTORY_CHANGES_RETENTION_DAYS; the
// other values are fixed defaults that tests narrow.

export const directoryChangesConfigSchema = z.object({
  // Microsoft publishes some records late. The settle lane reads every window again once it is
  // this old, so a record published up to this late is collected; one published later is not.
  lateArrivalMinutes: z.number().int().min(5).max(1_440).default(240),
  // How long a stored change is kept before the scheduler's daily retention deletes it.
  retentionDays: z.number().int().min(365).max(3_650).default(3_650),
  // Records one time window may hold. A window with more is split in half, down to one
  // millisecond; a one-millisecond window with more keeps what it read and is recorded as a gap.
  windowRecordLimit: z.number().int().min(2).max(10_000).default(5_000),
  // Time a run spends reading before it saves its place and asks to continue. No request starts
  // after it. The scheduler stops a run at 120 s and one request may take 90 s, so 25 s leaves
  // the last request room to finish.
  runBudgetSeconds: z.number().int().min(1).max(100).default(25),
});

export type DirectoryChangesConfig = z.infer<typeof directoryChangesConfigSchema>;

export const defaultDirectoryChangesConfig: DirectoryChangesConfig =
  directoryChangesConfigSchema.parse({});

// Reads DIRECTORY_CHANGES_LATE_MINUTES (60 to 1440) and DIRECTORY_CHANGES_RETENTION_DAYS (365 to
// 3650). A value outside its range, or one that is not a whole number, fails startup rather than
// silently keeping less history or missing late records.
export function directoryChangesConfigFromEnv(
  env: Record<string, string | undefined>,
): DirectoryChangesConfig {
  const read = (name: string, min: number, max: number): number | undefined => {
    const raw = env[name]?.trim();
    if (!raw) return undefined;
    const value = Number(raw);
    if (!/^\d+$/.test(raw) || value < min || value > max)
      throw new Error(`${name} must be a whole number from ${min} to ${max}.`);
    return value;
  };
  const late = read("DIRECTORY_CHANGES_LATE_MINUTES", 60, 1_440);
  const retention = read("DIRECTORY_CHANGES_RETENTION_DAYS", 365, 3_650);
  return directoryChangesConfigSchema.parse({
    ...(late === undefined ? {} : { lateArrivalMinutes: late }),
    ...(retention === undefined ? {} : { retentionDays: retention }),
  });
}

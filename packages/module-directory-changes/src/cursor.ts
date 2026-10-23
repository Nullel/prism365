import { z } from "zod";

// Where the collector is in Microsoft's directory audit log, saved after every window it
// finishes. Everything from `start` up to `through` has been read completely at least once.
// "history" is the first load, read oldest first from 30 days back; "live" follows new records
// and re-reads a late-arrival margin behind `through` on every run.

export const MINUTE = 60_000;
export const DAY = 24 * 60 * MINUTE;
export const HISTORY_DAYS = 30;
export const MAX_SPAN_MS = DAY;
export const MIN_SPAN_MS = MINUTE;

export const cursorSchema = z
  .object({
    v: z.literal(1),
    phase: z.enum(["history", "live"]),
    start: z.string().datetime(),
    through: z.string().datetime(),
    // The window size in use: halved when a window holds too many records or Microsoft does not
    // answer for it, doubled back after each window that reads cleanly.
    span: z.number().int().min(MIN_SPAN_MS).max(MAX_SPAN_MS).optional(),
    // The start of the window being read when the cursor was last saved. Still set when the next
    // run starts means the run was stopped mid-window (its time limit, or a restart), so that
    // window counts as failed instead of being retried at the same size forever.
    trying: z.string().datetime().optional(),
    // How many runs in a row a minimum-size window starting at `at` could not be read.
    strikes: z
      .object({ at: z.string().datetime(), count: z.number().int().min(1).max(100) })
      .optional(),
  })
  .refine((value) => Date.parse(value.start) <= Date.parse(value.through), {
    message: "The cursor's start must not be after its position.",
  });

export type DirectoryChangesCursor = z.infer<typeof cursorSchema>;

// A cursor that is missing, from a later version, or unreadable is null: collection starts the
// history again, and the records it reads twice change nothing.
export function parseCursor(value: string | null | undefined): DirectoryChangesCursor | null {
  if (!value) return null;
  try {
    const parsed = cursorSchema.safeParse(JSON.parse(value));
    return parsed.success ? parsed.data : null;
  } catch {
    return null;
  }
}

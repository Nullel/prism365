import { z } from "zod";

// Where the collector is in Microsoft's directory audit log, saved after every window it
// finishes. "history" is the first load, read oldest first from 30 days back; "live" follows new
// records.
//
// Two lanes read forward. The main lane is at `through`: everything from `start` up to it has
// been read completely at least once. The settle lane is at `settle.at`, never past `through`:
// everything before it has also been read at least the late-arrival margin after it happened,
// so records Microsoft publishes late are in. Each lane saves its own position, so neither has
// to start over when a run ends early. The settle lane reads every window again, including the
// ones the main lane read when they were already old: one empty answer from Microsoft, which
// Learn says can happen under load, is then not the only read of that window.

export const MINUTE = 60_000;
export const DAY = 24 * 60 * MINUTE;
export const HISTORY_DAYS = 30;
export const MAX_SPAN_MS = DAY;
// Microsoft's filter compares times to the millisecond, so a window can be split down to one.
export const MIN_SPAN_MS = 1;

const laneState = {
  // The window size in use: halved when a window holds too many records or Microsoft does not
  // answer for it, and doubled after each quiet window once the lane is past `hold`.
  span: z.number().int().min(MIN_SPAN_MS).max(MAX_SPAN_MS).optional(),
  hold: z.string().datetime().optional(),
  // The start of the window being read when the cursor was last saved. Still set when the next
  // run starts means the run was stopped mid-window (its time limit, or a restart), so that
  // window counts as failed instead of being retried at the same size forever.
  trying: z.string().datetime().optional(),
  // How many runs in a row a minimum-size window starting at `at` could not be read.
  strikes: z
    .object({ at: z.string().datetime(), count: z.number().int().min(1).max(100) })
    .optional(),
};

export const cursorSchema = z
  .object({
    v: z.literal(1),
    phase: z.enum(["history", "live"]),
    start: z.string().datetime(),
    through: z.string().datetime(),
    ...laneState,
    settle: z.object({ at: z.string().datetime(), ...laneState }),
    // Where a span that fell out of Microsoft's retention before it was read began. Kept while
    // the main lane stays behind, so the one finding about that span widens instead of a new one
    // being raised on every run.
    missed: z.string().datetime().optional(),
  })
  .refine(
    (value) =>
      Date.parse(value.start) <= Date.parse(value.settle.at) &&
      Date.parse(value.settle.at) <= Date.parse(value.through),
    { message: "A cursor's lanes must lie between its start and its position." },
  );

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

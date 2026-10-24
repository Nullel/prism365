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
  // The window size was just doubled. When a window of a size the lane had just grown to fails,
  // windows stay at most `ceiling.span` until the lane passes `ceiling.until`, and each such
  // failure in a row (`backoff`) doubles how far that is, so a size Microsoft cannot answer is
  // not tried again every other window.
  grown: z.literal(true).optional(),
  backoff: z.number().int().min(1).max(10).optional(),
  ceiling: z
    .object({
      span: z.number().int().min(MIN_SPAN_MS).max(MAX_SPAN_MS),
      until: z.string().datetime(),
    })
    .optional(),
  // The start of the window being read when the cursor was last saved. Still set when the next
  // run starts means the run was stopped mid-window (its time limit, or a restart), so that
  // window counts as failed instead of being retried at the same size forever.
  trying: z.string().datetime().optional(),
  // How many times in a row a minimum-size window starting at `at` could not be read, and when
  // the first of those tries was.
  strikes: z
    .object({
      at: z.string().datetime(),
      count: z.number().int().min(0).max(100),
      since: z.string().datetime().optional(),
    })
    .optional(),
  // A try at the minimum-size window starting at `at` that failed and could not be counted: the
  // run ran out of time before it knew whether Microsoft was answering anything. The next run
  // finds out first.
  uncounted: z
    .object({
      at: z.string().datetime(),
      failure: z.enum(["slow", "failing", "refused", "unreadable"]),
    })
    .optional(),
};

const span = z.object({ from: z.string().datetime(), to: z.string().datetime() });

export const cursorSchema = z
  .object({
    v: z.literal(1),
    phase: z.enum(["history", "live"]),
    start: z.string().datetime(),
    through: z.string().datetime(),
    ...laneState,
    settle: z.object({ at: z.string().datetime(), ...laneState }),
    // The last span that fell out of Microsoft's retention before the main lane read it, and
    // the last one the settle lane had to skip before reading it again. A new span that starts
    // within the margin of the last one's end widens the same finding instead of raising another.
    missed: span.optional(),
    unsettled: span.optional(),
    // The spans the main lane read while they were younger than the late-arrival margin and the
    // settle lane has not read again yet, oldest first. At most 8 are written (16 are read). What the main lane read when it was
    // already old enough to have settled (the history load, a catch-up after collection was
    // stopped) is not in them, so the settle lane losing it to Microsoft's retention loses
    // nothing.
    fresh: z.array(span).max(16).optional(),
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

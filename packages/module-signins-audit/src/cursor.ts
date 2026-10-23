import { z } from "zod";

// Where the sign-in collector is in Microsoft's log, saved after every window it finishes.
// "backfill" is the first load of history, read oldest first from start; "live" follows new
// sign-ins from through. span is the window size the collector is currently using: it shrinks
// when a window is too busy to read within one run and grows back once windows are quick again.

export const HOUR = 3_600_000;
export const WINDOW_MS = 6 * HOUR;
export const MIN_SPLIT_MS = 5 * 60_000;

// A lane is one forward reading position with its own window state:
// - hold: keep windows at span until the lane passes this time (the end of a window that was too
//   busy).
// - strikes: how many runs in a row a minimum-size window starting at `at` could not be read.
// - trying: the start of the window the lane was reading when the run last saved. Still set when
//   the next run starts means that run was stopped mid-window (its time limit, or a restart), so
//   that window counts as failed instead of being retried at the same size forever.
const windowState = {
  span: z.number().int().min(MIN_SPLIT_MS).max(WINDOW_MS).optional(),
  hold: z.string().datetime().optional(),
  strikes: z
    .object({ at: z.string().datetime(), count: z.number().int().min(1).max(100) })
    .optional(),
  trying: z.string().datetime().optional(),
  // A minimum-size window could not be read without the whole run's budget: the next run reads
  // this lane first.
  owed: z.literal(true).optional(),
};

// Live collection has two lanes. The main one follows new sign-ins from `through`. `settle`
// re-reads behind it: everything before `settle.at` has been read at least lateArrivalMinutes
// after it happened, so records Microsoft publishes late are collected. Older cursors have no
// settle lane; it then starts lateArrivalMinutes before `through`.
export const cursorSchema = z.discriminatedUnion("phase", [
  z.object({
    v: z.literal(1),
    phase: z.literal("backfill"),
    start: z.string().datetime(),
    next: z.string().datetime(),
    ...windowState,
  }),
  z.object({
    v: z.literal(1),
    phase: z.literal("live"),
    through: z.string().datetime(),
    ...windowState,
    settle: z.object({ at: z.string().datetime(), ...windowState }).optional(),
  }),
]);

export type SigninsCursor = z.infer<typeof cursorSchema>;

export function parseCursor(value: string | null | undefined): SigninsCursor | null {
  if (!value) return null;
  try {
    const parsed = cursorSchema.safeParse(JSON.parse(value));
    return parsed.success ? parsed.data : null;
  } catch {
    return null;
  }
}

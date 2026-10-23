import { GraphBoundaryError, type GraphRequestContext } from "@prism365/graph";
import type { ScopedGraph } from "@prism365/modules";
import { z } from "zod";

// Reads for the audits whose Microsoft endpoints share one strict throttle: Conditional Access
// policies, named locations and risky users allow 1 request per second per tenant, and Microsoft
// answers an excess with a 429 that carries no Retry-After. The Graph client then falls back to
// 0.5, 1 and 2 seconds, which is too short to be sure of getting through, so these audits space
// their own requests instead of relying on the retry.
// Microsoft Graph throttling documentation verified 2026-10-22:
// https://learn.microsoft.com/en-us/graph/throttling-limits

// One second between the end of a request and the start of the next, plus a margin for clock
// differences between Prism365 and Microsoft.
export const REQUEST_SPACING_MILLISECONDS = 1_100;

// Waits, and stops waiting the moment the audit's time limit passes. Tests pass a version that
// records the wait and returns at once.
export type Pause = (milliseconds: number, signal: AbortSignal) => Promise<void>;

export const abortablePause: Pause = (milliseconds, signal) =>
  new Promise<void>((resolve, reject) => {
    if (signal.aborted) {
      reject(signal.reason);
      return;
    }
    const stop = () => {
      clearTimeout(timer);
      reject(signal.reason);
    };
    const timer = setTimeout(() => {
      signal.removeEventListener("abort", stop);
      resolve();
    }, milliseconds);
    signal.addEventListener("abort", stop, { once: true });
  });

// Spaces the requests of one run. Call it before every request to a throttled endpoint: the first
// goes straight out, each later one waits. Endpoints with their own limit (users, subscriptions)
// do not need it.
export interface Pacer {
  beforeRequest(signal: AbortSignal): Promise<void>;
}

export function createPacer(
  pause: Pause = abortablePause,
  spacing: number = REQUEST_SPACING_MILLISECONDS,
): Pacer {
  let requested = false;
  return {
    async beforeRequest(signal) {
      signal.throwIfAborted();
      if (requested) {
        await pause(spacing, signal);
        // A pause that ended because the audit was stopped must not let one more request out.
        signal.throwIfAborted();
      }
      requested = true;
    },
  };
}

const pageSchema = z.object({
  value: z.array(z.unknown()).max(1_000),
  "@odata.nextLink": z.string().max(16_384).nullish(),
});

export interface PacedRead {
  items: unknown[];
  // false when Microsoft had more than the page or record limit allowed.
  complete: boolean;
}

// A collection read one page at a time with a wait before every request. The Graph client's own
// collection read pages in one call with no place to wait, so this follows the continuation links
// itself; each request still goes through the scoped client, which checks that the link stays on
// the pack's endpoints and refuses anything but a GET.
export async function readPaced(
  graph: ScopedGraph,
  path: string,
  request: GraphRequestContext,
  options: { pacer: Pacer; signal: AbortSignal; maxPages: number; maxRecords: number },
): Promise<PacedRead> {
  const items: unknown[] = [];
  const seen = new Set<string>();
  let next: string | null | undefined = path;
  for (let pages = 0; next; pages += 1) {
    if (pages >= options.maxPages) return { items, complete: false };
    if (seen.has(next))
      throw new GraphBoundaryError("Graph pagination repeated a continuation link.", "pagination");
    seen.add(next);
    await options.pacer.beforeRequest(options.signal);
    const page: z.infer<typeof pageSchema> = await graph.get(next, request, pageSchema);
    for (const raw of page.value) {
      if (items.length >= options.maxRecords) return { items, complete: false };
      items.push(raw);
    }
    next = page["@odata.nextLink"];
    if (next && items.length >= options.maxRecords) return { items, complete: false };
  }
  return { items, complete: true };
}

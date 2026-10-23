import {
  type GraphRequestContext,
  GraphRequestError,
  GraphRetryDeferredError,
} from "@prism365/graph";
import { z } from "zod";
import type { AuditContext } from "./definition.ts";
import { abortablePause } from "./paced-read.ts";

// Paging by hand, for the audits that must count their requests, pause between pages or keep
// what they read when Microsoft asks them to slow down. (paced-read.ts spaces every request for the
// endpoints with a strict per-second limit; these audits need a budget and a partial result.) The Graph client underneath still checks
// every page against the endpoint allowlist of the read pack it is asked with.

export const pageSchema = z.object({
  value: z.array(z.unknown()).max(1_000),
  "@odata.nextLink": z.string().max(16_384).nullish(),
});

export interface PagedRead {
  items: unknown[];
  // false when more records exist than were read: a limit was reached or Microsoft throttled.
  complete: boolean;
  // true when the read stopped because Microsoft asked Prism365 to slow down.
  throttled: boolean;
}

export interface PageOptions {
  maxRecords: number;
  maxPages?: number;
  // Asked before every request; false stops the read, for a request budget that ran out.
  allow?: () => boolean;
  // A pause between two pages, in milliseconds.
  pageDelayMilliseconds?: number;
}

// Follows a collection page by page. Throttling after the first page keeps what was read and
// says so; on the first page there is nothing to keep, so the error goes to the caller.
export async function readPages(
  context: AuditContext,
  request: GraphRequestContext,
  path: string,
  options: PageOptions,
): Promise<PagedRead> {
  const maxPages = options.maxPages ?? 100;
  const items: unknown[] = [];
  let next: string | undefined = path;
  let pages = 0;
  while (next) {
    context.signal.throwIfAborted();
    if (pages >= maxPages || (options.allow && !options.allow()))
      return { items, complete: false, throttled: false };
    const delay = options.pageDelayMilliseconds ?? 0;
    if (pages > 0 && delay > 0) await abortablePause(delay, context.signal);
    let page: z.output<typeof pageSchema>;
    try {
      page = await context.graph.get(next, request, pageSchema);
    } catch (error) {
      if (error instanceof GraphRetryDeferredError && pages > 0)
        return { items, complete: false, throttled: true };
      throw error;
    }
    pages += 1;
    for (const record of page.value) {
      if (items.length >= options.maxRecords) return { items, complete: false, throttled: false };
      items.push(record);
    }
    next = page["@odata.nextLink"] ?? undefined;
    if (items.length >= options.maxRecords && next)
      return { items, complete: false, throttled: false };
  }
  return { items, complete: true, throttled: false };
}

// Tries a request in several forms, richest first, and moves on only when Microsoft answers 400.
// For query options its documentation does not confirm ($select on a list, a page size): a tenant
// that rejects the richer form still gets the audit. Any other failure goes to the caller.
export async function firstAccepted<T>(
  forms: readonly string[],
  read: (path: string) => Promise<T>,
): Promise<T> {
  let failure: unknown;
  for (const path of forms) {
    try {
      return await read(path);
    } catch (error) {
      if (!(error instanceof GraphRequestError) || error.status !== 400) throw error;
      failure = error;
    }
  }
  throw failure ?? new Error("No form of the request was given.");
}

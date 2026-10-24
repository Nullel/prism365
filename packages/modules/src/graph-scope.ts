import type { GraphFeatureId, GraphHttpClient, GraphRequestContext } from "@prism365/graph";
import type { z } from "zod";
import type { ScopedGraph } from "./contract.ts";

export class ModuleBoundaryError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ModuleBoundaryError";
  }
}

// Narrows the shared Graph client to one module's declared packs. The client underneath still
// enforces the per-pack endpoint allowlist, read-only methods, redirects, size and retry budget;
// this layer adds "only the packs you declared" and "always validate the response".
// `signal`, when given, stops paging and the request in flight once aborted, so a run that has
// ended stops reading instead of finishing in the background.
export function scopeGraph(
  client: GraphHttpClient,
  packs: readonly GraphFeatureId[],
  guard: () => void = () => {},
  signal?: AbortSignal,
): ScopedGraph {
  const allowed = new Set<string>(packs);
  const check = (context: GraphRequestContext) => {
    guard();
    if (!allowed.has(context.feature))
      throw new ModuleBoundaryError(
        `Graph pack ${context.feature} is not declared by this module.`,
      );
  };
  return {
    async get<T>(path: string, context: GraphRequestContext, schema: z.ZodType<T>) {
      check(context);
      return client.request<T>(path, context, signal ? { signal } : {}, schema);
    },
    async getCollection<T>(
      path: string,
      context: GraphRequestContext,
      options: { schema: z.ZodType<T>; maxPages?: number; maxRecords?: number; deadline?: number },
    ) {
      check(context);
      return client.getCollection<T>(path, context, signal ? { ...options, signal } : options);
    },
    async getReport(path: string, context: GraphRequestContext, options = {}) {
      check(context);
      return client.getReport(path, context, signal ? { ...options, signal } : options);
    },
  };
}

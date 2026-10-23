import { auditRowSchema } from "@prism365/contracts";
import { GraphHttpClient } from "@prism365/graph";
import { scopeGraph } from "@prism365/modules";
import { expect } from "vitest";
import { type Audit, type AuditContext, auditDefinition } from "../src/index.ts";

// Shared by the tests of the devices, app consents and Microsoft 365 app usage audits. Synthetic
// data only: invented names, example.invalid domains, made-up IDs.

export const now = new Date("2026-10-22T12:00:00Z");
export const daysAgo = (days: number) => new Date(now.getTime() - days * 86_400_000).toISOString();

export type Route = (url: URL) => Response | undefined;

export interface Fake {
  context: AuditContext;
  // Every request the audit made, in order.
  requests: URL[];
}

export interface FakeOptions {
  granted?: readonly string[];
  scope?: ReadonlySet<string> | null;
  tenantId?: string;
  signal?: AbortSignal;
}

// A Graph client over a fake network: the real client, with its endpoint allowlist, read-only
// boundary and size limits, narrowed to the audit's packs the way the runner narrows it. Retries
// are off and instant, so a 429 shows up at once.
export function fake(audit: Audit, route: Route, options: FakeOptions = {}): Fake {
  const requests: URL[] = [];
  const client = new GraphHttpClient({
    getAccessToken: async () => "synthetic",
    maxResponseBytes: 8_388_608,
    maxRetries: 0,
    sleep: async () => {},
    fetchImplementation: async (input, init) => {
      expect((init?.method ?? "GET").toUpperCase()).toBe("GET");
      const url = new URL(String(input));
      requests.push(url);
      return route(url) ?? new Response("{}", { status: 404 });
    },
  });
  const signal = options.signal ?? new AbortController().signal;
  return {
    requests,
    context: {
      graph: scopeGraph(client, [...audit.packs, ...audit.optionalPacks], () => {}, signal),
      now,
      granted: new Set(options.granted ?? []),
      signal,
      scope: options.scope ?? null,
      ...(options.tenantId ? { tenantId: options.tenantId } : {}),
    },
  };
}

// One page of a collection with a $skiptoken continuation, as Graph pages. `size` is the page
// size when the request names none.
export function paged(url: URL, items: readonly unknown[], size = 100): Response {
  const requested = Number(url.searchParams.get("$top"));
  const top = Number.isFinite(requested) && requested > 0 ? Math.min(requested, size) : size;
  const offset = Number(url.searchParams.get("$skiptoken") ?? "0") || 0;
  const body: { value: unknown[]; "@odata.nextLink"?: string } = {
    value: items.slice(offset, offset + top),
  };
  if (offset + top < items.length) {
    const next = new URL(url);
    next.searchParams.set("$skiptoken", String(offset + top));
    body["@odata.nextLink"] = next.href;
  }
  return Response.json(body);
}

// The checks every audit's published table has to pass: a default view and sort that exist, rows
// that fit the row contract and the audit's own columns, and tile links that open a real view.
export async function expectValidTable(audit: Audit, route: Route, options?: FakeOptions) {
  const definition = auditDefinition(audit);
  const keys = new Set(definition.columns.map((column) => column.key));
  expect(keys.has(definition.defaultSort.column), definition.id).toBe(true);
  expect(definition.views.map((view) => view.id)).toContain(definition.defaultView);
  expect(definition.permissions.length, definition.id).toBeGreaterThan(0);
  const result = await audit.run(fake(audit, route, options).context);
  for (const item of result.rows) {
    auditRowSchema.parse(item);
    for (const key of Object.keys(item.cells))
      expect(keys.has(key), `${definition.id}.${key}`).toBe(true);
  }
  for (const tile of result.summary)
    if (tile.view) expect(definition.views.map((view) => view.id)).toContain(tile.view);
  const viewTags = definition.views.flatMap((view) => view.tags);
  const tags = new Set(result.rows.flatMap((item) => item.tags));
  for (const tag of tags) expect(viewTags, `${definition.id} tag ${tag}`).toContain(tag);
  return result;
}

export function find<T extends { key: string }>(rows: T[], key: string): T {
  const found = rows.find((item) => item.key === key);
  if (!found) throw new Error(`No row ${key}`);
  return found;
}

import type { SyntheticTenant } from "../tenant.ts";

// One request to the synthetic Graph, already split for matching: `path` has no /v1.0/ prefix and
// no trailing slash, and is percent-decoded.
export interface DemoGraphRequest {
  url: URL;
  path: string;
  select: string | null;
  tenant: SyntheticTenant;
}

export interface DemoGraphTools {
  json(body: unknown, status?: number): Response;
  // One page of a collection with a $skiptoken continuation, honouring $top and $select.
  page<T extends object>(url: URL, items: readonly T[], select: string | null): Response;
  notFound(message?: string): Response;
  // A usage report as Microsoft serves it: `body` is kept under a one-time token on the download
  // host, and the answer is the 302 Found whose Location is the link to it.
  download(body: string): Response;
}

// A feature's part of the synthetic Graph. Returns null for paths it does not serve. Everything it
// returns must be invented: example domains only, and addresses from the tenant's own users.
export type DemoGraphExtension = (
  request: DemoGraphRequest,
  tools: DemoGraphTools,
) => Response | null;

import { GraphRequestError, GraphRetryDeferredError } from "@prism365/graph";
import { describe, expect, it } from "vitest";
import { appConsentsAudit, firstAccepted, readPages } from "../src/index.ts";
import { fake, paged, type Route } from "./apps-devices-usage-support.ts";

// The paging helpers behind the app consents and usage audits, read through the real Graph client
// with the application pack.
const application = {
  feature: "directory.applications.read",
  permission: "Application.Read.All",
} as const;
const items = Array.from({ length: 25 }, (_, index) => ({ id: `item-${index}` }));
const list: Route = (url) =>
  url.pathname === "/v1.0/servicePrincipals" ? paged(url, items, 10) : undefined;
const context = (route: Route = list) => fake(appConsentsAudit, route).context;

describe("reading pages by hand", () => {
  it("follows every page and says it read everything", async () => {
    const read = await readPages(context(), application, "servicePrincipals", { maxRecords: 100 });
    expect(read.items).toHaveLength(25);
    expect(read).toMatchObject({ complete: true, throttled: false });
  });

  it("stops at the record limit, whether or not the page ends there", async () => {
    const cut = await readPages(context(), application, "servicePrincipals", { maxRecords: 12 });
    expect(cut.items).toHaveLength(12);
    expect(cut.complete).toBe(false);
    // Exactly the last record is complete: nothing is left to read.
    const exact = await readPages(context(), application, "servicePrincipals", { maxRecords: 25 });
    expect(exact.items).toHaveLength(25);
    expect(exact.complete).toBe(true);
    // A limit on a page boundary with more to come is not complete.
    const ctx = fake(appConsentsAudit, list);
    const edge = await readPages(ctx.context, application, "servicePrincipals", {
      maxRecords: 20,
    });
    expect(edge).toMatchObject({ complete: false, throttled: false });
    expect(edge.items).toHaveLength(20);
    // It stops there, without asking for a page it would only discard.
    expect(ctx.requests).toHaveLength(2);
  });

  it("stops at the page limit", async () => {
    const ctx = fake(appConsentsAudit, list);
    const read = await readPages(ctx.context, application, "servicePrincipals", {
      maxRecords: 100,
      maxPages: 2,
    });
    expect(read.items).toHaveLength(20);
    expect(read.complete).toBe(false);
    expect(ctx.requests).toHaveLength(2);
  });

  it("stops before a request the caller does not allow", async () => {
    const ctx = fake(appConsentsAudit, list);
    let allowed = 2;
    const read = await readPages(ctx.context, application, "servicePrincipals", {
      maxRecords: 100,
      allow: () => allowed-- > 0,
    });
    expect(read.items).toHaveLength(20);
    expect(read.complete).toBe(false);
    expect(ctx.requests).toHaveLength(2);
    const none = await readPages(ctx.context, application, "servicePrincipals", {
      maxRecords: 100,
      allow: () => false,
    });
    expect(none).toEqual({ items: [], complete: false, throttled: false });
  });

  it("keeps the pages it read when throttled afterwards", async () => {
    const route: Route = (url) =>
      url.searchParams.has("$skiptoken") ? new Response("{}", { status: 429 }) : list(url);
    const read = await readPages(context(route), application, "servicePrincipals", {
      maxRecords: 100,
    });
    expect(read.items).toHaveLength(10);
    expect(read).toMatchObject({ complete: false, throttled: true });
  });

  it("raises throttling on the first page, where there is nothing to keep", async () => {
    await expect(
      readPages(
        context(() => new Response("{}", { status: 429 })),
        application,
        "servicePrincipals",
        {
          maxRecords: 100,
        },
      ),
    ).rejects.toBeInstanceOf(GraphRetryDeferredError);
  });

  it("raises other failures, even after some pages were read", async () => {
    const route: Route = (url) =>
      url.searchParams.has("$skiptoken") ? new Response("{}", { status: 500 }) : list(url);
    await expect(
      readPages(context(route), application, "servicePrincipals", { maxRecords: 100 }),
    ).rejects.toMatchObject({ status: 500 });
  });
});

describe("trying a request in several forms", () => {
  const forms = ["rich", "plain"];

  it("moves on only when Microsoft answers 400", async () => {
    const tried: string[] = [];
    const result = await firstAccepted(forms, async (path) => {
      tried.push(path);
      if (path === "rich") throw new GraphRequestError(400);
      return "ok";
    });
    expect(result).toBe("ok");
    expect(tried).toEqual(["rich", "plain"]);
  });

  it("does not move on for any other failure", async () => {
    for (const status of [401, 403, 404, 429, 500]) {
      const tried: string[] = [];
      await expect(
        firstAccepted(forms, async (path) => {
          tried.push(path);
          throw new GraphRequestError(status);
        }),
      ).rejects.toMatchObject({ status });
      expect(tried, String(status)).toEqual(["rich"]);
    }
    await expect(
      firstAccepted(forms, async () => {
        throw new Error("other");
      }),
    ).rejects.toThrow("other");
  });

  it("raises the last 400 when every form is rejected", async () => {
    let call = 0;
    await expect(
      firstAccepted(forms, async () => {
        call += 1;
        throw new GraphRequestError(400, `request-${call}`);
      }),
    ).rejects.toMatchObject({ status: 400, requestId: "request-2" });
  });
});

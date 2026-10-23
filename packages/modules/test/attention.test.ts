import { randomUUID } from "node:crypto";
import { PGliteModuleDatabase } from "@prism365/storage";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  type AttentionInput,
  attentionCounts,
  defaultAttentionPreferences,
  listAttention,
  listAttentionWindow,
  markAttentionRead,
  readAttentionPreferences,
  retireAttention,
  setAttentionMuted,
  setAttentionReview,
  upsertAttention,
  writeAttentionPreferences,
} from "../src/index.ts";

// Synthetic tenant and invented Microsoft incident identifiers only.
const tenant = randomUUID();
const actor = randomUUID();
let database: PGliteModuleDatabase;

beforeEach(async () => {
  database = await PGliteModuleDatabase.open({ scopes: [] });
}, 60_000);
afterEach(async () => {
  await database.close();
});

function incident(
  overrides: Partial<AttentionInput> = {},
  provider: Partial<AttentionInput["provider"]> = {},
): AttentionInput {
  return {
    kind: "service.incident",
    ref: "EX000001",
    class: "incident",
    title: "Some users cannot send mail",
    summary: "Users may be unable to send mail.",
    services: ["Exchange Online"],
    provider: {
      status: "Investigating",
      labels: ["Incident"],
      deadline: null,
      updatedAt: new Date("2026-10-22T10:00:00Z"),
      active: true,
      impact: "Users may be unable to send mail.",
      workaround: null,
      ...provider,
    },
    localLabels: [],
    severity: "high",
    link: "https://admin.microsoft.com/#/servicehealth/:/alerts/EX000001",
    meaningful: { Status: provider.status ?? "Investigating", Classification: "Incident" },
    observedAt: new Date("2026-10-22T10:05:00Z"),
    ...overrides,
  };
}

const upsert = (input: AttentionInput, preferences = defaultAttentionPreferences) =>
  database.tenant(tenant, (sql) =>
    upsertAttention(sql, tenant, "service.health", input, preferences),
  );
const items = (filter: "open" | "unread" | "all" | "muted" = "all") =>
  database.tenant(tenant, (sql) => listAttention(sql, tenant, { filter }));

describe("attention store", () => {
  it("creates one unread item and does not renotify when the same facts repeat", async () => {
    expect((await upsert(incident())).change).toBe("created");
    expect(await items("unread")).toHaveLength(1);
    await database.tenant(tenant, (sql) => markAttentionRead(sql, tenant, null));
    expect((await upsert(incident({ summary: "Reworded by Microsoft." }))).change).toBe(
      "unchanged",
    );
    const [item] = await items();
    expect(item).toMatchObject({ unread: false, revision: 1, summary: "Reworded by Microsoft." });
    expect(await items()).toHaveLength(1);
  });

  it("marks a meaningful change unread with a note, and resolution separately", async () => {
    await upsert(incident());
    await database.tenant(tenant, (sql) => markAttentionRead(sql, tenant, null));
    const changed = await upsert(
      incident(
        { meaningful: { Status: "Restoring service", Classification: "Incident" } },
        {
          status: "Restoring service",
        },
      ),
    );
    expect(changed.change).toBe("changed");
    let [item] = await items();
    expect(item).toMatchObject({ unread: true, revision: 2 });
    expect(item?.changeNote).toContain("Status: Investigating to Restoring service");

    const resolved = await upsert(
      incident(
        { meaningful: { Status: "Service restored", Classification: "Incident" } },
        {
          status: "Service restored",
          active: false,
        },
      ),
    );
    expect(resolved.change).toBe("resolved");
    [item] = await items();
    expect(item?.provider.active).toBe(false);
    expect(item?.changeNote).toContain("Resolved by Microsoft.");
  });

  it("keeps the administrator's review apart from Microsoft's status, and reopens on recurrence", async () => {
    const { id } = await upsert(incident());
    await database.tenant(tenant, (sql) => setAttentionReview(sql, tenant, id, "done", actor));
    let [item] = await items();
    expect(item).toMatchObject({ review: { state: "done" }, unread: false });
    expect(item?.provider.status).toBe("Investigating");
    await upsert(
      incident(
        { meaningful: { Status: "Service restored", Classification: "Incident" } },
        {
          active: false,
          status: "Service restored",
        },
      ),
    );
    [item] = await items();
    expect(item?.review.state).toBe("done");
    expect((await upsert(incident())).change).toBe("reopened");
    [item] = await items();
    expect(item?.review.state).toBe("new");
  });

  it("never notifies for a muted item, and notifies only about important services when set", async () => {
    const { id } = await upsert(incident());
    await database.tenant(tenant, (sql) => setAttentionMuted(sql, tenant, id, true));
    await upsert(incident({ meaningful: { Status: "Mitigated" } }, { status: "Mitigated" }));
    expect(await items("unread")).toEqual([]);
    expect(await items("muted")).toHaveLength(1);

    const teamsOnly = { ...defaultAttentionPreferences, services: ["Microsoft Teams"] };
    await upsert(incident({ ref: "EX000002" }), teamsOnly);
    await upsert(incident({ ref: "TM000003", services: ["Microsoft Teams"] }), teamsOnly);
    const unread = await items("unread");
    expect(unread.map((item) => item.ref)).toEqual(["TM000003"]);
    expect(unread[0]?.localLabels).toContain("Important service");
    expect(await database.tenant(tenant, (sql) => attentionCounts(sql, tenant))).toEqual({
      unread: 1,
      active: 2,
    });
  });

  it("respects the per-class notification switches", async () => {
    const quiet = {
      ...defaultAttentionPreferences,
      notify: { ...defaultAttentionPreferences.notify, advisory: false },
    };
    await upsert(incident({ ref: "EX000004", class: "advisory" }), quiet);
    expect(await items("unread")).toEqual([]);
    expect(await items("open")).toHaveLength(1);
  });

  it("retires items of one kind that the provider stopped listing", async () => {
    await upsert(incident({ kind: "service.message", ref: "MC000001", class: "message.action" }));
    await upsert(incident({ kind: "service.message", ref: "MC000002", class: "message.action" }));
    const retired = await database.tenant(tenant, (sql) =>
      retireAttention(
        sql,
        tenant,
        "service.health",
        "service.message",
        ["MC000002"],
        new Date(),
        "No longer listed by Microsoft.",
      ),
    );
    expect(retired).toBe(1);
    expect((await items("open")).map((item) => item.ref)).toEqual(["MC000002"]);
  });

  it("refuses malformed items and non-HTTPS links", async () => {
    await expect(upsert(incident({ link: "http://example.invalid/x" }))).rejects.toMatchObject({
      name: "AttentionInputError",
    });
    await expect(upsert(incident({ kind: "Bad Kind" }))).rejects.toMatchObject({
      name: "AttentionInputError",
    });
  });

  it("stores preferences per tenant with defaults", async () => {
    expect(await database.tenant(tenant, (sql) => readAttentionPreferences(sql, tenant))).toEqual(
      defaultAttentionPreferences,
    );
    const saved = await database.tenant(tenant, (sql) =>
      writeAttentionPreferences(
        sql,
        tenant,
        { ...defaultAttentionPreferences, services: ["Exchange Online"] },
        actor,
      ),
    );
    expect(saved.services).toEqual(["Exchange Online"]);
    const other = randomUUID();
    expect(await database.tenant(other, (sql) => readAttentionPreferences(sql, other))).toEqual(
      defaultAttentionPreferences,
    );
  });

  it("never lets read or resolved items push an active incident out of the panel window", async () => {
    for (let index = 0; index < 305; index += 1)
      await upsert(
        incident(
          { ref: `EX9${index}`, observedAt: new Date("2026-10-22T10:00:00Z") },
          { active: false },
        ),
      );
    const { id } = await upsert(incident({ ref: "EX-ACTIVE" }));
    await database.tenant(tenant, (sql) => setAttentionReview(sql, tenant, id, "reviewed", actor));
    const window = await database.tenant(tenant, (sql) =>
      listAttentionWindow(sql, tenant, {
        module: "service.health",
        kind: "service.incident",
        since: new Date("2026-10-21T00:00:00Z"),
        limit: 300,
      }),
    );
    expect(window[0]?.ref).toBe("EX-ACTIVE");
  });

  it("notices a change after the first 80 characters, and updates the class on escalation", async () => {
    const long = `${"x".repeat(100)} first`;
    await upsert(
      incident({ class: "advisory", meaningful: { Impact: long, Classification: "Advisory" } }),
    );
    const result = await upsert(
      incident({
        class: "incident",
        meaningful: { Impact: `${"x".repeat(100)} second`, Classification: "Incident" },
      }),
    );
    expect(result.change).toBe("changed");
    const [item] = await items();
    expect(item).toMatchObject({ revision: 2, class: "incident" });
    expect(item?.changeNote).toContain("Impact changed");
  });

  it("cleans odd provider text instead of rejecting the item", async () => {
    const result = await upsert(
      incident(
        {
          title: "  ",
          services: ["", "Exchange Online", "S".repeat(150)],
          meaningful: { Impact: "i".repeat(3_000), Title: "bad\u0000title" },
          summary: "has\u0000nul",
        },
        { labels: ["", "Incident"], impact: "i".repeat(5_000) },
      ),
    );
    expect(result.change).toBe("created");
    const [item] = await items();
    expect(item?.title).toBe("EX000001");
    expect(item?.services).toEqual(["Exchange Online", "S".repeat(120)]);
    expect(item?.summary).toBe("hasnul");
  });

  it("keeps the administrator's review when a post drops out of the feed and returns", async () => {
    const post = incident({ kind: "service.message", ref: "MC000009", class: "message.action" });
    const { id } = await upsert(post);
    await database.tenant(tenant, (sql) => setAttentionReview(sql, tenant, id, "done", actor));
    await database.tenant(tenant, (sql) =>
      retireAttention(
        sql,
        tenant,
        "service.health",
        "service.message",
        [],
        new Date(),
        "No longer listed by Microsoft.",
      ),
    );
    expect((await upsert(post)).change).toBe("reopened");
    const [item] = await items();
    expect(item?.review.state).toBe("done");
  });

  it("hides another tenant's preferences and items through row-level security", async () => {
    await upsert(incident());
    await database.tenant(tenant, (sql) =>
      writeAttentionPreferences(
        sql,
        tenant,
        { ...defaultAttentionPreferences, services: ["Exchange Online"] },
        actor,
      ),
    );
    const other = randomUUID();
    const seen = await database.tenant(other, async (sql) => ({
      preferences: (await sql.query("SELECT key FROM prism365_preference")).rows,
      items: (await sql.query("SELECT id FROM prism365_attention")).rows,
    }));
    expect(seen).toEqual({ preferences: [], items: [] });
  });
});

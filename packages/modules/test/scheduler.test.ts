import { randomUUID } from "node:crypto";
import type { PGliteModuleDatabase } from "@prism365/storage";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { z } from "zod";
import {
  type CollectionCredential,
  CredentialUnavailableError,
  type FindingNotification,
  latestRuns,
  listFindings,
  type ModuleRegistry,
  ModuleScheduler,
  type SchedulerEvent,
  setFindingState,
} from "../src/index.ts";
import {
  actor,
  fakeGraph,
  finding,
  openRuntime,
  type ProbeBehavior,
  probeModule,
  tenantA,
  tenantB,
} from "./fixtures.ts";

const userSchema = z.object({ id: z.string(), displayName: z.string() });
const usersContext = { feature: "directory.users.read", permission: "User.Read.All" } as const;

function credential(overrides: Partial<CollectionCredential> = {}): CollectionCredential {
  return {
    kind: "delegated",
    tenantId: tenantA,
    actorId: actor,
    grantedPermissions: async () => ["User.Read", "User.Read.All"],
    accessToken: async () => "synthetic-token",
    ...overrides,
  };
}

let behavior: { run: ProbeBehavior };
let registry: ModuleRegistry;
let database: PGliteModuleDatabase;
let clock: Date;
let events: SchedulerEvent[];
let delivered: FindingNotification[];
let scheduler: ModuleScheduler;
let graph: ReturnType<typeof fakeGraph>;

function build(
  options: { timeoutMilliseconds?: number; sinkFails?: boolean; catchUpDelay?: number } = {},
) {
  scheduler = new ModuleScheduler({
    enabledPacks: registry.list().flatMap((module) => [...module.packs]),
    registry,
    database,
    graphClient: graph.factory,
    now: () => clock,
    tickMilliseconds: 60_000,
    ...(options.catchUpDelay ? { catchUpDelayMilliseconds: options.catchUpDelay } : {}),
    ...(options.timeoutMilliseconds
      ? { timeoutMilliseconds: () => options.timeoutMilliseconds ?? 0 }
      : {}),
    sinks: [
      {
        id: "test-sink",
        minimumSeverity: "medium",
        deliver: async (notification) => {
          if (options.sinkFails) throw new Error("synthetic sink failure");
          delivered.push(notification);
        },
      },
    ],
    onEvent: (event) => events.push(event),
  });
}

const runs = () => database.tenant(tenantA, (sql) => latestRuns(sql, tenantA));
const advance = (seconds: number) => {
  clock = new Date(clock.getTime() + seconds * 1_000);
};

beforeEach(async () => {
  behavior = { run: async () => undefined };
  ({ registry, database } = await openRuntime(probeModule(behavior)));
  clock = new Date("2026-10-22T12:00:00Z");
  events = [];
  delivered = [];
  graph = fakeGraph();
  build();
}, 60_000);

afterEach(async () => {
  await scheduler?.stop();
  await database?.close();
});

describe("module scheduler", () => {
  it("does nothing until a credential is bound", async () => {
    let calls = 0;
    behavior.run = async () => {
      calls += 1;
      return undefined;
    };
    await scheduler.tick();
    expect(calls).toBe(0);
    expect(scheduler.status().state).toBe("stopped");
  });

  it("never runs a module whose pack is not enabled, even when its permission is granted", async () => {
    let calls = 0;
    behavior.run = async () => {
      calls += 1;
      return undefined;
    };
    // The probe module's permission (User.Read.All) is granted, but its pack is not enabled: a
    // permission shared with an enabled pack must not switch the module on.
    scheduler = new ModuleScheduler({
      enabledPacks: ["tenant.licenses.read"],
      registry,
      database,
      graphClient: graph.factory,
      now: () => clock,
      onEvent: (event) => events.push(event),
    });
    scheduler.bind(credential());
    await scheduler.tick();
    expect(calls).toBe(0);
    expect(events).toContainEqual(
      expect.objectContaining({ type: "module.skipped", missing: ["directory.users.read"] }),
    );
  });

  it("runs each collector on its cadence and records the run", async () => {
    let calls = 0;
    behavior.run = async (context) => {
      calls += 1;
      const page = await context.graph.getCollection("users?$select=id,displayName", usersContext, {
        schema: userSchema,
      });
      context.count.fetched(page.items.length);
      return { complete: page.complete };
    };
    scheduler.bind(credential());
    await scheduler.tick();
    await scheduler.tick();
    expect(calls).toBe(1);
    advance(299);
    await scheduler.tick();
    expect(calls).toBe(1);
    advance(1);
    await scheduler.tick();
    expect(calls).toBe(2);
    expect(await runs()).toEqual([
      expect.objectContaining({
        module: "test.probe",
        collector: "poll",
        status: "succeeded",
        fetched: 1,
        credentialKind: "delegated",
      }),
    ]);
    expect(graph.calls[0]).toContain("https://graph.microsoft.com/v1.0/users");
  });

  it("never overlaps cycles when ticks arrive during a run", async () => {
    let calls = 0;
    let release: () => void = () => {};
    behavior.run = async () => {
      calls += 1;
      await new Promise<void>((resolve) => {
        release = resolve;
      });
      return undefined;
    };
    scheduler.bind(credential());
    const first = scheduler.tick();
    const second = scheduler.tick();
    await new Promise((resolve) => setTimeout(resolve, 20));
    release();
    await Promise.all([first, second]);
    expect(calls).toBe(1);
  });

  it("skips a module whose permissions the credential does not hold", async () => {
    let calls = 0;
    behavior.run = async () => {
      calls += 1;
      return undefined;
    };
    scheduler.bind(credential({ grantedPermissions: async () => ["User.Read"] }));
    await scheduler.tick();
    await scheduler.tick();
    expect(calls).toBe(0);
    expect(scheduler.status().skipped).toEqual([
      { module: "test.probe", missing: ["User.Read.All"] },
    ]);
    expect(events.filter((event) => event.type === "module.skipped")).toHaveLength(1);
  });

  it("fails a run that calls a Graph pack the module did not declare", async () => {
    behavior.run = async (context) => {
      await context.graph.get(
        "subscribedSkus",
        { feature: "tenant.licenses.read", permission: "LicenseAssignment.Read.All" },
        z.unknown(),
      );
      return undefined;
    };
    scheduler.bind(credential());
    await scheduler.tick();
    expect((await runs())[0]).toMatchObject({ status: "failed", errorCategory: "boundary.module" });
    expect(graph.calls).toEqual([]);
  });

  it("refuses Graph calls and nested storage calls inside a module transaction", async () => {
    const seen: string[] = [];
    behavior.run = async (context) => {
      await context.db(async () => {
        await context.graph
          .getCollection("users", usersContext, { schema: userSchema })
          .catch((error: Error) => seen.push(error.name));
        await context.cursor.get().catch((error: Error) => seen.push(error.name));
      });
      return undefined;
    };
    scheduler.bind(credential());
    await scheduler.tick();
    expect(seen).toEqual(["ModuleBoundaryError", "ModuleBoundaryError"]);
    expect(graph.calls).toEqual([]);
  });

  it("lets a collector read a setting it declared, for its own tenant, and nothing else", async () => {
    await scheduler.stop();
    await database.close();
    ({ registry, database } = await openRuntime({
      ...probeModule(behavior),
      settings: ["test.canaries", "test.unset"],
    }));
    build();
    const save = (tenant: string, key: string, value: unknown) =>
      database.tenant(tenant, (sql) =>
        sql.query(
          `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
           VALUES ($1,$2,$3::jsonb,now(),$4)`,
          [tenant, key, JSON.stringify(value), actor],
        ),
      );
    await save(tenantA, "test.canaries", { accounts: ["mine"] });
    await save(tenantB, "test.canaries", { accounts: ["theirs"] });
    await save(tenantA, "attention", { secret: true });
    const seen: unknown[] = [];
    behavior.run = async (context) => {
      seen.push(await context.settings.get("test.canaries"));
      seen.push(await context.settings.get("test.unset"));
      // A key the module did not declare is refused, even though the row exists.
      await context.settings.get("attention").catch((error: Error) => seen.push(error.name));
      return undefined;
    };
    scheduler.bind(credential());
    await scheduler.tick();
    expect(seen).toEqual([{ accounts: ["mine"] }, null, "ModuleBoundaryError"]);
    expect((await runs())[0]).toMatchObject({ status: "succeeded" });
  });

  it("times out a slow run and discards anything it tries to write afterwards", async () => {
    build({ timeoutMilliseconds: 30 });
    let lateWrite: unknown = null;
    behavior.run = (context) =>
      new Promise((resolve) => {
        setTimeout(() => {
          context
            .db((sql) =>
              sql.query(
                `INSERT INTO ${context.table("item")} (tenant_id,id,value,created_at) VALUES ($1,'late','x',now())`,
                [context.tenantId],
              ),
            )
            .then(
              () => {
                lateWrite = "written";
              },
              (error: Error) => {
                lateWrite = error.name;
              },
            )
            .finally(() => resolve(undefined));
        }, 80);
      });
    scheduler.bind(credential());
    await scheduler.tick();
    expect((await runs())[0]).toMatchObject({ status: "failed", errorCategory: "timeout" });
    await new Promise((resolve) => setTimeout(resolve, 120));
    expect(lateWrite).toBe("RunEndedError");
    const items = await database.tenant(tenantA, (sql) =>
      sql.query("SELECT id FROM mod_test_probe_item"),
    );
    expect(items.rows).toEqual([]);
  });

  it("stops collecting when the credential is lost, and blocks further token use", async () => {
    behavior.run = async (context) => {
      await context.graph.getCollection("users", usersContext, { schema: userSchema });
      return undefined;
    };
    scheduler.bind(
      credential({
        accessToken: async () => {
          throw new CredentialUnavailableError();
        },
      }),
    );
    await scheduler.tick();
    expect((await runs())[0]).toMatchObject({
      status: "failed",
      errorCategory: "credential_unavailable",
    });
    expect(scheduler.status()).toMatchObject({
      state: "stopped",
      lastStopReason: "credential_unavailable",
    });
    expect(events).toContainEqual(
      expect.objectContaining({ type: "collection.unbound", reason: "credential_unavailable" }),
    );
  });

  it("stops when the operator stops collection mid-run", async () => {
    let tokenRequests = 0;
    let after: string | null = null;
    behavior.run = async (context) => {
      scheduler.unbind("operator_stopped");
      await context.graph
        .getCollection("users", usersContext, { schema: userSchema })
        .catch((error: Error) => {
          after = error.name;
        });
      return undefined;
    };
    scheduler.bind(
      credential({
        accessToken: async () => {
          tokenRequests += 1;
          return "synthetic-token";
        },
      }),
    );
    await scheduler.tick();
    expect(after).toBe("CredentialUnavailableError");
    expect(tokenRequests).toBe(0);
  });

  it("notifies sinks on new and reopened findings, not on repeats", async () => {
    behavior.run = async (context) => {
      await context.findings.upsert(finding({ observedAt: context.now }));
      return undefined;
    };
    scheduler.bind(credential());
    await scheduler.tick();
    advance(300);
    await scheduler.tick();
    expect(delivered.map((item) => item.change)).toEqual(["created"]);
    const [record] = await database.tenant(tenantA, (sql) => listFindings(sql, tenantA));
    if (!record) throw new Error("finding missing");
    await database.tenant(tenantA, (sql) =>
      setFindingState(sql, tenantA, record.id, "resolved", actor),
    );
    advance(300);
    await scheduler.tick();
    expect(delivered.map((item) => item.change)).toEqual(["created", "reopened"]);
    expect((await runs())[0]).toMatchObject({ status: "succeeded", findings: 1 });
  });

  it("keeps collecting when a sink fails", async () => {
    build({ sinkFails: true });
    behavior.run = async (context) => {
      await context.findings.upsert(finding({ observedAt: context.now }));
      return undefined;
    };
    scheduler.bind(credential());
    await scheduler.tick();
    expect(events).toContainEqual({ type: "sink.failed", sink: "test-sink" });
    expect((await runs())[0]).toMatchObject({ status: "succeeded" });
    expect(scheduler.status().state).toBe("collecting");
  });

  it("persists the collector cursor between runs", async () => {
    const seen: (string | null)[] = [];
    behavior.run = async (context) => {
      seen.push(await context.cursor.get());
      await context.cursor.set(context.now.toISOString());
      return undefined;
    };
    scheduler.bind(credential());
    await scheduler.tick();
    advance(300);
    await scheduler.tick();
    expect(seen).toEqual([null, "2026-10-22T12:00:00.000Z"]);
  });

  it("applies declared retention to this tenant only and marks abandoned runs", async () => {
    const old = new Date(clock.getTime() - 40 * 86_400_000).toISOString();
    const recent = new Date(clock.getTime() - 86_400_000).toISOString();
    for (const tenant of [tenantA, tenantB])
      await database.tenant(tenant, (sql) =>
        sql.query(
          "INSERT INTO mod_test_probe_item (tenant_id,id,value,created_at) VALUES ($1,'old','x',$2),($1,'recent','x',$3)",
          [tenant, old, recent],
        ),
      );
    await database.tenant(tenantA, (sql) =>
      sql.query(
        "INSERT INTO prism365_module_run (id,tenant_id,module,collector,credential_kind,started_at,status) VALUES ($1,$2,'test.probe','other','delegated',$3,'running')",
        [randomUUID(), tenantA, recent],
      ),
    );
    scheduler.bind(credential());
    await scheduler.tick();
    const ids = (tenant: string) =>
      database.tenant(tenant, (sql) =>
        sql.query<{ id: string }>("SELECT id FROM mod_test_probe_item ORDER BY id"),
      );
    expect((await ids(tenantA)).rows.map((row) => row.id)).toEqual(["recent"]);
    expect((await ids(tenantB)).rows.map((row) => row.id)).toEqual(["old", "recent"]);
    expect((await runs()).find((run) => run.collector === "other")?.status).toBe("abandoned");
  });
  it("runs a collector again before its cadence while it reports more to read", async () => {
    let calls = 0;
    behavior.run = async () => {
      calls += 1;
      return { complete: true, more: calls < 3 };
    };
    scheduler.bind(credential());
    for (let tick = 0; tick < 5; tick += 1) await scheduler.tick();
    // Runs 1 and 2 had more to read; run 3 did not, so the cadence applies again.
    expect(calls).toBe(3);
    expect((await runs())[0]?.status).toBe("succeeded");
  });

  it("catches up on its own once started, and stops when there is no more", async () => {
    build({ catchUpDelay: 10 });
    let calls = 0;
    behavior.run = async () => {
      calls += 1;
      return { complete: true, more: calls < 3 };
    };
    scheduler.bind(credential());
    scheduler.start();
    await scheduler.tick();
    for (let wait = 0; wait < 200 && calls < 3; wait += 1)
      await new Promise((resolve) => setTimeout(resolve, 10));
    await new Promise((resolve) => setTimeout(resolve, 100));
    expect(calls).toBe(3);
  });

  it("forgets pending catch-up when collection stops", async () => {
    let calls = 0;
    behavior.run = async () => {
      calls += 1;
      return { complete: true, more: true };
    };
    scheduler.bind(credential());
    await scheduler.tick();
    scheduler.unbind("operator_stopped");
    scheduler.bind(credential());
    await scheduler.tick();
    expect(calls).toBe(1);
  });
});

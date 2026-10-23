import { mkdir } from "node:fs/promises";
import { resolve } from "node:path";
import { GraphHttpClient } from "@prism365/graph";
import {
  createMessageCenterModule,
  createServiceHealthModule,
} from "@prism365/module-service-health";
import { createSigninsAuditModule } from "@prism365/module-signins-audit";
import {
  type CollectionCredential,
  CredentialUnavailableError,
  ModuleRegistry,
  ModuleScheduler,
  type SchedulerEvent,
} from "@prism365/modules";
import {
  type ModuleDatabase,
  PGliteModuleDatabase,
  PostgresModuleDatabase,
} from "@prism365/storage";
import type { AppConfig } from "./config.ts";
import { digest, type IdentityService, type IdentitySession } from "./identity.ts";

export interface ModuleRuntime {
  registry: ModuleRegistry;
  database: ModuleDatabase;
  scheduler: ModuleScheduler;
}

// Built-in modules. Migrations do not depend on settings, so the migrator can call this without
// a live configuration.
// Every built-in module is registered. One whose read pack this installation has not enabled, or
// whose permission the signed-in administrator has not granted, is skipped by the scheduler and
// reported as unavailable, so a fresh install works with whatever it has been given.
export function createModuleRegistry(config?: AppConfig): ModuleRegistry {
  return new ModuleRegistry()
    .register(createSigninsAuditModule(config?.signinsAudit))
    .register(createServiceHealthModule())
    .register(createMessageCenterModule());
}

// Collects with the signed-in operator's delegated session. The session handle stays in process
// memory, so collection ends when the operator signs out, the session expires, the operator is
// removed from the allowlist, or the process restarts. Only a digest is exposed for comparison.
export class DelegatedSessionCredential implements CollectionCredential {
  readonly kind = "delegated" as const;
  readonly tenantId: string;
  readonly actorId: string;
  readonly sessionKey: string;
  readonly #identity: IdentityService;
  readonly #handle: string;
  readonly #permitted: () => readonly string[];

  constructor(
    identity: IdentityService,
    handle: string,
    session: IdentitySession,
    permitted: () => readonly string[],
  ) {
    this.#identity = identity;
    this.#handle = handle;
    this.#permitted = permitted;
    this.tenantId = session.tenantId;
    this.actorId = session.actorId;
    this.sessionKey = digest(handle);
  }

  // Consented scopes, narrowed to the read packs this installation enabled.
  async grantedPermissions(): Promise<readonly string[]> {
    const session = await this.#identity.session(this.#handle);
    if (!session) throw new CredentialUnavailableError();
    const allowed = new Set(this.#permitted());
    return session.scopes.filter((scope) => allowed.has(scope));
  }

  // Set when Microsoft refused a token for a session that still exists (a sign-in frequency
  // policy, a revoked refresh token): signing in again is the only cure, so the session is not
  // used to resume collection. A missing session or a storage error does not set it.
  refused = false;

  async accessToken(permissions: readonly string[]): Promise<string> {
    try {
      return await this.#identity.accessToken(this.#handle, [...permissions]);
    } catch (error) {
      const name = error instanceof Error ? error.name : "";
      if (name === "InteractionRequiredAuthError") {
        this.refused = true;
        throw new CredentialUnavailableError();
      }
      if (!(await this.#identity.session(this.#handle))) throw new CredentialUnavailableError();
      // Anything else is treated as transient: this run fails, collection continues.
      throw new Error("Token acquisition failed.");
    }
  }
}

export function logSchedulerEvent(event: SchedulerEvent): void {
  // Categories and counts only: no tenant identifiers, provider messages, or account data.
  if (event.type === "run.finished")
    console.log(
      `Prism365 module ${event.module}/${event.collector}: ${event.status}${event.category ? ` (${event.category})` : ""}, fetched ${event.fetched}, findings ${event.findings}${event.skipped ? `, ${event.skipped} unreadable records left out` : ""}`,
    );
  else if (event.type === "collection.bound") console.log("Prism365 collection started.");
  else if (event.type === "collection.unbound")
    console.log(`Prism365 collection stopped (${event.reason}).`);
  else if (event.type === "module.skipped")
    console.log(`Prism365 module ${event.module} skipped: missing ${event.missing.join(", ")}.`);
  else if (event.type === "sink.failed") console.warn(`Prism365 sink ${event.sink} failed.`);
  else console.warn(`Prism365 ${event.task} maintenance failed.`);
}

// Live mode only. PostgreSQL when the deployment configured it, otherwise PGlite: in memory by
// default, on disk when PRISM_DATA_DIR is set.
export async function openModuleRuntime(config: AppConfig): Promise<ModuleRuntime> {
  const registry = createModuleRegistry(config);
  let database: ModuleDatabase;
  if (config.databaseUrl) {
    const postgres = new PostgresModuleDatabase({ connectionString: config.databaseUrl });
    try {
      await postgres.verify();
      const tenantId = config.identity?.tenantId;
      // Core tables the API uses directly. The allowed-table set is fixed when the database is
      // verified, so starting before `npm run db:migrate` finished would refuse every query on
      // them until a restart; refusing to start makes the container retry instead.
      if (tenantId)
        for (const table of ["prism365_attention", "prism365_preference", "prism365_audit_run"]) {
          const exists = await postgres.tenant(tenantId, (sql) =>
            sql.query<{ present: boolean }>("SELECT to_regclass($1) IS NOT NULL AS present", [
              `public.${table}`,
            ]),
          );
          if (!exists.rows[0]?.present)
            throw new Error(`The core schema is not migrated (${table}). Run npm run db:migrate.`);
        }
      if (tenantId)
        for (const module of registry.list())
          for (const table of registry.tables(module.id)) {
            const exists = await postgres.tenant(tenantId, (sql) =>
              sql.query<{ present: boolean }>("SELECT to_regclass($1) IS NOT NULL AS present", [
                `public.${table}`,
              ]),
            );
            if (!exists.rows[0]?.present)
              throw new Error(`Module ${module.id} is not migrated. Run npm run db:migrate.`);
          }
    } catch (error) {
      await postgres.close();
      throw error;
    }
    database = postgres;
  } else {
    let dataDir: string | undefined;
    if (config.moduleDataDir) {
      dataDir = resolve(config.moduleDataDir);
      await mkdir(dataDir, { recursive: true, mode: 0o700 });
    }
    database = await PGliteModuleDatabase.open({
      ...(dataDir ? { dataDir } : {}),
      scopes: registry.migrationScopes(),
    });
  }
  const scheduler = new ModuleScheduler({
    enabledPacks: config.enabledFeatures,
    registry,
    database,
    // Background reads tolerate slower Graph queries and longer throttling waits than the
    // interactive investigation does.
    graphClient: (getAccessToken) =>
      new GraphHttpClient({
        getAccessToken,
        // Sign-in log queries routinely take longer than 30 seconds.
        requestTimeoutMilliseconds: 90_000,
        // Sign-in records carry every Conditional Access policy evaluated; a page of them can
        // pass the interactive 1 MB budget.
        maxResponseBytes: 8_388_608,
        maxRetries: 3,
        maxRetryWaitMilliseconds: 30_000,
      }),
    onEvent: logSchedulerEvent,
  });
  return { registry, database, scheduler };
}

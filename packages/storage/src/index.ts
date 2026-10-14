import { createCipheriv, createDecipheriv, randomBytes, randomUUID } from "node:crypto";
import { Pool, type PoolClient } from "pg";

export type StateKind = "oauth" | "session" | "aiGrant" | "investigation" | "automationPackage";
export interface StateStore {
  readonly durable: boolean;
  get(tenantId: string, kind: StateKind, key: string): Promise<unknown | null>;
  take(tenantId: string, kind: StateKind, key: string): Promise<unknown | null>;
  replace(
    tenantId: string,
    kind: StateKind,
    key: string,
    value: unknown,
    expiresAt: Date,
  ): Promise<boolean>;
  list(tenantId: string, kind: StateKind): Promise<unknown[]>;
  purgeExpired(tenantId: string): Promise<void>;
  set(
    tenantId: string,
    kind: StateKind,
    key: string,
    value: unknown,
    expiresAt: Date,
  ): Promise<void>;
  delete(tenantId: string, kind: StateKind, key: string): Promise<void>;
  appendAudit(tenantId: string, actorId: string, event: Record<string, unknown>): Promise<void>;
  close(): Promise<void>;
}

function validateTenant(tenantId: string) {
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(tenantId)) {
    throw new Error("State storage requires a valid tenant UUID.");
  }
}

export class MemoryStateStore implements StateStore {
  readonly durable: boolean = false;
  readonly #states = new Map<string, { value: unknown; expiresAt: number }>();
  readonly audit: { tenantId: string; actorId: string; event: Record<string, unknown> }[] = [];

  #key(tenantId: string, kind: StateKind, key: string) {
    validateTenant(tenantId);
    return `${tenantId.toLowerCase()}:${kind}:${key}`;
  }

  async get(tenantId: string, kind: StateKind, key: string): Promise<unknown | null> {
    const stateKey = this.#key(tenantId, kind, key);
    const item = this.#states.get(stateKey);
    if (!item || item.expiresAt <= Date.now()) {
      this.#states.delete(stateKey);
      return null;
    }
    return structuredClone(item.value);
  }

  async set(tenantId: string, kind: StateKind, key: string, value: unknown, expiresAt: Date) {
    for (const [stateKey, item] of this.#states) {
      if (item.expiresAt <= Date.now()) this.#states.delete(stateKey);
    }
    if (this.#states.size >= 1_000 && !this.#states.has(this.#key(tenantId, kind, key)))
      throw new Error("Development state capacity was reached.");
    this.#states.set(this.#key(tenantId, kind, key), {
      value: structuredClone(value),
      expiresAt: expiresAt.getTime(),
    });
  }

  async take(tenantId: string, kind: StateKind, key: string): Promise<unknown | null> {
    const stateKey = this.#key(tenantId, kind, key);
    const item = this.#states.get(stateKey);
    this.#states.delete(stateKey);
    return item && item.expiresAt > Date.now() ? structuredClone(item.value) : null;
  }

  async replace(tenantId: string, kind: StateKind, key: string, value: unknown, expiresAt: Date) {
    const stateKey = this.#key(tenantId, kind, key);
    const item = this.#states.get(stateKey);
    if (!item || item.expiresAt <= Date.now()) return false;
    this.#states.set(stateKey, { value: structuredClone(value), expiresAt: expiresAt.getTime() });
    return true;
  }

  async list(tenantId: string, kind: StateKind) {
    const prefix = this.#key(tenantId, kind, "");
    return [...this.#states.entries()]
      .filter(([key, item]) => key.startsWith(prefix) && item.expiresAt > Date.now())
      .slice(0, 50)
      .map(([, item]) => structuredClone(item.value));
  }

  async purgeExpired(tenantId: string) {
    validateTenant(tenantId);
    for (const [key, item] of this.#states)
      if (key.startsWith(`${tenantId.toLowerCase()}:`) && item.expiresAt <= Date.now())
        this.#states.delete(key);
  }

  async delete(tenantId: string, kind: StateKind, key: string) {
    this.#states.delete(this.#key(tenantId, kind, key));
  }

  async appendAudit(tenantId: string, actorId: string, event: Record<string, unknown>) {
    validateTenant(tenantId);
    if (this.audit.length >= 10_000) throw new Error("Development audit capacity was reached.");
    this.audit.push({ tenantId, actorId, event: structuredClone(event) });
  }

  async close() {
    this.#states.clear();
    this.audit.length = 0;
  }
}

export class StateCipher {
  readonly #key: Buffer;
  constructor(hexKey: string) {
    if (!/^[a-f0-9]{64}$/i.test(hexKey)) {
      throw new Error("STATE_ENCRYPTION_KEY must be a 32-byte hex key.");
    }
    this.#key = Buffer.from(hexKey, "hex");
  }

  encrypt(value: unknown, binding: string): Buffer {
    const nonce = randomBytes(12);
    const cipher = createCipheriv("aes-256-gcm", this.#key, nonce);
    cipher.setAAD(Buffer.from(binding));
    const content = Buffer.from(JSON.stringify(value));
    if (content.length > 5_000_000) throw new Error("Stored state exceeds the size limit.");
    const encrypted = Buffer.concat([cipher.update(content), cipher.final()]);
    return Buffer.concat([nonce, cipher.getAuthTag(), encrypted]);
  }

  decrypt(content: Buffer, binding: string): unknown {
    if (content.length < 29) throw new Error("Stored state has an invalid encrypted envelope.");
    const decipher = createDecipheriv("aes-256-gcm", this.#key, content.subarray(0, 12));
    decipher.setAAD(Buffer.from(binding));
    decipher.setAuthTag(content.subarray(12, 28));
    return JSON.parse(
      Buffer.concat([decipher.update(content.subarray(28)), decipher.final()]).toString("utf8"),
    ) as unknown;
  }
}

export class PostgresStateStore implements StateStore {
  readonly durable: boolean = true;
  readonly #pool: Pool;
  readonly #cipher: StateCipher;
  constructor(options: { connectionString: string; encryptionKey: string }) {
    this.#pool = new Pool({
      connectionString: options.connectionString,
      max: 5,
      connectionTimeoutMillis: 5_000,
      statement_timeout: 10_000,
    });
    // pg removes failed idle clients. Handle that event without emitting credential-bearing
    // driver objects; active operation failures still reject and are sanitized by the API.
    this.#pool.on("error", () => {});
    this.#cipher = new StateCipher(options.encryptionKey);
  }

  async verifyRole() {
    const result = await this.#pool.query<{ rolsuper: boolean; rolbypassrls: boolean }>(
      "SELECT rolsuper, rolbypassrls FROM pg_roles WHERE rolname = current_user",
    );
    if (!result.rows[0] || result.rows[0].rolsuper || result.rows[0].rolbypassrls) {
      throw new Error(
        "The application database role must not be superuser or bypass row security.",
      );
    }
    const ownership = await this.#pool.query(
      "SELECT 1 FROM pg_tables WHERE schemaname='public' AND tablename IN ('prism365_state','prism365_audit') AND pg_has_role(current_user,tableowner,'MEMBER')",
    );
    if (ownership.rowCount)
      throw new Error("The application database role must not own Prism365 tables.");
    const tables = await this.#pool.query<{
      relrowsecurity: boolean;
      relforcerowsecurity: boolean;
    }>(
      "SELECT relrowsecurity,relforcerowsecurity FROM pg_class WHERE oid IN (to_regclass('public.prism365_state'),to_regclass('public.prism365_audit'))",
    );
    if (
      tables.rowCount !== 2 ||
      tables.rows.some((table) => !table.relrowsecurity || !table.relforcerowsecurity)
    )
      throw new Error("Prism365 tables require the migration and forced row security.");
    const permissions = await this.#pool.query<{ can_update: boolean; can_delete: boolean }>(
      "SELECT has_table_privilege(current_user,'public.prism365_audit','UPDATE') AS can_update,has_table_privilege(current_user,'public.prism365_audit','DELETE') AS can_delete",
    );
    if (permissions.rows[0]?.can_update || permissions.rows[0]?.can_delete)
      throw new Error("The runtime audit ledger must permit append and read only.");
  }

  async #transaction<T>(tenantId: string, operation: (client: PoolClient) => Promise<T>) {
    validateTenant(tenantId);
    const client = await this.#pool.connect();
    try {
      await client.query("BEGIN");
      await client.query("SELECT set_config('prism365.tenant_id', $1, true)", [
        tenantId.toLowerCase(),
      ]);
      const result = await operation(client);
      await client.query("COMMIT");
      return result;
    } catch (error) {
      await client.query("ROLLBACK");
      throw error;
    } finally {
      client.release();
    }
  }

  async get(tenantId: string, kind: StateKind, key: string): Promise<unknown | null> {
    return this.#transaction(tenantId, async (client) => {
      const result = await client.query<{ encrypted: Buffer }>(
        "SELECT encrypted FROM prism365_state WHERE tenant_id=$1 AND kind=$2 AND key=$3 AND expires_at>now()",
        [tenantId, kind, key],
      );
      const row = result.rows[0];
      return row
        ? this.#cipher.decrypt(row.encrypted, `${tenantId.toLowerCase()}:${kind}:${key}`)
        : null;
    });
  }

  async set(tenantId: string, kind: StateKind, key: string, value: unknown, expiresAt: Date) {
    const encrypted = this.#cipher.encrypt(value, `${tenantId.toLowerCase()}:${kind}:${key}`);
    await this.#transaction(tenantId, async (client) => {
      await client.query(
        "INSERT INTO prism365_state(tenant_id,kind,key,encrypted,expires_at) VALUES($1,$2,$3,$4,$5) ON CONFLICT(tenant_id,kind,key) DO UPDATE SET encrypted=EXCLUDED.encrypted, expires_at=EXCLUDED.expires_at",
        [tenantId, kind, key, encrypted, expiresAt],
      );
    });
  }

  async take(tenantId: string, kind: StateKind, key: string): Promise<unknown | null> {
    return this.#transaction(tenantId, async (client) => {
      const result = await client.query<{ encrypted: Buffer; expires_at: Date }>(
        "DELETE FROM prism365_state WHERE tenant_id=$1 AND kind=$2 AND key=$3 RETURNING encrypted,expires_at",
        [tenantId, kind, key],
      );
      const row = result.rows[0];
      return row && row.expires_at.getTime() > Date.now()
        ? this.#cipher.decrypt(row.encrypted, `${tenantId.toLowerCase()}:${kind}:${key}`)
        : null;
    });
  }

  async replace(tenantId: string, kind: StateKind, key: string, value: unknown, expiresAt: Date) {
    const encrypted = this.#cipher.encrypt(value, `${tenantId.toLowerCase()}:${kind}:${key}`);
    return this.#transaction(tenantId, async (client) => {
      const result = await client.query(
        "UPDATE prism365_state SET encrypted=$4,expires_at=$5 WHERE tenant_id=$1 AND kind=$2 AND key=$3 AND expires_at>now()",
        [tenantId, kind, key, encrypted, expiresAt],
      );
      return result.rowCount === 1;
    });
  }

  async list(tenantId: string, kind: StateKind) {
    return this.#transaction(tenantId, async (client) => {
      const result = await client.query<{ key: string; encrypted: Buffer }>(
        "SELECT key,encrypted FROM prism365_state WHERE tenant_id=$1 AND kind=$2 AND expires_at>now() ORDER BY key LIMIT 50",
        [tenantId, kind],
      );
      return result.rows.map((row) =>
        this.#cipher.decrypt(row.encrypted, `${tenantId.toLowerCase()}:${kind}:${row.key}`),
      );
    });
  }

  async purgeExpired(tenantId: string) {
    await this.#transaction(tenantId, async (client) => {
      await client.query("DELETE FROM prism365_state WHERE tenant_id=$1 AND expires_at<=now()", [
        tenantId,
      ]);
    });
  }

  async delete(tenantId: string, kind: StateKind, key: string) {
    await this.#transaction(tenantId, async (client) => {
      await client.query("DELETE FROM prism365_state WHERE tenant_id=$1 AND kind=$2 AND key=$3", [
        tenantId,
        kind,
        key,
      ]);
    });
  }

  async appendAudit(tenantId: string, actorId: string, event: Record<string, unknown>) {
    await this.#transaction(tenantId, async (client) => {
      await client.query(
        "INSERT INTO prism365_audit(id,tenant_id,actor_id,event) VALUES($1,$2,$3,$4)",
        [randomUUID(), tenantId, actorId, JSON.stringify(event)],
      );
    });
  }

  async close() {
    await this.#pool.end();
  }
}

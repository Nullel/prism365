import { randomBytes } from "node:crypto";
import { describe, expect, it } from "vitest";
import { MemoryStateStore, StateCipher } from "../src/index.ts";

const tenantA = "11111111-1111-4111-8111-111111111111";
const tenantB = "22222222-2222-4222-8222-222222222222";

describe("tenant state storage", () => {
  it("isolates matching keys by tenant, expires state, and invalidates deleted sessions", async () => {
    const store = new MemoryStateStore();
    await store.set(
      tenantA,
      "session",
      "same-key",
      { access: "synthetic" },
      new Date(Date.now() + 60_000),
    );
    expect(await store.get(tenantB, "session", "same-key")).toBeNull();
    expect(await store.get(tenantA, "session", "same-key")).toEqual({ access: "synthetic" });
    await store.delete(tenantA, "session", "same-key");
    expect(await store.get(tenantA, "session", "same-key")).toBeNull();
    await store.set(tenantA, "session", "expired", {}, new Date(0));
    expect(await store.get(tenantA, "session", "expired")).toBeNull();
  });

  it("encrypts state and rejects cross-tenant relocation or tampering", () => {
    const cipher = new StateCipher(randomBytes(32).toString("hex"));
    const bytes = cipher.encrypt({ token: "synthetic-token" }, `${tenantA}:session:key`);
    expect(bytes.includes(Buffer.from("synthetic-token"))).toBe(false);
    expect(cipher.decrypt(bytes, `${tenantA}:session:key`)).toEqual({ token: "synthetic-token" });
    expect(() => cipher.decrypt(bytes, `${tenantB}:session:key`)).toThrow();
    bytes[bytes.length - 1] = (bytes[bytes.length - 1] ?? 0) ^ 1;
    expect(() => cipher.decrypt(bytes, `${tenantA}:session:key`)).toThrow();
  });
});

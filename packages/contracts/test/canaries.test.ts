import { describe, expect, it } from "vitest";
import {
  canaryAccountSchema,
  canaryChangeSchema,
  canaryListSchema,
  MAX_CANARY_ACCOUNTS,
} from "../src/index.ts";

const account = {
  userId: "44444444-4444-4444-8444-444444444444",
  principal: "decoy.account@example.invalid",
  addedAt: "2026-10-22T12:00:00.000Z",
  addedBy: "22222222-2222-4222-8222-222222222222",
};

describe("canary contracts", () => {
  it("holds at most 100 accounts", () => {
    expect(MAX_CANARY_ACCOUNTS).toBe(100);
    const many = (count: number) =>
      Array.from({ length: count }, (_value, index) => ({ ...account, userId: `user-${index}` }));
    expect(canaryListSchema.safeParse({ accounts: many(100) }).success).toBe(true);
    expect(canaryListSchema.safeParse({ accounts: many(101) }).success).toBe(false);
  });

  it("needs a user id, a principal name, a time and who added it", () => {
    expect(canaryAccountSchema.safeParse(account).success).toBe(true);
    for (const key of Object.keys(account))
      expect(canaryAccountSchema.safeParse({ ...account, [key]: "" }).success, key).toBe(false);
    expect(canaryAccountSchema.safeParse({ ...account, userId: "not an id" }).success).toBe(false);
    expect(canaryAccountSchema.safeParse({ ...account, addedAt: "yesterday" }).success).toBe(false);
  });

  it("accepts a partial change, and refuses one that names no account or too many", () => {
    expect(canaryChangeSchema.safeParse({ add: ["user-1"] }).success).toBe(true);
    expect(canaryChangeSchema.safeParse({ remove: ["user-1"] }).success).toBe(true);
    expect(canaryChangeSchema.safeParse({ add: ["user-1"], remove: ["user-2"] }).success).toBe(
      true,
    );
    expect(canaryChangeSchema.safeParse({}).success).toBe(false);
    expect(canaryChangeSchema.safeParse({ add: [] }).success).toBe(false);
    expect(canaryChangeSchema.safeParse({ add: ["a b"] }).success).toBe(false);
    expect(
      canaryChangeSchema.safeParse({
        add: Array.from({ length: 101 }, (_value, index) => `user-${index}`),
      }).success,
    ).toBe(false);
    // The same account on both sides is ambiguous.
    expect(canaryChangeSchema.safeParse({ add: ["user-1"], remove: ["user-1"] }).success).toBe(
      false,
    );
  });
});

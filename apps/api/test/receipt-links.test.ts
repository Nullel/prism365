import type { AuditRow } from "@prism365/contracts";
import { describe, expect, it } from "vitest";
import { itemIsLinked, runAccounts } from "../src/receipt-links.ts";

// A receipt's account is the tool's claim; it links only when the run behind the handoff lists it.
const A = "44444444-4444-4444-8444-444444444441";
const B = "44444444-4444-4444-8444-444444444442";

const row = (personId: string | null, principal: string | null): AuditRow => ({
  key: `${personId}`,
  personId,
  tags: [],
  cells: { person: { name: "Synthetic", principal } },
});

describe("receipt item links", () => {
  const accounts = runAccounts([
    row(A, "ines@example.invalid"),
    row(B, null),
    row(null, "x@y.invalid"),
  ]);

  it("link an item whose id and principal name belong to one listed account", () => {
    expect(itemIsLinked({ userId: A, userPrincipalName: "Ines@Example.invalid" }, accounts)).toBe(
      true,
    );
  });

  it("link an item that names only a listed id", () => {
    expect(itemIsLinked({ userId: B, userPrincipalName: null }, accounts)).toBe(true);
  });

  it("do not link an id with another account's principal name", () => {
    expect(itemIsLinked({ userId: A, userPrincipalName: "other@example.invalid" }, accounts)).toBe(
      false,
    );
  });

  it("do not link an id the run does not list, or a name with no id", () => {
    expect(
      itemIsLinked(
        { userId: "44444444-4444-4444-8444-444444444499", userPrincipalName: null },
        accounts,
      ),
    ).toBe(false);
    expect(
      itemIsLinked({ userId: null, userPrincipalName: "ines@example.invalid" }, accounts),
    ).toBe(false);
  });

  it("link nothing when there is no run to check against", () => {
    expect(itemIsLinked({ userId: A, userPrincipalName: null }, null)).toBe(false);
  });
});

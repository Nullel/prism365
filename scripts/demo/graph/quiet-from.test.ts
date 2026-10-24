import { describe, expect, it } from "vitest";
import { buildSyntheticTenant } from "../tenant.ts";

// The recorder builds the tenant as it stood on earlier days to leave audit runs behind. Accounts
// that went quiet keep the dates they were last seen on, so one that is past 30 days of silence
// now was still active then, and the This week page has accounts to show.

const DAY = 86_400_000;
const anchor = new Date("2026-10-22T12:00:00.000Z");
const earlier = new Date(anchor.getTime() - 7 * DAY);

function lastSuccess(tenant: ReturnType<typeof buildSyntheticTenant>, name: string): number {
  const user = tenant.users.find((candidate) => candidate.displayName === name);
  const at = user?.signInActivity?.lastSuccessfulSignInDateTime;
  if (!at) throw new Error(`${name} has no last sign-in.`);
  return Date.parse(at);
}

describe("the synthetic tenant as it stood earlier", () => {
  const now = buildSyntheticTenant(anchor);
  const then = buildSyntheticTenant(earlier, { quietFrom: anchor });
  const alone = buildSyntheticTenant(earlier);

  it("keeps the date a quiet account was last seen when only the clock moves back", () => {
    // 35 days of silence at the anchor, so 28 a week earlier: not yet past 30 days.
    expect(lastSuccess(then, "Ravenna Doyle")).toBe(lastSuccess(now, "Ravenna Doyle"));
    expect((earlier.getTime() - lastSuccess(then, "Ravenna Doyle")) / DAY).toBeGreaterThan(27);
    expect((earlier.getTime() - lastSuccess(then, "Ravenna Doyle")) / DAY).toBeLessThan(29);
    expect((anchor.getTime() - lastSuccess(now, "Ravenna Doyle")) / DAY).toBeGreaterThan(34);
  });

  it("moves every date with the clock when it is not asked to", () => {
    expect(lastSuccess(alone, "Ravenna Doyle")).toBe(lastSuccess(now, "Ravenna Doyle") - 7 * DAY);
  });

  it("leaves accounts in daily use, and the anchor's own tenant, as they were", () => {
    const withOption = buildSyntheticTenant(anchor, { quietFrom: anchor });
    expect(JSON.stringify(withOption.users)).toBe(JSON.stringify(now.users));
    expect(then.users.map((user) => user.id)).toEqual(alone.users.map((user) => user.id));
  });
});

import { describe, expect, it } from "vitest";
import {
  cleanExportText,
  exportedAtProblem,
  isStaleExport,
  MAILBOX_ACCESS_CLOCK_SKEW_MS,
  mailboxAccessExportSchema,
  mailboxAccessProblem,
  visibleLength,
} from "../src/index.ts";

// The mailbox access export format (ADR-0012). Synthetic ids and .invalid addresses only.
const MAILBOX = "eeeeeeee-0000-4000-8000-000000000001";
const DAY = 86_400_000;
const NOW = Date.parse("2026-11-06T12:00:00Z");

const base = (grants: unknown[], unresolved = 0) => ({
  format: "prism365.mailbox-access",
  version: 1,
  tenantId: "11111111-1111-4111-8111-111111111111",
  exportedAt: "2026-11-05T09:30:00Z",
  includeUserMailboxes: false,
  mailboxes: [
    {
      id: MAILBOX,
      address: "lab.results@example.invalid",
      displayName: "Lab results",
      type: "SharedMailbox",
    },
  ],
  grants,
  unresolved,
});
const grant = (index: number) => ({
  mailboxId: MAILBOX,
  right: "FullAccess",
  trustee: {
    id: `cccccccc-0000-4000-8000-${String(index).padStart(12, "0")}`,
    address: null,
    name: "P",
    kind: "user",
  },
});
const problem = (input: unknown) => {
  const parsed = mailboxAccessExportSchema.safeParse(input);
  return parsed.success ? null : mailboxAccessProblem(input, parsed.error);
};

describe("the mailbox access export", () => {
  it("takes 200,000 grants and refuses one more", () => {
    const grants = Array.from({ length: 200_000 }, (_, index) => grant(index));
    expect(problem(base(grants))).toBeNull();
    grants.push(grant(200_000));
    expect(problem(base(grants))).toBe("grants has more than 200,000 entries.");
  }, 60_000);

  it("refuses a body that is not an object, and says which field is missing", () => {
    expect(problem("text")).toMatch(/not a JSON object/);
    const { tenantId: _tenant, ...rest } = base([]);
    expect(problem(rest)).toBe("tenantId is missing.");
  });

  it("refuses an address that is not one, and a time zone offset no zone uses", () => {
    const file = base([]);
    file.mailboxes[0] = {
      ...file.mailboxes[0],
      address: "not an address",
    } as (typeof file.mailboxes)[number];
    expect(problem(file)).toBe("mailboxes[0].address is not an email address.");
    expect(problem({ ...base([]), exportedAt: "2026-11-05T09:30:00+16:00" })).toMatch(
      /beyond 15:59/,
    );
  });

  it("refuses a user or group trustee without an id", () => {
    const file = base([
      { ...grant(1), trustee: { id: null, address: null, name: "Group", kind: "group" } },
    ]);
    expect(problem(file)).toMatch(
      /grants\[0\]\.trustee\.id is missing: a user or group trustee needs its directory id/,
    );
  });

  it("removes controls, NUL and direction marks, and checks a bound on the text as sent", () => {
    expect(cleanExportText(" A\u0000B\u202EC\nD\u0085E ")).toBe("ABC D E");
    const file = base([]);
    // 257 characters as sent, even though cleaning would shorten it.
    file.mailboxes[0] = {
      ...file.mailboxes[0],
      displayName: `${"x".repeat(250)}\u0000\u0000\u0000\u0000\u0000\u0000\u0000`,
    } as (typeof file.mailboxes)[number];
    expect(problem(file)).toBe("mailboxes[0].displayName is longer than 256 characters.");
  });

  it("turns line and paragraph separators into spaces, so a name stays one line", () => {
    expect(cleanExportText("Finance\u2028Full Access directly\u2029x")).toBe(
      "Finance Full Access directly x",
    );
    const file = base([
      {
        ...grant(1),
        trustee: { id: null, address: null, name: "Finance\u2028Team", kind: "other" },
      },
    ]);
    const parsed = mailboxAccessExportSchema.parse(file);
    expect(parsed.grants[0]?.trustee.name).toBe("Finance Team");
  });

  it("refuses a name or an address made only of characters that show nothing", () => {
    expect(visibleLength("\u200B\u3164\u{E0041} ")).toBe(0);
    for (const name of ["\u200B", "\u3164", "\u200B\u200C\u2060", "\u{E0041}\u{E0042}"])
      expect(
        problem(base([{ ...grant(1), trustee: { id: null, address: null, name, kind: "other" } }])),
        JSON.stringify(name),
      ).toBe("grants[0].trustee.name is empty or shows nothing.");
    const file = base([]);
    file.mailboxes[0] = {
      ...file.mailboxes[0],
      address: "\u200B@\u200B",
    } as (typeof file.mailboxes)[number];
    expect(problem(file)).toBe("mailboxes[0].address is shorter than 3 visible characters.");
    // A visible name with an invisible character inside is still a name.
    expect(
      problem(
        base([
          { ...grant(1), trustee: { id: null, address: null, name: "Pat\u200B", kind: "other" } },
        ]),
      ),
    ).toBeNull();
  });

  it("echoes an unknown field's name only short and cleaned", () => {
    const message = problem({ ...base([]), [`x${"y".repeat(200)}\u202E`]: 1 }) ?? "";
    expect(message.length).toBeLessThan(260);
    expect(message).not.toContain("\u202E");
  });
});

describe("an export's date", () => {
  it("is accepted from 90 days back to a few minutes ahead", () => {
    expect(exportedAtProblem(new Date(NOW - 90 * DAY + 1_000).toISOString(), NOW)).toBeNull();
    expect(
      exportedAtProblem(new Date(NOW + MAILBOX_ACCESS_CLOCK_SKEW_MS).toISOString(), NOW),
    ).toBeNull();
    expect(exportedAtProblem(new Date(NOW - 90 * DAY - 1_000).toISOString(), NOW)).toMatch(
      /more than 90 days ago/,
    );
    expect(
      exportedAtProblem(new Date(NOW + MAILBOX_ACCESS_CLOCK_SKEW_MS + 1_000).toISOString(), NOW),
    ).toMatch(/in the future/);
  });

  it("is stale after 30 days", () => {
    expect(isStaleExport(new Date(NOW - 30 * DAY).toISOString(), NOW)).toBe(false);
    expect(isStaleExport(new Date(NOW - 30 * DAY - 1_000).toISOString(), NOW)).toBe(true);
  });
});

import type { AccessAuditResult } from "@prism365/contracts";
import { describe, expect, it } from "vitest";
import { type SummaryLine, type SummaryPart, summaryLines } from "../src/personAccessSummary.ts";

// What the Access tab says about an account from each audit's rows, as lines of parts.
type Cells = Record<string, string | number | boolean | string[] | null>;
const row = (cells: Cells, key = "r") => ({
  key,
  cells: Object.entries(cells).map(([name, value]) => ({
    key: name,
    label: name,
    kind: "text" as const,
    value,
  })),
});
const result = (auditId: string, rows: ReturnType<typeof row>[]): AccessAuditResult => ({
  status: "ok",
  auditId,
  name: auditId,
  finishedAt: "2026-10-20T10:00:00.000Z",
  scope: null,
  matchedBy: "account",
  rows,
  truncated: false,
});
const lines = (...results: AccessAuditResult[]) => {
  const byId = new Map(results.map((entry) => [entry.auditId, entry]));
  return (id: string) => summaryLines(byId, id);
};
const say = (line: SummaryLine) =>
  line.parts
    .map((part) => [part.prefix, part.text, part.date].filter(Boolean).join(" "))
    .join(" | ");
const tones = (line: SummaryLine) =>
  Object.fromEntries(line.parts.filter((part) => part.tone).map((part) => [part.text, part.tone]));

describe("summary lines", () => {
  it("lists eligible and time-bound roles only, one line each, with the end date", () => {
    const of = lines(
      result("directory.privileged-roles", [
        row({ role: "Global Administrator", privileged: true, held: "Permanent", ends: null }, "a"),
        row(
          { role: "Exchange Administrator", privileged: true, held: "Eligible", ends: null },
          "b",
        ),
        row(
          {
            role: "Helpdesk Administrator",
            privileged: false,
            held: "Time-bound",
            ends: "2026-12-01T00:00:00.000Z",
          },
          "c",
        ),
      ]),
    )("directory.privileged-roles");
    expect(of.map((line) => line.title)).toEqual([
      "Exchange Administrator",
      "Helpdesk Administrator",
    ]);
    expect(of.map((line) => line.privileged)).toEqual([true, false]);
    expect(say(of[0] as SummaryLine)).toBe("Eligible");
    expect(say(of[1] as SummaryLine)).toBe("Time-bound | ends 2026-12-01T00:00:00.000Z");
  });

  it("says whether MFA is registered, the strongest method and the methods", () => {
    const [line] = lines(
      result("mfa.registration", [
        row({ registered: true, strength: "Authenticator app", methods: ["Phone", "App"] }),
      ]),
    )("mfa.registration");
    expect(say(line as SummaryLine)).toBe(
      "MFA registered | Strongest method Authenticator app | Methods Phone, App",
    );
    const [none] = lines(result("mfa.registration", [row({ registered: false })]))(
      "mfa.registration",
    );
    expect(say(none as SummaryLine)).toBe("No MFA registered");
    expect(tones(none as SummaryLine)).toEqual({ "No MFA registered": "bad" });
  });

  it("gives one line per device, colours compliance and names a jailbreak only when Yes", () => {
    const of = lines(
      result("devices.managed", [
        row(
          {
            device: "PLANT-LT-01",
            os: "Windows",
            osVersion: "11 23H2",
            compliance: "Compliant",
            lastCheckIn: "2026-10-19T08:00:00.000Z",
            encrypted: true,
            jailbroken: "No",
          },
          "a",
        ),
        row(
          {
            device: "PLANT-PH-02",
            os: "iOS",
            osVersion: "17.4",
            compliance: "Not compliant",
            encrypted: false,
            jailbroken: "Yes",
          },
          "b",
        ),
      ]),
    )("devices.managed");
    expect(of.map((line) => line.title)).toEqual(["PLANT-LT-01", "PLANT-PH-02"]);
    expect(say(of[0] as SummaryLine)).toBe(
      "Windows 11 23H2 | Compliant | last check-in 2026-10-19T08:00:00.000Z | Encrypted",
    );
    expect(tones(of[0] as SummaryLine)).toEqual({ Compliant: "good" });
    expect(tones(of[1] as SummaryLine)).toEqual({
      "Not compliant": "bad",
      "Not encrypted": "warning",
      "Jailbroken or rooted": "bad",
    });
  });

  it("reads a mailbox as used of the quota, with archive, type and a status only when not OK", () => {
    const [ok] = lines(
      result("mail.mailbox-storage", [
        row({
          used: 12.34,
          sendQuota: 99,
          percent: 12.5,
          archive: true,
          recipientType: "User",
          status: "OK",
        }),
      ]),
    )("mail.mailbox-storage");
    expect(say(ok as SummaryLine)).toBe("Used 12.3 of 99 GB (12.5%) | Has an archive | User");
    const [full] = lines(
      result("mail.mailbox-storage", [
        row({ used: 49.5, sendQuota: 50, percent: 99, archive: false, status: "Near full" }),
      ]),
    )("mail.mailbox-storage");
    expect(say(full as SummaryLine)).toBe("Used 49.5 of 50 GB (99%) | No archive | Near full");
    expect(tones(full as SummaryLine)).toEqual({ "Near full": "warning" });
  });

  it("reads a password as last changed with its age, the policy and only the flags that are Yes", () => {
    const [line] = lines(
      result("accounts.passwords", [
        row({
          lastChanged: "2026-07-01T00:00:00.000Z",
          days: 112,
          policy: "180 days",
          neverExpires: false,
          weak: true,
          status: "OK",
        }),
      ]),
    )("accounts.passwords");
    expect(say(line as SummaryLine)).toBe(
      "Last changed 2026-07-01T00:00:00.000Z | 112 days old | Expiry policy 180 days | Strong password not required",
    );
    const [flagged] = lines(
      result("accounts.passwords", [row({ neverExpires: true, status: "Never expires" })]),
    )("accounts.passwords");
    expect(say(flagged as SummaryLine)).toBe("Never expires for this account | Never expires");
  });

  it("reads risk as level, state and detail", () => {
    const [line] = lines(
      result("identity.risky-users", [
        row({ level: "High", state: "At risk", detail: "Leaked credentials" }),
      ]),
    )("identity.risky-users");
    expect(say(line as SummaryLine)).toBe("High risk | State At risk | Leaked credentials");
    expect(lines(result("identity.risky-users", []))("identity.risky-users")).toEqual([]);
  });

  it("folds inactivity into activity at 30 days or more and leaves a shorter gap out", () => {
    const activity = result("accounts.activity", [
      row({
        verdict: "Likely unused",
        lastSuccess: "2026-08-01T00:00:00.000Z",
        mailLast: "2026-09-01T00:00:00.000Z",
        teamsLast: null,
      }),
    ]);
    const long = lines(
      activity,
      result("accounts.inactive", [row({ days: 81 })]),
    )("accounts.activity");
    expect(say(long[0] as SummaryLine)).toBe(
      "Likely unused | Last successful sign-in 2026-08-01T00:00:00.000Z | Inactive for 81 days",
    );
    expect(say(long[1] as SummaryLine)).toContain(
      "Mail 2026-09-01T00:00:00.000Z | Teams none | OneDrive none",
    );
    const short = lines(
      activity,
      result("accounts.inactive", [row({ days: 29 })]),
    )("accounts.activity");
    expect(say(short[0] as SummaryLine)).not.toContain("Inactive for");
  });

  it("lists a license to reclaim only when the person is a candidate", () => {
    const activity = result("accounts.activity", [row({ verdict: "Active" })]);
    const none = lines(activity, result("licenses.reclaim", []))("accounts.activity");
    expect(none.some((line) => line.title === "License to reclaim")).toBe(false);
    const some = lines(
      activity,
      result("licenses.reclaim", [
        row({ status: "Inactive 90 days", licenses: ["Business Premium"] }),
      ]),
    )("accounts.activity");
    const reclaim = some.find((line) => line.title === "License to reclaim") as SummaryLine;
    expect(say(reclaim)).toBe("Inactive 90 days | Business Premium");
  });

  it("has nothing to say for an audit with no row, or one it does not know", () => {
    expect(lines(result("mfa.registration", []))("mfa.registration")).toEqual([]);
    expect(lines()("not.an.audit")).toEqual([]);
    const part: SummaryPart = { text: "x" };
    expect(part.text).toBe("x");
  });
});

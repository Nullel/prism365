import type { AuditRow } from "@prism365/contracts";
import { type InputId, INPUTS } from "../src/combined-risks-inputs.ts";
import type { StoredFinding, StoredResults, StoredRun } from "../src/index.ts";

// Synthetic accounts for the Combined risks tests: invented names on example.invalid, made-up IDs.
// A stored run is built from the rows an input audit would have produced, with only the columns
// the Combined risks audit reads.

export const now = new Date("2026-10-22T12:00:00Z");
const DAY = 86_400_000;
export const daysAgo = (days: number) => new Date(now.getTime() - days * DAY).toISOString();

export const who = (name: string, principal: string | null = null) => ({ name, principal });

export function row(
  key: string,
  personId: string | null,
  cells: Record<string, unknown>,
  tags: string[] = [],
): AuditRow {
  return { key, personId, tags, cells: cells as AuditRow["cells"] };
}

export interface RunOf {
  rows: AuditRow[];
  finishedAt?: string;
  complete?: boolean;
  scope?: string[] | null;
}

export function storedRun(run: RunOf): StoredRun {
  return {
    finishedAt: run.finishedAt ?? daysAgo(1),
    complete: run.complete ?? true,
    scope: run.scope ?? null,
    rows: run.rows,
  };
}

export type Runs = Partial<Record<InputId, RunOf>>;

export function fakeStored(
  runs: Runs,
  findings: StoredFinding[] = [],
): StoredResults & { asked: string[] } {
  const asked: string[] = [];
  const byAudit = new Map(
    (Object.entries(runs) as [InputId, RunOf][]).map(([id, run]) => [
      INPUTS[id].auditId,
      storedRun(run),
    ]),
  );
  return {
    asked,
    async latestRun(auditId) {
      asked.push(auditId);
      return byAudit.get(auditId) ?? null;
    },
    async openFindings() {
      return findings;
    },
  };
}

export const finding = (over: Partial<StoredFinding> = {}): StoredFinding => ({
  ruleId: "signins.unusual-location",
  severity: "high",
  subjectKind: "person",
  subjectId: "u-ada",
  title: "Sign-in from a new country",
  lastSeen: daysAgo(1),
  ...over,
});

// The people of the main fixture.
//   u-ada    Global Administrator (permanent), no MFA, dormant for 90 days, one weak device
//   u-bo     Guest holding Helpdesk Administrator, at high risk, no MFA
//   u-cy     Disabled, still licensed and holding a role
//   u-di     Licensed, never signed in, a user mailbox
//   u-ed     Fine: MFA registered, active
//   u-fay    Weak password, no MFA
//   u-room   Licensed, quiet, a room mailbox
export const ids = {
  ada: "u-ada",
  bo: "u-bo",
  cy: "u-cy",
  di: "u-di",
  ed: "u-ed",
  fay: "u-fay",
  room: "u-room",
};
export const names = {
  ada: who("Ada Admin", "ada@a.example"),
  bo: who("Bo Guest", "bo_b.example#EXT#@a.example"),
  cy: who("Cy Leaver", "cy@a.example"),
  di: who("Di Newhire", "di@a.example"),
  ed: who("Ed Steady", "ed@a.example"),
  fay: who("Fay Frontline", "fay@a.example"),
  room: who("Meeting Room", "room@a.example"),
};

export function fixture(): Runs {
  const person = (key: keyof typeof ids) => names[key];
  return {
    activity: {
      rows: [
        row(ids.ada, ids.ada, {
          person: person("ada"),
          status: "Enabled",
          type: "Member",
          lastSuccess: daysAgo(90),
          days: 90,
          verdict: "Quiet everywhere",
          licenses: ["Microsoft 365 E5"],
          groupIds: ["g-admins"],
        }),
        row(ids.bo, ids.bo, {
          person: person("bo"),
          status: "Enabled",
          type: "Guest",
          lastSuccess: daysAgo(2),
          days: 2,
          verdict: "Active",
          licenses: [],
          groupIds: ["g-guests"],
        }),
        row(ids.cy, ids.cy, {
          person: person("cy"),
          status: "Disabled",
          type: "Member",
          lastSuccess: daysAgo(200),
          days: 200,
          verdict: "Quiet everywhere",
          licenses: ["Microsoft 365 E3"],
          groupIds: [],
        }),
        row(ids.di, ids.di, {
          person: person("di"),
          status: "Enabled",
          type: "Member",
          lastSuccess: null,
          days: 120,
          verdict: "Never signed in",
          licenses: ["Microsoft 365 E3"],
          groupIds: [],
        }),
        row(ids.ed, ids.ed, {
          person: person("ed"),
          status: "Enabled",
          type: "Member",
          lastSuccess: daysAgo(1),
          days: 1,
          verdict: "Active",
          licenses: ["Microsoft 365 E3"],
          groupIds: ["g-staff"],
        }),
        row(ids.fay, ids.fay, {
          person: person("fay"),
          status: "Enabled",
          type: "Member",
          lastSuccess: daysAgo(3),
          days: 3,
          verdict: "Active",
          licenses: ["Microsoft 365 E3"],
          groupIds: ["g-staff"],
        }),
        row(ids.room, ids.room, {
          person: person("room"),
          status: "Enabled",
          type: "Member",
          lastSuccess: null,
          days: 400,
          verdict: "Quiet everywhere",
          licenses: ["Exchange Online Plan 2"],
          groupIds: [],
        }),
      ],
    },
    mfa: {
      rows: [
        row(ids.ada, ids.ada, { person: person("ada"), registered: false, type: "Member" }),
        row(ids.bo, ids.bo, { person: person("bo"), registered: false, type: "Guest" }),
        row(ids.ed, ids.ed, { person: person("ed"), registered: true, type: "Member" }),
        row(ids.fay, ids.fay, { person: person("fay"), registered: false, type: "Member" }),
      ],
    },
    roles: {
      rows: [
        row("r1", ids.ada, {
          holder: person("ada"),
          holderType: "User",
          role: "Global Administrator",
          privileged: true,
          held: "Permanent",
          enabled: true,
        }),
        row("r2", ids.bo, {
          holder: person("bo"),
          holderType: "Guest",
          role: "Helpdesk Administrator",
          privileged: false,
          held: "Permanent",
          enabled: true,
        }),
        row("r3", ids.cy, {
          holder: person("cy"),
          holderType: "User",
          role: "User Administrator",
          privileged: true,
          held: "Eligible",
          enabled: false,
        }),
        row("r4", null, {
          holder: who("Sync app"),
          holderType: "Service principal",
          role: "Global Administrator",
          privileged: true,
          held: "Permanent",
          enabled: true,
        }),
      ],
    },
    risky: {
      rows: [
        row(ids.bo, ids.bo, { person: person("bo"), level: "High", state: "At risk" }),
        row(ids.ed, ids.ed, { person: person("ed"), level: "High", state: "Remediated" }),
      ],
    },
    guests: { rows: [row(ids.bo, ids.bo, { person: person("bo"), status: "Enabled" })] },
    passwords: {
      rows: [
        row(ids.fay, ids.fay, {
          person: person("fay"),
          enabled: "Enabled",
          days: 400,
          neverExpires: true,
          weak: false,
        }),
        row(ids.ed, ids.ed, {
          person: person("ed"),
          enabled: "Enabled",
          days: 30,
          neverExpires: false,
          weak: false,
        }),
      ],
    },
    devices: {
      rows: [
        row("d1", ids.ada, { device: "LAP-ADA", compliance: "Not compliant", encrypted: true }),
        row("d2", ids.ed, { device: "LAP-ED", compliance: "Compliant", encrypted: true }),
      ],
    },
    mailbox: {
      rows: [
        row("m1", null, { mailbox: "room@a.example", recipientType: "RoomMailbox" }),
        row("m2", null, { mailbox: "di@a.example", recipientType: "UserMailbox" }),
      ],
    },
  };
}

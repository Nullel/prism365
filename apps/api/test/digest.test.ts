import { randomUUID } from "node:crypto";
import {
  finishRun,
  insertRun,
  type RunOutcome,
  writeAuditSchedule,
  builtInAudits,
} from "@prism365/audits";
import type { AuditRow, AuditSummaryItem, DigestDays } from "@prism365/contracts";
import { digestSchema } from "@prism365/contracts";
import {
  INCIDENT_KIND,
  HEALTH_MODULE,
  MESSAGE_KIND,
  MESSAGES_MODULE,
} from "@prism365/module-service-health";
import {
  type AttentionInput,
  defaultAttentionPreferences,
  upsertAttention,
} from "@prism365/modules";
import { PGliteModuleDatabase, type SqlClient } from "@prism365/storage";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import {
  auditDriftSection,
  buildDigest,
  collectionSection,
  type DigestSource,
  digestWindow,
  digestSectionsRead,
  directoryChangesSection,
  microsoftSection,
  quietAccountsSection,
  signInFindingsSection,
} from "../src/digest.ts";
import { createModuleRegistry } from "../src/modules.ts";

vi.setConfig({ testTimeout: 30_000 });

// Every section builder on a real (PGlite) database with seeded rows. Synthetic tenants, people
// and Microsoft posts only; the clock is fixed so the edges of a range can be hit exactly.

const SIGNINS = "signins.audit";
const CHANGES = "directory.changes";
const INACTIVE = "accounts.inactive";
const actor = { id: "22222222-2222-4222-8222-222222222222", name: "Synthetic admin" };
const now = new Date("2026-10-22T12:00:00.000Z");
const HOUR = 3_600_000;
const DAY = 24 * HOUR;
const week = digestWindow(now, 7);
const available: DigestSource = { reason: null, collected: true };

let database: PGliteModuleDatabase;
beforeAll(async () => {
  database = await PGliteModuleDatabase.open({
    scopes: createModuleRegistry().migrationScopes(),
  });
}, 60_000);
afterAll(async () => {
  await database?.close();
});

// A tenant of its own for each group of tests, so the rows of one never reach another.
const newTenant = () => randomUUID();
const on = <T>(tenant: string, operation: (sql: SqlClient) => Promise<T>) =>
  database.tenant(tenant, operation);
const at = (milliseconds: number) => new Date(milliseconds);
const ago = (milliseconds: number) => new Date(now.getTime() - milliseconds);

// Findings --------------------------------------------------------------------------------------

interface SeededFinding {
  module?: string;
  rule?: string;
  severity?: "info" | "low" | "medium" | "high" | "critical";
  state?: "open" | "acknowledged" | "resolved";
  firstSeen: Date;
  stateChangedAt?: Date | null;
  score?: number;
  title?: string;
  subject?: {
    kind: "tenant" | "service" | "resource" | "person";
    id: string;
    label: string | null;
  };
}

async function finding(tenant: string, seed: SeededFinding): Promise<string> {
  const id = randomUUID();
  await on(tenant, (sql) =>
    sql.query(
      `INSERT INTO prism365_finding (id, tenant_id, module, rule_id, rule_version, dedupe_key,
         purpose, severity, score, subject_kind, subject_id, subject_label, title, summary,
         evidence, limitations, window_start, window_end, state, occurrences, first_seen,
         last_seen, state_changed_at)
       VALUES ($1,$2,$3,$4,1,$14,'security',$5,$6,$7,$8,$9,$10,'Synthetic.','[]'::jsonb,'[]'::jsonb,
         $11,$11,$12,1,$11,$11,$13)`,
      [
        id,
        tenant,
        seed.module ?? SIGNINS,
        seed.rule ?? "signins.unfamiliar_signin",
        seed.severity ?? "medium",
        seed.score ?? 50,
        seed.subject?.kind ?? "person",
        seed.subject?.id ?? "person-1",
        seed.subject ? seed.subject.label : "Rowan Ellis",
        seed.title ?? "An unfamiliar sign-in",
        seed.firstSeen.toISOString(),
        seed.state ?? "open",
        seed.stateChangedAt?.toISOString() ?? null,
        id,
      ],
    ),
  );
  return id;
}

// Module runs: a collector's history, which the "never collected" check and the collection
// section read.
async function moduleRun(
  tenant: string,
  module: string,
  collector: string,
  startedAt: Date,
  status: "succeeded" | "partial" | "failed" | "abandoned" | "running",
  errorCategory: string | null = null,
) {
  await on(tenant, (sql) =>
    sql.query(
      `INSERT INTO prism365_module_run (id, tenant_id, module, collector, credential_kind,
         started_at, finished_at, status, error_category)
       VALUES ($1,$2,$3,$4,'delegated',$5,$6,$7,$8)`,
      [
        randomUUID(),
        tenant,
        module,
        collector,
        startedAt.toISOString(),
        status === "running" ? null : new Date(startedAt.getTime() + 2_000).toISOString(),
        status,
        errorCategory,
      ],
    ),
  );
}

describe("sign-in findings", () => {
  it("counts the findings opened in the range by severity, from the start to just before the end", async () => {
    const tenant = newTenant();
    await finding(tenant, { firstSeen: week.from, severity: "critical" });
    await finding(tenant, { firstSeen: at(week.from.getTime() - 1), severity: "critical" });
    await finding(tenant, { firstSeen: ago(HOUR), severity: "high" });
    await finding(tenant, { firstSeen: ago(2 * HOUR), severity: "high", state: "resolved" });
    await finding(tenant, { firstSeen: at(week.to.getTime() - 1), severity: "low" });
    await finding(tenant, { firstSeen: week.to, severity: "info" });
    // Another module's findings and the collector's own gaps are not sign-in findings.
    await finding(tenant, { firstSeen: ago(HOUR), module: CHANGES, severity: "high" });
    await finding(tenant, {
      firstSeen: ago(HOUR),
      severity: "medium",
      rule: "signins.collection_gap",
    });
    const section = await on(tenant, (sql) => signInFindingsSection(sql, tenant, week, available));
    expect(section.opened).toEqual({ critical: 1, high: 2, medium: 0, low: 1, info: 0 });
    expect(section.state).toBe("ok");
    expect(section.note).toContain("4 sign-in findings opened this week");
  });

  it("lists the ten most severe findings that are still open, whenever they were opened", async () => {
    const tenant = newTenant();
    const old = ago(90 * DAY);
    await finding(tenant, { firstSeen: old, severity: "critical", title: "Old critical" });
    // The higher score is the older finding, so the score, not the age, puts it first.
    await finding(tenant, {
      firstSeen: old,
      severity: "high",
      score: 90,
      title: "High, scored 90",
    });
    await finding(tenant, {
      firstSeen: ago(80 * DAY),
      severity: "high",
      score: 10,
      title: "High, scored 10",
    });
    for (let n = 0; n < 9; n += 1)
      await finding(tenant, { firstSeen: ago(n * HOUR), severity: "low", title: `Low ${n}` });
    // Not open: acknowledged and resolved findings are not listed or counted.
    await finding(tenant, { firstSeen: old, severity: "critical", state: "acknowledged" });
    await finding(tenant, { firstSeen: old, severity: "critical", state: "resolved" });
    const section = await on(tenant, (sql) => signInFindingsSection(sql, tenant, week, available));
    expect(section.openTotal).toBe(12);
    expect(section.open).toHaveLength(10);
    expect(section.open.slice(0, 3).map((item) => item.title)).toEqual([
      "Old critical",
      "High, scored 90",
      "High, scored 10",
    ]);
    // The newest of the equally severe lows come first.
    expect(section.open[3]?.title).toBe("Low 0");
    expect(section.open.every((item) => item.severity !== "info")).toBe(true);
  });

  it("names the subject by its label, else its id, and links only people", async () => {
    const tenant = newTenant();
    await finding(tenant, {
      firstSeen: ago(HOUR),
      severity: "critical",
      subject: { kind: "person", id: "u-1", label: "Rowan Ellis" },
    });
    await finding(tenant, {
      firstSeen: ago(HOUR),
      severity: "high",
      subject: { kind: "person", id: "u-2", label: null },
    });
    await finding(tenant, {
      firstSeen: ago(HOUR),
      severity: "medium",
      subject: { kind: "resource", id: "198.51.100.7", label: null },
    });
    await finding(tenant, {
      firstSeen: ago(HOUR),
      severity: "low",
      subject: { kind: "service", id: "signins.audit", label: "Sign-in collection" },
    });
    const section = await on(tenant, (sql) => signInFindingsSection(sql, tenant, week, available));
    expect(section.open.map((item) => [item.subjectLabel, item.personId])).toEqual([
      ["Rowan Ellis", "u-1"],
      ["u-2", "u-2"],
      ["198.51.100.7", null],
      ["Sign-in collection", null],
    ]);
  });

  it("counts a finding as resolved in the range by when it was resolved, not opened", async () => {
    const tenant = newTenant();
    const old = ago(60 * DAY);
    const resolved = (changed: Date, firstSeen = old) =>
      finding(tenant, { firstSeen, state: "resolved", stateChangedAt: changed });
    await resolved(week.from);
    await resolved(at(week.from.getTime() - 1));
    await resolved(ago(HOUR), ago(2 * HOUR));
    await resolved(week.to);
    // Resolved long ago but opened in the range cannot happen; resolved in the range, opened
    // before it, is the usual case and is counted.
    const section = await on(tenant, (sql) => signInFindingsSection(sql, tenant, week, available));
    expect(section.resolved).toBe(2);
    expect(section.note).toContain("2 resolved this week");
  });

  it("says plainly when nothing was opened, is open or was resolved", async () => {
    const tenant = newTenant();
    const section = await on(tenant, (sql) => signInFindingsSection(sql, tenant, week, available));
    expect(section).toMatchObject({
      state: "empty",
      note: "No new sign-in findings this week.",
      openTotal: 0,
      resolved: 0,
    });
    const month = await on(tenant, (sql) =>
      signInFindingsSection(sql, tenant, digestWindow(now, 14), available),
    );
    expect(month.note).toBe("No new sign-in findings in the last 14 days.");
    await finding(tenant, { firstSeen: ago(30 * DAY) });
    const older = await on(tenant, (sql) => signInFindingsSection(sql, tenant, week, available));
    expect(older.state).toBe("ok");
    expect(older.note).toBe("No new sign-in findings this week. 1 still open.");
  });

  it("is unavailable, with the reason, when the module is off or nothing was ever collected", async () => {
    const tenant = newTenant();
    await finding(tenant, { firstSeen: ago(HOUR) });
    const off = await on(tenant, (sql) =>
      signInFindingsSection(sql, tenant, week, {
        reason: "The sign-in audit is not available in this installation.",
        collected: true,
      }),
    );
    expect(off).toMatchObject({
      state: "unavailable",
      note: "The sign-in audit is not available in this installation.",
      openTotal: 0,
      open: [],
    });
    const fresh = newTenant();
    const never = await on(fresh, (sql) =>
      signInFindingsSection(sql, fresh, week, { reason: null, collected: false }),
    );
    expect(never.state).toBe("unavailable");
    expect(never.note).toContain("not been collected yet");
    // Findings that exist are shown even when no run is on record (run history is kept 30 days).
    const stored = await on(tenant, (sql) =>
      signInFindingsSection(sql, tenant, week, { reason: null, collected: false }),
    );
    expect(stored.state).toBe("ok");
  });
});

// Directory changes -----------------------------------------------------------------------------

async function directoryEvent(
  tenant: string,
  id: string,
  activityAt: Date,
  category: string | null,
) {
  await on(tenant, (sql) =>
    sql.query(
      `INSERT INTO mod_directory_changes_event (tenant_id, id, activity_at, category, activity,
         result, initiated_by_kind, targets, details, search_text, processed, collected_at)
       VALUES ($1,$2,$3,$4,'Update user','success','user','[]'::jsonb,'[]'::jsonb,'',true,$3)`,
      [tenant, id, activityAt.toISOString(), category],
    ),
  );
}

describe("directory changes", () => {
  it("counts the changes in the range by category, from the start to just before the end", async () => {
    const tenant = newTenant();
    await directoryEvent(tenant, "before", at(week.from.getTime() - 1), "UserManagement");
    await directoryEvent(tenant, "start", week.from, "UserManagement");
    await directoryEvent(tenant, "role-1", ago(3 * DAY), "RoleManagement");
    await directoryEvent(tenant, "role-2", ago(2 * DAY), "RoleManagement");
    await directoryEvent(tenant, "none", ago(DAY), null);
    await directoryEvent(tenant, "last", at(week.to.getTime() - 1), "GroupManagement");
    await directoryEvent(tenant, "end", week.to, "GroupManagement");
    const section = await on(tenant, (sql) =>
      directoryChangesSection(sql, tenant, week, available),
    );
    expect(section.total).toBe(5);
    expect(section.categories).toEqual([
      { category: "RoleManagement", count: 2 },
      { category: "GroupManagement", count: 1 },
      { category: "UserManagement", count: 1 },
      { category: null, count: 1 },
    ]);
    expect(section.state).toBe("ok");
    expect(section.note).toContain("5 changes this week");
  });

  it("lists the changes the rules flagged in the range, most severe first, without collection gaps", async () => {
    const tenant = newTenant();
    const flagged = (overrides: SeededFinding) =>
      finding(tenant, { module: CHANGES, rule: "role.privileged-assigned", ...overrides });
    await flagged({ firstSeen: week.from, severity: "high", title: "Starts the range" });
    await flagged({
      firstSeen: at(week.from.getTime() - 1),
      severity: "critical",
      title: "Too early",
    });
    await flagged({ firstSeen: week.to, severity: "critical", title: "At the end" });
    await flagged({
      firstSeen: ago(DAY),
      severity: "critical",
      title: "Global Administrator assigned",
      subject: { kind: "person", id: "u-9", label: "Rowan Ellis" },
    });
    await flagged({
      firstSeen: ago(2 * DAY),
      severity: "medium",
      title: "Policy changed",
      subject: { kind: "resource", id: "policy-1", label: "Require MFA" },
      state: "resolved",
    });
    await flagged({ firstSeen: ago(HOUR), rule: "collection.gap", title: "A gap" });
    // A sign-in finding in the same range is not a directory change.
    await finding(tenant, { firstSeen: ago(HOUR), severity: "critical" });
    const section = await on(tenant, (sql) =>
      directoryChangesSection(sql, tenant, week, available),
    );
    expect(section.flaggedTotal).toBe(3);
    expect(section.flagged.map((item) => item.title)).toEqual([
      "Global Administrator assigned",
      "Starts the range",
      "Policy changed",
    ]);
    expect(section.flagged[0]).toMatchObject({ subjectLabel: "Rowan Ellis", personId: "u-9" });
    expect(section.flagged[2]).toMatchObject({ subjectLabel: "Require MFA", personId: null });
    expect(section.state).toBe("ok");
    expect(section.note).toContain("3 flagged");
  });

  it("caps the flagged list and says how many there were", async () => {
    const tenant = newTenant();
    for (let n = 0; n < 25; n += 1)
      await finding(tenant, {
        module: CHANGES,
        rule: "app.consent-granted",
        firstSeen: ago((n + 1) * HOUR),
        severity: "high",
      });
    const section = await on(tenant, (sql) =>
      directoryChangesSection(sql, tenant, week, available),
    );
    expect(section.flagged).toHaveLength(20);
    expect(section.flaggedTotal).toBe(25);
  });

  it("reports how far collection has read, and says when the range reaches before it", async () => {
    const tenant = newTenant();
    // Collection began three days ago: the history reaches back to the first change stored.
    await directoryEvent(tenant, "first", ago(3 * DAY), "UserManagement");
    const bare = await on(tenant, (sql) => directoryChangesSection(sql, tenant, week, available));
    expect(bare.coverage).toMatchObject({ phase: "none", from: ago(3 * DAY).toISOString() });
    expect(bare.note).toContain("Collection started partway through this range");
    expect(bare.note).not.toContain("still loading");
    await on(tenant, (sql) =>
      sql.query(
        `INSERT INTO prism365_cursor (tenant_id, module, collector, value, updated_at)
         VALUES ($1,$2,'audits',$3,$4)`,
        [
          tenant,
          CHANGES,
          JSON.stringify({
            v: 1,
            phase: "history",
            start: ago(3 * DAY).toISOString(),
            through: ago(HOUR).toISOString(),
            settle: { at: ago(3 * DAY).toISOString() },
          }),
          now.toISOString(),
        ],
      ),
    );
    const loading = await on(tenant, (sql) =>
      directoryChangesSection(sql, tenant, week, available),
    );
    expect(loading.coverage).toMatchObject({
      phase: "history",
      from: ago(3 * DAY).toISOString(),
      through: ago(HOUR).toISOString(),
    });
    expect(loading.note).toContain("History is still loading");
    // A range that starts after collection did has no such caveat.
    const day = await on(tenant, (sql) =>
      directoryChangesSection(sql, tenant, digestWindow(now, 1), available),
    );
    expect(day.note).not.toContain("partway");
  });

  it("is empty when a collected range holds nothing, and unavailable when nothing was collected", async () => {
    const tenant = newTenant();
    await directoryEvent(tenant, "old", ago(20 * DAY), "UserManagement");
    const empty = await on(tenant, (sql) => directoryChangesSection(sql, tenant, week, available));
    expect(empty).toMatchObject({ state: "empty", total: 0, flaggedTotal: 0 });
    expect(empty.note).toContain("No directory changes this week");
    const fresh = newTenant();
    const never = await on(fresh, (sql) => directoryChangesSection(sql, fresh, week, available));
    expect(never.state).toBe("unavailable");
    expect(never.note).toContain("not been collected yet");
    const off = await on(fresh, (sql) =>
      directoryChangesSection(sql, fresh, week, {
        reason: "The directory change feed is not available in this installation.",
      }),
    );
    expect(off).toMatchObject({
      state: "unavailable",
      note: "The directory change feed is not available in this installation.",
      total: 0,
      coverage: { phase: "none", from: null },
    });
  });
});

// Audit results ---------------------------------------------------------------------------------

interface SeededRun {
  auditId?: string;
  startedAt: Date;
  finishedAt?: Date;
  version?: number;
  scope?: string[] | null;
  complete?: boolean;
  status?: "succeeded" | "failed";
  rows?: AuditRow[];
  summary?: AuditSummaryItem[];
}

// Stores a run and finishes it the way the runner does, which leaves its trend point.
async function auditRun(tenant: string, seed: SeededRun): Promise<string> {
  const id = randomUUID();
  const auditId = seed.auditId ?? INACTIVE;
  await on(tenant, (sql) =>
    insertRun(sql, tenant, {
      id,
      auditId,
      auditVersion: seed.version ?? 3,
      startedAt: seed.startedAt,
      startedBy: actor,
      scope: seed.scope ?? null,
    }),
  );
  const outcome: RunOutcome =
    seed.status === "failed"
      ? { status: "failed", code: "failed", message: "Synthetic failure." }
      : {
          status: "succeeded",
          complete: seed.complete ?? true,
          summary: seed.summary ?? [],
          notes: [],
          rows: seed.rows ?? [],
        };
  await on(tenant, (sql) =>
    finishRun(
      sql,
      tenant,
      id,
      seed.finishedAt ?? new Date(seed.startedAt.getTime() + 5_000),
      outcome,
    ),
  );
  return id;
}

// An account row of the inactive accounts audit: `days` is how long it has been quiet, and the
// tags are the cumulative ranges the audit gives it.
function account(key: string, days: number | null, name: string | null = null): AuditRow {
  const tags = [7, 30, 60, 90].filter((threshold) => days !== null && days >= threshold);
  return {
    key,
    personId: key,
    tags: tags.map((threshold) => `d${threshold}`),
    cells: {
      person: { name: name ?? `Person ${key}`, principal: `${key}@example.invalid` },
      days,
    },
  };
}

describe("accounts that went quiet", () => {
  it("lists the accounts in the 30-day view now that were not in it in the earlier result", async () => {
    const tenant = newTenant();
    const earlier = await auditRun(tenant, {
      startedAt: ago(8 * DAY),
      rows: [
        // Active then, quiet now: went quiet.
        account("went-quiet", 20),
        // Quiet then and now: not new.
        account("long-quiet", 120),
        // Quiet then, active now: left the view.
        account("came-back", 45),
        account("always-active", 2),
      ],
    });
    const latest = await auditRun(tenant, {
      startedAt: ago(HOUR),
      rows: [
        account("went-quiet", 31, "Rowan Ellis"),
        account("long-quiet", 128),
        account("came-back", 1),
        account("always-active", 3),
        // Not in the earlier result at all (new, or outside its scope then).
        account("not-listed-before", 40),
      ],
    });
    const section = await on(tenant, (sql) => quietAccountsSection(sql, tenant, week));
    expect(section.state).toBe("ok");
    expect(section.items.map((item) => item.personId)).toEqual(["not-listed-before", "went-quiet"]);
    expect(section.total).toBe(2);
    expect(section.items[1]).toEqual({
      personId: "went-quiet",
      name: "Rowan Ellis",
      principal: "went-quiet@example.invalid",
      days: 31,
    });
    expect(section.earlier?.runId).toBe(earlier);
    expect(section.latest?.runId).toBe(latest);
    expect(section.differences).toEqual([]);
    expect(section.auditId).toBe(INACTIVE);
    expect(section.view).toBe("d30");
    expect(section.note).toContain("2 accounts");
  });

  it("compares with the newest result that started at or before the range began", async () => {
    const tenant = newTenant();
    await auditRun(tenant, { startedAt: ago(20 * DAY), rows: [account("a", 5), account("b", 5)] });
    // Exactly at the start of the range: the baseline.
    const baseline = await auditRun(tenant, {
      startedAt: week.from,
      rows: [account("a", 5), account("b", 100)],
    });
    // A moment after the start is a result from inside the range, not a baseline.
    await auditRun(tenant, {
      startedAt: at(week.from.getTime() + 1),
      rows: [account("a", 100), account("b", 100)],
    });
    // A failed run is never a baseline.
    await auditRun(tenant, { startedAt: ago(3 * DAY), status: "failed" });
    await auditRun(tenant, { startedAt: ago(HOUR), rows: [account("a", 40), account("b", 110)] });
    const section = await on(tenant, (sql) => quietAccountsSection(sql, tenant, week));
    expect(section.earlier?.runId).toBe(baseline);
    expect(section.items.map((item) => item.personId)).toEqual(["a"]);
  });

  it("ignores a result that has not started yet and takes the newest one before the end", async () => {
    const tenant = newTenant();
    await auditRun(tenant, { startedAt: ago(9 * DAY), rows: [account("a", 5)] });
    const latest = await auditRun(tenant, { startedAt: ago(HOUR), rows: [account("a", 40)] });
    await auditRun(tenant, { startedAt: at(now.getTime() + HOUR), rows: [account("a", 2)] });
    const section = await on(tenant, (sql) => quietAccountsSection(sql, tenant, week));
    expect(section.latest?.runId).toBe(latest);
    expect(section.total).toBe(1);
  });

  it("lists at most 50, longest quiet first, and still counts them all", async () => {
    const tenant = newTenant();
    await auditRun(tenant, { startedAt: ago(8 * DAY), rows: [] });
    await auditRun(tenant, {
      startedAt: ago(HOUR),
      rows: Array.from({ length: 60 }, (_, n) =>
        account(`user-${String(n).padStart(2, "0")}`, 31 + n),
      ),
    });
    const section = await on(tenant, (sql) => quietAccountsSection(sql, tenant, week));
    expect(section.total).toBe(60);
    expect(section.items).toHaveLength(50);
    expect(section.items[0]?.days).toBe(90);
    expect(section.items[49]?.days).toBe(41);
  });

  it("is empty when no account went quiet, and says how the two results compare", async () => {
    const tenant = newTenant();
    await auditRun(tenant, { startedAt: ago(8 * DAY), rows: [account("a", 100)] });
    await auditRun(tenant, { startedAt: ago(HOUR), rows: [account("a", 108)] });
    const section = await on(tenant, (sql) => quietAccountsSection(sql, tenant, week));
    expect(section).toMatchObject({ state: "empty", total: 0, items: [] });
    expect(section.note).toContain("No accounts");
  });

  it("flags results that do not compare directly: another version, another scope, a partial read", async () => {
    const check = async (
      earlier: Partial<SeededRun>,
      latest: Partial<SeededRun>,
    ): Promise<string[]> => {
      const tenant = newTenant();
      await auditRun(tenant, { startedAt: ago(8 * DAY), rows: [account("a", 5)], ...earlier });
      await auditRun(tenant, { startedAt: ago(HOUR), rows: [account("a", 40)], ...latest });
      const section = await on(tenant, (sql) => quietAccountsSection(sql, tenant, week));
      expect(section.state).toBe("ok");
      return section.differences;
    };
    expect(await check({}, {})).toEqual([]);
    expect(await check({ version: 2 }, { version: 3 })).toEqual(["version"]);
    expect(await check({ scope: ["a.example"] }, { scope: null })).toEqual(["scope"]);
    expect(
      await check({ scope: ["a.example", "b.example"] }, { scope: ["b.example", "a.example"] }),
    ).toEqual([]);
    expect(await check({ complete: false }, {})).toEqual(["incomplete"]);
    expect(await check({}, { complete: false })).toEqual(["incomplete"]);
    expect(await check({ version: 2, scope: ["a.example"], complete: false }, {})).toEqual([
      "version",
      "scope",
      "incomplete",
    ]);
    const warned = newTenant();
    await auditRun(warned, { startedAt: ago(8 * DAY), rows: [account("a", 5)], version: 2 });
    await auditRun(warned, { startedAt: ago(HOUR), rows: [account("a", 40)] });
    const section = await on(warned, (sql) => quietAccountsSection(sql, warned, week));
    expect(section.note).toContain("do not compare directly");
  });

  it("explains each reason there is nothing to compare, and suggests a schedule", async () => {
    const none = newTenant();
    const never = await on(none, (sql) => quietAccountsSection(sql, none, week));
    expect(never).toMatchObject({ state: "unavailable", total: 0, latest: null, earlier: null });
    expect(never.note).toContain("has not finished a run");

    const onlyNew = newTenant();
    await auditRun(onlyNew, { startedAt: ago(2 * DAY), rows: [account("a", 40)] });
    await auditRun(onlyNew, { startedAt: ago(HOUR), rows: [account("a", 41)] });
    const noOlder = await on(onlyNew, (sql) => quietAccountsSection(sql, onlyNew, week));
    expect(noOlder.state).toBe("unavailable");
    expect(noOlder.note).toContain("no kept result from before this range");
    expect(noOlder.note).toContain("daily or weekly");
    // Reaching back further than a week asks more of the schedule, so the note says what each covers.
    expect(noOlder.note).toContain("10");

    const stale = newTenant();
    await auditRun(stale, { startedAt: ago(30 * DAY), rows: [account("a", 40)] });
    await auditRun(stale, { startedAt: ago(20 * DAY), rows: [account("a", 41)] });
    const old = await on(stale, (sql) => quietAccountsSection(sql, stale, week));
    expect(old.state).toBe("unavailable");
    expect(old.note).toContain("older than this range");

    const onlyFailed = newTenant();
    await auditRun(onlyFailed, { startedAt: ago(HOUR), status: "failed" });
    const failed = await on(onlyFailed, (sql) => quietAccountsSection(sql, onlyFailed, week));
    expect(failed.state).toBe("unavailable");
    expect(failed.note).toContain("has not finished a run");
  });

  it("reads rows with missing cells without failing", async () => {
    const tenant = newTenant();
    await auditRun(tenant, { startedAt: ago(8 * DAY), rows: [] });
    await auditRun(tenant, {
      startedAt: ago(HOUR),
      rows: [
        { key: "bare", personId: null, tags: ["d30"], cells: {} },
        { key: "odd", personId: "odd", tags: ["d30"], cells: { person: "text", days: "many" } },
      ],
    });
    const section = await on(tenant, (sql) => quietAccountsSection(sql, tenant, week));
    // Neither has a day count or a name, so they fall back to their keys.
    expect(section.items).toEqual([
      { personId: null, name: null, principal: null, days: null },
      { personId: "odd", name: null, principal: null, days: null },
    ]);
  });
});

describe("audit drift", () => {
  const tiles = (inactive: number, attempts = 0, accounts = 100): AuditSummaryItem[] => [
    { label: "Accounts", value: accounts, tone: "neutral", view: "all" },
    { label: "Inactive 30+ days", value: inactive, tone: "warning", view: "d30" },
    {
      label: "Attempts, no success",
      value: attempts,
      tone: attempts > 0 ? "critical" : "neutral",
      view: "no-success",
    },
  ];
  const drift = (tenant: string, window = week) =>
    on(tenant, (sql) => auditDriftSection(sql, tenant, window));

  it("compares each tile at the start of the range, the newest point at or before it, with now", async () => {
    const tenant = newTenant();
    await auditRun(tenant, { startedAt: ago(20 * DAY), summary: tiles(2) });
    const start = await auditRun(tenant, {
      startedAt: ago(8 * DAY),
      finishedAt: week.from,
      summary: tiles(10, 0),
    });
    // A point a moment after the start is inside the range, not the start of it.
    await auditRun(tenant, {
      startedAt: ago(5 * DAY),
      finishedAt: at(week.from.getTime() + 1),
      summary: tiles(99, 9),
    });
    const newest = await auditRun(tenant, { startedAt: ago(2 * HOUR), summary: tiles(15, 3) });
    const section = await drift(tenant);
    expect(section.state).toBe("ok");
    expect(section.compared).toBe(1);
    const [audit] = section.audits;
    expect(audit?.auditId).toBe(INACTIVE);
    expect(audit?.title).toBe("Inactive accounts");
    expect(audit?.from.runId).toBe(start);
    expect(audit?.to.runId).toBe(newest);
    // Only the tiles that changed, the critical one first.
    expect(audit?.tiles).toEqual([
      {
        id: "attempts-no-success",
        label: "Attempts, no success",
        tone: "critical",
        from: 0,
        to: 3,
        change: 3,
      },
      {
        id: "inactive-30-days",
        label: "Inactive 30+ days",
        tone: "warning",
        from: 10,
        to: 15,
        change: 5,
      },
    ]);
    expect(audit?.differences).toEqual([]);
    expect(section.note).toContain("1 audit changed this week");
  });

  it("reports a fall as a negative change and puts critical and warning tiles first", async () => {
    const tenant = newTenant();
    await auditRun(tenant, {
      startedAt: ago(9 * DAY),
      summary: [
        { label: "Fine", value: 5, tone: "good", view: null },
        { label: "Plain", value: 5, tone: "neutral", view: null },
        { label: "Careful", value: 9, tone: "warning", view: null },
        { label: "Urgent", value: 4, tone: "critical", view: null },
      ],
    });
    await auditRun(tenant, {
      startedAt: ago(HOUR),
      summary: [
        { label: "Fine", value: 6, tone: "good", view: null },
        { label: "Plain", value: 7, tone: "neutral", view: null },
        { label: "Careful", value: 2, tone: "warning", view: null },
        { label: "Urgent", value: 1, tone: "critical", view: null },
      ],
    });
    const [audit] = (await drift(tenant)).audits;
    expect(audit?.tiles.map((tile) => [tile.label, tile.change])).toEqual([
      ["Urgent", -3],
      ["Careful", -7],
      ["Fine", 1],
      ["Plain", 2],
    ]);
  });

  it("orders the audits by their most serious change, then by the catalog", async () => {
    const tenant = newTenant();
    const two = (warning: number, critical?: number): AuditSummaryItem[] => [
      { label: "Count", value: warning, tone: "warning", view: null },
      ...(critical === undefined
        ? []
        : [{ label: "Bad", value: critical, tone: "critical" as const, view: null }]),
    ];
    for (const auditId of [INACTIVE, "mfa.registration", "guests.review"]) {
      await auditRun(tenant, { auditId, startedAt: ago(9 * DAY), summary: two(1, 0) });
    }
    await auditRun(tenant, { auditId: INACTIVE, startedAt: ago(HOUR), summary: two(5, 0) });
    await auditRun(tenant, {
      auditId: "mfa.registration",
      startedAt: ago(HOUR),
      summary: two(1, 2),
    });
    await auditRun(tenant, { auditId: "guests.review", startedAt: ago(HOUR), summary: two(8, 0) });
    const section = await drift(tenant);
    expect(section.audits.map((audit) => audit.auditId)).toEqual([
      "mfa.registration",
      INACTIVE,
      "guests.review",
    ]);
  });

  it("says when two points do not compare directly, whichever of them is the odd one", async () => {
    const check = async (earlier: Partial<SeededRun>, latest: Partial<SeededRun>) => {
      const tenant = newTenant();
      await auditRun(tenant, { startedAt: ago(9 * DAY), summary: tiles(1), ...earlier });
      await auditRun(tenant, { startedAt: ago(HOUR), summary: tiles(2), ...latest });
      return (await drift(tenant)).audits[0]?.differences;
    };
    expect(await check({}, {})).toEqual([]);
    expect(await check({ version: 2 }, {})).toEqual(["version"]);
    expect(await check({ scope: ["a.example"] }, {})).toEqual(["scope"]);
    expect(await check({}, { scope: ["a.example"] })).toEqual(["scope"]);
    expect(await check({ scope: ["b.example", "a.example"] }, { scope: ["a.example"] })).toEqual([
      "scope",
    ]);
    expect(await check({ complete: false }, {})).toEqual(["incomplete"]);
    expect(await check({}, { complete: false })).toEqual(["incomplete"]);
    expect(await check({ version: 2, scope: ["a.example"], complete: false }, {})).toEqual([
      "version",
      "scope",
      "incomplete",
    ]);
  });

  it("does not compare a tile that only one of the two points has", async () => {
    const tenant = newTenant();
    await auditRun(tenant, { startedAt: ago(9 * DAY), summary: tiles(1) });
    await auditRun(tenant, {
      startedAt: ago(HOUR),
      summary: [
        ...tiles(1),
        { label: "Outside the scope", value: 4, tone: "critical", view: null },
      ],
    });
    const section = await drift(tenant);
    expect(section.compared).toBe(1);
    expect(section.audits).toEqual([]);
    expect(section.state).toBe("empty");
    expect(section.note).toBe("No audit tile changed this week.");
  });

  it("counts audits it cannot compare: no point from before the range, or none newer", async () => {
    const tenant = newTenant();
    // Only points inside the range: nothing to compare them with.
    await auditRun(tenant, {
      auditId: "mfa.registration",
      startedAt: ago(2 * DAY),
      summary: tiles(1),
    });
    await auditRun(tenant, {
      auditId: "mfa.registration",
      startedAt: ago(HOUR),
      summary: tiles(2),
    });
    // The newest point is older than the range: nothing newer.
    await auditRun(tenant, {
      auditId: "guests.review",
      startedAt: ago(30 * DAY),
      summary: tiles(1),
    });
    await auditRun(tenant, {
      auditId: "guests.review",
      startedAt: ago(20 * DAY),
      summary: tiles(2),
    });
    // Points of an audit this build does not know are left out.
    await auditRun(tenant, {
      auditId: "retired.audit",
      startedAt: ago(9 * DAY),
      summary: tiles(1),
    });
    await auditRun(tenant, { auditId: "retired.audit", startedAt: ago(HOUR), summary: tiles(2) });
    const none = await drift(tenant);
    expect(none).toMatchObject({ state: "unavailable", compared: 0, uncompared: 2, audits: [] });
    expect(none.note).toContain("daily or weekly");
    await auditRun(tenant, { auditId: INACTIVE, startedAt: ago(9 * DAY), summary: tiles(1) });
    await auditRun(tenant, { auditId: INACTIVE, startedAt: ago(HOUR), summary: tiles(4) });
    const some = await drift(tenant);
    expect(some).toMatchObject({ state: "ok", compared: 1, uncompared: 2 });
    expect(some.note).toContain("2 could not be compared");
  });

  it("ignores a point from the future, and a failed run leaves none", async () => {
    const tenant = newTenant();
    await auditRun(tenant, { startedAt: ago(9 * DAY), summary: tiles(1) });
    await auditRun(tenant, { startedAt: ago(2 * HOUR), summary: tiles(2) });
    await auditRun(tenant, { startedAt: ago(HOUR), status: "failed" });
    await auditRun(tenant, {
      startedAt: at(now.getTime() + HOUR),
      finishedAt: week.to,
      summary: tiles(50),
    });
    const [audit] = (await drift(tenant)).audits;
    expect(audit?.tiles.find((tile) => tile.id === "inactive-30-days")?.to).toBe(2);
  });

  it("is unavailable, and says why, when no audit has a result yet", async () => {
    const tenant = newTenant();
    const section = await drift(tenant);
    expect(section).toMatchObject({ state: "unavailable", compared: 0, uncompared: 0 });
    expect(section.note).toContain("daily or weekly");
  });
});

// Microsoft -------------------------------------------------------------------------------------

function attention(
  kind: string,
  ref: string,
  options: {
    title?: string;
    services?: string[];
    deadline?: Date | null;
    active?: boolean;
    observedAt: Date;
    severity?: AttentionInput["severity"];
  },
): AttentionInput {
  const active = options.active ?? true;
  return {
    kind,
    ref,
    class: kind === INCIDENT_KIND ? "incident" : "message.action",
    title: options.title ?? ref,
    summary: "Synthetic.",
    services: options.services ?? ["Exchange Online"],
    provider: {
      status: active ? "Investigating" : "Service restored",
      labels: [],
      deadline: options.deadline ?? null,
      updatedAt: options.observedAt,
      active,
      impact: null,
      workaround: null,
    },
    localLabels: [],
    severity: options.severity ?? "medium",
    link: null,
    meaningful: { Title: options.title ?? ref, Active: active },
    observedAt: options.observedAt,
  };
}

async function report(tenant: string, module: string, input: AttentionInput): Promise<string> {
  return on(
    tenant,
    async (sql) =>
      (await upsertAttention(sql, tenant, module, input, defaultAttentionPreferences)).id,
  );
}

describe("Microsoft", () => {
  const microsoft = (tenant: string, sources = { health: available, messages: available }) =>
    on(tenant, (sql) => microsoftSection(sql, tenant, week, sources));

  it("lists the posts with an act-by date in the next 30 days, soonest first", async () => {
    const tenant = newTenant();
    const post = (ref: string, deadline: Date | null, title = ref) =>
      report(
        tenant,
        MESSAGES_MODULE,
        attention(MESSAGE_KIND, ref, { title, deadline, observedAt: ago(DAY) }),
      );
    const later = await post("later", at(now.getTime() + 20 * DAY), "Later post");
    const sooner = await post("sooner", at(now.getTime() + 2 * DAY), "Sooner post");
    await post("now", now, "Due now");
    // A past date, a date exactly 30 days out and a post with no date are not in the next 30 days.
    await post("past", at(now.getTime() - 1));
    await post("edge", at(now.getTime() + 30 * DAY));
    await post("undated", null);
    const section = await microsoft(tenant);
    expect(section.posts.total).toBe(3);
    expect(section.posts.items.map((item) => item.title)).toEqual([
      "Due now",
      "Sooner post",
      "Later post",
    ]);
    expect(section.posts.items[1]).toMatchObject({
      id: sooner,
      actBy: at(now.getTime() + 2 * DAY).toISOString(),
      services: ["Exchange Online"],
    });
    expect(section.posts.items[2]?.id).toBe(later);
    expect(section.posts.state).toBe("ok");
    expect(section.posts.note).toContain("3 Message center posts have an act-by date");
  });

  it("leaves out posts that were handled, muted or withdrawn, as the notification list does", async () => {
    const tenant = newTenant();
    const deadline = at(now.getTime() + 5 * DAY);
    const kept = await report(
      tenant,
      MESSAGES_MODULE,
      attention(MESSAGE_KIND, "kept", { deadline, observedAt: ago(DAY) }),
    );
    const done = await report(
      tenant,
      MESSAGES_MODULE,
      attention(MESSAGE_KIND, "done", { deadline, observedAt: ago(DAY) }),
    );
    const muted = await report(
      tenant,
      MESSAGES_MODULE,
      attention(MESSAGE_KIND, "muted", { deadline, observedAt: ago(DAY) }),
    );
    await report(
      tenant,
      MESSAGES_MODULE,
      attention(MESSAGE_KIND, "withdrawn", { deadline, active: false, observedAt: ago(DAY) }),
    );
    await on(tenant, async (sql) => {
      await sql.query(`UPDATE prism365_attention SET review_state='done' WHERE id=$1`, [done]);
      await sql.query(`UPDATE prism365_attention SET muted=true WHERE id=$1`, [muted]);
    });
    const section = await microsoft(tenant);
    expect(section.posts.items.map((item) => item.id)).toEqual([kept]);
  });

  it("caps the posts and says how many there were", async () => {
    const tenant = newTenant();
    for (let n = 0; n < 23; n += 1)
      await report(
        tenant,
        MESSAGES_MODULE,
        attention(MESSAGE_KIND, `post-${n}`, {
          deadline: at(now.getTime() + (n + 1) * HOUR),
          observedAt: ago(DAY),
        }),
      );
    const section = await microsoft(tenant);
    expect(section.posts.total).toBe(23);
    expect(section.posts.items).toHaveLength(20);
  });

  it("counts incidents opened and resolved in the range, and lists the ones still open", async () => {
    const tenant = newTenant();
    const incident = (ref: string, options: Parameters<typeof attention>[2]) =>
      report(tenant, HEALTH_MODULE, attention(INCIDENT_KIND, ref, options));
    // Opened in the range, still open.
    const open = await incident("EX1", {
      title: "Mail delays",
      observedAt: ago(DAY),
      severity: "high",
      services: ["Exchange Online", "Microsoft Teams"],
    });
    // Opened at the very start of the range, then resolved inside it.
    await incident("EX2", { observedAt: week.from });
    await incident("EX2", { observedAt: ago(2 * HOUR), active: false });
    // Opened just before the range, resolved inside it: resolved only.
    await incident("EX3", { observedAt: at(week.from.getTime() - 1) });
    await incident("EX3", { observedAt: ago(HOUR), active: false });
    // Opened and resolved before the range, and opened at its end.
    await incident("EX4", { observedAt: ago(20 * DAY) });
    await incident("EX4", { observedAt: ago(19 * DAY), active: false });
    await incident("EX5", { observedAt: week.to });
    // Resolved exactly at the start of the range counts; exactly at its end does not.
    await incident("EX7", { observedAt: ago(20 * DAY) });
    await incident("EX7", { observedAt: week.from, active: false });
    await incident("EX8", { observedAt: ago(10 * DAY) });
    await incident("EX8", { observedAt: week.to, active: false });
    // A still open incident from before the range is listed, and not counted as opened.
    const older = await incident("EX6", { observedAt: ago(15 * DAY), severity: "critical" });
    const section = await microsoft(tenant);
    expect(section.incidents).toMatchObject({ opened: 2, resolved: 3, openTotal: 3, state: "ok" });
    expect(section.incidents.open.map((item) => item.id)).toEqual([
      older,
      open,
      expect.any(String),
    ]);
    expect(section.incidents.open[1]).toMatchObject({
      title: "Mail delays",
      severity: "high",
      status: "Investigating",
      services: ["Exchange Online", "Microsoft Teams"],
      firstSeen: ago(DAY).toISOString(),
    });
    expect(section.incidents.note).toContain("2 opened and 3 resolved this week");
    expect(section.incidents.note).toContain("3 open now");
  });

  it("combines the two sources, and marks one that is off without hiding the other", async () => {
    const tenant = newTenant();
    const off = {
      reason: "Message center is not available in this installation.",
      collected: true,
    };
    const partial = await microsoft(tenant, { health: available, messages: off });
    expect(partial.posts).toMatchObject({
      state: "unavailable",
      note: "Message center is not available in this installation.",
      items: [],
    });
    expect(partial.incidents.state).toBe("empty");
    expect(partial.incidents.note).toBe("Microsoft reported no incidents this week.");
    expect(partial.state).toBe("empty");
    const neither = await microsoft(tenant, {
      health: { reason: "Service health is off.", collected: true },
      messages: off,
    });
    expect(neither.state).toBe("unavailable");
    expect(neither.note).toContain("Service health is off.");
    expect(neither.note).toContain("Message center is not available");
    await report(tenant, HEALTH_MODULE, attention(INCIDENT_KIND, "EX9", { observedAt: ago(HOUR) }));
    expect((await microsoft(tenant, { health: available, messages: off })).state).toBe("ok");
  });

  it("is unavailable, not empty, for a source that was never checked", async () => {
    const tenant = newTenant();
    const never = { reason: null, collected: false };
    const section = await microsoft(tenant, { health: never, messages: never });
    expect(section.state).toBe("unavailable");
    expect(section.posts.note).toContain("not been checked yet");
    expect(section.incidents.note).toContain("not been checked yet");
    await report(
      tenant,
      MESSAGES_MODULE,
      attention(MESSAGE_KIND, "stored", {
        deadline: at(now.getTime() + DAY),
        observedAt: ago(DAY),
      }),
    );
    expect((await microsoft(tenant, { health: never, messages: never })).posts.state).toBe("ok");
  });
});

// Collection ------------------------------------------------------------------------------------

describe("collection", () => {
  const modules = [
    {
      id: SIGNINS,
      title: "Sign-in audit",
      collectors: [{ id: "signins", title: "Sign-ins" }],
    },
    {
      id: CHANGES,
      title: "Directory changes",
      collectors: [{ id: "audits", title: "Directory audit log" }],
    },
  ];
  const collection = (tenant: string, collecting = true, list = modules) =>
    on(tenant, (sql) => collectionSection(sql, tenant, week, { collecting, modules: list }));

  it("gives each collector's last success and newest run, and whether collection is on", async () => {
    const tenant = newTenant();
    await moduleRun(tenant, SIGNINS, "signins", ago(3 * HOUR), "succeeded");
    await moduleRun(tenant, SIGNINS, "signins", ago(2 * HOUR), "partial");
    await moduleRun(tenant, SIGNINS, "signins", ago(HOUR), "failed", "throttled");
    const on1 = await collection(tenant, true);
    expect(on1.state).toBe("ok");
    expect(on1.collecting).toBe(true);
    expect(on1.collectors).toEqual([
      {
        module: SIGNINS,
        moduleTitle: "Sign-in audit",
        collector: "signins",
        collectorTitle: "Sign-ins",
        lastSuccess: new Date(ago(3 * HOUR).getTime() + 2_000).toISOString(),
        lastRun: { status: "failed", at: ago(HOUR).toISOString() },
      },
      {
        module: CHANGES,
        moduleTitle: "Directory changes",
        collector: "audits",
        collectorTitle: "Directory audit log",
        lastSuccess: null,
        lastRun: null,
      },
    ]);
    const off = await collection(tenant, false);
    expect(off.collecting).toBe(false);
    expect(off.note).toContain("Collection is off");
  });

  it("counts failed runs in the range, from the start to just before the end, and lists the newest categories once each", async () => {
    const tenant = newTenant();
    await moduleRun(tenant, SIGNINS, "signins", at(week.from.getTime() - 1), "failed", "timeout");
    await moduleRun(tenant, SIGNINS, "signins", week.from, "failed", "storage_error");
    await moduleRun(tenant, SIGNINS, "signins", ago(2 * DAY), "failed", "throttled");
    await moduleRun(tenant, CHANGES, "audits", ago(DAY), "failed", "access_denied");
    await moduleRun(tenant, SIGNINS, "signins", ago(HOUR), "failed", "throttled");
    await moduleRun(tenant, SIGNINS, "signins", ago(30 * 60_000), "failed", null);
    await moduleRun(tenant, SIGNINS, "signins", week.to, "failed", "timeout");
    // Runs that did not fail, and runs of modules that are not installed, are not failures.
    await moduleRun(tenant, SIGNINS, "signins", ago(HOUR), "partial", "timeout");
    await moduleRun(tenant, SIGNINS, "signins", ago(HOUR), "abandoned");
    await moduleRun(tenant, "retired.module", "poll", ago(HOUR), "failed", "collector_error");
    const section = await collection(tenant);
    expect(section.failedRuns).toBe(5);
    expect(section.errors).toEqual(["throttled", "access_denied", "storage_error"]);
    expect(section.note).toContain("5 collector runs failed this week");
  });

  it("keeps the newest five categories", async () => {
    const tenant = newTenant();
    const categories = ["a", "b", "c", "d", "e", "f", "g"];
    for (const [n, category] of categories.entries())
      await moduleRun(tenant, SIGNINS, "signins", ago((n + 1) * HOUR), "failed", category);
    const section = await collection(tenant);
    expect(section.failedRuns).toBe(7);
    expect(section.errors).toEqual(["a", "b", "c", "d", "e"]);
  });

  it("counts the gaps the collectors recorded in the range, of any state", async () => {
    const tenant = newTenant();
    const gap = (module: string, rule: string, firstSeen: Date, state?: "resolved") =>
      finding(tenant, { module, rule, firstSeen, ...(state ? { state } : {}) });
    await gap(CHANGES, "collection.gap", week.from);
    await gap(CHANGES, "collection.gap", ago(DAY), "resolved");
    await gap(SIGNINS, "signins.collection_gap", ago(HOUR));
    await gap(CHANGES, "collection.gap", at(week.from.getTime() - 1));
    await gap(SIGNINS, "signins.collection_gap", week.to);
    // A gap-like rule name in another module, and a different rule, are not gaps.
    await gap(SIGNINS, "collection.gap", ago(HOUR));
    await gap(CHANGES, "role.privileged-assigned", ago(HOUR));
    const section = await collection(tenant);
    expect(section.gaps).toBe(3);
    expect(section.note).toContain("3 gaps");
  });

  it("lists the audits that run on their own, in the catalog's order, with when each last ran", async () => {
    const tenant = newTenant();
    await on(tenant, (sql) =>
      writeAuditSchedule(
        sql,
        tenant,
        {
          audits: {
            "mfa.registration": "daily",
            [INACTIVE]: "weekly",
            "guests.review": "off",
          },
        },
        actor.id,
        builtInAudits,
      ),
    );
    await auditRun(tenant, { startedAt: ago(3 * DAY) });
    const newest = ago(2 * HOUR);
    await auditRun(tenant, { startedAt: newest, status: "failed" });
    // A run of an audit that is off does not make it scheduled.
    await auditRun(tenant, { auditId: "guests.review", startedAt: ago(HOUR) });
    const section = await collection(tenant);
    expect(section.audits).toEqual([
      {
        auditId: INACTIVE,
        title: "Inactive accounts",
        frequency: "weekly",
        lastRun: { status: "failed", at: newest.toISOString() },
      },
      { auditId: "mfa.registration", title: "MFA registration", frequency: "daily", lastRun: null },
    ]);
  });

  it("reports an installation with nothing to collect as empty", async () => {
    const tenant = newTenant();
    const section = await collection(tenant, false, []);
    expect(section).toMatchObject({ state: "empty", collectors: [], audits: [] });
  });
});

// The whole digest ------------------------------------------------------------------------------

describe("the digest", () => {
  const registryModules = createModuleRegistry()
    .list()
    .map((module) => ({
      id: module.id,
      title: module.title,
      collectors: module.collectors.map(({ id, title }) => ({ id, title })),
    }));
  const build = (tenant: string, days: DigestDays, unavailable = {}) =>
    on(tenant, (sql) =>
      buildDigest(sql, {
        tenantId: tenant,
        now,
        days,
        collecting: true,
        modules: registryModules,
        unavailable,
      }),
    );

  it("fills every section for the range and matches its contract", async () => {
    const tenant = newTenant();
    await moduleRun(tenant, SIGNINS, "signins", ago(HOUR), "succeeded");
    await moduleRun(tenant, CHANGES, "audits", ago(HOUR), "succeeded");
    await moduleRun(tenant, HEALTH_MODULE, "health", ago(HOUR), "succeeded");
    await moduleRun(tenant, MESSAGES_MODULE, "messages", ago(HOUR), "succeeded");
    await finding(tenant, { firstSeen: ago(HOUR), severity: "high" });
    await directoryEvent(tenant, "e1", ago(HOUR), "UserManagement");
    await auditRun(tenant, {
      startedAt: ago(8 * DAY),
      rows: [account("a", 5)],
      summary: [{ label: "Inactive 30+ days", value: 1, tone: "warning", view: "d30" }],
    });
    await auditRun(tenant, {
      startedAt: ago(HOUR),
      rows: [account("a", 40)],
      summary: [{ label: "Inactive 30+ days", value: 2, tone: "warning", view: "d30" }],
    });
    const digest = await build(tenant, 7);
    expect(digestSchema.parse(digest)).toEqual(digest);
    expect(digest).toMatchObject({
      days: 7,
      from: week.from.toISOString(),
      generatedAt: now.toISOString(),
      signInFindings: { state: "ok" },
      directoryChanges: { state: "ok", total: 1 },
      quietAccounts: { state: "ok", total: 1 },
      auditDrift: { state: "ok", compared: 1 },
      microsoft: { state: "empty" },
      collection: { state: "ok", collecting: true },
    });
    expect(digest.collection.collectors).toHaveLength(
      registryModules.flatMap((module) => module.collectors).length,
    );
  });

  it("uses the range it was given", async () => {
    const tenant = newTenant();
    await moduleRun(tenant, SIGNINS, "signins", ago(HOUR), "succeeded");
    await finding(tenant, { firstSeen: ago(10 * DAY) });
    const counts = async (days: DigestDays) =>
      Object.values((await build(tenant, days)).signInFindings.opened).reduce((a, b) => a + b, 0);
    expect(await counts(1)).toBe(0);
    expect(await counts(7)).toBe(0);
    expect(await counts(14)).toBe(1);
    expect(await counts(30)).toBe(1);
    expect((await build(tenant, 30)).from).toBe(ago(30 * DAY).toISOString());
  });

  it("marks only the sections whose source is off, each with its reason", async () => {
    const tenant = newTenant();
    const digest = await build(tenant, 7, {
      signins: "The sign-in audit is off.",
      changes: "The change feed is off.",
      health: "Service health is off.",
      messages: "Message center is off.",
    });
    expect(digest.signInFindings).toMatchObject({
      state: "unavailable",
      note: "The sign-in audit is off.",
    });
    expect(digest.directoryChanges).toMatchObject({
      state: "unavailable",
      note: "The change feed is off.",
    });
    expect(digest.microsoft.state).toBe("unavailable");
    // Audit results and collection runs are core data, not a module's.
    expect(digest.quietAccounts.state).toBe("unavailable");
    expect(digest.collection.state).toBe("ok");
    const partial = await build(tenant, 7, { changes: "The change feed is off." });
    expect(partial.signInFindings.note).toContain("not been collected yet");
    expect(partial.directoryChanges.note).toBe("The change feed is off.");
  });
});

describe("the sections a read touches", () => {
  it("names every section when every source is on", () => {
    expect(digestSectionsRead({})).toEqual([
      "signInFindings",
      "directoryChanges",
      "quietAccounts",
      "auditDrift",
      "microsoft",
      "collection",
    ]);
  });

  it("leaves out the section of a source that is off, and keeps the core ones", () => {
    expect(digestSectionsRead({ signins: "off" })).not.toContain("signInFindings");
    expect(digestSectionsRead({ signins: "off" })).toContain("directoryChanges");
    expect(digestSectionsRead({ changes: "off" })).not.toContain("directoryChanges");
    expect(digestSectionsRead({ changes: "off" })).toContain("signInFindings");
    const everything = digestSectionsRead({
      signins: "off",
      changes: "off",
      health: "off",
      messages: "off",
    });
    expect(everything).toEqual(["quietAccounts", "auditDrift", "collection"]);
  });

  it("reads Microsoft's section while either of its two sources is on", () => {
    expect(digestSectionsRead({ health: "off" })).toContain("microsoft");
    expect(digestSectionsRead({ messages: "off" })).toContain("microsoft");
    expect(digestSectionsRead({ health: "off", messages: "off" })).not.toContain("microsoft");
  });
});

import { auditRowSchema, COMBINED_RISKS_AUDIT_ID, COMBINED_RISKS_TILES } from "@prism365/contracts";
import { describe, expect, it } from "vitest";
import { type InputId, joinInputs, readInputs } from "../src/combined-risks-inputs.ts";
import {
  type AuditContext,
  auditDefinition,
  builtInAudits,
  combinedRisksAudit,
  createCombinedRisksAudit,
  INPUT_IDS,
  INPUTS,
  RULES,
  SECURITY_DEFAULTS_KEY,
  SECURITY_DEFAULTS_ON,
} from "../src/index.ts";
import { fake } from "./apps-devices-usage-support.ts";
import {
  daysAgo,
  fakeStored,
  finding,
  fixture,
  ids,
  names,
  now,
  type Runs,
  row,
  who,
} from "./combined-risks-support.ts";

// The audit as a whole, over synthetic stored results. Conditional Access is a fake network.

const POLICY_PATH = "/v1.0/identity/conditionalAccess/policies";
const mfaPolicy = (over: Record<string, unknown> = {}) => ({
  id: "ca-1",
  displayName: "Require MFA for everyone",
  state: "enabled",
  conditions: { users: { includeUsers: ["All"], excludeUsers: [] } },
  grantControls: { operator: "OR", builtInControls: ["mfa"] },
  ...over,
});

interface Setup {
  runs?: Runs;
  findings?: ReturnType<typeof finding>[];
  policies?: unknown[] | Response;
  granted?: string[];
  scope?: ReadonlySet<string> | null;
}

async function run(setup: Setup = {}) {
  const audit = createCombinedRisksAudit({ pause: async () => {} });
  const stored = fakeStored(setup.runs ?? fixture(), setup.findings ?? []);
  const network = fake(
    audit,
    (url) => {
      if (url.pathname !== POLICY_PATH) return undefined;
      const policies = setup.policies ?? [];
      return policies instanceof Response ? policies : Response.json({ value: policies });
    },
    {
      granted: setup.granted ?? [],
      scope: setup.scope ?? null,
    },
  );
  const context: AuditContext = { ...network.context, stored };
  const result = await audit.run(context);
  return { result, requests: network.requests, stored };
}

const rowFor = (rows: { key: string }[], key: string) => {
  const found = rows.find((item) => item.key === key) as ReturnType<typeof row> | undefined;
  if (!found) throw new Error(`No row ${key}`);
  return found;
};
const tile = (result: Awaited<ReturnType<typeof run>>["result"], label: string) =>
  result.summary.find((item) => item.label === label)?.value;
const ruleLabels = (id: string, result: Awaited<ReturnType<typeof run>>["result"]) =>
  (rowFor(result.rows, id).cells.rules as string[]).sort();
const label = (id: string) => RULES.find((item) => item.id === id)?.label ?? id;

describe("the published definition", () => {
  const definition = auditDefinition(combinedRisksAudit);

  it("is listed last, needs no permission, and offers Conditional Access as an option", () => {
    expect(builtInAudits[builtInAudits.length - 1]?.definition.id).toBe(COMBINED_RISKS_AUDIT_ID);
    expect(definition).toMatchObject({
      id: "identity.combined-risks",
      permissions: [],
      optionalPermissions: ["Policy.Read.All"],
      scoped: true,
      handoff: true,
    });
    expect(combinedRisksAudit.usesStoredResults).toBe(true);
    // Tenant settings reads one stored run as well, but it also has a pack of its own.
    expect(
      builtInAudits
        .filter((audit) => audit.usesStoredResults)
        .map((audit) => audit.definition.id)
        .sort(),
    ).toEqual(["identity.combined-risks", "tenant.settings"]);
  });

  it("has a view for every rule, every severity, and a default that exists", () => {
    const views = definition.views.map((view) => view.id);
    for (const rule of RULES) expect(views).toContain(rule.id);
    for (const severity of ["critical", "high", "medium"]) expect(views).toContain(severity);
    expect(views).toContain(definition.defaultView);
    expect(definition.columns.map((column) => column.key)).toContain(definition.defaultSort.column);
  });

  it("sorts severity labels alphabetically in severity order, which the default sort relies on", async () => {
    const { result } = await run();
    const labels = [...new Set(result.rows.map((item) => item.cells.severity as string))];
    expect([...labels].sort()).toEqual(labels);
    expect(labels[0]).toBe("Critical");
  });
});

describe("the columns it reads", () => {
  it("all exist in the audit that publishes them, so a rename fails here", () => {
    for (const id of INPUT_IDS) {
      const audit = builtInAudits.find((item) => item.definition.id === INPUTS[id].auditId);
      expect(audit, `${id} is a built-in audit`).toBeDefined();
      const keys = new Set(auditDefinition(audit as never).columns.map((column) => column.key));
      for (const column of INPUTS[id].columns)
        expect(keys.has(column), `${id}.${column}`).toBe(true);
    }
  });

  it("finds the hidden owner IDs column in Groups and Teams, which the join reads", () => {
    const groups = builtInAudits.find((item) => item.definition.id === INPUTS.groups.auditId);
    const column = auditDefinition(groups as never).columns.find((item) => item.key === "ownerIds");
    expect(column?.hidden).toBe(true);
    expect(INPUTS.groups.columns).toContain("ownerIds");
  });
});

describe("group ownership", () => {
  const groups = (over: Record<string, string[] | null> = {}): NonNullable<Runs["groups"]> => ({
    rows: [
      row("g-wiki", null, { group: "Wiki owners", ownerIds: over.wiki ?? [ids.ed, ids.cy] }),
      row("g-lab", null, { group: "Lab", ownerIds: over.lab ?? [ids.cy, ids.cy] }),
      row("g-unread", null, { group: "Unread", ownerIds: null }),
    ],
  });
  const withGroups = (runs: Runs, groupRuns = groups()): Runs => ({ ...runs, groups: groupRuns });

  it("joins the owners of each group to the accounts it already knows", () => {
    const stored = fixture();
    const inputs = {} as Parameters<typeof joinInputs>[0];
    // joinInputs takes freshness records; readInputs builds them from the stored runs.
    return readInputs(fakeStored(withGroups(stored)), now).then((read) => {
      Object.assign(inputs, read);
      const joined = joinInputs(inputs, null, new Set(INPUT_IDS));
      expect(joined.people.get(ids.cy)?.ownedGroups).toEqual(["g-wiki", "g-lab"]);
      expect(joined.people.get(ids.ed)?.ownedGroups).toEqual(["g-wiki"]);
      expect(joined.people.get(ids.ada)?.ownedGroups).toEqual([]);
      // An owner no other input mentions gets no record of their own.
      expect(joined.people.size).toBe(
        joinInputs(read, null, new Set(INPUT_IDS.filter((id) => id !== "groups"))).people.size,
      );
    });
  });

  it("does not use ownership from Groups and Teams when that input is not usable", async () => {
    const read = await readInputs(fakeStored(withGroups(fixture())), now);
    const usable = new Set(INPUT_IDS.filter((id) => id !== "groups"));
    expect(joinInputs(read, null, usable).people.get(ids.cy)?.ownedGroups).toEqual([]);
  });

  it("adds ownership to the reasons for a disabled account and a guest", async () => {
    const runs = withGroups(fixture(), groups({ wiki: [ids.cy, ids.bo], lab: [ids.bo] }));
    const { result } = await run({ runs });
    expect(rowFor(result.rows, ids.cy).cells.reasons).toContain("1 license, ownership of 1 group");
    expect(rowFor(result.rows, ids.bo).cells.reasons).toContain("and owns 2 groups.");
  });

  it("flags a guest whose only power is owning a group", async () => {
    const runs = fixture();
    runs.roles = { rows: (runs.roles?.rows ?? []).filter((item) => item.personId !== ids.bo) };
    const without = await run({ runs: structuredClone(runs) });
    expect(ruleLabels(ids.bo, without.result)).not.toContain(label("guest-with-power"));
    const { result } = await run({ runs: withGroups(runs, groups({ wiki: [ids.bo] })) });
    expect(ruleLabels(ids.bo, result)).toContain(label("guest-with-power"));
    expect(rowFor(result.rows, ids.bo).cells.reasons).toContain(
      "Is a guest account that owns 1 group.",
    );
  });
});

describe("the result", () => {
  it("lists one row per account that trips a rule, worst first, and none for the rest", async () => {
    const { result } = await run({ findings: [finding()], granted: ["Policy.Read.All"] });
    expect(result.rows.map((item) => item.key)).toEqual([ids.ada, ids.bo, ids.cy, ids.fay, ids.di]);
    for (const item of result.rows) auditRowSchema.parse(item);
    const columns = new Set(
      auditDefinition(combinedRisksAudit).columns.map((column) => column.key),
    );
    for (const item of result.rows)
      for (const key of Object.keys(item.cells)) expect(columns.has(key)).toBe(true);
    const viewTags = new Set(
      auditDefinition(combinedRisksAudit).views.flatMap((view) => view.tags),
    );
    for (const item of result.rows)
      for (const tag of item.tags) expect(viewTags.has(tag), tag).toBe(true);
  });

  it("names each account's rules, reasons and highest severity", async () => {
    const { result } = await run({ findings: [finding()] });
    const ada = rowFor(result.rows, ids.ada);
    expect(ada.cells).toMatchObject({
      severity: "Critical",
      ruleCount: 5,
      enabled: true,
      privileged: true,
      domain: "a.example",
    });
    expect(ruleLabels(ids.ada, result)).toEqual(
      [
        label("admin-unprotected"),
        label("admin-signin-finding"),
        label("dormant-admin"),
        label("admin-weak-device"),
        label("paying-for-nothing"),
      ].sort(),
    );
    // Five rules, five whole sentences one after another, never joined with commas.
    const reasons = ada.cells.reasons as string;
    expect(typeof reasons).toBe("string");
    expect(reasons).toContain("Holds Global Administrator (permanent) and has no MFA registered. ");
    expect(reasons).not.toMatch(/\.,/);
    expect(reasons.match(/\. [A-Z]/g)).toHaveLength(4);
    expect(rowFor(result.rows, ids.bo).cells).toMatchObject({
      severity: "Critical",
      privileged: false,
    });
    expect(ruleLabels(ids.bo, result)).toEqual(
      [label("risky-unprotected"), label("guest-with-power")].sort(),
    );
    expect(rowFor(result.rows, ids.fay).cells).toMatchObject({ severity: "High", ruleCount: 1 });
    expect(rowFor(result.rows, ids.cy).cells).toMatchObject({ severity: "High", enabled: false });
    expect(ruleLabels(ids.cy, result)).toEqual(
      [label("dormant-admin"), label("disabled-still-holding"), label("paying-for-nothing")].sort(),
    );
    expect(ruleLabels(ids.di, result)).toEqual([label("paying-for-nothing")]);
  });

  it("leaves out a remediated risk, a room mailbox and an account with nothing wrong", async () => {
    const { result } = await run();
    const keys = result.rows.map((item) => item.key);
    expect(keys).not.toContain(ids.ed);
    expect(keys).not.toContain(ids.room);
  });

  it("counts accounts by their highest severity", async () => {
    const { result } = await run({ findings: [finding()] });
    expect(tile(result, COMBINED_RISKS_TILES.critical)).toBe(2);
    expect(tile(result, COMBINED_RISKS_TILES.high)).toBe(2);
    expect(tile(result, "Medium combined risks")).toBe(1);
    expect(tile(result, "Accounts affected")).toBe(5);
    expect(tile(result, "Rules not checked")).toBe(0);
    expect(result.summary[0]).toMatchObject({ tone: "critical", view: "critical" });
    expect(result.summary.map((item) => item.label).slice(0, 2)).toEqual([
      "Critical combined risks",
      "High combined risks",
    ]);
  });

  it("lists the audits and dates behind a combination", async () => {
    const runs = fixture();
    runs.roles = {
      ...(runs.roles as object),
      rows: runs.roles?.rows ?? [],
      finishedAt: "2026-10-20T08:00:00.000Z",
    } as never;
    const { result } = await run({ runs });
    const sources = rowFor(result.rows, ids.ada).cells.sources as string[];
    expect(sources).toContain("Privileged roles (2026-10-20)");
    expect(sources).toContain("MFA registration (2026-10-21)");
    expect(sources.some((item) => item.startsWith("Managed devices"))).toBe(true);
  });

  it("is complete only when every rule could be checked", async () => {
    expect((await run({ findings: [finding()] })).result.complete).toBe(true);
    const runs = fixture();
    delete runs.devices;
    expect((await run({ runs })).result.complete).toBe(false);
  });

  it("gives the inventory of domains it saw", async () => {
    const { result } = await run();
    // Bo's guest sign-in name is on a.example too, so all seven accounts are on one domain.
    expect(result.domains).toEqual([{ domain: "a.example", accounts: 7 }]);
  });
});

describe("freshness", () => {
  const withRoles = (finishedAt: string | undefined, drop = false) => {
    const runs = fixture();
    if (drop) delete runs.roles;
    else if (runs.roles && finishedAt) runs.roles = { ...runs.roles, finishedAt };
    return runs;
  };
  const sevenDays = 7 * 86_400_000;

  it("uses an input that finished exactly seven days ago", async () => {
    const { result } = await run({
      runs: withRoles(new Date(now.getTime() - sevenDays).toISOString()),
    });
    expect(tile(result, "Rules not checked")).toBe(0);
  });

  it("treats an input one millisecond older than seven days as missing", async () => {
    const stale = new Date(now.getTime() - sevenDays - 1).toISOString();
    const { result } = await run({ runs: withRoles(stale) });
    // Every rule that reads Privileged roles is off for everyone.
    expect(tile(result, "Rules not checked")).toBe(6);
    // Ada is still listed for what does not need roles, and for nothing that does.
    expect(ruleLabels(ids.ada, result)).toEqual([label("paying-for-nothing")]);
    expect(result.notes.join("\n")).toContain(
      "Privileged roles last ran 2026-10-15, more than 7 days ago",
    );
  });

  it("treats an input that never ran as missing and says to run it", async () => {
    const { result } = await run({ runs: withRoles(undefined, true) });
    expect(result.notes.join("\n")).toContain("Privileged roles has never run");
    expect(result.notes).toContain(
      `${label("admin-unprotected")} was not checked. Run Privileged roles.`,
    );
    expect(result.notes).toContain(
      `${label("dormant-admin")} was not checked. Run Privileged roles.`,
    );
  });

  it("never combines a fresh input with a missing one: the other half of a rule alone does not trip it", async () => {
    // MFA registration is fresh and says Ada has none, but without Privileged roles she is not known
    // to be an administrator, so nothing about her is flagged as one.
    const { result } = await run({ runs: withRoles(undefined, true) });
    expect(ruleLabels(ids.ada, result)).toEqual([label("paying-for-nothing")]);
    expect(rowFor(result.rows, ids.bo).cells.rules).toEqual([label("risky-unprotected")]);
  });

  it("falls back to Inactive accounts for sign-in dates when Account activity is stale", async () => {
    const runs = fixture();
    runs.activity = { ...(runs.activity as object), rows: [], finishedAt: daysAgo(9) } as never;
    runs.inactive = {
      rows: [
        row(ids.ada, ids.ada, {
          person: names.ada,
          status: "Enabled",
          type: "Member",
          lastSuccess: daysAgo(95),
          days: 95,
          licenses: [],
        }),
      ],
    };
    const { result } = await run({ runs });
    expect(ruleLabels(ids.ada, result)).toContain(label("dormant-admin"));
    // Paying for nothing needs Account activity itself, so it is the one that goes off.
    expect(result.notes).toContain(
      `${label("paying-for-nothing")} was not checked. Run Account activity.`,
    );
  });

  it("reports when nothing has run lately", async () => {
    const { result } = await run({ runs: {} });
    expect(result.rows).toEqual([]);
    expect(tile(result, "Rules not checked")).toBe(9);
    expect(result.notes.join("\n")).toContain("nothing was checked");
  });
});

describe("matching people", () => {
  it("joins by directory ID, not by name", async () => {
    const runs = fixture();
    // The same display name on a different account must not take on Ada's roles.
    runs.mfa?.rows.push(
      row("x", "u-other", { person: names.ada, registered: false, type: "Member" }),
    );
    const { result } = await run({ runs });
    expect(result.rows.map((item) => item.key)).not.toContain("u-other");
  });

  it("ignores a row with no person ID, such as an application holding a role", async () => {
    const { result } = await run();
    expect(result.rows.every((item) => item.personId !== null)).toBe(true);
  });

  it("matches a mailbox by sign-in name, in any letter case", async () => {
    const runs = fixture();
    runs.mailbox = {
      rows: [row("m", null, { mailbox: "DI@A.EXAMPLE", recipientType: "SharedMailbox" })],
    };
    const { result } = await run({ runs });
    expect(result.rows.map((item) => item.key)).not.toContain(ids.di);
  });

  it("never matches a hashed mailbox name, and says so", async () => {
    const runs = fixture();
    runs.mailbox = {
      rows: [
        row("m1", null, {
          mailbox: "0123456789abcdef0123456789abcdef",
          recipientType: "SharedMailbox",
        }),
        row("m2", null, { mailbox: "room@a.example", recipientType: "UserMailbox" }),
      ],
    };
    const { result } = await run({ runs });
    expect(result.notes.join("\n")).toContain("1 mailbox has a hidden name");
    // The room's type now reads as a user mailbox, so it is flagged: nothing guessed a type for it.
    expect(result.rows.map((item) => item.key)).toContain(ids.room);
  });

  it("does not check paying for an unused account when Mailbox storage lists nothing", async () => {
    const runs = fixture();
    runs.mailbox = { rows: [] };
    const { result } = await run({ runs });
    expect(result.rows.map((item) => item.key)).not.toContain(ids.di);
    expect(result.notes.join("\n")).toContain("Mailbox storage listed no mailboxes");
    expect(result.notes).toContain(
      `${label("paying-for-nothing")} was not checked. Run Mailbox storage.`,
    );
  });
});

describe("who counts as a guest", () => {
  // Each input that can say so, alone, makes the account a guest; an account none of them names is not.
  const guestFrom = async (id: InputId, cells: Record<string, unknown>) => {
    const stored = fakeStored({ [id]: { rows: [row("g", "u-g", cells)] } });
    const inputs = await readInputs(stored, now);
    return joinInputs(inputs, [], new Set([id])).people.get("u-g")?.guest;
  };

  it("takes the type from Account activity, Inactive accounts and MFA registration", async () => {
    for (const id of ["activity", "inactive", "mfa"] as const) {
      expect(await guestFrom(id, { person: names.bo, type: "Guest" }), id).toBe(true);
      expect(await guestFrom(id, { person: names.bo, type: "Member" }), id).toBe(false);
    }
  });

  it("takes membership of the Guest accounts audit, and the holder type in Privileged roles", async () => {
    expect(await guestFrom("guests", { person: names.bo, status: "Enabled" })).toBe(true);
    expect(
      await guestFrom("roles", {
        holder: names.bo,
        holderType: "Guest",
        role: "Helpdesk Administrator",
      }),
    ).toBe(true);
    expect(
      await guestFrom("roles", {
        holder: names.bo,
        holderType: "User",
        role: "Helpdesk Administrator",
      }),
    ).toBe(false);
  });
});

describe("Conditional Access", () => {
  it("reads nothing from Microsoft when the pack is not granted, and says MFA registration alone was used", async () => {
    const { requests, result } = await run();
    expect(requests).toEqual([]);
    expect(result.notes.join("\n")).toContain("looked at MFA registration only");
  });

  it("reads the policy list once, with GET, when granted", async () => {
    const { requests } = await run({ granted: ["Policy.Read.All"], policies: [mfaPolicy()] });
    expect(requests.map((url) => url.pathname)).toEqual([POLICY_PATH]);
  });

  it("flags an administrator with MFA who is left out of the policy by a direct group", async () => {
    const runs = fixture();
    runs.mfa = {
      rows: [row(ids.ada, ids.ada, { person: names.ada, registered: true, type: "Member" })],
    };
    const policy = mfaPolicy({
      conditions: { users: { includeUsers: ["All"], excludeGroups: ["G-ADMINS"] } },
    });
    const { result } = await run({ runs, granted: ["Policy.Read.All"], policies: [policy] });
    const ada = rowFor(result.rows, ids.ada);
    expect(ada.cells.reasons as string).toContain(
      "Holds Global Administrator (permanent) and is left out of Require MFA for everyone, which requires MFA.",
    );
    expect(ada.cells.sources as string[]).toContain(
      "Conditional Access policies (read 2026-10-22)",
    );
    expect(result.notes.join("\n")).toContain("1 enabled policy requires MFA");
  });

  it("flags an administrator left out by role", async () => {
    const runs = fixture();
    runs.mfa = {
      rows: [row(ids.ada, ids.ada, { person: names.ada, registered: true, type: "Member" })],
    };
    const policy = mfaPolicy({
      conditions: {
        users: { includeUsers: ["All"], excludeRoles: ["62e90394-69f5-4237-9190-012177145e10"] },
      },
    });
    const { result } = await run({ runs, granted: ["Policy.Read.All"], policies: [policy] });
    expect(ruleLabels(ids.ada, result)).toContain(label("admin-unprotected"));
  });

  it("does not flag a covered administrator, or one left out of a policy in report-only mode", async () => {
    const runs = fixture();
    runs.mfa = {
      rows: [row(ids.ada, ids.ada, { person: names.ada, registered: true, type: "Member" })],
    };
    const reportOnly = mfaPolicy({
      state: "enabledForReportingButNotEnforced",
      conditions: { users: { includeUsers: ["All"], excludeUsers: [ids.ada] } },
    });
    for (const policies of [[mfaPolicy()], [reportOnly]]) {
      const { result } = await run({ runs, granted: ["Policy.Read.All"], policies });
      expect(ruleLabels(ids.ada, result)).not.toContain(label("admin-unprotected"));
    }
  });

  it("keeps going on MFA registration alone when Microsoft refuses the policies", async () => {
    const { result } = await run({
      granted: ["Policy.Read.All"],
      policies: new Response("{}", { status: 403 }),
    });
    expect(result.complete).toBe(false);
    expect(result.notes.join("\n")).toContain("would not return the Conditional Access policies");
    expect(ruleLabels(ids.ada, result)).toContain(label("admin-unprotected"));
  });

  it("is not read at all when no rule that uses it can run", async () => {
    const runs = fixture();
    delete runs.mfa;
    const { requests } = await run({ runs, granted: ["Policy.Read.All"], policies: [mfaPolicy()] });
    expect(requests).toEqual([]);
  });

  it("says when Account activity did not list group IDs", async () => {
    const runs = fixture();
    runs.activity = {
      rows: (runs.activity?.rows ?? []).map((item) => ({
        ...item,
        cells: { ...item.cells, groupIds: null },
      })),
    };
    const { result } = await run({ runs, granted: ["Policy.Read.All"], policies: [mfaPolicy()] });
    expect(result.notes.join("\n")).toContain("did not list group IDs");
  });
});

describe("Tenant settings", () => {
  const defaultsRow = (value: string, status = "Meets") =>
    row(SECURITY_DEFAULTS_KEY, null, { setting: "Security defaults", value, status });
  const settingsRun = (
    value: string,
    over: { status?: string; days?: number; complete?: boolean; scope?: string[] } = {},
  ) => ({
    rows: [defaultsRow(value, over.status)],
    finishedAt: daysAgo(over.days ?? 1),
    ...(over.complete === undefined ? {} : { complete: over.complete }),
    ...(over.scope ? { scope: over.scope } : {}),
  });
  // Ada has MFA registered but is left out of the policy by her group, so only the policy half of
  // the rule can trip for her.
  const leftOut = () => {
    const runs = fixture();
    runs.mfa = {
      rows: [row(ids.ada, ids.ada, { person: names.ada, registered: true, type: "Member" })],
    };
    const policy = mfaPolicy({
      conditions: { users: { includeUsers: ["All"], excludeGroups: ["G-ADMINS"] } },
    });
    return { runs, policy };
  };
  // Ada with MFA registered and no policy aimed at her at all.
  const registered = () => {
    const runs = fixture();
    runs.mfa = {
      rows: [row(ids.ada, ids.ada, { person: names.ada, registered: true, type: "Member" })],
    };
    return runs;
  };
  const trips = (result: Awaited<ReturnType<typeof run>>["result"]) =>
    result.rows.some((item) => item.key === ids.ada) &&
    ruleLabels(ids.ada, result).includes(label("admin-unprotected"));

  it("reads the policies live and flags the admin even when a stored run says security defaults are on", async () => {
    const { runs, policy } = leftOut();
    runs.settings = settingsRun(SECURITY_DEFAULTS_ON);
    const { result, requests } = await run({
      runs,
      granted: ["Policy.Read.All"],
      policies: [policy],
    });
    expect(requests.map((url) => url.pathname)).toEqual([POLICY_PATH]);
    expect(trips(result)).toBe(true);
    expect(result.notes.join("\n")).not.toContain("Security defaults are on");
  });

  it("uses a fresh security defaults row when the pack is not granted, so no one is left out", async () => {
    const runs = registered();
    runs.settings = settingsRun(SECURITY_DEFAULTS_ON);
    const { result, requests } = await run({ runs });
    expect(requests).toEqual([]);
    expect(trips(result)).toBe(false);
    const notes = result.notes.join("\n");
    expect(notes).toContain("Security defaults are on (Tenant settings, 2026-10-21)");
    expect(notes).toContain("so no one is left out of MFA");
    expect(notes).not.toContain("not enabled or granted");
    expect(result.complete).toBe(true);
    expect(notes).toContain("Tenant settings (2026-10-21)");
  });

  it("uses the same row when Microsoft refuses the policies, and says so", async () => {
    const { result } = await run({
      runs: Object.assign(registered(), { settings: settingsRun(SECURITY_DEFAULTS_ON) }),
      granted: ["Policy.Read.All"],
      policies: new Response("{}", { status: 403 }),
    });
    const notes = result.notes.join("\n");
    expect(notes).toContain("Microsoft would not return the Conditional Access policies.");
    expect(notes).toContain("Security defaults are on (Tenant settings");
    expect(result.complete).toBe(true);
  });

  it("does not check the left-out part without a fresh run and without the pack, as before", async () => {
    const none = await run({ runs: registered() });
    expect(none.result.notes.join("\n")).toContain("looked at MFA registration only");
    expect(none.result.notes.join("\n")).not.toContain("Security defaults are on");
    expect(none.result.notes.join("\n")).not.toContain("Tenant settings");
    const stale = registered();
    stale.settings = settingsRun(SECURITY_DEFAULTS_ON, { days: 8 });
    const old = await run({ runs: stale });
    expect(old.result.notes.join("\n")).not.toContain("Security defaults are on");
  });

  it("uses a run exactly seven days old, and ignores a row that is not security defaults on", async () => {
    const edge = registered();
    edge.settings = settingsRun(SECURITY_DEFAULTS_ON, { days: 7 });
    expect((await run({ runs: edge })).result.notes.join("\n")).toContain(
      "Security defaults are on",
    );
    for (const settings of [
      settingsRun("Off, with a Conditional Access policy that requires MFA for all users"),
      settingsRun("Off. No enabled Conditional Access policy requires MFA for all users", {
        status: "Review",
      }),
      settingsRun(SECURITY_DEFAULTS_ON, { status: "Not read" }),
    ]) {
      const runs = registered();
      runs.settings = settings;
      const { result } = await run({ runs });
      expect(result.notes.join("\n")).not.toContain("Security defaults are on");
    }
  });

  it("still flags an administrator with no MFA registered, and says nothing about a policy", async () => {
    const runs = fixture();
    runs.settings = settingsRun(SECURITY_DEFAULTS_ON);
    const { result } = await run({ runs });
    const ada = rowFor(result.rows, ids.ada);
    expect(ada.cells.reasons as string).toContain(
      "Holds Global Administrator (permanent) and has no MFA registered.",
    );
    expect(ada.cells.reasons as string).not.toContain("left out of");
  });

  it("does not call Combined risks partial or limited because Tenant settings was, and lists it only when used", async () => {
    const runs = fixture();
    runs.settings = settingsRun(SECURITY_DEFAULTS_ON, { complete: false, scope: ["a.example"] });
    const { result } = await run({ runs, granted: ["Policy.Read.All"], policies: [mfaPolicy()] });
    const notes = result.notes.join("\n");
    expect(notes).not.toContain("Tenant settings");
    expect(notes).not.toContain("partial result, so an account may be missing");
    expect(notes).not.toContain("ran limited to some domains");
    expect(result.complete).toBe(true);
    const used = await run({
      runs: Object.assign(registered(), {
        settings: settingsRun(SECURITY_DEFAULTS_ON, { complete: false }),
      }),
    });
    expect(used.result.notes.join("\n")).not.toContain("an account may be missing");
  });
});

describe("scope and notes", () => {
  it("covers only the domains in the audit scope and says how many it left out", async () => {
    const runs = fixture();
    runs.mfa?.rows.push(
      row("fb", "u-fb", {
        person: who("Fi Beta", "fi@b.example"),
        registered: false,
        type: "Member",
      }),
    );
    runs.roles?.rows.push(
      row("rb", "u-fb", {
        holder: who("Fi Beta", "fi@b.example"),
        holderType: "User",
        role: "Global Administrator",
        privileged: true,
        held: "Permanent",
        enabled: true,
      }),
    );
    const all = await run({ runs });
    expect(all.result.rows.map((item) => item.key)).toContain("u-fb");
    const scoped = await run({ runs, scope: new Set(["a.example"]) });
    expect(scoped.result.rows.map((item) => item.key)).not.toContain("u-fb");
    expect(scoped.result.notes).toContain("Limited to accounts in: a.example.");
    expect(scoped.result.notes).toContain(
      "1 account in another domain was left out by the audit scope.",
    );
    expect(scoped.result.domains?.map((item) => item.domain)).toEqual(["a.example", "b.example"]);
  });

  it("says that group ownership was not checked while Groups and Teams is missing", async () => {
    const { result } = await run();
    expect(result.notes.join("\n")).toContain("Owning a group was not checked");
  });

  it("says which groups it can see ownership of once Groups and Teams is used", async () => {
    const { result } = await run({
      runs: { ...fixture(), groups: { rows: [row("g1", null, { group: "G", ownerIds: [] })] } },
    });
    const text = result.notes.join("\n");
    expect(text).toContain("Owning a group is checked against Groups and Teams");
    expect(text).not.toContain("Owning a group was not checked");
  });

  it("says when an input was limited to some domains or only partly read", async () => {
    const runs = fixture();
    runs.mfa = {
      ...(runs.mfa as object),
      rows: runs.mfa?.rows ?? [],
      scope: ["a.example"],
      complete: false,
    } as never;
    const { result } = await run({ runs });
    expect(result.notes.join("\n")).toContain("MFA registration ran limited to some domains");
    expect(result.notes.join("\n")).toContain("MFA registration returned a partial result");
    expect(result.complete).toBe(false);
  });

  it("lists the inputs used with the dates they ran", async () => {
    const { result } = await run();
    const used = result.notes.find((note) => note.startsWith("Inputs used"));
    expect(used).toContain("Account activity (2026-10-21)");
    expect(used).toContain("MFA registration (2026-10-21)");
    expect(used).not.toContain("Inactive accounts");
  });

  it("keeps every note within the 500 characters the result stores", async () => {
    const { result } = await run({ runs: { activity: { rows: [] } } });
    expect(result.notes.length).toBeLessThanOrEqual(20);
    for (const note of result.notes) expect(note.length).toBeLessThanOrEqual(500);
  });

  it("refuses to run without the stored results", async () => {
    const audit = createCombinedRisksAudit();
    const network = fake(audit, () => undefined);
    await expect(audit.run(network.context)).rejects.toMatchObject({ code: "no_results" });
  });

  it("asks only for the stored runs of its own inputs", async () => {
    const { stored } = await run();
    expect([...new Set(stored.asked)].sort()).toEqual(
      INPUT_IDS.map((id) => INPUTS[id].auditId).sort(),
    );
  });
});

import {
  type AuditDefinitionInput,
  type AuditRow,
  type AuditSummaryItem,
  COMBINED_RISKS_AUDIT_ID,
  COMBINED_RISKS_TILES,
} from "@prism365/contracts";
import { GraphRequestError } from "@prism365/graph";
import {
  FRESH_DAYS,
  type Freshness,
  INPUT_IDS,
  INPUTS,
  type InputId,
  joinInputs,
  readInputs,
  SIGNINS_MODULE,
} from "./combined-risks-inputs.ts";
import { type MfaPolicy, mfaPolicies } from "./combined-risks-reach.ts";
import { RULES, type Rule } from "./combined-risks-rules.ts";
import type { Audit } from "./definition.ts";
import { AuditError } from "./definition.ts";
import { byName } from "./directory.ts";
import { createPacer, type Pause, readPaced } from "./paced-read.ts";
import {
  domainColumn,
  domainInventory,
  excludedNote,
  inAuditScope,
  principalDomain,
  scopeNote,
} from "./scope.ts";

// Each audit looks at one thing; the accounts that matter most trip several at once. This audit
// joins the latest results of the others per person and flags the dangerous combinations, with
// the audits and run dates behind each. It reads nothing from Microsoft except the Conditional
// Access policies, and only when a rule needs them. It needs no permission of its own.

type Column = AuditDefinitionInput["columns"][number];

const SEVERITY_LABEL = { critical: "Critical", high: "High", medium: "Medium" } as const;
// Alphabetical order is also severity order (Critical, High, Medium), which the default sort on
// the severity column relies on. A test checks it.
const RANK = { critical: 3, high: 2, medium: 1 } as const;

const policyContext = {
  feature: "identity.conditional-access.read",
  permission: "Policy.Read.All",
} as const;
const MAX_POLICIES = 1_000;
const MAX_PAGES = 20;
const MAX_NOTES = 20;

const dateOf = (iso: string) => iso.slice(0, 10);
const plural = (count: number, one: string, many: string) => `${count} ${count === 1 ? one : many}`;
const listOf = (items: readonly string[]) =>
  items.length <= 1
    ? items.join("")
    : `${items.slice(0, -1).join(", ")} and ${items[items.length - 1]}`;

const columns: Column[] = [
  { key: "person", label: "Account", kind: "person" },
  {
    key: "severity",
    label: "Highest severity",
    kind: "enum",
    facet: true,
    help: "The most serious rule the account trips.",
  },
  {
    key: "ruleCount",
    label: "Rules tripped",
    kind: "number",
    hidden: true,
  },
  { key: "rules", label: "Combinations", kind: "list" },
  { key: "reasons", label: "Why", kind: "text" },
  {
    key: "sources",
    label: "Based on",
    kind: "list",
    help: "The audits, and the dates they ran, behind the combinations. Results older than seven days are not used.",
  },
  { key: "enabled", label: "Account enabled", kind: "boolean", empty: "Unknown" },
  {
    key: "privileged",
    label: "Privileged",
    kind: "boolean",
    empty: "Unknown",
    help: "Holds a privileged directory role, active or eligible, according to the Privileged roles audit.",
  },
  domainColumn,
];

const views = [
  { id: "all", label: "All accounts", tags: [] },
  { id: "critical", label: "Critical", tags: ["critical"] },
  { id: "high", label: "High", tags: ["high"] },
  { id: "medium", label: "Medium", tags: ["medium"] },
  { id: "privileged", label: "Privileged accounts", tags: ["privileged"] },
  ...RULES.map((rule) => ({ id: rule.id, label: rule.label, tags: [rule.id] })),
];

// What each input contributed to this run, for the notes and the sources column.
interface Standing {
  freshness: Freshness;
  usable: boolean;
}

function standingNote(id: InputId, standing: Standing): string {
  const { freshness } = standing;
  if (freshness.state === "never") return `${INPUTS[id].title} has never run`;
  if (freshness.state === "stale")
    return `${INPUTS[id].title} last ran ${dateOf(freshness.finishedAt)}, more than ${FRESH_DAYS} days ago`;
  return `${INPUTS[id].title} listed no mailboxes, which happens when Microsoft hides names`;
}

// The groups of a rule that no usable input can meet.
const unmetGroups = (rule: Rule, usable: ReadonlySet<InputId>) =>
  rule.needs.filter((group) => !group.some((id) => usable.has(id)));

type PolicyRead =
  | { state: "read"; policies: MfaPolicy[]; skipped: number; truncated: boolean }
  | { state: "not-granted" }
  | { state: "failed" };

export interface CombinedRisksOptions {
  // How the audit waits between Conditional Access requests. Tests replace it.
  pause?: Pause;
}

export function createCombinedRisksAudit(options: CombinedRisksOptions = {}): Audit {
  return {
    definition: {
      id: COMBINED_RISKS_AUDIT_ID,
      version: 1,
      title: "Combined risks",
      description:
        "Accounts that trip several audits at once: administrators without MFA, at-risk accounts without MFA, dormant administrators, guests with directory roles and more, each with the audits and dates behind it.",
      category: "Identity",
      subject: { one: "account", many: "accounts" },
      columns,
      views,
      defaultView: "all",
      defaultSort: { column: "severity", direction: "asc" },
      requirements: [
        "Run the audits it joins first: Account activity, MFA registration, Privileged roles, Risky users, Guest accounts, Passwords, Managed devices, Mailbox storage and Groups and Teams. A result older than seven days is not used.",
        "Conditional Access policies are read when the Conditional Access read pack is enabled and granted. Without them only MFA registration is checked.",
      ],
      handoff: true,
    },
    scoped: true,
    usesStoredResults: true,
    packs: [],
    optionalPacks: ["identity.conditional-access.read"],
    forbidden: "This audit asks Microsoft for nothing it needs permission for.",
    async run(context) {
      const { stored } = context;
      if (!stored)
        throw new AuditError("no_results", "This audit could not read the stored audit results.");
      context.signal.throwIfAborted();
      const inputs = await readInputs(stored, context.now);
      const findings = await stored.openFindings(SIGNINS_MODULE);

      // An input is usable when it ran lately. Mailbox storage must also have listed mailboxes:
      // with names hidden it lists none, and then a shared mailbox cannot be told from any other.
      const usable = new Set<InputId>();
      const standing = {} as Record<InputId, Standing>;
      for (const id of INPUT_IDS) {
        const freshness = inputs[id];
        const listed = freshness.state === "fresh" && freshness.run.rows.length > 0;
        const ok = freshness.state === "fresh" && (id !== "mailbox" || listed);
        if (ok) usable.add(id);
        standing[id] = { freshness, usable: ok };
      }

      const checked = RULES.filter((rule) => unmetGroups(rule, usable).length === 0);
      const unchecked = RULES.filter((rule) => !checked.includes(rule));

      // The policies are read only when a rule that looks at them can run.
      const wantsPolicies = checked.some(
        (rule) => rule.id === "admin-unprotected" || rule.id === "risky-unprotected",
      );
      let policyRead: PolicyRead = { state: "not-granted" };
      if (wantsPolicies && context.granted.has(policyContext.permission)) {
        try {
          const read = await readPaced(
            context.graph,
            "identity/conditionalAccess/policies",
            policyContext,
            {
              pacer: createPacer(options.pause),
              signal: context.signal,
              maxPages: MAX_PAGES,
              maxRecords: MAX_POLICIES,
            },
          );
          const parsed = mfaPolicies(read.items);
          policyRead = { state: "read", ...parsed, truncated: !read.complete };
        } catch (error) {
          // A refusal or throttling loses the policy half of two rules; a session that ended
          // is not swallowed, the runner explains it.
          if (!(error instanceof GraphRequestError) || error.status === 401) throw error;
          policyRead = { state: "failed" };
        }
      }
      const policies = policyRead.state === "read" ? policyRead.policies : null;

      const joined = joinInputs(inputs, findings, usable);
      const principals = [...joined.people.values()].map((person) => person.principal);

      const rows: AuditRow[] = [];
      const severityOrder = new Map<string, number>();
      let outside = 0;
      for (const person of joined.people.values()) {
        const trips = checked.flatMap((rule) => {
          const trip = rule.check(person, { now: context.now, policies });
          return trip ? [{ rule, trip }] : [];
        });
        if (trips.length === 0) continue;
        if (!inAuditScope(context.scope, person.principal)) {
          outside += 1;
          continue;
        }
        const highest = trips.reduce<Rule["severity"]>(
          (best, { rule }) => (RANK[rule.severity] > RANK[best] ? rule.severity : best),
          "medium",
        );
        const sources: string[] = [];
        const addSource = (label: string) => {
          if (!sources.includes(label)) sources.push(label);
        };
        for (const { rule, trip } of trips) {
          for (const group of rule.needs) {
            const id = group.find((candidate) => usable.has(candidate));
            const freshness = id ? inputs[id] : null;
            if (id && freshness?.state === "fresh")
              addSource(`${INPUTS[id].title} (${dateOf(freshness.finishedAt)})`);
          }
          if (rule.id === "admin-signin-finding")
            addSource(`Sign-in findings (open on ${dateOf(context.now.toISOString())})`);
          if (trip.usedPolicies)
            addSource(`Conditional Access policies (read ${dateOf(context.now.toISOString())})`);
        }
        const privileged = usable.has("roles")
          ? person.roles.some((role) => role.privileged)
          : null;
        severityOrder.set(person.id, RANK[highest]);
        rows.push({
          key: person.id,
          personId: person.id,
          tags: [
            highest,
            ...trips.map(({ rule }) => rule.id),
            ...(privileged ? ["privileged"] : []),
          ],
          cells: {
            person: { name: person.name, principal: person.principal },
            severity: SEVERITY_LABEL[highest],
            ruleCount: trips.length,
            rules: trips.map(({ rule }) => rule.label),
            // Whole sentences, so they read one after another rather than as a comma list.
            reasons: trips
              .map(({ trip }) => trip.reason)
              .join(" ")
              .slice(0, 2_000),
            sources,
            enabled: person.enabled,
            privileged,
            domain: principalDomain(person.principal),
          },
        });
      }
      rows.sort(
        (a, b) =>
          (severityOrder.get(b.key) ?? 0) - (severityOrder.get(a.key) ?? 0) ||
          Number(b.cells.ruleCount) - Number(a.cells.ruleCount) ||
          byName(a, b),
      );

      const count = (tag: string) => rows.filter((row) => row.tags.includes(tag)).length;
      const critical = count("critical");
      const high = count("high");
      const summary: AuditSummaryItem[] = [
        {
          label: COMBINED_RISKS_TILES.critical,
          value: critical,
          tone: critical > 0 ? "critical" : "neutral",
          view: "critical",
        },
        {
          label: COMBINED_RISKS_TILES.high,
          value: high,
          tone: high > 0 ? "warning" : "neutral",
          view: "high",
        },
        { label: "Medium combined risks", value: count("medium"), tone: "neutral", view: "medium" },
        { label: "Accounts affected", value: rows.length, tone: "neutral", view: "all" },
        {
          label: "Rules not checked",
          value: unchecked.length,
          tone: unchecked.length > 0 ? "warning" : "neutral",
          view: null,
        },
      ];

      const date = (id: InputId) => {
        const freshness = inputs[id];
        return freshness.state === "fresh" ? dateOf(freshness.finishedAt) : "";
      };
      const usedInputs = INPUT_IDS.filter((id) => usable.has(id)).map(
        (id) => `${INPUTS[id].title} (${date(id)})`,
      );
      const notes: string[] = [];
      const limited = scopeNote(context.scope);
      if (limited) notes.push(limited);
      const left = excludedNote(outside);
      if (left) notes.push(left);
      notes.push(
        usedInputs.length > 0
          ? `Inputs used, with the date each audit ran: ${listOf(usedInputs)}. Sign-in findings: ${plural(findings.length, "open finding", "open findings")}.`
          : `No input audit has run in the last ${FRESH_DAYS} days, so nothing was checked.`,
      );
      const gaps = INPUT_IDS.filter(
        (id) =>
          !usable.has(id) &&
          unchecked.some((rule) => unmetGroups(rule, usable).some((group) => group.includes(id))),
      );
      if (gaps.length > 0)
        notes.push(
          `Not used: ${listOf(gaps.map((id) => standingNote(id, standing[id])))}. Two audits are never joined across such a gap.`,
        );
      for (const rule of unchecked) {
        const asks = unmetGroups(rule, usable).map((group) =>
          group.map((id) => INPUTS[id].title).join(" or "),
        );
        notes.push(`${rule.label} was not checked. Run ${listOf(asks)}.`);
      }
      if (policyRead.state === "read") {
        notes.push(
          `Conditional Access: ${plural(policyRead.policies.length, "enabled policy requires", "enabled policies require")} MFA${policyRead.skipped > 0 ? `, and ${plural(policyRead.skipped, "policy was", "policies were")} unreadable` : ""}${policyRead.truncated ? ", and the list was cut short" : ""}. Only direct group memberships are matched and nested groups are not followed, roles are matched for the privileged built-in roles, and apps, locations and risk levels are not weighed.`,
        );
        const noGroups =
          usable.has("activity") &&
          [...joined.people.values()].every((person) => person.groupIds === null);
        if (noGroups && joined.people.size > 0)
          notes.push(
            "Account activity did not list group IDs, so policies aimed at groups were not matched. Run Account activity again.",
          );
      } else if (wantsPolicies)
        notes.push(
          policyRead.state === "failed"
            ? "Microsoft would not return the Conditional Access policies, so the rules about MFA looked at MFA registration only."
            : "The Conditional Access read pack is not enabled or granted for this session, so the rules about MFA looked at MFA registration only.",
        );
      notes.push(
        usable.has("groups")
          ? "Owning a group is checked against Groups and Teams, which reads owners for Microsoft 365 groups, teams and groups that can be given a directory role. Ownership of any other group is not seen."
          : `Owning a group was not checked, because Groups and Teams has not run in the last ${FRESH_DAYS} days. Guest with directory power then looks at roles only, and Disabled but still holding access at roles and licenses.`,
      );
      if (joined.hashedMailboxes > 0)
        notes.push(
          `${plural(joined.hashedMailboxes, "mailbox has", "mailboxes have")} a hidden name and was not matched to anyone.`,
        );
      if (joined.limited.length > 0)
        notes.push(
          `${listOf(joined.limited.map((id) => INPUTS[id].title))} ran limited to some domains, so accounts in other domains are not checked against ${joined.limited.length === 1 ? "it" : "them"}.`,
        );
      if (joined.partial.length > 0)
        notes.push(
          `${listOf(joined.partial.map((id) => INPUTS[id].title))} returned a partial result, so an account may be missing from it.`,
        );
      if (notes.length > MAX_NOTES)
        notes.splice(
          MAX_NOTES - 1,
          notes.length - MAX_NOTES + 1,
          `${notes.length - MAX_NOTES + 1} more notes were left out.`,
        );

      return {
        rows,
        summary,
        notes,
        complete:
          unchecked.length === 0 &&
          joined.partial.length === 0 &&
          policyRead.state !== "failed" &&
          !(policyRead.state === "read" && policyRead.truncated),
        domains: domainInventory(principals),
      };
    },
  };
}

export const combinedRisksAudit: Audit = createCombinedRisksAudit();

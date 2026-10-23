import {
  type AuditRow,
  type AuditScope,
  type AuditSettings,
  auditDomainSchema,
  auditScopeSchema,
  auditSettingsSchema,
  MAX_AUDIT_SETTINGS_DOMAINS,
} from "@prism365/contracts";
import type { SqlClient } from "@prism365/storage";
import type { Audit } from "./definition.ts";
import { latestSucceededRun } from "./store.ts";

// The audit scope: which account domains the account audits cover. It is one setting per tenant
// in prism365_preference, read when a run starts and applied by the audits that declare `scoped`.

const PREFERENCE_KEY = "audits.scope";
// Run notes are cut to 500 characters, so the scope note stays under that.
const NOTE_LIMIT = 480;

// The Domain column every scoped audit has: a filter for one domain that leaves the saved scope
// alone. Hidden by default because the account cell already shows the full principal name.
export const domainColumn = {
  key: "domain",
  label: "Domain",
  kind: "enum",
  facet: true,
  hidden: true,
  help: "The part of the sign-in name after the @. Filter by it to look at one domain without changing the audit scope.",
} as const;

const noScope: AuditScope = { domains: [] };

function parsed(raw: unknown): AuditScope {
  try {
    const value = typeof raw === "string" ? (JSON.parse(raw) as unknown) : raw;
    const scope = auditScopeSchema.safeParse(value);
    return scope.success ? scope.data : noScope;
  } catch {
    return noScope;
  }
}

// The saved scope. A missing or damaged value reads as every domain.
export async function readAuditScope(sql: SqlClient, tenantId: string): Promise<AuditScope> {
  const result = await sql.query<{ value: unknown }>(
    "SELECT value FROM prism365_preference WHERE tenant_id=$1 AND key=$2",
    [tenantId, PREFERENCE_KEY],
  );
  const stored = result.rows[0];
  return stored ? parsed(stored.value) : noScope;
}

export async function writeAuditScope(
  sql: SqlClient,
  tenantId: string,
  scope: AuditScope,
  actorId: string,
): Promise<AuditScope> {
  const value = auditScopeSchema.parse(scope);
  await sql.query(
    `INSERT INTO prism365_preference (tenant_id, key, value, updated_at, updated_by)
     VALUES ($1,$2,$3::jsonb,now(),$4)
     ON CONFLICT (tenant_id, key) DO UPDATE SET value=EXCLUDED.value, updated_at=EXCLUDED.updated_at,
       updated_by=EXCLUDED.updated_by`,
    [tenantId, PREFERENCE_KEY, JSON.stringify(value), actorId],
  );
  return value;
}

// The domain of a principal name: the text after the last @, in lower case. null for a missing
// name or one with nothing after the @.
export function principalDomain(principal: string | null | undefined): string | null {
  if (!principal) return null;
  const at = principal.lastIndexOf("@");
  if (at < 0) return null;
  const domain = principal
    .slice(at + 1)
    .trim()
    .toLowerCase();
  return domain || null;
}

// Whether an account is covered. With no scope (or an empty one) every account is, including one
// with no principal name; with a scope, an account needs a principal name in one of its domains.
export function inAuditScope(
  scope: ReadonlySet<string> | null,
  principal: string | null | undefined,
): boolean {
  if (!scope || scope.size === 0) return true;
  const domain = principalDomain(principal);
  return domain !== null && scope.has(domain);
}

// The note a scoped audit adds to its result: "Limited to accounts in: a.example, b.example."
// null when every domain is covered. A long list is cut with a count so one note holds it all.
export function scopeNote(scope: ReadonlySet<string> | null): string | null {
  if (!scope || scope.size === 0) return null;
  const domains = [...scope].sort();
  const lead = "Limited to accounts in: ";
  const all = `${lead}${domains.join(", ")}.`;
  if (all.length <= NOTE_LIMIT) return all;
  for (let shown = domains.length - 1; shown >= 1; shown -= 1) {
    const note = `${lead}${domains.slice(0, shown).join(", ")}, and ${domains.length - shown} more.`;
    if (note.length <= NOTE_LIMIT) return note;
  }
  return `${lead}${domains.length} domains.`;
}

// Every domain among these principal names, with its number of accounts, before any scope is
// applied. Stored with each run, so the chooser lists domains the current scope leaves out.
export function domainInventory(
  principals: readonly (string | null | undefined)[],
): AuditSettings["domains"] {
  const counts = new Map<string, number>();
  for (const principal of principals) {
    const domain = auditDomainSchema.safeParse(principalDomain(principal) ?? "");
    if (domain.success) counts.set(domain.data, (counts.get(domain.data) ?? 0) + 1);
  }
  return [...counts.entries()]
    .map(([domain, accounts]) => ({ domain, accounts }))
    .sort((a, b) => b.accounts - a.accounts || a.domain.localeCompare(b.domain))
    .slice(0, MAX_AUDIT_SETTINGS_DOMAINS)
    .sort((a, b) => a.domain.localeCompare(b.domain));
}

// Said in every scoped result that left accounts out, so a limited view is never mistaken for
// the whole tenant.
export function excludedNote(excluded: number): string | null {
  if (excluded <= 0) return null;
  return excluded === 1
    ? "1 account in another domain was left out by the audit scope."
    : `${excluded} accounts in other domains were left out by the audit scope.`;
}

// Domains with the number of distinct accounts in the given results, by their Domain column.
// The same account in several audits counts once.
export function domainCounts(
  runs: readonly { rows: readonly AuditRow[] }[],
): AuditSettings["domains"] {
  const accounts = new Map<string, Set<string>>();
  for (const run of runs) {
    for (const row of run.rows) {
      const cell = row.cells.domain;
      if (typeof cell !== "string") continue;
      const domain = auditDomainSchema.safeParse(cell);
      if (!domain.success) continue;
      const keys = accounts.get(domain.data) ?? new Set<string>();
      keys.add(row.key);
      accounts.set(domain.data, keys);
    }
  }
  return [...accounts.entries()]
    .map(([domain, keys]) => ({ domain, accounts: keys.size }))
    .sort((a, b) => b.accounts - a.accounts || a.domain.localeCompare(b.domain))
    .slice(0, MAX_AUDIT_SETTINGS_DOMAINS)
    .sort((a, b) => a.domain.localeCompare(b.domain));
}

// The saved scope and the domains to choose from. The domains come from the newest successful
// run of each audit that uses the scope, so no call to Microsoft is made.
export async function readAuditSettings(
  sql: SqlClient,
  tenantId: string,
  audits: readonly Audit[],
): Promise<AuditSettings> {
  const runs = [];
  for (const audit of audits) {
    if (!audit.scoped) continue;
    const run = await latestSucceededRun(sql, tenantId, audit.definition.id);
    if (run) runs.push(run);
  }
  // Each run's inventory covers every domain it read, whatever its scope. Runs made before the
  // inventory existed fall back to the domains of their rows.
  const counts = new Map<string, number>();
  for (const run of runs) {
    const seen = run.domains.length > 0 ? run.domains : domainCounts([run]);
    for (const { domain, accounts } of seen)
      counts.set(domain, Math.max(counts.get(domain) ?? 0, accounts));
  }
  const domains = [...counts.entries()]
    .map(([domain, accounts]) => ({ domain, accounts }))
    .filter((item) => auditDomainSchema.safeParse(item.domain).success)
    .sort((a, b) => b.accounts - a.accounts || a.domain.localeCompare(b.domain))
    .slice(0, MAX_AUDIT_SETTINGS_DOMAINS)
    .sort((a, b) => a.domain.localeCompare(b.domain));
  return auditSettingsSchema.parse({ scope: await readAuditScope(sql, tenantId), domains });
}

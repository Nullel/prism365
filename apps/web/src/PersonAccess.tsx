import {
  type AccessAuditResult,
  type AccessMembership,
  type AccessMembershipKind,
  type AccessPolicy,
  type CombinedRisks,
  COMBINED_RISKS_AUDIT_ID,
  isStaleExport,
  MAILBOX_ACCESS_STALE_DAYS,
  type MailboxHolder,
  type MailboxRight,
  type PersonAccess as PersonAccessData,
  type PersonMailboxAccess,
  type PersonMailboxGrant,
  type PolicyReason,
  type PrismSession,
  personAccessSchema,
  personMailboxAccessSchema,
  type TrusteeKind,
} from "@prism365/contracts";
import { type ReactNode, useEffect, useState } from "react";
import { day, number, RelativeTime, readJson, SeverityDot, SkeletonRows } from "./common.tsx";
import {
  lineKey,
  partKey,
  type SummaryLine,
  type SummaryPart,
  summaryLines,
} from "./personAccessSummary.ts";

// The Access tab on a person's page: everything Prism365 can see about this account's access, each
// section with where it came from and when, and a plain line when it cannot be shown. The groups,
// roles, ownership and Conditional Access are read live when the tab first opens. The rest comes
// from the audits' stored results, so it is only as new as the audit's last run.

type NotGranted = Extract<PersonAccessData["memberships"], { status: "not_granted" }>;

const accessLink = <a href="#/access">Review Microsoft access</a>;

// What a section that was refused needs, in words, with the way to get it.
export function CannotShow({
  section,
}: {
  section: NotGranted | { status: "failed"; message: string };
}) {
  if (section.status === "failed") return <p className="empty-data">{section.message}</p>;
  return (
    <p className="empty-data">
      {section.enabled ? (
        <>
          This needs the {section.permission} permission, which the signed-in administrator has not
          granted. {accessLink}.
        </>
      ) : (
        <>
          This needs the {section.pack} read pack, which this installation has not enabled. Add it
          to ENABLED_READ_FEATURES and restart. {accessLink}.
        </>
      )}
    </p>
  );
}

const kindHeadings: Record<AccessMembershipKind, string> = {
  microsoft365: "Microsoft 365 groups",
  security: "Security groups",
  "mail-security": "Mail-enabled security groups",
  distribution: "Distribution lists",
  dynamic: "Dynamic groups",
  role: "Directory roles held now",
  "administrative-unit": "Administrative units",
  hidden: "Hidden by Microsoft",
};

function MembershipItem({ entry }: { entry: AccessMembership }) {
  if (entry.kind === "hidden" || entry.name === null)
    return (
      <li>
        <span className="person-muted">Hidden by Microsoft</span> <code>{entry.id}</code>
      </li>
    );
  return (
    <li>
      <strong>{entry.name}</strong>
      {entry.kind === "role" || entry.kind === "administrative-unit" ? null : (
        // Microsoft lists the groups an account is in, nested or not, but not the path to a nested
        // one, so it can only be called nested.
        <span className="person-muted"> {entry.relation === "direct" ? "Direct" : "Nested"}</span>
      )}
      {entry.dynamic && entry.kind === "microsoft365" ? (
        <span className="person-muted"> Dynamic</span>
      ) : null}
      {entry.hasGroupMailbox ? <span className="person-access__tag">Group mailbox</span> : null}
    </li>
  );
}

function Source({ children }: { children: ReactNode }) {
  return <p className="signin-coverage person-access__source">{children}</p>;
}

function PartView({ part }: { part: SummaryPart }) {
  const body = part.date ? <RelativeTime value={part.date} /> : part.text;
  const words = (
    <>
      {part.prefix ? `${part.prefix} ` : null}
      {body}
    </>
  );
  if (part.tone === "bad") return <SeverityDot severity="high" label={part.text ?? ""} />;
  if (part.tone === "warning") return <SeverityDot severity="medium" label={part.text ?? ""} />;
  return part.tone === "good" ? <span className="person-access__good">{words}</span> : words;
}

function LineView({ line }: { line: SummaryLine }) {
  return (
    <li>
      {line.title ? <strong>{line.title}</strong> : null}
      {line.privileged ? (
        <>
          {" "}
          <SeverityDot severity="high" label="Privileged" />
        </>
      ) : null}
      {line.parts.map((part, index) => (
        <span className="person-access__part" key={partKey(part)}>
          {index > 0 ? ", " : line.title ? ": " : null}
          <PartView part={part} />
        </span>
      ))}
    </li>
  );
}

const auditLink = (result: AccessAuditResult) => (
  <a href={`#/audits/${encodeURIComponent(result.auditId)}`}>{result.name}</a>
);

// The source line of a stored result: which audit, how the row was found and how old it is.
function AuditSource({ result }: { result: Extract<AccessAuditResult, { status: "ok" }> }) {
  return (
    <Source>
      From the {auditLink(result)} audit
      {result.matchedBy === "sign-in name" ? ", matched by sign-in name" : ""},{" "}
      <RelativeTime value={result.finishedAt} />
      {result.scope ? `, limited to ${result.scope.join(", ")}` : ""}
    </Source>
  );
}

// The line that stands for a result with nothing to show. Risk says the person was not flagged,
// since that is good news and not a gap.
function EmptyAudit({ result }: { result: Extract<AccessAuditResult, { status: "ok" }> }) {
  if (result.auditId === "identity.risky-users")
    return (
      <p className="empty-data">
        Not flagged in the latest {auditLink(result)} run,{" "}
        <RelativeTime value={result.finishedAt} />.
      </p>
    );
  return (
    <p className="empty-data">Nothing about this person in the latest {auditLink(result)} run.</p>
  );
}

// One audit's rows for this person as a few short lines, from the fields that matter
// (personAccessSummary.ts). A result that cannot be shown says what to do about it.
export function AuditSummary({
  result,
  results,
}: {
  result: AccessAuditResult;
  results: ReadonlyMap<string, AccessAuditResult>;
}) {
  if (result.status === "no_run")
    return (
      <p className="empty-data">
        <a href={`#/audits/${encodeURIComponent(result.auditId)}`}>Run the {result.name} audit</a>{" "}
        to see this here.
      </p>
    );
  if (result.status === "failed")
    return (
      <div className="person-access__audit">
        <Source>From the {auditLink(result)} audit</Source>
        <p className="empty-data">{result.message}</p>
      </div>
    );
  const lines = summaryLines(results, result.auditId);
  return (
    <div className="person-access__audit">
      <AuditSource result={result} />
      {lines.length === 0 ? (
        <EmptyAudit result={result} />
      ) : (
        <ul className="person-access__list">
          {lines.map((line) => (
            <LineView key={lineKey(line)} line={line} />
          ))}
        </ul>
      )}
      {result.truncated ? (
        <p className="signin-coverage">Showing the first {result.rows.length} rows.</p>
      ) : null}
    </div>
  );
}

const combinedAuditHref = `#/audits/${encodeURIComponent(COMBINED_RISKS_AUDIT_ID)}`;

// What the latest Combined risks run said about this account, first on the tab because it is the
// one answer that already weighs several audits together. Stored, so it is as new as the run.
export function CombinedRisksSection({ section }: { section: CombinedRisks | undefined }) {
  // An answer recorded before the section existed has nothing to show.
  if (!section) return null;
  return (
    <section className="panel person-access" aria-label="Combined risks">
      <h3>Combined risks</h3>
      {section.status === "failed" ? (
        <p className="empty-data">{section.message}</p>
      ) : section.status === "no_run" ? (
        <p className="empty-data">
          Combined risks has not run yet. <a href="#/audits">Run a full review</a> to see them here.
        </p>
      ) : (
        <>
          <Source>
            From the <a href={combinedAuditHref}>Combined risks</a> audit,{" "}
            <RelativeTime value={section.finishedAt} />
          </Source>
          {section.severity === null ? (
            <p className="empty-data">No combined risks in the latest run.</p>
          ) : (
            <>
              <p>
                <SeverityDot severity={section.severity} />
              </p>
              <ul className="person-access__list">
                {section.rules.map((rule) => (
                  <li key={rule.id}>
                    <strong>{rule.label}</strong> <SeverityDot severity={rule.severity} />
                  </li>
                ))}
              </ul>
              {section.reasons ? <p>{section.reasons}</p> : null}
            </>
          )}
          {section.gaps.length > 0 || !section.complete ? (
            <>
              <p className="person-muted">
                That run could not check everything{section.gaps.length > 0 ? ":" : "."}
              </p>
              <ul className="person-access__list person-muted">
                {section.gaps.map((gap) => (
                  <li key={gap}>{gap}</li>
                ))}
              </ul>
            </>
          ) : null}
        </>
      )}
    </section>
  );
}

const ownedLabels = {
  group: "Group",
  team: "Team",
  application: "Application",
  "service-principal": "Service principal",
} as const;

const stateLabels: Record<AccessPolicy["state"], string> = {
  on: "On",
  "report-only": "Report only",
  off: "Off",
  unknown: "State not recognised",
};

function reasonText(reason: PolicyReason): string {
  switch (reason.kind) {
    case "all":
      return "all users";
    case "person":
      return "this person by name";
    case "guests":
      return "guests or external users";
    case "group":
      return reason.name ? `the group ${reason.name}` : "a group they are in";
    case "role":
      return reason.name ? `the role ${reason.name}` : "a role they hold";
  }
}
const reasonList = (reasons: PolicyReason[]) => reasons.map(reasonText).join(", ");

// How an exclusion reads: by name, or through the group or role that carries it.
function exclusionText(reason: PolicyReason): string {
  switch (reason.kind) {
    case "person":
      return "by name";
    case "guests":
      return "as a guest or external user";
    case "group":
      return reason.name ? `through the group ${reason.name}` : "through a group they are in";
    case "role":
      return reason.name ? `through the role ${reason.name}` : "through a role they hold";
    case "all":
      return "for all users";
  }
}

function PolicyItem({ policy }: { policy: AccessPolicy }) {
  return (
    <li>
      <strong>{policy.name ?? "Unnamed policy"}</strong>
      <span className="person-muted">
        {" "}
        {stateLabels[policy.state]}. {policy.controls}.
      </span>
      <div className="person-access__why">
        {policy.effect === "excluded" ? (
          <>
            <SeverityDot severity="medium" label="Excluded" /> Left out{" "}
            {policy.excludedBy.map(exclusionText).join(" and ")}.
            {policy.targeted
              ? ` It would otherwise reach them through ${reasonList(policy.includedBy)}.`
              : " It would not have reached them anyway."}
          </>
        ) : (
          <>Applies through {reasonList(policy.includedBy)}.</>
        )}
      </div>
    </li>
  );
}

function ConditionalAccessSection({ section }: { section: PersonAccessData["conditionalAccess"] }) {
  if (section.status !== "ok")
    return (
      <section className="panel person-access" aria-label="Conditional Access">
        <h3>Conditional Access</h3>
        <Source>Live from Microsoft</Source>
        <CannotShow section={section} />
      </section>
    );
  const excluded = section.policies.filter((policy) => policy.effect === "excluded");
  const applies = section.policies.filter((policy) => policy.effect === "applies");
  // An exclusion from a policy that is enforced is the risky case, so it comes first.
  const exclusionsFirst = excluded.some((policy) => policy.state === "on");
  const groups = [
    { id: "applies", title: "Policies that apply", items: applies },
    { id: "excluded", title: "Policies that leave them out", items: excluded },
  ];
  if (exclusionsFirst) groups.reverse();
  return (
    <section className="panel person-access" aria-label="Conditional Access">
      <h3>Conditional Access</h3>
      <Source>
        Live from Microsoft, <RelativeTime value={section.readAt} />
      </Source>
      {section.policies.length === 0 ? (
        <p className="empty-data">No Conditional Access policy names or reaches this person.</p>
      ) : (
        groups
          .filter((group) => group.items.length > 0)
          .map((group) => (
            <div key={group.id}>
              <h4>{group.title}</h4>
              <ul className="person-access__list">
                {group.items.map((policy) => (
                  <PolicyItem key={policy.id} policy={policy} />
                ))}
              </ul>
            </div>
          ))
      )}
      <p className="signin-coverage">
        Conditions such as apps, platforms and locations still decide whether a policy applies at a
        given sign-in.
      </p>
      {section.unaffected > 0 ? (
        <p className="signin-coverage">
          {number.format(section.unaffected)}{" "}
          {section.unaffected === 1 ? "other policy does" : "other policies do"} not reach this
          person.
        </p>
      ) : null}
      {section.membershipsKnown ? null : (
        <p className="signin-coverage">
          Group and role memberships could not be read, so only policies that name everyone, this
          person or guests are matched.
        </p>
      )}
      {section.unreadable > 0 ? (
        <p className="signin-coverage">
          {number.format(section.unreadable)} {section.unreadable === 1 ? "policy" : "policies"}{" "}
          could not be read and {section.unreadable === 1 ? "is" : "are"} not counted.
        </p>
      ) : null}
      {section.truncated ? (
        <p className="signin-coverage">The tenant has more policies than were read.</p>
      ) : null}
    </section>
  );
}

function GroupsSection({
  memberships,
  results,
}: {
  memberships: PersonAccessData["memberships"];
  results: ReadonlyMap<string, AccessAuditResult>;
}) {
  const roles = results.get("directory.privileged-roles") ?? null;
  return (
    <section className="panel person-access" aria-label="Groups and roles">
      <h3>Groups and roles</h3>
      <Source>
        Live from Microsoft
        {memberships.status === "ok" ? (
          <>
            , <RelativeTime value={memberships.readAt} />
          </>
        ) : null}
      </Source>
      {memberships.status !== "ok" ? (
        <CannotShow section={memberships} />
      ) : memberships.entries.length === 0 ? (
        <p className="empty-data">This account is not in any group and holds no directory role.</p>
      ) : (
        (Object.keys(kindHeadings) as AccessMembershipKind[]).flatMap((kind) => {
          const entries = memberships.entries.filter((entry) => entry.kind === kind);
          return entries.length === 0
            ? []
            : [
                <div key={kind}>
                  <h4>
                    {kindHeadings[kind]} ({number.format(entries.length)})
                  </h4>
                  <ul className="person-access__list">
                    {entries.map((entry) => (
                      <MembershipItem key={entry.id} entry={entry} />
                    ))}
                  </ul>
                </div>,
              ];
        })
      )}
      {memberships.status === "ok" && memberships.truncated ? (
        <p className="signin-coverage">
          Showing the first {number.format(memberships.entries.length)} memberships.
        </p>
      ) : null}
      <h4>Eligible and time-bound roles</h4>
      {roles ? <AuditSummary result={roles} results={results} /> : null}
    </section>
  );
}

function OwnsSection({ owned }: { owned: PersonAccessData["owned"] }) {
  return (
    <section className="panel person-access" aria-label="Owns">
      <h3>Owns</h3>
      <Source>
        Live from Microsoft
        {owned.status === "ok" ? (
          <>
            , <RelativeTime value={owned.readAt} />
          </>
        ) : null}
      </Source>
      {owned.status !== "ok" ? (
        <CannotShow section={owned} />
      ) : owned.entries.length === 0 ? (
        <p className="empty-data">This account owns no groups, teams or applications.</p>
      ) : (
        <>
          <ul className="person-access__list">
            {owned.entries.map((entry) => (
              <li key={entry.id}>
                <strong>{entry.name ?? "Hidden by Microsoft"}</strong>
                <span className="person-muted"> {ownedLabels[entry.kind]}</span>
              </li>
            ))}
          </ul>
          <p className="signin-coverage">
            Hand ownership over before this account is removed, so nothing it owns is left without
            an owner.
          </p>
          {owned.truncated ? (
            <p className="signin-coverage">
              Showing the first {number.format(owned.entries.length)} objects.
            </p>
          ) : null}
        </>
      )}
    </section>
  );
}

// The audits each stored section draws on, in the order the tab shows them.
const auditSections: { title: string; audits: string[] }[] = [
  { title: "Sign-in methods", audits: ["mfa.registration"] },
  { title: "Devices", audits: ["devices.managed"] },
  { title: "Mailbox", audits: ["mail.mailbox-storage"] },
  { title: "Password", audits: ["accounts.passwords"] },
  { title: "Risk", audits: ["identity.risky-users"] },
  // Inactive accounts and Licenses to reclaim are folded into this one (personAccessSummary.ts).
  { title: "Activity", audits: ["accounts.activity"] },
];

export const rightLabels: Record<MailboxRight, string> = {
  FullAccess: "Full Access",
  SendAs: "Send As",
  SendOnBehalf: "Send on Behalf",
};
const trusteeLabels: Record<TrusteeKind, string | null> = {
  user: null,
  group: "Group",
  other: null,
  unresolved: "Not resolved",
};

const groupMailboxNote = (
  <p className="signin-coverage">
    The Microsoft 365 group mailboxes this account can use are listed under Groups and roles.
  </p>
);

// "Full Access directly", "Send As through All plant staff".
function grantText(grant: PersonMailboxGrant): string {
  const how = grant.via.kind === "direct" ? "directly" : `through ${grant.via.name}`;
  return `${rightLabels[grant.right]} ${how}`;
}

// The person's grants, one line per mailbox, in the order the server sorted them.
function byMailbox(grants: readonly PersonMailboxGrant[]) {
  const mailboxes = new Map<string, { mailbox: PersonMailboxGrant["mailbox"]; rights: string[] }>();
  for (const grant of grants) {
    const entry = mailboxes.get(grant.mailbox.id) ?? { mailbox: grant.mailbox, rights: [] };
    entry.rights.push(grantText(grant));
    mailboxes.set(grant.mailbox.id, entry);
  }
  return [...mailboxes.values()];
}

function HolderItem({ holder }: { holder: MailboxHolder }) {
  const tag = trusteeLabels[holder.trustee.kind];
  return (
    <li>
      <strong>{holder.trustee.name}</strong>
      {holder.trustee.address ? (
        <span className="person-muted"> {holder.trustee.address}</span>
      ) : null}
      {tag ? <span className="person-access__tag">{tag}</span> : null}
      <span className="person-muted"> {rightLabels[holder.right]}</span>
    </li>
  );
}

// Why grants to groups may be missing: the groups were not, or not all, read.
function GroupsNote({ groups }: { groups: PersonMailboxAccess["groups"] }) {
  if (groups.status === "ok")
    return groups.truncated ? (
      <p className="signin-coverage">
        Only the first {number.format(groups.count)} groups this account is in were checked, so a
        right granted to another group may be missing.
      </p>
    ) : null;
  if (groups.status === "not_read") return null;
  return (
    <>
      <p className="signin-coverage">
        Only rights granted to this account directly are shown: the groups it is in could not be
        read, so rights granted to a group are not matched.
      </p>
      <CannotShow section={groups} />
    </>
  );
}

// The Shared mailboxes section, from the imported Exchange export (ADR-0012): what the person can
// use, how, and for a mailbox, who can use it. Every line is the export's claim, so the source and
// its age come first.
export function SharedMailboxesSection({
  mailbox,
  error,
}: {
  // null while it loads.
  mailbox: PersonMailboxAccess | null;
  error: string;
}) {
  const body = (() => {
    if (error) return <p className="empty-data">{error}</p>;
    if (!mailbox) return <SkeletonRows rows={2} label="Reading shared mailbox access" />;
    const stored = mailbox.import;
    if (!stored)
      return (
        <>
          <p className="empty-data">
            Microsoft's read-only APIs do not show Full Access, Send As or Send on Behalf, and no
            mailbox access export is imported. <a href="#/access">Import one on Microsoft access</a>{" "}
            to see them here.
          </p>
          {groupMailboxNote}
        </>
      );
    const stale = isStaleExport(stored.exportedAt);
    const mailboxes = byMailbox(mailbox.grants);
    return (
      <>
        <Source>
          {stale ? (
            <>
              <SeverityDot
                severity="medium"
                label={`Older than ${MAILBOX_ACCESS_STALE_DAYS} days`}
              />{" "}
            </>
          ) : null}
          From the export of {day(stored.exportedAt)}, imported{" "}
          <RelativeTime value={stored.importedAt} />
        </Source>
        {mailboxes.length === 0 ? (
          <p className="empty-data">
            The export gives this account no Full Access, Send As or Send on Behalf on any{" "}
            {stored.includeUserMailboxes ? "mailbox" : "shared mailbox"}
            {mailbox.groups.status === "ok" ? "" : " directly"}.
          </p>
        ) : (
          <ul className="person-access__list">
            {mailboxes.map(({ mailbox: shared, rights }) => (
              <li key={shared.id}>
                <strong>{shared.name || shared.address}</strong>
                {shared.name ? <span className="person-muted"> {shared.address}</span> : null}
                {shared.type === "UserMailbox" ? (
                  <span className="person-access__tag">User mailbox</span>
                ) : null}
                <div className="person-access__why">{rights.join(", ")}</div>
              </li>
            ))}
          </ul>
        )}
        {mailbox.grantsTruncated ? (
          <p className="signin-coverage">
            Showing the first {number.format(mailbox.grants.length)} rights.
          </p>
        ) : null}
        <GroupsNote groups={mailbox.groups} />
        {stored.includeUserMailboxes ? null : (
          // A shared-only export says nothing about user mailboxes, this account's own included.
          <p className="signin-coverage">
            The export covers shared mailboxes only, so this account's own mailbox and other user
            mailboxes are not in it. Run the script with -IncludeUserMailboxes to include them.
          </p>
        )}
        {mailbox.holders.length > 0 ? (
          <>
            <h4>Who can use this mailbox</h4>
            <ul className="person-access__list">
              {mailbox.holders.map((holder) => (
                <HolderItem
                  key={`${holder.right}:${holder.trustee.id ?? holder.trustee.name}`}
                  holder={holder}
                />
              ))}
            </ul>
            {mailbox.holdersTruncated ? (
              <p className="signin-coverage">
                Showing the first {number.format(mailbox.holders.length)} holders.
              </p>
            ) : null}
          </>
        ) : null}
        {stored.unresolved > 0 ? (
          <p className="signin-coverage">
            {number.format(stored.unresolved)}{" "}
            {stored.unresolved === 1
              ? "grant in the export names a trustee"
              : "grants in the export name trustees"}{" "}
            the script could not resolve, so access through{" "}
            {stored.unresolved === 1 ? "it" : "them"} is not shown here.
          </p>
        ) : null}
        <p className="signin-coverage">
          This is what the export says. Prism365 cannot check it against Exchange.
          {stale
            ? " Access may have changed since; run the export again for a current picture."
            : ""}
        </p>
        {groupMailboxNote}
      </>
    );
  })();
  return (
    <section className="panel person-access" aria-label="Shared mailboxes">
      <h3>Shared mailboxes</h3>
      {body}
    </section>
  );
}

// The whole tab as a function of what has loaded. It holds no state, so a test can draw it with an
// answer in hand.
export function PersonAccessView({
  access,
  error,
  mailbox = null,
  mailboxError = "",
}: {
  // null while the answer loads.
  access: PersonAccessData | null;
  error: string;
  // The Shared mailboxes section's own answer, null while it loads.
  mailbox?: PersonMailboxAccess | null;
  mailboxError?: string;
}) {
  if (error) return <p className="empty-data">{error}</p>;
  if (!access)
    return (
      <section className="panel person-access" aria-label="Access">
        <SkeletonRows rows={5} label="Reading this person's access" />
      </section>
    );
  const results = access.audits.status === "ok" ? access.audits.results : [];
  const byId = new Map(results.map((result) => [result.auditId, result]));
  return (
    <>
      <CombinedRisksSection section={access.combinedRisks} />
      <GroupsSection memberships={access.memberships} results={byId} />
      <OwnsSection owned={access.owned} />
      <ConditionalAccessSection section={access.conditionalAccess} />
      {auditSections.map((group) => (
        <section className="panel person-access" aria-label={group.title} key={group.title}>
          <h3>{group.title}</h3>
          {access.audits.status !== "ok" ? (
            <CannotShow section={access.audits} />
          ) : (
            group.audits.map((id) => {
              const result = byId.get(id);
              return result ? <AuditSummary key={id} result={result} results={byId} /> : null;
            })
          )}
        </section>
      ))}
      <SharedMailboxesSection mailbox={mailbox} error={mailboxError} />
    </>
  );
}

// Loads when the tab first opens, not with the page, and keeps the answer when the tab is left.
export function PersonAccess({
  session,
  userId,
  active,
}: {
  session: PrismSession;
  userId: string;
  active: boolean;
}) {
  const [opened, setOpened] = useState(false);
  const [access, setAccess] = useState<PersonAccessData | null>(null);
  const [error, setError] = useState("");
  const [mailbox, setMailbox] = useState<PersonMailboxAccess | null>(null);
  const [mailboxError, setMailboxError] = useState("");
  const live = session.mode === "live" && session.authenticated;

  useEffect(() => {
    if (active) setOpened(true);
  }, [active]);

  // The shared mailbox answer is its own request, so a slow or failed one never holds up the rest.
  useEffect(() => {
    if (!live || !opened) return;
    const controller = new AbortController();
    readJson(
      `/api/v1/people/${encodeURIComponent(userId)}/mailbox-access`,
      (value) => personMailboxAccessSchema.parse(value),
      { signal: controller.signal },
    )
      .then((value) => {
        setMailbox(value);
        setMailboxError("");
      })
      .catch((loadError: unknown) => {
        if (controller.signal.aborted) return;
        setMailboxError(
          loadError instanceof Error
            ? loadError.message
            : "Shared mailbox access could not be loaded.",
        );
      });
    return () => controller.abort();
  }, [live, opened, userId]);

  useEffect(() => {
    if (!live || !opened) return;
    const controller = new AbortController();
    readJson(
      `/api/v1/people/${encodeURIComponent(userId)}/access`,
      (value) => personAccessSchema.parse(value),
      { signal: controller.signal },
    )
      .then((value) => {
        setAccess(value);
        setError("");
      })
      .catch((loadError: unknown) => {
        if (controller.signal.aborted) return;
        setError(loadError instanceof Error ? loadError.message : "Access could not be loaded.");
      });
    return () => controller.abort();
  }, [live, opened, userId]);

  if (!live || !opened) return null;
  return (
    <PersonAccessView access={access} error={error} mailbox={mailbox} mailboxError={mailboxError} />
  );
}

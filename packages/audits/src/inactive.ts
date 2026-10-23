import type { AuditRow } from "@prism365/contracts";
import type { Audit } from "./definition.ts";
import {
  activityOf,
  basisLabel,
  byName,
  coverageNotes,
  displayDate,
  INACTIVITY_THRESHOLDS,
  inactivityTags,
  licenseNamesNote,
  licensesOf,
  personCell,
  readLicenseNames,
  readUsers,
  statusOf,
  tagOf,
  typeOf,
  unreadableNote,
} from "./directory.ts";
import {
  domainColumn,
  domainInventory,
  excludedNote,
  inAuditScope,
  principalDomain,
  scopeNote,
} from "./scope.ts";

const thresholdLabel = (days: number) => (days === 365 ? "1 year+" : `${days}+ days`);

// Every account, with how long it has gone without signing in.
export const inactiveAccountsAudit: Audit = {
  definition: {
    id: "accounts.inactive",
    // Accounts an action tool could act on (docs/handoff.md).
    handoff: true,
    version: 3,
    title: "Inactive accounts",
    description:
      "Every account and how long it has gone without signing in, from Microsoft's last sign-in dates. Pick a range or sort by days.",
    category: "Accounts",
    subject: { one: "account", many: "accounts" },
    columns: [
      { key: "person", label: "Account", kind: "person" },
      {
        key: "days",
        label: "Inactive for",
        kind: "days",
        empty: "Unknown",
        help: "Days since the last successful sign-in. With no success on record, since the last attempt; with no sign-in at all, since the account was created.",
      },
      { key: "basis", label: "Counted from", kind: "enum" },
      {
        key: "lastSuccess",
        label: "Last successful sign-in",
        kind: "date",
        empty: "None recorded",
        help: "Interactive or background. Microsoft has recorded this since 1 December 2023.",
      },
      {
        key: "lastInteractive",
        label: "Last interactive attempt",
        kind: "date",
        hidden: true,
        help: "The last time the person signed in themselves, successfully or not.",
      },
      {
        key: "lastNonInteractive",
        label: "Last background attempt",
        kind: "date",
        hidden: true,
        help: "The last time an app signed in for the person, successfully or not.",
      },
      { key: "status", label: "Status", kind: "enum", facet: true },
      { key: "type", label: "Type", kind: "enum", facet: true },
      domainColumn,
      { key: "licenses", label: "Licenses", kind: "list" },
      { key: "licensed", label: "Licensed", kind: "enum", facet: true, hidden: true },
      { key: "created", label: "Created", kind: "date", empty: "Unknown" },
      { key: "source", label: "Source", kind: "enum", facet: true, hidden: true },
      { key: "department", label: "Department", kind: "text", hidden: true },
      { key: "jobTitle", label: "Job title", kind: "text", hidden: true },
    ],
    views: [
      { id: "all", label: "All accounts", tags: [] },
      ...INACTIVITY_THRESHOLDS.map((days) => ({
        id: `d${days}`,
        label: thresholdLabel(days),
        tags: [`d${days}`],
      })),
      { id: "never", label: "Never signed in", tags: ["never"] },
      { id: "no-success", label: "Attempts, no success", tags: ["no-success"] },
    ],
    defaultView: "d30",
    defaultSort: { column: "days", direction: "desc" },
    requirements: [
      "Microsoft Entra ID P1 or P2 in the tenant, which Microsoft requires for sign-in dates.",
      "A role that can read sign-in activity, such as Reports Reader, Security Reader or Global Reader.",
    ],
  },
  packs: ["directory.users.read", "audit.signins.read"],
  optionalPacks: ["tenant.licenses.read"],
  scoped: true,
  forbidden:
    "Microsoft refused to return last sign-in dates. They need AuditLog.Read.All, Microsoft Entra ID P1 or P2, and a role such as Reports Reader, Security Reader or Global Reader.",
  async run(context) {
    const read = await readUsers(context, { signInActivity: "required" });
    const names = await readLicenseNames(context);
    // Filtered first, so the rows, tiles and notes below all describe the same accounts.
    const covered = read.users.filter((user) =>
      inAuditScope(context.scope, user.userPrincipalName),
    );
    const limited = scopeNote(context.scope);
    const left = excludedNote(read.users.length - covered.length);
    // Attempts with no success on record are worth checking whatever their domain (an
    // emergency-access account often sits on the onmicrosoft.com domain), so the scope never
    // hides one silently.
    const outsideNoSuccess = read.users.filter(
      (user) =>
        !inAuditScope(context.scope, user.userPrincipalName) &&
        activityOf(user, context.now).noSuccess,
    ).length;
    const rows: AuditRow[] = covered.map((user) => {
      const activity = activityOf(user, context.now);
      const licenses = licensesOf(user, names);
      const status = statusOf(user);
      const type = typeOf(user.userType);
      return {
        key: user.id,
        personId: user.id,
        tags: [
          ...inactivityTags(activity),
          ...tagOf(status),
          ...tagOf(type),
          ...(licenses.length > 0 ? ["licensed"] : []),
        ],
        cells: {
          person: personCell(user),
          days: activity.days,
          basis: basisLabel(activity),
          lastSuccess: displayDate(user.signInActivity?.lastSuccessfulSignInDateTime ?? null),
          lastInteractive: displayDate(user.signInActivity?.lastSignInDateTime ?? null),
          lastNonInteractive: displayDate(
            user.signInActivity?.lastNonInteractiveSignInDateTime ?? null,
          ),
          status,
          type,
          domain: principalDomain(user.userPrincipalName),
          licenses,
          licensed: licenses.length > 0 ? "Licensed" : "Unlicensed",
          created: displayDate(user.createdDateTime),
          source: user.onPremisesSyncEnabled ? "Synced from on-premises" : "Cloud",
          department: user.department,
          jobTitle: user.jobTitle,
        },
      };
    });
    rows.sort(byName);
    const count = (tag: string) => rows.filter((row) => row.tags.includes(tag)).length;
    const noSuccess = count("no-success");
    const unreadable = count("unreadable");
    return {
      domains: domainInventory(read.users.map((user) => user.userPrincipalName)),
      rows,
      complete: read.complete && read.skipped === 0 && unreadable === 0,
      summary: [
        { label: "Accounts", value: rows.length, tone: "neutral", view: "all" },
        { label: "Inactive 30+ days", value: count("d30"), tone: "warning", view: "d30" },
        { label: "Inactive 90+ days", value: count("d90"), tone: "warning", view: "d90" },
        { label: "Never signed in", value: count("never"), tone: "neutral", view: "never" },
        {
          label: "Attempts, no success",
          value: noSuccess,
          tone: noSuccess > 0 ? "critical" : "neutral",
          view: "no-success",
        },
        ...(outsideNoSuccess > 0
          ? [
              {
                label: "Attempts, no success outside the scope",
                value: outsideNoSuccess,
                tone: "critical" as const,
                view: null,
              },
            ]
          : []),
      ],
      notes: [
        ...(limited ? [limited] : []),
        ...(left ? [left] : []),
        ...(outsideNoSuccess > 0
          ? [
              `${outsideNoSuccess === 1 ? "1 account" : `${outsideNoSuccess} accounts`} with sign-in attempts and no success on record ${outsideNoSuccess === 1 ? "is" : "are"} in domains the audit scope leaves out. Choose All domains and run again to see ${outsideNoSuccess === 1 ? "it" : "them"}.`,
            ]
          : []),
        "Days count from the last successful sign-in, interactive or background. Microsoft has recorded successful sign-ins only since 1 December 2023, so an account with none since then is counted from its last attempt, which can only make it look more recent.",
        "Attempts, no success: accounts that tried to sign in after December 2023 with no successful sign-in on record. Check them for a broken account or someone guessing passwords.",
        "Never signed in means Microsoft has no sign-in on record. Its records start in April 2020, so these accounts are counted from creation or April 2020, whichever is later.",
        "Shared mailboxes, rooms and equipment are accounts too and normally never sign in.",
        ...unreadableNote(unreadable),
        ...(names ? [] : [licenseNamesNote]),
        ...coverageNotes(read),
      ],
    };
  },
};

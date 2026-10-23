import { type AuditRow, cleanText } from "@prism365/contracts";
import { z } from "zod";
import type { Audit } from "./definition.ts";
import { byName, DAY_MS, MAX_ACCOUNTS } from "./directory.ts";
import { createPacer, type Pause, readPaced } from "./paced-read.ts";
import {
  domainColumn,
  domainInventory,
  excludedNote,
  inAuditScope,
  principalDomain,
  scopeNote,
} from "./scope.ts";

// Users Microsoft Entra ID Protection has marked as risky. From Microsoft Graph v1.0, checked
// 2026-10-22:
// https://learn.microsoft.com/en-us/graph/api/riskyuser-list?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/resources/riskyuser?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/resources/riskdetail?view=graph-rest-1.0

const riskContext = {
  feature: "identity.risk.read",
  permission: "IdentityRiskyUser.Read.All",
} as const;

// The largest page Microsoft documents for this endpoint.
const PAGE_SIZE = 500;
const MAX_PAGES = 25;
const CLOSED_WINDOW_DAYS = 30;

const P2_SENTENCE = "Risky user data in Microsoft Graph needs Microsoft Entra ID P2.";

const text = z
  .string()
  .transform((value) => cleanText(value.slice(0, 500)))
  .nullish()
  .catch(null)
  .transform((value) => value ?? null);
const flag = z
  .boolean()
  .nullish()
  .catch(null)
  .transform((value) => value ?? null);

const riskyUserSchema = z.object({
  id: z
    .string()
    .min(1)
    .max(200)
    .regex(/^[A-Za-z0-9._@:-]+$/),
  isDeleted: flag,
  isProcessing: flag,
  riskLastUpdatedDateTime: z
    .string()
    .datetime({ offset: true })
    .nullish()
    .catch(null)
    .transform((value) => value ?? null),
  riskLevel: text,
  riskState: text,
  riskDetail: text,
  userDisplayName: text,
  userPrincipalName: text,
});

const levelLabels: Record<string, string> = {
  high: "High",
  medium: "Medium",
  low: "Low",
  hidden: "Hidden",
  none: "None",
};

const stateLabels: Record<string, string> = {
  atrisk: "At risk",
  confirmedcompromised: "Confirmed compromised",
  remediated: "Remediated",
  dismissed: "Dismissed",
  confirmedsafe: "Confirmed safe",
  none: "None",
};

// Why the risk is what it is. The first group is what Microsoft always returns; the second only
// comes back when a request asks for unknown enum members, which this audit does not, but a
// tenant may still send them, so they read as words too.
const detailLabels: Record<string, string> = {
  none: "None",
  adminGeneratedTemporaryPassword: "Admin set a temporary password",
  userPerformedSecuredPasswordChange: "User changed the password securely",
  userPerformedSecuredPasswordReset: "User reset the password securely",
  adminConfirmedSigninSafe: "Admin confirmed the sign-in was safe",
  aiConfirmedSigninSafe: "Microsoft's automated review confirmed the sign-in was safe",
  userPassedMFADrivenByRiskBasedPolicy: "User passed MFA required by a risk policy",
  adminDismissedAllRiskForUser: "Admin dismissed all risk for the user",
  adminConfirmedSigninCompromised: "Admin confirmed the sign-in was compromised",
  hidden: "Hidden",
  adminConfirmedUserCompromised: "Admin confirmed the user was compromised",
  m365DAdminDismissedDetection: "Microsoft 365 Defender admin dismissed the detection",
  adminDismissedAllRiskForServicePrincipal: "Admin dismissed all risk for the service principal",
  adminConfirmedServicePrincipalCompromised:
    "Admin confirmed the service principal was compromised",
  userChangedPasswordOnPremises: "User changed the password on premises",
  adminDismissedRiskForSignIn: "Admin dismissed the risk for the sign-in",
  adminConfirmedAccountSafe: "Admin confirmed the account was safe",
  microsoftRevokedSessions: "Microsoft revoked the sessions",
  unknownFutureValue: "Unknown",
};
const detailByLowerKey = new Map(
  Object.entries(detailLabels).map(([key, value]) => [key.toLowerCase(), value] as const),
);

function humanize(value: string): string {
  const spaced = value.replace(/([a-z0-9])([A-Z])/g, "$1 $2").toLowerCase();
  return `${spaced.charAt(0).toUpperCase()}${spaced.slice(1)}`;
}

function labelled(value: string | null, labels: Record<string, string>): string {
  if (!value) return "Unknown";
  return labels[value.toLowerCase()] ?? "Unknown";
}

function detailOf(value: string | null): string | null {
  if (!value) return null;
  return detailByLowerKey.get(value.toLowerCase()) ?? humanize(value);
}

export interface RiskyUsersOptions {
  // How the audit waits between pages. Tests replace it.
  pause?: Pause;
}

export function createRiskyUsersAudit(options: RiskyUsersOptions = {}): Audit {
  return {
    definition: {
      id: "identity.risky-users",
      version: 1,
      title: "Risky users",
      description:
        "Accounts Microsoft Entra ID Protection marks as at risk or compromised, with the risk level, why, and whether the risk has been dealt with.",
      category: "Identity",
      subject: { one: "user", many: "users" },
      columns: [
        { key: "person", label: "User", kind: "person" },
        {
          key: "level",
          label: "Risk level",
          kind: "enum",
          facet: true,
          help: "Microsoft's level for the account. Hidden means the account is not covered by ID Protection.",
        },
        { key: "state", label: "Risk state", kind: "enum", facet: true },
        {
          key: "detail",
          label: "Risk detail",
          kind: "enum",
          facet: true,
          empty: "None",
          help: "Why Microsoft last changed the risk, such as a password reset or an admin decision.",
        },
        { key: "updated", label: "Risk updated", kind: "date", empty: "Unknown" },
        { key: "days", label: "Since update", kind: "days", empty: "Unknown" },
        {
          key: "processing",
          label: "Being processed",
          kind: "boolean",
          help: "Yes while Microsoft is still working out the user's risk, so the level may change.",
        },
        domainColumn,
      ],
      views: [
        { id: "at-risk", label: "At risk", tags: ["at-risk"] },
        { id: "high-risk", label: "High risk", tags: ["high", "at-risk"] },
        { id: "closed", label: "Remediated or dismissed", tags: ["closed"] },
        { id: "all", label: "All", tags: [] },
      ],
      defaultView: "at-risk",
      defaultSort: { column: "updated", direction: "desc" },
      requirements: [
        "Microsoft Entra ID P2.",
        "A role such as Global Reader, Security Reader, Security Operator or Security Administrator.",
      ],
    },
    packs: ["identity.risk.read"],
    optionalPacks: [],
    scoped: true,
    forbidden: `Microsoft refused to list risky users. ${P2_SENTENCE} It also needs IdentityRiskyUser.Read.All and a role such as Global Reader, Security Reader, Security Operator or Security Administrator.`,
    async run(context) {
      context.signal.throwIfAborted();
      const read = await readPaced(
        context.graph,
        `identityProtection/riskyUsers?$top=${PAGE_SIZE}`,
        riskContext,
        {
          pacer: createPacer(options.pause),
          signal: context.signal,
          maxPages: MAX_PAGES,
          maxRecords: MAX_ACCOUNTS,
        },
      );
      let skipped = 0;
      let deleted = 0;
      let outside = 0;
      let outsideUrgent = 0;
      const principals: (string | null)[] = [];
      const rows: AuditRow[] = [];
      for (const raw of read.items) {
        const parsed = riskyUserSchema.safeParse(raw);
        if (!parsed.success) {
          skipped += 1;
          continue;
        }
        const user = parsed.data;
        // Microsoft keeps a deleted user listed until its risk is closed. There is no one left to
        // act on, so it is counted in a note instead of the table.
        if (user.isDeleted === true) {
          if (inAuditScope(context.scope, user.userPrincipalName)) deleted += 1;
          continue;
        }
        const level = labelled(user.riskLevel, levelLabels);
        const state = labelled(user.riskState, stateLabels);
        const open = state === "At risk" || state === "Confirmed compromised";
        principals.push(user.userPrincipalName);
        if (!inAuditScope(context.scope, user.userPrincipalName)) {
          outside += 1;
          // A compromised or high-risk account matters whatever its domain, so the scope never
          // hides one silently.
          if (state === "Confirmed compromised" || (level === "High" && open)) outsideUrgent += 1;
          continue;
        }
        const updated = user.riskLastUpdatedDateTime;
        const days =
          updated === null
            ? null
            : Math.max(0, Math.floor((context.now.getTime() - Date.parse(updated)) / DAY_MS));
        const closed = state === "Remediated" || state === "Dismissed";
        rows.push({
          key: user.id,
          personId: user.id,
          tags: [
            ...(open ? ["at-risk"] : []),
            ...(state === "Confirmed compromised" ? ["compromised"] : []),
            ...(closed ? ["closed"] : []),
            ...(level === "High" ? ["high"] : []),
            ...(level === "Medium" ? ["medium"] : []),
            ...(level === "Low" ? ["low"] : []),
            ...(level === "Hidden" ? ["hidden"] : []),
          ],
          cells: {
            person: { name: user.userDisplayName, principal: user.userPrincipalName },
            level,
            state,
            detail: detailOf(user.riskDetail),
            updated,
            days,
            processing: user.isProcessing,
            domain: principalDomain(user.userPrincipalName),
          },
        });
      }
      rows.sort(byName);
      const count = (...tags: string[]) =>
        rows.filter((row) => tags.every((tag) => row.tags.includes(tag))).length;
      const recentlyClosed = rows.filter(
        (row) =>
          row.tags.includes("closed") &&
          typeof row.cells.days === "number" &&
          row.cells.days <= CLOSED_WINDOW_DAYS,
      ).length;
      const compromised = count("compromised");
      const highOpen = count("high", "at-risk");
      const mediumOpen = count("medium", "at-risk");
      const limited = scopeNote(context.scope);
      const left = excludedNote(outside);
      // Never "good" under a scope: a clean count there says nothing about the other domains.
      const clean = context.scope ? "neutral" : "good";
      return {
        domains: domainInventory(principals),
        rows,
        complete: read.complete && skipped === 0,
        summary: [
          {
            label: "Confirmed compromised",
            value: compromised,
            tone: compromised > 0 ? "critical" : clean,
            view: "at-risk",
          },
          {
            label: "High risk, still at risk",
            value: highOpen,
            tone: highOpen > 0 ? "critical" : clean,
            view: "high-risk",
          },
          {
            label: "Medium risk, still at risk",
            value: mediumOpen,
            tone: mediumOpen > 0 ? "warning" : "neutral",
            view: "at-risk",
          },
          {
            label: "Remediated or dismissed in the last 30 days",
            value: recentlyClosed,
            tone: "neutral",
            view: "closed",
          },
          ...(outsideUrgent > 0
            ? [
                {
                  label: "High risk or compromised outside the scope",
                  value: outsideUrgent,
                  tone: "critical" as const,
                  view: null,
                },
              ]
            : []),
        ],
        notes: [
          ...(limited ? [limited] : []),
          ...(left ? [left] : []),
          ...(outsideUrgent > 0
            ? [
                `${outsideUrgent === 1 ? "1 user" : `${outsideUrgent} users`} at high risk or confirmed compromised ${outsideUrgent === 1 ? "is" : "are"} in domains the audit scope leaves out. Choose All domains and run again to see ${outsideUrgent === 1 ? "it" : "them"}.`,
              ]
            : []),
          ...(read.items.length === 0
            ? [
                `Microsoft returned no risky users. That can mean none exist, or that the tenant lacks the license. ${P2_SENTENCE}`,
              ]
            : []),
          ...(deleted > 0
            ? [
                `${deleted} deleted ${deleted === 1 ? "user" : "users"} that Microsoft still lists ${deleted === 1 ? "was" : "were"} left out.`,
              ]
            : []),
          "Microsoft keeps a risky user listed until the risk is remediated, so an old entry stays until someone closes it.",
          "Hidden means the account is not covered by ID Protection, so Microsoft reports no risk for it.",
          "These are Microsoft's detections. Prism365 adds none and judges nothing itself.",
          ...(skipped > 0
            ? [
                `${skipped} ${skipped === 1 ? "user" : "users"} could not be read and ${skipped === 1 ? "is" : "are"} left out.`,
              ]
            : []),
          ...(read.complete
            ? []
            : [
                `Stopped at ${MAX_ACCOUNTS.toLocaleString("en-US")} risky users; the rest were not read.`,
              ]),
        ],
      };
    },
  };
}

export const riskyUsersAudit: Audit = createRiskyUsersAudit();

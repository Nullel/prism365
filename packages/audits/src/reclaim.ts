import type { AuditRow } from "@prism365/contracts";
import type { Audit } from "./definition.ts";
import {
  activityOf,
  byName,
  coverageNotes,
  type DirectoryUser,
  displayDate,
  inactivityTags,
  licenseNamesNote,
  licensesOf,
  personCell,
  readLicenseNames,
  readUsers,
  signInNote,
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

const RECLAIM_AFTER_DAYS = 30;

function assignment(user: DirectoryUser): { label: string; byGroup: boolean } {
  const states = user.licenseAssignmentStates;
  if (!states) return { label: "Unknown", byGroup: false };
  const held = new Set(user.assignedLicenses);
  const relevant = states.filter((state) => state.skuId && held.has(state.skuId));
  const byGroup = relevant.some((state) => Boolean(state.assignedByGroup));
  const direct = relevant.some((state) => !state.assignedByGroup);
  if (byGroup && direct) return { label: "Directly and by group", byGroup };
  if (byGroup) return { label: "By group", byGroup };
  return { label: "Directly", byGroup };
}

// Licensed accounts nobody is using: disabled, or no sign-in for 30 days or more.
export const licenseReclaimAudit: Audit = {
  definition: {
    id: "licenses.reclaim",
    version: 3,
    title: "Licenses to reclaim",
    description:
      "Licensed accounts that are disabled or have not signed in for 30 days or more, and how each license was assigned.",
    category: "Licenses",
    subject: { one: "account", many: "accounts" },
    columns: [
      { key: "person", label: "Account", kind: "person" },
      { key: "status", label: "Status", kind: "enum", facet: true },
      {
        key: "days",
        label: "Inactive for",
        kind: "days",
        empty: "Unknown",
        help: "Counted the same way as the Inactive accounts audit. Empty when last sign-in dates could not be read.",
      },
      {
        key: "lastSuccess",
        label: "Last successful sign-in",
        kind: "date",
        empty: "None recorded",
      },
      { key: "licenses", label: "Licenses", kind: "list" },
      { key: "licenseCount", label: "Count", kind: "number" },
      {
        key: "assignment",
        label: "Assigned",
        kind: "enum",
        facet: true,
        help: "A license assigned by a group comes back if you remove it directly. Remove the account from the group instead.",
      },
      { key: "type", label: "Type", kind: "enum", facet: true, hidden: true },
      domainColumn,
      { key: "created", label: "Created", kind: "date", hidden: true, empty: "Unknown" },
    ],
    views: [
      { id: "all", label: "All candidates", tags: [] },
      { id: "disabled", label: "Disabled", tags: ["disabled"] },
      { id: "d60", label: "60+ days", tags: ["d60"] },
      { id: "d90", label: "90+ days", tags: ["d90"] },
      { id: "d180", label: "180+ days", tags: ["d180"] },
      { id: "never", label: "Never signed in", tags: ["never"] },
      { id: "group", label: "Assigned by group", tags: ["group"] },
    ],
    defaultView: "all",
    defaultSort: { column: "days", direction: "desc" },
    requirements: [
      "Without sign-in dates (AuditLog.Read.All with Microsoft Entra ID P1 or P2), only disabled accounts are listed.",
    ],
  },
  packs: ["directory.users.read"],
  optionalPacks: ["audit.signins.read", "tenant.licenses.read"],
  scoped: true,
  forbidden: "Microsoft refused to list accounts. Listing them needs User.Read.All.",
  async run(context) {
    const read = await readUsers(context, { signInActivity: "optional", groupAssignment: true });
    const names = await readLicenseNames(context);
    const datesRead = read.signInActivity === "read";
    const rows: AuditRow[] = [];
    // Licensed, enabled accounts whose dates could not be read: not listed, but counted.
    let unreadable = 0;
    const byProduct = new Map<string, number>();
    let outside = 0;
    for (const user of read.users) {
      if (!inAuditScope(context.scope, user.userPrincipalName)) {
        outside += 1;
        continue;
      }
      if (user.assignedLicenses.length === 0) continue;
      const activity = datesRead ? activityOf(user, context.now) : null;
      const status = statusOf(user);
      if (status !== "Disabled" && activity?.basis === "unreadable") unreadable += 1;
      if (status !== "Disabled" && !((activity?.days ?? 0) >= RECLAIM_AFTER_DAYS)) continue;
      const licenses = licensesOf(user, names);
      for (const license of licenses) byProduct.set(license, (byProduct.get(license) ?? 0) + 1);
      const assigned = assignment(user);
      rows.push({
        key: user.id,
        personId: user.id,
        tags: [
          ...(activity ? inactivityTags(activity) : []),
          ...tagOf(status),
          ...(assigned.byGroup ? ["group"] : []),
        ],
        cells: {
          person: personCell(user),
          status,
          days: activity?.days ?? null,
          lastSuccess: displayDate(user.signInActivity?.lastSuccessfulSignInDateTime ?? null),
          licenses,
          licenseCount: licenses.length,
          assignment: assigned.label,
          type: typeOf(user.userType),
          domain: principalDomain(user.userPrincipalName),
          created: displayDate(user.createdDateTime),
        },
      });
    }
    rows.sort(byName);
    const count = (tag: string) => rows.filter((row) => row.tags.includes(tag)).length;
    const held = rows.reduce((total, row) => total + Number(row.cells.licenseCount ?? 0), 0);
    const limited = scopeNote(context.scope);
    const left = excludedNote(outside);
    const products = [...byProduct.entries()]
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
      .slice(0, 12)
      .map(([name, seats]) => `${name} ${seats}`);
    return {
      domains: domainInventory(read.users.map((user) => user.userPrincipalName)),
      rows,
      complete: read.complete && read.skipped === 0 && datesRead && unreadable === 0,
      summary: [
        { label: "Accounts", value: rows.length, tone: "neutral", view: "all" },
        { label: "Licenses held", value: held, tone: held > 0 ? "warning" : "neutral", view: null },
        { label: "Disabled", value: count("disabled"), tone: "warning", view: "disabled" },
        { label: "Inactive 90+ days", value: count("d90"), tone: "warning", view: "d90" },
        { label: "Assigned by group", value: count("group"), tone: "neutral", view: "group" },
      ],
      notes: [
        ...(limited ? [limited] : []),
        ...(left ? [left] : []),
        ...(products.length > 0 ? [`By product: ${products.join(", ")}.`] : []),
        "Microsoft does not report prices, so Prism365 counts seats, not cost.",
        ...(read.groupAssignment
          ? []
          : ["How licenses were assigned could not be read, so Assigned shows Unknown."]),
        ...signInNote(read),
        ...unreadableNote(unreadable),
        ...(names ? [] : [licenseNamesNote]),
        ...coverageNotes(read),
      ],
    };
  },
};

import type { AuditRow } from "@prism365/contracts";
import type { Audit } from "./definition.ts";
import {
  activityOf,
  byName,
  coverageNotes,
  DAY_MS,
  displayDate,
  guestDomain,
  inactivityTags,
  isGuest,
  licenseNamesNote,
  licensesOf,
  personCell,
  readLicenseNames,
  readUsers,
  signInNote,
  statusOf,
  tagOf,
  unreadableNote,
} from "./directory.ts";

const PENDING_REVIEW_DAYS = 30;

function invitation(state: string | null): "Accepted" | "Not accepted" | "Unknown" {
  const value = state?.toLowerCase();
  if (value === "accepted") return "Accepted";
  if (value === "pendingacceptance") return "Not accepted";
  return "Unknown";
}

// External guests: invitations never accepted, guests who never signed in, guests gone quiet.
export const guestAccountsAudit: Audit = {
  definition: {
    id: "guests.review",
    version: 1,
    title: "Guest accounts",
    description:
      "External guests in the tenant: invitations never accepted, guests who never signed in, and guests who have gone quiet.",
    category: "Accounts",
    subject: { one: "guest", many: "guests" },
    columns: [
      { key: "person", label: "Guest", kind: "person" },
      { key: "email", label: "Email", kind: "text" },
      { key: "domain", label: "Home domain", kind: "enum", facet: true },
      { key: "invitation", label: "Invitation", kind: "enum", facet: true },
      {
        key: "invited",
        label: "Invited",
        kind: "date",
        empty: "Unknown",
        help: "When the guest account was created, which is when the invitation was sent.",
      },
      { key: "stateChanged", label: "Invitation changed", kind: "date", hidden: true },
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
      { key: "status", label: "Status", kind: "enum", facet: true },
      { key: "licenses", label: "Licenses", kind: "list", hidden: true },
    ],
    views: [
      { id: "all", label: "All guests", tags: [] },
      { id: "pending", label: "Not accepted", tags: ["pending"] },
      { id: "pending-30", label: "Not accepted after 30 days", tags: ["pending-30"] },
      { id: "never", label: "Never signed in", tags: ["never"] },
      { id: "d90", label: "Inactive 90+ days", tags: ["d90"] },
      { id: "d180", label: "Inactive 180+ days", tags: ["d180"] },
    ],
    defaultView: "all",
    defaultSort: { column: "days", direction: "desc" },
    requirements: [
      "Sign-in dates need AuditLog.Read.All and Microsoft Entra ID P1 or P2. Without them, only invitation states are shown.",
    ],
  },
  packs: ["directory.users.read"],
  optionalPacks: ["audit.signins.read", "tenant.licenses.read"],
  scoped: false,
  forbidden: "Microsoft refused to list accounts. Listing them needs User.Read.All.",
  async run(context) {
    const read = await readUsers(context, { signInActivity: "optional" });
    const names = await readLicenseNames(context);
    const datesRead = read.signInActivity === "read";
    const rows: AuditRow[] = [];
    for (const user of read.users) {
      if (!isGuest(user)) continue;
      const activity = datesRead ? activityOf(user, context.now) : null;
      const state = invitation(user.externalUserState);
      const invited = displayDate(user.createdDateTime);
      const created = invited ? Date.parse(invited) : null;
      const pendingDays =
        created === null ? 0 : Math.floor((context.now.getTime() - created) / DAY_MS);
      const status = statusOf(user);
      rows.push({
        key: user.id,
        personId: user.id,
        tags: [
          ...(activity ? inactivityTags(activity) : []),
          ...(state === "Not accepted" ? ["pending"] : []),
          ...(state === "Not accepted" && pendingDays >= PENDING_REVIEW_DAYS ? ["pending-30"] : []),
          ...tagOf(status),
        ],
        cells: {
          person: personCell(user),
          email: user.mail,
          domain: guestDomain(user),
          invitation: state,
          invited,
          stateChanged: user.externalUserStateChangeDateTime,
          days: activity?.days ?? null,
          lastSuccess: displayDate(user.signInActivity?.lastSuccessfulSignInDateTime ?? null),
          status,
          licenses: licensesOf(user, names),
        },
      });
    }
    rows.sort(byName);
    const count = (tag: string) => rows.filter((row) => row.tags.includes(tag)).length;
    const domains = new Set(rows.map((row) => row.cells.domain).filter(Boolean)).size;
    const unreadable = count("unreadable");
    return {
      rows,
      complete: read.complete && read.skipped === 0 && datesRead && unreadable === 0,
      summary: [
        { label: "Guests", value: rows.length, tone: "neutral", view: "all" },
        { label: "Not accepted", value: count("pending"), tone: "warning", view: "pending" },
        { label: "Never signed in", value: count("never"), tone: "warning", view: "never" },
        { label: "Inactive 90+ days", value: count("d90"), tone: "warning", view: "d90" },
        { label: "Home domains", value: domains, tone: "neutral", view: null },
      ],
      notes: [
        "Every guest here can sign in to whatever was shared with it until the account is disabled or removed.",
        ...signInNote(read),
        ...unreadableNote(unreadable),
        ...(names ? [] : [licenseNamesNote]),
        ...coverageNotes(read),
      ],
    };
  },
};

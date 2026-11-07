import { guid, RESEARCH_DOMAIN, type SyntheticTenant } from "./tenant.ts";

// The demo's mailbox access export (ADR-0012), as Export-MailboxAccess.ps1 would write it for the
// synthetic tenant. Prism365 never calls Exchange, so the recorder imports this file the way an
// administrator imports theirs.
//
// Two shared mailboxes: Lab Results, which the laboratory team reaches through its group and the
// technician may send as, and an instrument bookings mailbox the technician holds directly, which
// a lab member sends on behalf of and the Madison file share readers reach through nesting. One
// Full Access trustee is a deleted account's SID that does not resolve.

// The groups from graph/groups.ts, by the same ids. Importing that file here would load it before
// the tenant it extends.
const LABORATORY = { id: guid("group:laboratory"), name: "Laboratory team" };
const FILE_SHARE_READERS = {
  id: guid("group:file-share-readers"),
  name: "Madison file share readers",
};
const BOOKINGS = {
  id: guid("mailbox:instrument-bookings"),
  address: `instrument.bookings@${RESEARCH_DOMAIN}`,
  displayName: "Instrument bookings",
};
export const UNRESOLVED_TRUSTEE = "S-1-5-21-3110593217-2318866104-1205783442-1187";

export function demoMailboxAccessExport(
  tenant: SyntheticTenant,
  exportedAt: Date,
  tenantId: string,
) {
  const byKey = (key: string) => {
    const found = tenant.users.find((user) => user.id === guid(`user:${key}`));
    if (!found) throw new Error(`The synthetic tenant has no ${key}.`);
    return found;
  };
  const lab = byKey("labresults");
  const technician = byKey("tvance");
  const member = tenant.users.find(
    (user) => user.department === "Laboratory" && user.userType === "Member" && user.accountEnabled,
  );
  if (!member) throw new Error("The synthetic tenant has no enabled laboratory member.");
  const person = (user: typeof lab) => ({
    id: user.id,
    address: user.mail,
    name: user.displayName,
    kind: "user" as const,
  });
  const group = (record: { id: string; name: string }) => ({
    id: record.id,
    address: null,
    name: record.name,
    kind: "group" as const,
  });
  return {
    format: "prism365.mailbox-access",
    version: 1,
    tenantId,
    exportedAt: exportedAt.toISOString().replace(/\.\d{3}Z$/, "Z"),
    includeUserMailboxes: false,
    mailboxes: [
      { id: lab.id, address: lab.mail, displayName: lab.displayName, type: "SharedMailbox" },
      { ...BOOKINGS, type: "SharedMailbox" },
    ],
    grants: [
      { mailboxId: lab.id, right: "FullAccess", trustee: group(LABORATORY) },
      {
        mailboxId: lab.id,
        right: "FullAccess",
        trustee: { id: null, address: null, name: UNRESOLVED_TRUSTEE, kind: "unresolved" },
      },
      { mailboxId: lab.id, right: "SendAs", trustee: person(technician) },
      { mailboxId: BOOKINGS.id, right: "FullAccess", trustee: person(technician) },
      { mailboxId: BOOKINGS.id, right: "FullAccess", trustee: group(FILE_SHARE_READERS) },
      { mailboxId: BOOKINGS.id, right: "SendOnBehalf", trustee: person(member) },
    ],
    unresolved: 1,
  };
}

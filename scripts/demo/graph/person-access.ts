import { privilegedRoles } from "@prism365/graph";
import { guid } from "../tenant.ts";
import { groupRecords, groupsOwnedBy } from "./groups.ts";
import type { DemoGraphExtension } from "./types.ts";
import { membershipsOf } from "./usage-reports.ts";

// What the person page's Access tab reads that the other features do not already serve: the
// groups and roles an account is in once nesting is followed (transitiveMemberOf) and the objects
// it owns (ownedObjects). Direct memberships are the ones usage-reports.ts serves, so the tab can
// tell direct from nested.
//
// The nestings: the group that gives accounts their Business Premium license is itself a member
// of the Madison file share readers security group, so everyone licensed through it is in that
// group too, without being a member of it. Administrators also own what they run: the groups from
// groups.ts that name them as owners, and the instrument gateway and lab portal registrations.

const FILE_SHARE_READERS = guid("group:file-share-readers");
const LICENSING = guid("group:licensing");
const IT_OPERATIONS = guid("group:it-operations-admins");
const TECHNICIAN = guid("user:tvance");

// Role template ids by name: the shared privileged roles plus the one this tenant adds.
const roleTemplates = new Map<string, string>([
  ...[...privilegedRoles].map(([id, name]) => [name, id] as [string, string]),
  ["Helpdesk Administrator", "729827e3-9c14-49f7-bb1b-9608f156bbb8"],
]);

// Applications (registrations, not their enterprise copies) by owner roster key.
const applicationOwners: Record<string, { key: string; name: string }[]> = {
  tvance: [{ key: "lab-portal", name: "Larkspur Lab Portal" }],
  admin: [{ key: "instrument-gateway", name: "Larkspur Instrument Gateway" }],
};

const PATH = /^users\/([^/]+)\/(transitiveMemberOf|ownedObjects)$/;

export const personAccess: DemoGraphExtension = (request, tools) => {
  const found = PATH.exec(request.path);
  if (!found) return null;
  const user = request.tenant.users.find((candidate) => candidate.id === found[1]);
  if (!user) return tools.notFound();
  const groups = new Map(
    groupRecords(request.tenant).map((record) => [
      record.id,
      { "@odata.type": "#microsoft.graph.group", ...record },
    ]),
  );

  if (found[2] === "ownedObjects") {
    const keys = Object.entries(applicationOwners).find(
      ([owner]) => guid(`user:${owner}`) === user.id,
    )?.[1];
    return tools.page(
      request.url,
      [
        ...groupsOwnedBy(request.tenant, user.id).map((record) => groups.get(record.id) ?? record),
        ...(keys ?? []).map((app) => ({
          "@odata.type": "#microsoft.graph.application",
          id: guid(`app-object:${app.key}`),
          displayName: app.name,
        })),
      ],
      request.url.searchParams.get("$select"),
    );
  }

  const direct = membershipsOf(user);
  const entries: object[] = direct.map((entry) => {
    const record = groups.get(entry.id);
    if (record) return record;
    // A directory role, with the template Conditional Access names it by.
    return { ...entry, roleTemplateId: roleTemplates.get(entry.displayName) ?? null };
  });
  const nested = groups.get(FILE_SHARE_READERS);
  if (nested && direct.some((entry) => entry.id === LICENSING)) entries.push(nested);
  // The technician reaches the IT operations admins group, which holds the Helpdesk Administrator
  // role, through a group nested in it that the synthetic tenant does not model.
  const itOperations = groups.get(IT_OPERATIONS);
  if (itOperations && user.id === TECHNICIAN) entries.push(itOperations);
  return tools.page(request.url, entries, request.url.searchParams.get("$select"));
};

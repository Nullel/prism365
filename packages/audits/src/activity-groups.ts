import { cleanText } from "@prism365/contracts";
import { GraphBoundaryError, GraphRequestError, GraphRetryDeferredError } from "@prism365/graph";
import { z } from "zod";
import type { AuditContext } from "./definition.ts";
import type { DirectoryUser } from "./directory.ts";

// Each account's direct memberships: its groups, its directory roles and its administrative
// units. Nested groups are not followed (transitiveMemberOf is a different call). Microsoft Graph
// v1.0 documentation verified 2026-10-23:
// https://learn.microsoft.com/en-us/graph/api/user-list-memberof?view=graph-rest-1.0

const memberContext = { feature: "directory.users.read", permission: "User.Read.All" } as const;

// One request per account is made for the first this many accounts; the rest are not asked, and
// the run says so. Accounts are asked for in the order the caller gives, most important first.
export const MAX_MEMBERSHIP_ACCOUNTS = 2_000;
// Requests in flight at once. A few at a time keeps a 2,000-account run well inside the audit's
// ten minutes without hammering the directory.
export const MEMBERSHIP_CONCURRENCY = 4;
const MAX_PAGES = 10;
const MAX_MEMBERSHIPS = 1_000;
const NAME_LENGTH = 200;

const text = z
  .string()
  .transform((value) => cleanText(value.slice(0, 500)))
  .nullish()
  .catch(null)
  .transform((value) => value ?? null);
const objectSchema = z.object({
  "@odata.type": text,
  id: z.string().min(1).max(200),
  displayName: text,
});

export interface Membership {
  groups: { id: string; name: string }[];
  roles: string[];
  units: number;
  // Memberships that were not read: a type Prism365 does not know, or an unreadable record, or
  // more memberships than one account is read for.
  unread: number;
}

export interface MembershipRead {
  byAccount: Map<string, Membership>;
  // Accounts asked for that got no answer.
  unanswered: number;
  // Accounts beyond the request limit, which were not asked.
  notAsked: number;
  // Accounts whose memberships were cut short.
  partial: number;
  // Why reading stopped before every account was asked, if it did.
  stopped: "throttled" | "refused" | null;
  status: number | null;
}

function kindOf(type: string | null): "group" | "role" | "unit" | null {
  const name = type?.toLowerCase().replace(/^#?microsoft\.graph\./, "");
  if (name === "group") return "group";
  if (name === "directoryrole") return "role";
  if (name === "administrativeunit") return "unit";
  return null;
}

function membershipOf(items: readonly unknown[], complete: boolean): Membership {
  const membership: Membership = { groups: [], roles: [], units: 0, unread: complete ? 0 : 1 };
  for (const raw of items) {
    const parsed = objectSchema.safeParse(raw);
    const kind = parsed.success ? kindOf(parsed.data["@odata.type"]) : null;
    if (!parsed.success || kind === null) {
      membership.unread += 1;
      continue;
    }
    const name = (parsed.data.displayName ?? "").slice(0, NAME_LENGTH);
    if (kind === "group")
      membership.groups.push({ id: parsed.data.id, name: name || "Unnamed group" });
    else if (kind === "role") membership.roles.push(name || "Unnamed role");
    else membership.units += 1;
  }
  return membership;
}

// Asks for each account's direct memberships, a few at a time. A 404 means the account is gone
// since the list was read and costs only that account. Throttling, or a refusal for any other
// reason, ends the reading and keeps what was read: the same refusal would meet every account.
// A session that ended is not swallowed; the runner explains it.
export async function readMemberships(
  context: AuditContext,
  accounts: readonly DirectoryUser[],
): Promise<MembershipRead> {
  const asked = accounts.slice(0, MAX_MEMBERSHIP_ACCOUNTS);
  const result: MembershipRead = {
    byAccount: new Map(),
    unanswered: 0,
    notAsked: accounts.length - asked.length,
    partial: 0,
    stopped: null,
    status: null,
  };
  let next = 0;
  let halted = false;
  const worker = async () => {
    while (!halted) {
      context.signal.throwIfAborted();
      const user = asked[next];
      next += 1;
      if (!user) return;
      try {
        const read = await context.graph.getCollection(
          `users/${encodeURIComponent(user.id)}/memberOf?$select=id,displayName`,
          memberContext,
          { schema: z.unknown(), maxPages: MAX_PAGES, maxRecords: MAX_MEMBERSHIPS },
        );
        const membership = membershipOf(read.items, read.complete);
        if (membership.unread > 0) result.partial += 1;
        result.byAccount.set(user.id, membership);
      } catch (error) {
        if (context.signal.aborted) throw error;
        if (error instanceof GraphRetryDeferredError) {
          result.stopped ??= "throttled";
          halted = true;
        } else if (error instanceof GraphRequestError && error.status === 404) {
          // Gone since the directory was read; nothing to say about it.
        } else if (error instanceof GraphRequestError && error.status !== 401) {
          result.stopped ??= "refused";
          result.status ??= error.status;
          halted = true;
        } else if (error instanceof GraphBoundaryError && error.reason !== "credential") {
          // One answer Prism365 could not read safely; the account is left without a result.
        } else {
          halted = true;
          throw error;
        }
      }
    }
  };
  await Promise.all(Array.from({ length: MEMBERSHIP_CONCURRENCY }, worker));
  result.unanswered = asked.length - result.byAccount.size;
  return result;
}

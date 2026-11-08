import {
  investigationLicenseSchema,
  type PersonDirectory,
  personDirectorySchema,
} from "@prism365/contracts";
import { z } from "zod";
import { type GraphHttpClient, GraphRequestError } from "./client.ts";

// What Microsoft's directory says about one account right now, for its profile page. Each part
// is read on its own: a missing permission or a slow call empties that part, never the page.
// Microsoft Graph v1.0 documentation verified 2026-10-22:
// https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/user-list-manager?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/user-list-directreports?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/resources/signinactivity?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/user-list-licensedetails?view=graph-rest-1.0

const directoryContext = { feature: "directory.users.read", permission: "User.Read.All" } as const;
const licenseContext = {
  feature: "user.licenses.read",
  permission: "LicenseAssignment.Read.All",
} as const;

const text = z
  .string()
  .transform((value) => value.slice(0, 500))
  .nullable()
  .optional()
  .transform((value) => value ?? null);
const flag = z
  .boolean()
  .nullable()
  .optional()
  .transform((value) => value ?? null);
const time = z
  .string()
  .datetime({ offset: true })
  .nullable()
  .optional()
  .catch(null)
  .transform((value) => value ?? null);

const profileFields = [
  "id",
  "displayName",
  "userPrincipalName",
  "mail",
  "jobTitle",
  "department",
  "companyName",
  "officeLocation",
  "employeeId",
  "userType",
  "accountEnabled",
  "createdDateTime",
  "lastPasswordChangeDateTime",
  "onPremisesSyncEnabled",
  "usageLocation",
  "city",
  "state",
  "country",
  "mobilePhone",
  "businessPhones",
].join(",");

const providerProfile = z.object({
  id: z.string().min(1).max(500),
  displayName: text,
  userPrincipalName: z.string().min(1).max(254),
  mail: text,
  jobTitle: text,
  department: text,
  companyName: text,
  officeLocation: text,
  employeeId: text,
  userType: text,
  accountEnabled: flag,
  createdDateTime: time,
  lastPasswordChangeDateTime: time,
  onPremisesSyncEnabled: flag,
  usageLocation: text,
  city: text,
  state: text,
  country: text,
  mobilePhone: text,
  businessPhones: z
    .array(z.string().max(500))
    .nullable()
    .optional()
    .catch(null)
    .transform((value) => (value ?? []).slice(0, 20)),
});

// The manager can be an organizational contact. The kind rule is the one the direct reports use:
// Microsoft names it in @odata.type, and without one an entry is a user only if it has a
// user principal name.
const isUserEntry = (kind: string | undefined, principal: string | null): boolean =>
  kind ? kind.toLowerCase() === "#microsoft.graph.user" : principal !== null;

const providerManager = z.object({
  "@odata.type": z.string().max(200).optional(),
  id: z.string().min(1).max(500),
  displayName: text,
  userPrincipalName: text,
  jobTitle: text,
});

// directReports can hold contacts as well as users. Microsoft names the kind in @odata.type; when
// it is missing, an entry counts as a user only if it carries a user principal name.
const providerReport = z.object({
  "@odata.type": z.string().max(200).optional(),
  id: z.string().min(1).max(500),
  displayName: text,
  userPrincipalName: text,
  jobTitle: text,
});

const reportsPageSize = 50;

// The first page of direct reports. `value` is taken whole and cut in code, so a page longer than
// asked for is capped rather than refused.
const providerReportsPage = z.object({
  value: z.array(z.unknown()),
  "@odata.nextLink": z.string().max(16_384).optional(),
});

const providerActivity = z.object({
  signInActivity: z
    .object({
      lastSignInDateTime: time,
      lastNonInteractiveSignInDateTime: time,
      lastSuccessfulSignInDateTime: time,
    })
    .nullable()
    .optional(),
});

export type PersonKey = { userId: string } | { userPrincipalName: string };

// Microsoft does not allow "/", "\" or "?" in a user principal name, and the Graph client refuses a
// user path that encodes them, so a name holding one is answered as not found without a request.
export function principalNameCanExist(name: string): boolean {
  return !/[/\\?]/.test(name);
}

function userPath(key: PersonKey): string {
  if ("userId" in key) return `users/${encodeURIComponent(key.userId)}`;
  // The OData key form also handles principal names that start with '$'.
  const escaped = encodeURIComponent(key.userPrincipalName.replaceAll("'", "''")).replaceAll(
    "'",
    "%27",
  );
  return `users('${escaped}')`;
}

export async function readPersonDirectory(options: {
  key: PersonKey;
  client: GraphHttpClient;
  grantedScopes: readonly string[];
}): Promise<PersonDirectory> {
  const has = (scope: string) => options.grantedScopes.includes(scope);
  const empty = {
    profile: null,
    manager: null,
    directReports: null,
    signInActivity: null,
    licenses: null,
  };
  if (!has("User.Read.All"))
    return personDirectorySchema.parse({
      state: "missingConsent",
      message: "Directory details need the User.Read.All feature pack.",
      ...empty,
    });

  if ("userPrincipalName" in options.key && !principalNameCanExist(options.key.userPrincipalName))
    return personDirectorySchema.parse({
      state: "notFound",
      message: "This account is not in the directory any more. Its stored sign-ins remain.",
      ...empty,
    });
  let profile: z.infer<typeof providerProfile>;
  try {
    profile = await options.client.request(
      `${userPath(options.key)}?$select=${profileFields}`,
      directoryContext,
      {},
      providerProfile,
    );
  } catch (error) {
    if (error instanceof GraphRequestError && error.status === 404)
      return personDirectorySchema.parse({
        state: "notFound",
        message: "This account is not in the directory any more. Its stored sign-ins remain.",
        ...empty,
      });
    return personDirectorySchema.parse({
      state: "unavailable",
      message: "Microsoft's directory could not be read just now. Refresh to try again.",
      ...empty,
    });
  }
  const byId = `users/${encodeURIComponent(profile.id)}`;

  const [manager, directReports, activity, licenses] = await Promise.all([
    options.client
      .request(
        `${byId}/manager?$select=id,displayName,userPrincipalName,jobTitle`,
        directoryContext,
        {},
        providerManager,
      )
      .then(({ "@odata.type": kind, ...person }) => ({
        state: "available" as const,
        data: { ...person, isUser: isUserEntry(kind, person.userPrincipalName) },
      }))
      // 404 means no manager is set.
      .catch((error: unknown) =>
        error instanceof GraphRequestError && error.status === 404
          ? { state: "available" as const, data: null }
          : { state: "unavailable" as const, data: null },
      ),
    // One page only. The answer is read with a single request, never with getCollection, because
    // that validates a next link, and Microsoft may give one in the directoryObjects form that the
    // pack does not allow. The link is only noted, never resolved or followed.
    options.client
      .request(
        `${byId}/directReports?$select=id,displayName,userPrincipalName,jobTitle&$top=${reportsPageSize}`,
        directoryContext,
        {},
        providerReportsPage,
      )
      .then((page) => {
        const people: z.infer<typeof providerReport>[] = [];
        let skipped = 0;
        for (const raw of page.value.slice(0, reportsPageSize)) {
          const parsed = providerReport.safeParse(raw);
          if (
            parsed.success &&
            isUserEntry(parsed.data["@odata.type"], parsed.data.userPrincipalName)
          )
            people.push(parsed.data);
          else skipped += 1;
        }
        return {
          state: "available" as const,
          more: page["@odata.nextLink"] !== undefined || page.value.length > reportsPageSize,
          skipped,
          data: people.map(({ "@odata.type": _kind, ...person }) => person),
        };
      })
      // Microsoft answers an empty collection when there are no reports, so a 404 means the
      // account was not found, not that it has none. 403 means Microsoft refused this read for the
      // signed-in administrator; any other failure leaves the list unread.
      .catch((error: unknown) => {
        const status = error instanceof GraphRequestError ? error.status : null;
        const none = { more: false, skipped: 0, data: [] };
        if (status === 403) return { state: "missingConsent" as const, ...none };
        return { state: "unavailable" as const, ...none };
      }),
    has("AuditLog.Read.All")
      ? options.client
          .request(`${byId}?$select=signInActivity`, directoryContext, {}, providerActivity)
          .then((data) => ({
            state: "available" as const,
            lastSignIn: data.signInActivity?.lastSignInDateTime ?? null,
            lastNonInteractiveSignIn: data.signInActivity?.lastNonInteractiveSignInDateTime ?? null,
            lastSuccessfulSignIn: data.signInActivity?.lastSuccessfulSignInDateTime ?? null,
          }))
          .catch(() => ({
            state: "unavailable" as const,
            lastSignIn: null,
            lastNonInteractiveSignIn: null,
            lastSuccessfulSignIn: null,
          }))
      : Promise.resolve({
          state: "missingConsent" as const,
          lastSignIn: null,
          lastNonInteractiveSignIn: null,
          lastSuccessfulSignIn: null,
        }),
    has("LicenseAssignment.Read.All")
      ? options.client
          .getCollection(
            `${byId}/licenseDetails?$select=id,skuId,skuPartNumber,servicePlans`,
            licenseContext,
            { schema: investigationLicenseSchema, maxPages: 4, maxRecords: 100 },
          )
          .then((result) => ({
            state: "available" as const,
            ...(result.complete
              ? {}
              : { message: "More licenses exist than this page reads; the list is incomplete." }),
            data: result.items,
          }))
          .catch(() => ({
            state: "unavailable" as const,
            message: "Licenses could not be read just now.",
            data: [],
          }))
      : Promise.resolve({
          state: "missingConsent" as const,
          message: "Licenses need the user.licenses.read feature pack.",
          data: [],
        }),
  ]);

  return personDirectorySchema.parse({
    state: "available",
    profile,
    manager,
    directReports,
    signInActivity: activity,
    licenses,
  });
}

// Resolves a principal name to the directory object ID that stored sign-ins are keyed on.
export async function lookupPersonId(options: {
  userPrincipalName: string;
  client: GraphHttpClient;
}): Promise<string | null> {
  if (!principalNameCanExist(options.userPrincipalName)) return null;
  try {
    const user = await options.client.request(
      `${userPath({ userPrincipalName: options.userPrincipalName })}?$select=id`,
      directoryContext,
      {},
      z.object({ id: z.string().min(1).max(200) }),
    );
    return user.id;
  } catch (error) {
    if (error instanceof GraphRequestError && error.status === 404) return null;
    throw error;
  }
}

import { type PersonDirectory, personDirectorySchema } from "@prism365/contracts";
import { investigationLicenseSchema } from "@prism365/contracts";
import { z } from "zod";
import { type GraphHttpClient, GraphRequestError } from "./client.ts";

// What Microsoft's directory says about one account right now, for its profile page. Each part
// is read on its own: a missing permission or a slow call empties that part, never the page.
// Microsoft Graph v1.0 documentation verified 2026-10-22:
// https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/user-list-manager?view=graph-rest-1.0
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

const providerManager = z.object({
  id: z.string().min(1).max(500),
  displayName: text,
  userPrincipalName: text,
  jobTitle: text,
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
  const empty = { profile: null, manager: null, signInActivity: null, licenses: null };
  if (!has("User.Read.All"))
    return personDirectorySchema.parse({
      state: "missingConsent",
      message: "Directory details need the User.Read.All feature pack.",
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

  const [manager, activity, licenses] = await Promise.all([
    options.client
      .request(
        `${byId}/manager?$select=id,displayName,userPrincipalName,jobTitle`,
        directoryContext,
        {},
        providerManager,
      )
      .then((data) => ({ state: "available" as const, data }))
      // 404 means no manager is set.
      .catch((error: unknown) =>
        error instanceof GraphRequestError && error.status === 404
          ? { state: "available" as const, data: null }
          : { state: "unavailable" as const, data: null },
      ),
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
    signInActivity: activity,
    licenses,
  });
}

// Resolves a principal name to the directory object ID that stored sign-ins are keyed on.
export async function lookupPersonId(options: {
  userPrincipalName: string;
  client: GraphHttpClient;
}): Promise<string | null> {
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

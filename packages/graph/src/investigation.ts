import {
  investigationEmailSchema,
  investigationLicenseSchema,
  investigationProfileSchema,
  investigationSignInSchema,
  type UserInvestigation,
  userInvestigationSchema,
} from "@prism365/contracts";
import { z } from "zod";
import { GraphBoundaryError, type GraphHttpClient, GraphRequestError } from "./client.ts";

// Microsoft Graph v1.0 documentation verified 2026-10-13:
// https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/user-list-licensedetails?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0
// This adapter accepts only a server-owned, tenant-bound delegated client.

const profileContext = { feature: "directory.users.read", permission: "User.Read.All" } as const;
const licenseContext = {
  feature: "user.licenses.read",
  permission: "LicenseAssignment.Read.All",
} as const;
const signInContext = { feature: "audit.signins.read", permission: "AuditLog.Read.All" } as const;
const nullableText = z.string().max(500).nullable().optional();

const providerProfileSchema = z
  .object({
    id: z.string().min(1).max(500),
    displayName: nullableText,
    userPrincipalName: investigationEmailSchema,
    mail: investigationEmailSchema.nullable().optional(),
    jobTitle: nullableText,
    department: nullableText,
    userType: nullableText,
  })
  .transform((profile) =>
    investigationProfileSchema.parse({
      ...profile,
      displayName: profile.displayName || "No display name",
      mail: profile.mail ?? null,
      jobTitle: profile.jobTitle ?? null,
      department: profile.department ?? null,
      userType: profile.userType ?? null,
    }),
  );

const providerSignInSchema = z.object({
  id: z.string().min(1).max(500),
  userId: z.string().min(1).max(500),
  createdDateTime: z.string().datetime({ offset: true }),
  appDisplayName: nullableText,
  clientAppUsed: nullableText,
  isInteractive: z.boolean().nullable().optional(),
  status: z.object({
    errorCode: z.number().int().nonnegative(),
    failureReason: nullableText,
    additionalDetails: nullableText,
  }),
  location: z.object({ city: nullableText, countryOrRegion: nullableText }).nullable().optional(),
});

export type UserInvestigationErrorCode =
  | "user_not_found"
  | "access_denied"
  | "missing_consent"
  | "provider_unavailable"
  | "invalid_input";

export class UserInvestigationError extends Error {
  readonly code: UserInvestigationErrorCode;
  readonly status: number;
  constructor(code: UserInvestigationErrorCode, status: number, message: string) {
    super(message);
    this.name = "UserInvestigationError";
    this.code = code;
    this.status = status;
  }
}

function validatedEmail(email: string): string {
  const result = investigationEmailSchema.safeParse(email);
  if (!result.success)
    throw new UserInvestigationError(
      "invalid_input",
      400,
      "Enter a valid user principal name (work email).",
    );
  return result.data;
}

function timeWindow(now: Date) {
  if (!Number.isFinite(now.getTime()))
    throw new UserInvestigationError("invalid_input", 400, "The investigation time is invalid.");
  return { start: new Date(now.getTime() - 7 * 86_400_000).toISOString(), end: now.toISOString() };
}

function sectionFailure(error: unknown): {
  state: "unknown" | "providerUnavailable";
  message: string;
} {
  if (error instanceof GraphRequestError && error.status === 403) {
    return {
      state: "unknown",
      message:
        "Microsoft denied access. Check administrator role, consent, feature licensing, and tenant restrictions; a 403 alone cannot identify the cause.",
    };
  }
  return {
    state: "providerUnavailable",
    message:
      "This section could not be read from Microsoft. Retry later; no missing data has been treated as inactivity.",
  };
}

const baseLimitations = [
  "Sign-in records are security evidence and cannot measure employee productivity, performance, or intent.",
  "The selected seven-day window is subject to Microsoft retention, licensing, collection delay, and event coverage; an empty result does not prove inactivity.",
  "IP addresses and Conditional Access policy details are omitted. Location fields are provider-reported context, not verified physical location.",
];

export async function investigateUser(options: {
  email: string;
  client: GraphHttpClient;
  grantedScopes?: readonly string[];
  now?: Date;
}): Promise<UserInvestigation> {
  const email = validatedEmail(options.email);
  const now = options.now ?? new Date();
  const window = timeWindow(now);
  const hasScope = (scope: string) =>
    options.grantedScopes === undefined || options.grantedScopes.includes(scope);
  if (!hasScope("User.Read.All"))
    throw new UserInvestigationError(
      "missing_consent",
      403,
      "User inventory requires the enabled User.Read.All feature pack.",
    );

  let profile: UserInvestigation["profile"]["data"];
  try {
    // OData function form also handles UPNs starting with '$'; quote escaping precedes URL encoding.
    const escapedUpn = encodeURIComponent(email.replaceAll("'", "''")).replaceAll("'", "%27");
    profile = await options.client.request(
      `users('${escapedUpn}')?$select=id,displayName,userPrincipalName,mail,jobTitle,department,userType`,
      profileContext,
      {},
      providerProfileSchema,
    );
    if (profile.userPrincipalName.toLowerCase() !== email.toLowerCase()) {
      throw new GraphBoundaryError("Microsoft returned a different user principal name.");
    }
  } catch (error) {
    if (error instanceof GraphRequestError && error.status === 404)
      throw new UserInvestigationError(
        "user_not_found",
        404,
        "No user was found for this principal name in the connected tenant.",
      );
    if (error instanceof GraphRequestError && error.status === 403)
      throw new UserInvestigationError(
        "access_denied",
        403,
        "Microsoft denied the user lookup. Check consent, administrator privileges, and tenant restrictions.",
      );
    throw new UserInvestigationError(
      "provider_unavailable",
      503,
      "Microsoft user lookup is temporarily unavailable or returned an invalid response.",
    );
  }
  const user = profile;
  const userPath = `users/${encodeURIComponent(user.id)}`;

  const readLicenses = async (): Promise<UserInvestigation["licenses"]> => {
    if (!hasScope("LicenseAssignment.Read.All"))
      return {
        state: "missingConsent",
        message:
          "Enable the user license feature pack (LicenseAssignment.Read.All) for this section.",
        data: [],
      };
    try {
      const result = await options.client.getCollection(
        `${userPath}/licenseDetails?$select=id,skuId,skuPartNumber,servicePlans`,
        licenseContext,
        { schema: investigationLicenseSchema, maxPages: 4, maxRecords: 100 },
      );
      if (!result.complete)
        return {
          state: "unknown",
          message: "The license response exceeded the bounded preview; completeness is unknown.",
          data: result.items,
        };
      return { state: "available", data: result.items };
    } catch (error) {
      return { ...sectionFailure(error), data: [] };
    }
  };
  const readSignIns = async (): Promise<UserInvestigation["signIns"]> => {
    if (!hasScope("AuditLog.Read.All"))
      return {
        state: "missingConsent",
        message: "Enable the sign-in feature pack (AuditLog.Read.All) for this section.",
        data: [],
        window,
        complete: false,
      };
    try {
      const escapedId = user.id.replaceAll("'", "''");
      const query = new URLSearchParams({
        $filter: `userId eq '${escapedId}' and createdDateTime ge ${window.start} and createdDateTime le ${window.end}`,
        $top: "50",
      });
      const result = await options.client.getCollection(
        `auditLogs/signIns?${query}`,
        signInContext,
        { schema: providerSignInSchema, maxPages: 4, maxRecords: 100 },
      );
      const data = result.items
        .map((entry) => {
          const timestamp = Date.parse(entry.createdDateTime);
          if (
            entry.userId !== user.id ||
            timestamp < Date.parse(window.start) ||
            timestamp > Date.parse(window.end)
          )
            throw new GraphBoundaryError(
              "Provider records escaped the investigated user or time window.",
            );
          return investigationSignInSchema.parse({
            id: entry.id,
            createdDateTime: entry.createdDateTime,
            appDisplayName: entry.appDisplayName ?? "Unknown application",
            clientAppUsed: entry.clientAppUsed ?? "Unknown client",
            isInteractive: entry.isInteractive ?? null,
            status: {
              errorCode: entry.status.errorCode,
              failureReason: entry.status.failureReason ?? null,
              additionalDetails: entry.status.additionalDetails ?? null,
            },
            location: {
              city: entry.location?.city ?? null,
              countryOrRegion: entry.location?.countryOrRegion ?? null,
            },
          });
        })
        .sort((a, b) => Date.parse(b.createdDateTime) - Date.parse(a.createdDateTime));
      return {
        state: "available",
        ...(result.complete
          ? {}
          : {
              message: "Showing at most 100 sign-ins from four pages; additional events may exist.",
            }),
        data,
        window,
        complete: result.complete,
      };
    } catch (error) {
      return { ...sectionFailure(error), data: [], window, complete: false };
    }
  };
  const [licenses, signIns] = await Promise.all([readLicenses(), readSignIns()]);
  return userInvestigationSchema.parse({
    mode: "live",
    generatedAt: now.toISOString(),
    targetEmail: email,
    profile: { state: "available", data: user },
    licenses,
    signIns,
    limitations: [
      ...baseLimitations,
      "License details require delegated access and an appropriate Entra role; background application access is not supported for this endpoint.",
    ],
  });
}

export function createDemoUserInvestigation(options: {
  email: string;
  now?: Date;
}): UserInvestigation {
  const email = validatedEmail(options.email);
  const now = options.now ?? new Date();
  const window = timeWindow(now);
  return userInvestigationSchema.parse({
    mode: "demo",
    generatedAt: now.toISOString(),
    targetEmail: email,
    profile: {
      state: "available",
      data: {
        id: "synthetic-user",
        displayName: "Alex Rivera (demo)",
        userPrincipalName: email,
        mail: email,
        jobTitle: "Systems Engineer",
        department: "Technology",
        userType: "Member",
      },
    },
    licenses: {
      state: "available",
      data: [
        {
          id: "synthetic-license",
          skuId: "synthetic-sku",
          skuPartNumber: "DEMO_M365",
          servicePlans: [
            {
              servicePlanId: "synthetic-plan",
              servicePlanName: "DEMO_EXCHANGE",
              provisioningStatus: "Success",
            },
          ],
        },
      ],
    },
    signIns: {
      state: "available",
      window,
      complete: true,
      data: Array.from({ length: 5 }, (_, index) => ({
        id: `synthetic-signin-${index + 1}`,
        createdDateTime: new Date(now.getTime() - (index + 1) * 3_600_000).toISOString(),
        appDisplayName: index % 2 ? "Microsoft Teams (demo)" : "Microsoft 365 (demo)",
        clientAppUsed: "Browser",
        isInteractive: true,
        status: {
          errorCode: index === 2 ? 500121 : 0,
          failureReason: index === 2 ? "Synthetic failed MFA attempt" : null,
          additionalDetails: null,
        },
        location: { city: "Demo city", countryOrRegion: "Demo region" },
      })),
    },
    limitations: [
      "All records in this view are synthetic. No tenant or account has been queried.",
      ...baseLimitations,
    ],
  });
}

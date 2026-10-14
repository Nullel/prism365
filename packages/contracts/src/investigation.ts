import { z } from "zod";

// HTML email syntax accepts valid UPN local parts such as '$user' and '#EXT#' guests.
export const investigationEmailSchema = z
  .string()
  .trim()
  .max(254)
  .pipe(z.email({ pattern: z.regexes.html5Email }));

const text = z.string().max(500);
const nullableText = text.nullable();
const sectionState = z.enum([
  "available",
  "missingConsent",
  "missingRole",
  "missingLicense",
  "unsupportedCloud",
  "notEnabled",
  "notImplemented",
  "providerUnavailable",
  "unknown",
]);

export const investigationProfileSchema = z.object({
  id: text.min(1),
  displayName: text.min(1),
  userPrincipalName: investigationEmailSchema,
  mail: investigationEmailSchema.nullable(),
  jobTitle: nullableText,
  department: nullableText,
  userType: nullableText,
});

export const investigationLicenseSchema = z.object({
  id: text.min(1),
  skuId: text.min(1),
  skuPartNumber: text.min(1),
  servicePlans: z
    .array(
      z.object({
        servicePlanId: text.min(1),
        servicePlanName: text.min(1),
        provisioningStatus: text.min(1),
      }),
    )
    .max(500),
});

export const investigationSignInSchema = z.object({
  id: text.min(1),
  createdDateTime: z.string().datetime({ offset: true }),
  appDisplayName: text,
  clientAppUsed: text,
  isInteractive: z.boolean().nullable(),
  status: z.object({
    errorCode: z.number().int().nonnegative(),
    failureReason: nullableText,
    additionalDetails: nullableText,
  }),
  location: z.object({ city: nullableText, countryOrRegion: nullableText }),
});

export const userInvestigationSchema = z.object({
  mode: z.enum(["demo", "live"]),
  generatedAt: z.string().datetime(),
  targetEmail: investigationEmailSchema,
  profile: z.object({
    state: sectionState,
    message: z.string().max(1_000).optional(),
    data: investigationProfileSchema.nullable(),
  }),
  licenses: z.object({
    state: sectionState,
    message: z.string().max(1_000).optional(),
    data: z.array(investigationLicenseSchema).max(100),
  }),
  signIns: z.object({
    state: sectionState,
    message: z.string().max(1_000).optional(),
    data: z.array(investigationSignInSchema).max(100),
    window: z.object({ start: z.string().datetime(), end: z.string().datetime() }),
    complete: z.boolean(),
  }),
  limitations: z.array(z.string().max(1_000)).max(30),
});

export type UserInvestigation = z.infer<typeof userInvestigationSchema>;
export type InvestigationProfile = z.infer<typeof investigationProfileSchema>;
export type InvestigationLicense = z.infer<typeof investigationLicenseSchema>;
export type InvestigationSignIn = z.infer<typeof investigationSignInSchema>;

import { z } from "zod";
import { investigationLicenseSchema } from "./investigation.ts";

// One account's page: what Prism365 has stored from its sign-ins, plus what Microsoft's
// directory says about it now. Full addresses and device details are shown here on purpose:
// this page exists so an administrator can tell whether an account is compromised.

// Provider text is cut to display length rather than rejected: one long failure reason or job
// title must never turn a whole page into an error.
const text = z.string().transform((value) => value.slice(0, 500));
const requiredText = z
  .string()
  .min(1)
  .transform((value) => value.slice(0, 500));
const nullableText = text.nullable();
const time = z.string().datetime({ offset: true });
const count = z.number().int().nonnegative();

export const personIdSchema = z
  .string()
  .min(1)
  .max(200)
  .regex(/^[A-Za-z0-9._@:-]+$/);

export const personListItemSchema = z.object({
  userId: personIdSchema,
  userPrincipalName: nullableText,
  displayName: nullableText,
  lastSignIn: time,
  lastSuccess: time.nullable(),
  signIns: count,
  failures: count,
  countries: count,
  openFindings: count,
});
export type PersonListItem = z.infer<typeof personListItemSchema>;

export const personListSchema = z.object({
  mode: z.enum(["demo", "live"]),
  days: z.number().int().min(1).max(90),
  items: z.array(personListItemSchema),
  truncated: z.boolean(),
});
export type PersonList = z.infer<typeof personListSchema>;

const span = { count, successes: count, failures: count, firstSeen: time, lastSeen: time };

export const personHistorySchema = z.object({
  userId: personIdSchema,
  userPrincipalName: nullableText,
  displayName: nullableText,
  totals: z.object({
    signIns: count,
    successes: count,
    failures: count,
    blocked: count,
    firstSeen: time.nullable(),
    lastSeen: time.nullable(),
    lastSuccess: time.nullable(),
  }),
  places: z.array(
    z.object({ country: nullableText, state: nullableText, city: nullableText, ...span }),
  ),
  addresses: z.array(
    z.object({
      ip: text,
      country: nullableText,
      state: nullableText,
      city: nullableText,
      ...span,
    }),
  ),
  devices: z.array(
    z.object({
      deviceId: nullableText,
      operatingSystem: nullableText,
      browser: nullableText,
      compliant: z.boolean().nullable(),
      managed: z.boolean().nullable(),
      trustType: nullableText,
      ...span,
    }),
  ),
  apps: z.array(z.object({ app: nullableText, ...span })),
  // What the sign-in audit has learned is normal for this account.
  baseline: z.array(
    z.object({
      kind: z.enum(["ip", "country", "state", "device"]),
      value: text,
      trusted: z.boolean(),
      firstSeen: time,
      lastSeen: time,
    }),
  ),
});
export type PersonHistory = z.infer<typeof personHistorySchema>;

export const personSigninSchema = z.object({
  id: text,
  createdAt: time,
  // success; failure: wrong password or account problem; mfaFailure: a failed or denied MFA
  // prompt; blocked: Conditional Access; interrupted: any other code, usually a normal step
  // such as an MFA prompt that the next record completes.
  result: z.enum(["success", "failure", "mfaFailure", "blocked", "interrupted"]),
  errorCode: z.number().int(),
  failureReason: nullableText,
  app: nullableText,
  resource: nullableText,
  clientApp: nullableText,
  ip: nullableText,
  city: nullableText,
  state: nullableText,
  country: nullableText,
  latitude: z.number().nullable(),
  longitude: z.number().nullable(),
  deviceId: nullableText,
  operatingSystem: nullableText,
  browser: nullableText,
  compliant: z.boolean().nullable(),
  managed: z.boolean().nullable(),
  trustType: nullableText,
  conditionalAccess: nullableText,
  riskLevel: nullableText,
  riskState: nullableText,
  riskEventTypes: z.array(text).transform((items) => items.slice(0, 50)),
  interactive: z.boolean().nullable(),
});
export type PersonSignin = z.infer<typeof personSigninSchema>;

export const personSigninPageSchema = z.object({
  items: z.array(personSigninSchema),
  // Pass back as `before` for the next, older page; null when there is none.
  next: z.string().max(600).nullable(),
});
export type PersonSigninPage = z.infer<typeof personSigninPageSchema>;

const directoryState = z.enum(["available", "missingConsent", "unavailable"]);

export const personDirectorySchema = z.object({
  state: z.enum(["available", "notFound", "missingConsent", "unavailable"]),
  message: z.string().max(500).optional(),
  profile: z
    .object({
      id: requiredText,
      displayName: nullableText,
      userPrincipalName: requiredText,
      mail: nullableText,
      jobTitle: nullableText,
      department: nullableText,
      companyName: nullableText,
      officeLocation: nullableText,
      employeeId: nullableText,
      userType: nullableText,
      accountEnabled: z.boolean().nullable(),
      createdDateTime: time.nullable(),
      lastPasswordChangeDateTime: time.nullable(),
      onPremisesSyncEnabled: z.boolean().nullable(),
      usageLocation: nullableText,
      city: nullableText,
      state: nullableText,
      country: nullableText,
      mobilePhone: nullableText,
      businessPhones: z.array(text).transform((items) => items.slice(0, 20)),
    })
    .nullable(),
  manager: z
    .object({
      state: directoryState,
      data: z
        .object({
          id: requiredText,
          displayName: nullableText,
          userPrincipalName: nullableText,
          jobTitle: nullableText,
        })
        .nullable(),
    })
    .nullable(),
  signInActivity: z
    .object({
      state: directoryState,
      lastSignIn: time.nullable(),
      lastNonInteractiveSignIn: time.nullable(),
      lastSuccessfulSignIn: time.nullable(),
    })
    .nullable(),
  licenses: z
    .object({
      state: directoryState,
      message: z.string().max(500).optional(),
      data: z.array(investigationLicenseSchema).max(100),
    })
    .nullable(),
});
export type PersonDirectory = z.infer<typeof personDirectorySchema>;

export const personLookupSchema = z.object({ userId: personIdSchema });

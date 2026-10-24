import type { MicrosoftCloud } from "./clouds.ts";

export type PermissionSource =
  | "microsoftGraph"
  | "officeManagementActivity"
  | "exchangeOnlinePowerShell";

interface PermissionPackDefinition {
  id: string;
  label: string;
  source: PermissionSource;
  permission: string;
  mode: "delegated" | "application" | "either";
  phase: "mvp" | "experimental" | "future";
  defaultEnabled: boolean;
  purpose: string;
  deniedBehavior: string;
  clouds: readonly MicrosoftCloud[];
  prerequisites: readonly string[];
  roleNotes: string;
  licenseNotes: string;
  documentationUrl: `https://${string}`;
  documentationLastVerified: `${number}-${number}-${number}`;
}

const allGraphClouds = [
  "worldwide",
  "gcc",
  "usGovernmentL4",
  "usGovernmentL5",
  "china",
] as const satisfies readonly MicrosoftCloud[];

const graphWithoutChina = [
  "worldwide",
  "gcc",
  "usGovernmentL4",
  "usGovernmentL5",
] as const satisfies readonly MicrosoftCloud[];

export const permissionPacks = [
  {
    id: "profile",
    label: "Signed-in profile",
    source: "microsoftGraph",
    permission: "User.Read",
    mode: "delegated",
    phase: "mvp",
    defaultEnabled: true,
    purpose: "Identify the signed-in administrator during an interactive session.",
    deniedBehavior: "Interactive sign-in is unavailable.",
    clouds: allGraphClouds,
    prerequisites: [],
    roleNotes: "No additional Entra role is normally required.",
    licenseNotes: "A Microsoft Entra work or school account is required.",
    documentationUrl: "https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-13",
  },
  {
    id: "directory.users.read",
    label: "User inventory",
    source: "microsoftGraph",
    permission: "User.Read.All",
    mode: "either",
    phase: "mvp",
    defaultEnabled: false,
    purpose: "Read selected directory user fields for account inventory and hygiene.",
    deniedBehavior: "User inventory is hidden; other reports remain available.",
    clouds: allGraphClouds,
    prerequisites: [],
    roleNotes: "Delegated access is also limited by the signed-in user's privileges.",
    licenseNotes: "Endpoint-specific Microsoft Entra licensing rules still apply.",
    documentationUrl: "https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-13",
  },
  {
    id: "user.licenses.read",
    label: "User license details",
    source: "microsoftGraph",
    permission: "LicenseAssignment.Read.All",
    mode: "delegated",
    phase: "mvp",
    defaultEnabled: false,
    purpose:
      "Read assigned product licenses and service-plan provisioning for one investigated user.",
    deniedBehavior:
      "User license details are unavailable; profile and sign-in sections remain independent.",
    clouds: allGraphClouds,
    prerequisites: ["Delegated administrator session"],
    roleNotes:
      "Directory Readers, Guest Inviter, License Administrator, User Administrator, or an appropriate custom role is required.",
    licenseNotes:
      "This endpoint does not support application permissions; returned details include direct and group assignments.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/user-list-licensedetails?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-13",
  },
  {
    id: "tenant.licenses.read",
    label: "Tenant license inventory",
    source: "microsoftGraph",
    permission: "LicenseAssignment.Read.All",
    mode: "either",
    phase: "mvp",
    defaultEnabled: false,
    purpose: "Explain which licensed capabilities can be evaluated in the tenant.",
    deniedBehavior: "License-dependent readiness remains unknown.",
    clouds: allGraphClouds,
    prerequisites: [],
    roleNotes: "Delegated callers need a supported directory role such as Directory Readers.",
    licenseNotes: "Returns purchased SKU inventory; it does not grant use of licensed features.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/subscribedsku-list?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-13",
  },
  {
    id: "service.health.read",
    label: "Service health",
    source: "microsoftGraph",
    permission: "ServiceHealth.Read.All",
    mode: "either",
    phase: "mvp",
    defaultEnabled: false,
    purpose: "Show current Microsoft 365 service health, incidents and advisories.",
    deniedBehavior: "The service-health panel says it cannot check, and why.",
    clouds: allGraphClouds,
    prerequisites: [],
    roleNotes:
      "Delegated callers also need an admin role that can view service health, such as Global Reader or Service Support Administrator.",
    licenseNotes: "Microsoft reports only the services the tenant subscribes to.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/serviceannouncement-list-issues?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-22",
  },
  {
    id: "service.messages.read",
    label: "Message center",
    source: "microsoftGraph",
    permission: "ServiceMessage.Read.All",
    mode: "either",
    phase: "mvp",
    defaultEnabled: false,
    purpose:
      "Read Message center posts: planned changes, retirements and actions Microsoft asks administrators to take.",
    deniedBehavior: "Message center items are left out of Admin attention.",
    clouds: allGraphClouds,
    prerequisites: [],
    roleNotes:
      "Delegated callers also need a role that can read Message center, such as Message Center Reader or Global Reader.",
    licenseNotes: "Microsoft lists posts for the services the tenant subscribes to.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/serviceannouncement-list-messages?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-22",
  },
  {
    id: "audit.signins.read",
    label: "Sign-in and directory audits",
    source: "microsoftGraph",
    permission: "AuditLog.Read.All",
    mode: "either",
    phase: "mvp",
    defaultEnabled: false,
    purpose: "Summarize sign-ins and directory changes for security operations.",
    deniedBehavior: "Audit and sign-in reports are unavailable.",
    clouds: allGraphClouds,
    prerequisites: [],
    roleNotes:
      "Delegated access requires a supported Entra role; Conditional Access detail needs policy read access.",
    licenseNotes: "Retention and sign-in fields vary with the tenant's Entra licensing.",
    documentationUrl: "https://learn.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-13",
  },
  {
    id: "reports.authentication-methods.read",
    label: "Authentication method registration",
    source: "microsoftGraph",
    permission: "AuditLog.Read.All",
    mode: "either",
    phase: "mvp",
    defaultEnabled: false,
    purpose:
      "Read which sign-in methods each account has registered, for the MFA registration audit.",
    deniedBehavior: "The MFA registration audit cannot run; other audits are unaffected.",
    clouds: allGraphClouds,
    prerequisites: [],
    roleNotes:
      "Delegated callers need Reports Reader, Security Reader, Security Administrator or Global Reader.",
    licenseNotes: "Microsoft's report leaves out disabled accounts.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/authenticationmethodsroot-list-userregistrationdetails?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-22",
  },
  {
    id: "identity.conditional-access.read",
    label: "Conditional Access policies",
    source: "microsoftGraph",
    permission: "Policy.Read.All",
    mode: "either",
    phase: "experimental",
    defaultEnabled: false,
    purpose: "Read Conditional Access policy definitions and related sign-in context.",
    deniedBehavior: "Policy details remain unavailable; sign-in summaries can still be shown.",
    clouds: allGraphClouds,
    prerequisites: ["audit.signins.read for sign-in correlation"],
    roleNotes:
      "Delegated callers require a supported role such as Security Reader or Global Reader.",
    licenseNotes: "Evaluating Conditional Access outcomes may require Entra ID P1 or P2 features.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/conditionalaccessroot-list-policies?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-13",
  },
  {
    id: "identity.risk.read",
    label: "Identity Protection risk",
    source: "microsoftGraph",
    permission: "IdentityRiskyUser.Read.All",
    mode: "either",
    phase: "experimental",
    defaultEnabled: false,
    purpose: "Read risky-user detections for identity-security investigations.",
    deniedBehavior:
      "Risk detections are unavailable; no risk score is inferred from other telemetry.",
    clouds: graphWithoutChina,
    prerequisites: [],
    roleNotes: "Delegated callers require a supported Entra role.",
    licenseNotes: "Microsoft Entra ID P2 is required for risky-user data.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/riskyuser-get?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-13",
  },
  {
    id: "reports.usage.read",
    label: "Microsoft 365 usage reports",
    source: "microsoftGraph",
    permission: "Reports.Read.All",
    mode: "either",
    phase: "mvp",
    defaultEnabled: false,
    purpose: "Show tenant-level service adoption and operational usage trends.",
    deniedBehavior: "Usage reports are unavailable.",
    clouds: ["worldwide", "gcc"],
    prerequisites: [],
    roleNotes: "Delegated callers require a supported reports or directory role.",
    licenseNotes: "Report availability depends on the tenant's Microsoft 365 subscriptions.",
    documentationUrl: "https://learn.microsoft.com/en-us/graph/reportroot-authorization",
    documentationLastVerified: "2026-10-13",
  },
  {
    id: "reports.settings.read",
    label: "Usage-report privacy setting",
    source: "microsoftGraph",
    permission: "ReportSettings.Read.All",
    mode: "either",
    phase: "mvp",
    defaultEnabled: false,
    purpose: "Respect whether the tenant conceals names in Microsoft 365 usage reports.",
    deniedBehavior: "Prism365 assumes identifiers must remain concealed.",
    clouds: ["worldwide", "gcc"],
    prerequisites: ["reports.usage.read"],
    roleNotes: "Delegated callers require a supported role.",
    licenseNotes: "This reads the tenant privacy setting; it does not change it.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/adminreportsettings-get?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-13",
  },
  {
    id: "security.score.read",
    label: "Microsoft Secure Score",
    source: "microsoftGraph",
    permission: "SecurityEvents.Read.All",
    mode: "either",
    phase: "experimental",
    defaultEnabled: false,
    purpose: "Read Secure Score snapshots and improvement actions.",
    deniedBehavior: "Secure Score views are unavailable.",
    clouds: graphWithoutChina,
    prerequisites: [],
    roleNotes: "Delegated callers require a supported security or global reader role.",
    licenseNotes: "Returned products and controls depend on tenant subscriptions.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/security-list-securescores?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-13",
  },
  {
    id: "security.incidents.read",
    label: "Microsoft Defender incidents",
    source: "microsoftGraph",
    permission: "SecurityIncident.Read.All",
    mode: "either",
    phase: "experimental",
    defaultEnabled: false,
    purpose: "Read Defender XDR incidents for security operations.",
    deniedBehavior: "Defender incident views are unavailable.",
    clouds: graphWithoutChina,
    prerequisites: [],
    roleNotes: "Delegated callers require a supported security role.",
    licenseNotes: "Requires an applicable Microsoft Defender product and entitlement.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/security-list-incidents?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-13",
  },
  {
    id: "devices.intune.read",
    label: "Intune managed devices",
    source: "microsoftGraph",
    permission: "DeviceManagementManagedDevices.Read.All",
    mode: "either",
    phase: "experimental",
    defaultEnabled: false,
    purpose: "Read normalized managed-device inventory and compliance state.",
    deniedBehavior: "Managed-device inventory is unavailable.",
    clouds: allGraphClouds,
    prerequisites: [],
    roleNotes: "Delegated callers also need an appropriate Intune role.",
    licenseNotes: "The tenant must have an active Intune license.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/intune-devices-manageddevice-list?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-13",
  },
  {
    id: "purview.entra.read",
    label: "Purview Entra audit search",
    source: "microsoftGraph",
    permission: "AuditLogsQuery-Entra.Read.All",
    mode: "either",
    phase: "experimental",
    defaultEnabled: false,
    purpose: "Run explicit Purview Audit queries for the Entra workload.",
    deniedBehavior: "Purview audit search is unavailable; Graph directory audits are separate.",
    clouds: ["worldwide", "gcc"],
    prerequisites: ["Microsoft Purview Audit must be available and enabled"],
    roleNotes: "The calling identity must satisfy Purview role and query requirements.",
    licenseNotes: "Audit retention and premium features depend on Microsoft 365 licensing.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/security-auditcoreroot-post-auditlogqueries?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-13",
  },
  {
    id: "audit.activity-feed.read",
    label: "Microsoft 365 Management Activity feed",
    source: "officeManagementActivity",
    permission: "ActivityFeed.Read",
    mode: "application",
    phase: "future",
    defaultEnabled: false,
    purpose: "Continuously ingest subscribed Microsoft 365 audit content blobs.",
    deniedBehavior: "Continuous activity-feed ingestion is unavailable.",
    clouds: ["worldwide", "gcc", "usGovernmentL4", "usGovernmentL5"],
    prerequisites: [
      "Unified audit logging must be enabled",
      "A content subscription must be started",
    ],
    roleNotes: "The access token tenant must match the tenant being queried.",
    licenseNotes:
      "Available audit workloads and retention depend on tenant services and licensing.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference",
    documentationLastVerified: "2026-10-13",
  },
  {
    id: "exchange.configuration.read",
    label: "Exchange Online configuration",
    source: "exchangeOnlinePowerShell",
    permission: "Exchange.ManageAsApp",
    mode: "application",
    phase: "future",
    defaultEnabled: false,
    purpose: "Read Exchange configuration that is not available through stable Graph endpoints.",
    deniedBehavior: "Exchange-specific configuration reports are unavailable.",
    clouds: allGraphClouds,
    prerequisites: ["Certificate or federated application credential", "Separate isolated worker"],
    roleNotes:
      "Exchange Online RBAC must restrict the service principal to the required cmdlets and scope.",
    licenseNotes: "Requires Exchange Online in the target tenant.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2",
    documentationLastVerified: "2026-10-13",
  },
] as const satisfies readonly PermissionPackDefinition[];

export type PermissionPack = (typeof permissionPacks)[number];
export type PermissionPackId = PermissionPack["id"];
export type GraphPermissionPack = Extract<PermissionPack, { source: "microsoftGraph" }>;
export type GraphFeatureId = GraphPermissionPack["id"];
export type GraphPermission = GraphPermissionPack["permission"];

export function getPermissionPack(id: PermissionPackId): PermissionPack {
  const pack = permissionPacks.find((candidate) => candidate.id === id);
  if (!pack) {
    throw new Error(`Unknown permission pack: ${id}`);
  }
  return pack;
}

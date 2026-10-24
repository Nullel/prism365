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
  // Other Microsoft pages for the pack's endpoints, when it reads more than one kind of resource.
  relatedDocumentationUrls?: readonly `https://${string}`[];
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
    purpose:
      "Read selected directory user fields for account inventory and hygiene, a user's manager, and a user's direct memberships (groups, directory roles and administrative units).",
    deniedBehavior: "User inventory is hidden; other reports remain available.",
    clouds: allGraphClouds,
    prerequisites: [],
    roleNotes:
      "Delegated access is also limited by the signed-in user's privileges. User.Read.All is Microsoft's documented least privilege for reading another user's memberships.",
    licenseNotes: "Endpoint-specific Microsoft Entra licensing rules still apply.",
    documentationUrl: "https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0",
    relatedDocumentationUrls: [
      "https://learn.microsoft.com/en-us/graph/api/user-list-manager?view=graph-rest-1.0",
      "https://learn.microsoft.com/en-us/graph/api/user-list-memberof?view=graph-rest-1.0",
    ],
    documentationLastVerified: "2026-10-23",
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
    label: "Sign-in logs",
    source: "microsoftGraph",
    permission: "AuditLog.Read.All",
    mode: "either",
    phase: "mvp",
    defaultEnabled: false,
    purpose: "Collect sign-ins for the sign-in audit and the People view.",
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
    id: "audit.directory.read",
    label: "Directory audit log",
    source: "microsoftGraph",
    permission: "AuditLog.Read.All",
    mode: "either",
    phase: "mvp",
    defaultEnabled: false,
    purpose:
      "Collect the directory audit log (who changed which user, group, role, app or policy) for the change feed.",
    deniedBehavior: "The directory change feed is unavailable; other features are unaffected.",
    clouds: allGraphClouds,
    prerequisites: [],
    roleNotes:
      "Delegated callers need a role such as Reports Reader, Security Reader or Global Reader.",
    licenseNotes:
      "Microsoft keeps directory audits 7 days without Entra ID P1 or P2 and 30 days with it.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/directoryaudit-list?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-23",
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
    id: "directory.applications.read",
    label: "Enterprise applications",
    source: "microsoftGraph",
    permission: "Application.Read.All",
    mode: "either",
    phase: "experimental",
    defaultEnabled: false,
    purpose:
      "Read enterprise applications and the application permissions granted to them, for the app consents audit.",
    deniedBehavior: "The app consents audit cannot run; other audits are unaffected.",
    clouds: allGraphClouds,
    prerequisites: [],
    roleNotes:
      "Delegated callers need a role that can read applications, such as Global Reader or Cloud Application Administrator.",
    licenseNotes: "No additional license is required.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/serviceprincipal-list?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-23",
  },
  {
    id: "directory.permission-grants.read",
    label: "Delegated permission grants",
    source: "microsoftGraph",
    permission: "Directory.Read.All",
    mode: "either",
    phase: "experimental",
    defaultEnabled: false,
    purpose:
      "Read which delegated permissions users and administrators consented to for each application.",
    deniedBehavior: "The app consents audit lists application permissions only.",
    clouds: allGraphClouds,
    prerequisites: ["directory.applications.read"],
    roleNotes:
      "Microsoft documents Directory.Read.All as the least-privileged permission for listing grants. Delegated callers need a role such as Global Reader.",
    licenseNotes: "No additional license is required.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/oauth2permissiongrant-list?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-23",
  },
  {
    id: "directory.roles.read",
    label: "Directory roles",
    source: "microsoftGraph",
    permission: "RoleManagement.Read.Directory",
    mode: "either",
    phase: "experimental",
    defaultEnabled: false,
    purpose:
      "Read directory role definitions and who holds each role, active or eligible, for the privileged roles audit.",
    deniedBehavior: "The privileged roles audit cannot run; other audits are unaffected.",
    clouds: allGraphClouds,
    prerequisites: [],
    roleNotes:
      "Delegated callers need Global Reader or Privileged Role Administrator for every endpoint; Directory Readers can read active assignments and definitions but not Privileged Identity Management schedules.",
    licenseNotes:
      "Eligible and time-bound assignments come from Privileged Identity Management, which needs Microsoft Entra ID P2 or ID Governance; without it only active assignments are listed.",
    documentationUrl:
      "https://learn.microsoft.com/en-us/graph/api/rbacapplication-list-roleassignments?view=graph-rest-1.0",
    relatedDocumentationUrls: [
      "https://learn.microsoft.com/en-us/graph/api/rbacapplication-list-roledefinitions?view=graph-rest-1.0",
      "https://learn.microsoft.com/en-us/graph/api/rbacapplication-list-roleeligibilityscheduleinstances?view=graph-rest-1.0",
      "https://learn.microsoft.com/en-us/graph/api/rbacapplication-list-roleassignmentscheduleinstances?view=graph-rest-1.0",
    ],
    documentationLastVerified: "2026-10-24",
  },
  {
    id: "directory.groups.read",
    label: "Groups and owners",
    source: "microsoftGraph",
    permission: "Group.Read.All",
    mode: "either",
    phase: "experimental",
    defaultEnabled: false,
    purpose:
      "Read groups, including the groups behind Microsoft Teams, and each group's owners, for the groups and Teams audit.",
    deniedBehavior: "The groups and Teams audit cannot run; other audits are unaffected.",
    clouds: allGraphClouds,
    prerequisites: [],
    roleNotes:
      "Delegated access is limited by the signed-in user's privileges; Global Reader or Directory Readers can read every group's owners.",
    licenseNotes: "No additional license is required.",
    documentationUrl: "https://learn.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0",
    relatedDocumentationUrls: [
      "https://learn.microsoft.com/en-us/graph/api/group-list-owners?view=graph-rest-1.0",
    ],
    documentationLastVerified: "2026-10-24",
  },
  {
    id: "directory.domains.read",
    label: "Domains",
    source: "microsoftGraph",
    permission: "Domain.Read.All",
    mode: "either",
    phase: "experimental",
    defaultEnabled: false,
    purpose:
      "Read each domain's password expiry policy, so the password audit can say when passwords expire.",
    deniedBehavior: "The password audit reports password age only, not expiry dates.",
    clouds: allGraphClouds,
    prerequisites: ["directory.users.read"],
    roleNotes: "Delegated callers need a role such as Global Reader or Directory Readers.",
    licenseNotes: "No additional license is required.",
    documentationUrl: "https://learn.microsoft.com/en-us/graph/api/domain-list?view=graph-rest-1.0",
    documentationLastVerified: "2026-10-24",
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

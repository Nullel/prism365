export const permissionPacks = [
  {
    id: "profile",
    label: "Signed-in profile",
    permission: "User.Read",
    mode: "delegated",
    phase: "mvp",
    defaultEnabled: true,
    purpose: "Identify the signed-in administrator during an interactive session.",
    deniedBehavior: "Interactive sign-in is unavailable.",
  },
  {
    id: "directory.users.read",
    label: "User inventory",
    permission: "User.Read.All",
    mode: "either",
    phase: "mvp",
    defaultEnabled: false,
    purpose: "Read selected directory user fields for account inventory and hygiene.",
    deniedBehavior: "User inventory is hidden; other reports remain available.",
  },
  {
    id: "audit.signins.read",
    label: "Sign-in and directory audits",
    permission: "AuditLog.Read.All",
    mode: "either",
    phase: "mvp",
    defaultEnabled: false,
    purpose: "Summarize sign-ins and directory changes for security operations.",
    deniedBehavior: "Audit and sign-in reports are unavailable.",
  },
  {
    id: "reports.usage.read",
    label: "Microsoft 365 usage reports",
    permission: "Reports.Read.All",
    mode: "either",
    phase: "mvp",
    defaultEnabled: false,
    purpose: "Show tenant-level service adoption and operational usage trends.",
    deniedBehavior: "Usage reports are unavailable.",
  },
  {
    id: "purview.entra.read",
    label: "Purview Entra audit search",
    permission: "AuditLogsQuery-Entra.Read.All",
    mode: "either",
    phase: "experimental",
    defaultEnabled: false,
    purpose: "Run explicit Purview Audit queries for the Entra workload.",
    deniedBehavior: "Purview audit search is unavailable; Graph directory audits are separate.",
  },
] as const;

export type PermissionPackId = (typeof permissionPacks)[number]["id"];
export type GraphPermission = (typeof permissionPacks)[number]["permission"];

export function getPermissionPack(id: PermissionPackId) {
  const pack = permissionPacks.find((candidate) => candidate.id === id);
  if (!pack) {
    throw new Error(`Unknown permission pack: ${id}`);
  }
  return pack;
}

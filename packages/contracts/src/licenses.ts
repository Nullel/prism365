// Microsoft's commercial names for common license part numbers. Anything else shows its part
// number. Source: Microsoft's "Product names and service plan identifiers for licensing".
const licenseNames: Record<string, string> = {
  AAD_PREMIUM: "Microsoft Entra ID P1",
  AAD_PREMIUM_P2: "Microsoft Entra ID P2",
  ATP_ENTERPRISE: "Microsoft Defender for Office 365 (Plan 1)",
  DESKLESSPACK: "Office 365 F3",
  EMS: "Enterprise Mobility + Security E3",
  EMSPREMIUM: "Enterprise Mobility + Security E5",
  ENTERPRISEPACK: "Office 365 E3",
  ENTERPRISEPREMIUM: "Office 365 E5",
  EXCHANGEDESKLESS: "Exchange Online Kiosk",
  EXCHANGEENTERPRISE: "Exchange Online (Plan 2)",
  EXCHANGESTANDARD: "Exchange Online (Plan 1)",
  FLOW_FREE: "Microsoft Power Automate Free",
  INTUNE_A: "Microsoft Intune Plan 1",
  MCOEV: "Microsoft Teams Phone Standard",
  MCOMEETADV: "Microsoft 365 Audio Conferencing",
  Microsoft_365_Copilot: "Microsoft 365 Copilot",
  O365_BUSINESS: "Microsoft 365 Apps for business",
  O365_BUSINESS_ESSENTIALS: "Microsoft 365 Business Basic",
  O365_BUSINESS_PREMIUM: "Microsoft 365 Business Standard",
  OFFICESUBSCRIPTION: "Microsoft 365 Apps for enterprise",
  POWER_BI_PRO: "Power BI Pro",
  POWER_BI_STANDARD: "Microsoft Fabric (Free)",
  PROJECTPROFESSIONAL: "Project Plan 3",
  SPB: "Microsoft 365 Business Premium",
  SPE_E3: "Microsoft 365 E3",
  SPE_E5: "Microsoft 365 E5",
  SPE_F1: "Microsoft 365 F3",
  STANDARDPACK: "Office 365 E1",
  STREAM: "Microsoft Stream",
  TEAMS_EXPLORATORY: "Microsoft Teams Exploratory",
  VISIOCLIENT: "Visio Plan 2",
  WINDOWS_STORE: "Windows Store for Business",
};

export function licenseName(partNumber: string): string {
  return Object.hasOwn(licenseNames, partNumber)
    ? (licenseNames[partNumber] ?? partNumber)
    : partNumber;
}

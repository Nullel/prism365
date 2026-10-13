export const microsoftCloudProfiles = {
  worldwide: {
    id: "worldwide",
    label: "Microsoft global cloud",
    authorityHost: "https://login.microsoftonline.com",
    graphBaseUrl: "https://graph.microsoft.com/v1.0",
    managementActivityBaseUrl: "https://manage.office.com",
  },
  gcc: {
    id: "gcc",
    label: "Microsoft 365 GCC",
    authorityHost: "https://login.microsoftonline.com",
    graphBaseUrl: "https://graph.microsoft.com/v1.0",
    managementActivityBaseUrl: "https://manage-gcc.office.com",
  },
  usGovernmentL4: {
    id: "usGovernmentL4",
    label: "Microsoft 365 GCC High",
    authorityHost: "https://login.microsoftonline.us",
    graphBaseUrl: "https://graph.microsoft.us/v1.0",
    managementActivityBaseUrl: "https://manage.office365.us",
  },
  usGovernmentL5: {
    id: "usGovernmentL5",
    label: "Microsoft 365 DoD",
    authorityHost: "https://login.microsoftonline.us",
    graphBaseUrl: "https://dod-graph.microsoft.us/v1.0",
    managementActivityBaseUrl: "https://manage.protection.apps.mil",
  },
  china: {
    id: "china",
    label: "Microsoft 365 operated by 21Vianet",
    authorityHost: "https://login.chinacloudapi.cn",
    graphBaseUrl: "https://microsoftgraph.chinacloudapi.cn/v1.0",
    managementActivityBaseUrl: null,
  },
} as const;

export type MicrosoftCloud = keyof typeof microsoftCloudProfiles;

export function getMicrosoftCloudProfile(cloud: MicrosoftCloud) {
  return microsoftCloudProfiles[cloud];
}

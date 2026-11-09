import type { DemoGraphExtension } from "./types.ts";

// The tenant-wide policies for Larkspur Analytical, as the Tenant settings audit reads them. The
// tenant has left several defaults as Microsoft ships them, so the audit has something to show:
// security defaults are off (the Conditional Access policy for everyone stands in for them), text
// messages still sign people in, guests can be invited by anyone, any user can register an
// application, and no one can ask for an app to be approved. User consent is limited to verified
// publishers, which meets the expectation.

const LIMITED = "10dae51f-b6af-4016-8d66-8c2a99b929b3";

const configuration = (id: string, state: "enabled" | "disabled") => ({
  "@odata.type": `#microsoft.graph.${id[0]?.toLowerCase()}${id.slice(1)}AuthenticationMethodConfiguration`,
  id,
  state,
});

const authenticationMethods = {
  id: "authenticationMethodsPolicy",
  displayName: "Authentication Methods Policy",
  policyMigrationState: "migrationComplete",
  registrationEnforcement: {
    authenticationMethodsRegistrationCampaign: {
      snoozeDurationInDays: 1,
      state: "disabled",
      excludeTargets: [],
      includeTargets: [],
    },
  },
  authenticationMethodConfigurations: [
    configuration("Sms", "enabled"),
    configuration("Voice", "disabled"),
    configuration("Email", "enabled"),
    configuration("Fido2", "enabled"),
    configuration("MicrosoftAuthenticator", "enabled"),
    configuration("TemporaryAccessPass", "disabled"),
    configuration("SoftwareOath", "enabled"),
    configuration("X509Certificate", "disabled"),
  ],
};

const authorization = {
  id: "authorizationPolicy",
  displayName: "Authorization Policy",
  allowInvitesFrom: "everyone",
  guestUserRoleId: LIMITED,
  allowEmailVerifiedUsersToJoinOrganization: false,
  allowedToSignUpEmailBasedSubscriptions: true,
  allowedToUseSSPR: true,
  blockMsolPowerShell: false,
  defaultUserRolePermissions: {
    allowedToCreateApps: true,
    allowedToCreateSecurityGroups: true,
    allowedToReadOtherUsers: true,
    permissionGrantPoliciesAssigned: ["ManagePermissionGrantsForSelf.microsoft-user-default-low"],
  },
};

const grantPolicies = [
  {
    id: "microsoft-user-default-low",
    displayName: "Allow user consent for apps from verified publishers, for selected permissions",
    description: "Users can consent to low-impact permissions of apps from verified publishers.",
  },
  {
    id: "microsoft-user-default-legacy",
    displayName: "Allow user consent for apps",
    description: "Users can consent to any app, for any permission they can consent to.",
  },
];

export const tenantSettings: DemoGraphExtension = ({ url, path }, { json, page }) => {
  switch (path) {
    case "policies/identitySecurityDefaultsEnforcementPolicy":
      return json({
        id: "00000000-0000-0000-0000-000000000005",
        displayName: "Security Defaults",
        isEnabled: false,
      });
    case "policies/authenticationMethodsPolicy":
      return json(authenticationMethods);
    case "policies/authorizationPolicy":
      return json(authorization);
    case "policies/adminConsentRequestPolicy":
      return json({ isEnabled: false, notifyReviewers: true, remindersEnabled: true, version: 0 });
    case "policies/permissionGrantPolicies":
      return page(url, grantPolicies, null);
    default:
      return null;
  }
};

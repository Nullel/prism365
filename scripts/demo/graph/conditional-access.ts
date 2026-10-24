import { guid, type SyntheticTenant } from "../tenant.ts";
import type { DemoGraphExtension } from "./types.ts";

// Conditional Access policies and named locations for Larkspur Analytical. The set is small and
// plants what the Conditional Access audit looks for: legacy authentication is blocked only in
// report-only mode, so the critical tile shows; MFA for everyone leaves out two accounts; and a
// compliant-device policy is switched off.

const DAY = 86_400_000;
const GLOBAL_ADMINISTRATOR = "62e90394-69f5-4237-9190-012177145e10";
const SECURITY_ADMINISTRATOR = "194ae4cb-b126-40b2-bd5b-6091b380977d";
const EXCHANGE_ADMINISTRATOR = "29232cdf-9323-42fd-ade2-1d097af3e4de";
const PRIVILEGED_ROLE_ADMINISTRATOR = "e8611ab8-c189-46e8-94e1-60213ab1f814";

// The ids are made when asked for, not when this file loads: the tenant module and the graph
// index import each other, and `guid` is only safe to call once both have loaded.
function ids() {
  return {
    madisonLab: guid("named-location:madison-lab"),
    countries: guid("named-location:approved-countries"),
    travellers: guid("group:frequent-travellers"),
    fieldStaff: guid("group:field-sampling-staff"),
    strength: guid("authentication-strength:phishing-resistant"),
  };
}

function userId(tenant: SyntheticTenant, name: string, fallback: number): string {
  return (
    (tenant.users.find((user) => user.displayName === name) ?? tenant.users[fallback])?.id ?? ""
  );
}

function namedLocations(now: number) {
  const at = (daysAgo: number) => new Date(now - daysAgo * DAY).toISOString();
  const id = ids();
  return [
    {
      "@odata.type": "#microsoft.graph.ipNamedLocation",
      id: id.madisonLab,
      displayName: "Madison laboratory",
      createdDateTime: at(400),
      modifiedDateTime: at(120),
      isTrusted: true,
      ipRanges: [
        { "@odata.type": "#microsoft.graph.iPv4CidrRange", cidrAddress: "203.0.113.0/24" },
      ],
    },
    {
      "@odata.type": "#microsoft.graph.countryNamedLocation",
      id: id.countries,
      displayName: "Approved countries",
      createdDateTime: at(380),
      modifiedDateTime: at(380),
      countriesAndRegions: ["US", "CA"],
      includeUnknownCountriesAndRegions: false,
    },
  ];
}

function policies(tenant: SyntheticTenant) {
  const now = tenant.now.getTime();
  const at = (daysAgo: number) => new Date(now - daysAgo * DAY).toISOString();
  const id = ids();
  // Two accounts kept out of the MFA policy so they can always sign in.
  const emergency = [userId(tenant, "Demo administrator", 0), userId(tenant, "Tobias Vance", 1)];
  const users = (over: Record<string, unknown> = {}) => ({
    includeUsers: ["All"],
    excludeUsers: [],
    includeGroups: [],
    excludeGroups: [],
    includeRoles: [],
    excludeRoles: [],
    ...over,
  });
  const conditions = (over: Record<string, unknown> = {}) => ({
    applications: { includeApplications: ["All"], excludeApplications: [] },
    clientAppTypes: ["all"],
    locations: null,
    platforms: null,
    signInRiskLevels: [],
    userRiskLevels: [],
    users: users(),
    ...over,
  });
  const grant = (builtInControls: string[], operator = "OR") => ({
    operator,
    builtInControls,
    customAuthenticationFactors: [],
    termsOfUse: [],
    authenticationStrength: null,
  });
  return [
    {
      id: guid("policy:block-legacy"),
      templateId: null,
      displayName: "Block legacy authentication",
      createdDateTime: at(21),
      modifiedDateTime: at(6),
      state: "enabledForReportingButNotEnforced",
      conditions: conditions({ clientAppTypes: ["exchangeActiveSync", "other"] }),
      grantControls: grant(["block"]),
      sessionControls: null,
    },
    {
      id: guid("policy:mfa-all-users"),
      templateId: null,
      displayName: "Require MFA for all users",
      createdDateTime: at(330),
      modifiedDateTime: at(45),
      state: "enabled",
      conditions: conditions({ users: users({ excludeUsers: emergency }) }),
      grantControls: grant(["mfa"]),
      sessionControls: {
        signInFrequency: {
          isEnabled: true,
          type: "days",
          value: 14,
          frequencyInterval: "timeBased",
          authenticationType: "primaryAndSecondaryAuthentication",
        },
      },
    },
    {
      id: guid("policy:admin-portals"),
      templateId: null,
      displayName: "Require phishing-resistant MFA for admin portals",
      createdDateTime: at(210),
      modifiedDateTime: at(210),
      state: "enabled",
      conditions: conditions({
        applications: { includeApplications: ["MicrosoftAdminPortals"], excludeApplications: [] },
        users: users({
          includeUsers: [],
          includeRoles: [
            GLOBAL_ADMINISTRATOR,
            SECURITY_ADMINISTRATOR,
            EXCHANGE_ADMINISTRATOR,
            PRIVILEGED_ROLE_ADMINISTRATOR,
          ],
        }),
      }),
      grantControls: {
        operator: "AND",
        builtInControls: [],
        customAuthenticationFactors: [],
        termsOfUse: [],
        authenticationStrength: {
          id: id.strength,
          displayName: "Phishing-resistant MFA",
          policyType: "builtIn",
          requirementsSatisfied: "mfa",
          allowedCombinations: ["fido2", "windowsHelloForBusiness"],
        },
      },
      sessionControls: {
        signInFrequency: {
          isEnabled: true,
          type: "hours",
          value: 4,
          frequencyInterval: "timeBased",
          authenticationType: "primaryAndSecondaryAuthentication",
        },
        persistentBrowser: { isEnabled: true, mode: "never" },
      },
    },
    {
      id: guid("policy:compliant-device"),
      templateId: null,
      displayName: "Require a compliant device for Office 365",
      createdDateTime: at(150),
      modifiedDateTime: at(33),
      state: "disabled",
      conditions: conditions({
        applications: { includeApplications: ["Office365"], excludeApplications: [] },
        users: users({ excludeGroups: [id.fieldStaff] }),
        platforms: { includePlatforms: ["windows", "macOS"], excludePlatforms: [] },
        locations: { includeLocations: ["All"], excludeLocations: [id.madisonLab] },
      }),
      grantControls: grant(["compliantDevice"]),
      sessionControls: null,
    },
    {
      id: guid("policy:approved-countries"),
      templateId: null,
      displayName: "Block sign-ins from outside approved countries",
      createdDateTime: at(300),
      modifiedDateTime: at(90),
      state: "enabled",
      conditions: conditions({
        users: users({ excludeGroups: [id.travellers] }),
        locations: { includeLocations: ["All"], excludeLocations: [id.countries] },
      }),
      grantControls: grant(["block"]),
      sessionControls: null,
    },
    {
      id: guid("policy:risky-sign-ins"),
      templateId: null,
      displayName: "Require MFA for risky sign-ins",
      createdDateTime: at(95),
      modifiedDateTime: at(95),
      state: "enabled",
      conditions: conditions({ signInRiskLevels: ["high", "medium"] }),
      grantControls: grant(["mfa"]),
      sessionControls: {
        signInFrequency: {
          isEnabled: true,
          frequencyInterval: "everyTime",
          authenticationType: "primaryAndSecondaryAuthentication",
        },
      },
    },
  ];
}

export const conditionalAccess: DemoGraphExtension = ({ url, path, tenant }, { page }) => {
  if (path === "identity/conditionalAccess/policies") return page(url, policies(tenant), null);
  if (path === "identity/conditionalAccess/namedLocations")
    return page(url, namedLocations(tenant.now.getTime()), null);
  return null;
};

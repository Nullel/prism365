import type { SyntheticTenant } from "../tenant.ts";
import type { DemoGraphExtension } from "./types.ts";

// Identity Protection risky users for Larkspur Analytical: one high risk and one medium risk still
// open, one remediated recently and one dismissed a while ago. The medium risk user is in the
// laboratory, which signs in on the research subsidiary's domain, so the audit scope has
// something to leave out.

const DAY = 86_400_000;

function riskyUserRecords(tenant: SyntheticTenant) {
  const now = tenant.now.getTime();
  const user = (name: string, fallback: number) => {
    const found = tenant.users.find((candidate) => candidate.displayName === name);
    const chosen = found ?? tenant.users[fallback];
    if (!chosen) throw new Error(`The synthetic tenant has no user for ${name}.`);
    return chosen;
  };
  const entry = (
    name: string,
    fallback: number,
    risk: { level: string; state: string; detail: string; daysAgo: number },
  ) => {
    const account = user(name, fallback);
    return {
      id: account.id,
      isDeleted: false,
      isProcessing: false,
      riskLastUpdatedDateTime: new Date(now - risk.daysAgo * DAY).toISOString(),
      riskLevel: risk.level,
      riskState: risk.state,
      riskDetail: risk.detail,
      userDisplayName: account.displayName,
      userPrincipalName: account.userPrincipalName,
    };
  };
  return [
    entry("Juniper Ashdown", 13, { level: "high", state: "atRisk", detail: "none", daysAgo: 2 }),
    entry("Dagny Holloway", 4, { level: "medium", state: "atRisk", detail: "none", daysAgo: 5 }),
    entry("Kasimir Orne", 5, {
      level: "none",
      state: "remediated",
      detail: "userPerformedSecuredPasswordReset",
      daysAgo: 9,
    }),
    entry("Sunniva Petrakis", 15, {
      level: "none",
      state: "dismissed",
      detail: "adminDismissedAllRiskForUser",
      daysAgo: 41,
    }),
  ];
}

export const riskyUsers: DemoGraphExtension = ({ url, path, tenant }, { page }) =>
  path === "identityProtection/riskyUsers" ? page(url, riskyUserRecords(tenant), null) : null;

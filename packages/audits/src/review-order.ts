import { COMBINED_RISKS_AUDIT_ID } from "@prism365/contracts";

// The order a full review runs the audits in (docs/audits.md, section 22), and what each one costs
// Microsoft's report limit. A test checks that every built-in audit is here exactly once, that the
// stages are in order, and that each audit sits in the stage its packs say.
//
//   directory  reads that have their own, generous limits
//   paced      Conditional Access policies and Risky users, which Microsoft limits to one request a
//              second for the whole tenant
//   reports    the usage-report audits; Microsoft allows 14 report downloads in ten minutes
//   summary    Combined risks, which joins the stored results of the rest

export type ReviewStage = "directory" | "paced" | "reports" | "summary";

export interface ReviewStep {
  auditId: string;
  stage: ReviewStage;
  // The most report downloads the audit can start in one run. 0 for an audit that downloads none.
  reportDownloads: number;
}

const step = (auditId: string, stage: ReviewStage, reportDownloads = 0): ReviewStep => ({
  auditId,
  stage,
  reportDownloads,
});

export const REVIEW_ORDER: readonly ReviewStep[] = [
  step("accounts.inactive", "directory"),
  step("licenses.reclaim", "directory"),
  step("mfa.registration", "directory"),
  step("guests.review", "directory"),
  step("accounts.passwords", "directory"),
  step("directory.privileged-roles", "directory"),
  step("apps.consents", "directory"),
  step("devices.managed", "directory"),
  step("licenses.capacity", "directory"),
  step("identity.conditional-access", "paced"),
  step("identity.risky-users", "paced"),
  // Seven reports for Account activity, the heaviest, then the audits that read one or two.
  step("accounts.activity", "reports", 7),
  step("usage.m365-apps", "reports", 2),
  step("mail.mailbox-storage", "reports", 1),
  step("groups.teams", "reports", 1),
  step(COMBINED_RISKS_AUDIT_ID, "summary"),
];

// Microsoft allows 14 report downloads in ten minutes; a review starts no more than 12, which
// leaves room for a download an administrator starts from the page at the same time.
export const MAX_REPORT_DOWNLOADS = 12;
export const REPORT_WINDOW_MILLISECONDS = 10 * 60_000;

import {
  automationManifestSchema,
  automationPreviewRequestSchema,
  automationPreviewSchema,
  automationWorkflowListSchema,
  automationWorkflowSummarySchema,
  type AutomationCapability,
  type AutomationManifest,
  type AutomationPreview,
  type AutomationPreviewRequest,
  type AutomationWorkflowSummary,
} from "@prism365/contracts";

type RegisteredWorkflow = {
  manifest: AutomationManifest;
  origin: "builtIn" | "imported";
  trust: "approved";
};

const permissionByCapability: Record<AutomationCapability, string> = {
  "identity.lookup": "User.Read.All",
  "identity.user.create": "User.Create",
  "identity.user.disable": "User.EnableDisableAccount.All + User.Read.All",
  "identity.sessions.revoke": "User.RevokeSessions.All",
  "licenses.assign": "LicenseAssignment.ReadWrite.All",
  "licenses.remove": "LicenseAssignment.ReadWrite.All",
  "groups.add": "GroupMember.ReadWrite.All",
  "groups.remove": "GroupMember.ReadWrite.All",
  "mailbox.handoff": "Exchange Online resource-scoped RBAC",
  "data.handoff": "Workload-specific resource-scoped authorization",
};

const standardOnboarding = automationManifestSchema.parse({
  schemaVersion: 1,
  id: "prism365.standard-onboarding",
  name: "Standard employee onboarding",
  version: "0.1.0",
  description: "Create an account, assign an approved license profile, and add approved groups.",
  lifecycle: "onboarding",
  runtime: "prism365-native",
  entrypoint: "workflows/standard-onboarding",
  publisher: "Prism365 sample library",
  parameters: [
    {
      key: "targetEmail",
      label: "Work email",
      type: "email",
      required: true,
      secret: false,
      helpText: "The work or school account to create.",
    },
  ],
  capabilities: ["identity.lookup", "identity.user.create", "licenses.assign", "groups.add"],
  allowedHosts: [],
  timeoutSeconds: 300,
  steps: [
    {
      id: "verify-target",
      label: "Verify the account does not already exist",
      description: "Check the selected tenant for an existing matching work account.",
      capability: "identity.lookup",
      risk: "readOnly",
      reversible: true,
      recovery: "No tenant change is made.",
    },
    {
      id: "create-account",
      label: "Create the user account",
      description: "Create a work account using an administrator-approved profile.",
      capability: "identity.user.create",
      risk: "high",
      reversible: true,
      recovery: "Disable and quarantine the new account; deletion requires a separate review.",
    },
    {
      id: "assign-license",
      label: "Assign the selected license profile",
      description: "Apply only the license profile selected in the reviewed runbook.",
      capability: "licenses.assign",
      risk: "moderate",
      reversible: true,
      recovery: "Remove the license assignment after dependency review.",
    },
    {
      id: "add-groups",
      label: "Add approved group memberships",
      description: "Add only static, non-role-assignable groups declared by the runbook.",
      capability: "groups.add",
      risk: "high",
      reversible: true,
      recovery: "Remove memberships added by this run after reviewing downstream access.",
    },
  ],
});

const secureOffboarding = automationManifestSchema.parse({
  schemaVersion: 1,
  id: "prism365.secure-offboarding",
  name: "Secure employee offboarding",
  version: "0.1.0",
  description: "Block access, revoke sessions, remove approved access, and prepare data handoff.",
  lifecycle: "offboarding",
  runtime: "powershell7",
  entrypoint: "scripts/secure-offboarding.ps1",
  publisher: "Prism365 sample package",
  parameters: [
    {
      key: "targetEmail",
      label: "Work email",
      type: "email",
      required: true,
      secret: false,
      helpText: "The work or school account to offboard.",
    },
  ],
  capabilities: [
    "identity.lookup",
    "identity.user.disable",
    "identity.sessions.revoke",
    "groups.remove",
    "licenses.remove",
    "mailbox.handoff",
    "data.handoff",
  ],
  allowedHosts: [],
  timeoutSeconds: 600,
  steps: [
    {
      id: "protect-target",
      label: "Verify the target and protected-account rules",
      description: "Stop if the account is the operator, break-glass protected, or privileged.",
      capability: "identity.lookup",
      risk: "readOnly",
      reversible: true,
      recovery: "No tenant change is made.",
    },
    {
      id: "disable-sign-in",
      label: "Block new sign-ins",
      description: "Set the account to disabled after authorization is rechecked.",
      capability: "identity.user.disable",
      risk: "high",
      reversible: true,
      recovery: "Re-enable the account after an authorized incident review.",
    },
    {
      id: "revoke-sessions",
      label: "Revoke active sessions",
      description: "Invalidate refresh tokens; provider propagation can take several minutes.",
      capability: "identity.sessions.revoke",
      risk: "high",
      reversible: false,
      recovery: "The user must authenticate again if the account is later restored.",
    },
    {
      id: "remove-groups",
      label: "Remove approved group memberships",
      description: "Exclude dynamic groups, role-assignable groups, and protected access paths.",
      capability: "groups.remove",
      risk: "high",
      reversible: true,
      recovery: "Restore memberships from the immutable before-state record.",
    },
    {
      id: "prepare-handoff",
      label: "Prepare mailbox and data handoff",
      description: "Pause for an authorized owner and workload-specific policy review.",
      capability: "mailbox.handoff",
      risk: "manualReview",
      reversible: false,
      recovery: "Follow the workload-specific recovery runbook and retention policy.",
    },
    {
      id: "remove-licenses",
      label: "Review and remove eligible licenses",
      description: "Remove only licenses whose service and retention dependencies are cleared.",
      capability: "licenses.remove",
      risk: "manualReview",
      reversible: true,
      recovery: "Reassign the recorded license set if capacity remains available.",
    },
  ],
});

const workflows: RegisteredWorkflow[] = [
  { manifest: standardOnboarding, origin: "builtIn", trust: "approved" },
  { manifest: secureOffboarding, origin: "imported", trust: "approved" },
];

function toSummary(workflow: RegisteredWorkflow): AutomationWorkflowSummary {
  return automationWorkflowSummarySchema.parse({
    id: workflow.manifest.id,
    name: workflow.manifest.name,
    version: workflow.manifest.version,
    description: workflow.manifest.description,
    lifecycle: workflow.manifest.lifecycle,
    runtime: workflow.manifest.runtime,
    publisher: workflow.manifest.publisher,
    origin: workflow.origin,
    trust: workflow.trust,
    capabilities: workflow.manifest.capabilities,
    stepCount: workflow.manifest.steps.length,
  });
}

function previewId(request: AutomationPreviewRequest, now: Date): string {
  const input = `${request.workflowId}|${request.targetEmail}|${now.toISOString()}`;
  let hash = 2166136261;
  for (const character of input) {
    hash ^= character.charCodeAt(0);
    hash = Math.imul(hash, 16777619);
  }
  return `preview_${(hash >>> 0).toString(16).padStart(8, "0")}`;
}

export function listAutomationWorkflows() {
  return automationWorkflowListSchema.parse({
    mode: "demo",
    importEnabled: false,
    executionEnabled: false,
    items: workflows.map(toSummary),
  });
}

export function createAutomationPreview(
  input: AutomationPreviewRequest,
  now = new Date(),
): AutomationPreview {
  const request = automationPreviewRequestSchema.parse(input);
  const workflow = workflows.find((candidate) => candidate.manifest.id === request.workflowId);
  if (!workflow) {
    throw new Error("Unknown lifecycle workflow.");
  }

  const requiredPermissions = [
    ...new Set(
      workflow.manifest.capabilities.map((capability) => permissionByCapability[capability]),
    ),
  ];
  const expiresAt = new Date(now.getTime() + 15 * 60 * 1_000);

  return automationPreviewSchema.parse({
    mode: "demo",
    previewId: previewId(request, now),
    workflow: toSummary(workflow),
    targetEmail: request.targetEmail,
    generatedAt: now.toISOString(),
    expiresAt: expiresAt.toISOString(),
    executionAllowed: false,
    confirmationPhrase: `${workflow.manifest.lifecycle.toUpperCase()} ${request.targetEmail}`,
    steps: workflow.manifest.steps.map((step, index) => ({
      ...step,
      order: index + 1,
      status: step.risk === "manualReview" ? "manualReview" : "ready",
      permission: permissionByCapability[step.capability],
    })),
    requiredPermissions,
    blockedReasons: [
      "Tenant write connections are disabled in this pre-alpha build.",
      "The isolated script runner, approval ledger, and reconciliation worker are not implemented.",
    ],
    limitations: [
      "This preview uses synthetic workflow metadata and does not inspect a Microsoft 365 tenant.",
      "A production preview must resolve the target, detect protected accounts, and record before-state evidence.",
    ],
  });
}

export { permissionByCapability };
export { inspectAutomationPackage, summarizeAutomationPackage } from "./registry.ts";
export type {
  AutomationPackageRecord,
  AutomationPackageSummary,
  AutomationPackageFile,
} from "./registry.ts";

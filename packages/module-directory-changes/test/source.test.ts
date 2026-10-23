import { describe, expect, it } from "vitest";
import {
  changeSubjects,
  NAME_LIMIT,
  readDirectoryAudit,
  TARGET_LIMIT,
  unquote,
  VALUE_LIMIT,
} from "../src/index.ts";

// Synthetic records shaped like Microsoft's documented samples; invented names and ids only.

const adminId = "0a1b2c3d-0000-4000-8000-00000000a001";
const personId = "0A1B2C3D-0000-4000-8000-00000000B002";

const record = (overrides: Record<string, unknown> = {}) => ({
  id: "Directory_0a1b2c3d_ABCDE_1",
  category: "UserManagement",
  correlationId: "0a1b2c3d-0000-4000-8000-00000000c003",
  result: "Success",
  resultReason: "",
  activityDisplayName: "Update user",
  activityDateTime: "2026-10-21T09:15:02.7215374Z",
  loggedByService: "Core Directory",
  operationType: "Update",
  initiatedBy: {
    user: {
      id: adminId,
      displayName: "Synthetic Admin",
      userPrincipalName: "admin@example.invalid",
      ipAddress: "192.0.2.10",
    },
    app: null,
  },
  targetResources: [
    {
      id: personId,
      displayName: "Synthetic Person",
      Type: "User",
      userPrincipalName: "person@example.invalid",
      modifiedProperties: [
        { displayName: "Department", oldValue: '"Research"', newValue: '"Sales"' },
        { displayName: "AccountEnabled", oldValue: "[true]", newValue: "[false]" },
        {
          displayName: "StrongAuthenticationMethod",
          oldValue: '[{"MethodType":6,"Default":true}]',
          newValue: "[]",
        },
      ],
    },
  ],
  additionalDetails: [{ key: "UserType", value: "Member" }],
  ...overrides,
});

describe("directory audit source", () => {
  it("normalizes a record: unquoted values, a case-insensitive target type, the actor", () => {
    const event = readDirectoryAudit(record());
    expect(event).toMatchObject({
      id: "Directory_0a1b2c3d_ABCDE_1",
      activity: "Update user",
      category: "UserManagement",
      result: "success",
      initiatedBy: {
        kind: "user",
        id: adminId,
        name: "Synthetic Admin",
        principal: "admin@example.invalid",
      },
      details: [{ key: "UserType", value: "Member" }],
    });
    expect(event?.activityAt.toISOString()).toBe("2026-10-21T09:15:02.721Z");
    const [target] = event?.targets ?? [];
    expect(target).toMatchObject({ id: personId, type: "User", name: "Synthetic Person" });
    expect(target?.properties).toEqual([
      { name: "Department", oldValue: "Research", newValue: "Sales" },
      { name: "AccountEnabled", oldValue: "true", newValue: "false" },
      {
        name: "StrongAuthenticationMethod",
        oldValue: '[{"MethodType":6,"Default":true}]',
        newValue: "",
      },
    ]);
    // The actor's address is read for nothing and kept nowhere.
    expect(JSON.stringify(event)).not.toContain("192.0.2.10");
  });

  it("cuts long values and says so, and bounds every list", () => {
    const long = "x".repeat(VALUE_LIMIT + 50);
    const event = readDirectoryAudit(
      record({
        activityDisplayName: `A${"n".repeat(NAME_LIMIT + 10)}\u0000`,
        targetResources: Array.from({ length: TARGET_LIMIT + 5 }, (_value, index) => ({
          id: `target-${index}`,
          type: "Group",
          modifiedProperties: Array.from({ length: 60 }, () => ({
            displayName: "Description",
            oldValue: JSON.stringify(long),
            newValue: "short",
          })),
        })),
        additionalDetails: Array.from({ length: 30 }, (_value, index) => ({
          key: `k${index}`,
          value: "v",
        })),
      }),
    );
    expect(event?.activity?.length).toBe(NAME_LIMIT);
    expect(event?.activity).not.toContain("\u0000");
    expect(event?.targets).toHaveLength(TARGET_LIMIT);
    expect(event?.targets[0]?.properties).toHaveLength(50);
    expect(event?.targets[0]?.properties[0]).toEqual({
      name: "Description",
      oldValue: "x".repeat(VALUE_LIMIT),
      newValue: "short",
      oldCut: true,
    });
    expect(event?.details).toHaveLength(20);
  });

  it("reads an application as the actor, and nobody when Microsoft names nobody", () => {
    const byApp = readDirectoryAudit(
      record({
        initiatedBy: {
          user: null,
          app: {
            appId: "0a1b2c3d-0000-4000-8000-00000000d004",
            displayName: "Synthetic Sync",
            servicePrincipalId: "0a1b2c3d-0000-4000-8000-00000000e005",
            servicePrincipalName: null,
          },
        },
      }),
    );
    expect(byApp?.initiatedBy).toEqual({
      kind: "app",
      id: "0a1b2c3d-0000-4000-8000-00000000e005",
      name: "Synthetic Sync",
      principal: "0a1b2c3d-0000-4000-8000-00000000d004",
    });
    const system = readDirectoryAudit(
      record({
        initiatedBy: {
          user: {
            id: "00000000-0000-0000-0000-000000000000",
            displayName: null,
            userPrincipalName: null,
            ipAddress: "10.0.0.0",
          },
          app: null,
        },
      }),
    );
    expect(system?.initiatedBy).toEqual({ kind: "unknown", id: null, name: null, principal: null });
  });

  it("tolerates odd shapes and leaves out only records it cannot store", () => {
    expect(
      readDirectoryAudit(record({ initiatedBy: 7, targetResources: "x", additionalDetails: null })),
    ).toMatchObject({ initiatedBy: { kind: "unknown" }, targets: [], details: [] });
    expect(readDirectoryAudit(record({ id: "" }))).toBeNull();
    expect(readDirectoryAudit(record({ id: undefined }))).toBeNull();
    expect(readDirectoryAudit(record({ activityDateTime: "yesterday" }))).toBeNull();
    expect(readDirectoryAudit("not a record")).toBeNull();
  });

  it("unquotes Microsoft's JSON values and keeps anything else as written", () => {
    expect(unquote('"Global Administrator"')).toBe("Global Administrator");
    expect(unquote('["a@example.invalid","b@example.invalid"]')).toBe(
      "a@example.invalid, b@example.invalid",
    );
    expect(unquote("[SkuName=SPB, ServicePlan=EXCHANGE]")).toBe(
      "[SkuName=SPB, ServicePlan=EXCHANGE]",
    );
    expect(unquote('{"a":1}')).toBe('{"a":1}');
  });

  it("lists the actor and each target once, by lower-case id", () => {
    const event = readDirectoryAudit(
      record({
        targetResources: [
          { id: personId, type: "User" },
          { id: personId.toLowerCase(), type: "User" },
          { id: "00000000-0000-0000-0000-000000000000", type: "Other" },
          { id: null, type: "Group" },
        ],
      }),
    );
    expect(event ? changeSubjects(event) : []).toEqual([
      { id: adminId, kind: "actor", type: "user" },
      { id: personId.toLowerCase(), kind: "target", type: "user" },
    ]);
  });
});

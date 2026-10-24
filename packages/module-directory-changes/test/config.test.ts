import { describe, expect, it } from "vitest";
import {
  createDirectoryChangesModule,
  defaultDirectoryChangesConfig,
  directoryChangesConfigFromEnv,
} from "../src/index.ts";

describe("directory change settings", () => {
  it("re-reads four hours behind and keeps ten years unless told otherwise", () => {
    expect(directoryChangesConfigFromEnv({})).toEqual(defaultDirectoryChangesConfig);
    expect(defaultDirectoryChangesConfig).toMatchObject({
      lateArrivalMinutes: 240,
      retentionDays: 3_650,
    });
    expect(createDirectoryChangesModule().retention).toEqual([
      { table: "event", timeColumn: "activity_at", days: 3_650 },
      { table: "subject", timeColumn: "activity_at", days: 3_650 },
    ]);
  });

  it("reads both settings from the environment, inside their ranges", () => {
    const config = directoryChangesConfigFromEnv({
      DIRECTORY_CHANGES_LATE_MINUTES: "60",
      DIRECTORY_CHANGES_RETENTION_DAYS: " 400 ",
    });
    expect(config).toMatchObject({ lateArrivalMinutes: 60, retentionDays: 400 });
    expect(createDirectoryChangesModule(config).retention.map((rule) => rule.days)).toEqual([
      400, 400,
    ]);
    for (const env of [
      { DIRECTORY_CHANGES_LATE_MINUTES: "59" },
      { DIRECTORY_CHANGES_LATE_MINUTES: "1441" },
      { DIRECTORY_CHANGES_LATE_MINUTES: "90.5" },
      { DIRECTORY_CHANGES_RETENTION_DAYS: "364" },
      { DIRECTORY_CHANGES_RETENTION_DAYS: "3651" },
      { DIRECTORY_CHANGES_RETENTION_DAYS: "ten years" },
    ])
      expect(() => directoryChangesConfigFromEnv(env), JSON.stringify(env)).toThrow(
        /DIRECTORY_CHANGES_/,
      );
  });
});

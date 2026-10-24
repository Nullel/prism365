import type { Audit } from "@prism365/audits";
import {
  GraphBoundaryError,
  GraphReportDownloadError,
  GraphReportHostError,
  GraphRequestError,
} from "@prism365/graph";
import { describe, expect, it } from "vitest";
import { auditFailureLogLine, describeAuditFailure } from "../src/audits.ts";

// What an administrator and the server log are told when a usage report cannot be read (ADR-0010).
// Every value is invented.

const audit = { forbidden: "Microsoft refused the read. It needs Reports Reader." } as Audit;

describe("usage report failures on the audit page", () => {
  it("name a report host Prism365 does not recognise", () => {
    const error = new GraphReportHostError(
      "its host is not a report download host",
      "reports.example",
      false,
    );
    expect(describeAuditFailure(error, audit)).toEqual({
      code: "report_host",
      message:
        "Microsoft sent the report from a host Prism365 does not recognise: reports.example.",
    });
  });

  it("say what was wrong when the host is the known one or there is none", () => {
    expect(
      describeAuditFailure(
        new GraphReportHostError("it names a port", "reports.office.com", true),
        audit,
      ),
    ).toEqual({
      code: "report_host",
      message:
        "Microsoft sent a report download link on reports.office.com that Prism365 does not accept: it names a port.",
    });
    expect(
      describeAuditFailure(
        new GraphReportHostError("it is not an absolute URL", null, false),
        audit,
      ),
    ).toEqual({
      code: "report_host",
      message: "Microsoft sent a report download link Prism365 could not read.",
    });
  });

  it("say a failed download link failed, not that the session or a role is missing", () => {
    for (const status of [401, 403, 404, 500]) {
      const outcome = describeAuditFailure(new GraphReportDownloadError(status), audit);
      expect(outcome, String(status)).toEqual({
        code: "report_download",
        message: "Microsoft's report download link failed. Run the audit again.",
      });
    }
    expect(
      describeAuditFailure(new GraphBoundaryError("The download failed.", "download"), audit),
    ).toMatchObject({ code: "report_download" });
    // A 403 from Graph itself still explains the role the audit needs.
    expect(describeAuditFailure(new GraphRequestError(403), audit)).toMatchObject({
      code: "forbidden",
      message: audit.forbidden,
    });
  });
});

describe("the run log line", () => {
  it("carries the refused host with its category", () => {
    const error = new GraphReportHostError(
      "its host is not a report download host",
      "reports.example",
      false,
    );
    expect(auditFailureLogLine("usage.example", "report_host", error)).toBe(
      "Prism365 audit usage.example failed (report_host, host reports.example).",
    );
    expect(
      auditFailureLogLine(
        "usage.example",
        "report_host",
        new GraphReportHostError("it is not an absolute URL", null, false),
      ),
    ).toBe("Prism365 audit usage.example failed (report_host, no host).");
  });

  it("keeps the other lines as they were", () => {
    expect(auditFailureLogLine("usage.example", "failed", new TypeError("x"))).toBe(
      "Prism365 audit usage.example failed (TypeError).",
    );
    expect(auditFailureLogLine("usage.example", "invalid_result", "text")).toBe(
      "Prism365 audit usage.example failed (string).",
    );
    expect(auditFailureLogLine("usage.example", "forbidden", new GraphRequestError(403))).toBe(
      null,
    );
  });
});

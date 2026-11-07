import {
  type MailboxAccessStatus,
  mailboxAccessStatusSchema,
  type PersonMailboxAccess,
  personMailboxAccessSchema,
} from "@prism365/contracts";
import { readFileSync } from "node:fs";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";
import {
  COMMANDS,
  type ExportImportState,
  importMessage,
  MailboxAccessView,
  RELEASE_TAGS_URL,
  readExportFile,
  SCRIPT_SHA256,
  SCRIPT_URL,
  type ScriptHashState,
  sha256Hex,
} from "../src/MailboxAccess.tsx";
import { SharedMailboxesSection } from "../src/PersonAccess.tsx";
import {
  at,
  findAll,
  type Host,
  hostTree,
  nestedInteractive,
  plainClick,
  textOf,
} from "./markup.ts";

// The Mailbox access section on the Microsoft access page and the Shared mailboxes section of a
// person's Access tab, drawn from answers in hand. Synthetic names and .invalid addresses only.
const NOW = "2026-11-06T12:00:00.000Z";
const DAY = 86_400_000;
const daysAgo = (days: number) => new Date(Date.parse(NOW) - days * DAY).toISOString();
const text = (html: string) =>
  html
    .replaceAll("<!-- -->", "")
    .replace(/<[^>]*>/g, " ")
    .replace(/&#x27;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/\s+/g, " ");

const stored = (over: Record<string, unknown> = {}) => ({
  id: "44444444-4444-4444-8444-444444444444",
  exportedAt: daysAgo(2),
  importedAt: daysAgo(1),
  importedBy: { id: "22222222-2222-4222-8222-222222222222", name: "Synthetic admin" },
  includeUserMailboxes: false,
  mailboxes: 12,
  grants: 40,
  unresolved: 3,
  ...over,
});
const status = (value: unknown): MailboxAccessStatus =>
  mailboxAccessStatusSchema.parse({ import: value });

const noop = () => {};
const view = (props: {
  status: MailboxAccessStatus | null;
  error?: string;
  importState?: ExportImportState;
  confirming?: boolean;
  readOnly?: boolean;
  scriptHash?: ScriptHashState;
  onClear?: () => void;
  onConfirmClear?: () => void;
}) =>
  createElement(MailboxAccessView, {
    status: props.status,
    error: props.error ?? "",
    importState: props.importState ?? { state: "idle" },
    confirming: props.confirming ?? false,
    clearing: false,
    readOnly: props.readOnly ?? false,
    scriptHash: props.scriptHash ?? { state: "ready", sha256: SCRIPT_SHA256 },
    onImport: noop,
    onClear: props.onClear ?? noop,
    onConfirmClear: props.onConfirmClear ?? noop,
    onCancelClear: noop,
  });
const press = (button: Host | undefined) => {
  if (!button) throw new Error("No such button.");
  (button.props.onClick as (event: unknown) => void)(plainClick());
};
const draw = (props: Parameters<typeof view>[0]) =>
  at(NOW, () => renderToStaticMarkup(view(props)));

describe("the Mailbox access section", () => {
  it("explains the export, offers the script and lists the commands", () => {
    const html = draw({ status: status(null) });
    const words = text(html);
    expect(words).toContain("Microsoft offers no read-only way for Prism365");
    expect(words).toContain("never holds an Exchange credential and never calls Exchange");
    expect(html).toContain(`href="${SCRIPT_URL}" download="Export-MailboxAccess.ps1"`);
    for (const command of COMMANDS) expect(html).toContain(command);
    expect(words).toContain("Get-ManagementRole -Cmdlet Get-MailboxPermission");
    expect(words).toContain("Nothing is imported.");
    expect(html).toContain('type="file"');
    expect(nestedInteractive(html)).toBe(0);
  });

  it("shows the current import with its date, importer and counts, unresolved trustees counted", () => {
    const words = text(draw({ status: status(stored()) }));
    expect(words).toMatch(
      /From the export of Nov 4, 2026,? .*, imported yesterday by Synthetic admin\./,
    );
    expect(words).toContain("12 mailboxes, shared mailboxes only, 40 grants.");
    expect(words).toContain("3 grants name trustees the script could not resolve");
    expect(words).not.toContain("Older than 30 days");
  });

  it("warns when the export is older than 30 days", () => {
    const html = draw({
      status: status(stored({ exportedAt: daysAgo(31), importedAt: daysAgo(31) })),
    });
    expect(html).toContain("sev--medium");
    expect(text(html)).toContain("Older than 30 days");
    expect(text(html)).toContain("Run the script again for a current export.");
  });

  it("asks before clearing, and clears only on the second press", () => {
    const onClear = vi.fn();
    const onConfirmClear = vi.fn();
    const first = at(NOW, () =>
      hostTree(view({ status: status(stored()), onClear, onConfirmClear })),
    );
    const clear = findAll(first, (node) => node.tag === "button" && textOf(node) === "Clear");
    expect(clear).toHaveLength(1);
    press(clear[0]);
    expect(onClear).toHaveBeenCalledTimes(1);
    expect(onConfirmClear).not.toHaveBeenCalled();
    const asked = at(NOW, () =>
      hostTree(view({ status: status(stored()), confirming: true, onClear, onConfirmClear })),
    );
    expect(
      findAll(asked, (node) => node.tag === "p")
        .map(textOf)
        .join(" "),
    ).toContain("Clear the import?");
    const confirm = findAll(
      asked,
      (node) => node.tag === "button" && textOf(node) === "Clear the import",
    );
    press(confirm[0]);
    expect(onConfirmClear).toHaveBeenCalledTimes(1);
  });

  it("shows an import's result and a refusal as text, never as markup", () => {
    const refused = draw({
      status: status(null),
      importState: {
        state: "failed",
        name: "<img src=x onerror=alert(1)>.json",
        message: 'mailboxes[0].displayName has a field Prism365 does not read: "<b>x</b>".',
      },
    });
    expect(refused).not.toContain("<img");
    expect(refused).not.toContain("<b>");
    expect(refused).toContain("&lt;img src=x onerror=alert(1)&gt;.json");
    const result = {
      import: mailboxAccessStatusSchema.parse({ import: stored() }).import,
      replaced: true,
    } as Parameters<typeof importMessage>[0];
    expect(at(NOW, () => importMessage(result))).toMatch(
      /^Imported the export of Nov 4, 2026.*: 12 mailboxes and 40 grants\. It replaces the previous import\.$/,
    );
  });

  it("offers no upload or clear in the public demo", () => {
    const html = draw({ status: status(stored()), readOnly: true });
    expect(html).not.toContain('type="file"');
    expect(html).not.toContain(">Clear<");
    expect(text(html)).toContain("The public demo cannot keep a file.");
  });

  it("is written without em dashes, middle dots or exclamation marks", () => {
    expect(text(draw({ status: status(stored({ exportedAt: daysAgo(40) })) }))).not.toMatch(
      /[—–·!]/,
    );
  });
});

describe("the script's SHA-256 on the page", () => {
  it("shows the hash and says it matches the script this version shipped", () => {
    const html = draw({ status: status(null) });
    expect(html).toContain(`<code>${SCRIPT_SHA256}</code>`);
    expect(text(html)).toContain("It matches the script this version of Prism365 shipped with.");
    expect(html).toContain(`href="${RELEASE_TAGS_URL}"`);
    expect(html).not.toContain("/blob/main/");
    expect(text(html)).toContain("at the release tag of the Prism365 version you run");
    expect(text(html)).toContain("read it before you run it");
    expect(text(html)).toContain("run Get-FileHash on the file you downloaded");
  });

  it("names the installed version, so an older install finds its own release tag", () => {
    const html = at(NOW, () =>
      renderToStaticMarkup(
        createElement(MailboxAccessView, {
          status: status(null),
          error: "",
          importState: { state: "idle" },
          confirming: false,
          clearing: false,
          readOnly: false,
          scriptHash: { state: "ready", sha256: SCRIPT_SHA256 },
          version: "1.4.2",
          onImport: noop,
          onClear: noop,
          onConfirmClear: noop,
          onCancelClear: noop,
        }),
      ),
    );
    expect(text(html)).toContain(
      "at the release tag of the version this server runs, Prism365 1.4.2",
    );
  });

  it("warns, and says not to run it, when the served copy is different", () => {
    const changed = "0".repeat(64);
    const html = draw({ status: status(null), scriptHash: { state: "ready", sha256: changed } });
    expect(html).toContain("sev--high");
    expect(html).toContain('role="alert"');
    expect(text(html)).toMatch(
      new RegExp(`SHA-256 of this copy: ${changed} ?\\. It does not match`),
    );
    expect(text(html)).toContain("Do not run it");
    expect(text(html)).not.toContain("It matches");
  });

  it("points to Get-FileHash when the browser cannot work the hash out", () => {
    const words = text(draw({ status: status(null), scriptHash: { state: "failed" } }));
    expect(words).toContain("could not work out the SHA-256 of this copy");
    expect(words).not.toContain("It matches");
    expect(text(draw({ status: status(null), scriptHash: { state: "reading" } }))).toContain(
      "Working out the SHA-256 of this copy.",
    );
  });

  it("works the hash out the way the browser will, to the published value", async () => {
    const bytes = readFileSync(
      new URL("../public/scripts/Export-MailboxAccess.ps1", import.meta.url),
    );
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    await expect(sha256Hex(buffer as ArrayBuffer, globalThis.crypto.subtle)).resolves.toBe(
      SCRIPT_SHA256,
    );
  });
});

describe("reading an export file", () => {
  const file = (body: string, size = body.length) => ({ size, text: async () => body });

  it("removes a byte order mark and parses the JSON", async () => {
    await expect(readExportFile(file('\uFEFF{"format":"x"}'))).resolves.toEqual({ format: "x" });
  });

  it("refuses a file over 25 MB before reading it, and one that is not JSON", async () => {
    const text = vi.fn(async () => "{}");
    await expect(readExportFile({ size: 25 * 1024 * 1024 + 1, text })).rejects.toThrow(
      /at most 25 MB/,
    );
    expect(text).not.toHaveBeenCalled();
    await expect(readExportFile(file("Mailbox,Right"))).rejects.toThrow(/not JSON/);
  });
});

// ---------------------------------------------------------------------------------------------
// The person's Access tab
// ---------------------------------------------------------------------------------------------

const LAB = "eeeeeeee-0000-4000-8000-000000000001";
const BILLING = "eeeeeeee-0000-4000-8000-000000000002";
const person = (over: Record<string, unknown> = {}): PersonMailboxAccess =>
  personMailboxAccessSchema.parse({
    import: stored(),
    groups: { status: "ok", readAt: NOW, count: 4, truncated: false },
    grants: [
      {
        mailbox: {
          id: BILLING,
          address: "billing@example.invalid",
          name: "Billing",
          type: "SharedMailbox",
        },
        right: "SendOnBehalf",
        via: { kind: "direct" },
      },
      {
        mailbox: {
          id: LAB,
          address: "lab.results@example.invalid",
          name: "Lab results",
          type: "SharedMailbox",
        },
        right: "FullAccess",
        via: { kind: "group", id: "bbbbbbbb-0000-4000-8000-000000000003", name: "All plant staff" },
      },
      {
        mailbox: {
          id: LAB,
          address: "lab.results@example.invalid",
          name: "Lab results",
          type: "SharedMailbox",
        },
        right: "SendAs",
        via: { kind: "direct" },
      },
    ],
    grantsTruncated: false,
    holders: [],
    holdersTruncated: false,
    ...over,
  });
const tab = (mailbox: PersonMailboxAccess | null, error = "") =>
  at(NOW, () => renderToStaticMarkup(createElement(SharedMailboxesSection, { mailbox, error })));

describe("the Access tab's Shared mailboxes section", () => {
  it("links to Microsoft access when nothing is imported", () => {
    const html = tab(person({ import: null, groups: { status: "not_read" }, grants: [] }));
    expect(text(html)).toContain("no mailbox access export is imported");
    expect(html).toContain('<a href="#/access">Import one on Microsoft access</a>');
    expect(text(html)).toContain(
      "Microsoft 365 group mailboxes this account can use are listed under Groups and roles",
    );
  });

  it("shows a fresh export's rights, directly and through a group, with its source", () => {
    const words = text(tab(person()));
    expect(words).toMatch(/From the export of Nov 4, 2026 ?, imported yesterday/);
    expect(words).not.toContain("Older than 30 days");
    expect(words).toContain("Billing billing@example.invalid Send on Behalf directly");
    expect(words).toContain(
      "Lab results lab.results@example.invalid Full Access through All plant staff, Send As directly",
    );
    expect(words).toContain("3 grants in the export name trustees the script could not resolve");
    expect(words).toContain(
      "This is what the export says. Prism365 cannot check it against Exchange.",
    );
  });

  it("warns when the export is stale", () => {
    const html = tab(
      person({ import: stored({ exportedAt: daysAgo(45), importedAt: daysAgo(44) }) }),
    );
    expect(html).toContain("sev--medium");
    expect(text(html)).toContain("Older than 30 days");
    expect(text(html)).toContain("Access may have changed since");
  });

  it("says when only direct rights could be matched, and why", () => {
    const words = text(
      tab(
        person({
          groups: {
            status: "not_granted",
            pack: "directory.users.read",
            permission: "User.Read.All",
            enabled: false,
          },
          grants: [],
        }),
      ),
    );
    expect(words).toContain(
      "no Full Access, Send As or Send on Behalf on any shared mailbox directly.",
    );
    expect(words).toContain("Only rights granted to this account directly are shown");
    expect(words).toContain("needs the directory.users.read read pack");
  });

  it("says a shared-only export leaves out user mailboxes, the account's own included", () => {
    const words = text(tab(person({ grants: [] })));
    expect(words).toContain("no Full Access, Send As or Send on Behalf on any shared mailbox.");
    expect(words).not.toContain("on any mailbox");
    expect(words).toContain(
      "The export covers shared mailboxes only, so this account's own mailbox and other user mailboxes are not in it. Run the script with -IncludeUserMailboxes to include them.",
    );
  });

  it("speaks of every mailbox only when the export includes user mailboxes", () => {
    const words = text(tab(person({ import: stored({ includeUserMailboxes: true }), grants: [] })));
    expect(words).toContain("no Full Access, Send As or Send on Behalf on any mailbox.");
    expect(words).not.toContain("shared mailboxes only");
  });

  it("lists who can use the person's own mailbox, unresolved trustees marked", () => {
    const words = text(
      tab(
        person({
          grants: [],
          holders: [
            {
              right: "FullAccess",
              trustee: {
                id: "bbbbbbbb-0000-4000-8000-000000000003",
                address: null,
                name: "All plant staff",
                kind: "group",
              },
            },
            {
              right: "FullAccess",
              trustee: {
                id: null,
                address: null,
                name: "S-1-5-21-1111111111-2222222222-3333333333-1105",
                kind: "unresolved",
              },
            },
          ],
        }),
      ),
    );
    expect(words).toContain("Who can use this mailbox");
    expect(words).toContain("All plant staff Group Full Access");
    expect(words).toContain(
      "S-1-5-21-1111111111-2222222222-3333333333-1105 Not resolved Full Access",
    );
  });

  it("shows names from the file as text, never as markup", () => {
    const html = tab(
      person({
        grants: [
          {
            mailbox: {
              id: LAB,
              address: "lab@example.invalid",
              name: '<script>alert("x")</script>',
              type: "SharedMailbox",
            },
            right: "FullAccess",
            via: {
              kind: "group",
              id: "bbbbbbbb-0000-4000-8000-000000000003",
              name: "<img src=x onerror=alert(1)>",
            },
          },
        ],
      }),
    );
    expect(html).not.toContain("<script>");
    expect(html).not.toContain("<img");
    expect(html).toContain("&lt;script&gt;");
  });

  it("shows an error as a line, and grey rows while it loads", () => {
    expect(text(tab(null, "Shared mailbox access could not be loaded."))).toContain(
      "Shared mailbox access could not be loaded.",
    );
    expect(tab(null)).toContain("skeleton__row");
  });
});

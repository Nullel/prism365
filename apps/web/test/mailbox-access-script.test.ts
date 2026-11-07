import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { mailboxAccessExportSchema } from "@prism365/contracts";
import { afterAll, describe, expect, it, vi } from "vitest";
import { SCRIPT_SHA256 } from "../src/MailboxAccess.tsx";

// A PowerShell process takes a second or two to start.
vi.setConfig({ testTimeout: 60_000 });

// Export-MailboxAccess.ps1 runs in an administrator's own Exchange Online session, so what it may
// call is checked here on its text: Get- cmdlets and a few plain helpers only, nothing that
// connects, writes to Exchange or runs other code, and no path that looks a trustee up with an
// empty identity (Exchange answers that with every recipient in the tenant). When PowerShell is
// installed, the script also runs against stand-ins for the Exchange cmdlets, and what it writes
// must pass the import's schema. Synthetic ids and .invalid addresses only.

const scriptPath = fileURLToPath(
  new URL("../public/scripts/Export-MailboxAccess.ps1", import.meta.url),
);
const source = readFileSync(scriptPath, "utf8");

// The script's code with comments and string contents removed. Subexpressions inside expandable
// strings ("$(...)") are code and are kept. Here-strings are not used, so they are not read.
function codeOf(text: string): string {
  let out = "";
  const stack: { kind: "code" | "dq"; depth: number }[] = [{ kind: "code", depth: 0 }];
  let index = 0;
  while (index < text.length) {
    const top = stack[stack.length - 1] ?? { kind: "code", depth: 0 };
    const char = text[index] ?? "";
    const next = text[index + 1] ?? "";
    if (top.kind === "code") {
      if (char === "<" && next === "#") {
        const end = text.indexOf("#>", index + 2);
        index = end < 0 ? text.length : end + 2;
        out += " ";
        continue;
      }
      if (char === "#") {
        const end = text.indexOf("\n", index);
        index = end < 0 ? text.length : end;
        continue;
      }
      if (char === "'") {
        let end = index + 1;
        while (end < text.length) {
          if (text[end] === "'" && text[end + 1] === "'") end += 2;
          else if (text[end] === "'") break;
          else end += 1;
        }
        out += "''";
        index = end + 1;
        continue;
      }
      if (char === '"') {
        out += '"';
        stack.push({ kind: "dq", depth: 0 });
        index += 1;
        continue;
      }
      if (char === "(") top.depth += 1;
      if (char === ")") {
        if (top.depth === 0 && stack.length > 1) {
          stack.pop();
          out += ")";
          index += 1;
          continue;
        }
        top.depth -= 1;
      }
      out += char;
      index += 1;
      continue;
    }
    if (char === "`") {
      index += 2;
      continue;
    }
    if (char === '"' && next === '"') {
      index += 2;
      continue;
    }
    if (char === '"') {
      out += '"';
      stack.pop();
      index += 1;
      continue;
    }
    if (char === "$" && next === "(") {
      out += " $(";
      stack.push({ kind: "code", depth: 0 });
      index += 2;
      continue;
    }
    index += 1;
  }
  return out;
}

const code = codeOf(source);

// Commands the script may call besides Get- cmdlets: the plain helpers it needs.
const HELPERS = new Set([
  "ConvertTo-Json",
  "Set-Content",
  "Out-File",
  "Write-Host",
  "Write-Progress",
  "Write-Error",
]);
const KEYWORDS = new Set([
  "if",
  "elseif",
  "else",
  "foreach",
  "for",
  "while",
  "do",
  "switch",
  "try",
  "catch",
  "finally",
  "function",
  "param",
  "return",
  "exit",
  "continue",
  "break",
  "throw",
]);
// Static members the script may use, as written: none of them reads or writes anything but the
// value it is given (File.Exists only asks whether the output file is there).
const STATIC_MEMBERS = new Set([
  "[guid]::TryParse",
  "[guid]::Empty",
  "[string]::IsNullOrWhiteSpace",
  "[System.IO.File]::Exists",
  "[Math]::Max",
  "[System.Globalization.CultureInfo]::InvariantCulture",
  "[System.Collections.Generic.List[object]]::new",
]);
const TYPES = new Set([
  "string",
  "int",
  "bool",
  "switch",
  "object",
  "guid",
  "ref",
  "ordered",
  "Math",
  "System.IO.File",
  "System.Globalization.CultureInfo",
  "System.Collections.Generic.List[object]",
]);
const METHODS = new Set([
  "Add",
  "Trim",
  "Substring",
  "ToString",
  "ToUniversalTime",
  "ContainsKey",
  "ToArray",
  "GetUnresolvedProviderPathFromPSPath",
  "TryParse",
  "IsNullOrWhiteSpace",
  "Exists",
  "Max",
  "new",
]);

// Every command the code calls by a Verb-Noun name.
function commandsIn(text: string): string[] {
  return [...text.matchAll(/(?<![\w$.\-:])([A-Za-z]+-[A-Za-z][A-Za-z0-9]*)/g)].map(
    (match) => match[1] ?? "",
  );
}

// The first word at each place a command could start: the start of a line, and after ; | { ( =.
function commandPositions(text: string): string[] {
  return [...text.matchAll(/(?:^|[;|{(=])[ \t]*([^\s;|{}()=,]+)(?=([ \t]*=(?!=))?)/gm)].flatMap(
    (match) => {
      const word = match[1] ?? "";
      // A hashtable key or a plain assignment target.
      if (match[2] !== undefined && /^[A-Za-z]\w*$/.test(word)) return [];
      return [word];
    },
  );
}

function problemsIn(text: string): string[] {
  const problems: string[] = [];
  if (/@['"]/.test(source)) problems.push("here-strings are not read by this test");
  if (text.includes("`")) problems.push("a backtick in code can disguise a command name");
  if (/invoke|scriptblock|addtype|\biex\b/i.test(text))
    problems.push("code that runs other code (Invoke, script blocks, Add-Type)");
  // This check reads text, not PowerShell's parse tree, so it refuses whole classes of syntax the
  // script does not need rather than trying to tell safe uses apart: & covers the call operator
  // and &&, > covers every redirection (>, >>, 2>), and Get-Command finds a command to run by a
  // name that is never written out.
  if (text.includes("&")) problems.push("uses & (the call operator or &&)");
  if (text.includes(">")) problems.push("uses > (a redirection)");
  if (/\bGet-Command\b/i.test(text)) problems.push("looks commands up with Get-Command");
  for (const match of text.matchAll(/\bImport-Module\b[^\n;|]*/gi))
    if (!/^Import-Module ExchangeOnlineManagement\s*$/.test(match[0]))
      problems.push(`imports something other than ExchangeOnlineManagement: ${match[0]}`);
  for (const command of commandsIn(text)) {
    const verb = command.split("-")[0]?.toLowerCase();
    if (verb !== "get" && !HELPERS.has(command) && command !== "Import-Module")
      problems.push(`calls ${command}`);
  }
  for (const match of text.matchAll(/\bfunction\s+([^\s({]+)/gi))
    if (!/^Get-[A-Za-z]+$/.test(match[1] ?? "")) problems.push(`defines ${match[1]}`);
  for (const word of commandPositions(text)) {
    if (/^[$[@'"\d!-]/.test(word)) continue;
    if (word.startsWith("&") || word.startsWith(".")) {
      problems.push(`runs ${word} with a call or dot-source operator`);
      continue;
    }
    if (KEYWORDS.has(word.toLowerCase())) continue;
    if (/^[A-Za-z]+-[A-Za-z][A-Za-z0-9]*$/.test(word)) continue;
    problems.push(`starts a command with ${word}`);
  }
  for (const match of text.matchAll(/(\[[A-Za-z][\w.]*(?:\[[\w.]+\])?\])::(\w+)/g))
    if (!STATIC_MEMBERS.has(`${match[1]}::${match[2]}`))
      problems.push(`uses ${match[1]}::${match[2]}`);
  for (const match of text.matchAll(/(?<![\w\])$])\[([A-Za-z][\w.]*(?:\[[\w.]+\])?)\]/g))
    if (!TYPES.has(match[1] ?? "") && match[1] !== "CmdletBinding")
      problems.push(`uses the type ${match[1]}`);
  for (const match of text.matchAll(/(?:\.|::)([A-Za-z]\w*)\(/g))
    if (!METHODS.has(match[1] ?? "")) problems.push(`calls the method ${match[1]}`);
  // The only file it writes is the output file, through one Set-Content.
  const writes = [...text.matchAll(/\b(Set-Content|Out-File)\b[^\n]*/g)].map((match) => match[0]);
  if (
    writes.length !== 1 ||
    !/^Set-Content -LiteralPath \$outputFile -Value \$json /.test(writes[0] ?? "")
  )
    problems.push("writes somewhere other than the output file");
  return problems;
}

// Lines the review found getting past an earlier text check, and their neighbours. Each is appended
// to the real script and must be refused, by the text check always and by the parse-tree check
// where PowerShell is installed. They are only ever parsed, never run.
const BYPASSES = [
  "$c = Get-Command -Verb Set -Noun Mailbox -ErrorAction Ignore; if ($c) { return & $c -Identity $tenantId -HiddenFromAddressListsEnabled $true }",
  "Get-Date && & $c",
  "Write-Host $json > C:\\elsewhere.json",
  "Write-Host $json >> C:\\elsewhere.json",
  "Get-Date 2> C:\\elsewhere.log",
  "Import-Module $env:TEMP\\evil.psm1",
  "Import-Module ExchangeOnlineManagement; Import-Module .\\other.psm1",
  ". $path",
];

describe("Export-MailboxAccess.ps1, as written", () => {
  it("calls only Get- cmdlets and the plain helpers, and never connects", () => {
    expect(problemsIn(code)).toEqual([]);
    const commands = new Set(commandsIn(code));
    for (const needed of [
      "Get-EXOMailbox",
      "Get-EXOMailboxPermission",
      "Get-EXORecipientPermission",
      "Get-EXORecipient",
      "Get-ConnectionInformation",
    ])
      expect(commands, needed).toContain(needed);
    expect([...commands].filter((command) => /^(Dis)?connect-/i.test(command))).toEqual([]);
  });

  it("catches what it must refuse", () => {
    const refused = [
      "Set-Mailbox -Identity $id",
      "Add-MailboxPermission -Identity $id",
      "Remove-MailboxPermission -Identity $id",
      "New-Object System.Net.WebClient",
      "Connect-ExchangeOnline",
      "Invoke-RestMethod https://example.invalid",
      "Invoke-Expression $text",
      "$ExecutionContext.InvokeCommand.InvokeScript($text)",
      "iex $text",
      "& $script",
      ". .\\other.ps1",
      "sc -Path x -Value y",
      "rm $outputFile",
      "Re`move-Item x",
      'Write-Host "$(Remove-Item x)"',
      "[System.IO.File]::WriteAllText($path, $json)",
      "[System.Net.WebClient]::new()",
      "Out-File -FilePath C:\\elsewhere.json",
      "function Set-Thing { }",
    ];
    for (const line of [...refused, ...BYPASSES])
      expect(problemsIn(codeOf(`${source}\n${line}\n`)).length, line).toBeGreaterThan(0);
  });

  it("lets through the one import it may need, and nothing longer", () => {
    expect(problemsIn(codeOf(`${source}\nImport-Module ExchangeOnlineManagement\n`))).toEqual([]);
    expect(
      problemsIn(codeOf(`${source}\nImport-Module ExchangeOnlineManagement -Prefix X\n`)).length,
    ).toBeGreaterThan(0);
  });

  it("looks a trustee up only after refusing an empty value", () => {
    const calls = [...code.matchAll(/Get-EXORecipient\b(?!Permission)/g)];
    expect(calls).toHaveLength(1);
    const body = /function Get-ResolvedTrustee\(\[object\]\$Raw\) \{([\s\S]*?)\n\}/.exec(code)?.[1];
    expect(body).toBeDefined();
    const guard = body?.indexOf("if ([string]::IsNullOrWhiteSpace($value)) { return $null }") ?? -1;
    const call = body?.indexOf("Get-EXORecipient -Identity $value") ?? -1;
    expect(guard).toBeGreaterThan(-1);
    expect(call).toBeGreaterThan(guard);
    // The value is the trimmed, bounded text, and nothing between the guard and the call changes it.
    expect(body?.slice(guard, call)).not.toMatch(/\$value\s*=/);
  });

  it("reads each mailbox's permissions only by a checked directory id", () => {
    const loop = code.slice(code.indexOf("foreach ($mailbox in $found) {\n  $index"));
    const check = loop.indexOf("if ($null -eq $id -or $null -eq $address");
    for (const call of [
      "Get-EXOMailboxPermission -Identity $id ",
      "Get-EXORecipientPermission -Identity $id -AccessRights SendAs ",
    ]) {
      expect(loop.indexOf(call), call).toBeGreaterThan(check);
      expect(check).toBeGreaterThan(-1);
    }
    expect(code).toMatch(/\$id = Get-DirectoryId -Value \(Get-FieldValue -Item \$mailbox/);
    // Get-DirectoryId answers only a non-empty GUID, or $null.
    expect(code).toMatch(
      /\[guid\]::TryParse\(\[string\]\$Value, \[ref\]\$parsed\) -and \$parsed -ne \[guid\]::Empty/,
    );
  });

  it("writes the fields the import reads, under the names it reads them by", () => {
    const keys = (block: string) => [...block.matchAll(/(\w+)\s*=/g)].map((match) => match[1]);
    const documentBlock = /\$document = \[ordered\]@\{([\s\S]*?)\n\}/.exec(code)?.[1] ?? "";
    expect(keys(documentBlock)).toEqual(Object.keys(mailboxAccessExportSchema.shape));
    const mailboxBlock = /\$mailboxes\.Add\(\[ordered\]@\{([\s\S]*?)\}\)/.exec(code)?.[1] ?? "";
    expect(keys(mailboxBlock)).toEqual(["id", "address", "displayName", "type"]);
    const grantBlock = /\$grants\.Add\(\[ordered\]@\{([\s\S]*?)\}\)/.exec(code)?.[1] ?? "";
    expect(keys(grantBlock)).toEqual(["mailboxId", "right", "trustee"]);
    for (const block of code.matchAll(/\$trustee = \[ordered\]@\{([\s\S]*?)\}/g))
      expect(keys(block[1] ?? "")).toEqual(["id", "address", "name", "kind"]);
  });

  it("has the SHA-256 that docs/SETUP.md publishes and the page compares against", () => {
    // Administrators check the file they download against the published hash before running it,
    // so a change to the script has to change both, on purpose.
    const sha256 = createHash("sha256").update(readFileSync(scriptPath)).digest("hex");
    const setup = readFileSync(new URL("../../../docs/SETUP.md", import.meta.url), "utf8");
    const published = /SHA-256 of Export-MailboxAccess\.ps1: `([0-9a-f]{64})`/.exec(setup)?.[1];
    expect(published, "the hash published in docs/SETUP.md").toBe(sha256);
    expect(SCRIPT_SHA256, "SCRIPT_SHA256 in apps/web/src/MailboxAccess.tsx").toBe(sha256);
  });

  it("matches the schema on a sample written by hand, as PowerShell would write it", () => {
    // ConvertTo-Json -Compress, with Windows PowerShell 5.1's byte order mark, which the page
    // removes before parsing.
    const sample =
      '\uFEFF{"format":"prism365.mailbox-access","version":1,"tenantId":"11111111-1111-4111-8111-111111111111","exportedAt":"2026-11-05T09:30:00Z","includeUserMailboxes":false,"mailboxes":[{"id":"eeeeeeee-0000-4000-8000-000000000001","address":"lab.results@example.invalid","displayName":"Lab results","type":"SharedMailbox"}],"grants":[{"mailboxId":"eeeeeeee-0000-4000-8000-000000000001","right":"FullAccess","trustee":{"id":"aaaaaaaa-0000-4000-8000-000000000001","address":"pat.pell@example.invalid","name":"Pat Pell","kind":"user"}},{"mailboxId":"eeeeeeee-0000-4000-8000-000000000001","right":"SendAs","trustee":{"id":null,"address":null,"name":"S-1-5-21-1111111111-2222222222-3333333333-1105","kind":"unresolved"}}],"unresolved":1}';
    const parsed = mailboxAccessExportSchema.safeParse(JSON.parse(sample.replace(/^\uFEFF/, "")));
    expect(parsed.success, JSON.stringify(parsed.error?.issues[0])).toBe(true);
  });
});

// ---------------------------------------------------------------------------------------------
// A run against stand-ins for the Exchange cmdlets
// ---------------------------------------------------------------------------------------------

const pwsh = spawnSync("pwsh", ["-NoLogo", "-NoProfile", "-NonInteractive", "-Command", "exit 0"], {
  env: { ...process.env, POWERSHELL_TELEMETRY_OPTOUT: "1", POWERSHELL_UPDATECHECK: "Off" },
  stdio: "ignore",
  timeout: 30_000,
});
const hasPwsh = pwsh.status === 0;

const TENANT = "11111111-1111-4111-8111-111111111111";
const LAB = "eeeeeeee-0000-4000-8000-000000000001";
const BILLING = "eeeeeeee-0000-4000-8000-000000000002";
const USERBOX = "eeeeeeee-0000-4000-8000-000000000003";
const PERSON = "aaaaaaaa-0000-4000-8000-000000000001";
const OTHER = "aaaaaaaa-0000-4000-8000-000000000002";
const ROBIN = "aaaaaaaa-0000-4000-8000-000000000003";
const GROUP = "bbbbbbbb-0000-4000-8000-000000000003";
const SID = "S-1-5-21-1111111111-2222222222-3333333333-1105";

const permission = (user: string, over: Record<string, unknown> = {}) => ({
  Identity: "Lab results",
  User: user,
  AccessRights: ["FullAccess"],
  IsInherited: false,
  Deny: false,
  ...over,
});
const recipient = (
  keys: string[],
  id: string | null,
  name: string,
  type: string,
  address: string,
) => ({
  keys,
  ExternalDirectoryObjectId: id,
  PrimarySmtpAddress: address,
  DisplayName: name,
  RecipientTypeDetails: type,
});
const fixture = (over: Record<string, unknown> = {}) => ({
  connections: [{ State: "Connected", TenantID: TENANT.toUpperCase() }],
  deny: null,
  mailboxes: [
    {
      ExternalDirectoryObjectId: LAB,
      PrimarySmtpAddress: "lab.results@example.invalid",
      DisplayName: "Lab results",
      RecipientTypeDetails: "SharedMailbox",
      GrantSendOnBehalfTo: ["Pat Pell", "", "   ", null],
    },
    {
      ExternalDirectoryObjectId: BILLING,
      PrimarySmtpAddress: "billing@example.invalid",
      DisplayName: "Billing",
      RecipientTypeDetails: "SharedMailbox",
      GrantSendOnBehalfTo: [],
    },
    // No directory id: left out, never read by an empty identity.
    {
      ExternalDirectoryObjectId: "",
      PrimarySmtpAddress: "noid@example.invalid",
      DisplayName: "No id",
      RecipientTypeDetails: "SharedMailbox",
      GrantSendOnBehalfTo: [],
    },
    {
      ExternalDirectoryObjectId: USERBOX,
      PrimarySmtpAddress: "robin.reyes@example.invalid",
      DisplayName: "Robin Reyes",
      RecipientTypeDetails: "UserMailbox",
      GrantSendOnBehalfTo: [],
    },
  ],
  fullAccess: {
    [LAB]: [
      permission("NT AUTHORITY\\SELF", { AccessRights: ["FullAccess", "ReadPermission"] }),
      permission("pat.pell@example.invalid"),
      // The same person by another name: one grant.
      permission("Pat Pell"),
      permission("Inherited admin", { IsInherited: true }),
      permission("Denied one", { Deny: true }),
      permission("Read only", { AccessRights: ["ReadPermission"] }),
      permission(SID),
      permission("All plant staff"),
      // Two recipients answer to this name: unresolved, never a guess.
      permission("Pat"),
      permission(""),
    ],
    [BILLING]: [permission("Robin Reyes")],
  },
  sendAs: {
    [LAB]: [
      {
        Identity: "Lab results",
        Trustee: "Pat Pell",
        AccessControlType: "Allow",
        AccessRights: ["SendAs"],
        IsInherited: false,
      },
      {
        Identity: "Lab results",
        Trustee: "Robin Reyes",
        AccessControlType: "Deny",
        AccessRights: ["SendAs"],
        IsInherited: false,
      },
      {
        Identity: "Lab results",
        Trustee: "NT AUTHORITY\\SELF",
        AccessControlType: "Allow",
        AccessRights: ["SendAs"],
        IsInherited: false,
      },
      {
        Identity: "Lab results",
        Trustee: "Inherited admin",
        AccessControlType: "Allow",
        AccessRights: ["SendAs"],
        IsInherited: true,
      },
    ],
  },
  recipients: [
    recipient(
      ["pat.pell@example.invalid", "Pat Pell", "Pat"],
      PERSON,
      "Pat Pell",
      "UserMailbox",
      "pat.pell@example.invalid",
    ),
    recipient(["Pat"], OTHER, "Pat Quinn", "UserMailbox", "pat.quinn@example.invalid"),
    recipient(["Robin Reyes"], ROBIN, "Robin Reyes", "UserMailbox", "robin.reyes@example.invalid"),
    recipient(
      ["All plant staff"],
      GROUP,
      "All plant staff",
      "MailUniversalSecurityGroup",
      "plant.staff@example.invalid",
    ),
    recipient(
      ["lab.results@example.invalid"],
      LAB,
      "Lab results",
      "SharedMailbox",
      "lab.results@example.invalid",
    ),
    recipient(
      ["billing@example.invalid"],
      BILLING,
      "Billing",
      "SharedMailbox",
      "billing@example.invalid",
    ),
  ],
  ...over,
});

// A stand-in ExchangeOnlineManagement 3.9.0 whose cmdlets answer from the fixture and log each
// call, and which answers an empty Get-EXORecipient identity with every recipient, as Exchange does.
const MODULE = `
$fixture = Get-Content -Raw -LiteralPath $env:PRISM365_FIXTURE | ConvertFrom-Json
function Write-Call([string]$Text) { Add-Content -LiteralPath $env:PRISM365_LOG -Value $Text }
function Get-ConnectionInformation { $fixture.connections }
function Get-EXOMailbox {
  param([string[]]$RecipientTypeDetails, $ResultSize, [string[]]$Properties)
  Write-Call "Get-EXOMailbox $($RecipientTypeDetails -join ',') $($Properties -join ',') $ResultSize"
  $fixture.mailboxes | Where-Object { $RecipientTypeDetails -contains $_.RecipientTypeDetails }
}
function Get-EXOMailboxPermission {
  param([string]$Identity, $ResultSize)
  Write-Call "Get-EXOMailboxPermission [$Identity]"
  if ([string]::IsNullOrWhiteSpace($Identity)) { Write-Call 'EMPTY'; return }
  if ($fixture.deny -eq 'Get-EXOMailboxPermission') { throw 'The operation could not be performed because Get-MailboxPermission is not allowed for your role.' }
  $found = $fixture.fullAccess.$Identity
  if ($null -ne $found) { $found }
}
function Get-EXORecipientPermission {
  param([string]$Identity, [string]$AccessRights, $ResultSize)
  Write-Call "Get-EXORecipientPermission [$Identity] $AccessRights"
  if ([string]::IsNullOrWhiteSpace($Identity)) { Write-Call 'EMPTY'; return }
  $found = $fixture.sendAs.$Identity
  if ($null -ne $found) { $found }
}
function Get-EXORecipient {
  param([string]$Identity, [string[]]$Properties)
  Write-Call "Get-EXORecipient [$Identity]"
  if ([string]::IsNullOrWhiteSpace($Identity)) { Write-Call 'EMPTY'; return $fixture.recipients }
  if ($null -ne $fixture.throwing -and $null -ne $fixture.throwing.$Identity) { throw $fixture.throwing.$Identity }
  $match = @($fixture.recipients | Where-Object { $_.keys -contains $Identity })
  if ($match.Count -eq 0) { throw "The object $Identity couldn't be found." }
  $match
}
`;
const RUNNER = `
param([string]$ModulePath, [string]$Script, [string]$Output, [switch]$IncludeUserMailboxes, [switch]$Force)
Import-Module -Name $ModulePath
& $Script -OutputPath $Output -IncludeUserMailboxes:$IncludeUserMailboxes -Force:$Force
exit $LASTEXITCODE
`;

// Reads each file with PowerShell's own parser, never running it, and lists what the script may
// not do: a command whose name is not written out, or is not on the list; the & or . invocation
// operators; any redirection; Import-Module of anything but ExchangeOnlineManagement; a function
// other than the script's own helpers.
const ANALYZER = `
param([string]$List)
$Paths = @(Get-Content -LiteralPath $List)
$commands = @('Get-Module', 'Get-ConnectionInformation', 'Get-Date', 'Get-EXOMailbox',
  'Get-EXOMailboxPermission', 'Get-EXORecipientPermission', 'Get-EXORecipient', 'Write-Host',
  'Write-Progress', 'Write-Error', 'ConvertTo-Json', 'Set-Content')
$helpers = @('Get-BoundedText', 'Get-FieldValue', 'Get-IsTrue', 'Get-DirectoryId', 'Get-Address',
  'Get-ResolvedTrustee')
$results = [ordered]@{}
foreach ($path in $Paths) {
  $tokens = $null
  $errors = $null
  $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
  $problems = [System.Collections.Generic.List[string]]::new()
  foreach ($parseError in @($errors)) { if ($null -ne $parseError) { $problems.Add("parse error: $($parseError.Message)") } }
  foreach ($node in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($helpers -notcontains $node.Name) { $problems.Add("defines $($node.Name)") }
  }
  foreach ($node in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
    $line = $node.Extent.StartLineNumber
    if ($node.InvocationOperator -ne [System.Management.Automation.Language.TokenKind]::Unknown) {
      $problems.Add("line $line uses the $($node.InvocationOperator) invocation operator")
    }
    $name = $node.GetCommandName()
    if ($null -eq $name -or $node.CommandElements[0] -isnot [System.Management.Automation.Language.StringConstantExpressionAst]) {
      $problems.Add("line $line runs a command whose name is not written out")
      continue
    }
    if ($name -eq 'Import-Module') {
      $elements = $node.CommandElements
      $literal = $elements.Count -eq 2 -and $elements[1] -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $elements[1].Value -eq 'ExchangeOnlineManagement'
      if (-not $literal) { $problems.Add("line $line imports something other than ExchangeOnlineManagement") }
      continue
    }
    if ($commands -cnotcontains $name -and $helpers -cnotcontains $name) { $problems.Add("line $line calls $name") }
  }
  foreach ($node in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.RedirectionAst] }, $true)) {
    $problems.Add("line $($node.Extent.StartLineNumber) redirects output")
  }
  $results[$path] = $problems.ToArray()
}
ConvertTo-Json -InputObject $results -Depth 4 -Compress
`;

const work = mkdtempSync(join(tmpdir(), "prism365-export-test-"));
afterAll(() => rmSync(work, { recursive: true, force: true }));
let runs = 0;

// What the analyzer finds in each file, by path.
function analyze(paths: string[]): Record<string, string[]> {
  const analyzer = join(work, "analyzer.ps1");
  const list = join(work, "analyze.txt");
  writeFileSync(analyzer, ANALYZER);
  writeFileSync(list, `${paths.join("\n")}\n`);
  const result = spawnSync(
    "pwsh",
    ["-NoLogo", "-NoProfile", "-NonInteractive", "-File", analyzer, "-List", list],
    {
      encoding: "utf8",
      timeout: 60_000,
      env: { ...process.env, POWERSHELL_TELEMETRY_OPTOUT: "1", POWERSHELL_UPDATECHECK: "Off" },
    },
  );
  if (result.status !== 0) throw new Error(`The analyzer failed: ${result.stdout}${result.stderr}`);
  return JSON.parse(result.stdout) as Record<string, string[]>;
}

function runScript(
  data: object,
  options: { includeUserMailboxes?: boolean; existing?: boolean; force?: boolean } = {},
) {
  runs += 1;
  const dir = join(work, `run-${runs}`);
  const moduleDir = join(dir, "ExchangeOnlineManagement");
  mkdirSync(moduleDir, { recursive: true });
  writeFileSync(
    join(moduleDir, "ExchangeOnlineManagement.psd1"),
    "@{ ModuleVersion = '3.9.0'; GUID = 'f1e2d3c4-b5a6-4978-8a9b-0c1d2e3f4a5b'; RootModule = 'ExchangeOnlineManagement.psm1'; FunctionsToExport = @('Get-ConnectionInformation','Get-EXOMailbox','Get-EXOMailboxPermission','Get-EXORecipientPermission','Get-EXORecipient') }\n",
  );
  writeFileSync(join(moduleDir, "ExchangeOnlineManagement.psm1"), MODULE);
  writeFileSync(join(dir, "runner.ps1"), RUNNER);
  writeFileSync(join(dir, "fixture.json"), JSON.stringify(data));
  const output = join(dir, "out.json");
  const log = join(dir, "calls.log");
  writeFileSync(log, "");
  if (options.existing) writeFileSync(output, "{}");
  const result = spawnSync(
    "pwsh",
    [
      "-NoLogo",
      "-NoProfile",
      "-NonInteractive",
      "-File",
      join(dir, "runner.ps1"),
      "-ModulePath",
      join(moduleDir, "ExchangeOnlineManagement.psd1"),
      "-Script",
      scriptPath,
      "-Output",
      output,
      ...(options.includeUserMailboxes ? ["-IncludeUserMailboxes"] : []),
      ...(options.force ? ["-Force"] : []),
    ],
    {
      encoding: "utf8",
      timeout: 60_000,
      env: {
        ...process.env,
        POWERSHELL_TELEMETRY_OPTOUT: "1",
        POWERSHELL_UPDATECHECK: "Off",
        PRISM365_FIXTURE: join(dir, "fixture.json"),
        PRISM365_LOG: log,
      },
    },
  );
  const calls = readFileSync(log, "utf8").split(/\r?\n/).filter(Boolean);
  const written = existsSync(output) ? readFileSync(output, "utf8").replace(/^\uFEFF/, "") : null;
  return { status: result.status, out: `${result.stdout}${result.stderr}`, calls, written };
}

describe.skipIf(!hasPwsh)("Export-MailboxAccess.ps1, run against stand-in cmdlets", () => {
  it("writes an export the import accepts, with only allowed grants", () => {
    const run = runScript(fixture());
    expect(run.status, run.out).toBe(0);
    const parsed = mailboxAccessExportSchema.safeParse(JSON.parse(run.written ?? "null"));
    expect(parsed.success, JSON.stringify(parsed.error?.issues[0])).toBe(true);
    const file = parsed.data;
    if (!file) return;
    expect(file.tenantId).toBe(TENANT);
    expect(Date.now() - Date.parse(file.exportedAt)).toBeLessThan(10 * 60_000);
    expect(file.includeUserMailboxes).toBe(false);
    expect(file.mailboxes.map((mailbox) => mailbox.id)).toEqual([LAB, BILLING]);
    const grants = file.grants.map((grant) => [
      grant.mailboxId,
      grant.right,
      grant.trustee.kind,
      grant.trustee.id ?? grant.trustee.name,
    ]);
    expect(grants).toEqual([
      [LAB, "FullAccess", "user", PERSON],
      [LAB, "FullAccess", "unresolved", SID],
      [LAB, "FullAccess", "group", GROUP],
      [LAB, "FullAccess", "unresolved", "Pat"],
      [LAB, "SendAs", "user", PERSON],
      [LAB, "SendOnBehalf", "user", PERSON],
      [BILLING, "FullAccess", "user", ROBIN],
    ]);
    expect(file.unresolved).toBe(2);
  });

  it("never asks Exchange about an empty identity, and looks each trustee up once", () => {
    const run = runScript(fixture());
    expect(run.status, run.out).toBe(0);
    expect(run.calls).not.toContain("EMPTY");
    expect(run.calls.filter((call) => call.startsWith("Get-EXORecipient ["))).toEqual([
      "Get-EXORecipient [lab.results@example.invalid]",
      "Get-EXORecipient [pat.pell@example.invalid]",
      "Get-EXORecipient [Pat Pell]",
      `Get-EXORecipient [${SID}]`,
      "Get-EXORecipient [All plant staff]",
      "Get-EXORecipient [Pat]",
      "Get-EXORecipient [Robin Reyes]",
    ]);
    expect(run.calls).toContain("Get-EXOMailbox SharedMailbox GrantSendOnBehalfTo Unlimited");
    expect(run.calls).toContain(`Get-EXORecipientPermission [${LAB}] SendAs`);
  });

  it("adds user mailboxes only when asked", () => {
    const run = runScript(fixture(), { includeUserMailboxes: true });
    expect(run.status, run.out).toBe(0);
    const file = mailboxAccessExportSchema.parse(JSON.parse(run.written ?? "null"));
    expect(file.includeUserMailboxes).toBe(true);
    expect(file.mailboxes.map((mailbox) => mailbox.type)).toEqual([
      "SharedMailbox",
      "SharedMailbox",
      "UserMailbox",
    ]);
  });

  it("refuses to run on no connection or on two", () => {
    for (const connections of [
      [],
      [
        { State: "Connected", TenantID: TENANT },
        { State: "Connected", TenantID: "99999999-9999-4999-8999-999999999999" },
      ],
    ]) {
      const run = runScript(fixture({ connections }));
      expect(run.status).toBe(1);
      expect(run.out).toMatch(
        new RegExp(`Found ${connections.length} Exchange Online connections`),
      );
      expect(run.written).toBeNull();
      expect(run.calls.some((call) => call.startsWith("Get-EXOMailbox"))).toBe(false);
    }
  });

  it("stops with the role check, writing nothing, when a cmdlet is not allowed", () => {
    const run = runScript(fixture({ deny: "Get-EXOMailboxPermission" }));
    expect(run.status).toBe(1);
    expect(run.out).toContain("Get-ManagementRole -Cmdlet Get-MailboxPermission");
    expect(run.out).toContain("Nothing was written");
    expect(run.written).toBeNull();
  });

  it("stops when no trustee can be resolved, rather than writing everyone as unresolved", () => {
    const run = runScript(fixture({ recipients: [] }));
    expect(run.status).toBe(1);
    expect(run.out).toContain("Get-EXORecipient could not look up the mailbox");
    expect(run.written).toBeNull();
  });

  it("does not replace an existing file without -Force", () => {
    const kept = runScript(fixture(), { existing: true });
    expect(kept.status).toBe(1);
    expect(kept.written).toBe("{}");
    const replaced = runScript(fixture(), { existing: true, force: true });
    expect(replaced.status, replaced.out).toBe(0);
    expect(
      mailboxAccessExportSchema.safeParse(JSON.parse(replaced.written ?? "null")).success,
    ).toBe(true);
  });

  it("keeps a trustee Exchange says matches several recipients as unresolved", () => {
    const data = fixture({
      throwing: {
        Lee: "The operation couldn't be performed because 'Lee' matches multiple entries.",
      },
    });
    (data.fullAccess as Record<string, unknown[]>)[BILLING] = [permission("Lee")];
    const run = runScript(data);
    expect(run.status, run.out).toBe(0);
    const file = mailboxAccessExportSchema.parse(JSON.parse(run.written ?? "null"));
    expect(
      file.grants.filter((grant) => grant.mailboxId === BILLING).map((grant) => grant.trustee),
    ).toEqual([{ id: null, address: null, name: "Lee", kind: "unresolved" }]);
  });

  it("stops, writing nothing, when looking a trustee up fails for any other reason", () => {
    const run = runScript(
      fixture({ throwing: { "Robin Reyes": "The session has been throttled. Try again later." } }),
    );
    expect(run.status).toBe(1);
    expect(run.out).toContain("Get-EXORecipient failed for the trustee Robin Reyes");
    expect(run.out).toContain("Nothing was written");
    expect(run.written).toBeNull();
  });
});

describe.skipIf(!hasPwsh)("Export-MailboxAccess.ps1, read by PowerShell's parser", () => {
  it("calls only listed commands, by name, with no invocation operator or redirection", () => {
    expect(analyze([scriptPath])).toEqual({ [scriptPath]: [] });
  });

  it("refuses every bypass the text check must also refuse", () => {
    const paths = BYPASSES.map((line, index) => {
      const path = join(work, `bypass-${index}.ps1`);
      writeFileSync(path, `${source}\n${line}\n`);
      return path;
    });
    const found = analyze(paths);
    for (const [index, line] of BYPASSES.entries())
      expect(found[paths[index] ?? ""]?.length ?? 0, line).toBeGreaterThan(0);
  });

  it("lets through Import-Module of ExchangeOnlineManagement by its literal name only", () => {
    const allowed = join(work, "import-allowed.ps1");
    const refused = join(work, "import-refused.ps1");
    writeFileSync(allowed, `${source}\nImport-Module ExchangeOnlineManagement\n`);
    writeFileSync(refused, `${source}\nImport-Module -Name $module\n`);
    const found = analyze([allowed, refused]);
    expect(found[allowed]).toEqual([]);
    expect(found[refused]?.length ?? 0).toBeGreaterThan(0);
  });
});

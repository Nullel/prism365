// What a pilot tenant must provide for every built-in audit and collector, generated from the
// code so docs/LIVE_PILOT.md cannot drift from it. Read-only: no network, no tenant, no env.
//
//   npm run pilot:matrix              print the generated section
//   npm run pilot:matrix -- --write   replace the section in docs/LIVE_PILOT.md
//   npm run pilot:matrix -- --check   exit 1 when the section in docs/LIVE_PILOT.md is stale
import { readFileSync, writeFileSync } from "node:fs";
import { auditDefinition, builtInAudits } from "@prism365/audits";
import { createDirectoryChangesModule } from "@prism365/module-directory-changes";
import {
  createMessageCenterModule,
  createServiceHealthModule,
} from "@prism365/module-service-health";
import { createSigninsAuditModule } from "@prism365/module-signins-audit";
import { supportedReadFeatures } from "../../apps/api/src/config.ts";

const root = new URL("../../", import.meta.url);
const docPath = new URL("docs/LIVE_PILOT.md", root);
const registryPath = new URL("apps/api/src/modules.ts", root);
const START = "<!-- pilot-matrix:start -->";
const END = "<!-- pilot-matrix:end -->";

// The modules apps/api/src/modules.ts registers. --check fails when that file registers more.
const modules = [
  createSigninsAuditModule(),
  createServiceHealthModule(),
  createMessageCenterModule(),
  createDirectoryChangesModule(),
];

type Feature = keyof typeof supportedReadFeatures;
const permissionOf = (feature: string) => supportedReadFeatures[feature as Feature] ?? feature;
const code = (items: readonly string[]) =>
  items.length > 0 ? items.map((item) => `\`${item}\``).join(", ") : "none";
const cell = (text: string) => text.replaceAll("|", "\\|");

function generate(): string {
  const required = new Set<string>();
  const optional = new Set<string>();
  const lines: string[] = [];

  lines.push(`#### Audits (${builtInAudits.length})`, "");
  lines.push("| Audit | Read packs | Permissions | Optional permissions | Roles and licences |");
  lines.push("|---|---|---|---|---|");
  for (const audit of builtInAudits) {
    const definition = auditDefinition(audit);
    for (const pack of audit.packs) required.add(pack);
    for (const pack of audit.optionalPacks) optional.add(pack);
    lines.push(
      `| ${cell(definition.title)} | ${code(audit.packs)} | ${code(definition.permissions)} | ${code(definition.optionalPermissions)} | ${cell(definition.requirements.join(" ")) || "none"} |`,
    );
  }

  lines.push("", `#### Collectors (${modules.length})`, "");
  lines.push("| Module | Read packs | Permissions | Requires |");
  lines.push("|---|---|---|---|");
  for (const module of modules) {
    for (const pack of module.packs) required.add(pack);
    const requires = (module as { requires?: readonly string[] }).requires ?? [];
    lines.push(
      `| ${cell(module.title)} | ${code(module.packs)} | ${code(module.packs.map(permissionOf))} | ${code(requires)} |`,
    );
  }

  const optionalOnly = [...optional].filter((pack) => !required.has(pack)).sort();
  // Supported features no audit or collector names are read elsewhere (People pages, the
  // investigation view), so the full-suite line keeps them.
  const elsewhere = Object.keys(supportedReadFeatures)
    .filter((feature) => !required.has(feature) && !optional.has(feature))
    .sort();
  const fullSuite = [...new Set([...required, ...elsewhere])].sort();
  const permissions = [...new Set(["User.Read", ...fullSuite.map(permissionOf)])].sort();
  const optionalPermissions = [...new Set(optionalOnly.map(permissionOf))]
    .filter((permission) => !permissions.includes(permission))
    .sort();
  const graphUrl = (permission: string) => `https://graph.microsoft.com/${permission}`;

  lines.push("", "#### `ENABLED_READ_FEATURES` for the full suite", "");
  lines.push(
    `Every pack an audit or collector requires, plus the packs other pages read (${code(elsewhere)}):`,
    "",
  );
  lines.push("```", `ENABLED_READ_FEATURES=${fullSuite.join(",")}`, "```", "");
  lines.push(
    `Optional only (an audit runs without them and shows fewer columns): ${code(optionalOnly)}. Add one only for what it adds.`,
  );

  lines.push("", "#### Delegated permissions for the registration", "");
  lines.push(`${code(permissions)}.`, "");
  lines.push("```bash");
  lines.push(
    `m365 entra app add --name "Prism365 (pilot)" --platform web --redirectUris "http://localhost:5173/api/v1/auth/callback" --apisDelegated "${permissions.map(graphUrl).join(",")}" --grantAdminConsent --withSecret --output json`,
  );
  lines.push("```");
  if (optionalPermissions.length > 0)
    lines.push(
      "",
      `For each optional pack you enable, append its permission to \`--apisDelegated\`: ${code(optionalPermissions.map(graphUrl))}.`,
    );
  return lines.join("\n");
}

function registeredModuleCount(): number {
  return (readFileSync(registryPath, "utf8").match(/\.register\(create\w+Module\(/g) ?? []).length;
}

function replaceSection(doc: string, body: string): string {
  const start = doc.indexOf(START);
  const end = doc.indexOf(END);
  if (start === -1 || end === -1 || end < start)
    throw new Error(`docs/LIVE_PILOT.md has no ${START} … ${END} section.`);
  return `${doc.slice(0, start + START.length)}\n${body}\n${doc.slice(end)}`;
}

const mode = process.argv[2] ?? "";
const body = generate();
if (mode === "--write") {
  writeFileSync(docPath, replaceSection(readFileSync(docPath, "utf8"), body));
  console.log("docs/LIVE_PILOT.md: pilot matrix section updated.");
} else if (mode === "--check") {
  const problems: string[] = [];
  const doc = readFileSync(docPath, "utf8");
  if (replaceSection(doc, body) !== doc)
    problems.push("docs/LIVE_PILOT.md is stale: run `npm run pilot:matrix -- --write`.");
  if (registeredModuleCount() !== modules.length)
    problems.push(
      `apps/api/src/modules.ts registers ${registeredModuleCount()} modules; scripts/pilot/matrix.ts lists ${modules.length}.`,
    );
  if (problems.length > 0) {
    for (const problem of problems) console.error(problem);
    process.exit(1);
  }
  console.log("Pilot matrix is current.");
} else if (mode === "") {
  console.log(body);
} else {
  console.error(`Unknown option ${mode}. Use --write or --check.`);
  process.exit(2);
}

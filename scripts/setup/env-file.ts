import { randomBytes } from "node:crypto";
import { chmod, open, realpath, rename, unlink } from "node:fs/promises";
import { basename, dirname, join } from "node:path";

// The keys `npm run setup` owns in .env. Every other line, comment and the order of the file
// stay as they are.
export const MANAGED_KEYS = [
  "DATA_MODE",
  "ENTRA_TENANT_ID",
  "ENTRA_CLIENT_ID",
  "ENTRA_ADMIN_OBJECT_IDS",
  "ENABLED_READ_FEATURES",
  "PUBLIC_ORIGIN",
  "ENTRA_CERTIFICATE_PATH",
  "ENTRA_PRIVATE_KEY_PATH",
] as const;
export type ManagedKey = (typeof MANAGED_KEYS)[number];
export type ManagedValues = Record<ManagedKey, string>;

export interface EnvChange {
  key: ManagedKey;
  before: string | null;
  after: string;
}

const ASSIGNMENT = /^(\s*(?:export\s+)?)([A-Za-z_][A-Za-z0-9_]*)(\s*=\s*)(.*)$/;

// A value and whatever follows it on the line (a closing quote's tail, or an inline comment).
function splitValue(raw: string): { value: string; suffix: string } {
  const quote = raw[0];
  if (quote === '"' || quote === "'" || quote === "`") {
    const end = raw.indexOf(quote, 1);
    if (end > 0) return { value: raw.slice(1, end), suffix: raw.slice(end + 1) };
  }
  const comment = /\s#/.exec(raw);
  const value = comment ? raw.slice(0, comment.index) : raw;
  return { value: value.trim(), suffix: comment ? raw.slice(comment.index) : "" };
}

function lines(text: string): string[] {
  return text.length === 0 ? [] : text.split(/(?<=\n)/);
}

function content(line: string): { body: string; end: string } {
  const end = /\r?\n$/.exec(line)?.[0] ?? "";
  return { body: line.slice(0, line.length - end.length), end };
}

// Active assignments, the last one winning as it does when the file is loaded.
export function readEnv(text: string): Map<string, string> {
  const values = new Map<string, string>();
  for (const line of lines(text)) {
    const match = ASSIGNMENT.exec(content(line).body);
    if (match?.[2]) values.set(match[2], splitValue(match[4] ?? "").value);
  }
  return values;
}

// The administrator list with the signed-in administrator added once; existing IDs are kept.
export function adminObjectIds(existing: string | undefined, adminId: string): string {
  const ids = (existing ?? "")
    .split(",")
    .map((id) => id.trim().toLowerCase())
    .filter(Boolean);
  const unique = [...new Set(ids)];
  if (!unique.includes(adminId.toLowerCase())) unique.push(adminId.toLowerCase());
  return unique.join(",");
}

// Characters a plain .env value cannot carry safely: quotes and # change how it is parsed, $ is
// expanded by Docker Compose, and line breaks end it.
export function assertEnvValue(key: string, value: string) {
  if (/[\r\n"'`#$]/.test(value) || value !== value.trim() || value.length > 4_096)
    throw new Error(
      `${key} cannot be written to .env as "${value.replace(/[\r\n]/g, " ")}". Choose a path without quotes, #, $ or line breaks.`,
    );
}

// Sets the managed keys. A key already assigned is changed where it is (every assignment of it,
// so a duplicate further down cannot undo it); a key only present as a commented example
// (`# KEY=...`) gets its line right after that comment; anything else is added at the end.
export function mergeEnv(
  text: string,
  values: ManagedValues,
): { text: string; changes: EnvChange[] } {
  for (const key of MANAGED_KEYS) assertEnvValue(key, values[key]);
  const before = readEnv(text);
  const eol = text.includes("\r\n") ? "\r\n" : "\n";
  const output = lines(text);
  const assigned = new Set<string>();
  for (const [index, line] of output.entries()) {
    const { body, end } = content(line);
    const match = ASSIGNMENT.exec(body);
    const key = match?.[2] as ManagedKey | undefined;
    if (!match || !key || !(MANAGED_KEYS as readonly string[]).includes(key)) continue;
    assigned.add(key);
    const { value, suffix } = splitValue(match[4] ?? "");
    if (value === values[key]) continue;
    output[index] = `${match[1]}${key}${match[3]}${values[key]}${suffix}${end}`;
  }
  const appended: string[] = [];
  for (const key of MANAGED_KEYS) {
    if (assigned.has(key)) continue;
    const example = new RegExp(`^\\s*#\\s*(?:export\\s+)?${key}\\s*=`);
    const at = output.findIndex((line) => example.test(content(line).body));
    if (at >= 0) {
      const { body, end } = content(output[at] ?? "");
      output[at] = `${body}${end || eol}`;
      output.splice(at + 1, 0, `${key}=${values[key]}${end || eol}`);
    } else appended.push(`${key}=${values[key]}${eol}`);
  }
  if (appended.length > 0) {
    const last = output.at(-1);
    if (last !== undefined && !/\n$/.test(last)) output[output.length - 1] = `${last}${eol}`;
    // A blank line sets the added block apart, unless the file already ends with one.
    if (last !== undefined && content(last).body.trim() !== "") output.push(eol);
    output.push(`# Set by npm run setup.${eol}`, ...appended);
  }
  const changes = MANAGED_KEYS.filter((key) => before.get(key) !== values[key]).map((key) => ({
    key,
    before: before.get(key) ?? null,
    after: values[key],
  }));
  return { text: output.join(""), changes };
}

// Replaces the file in one step: a temporary file beside it, owner-only, then a rename. When
// .env is a symbolic link, the file it points to is the one replaced.
export async function writeEnvFile(path: string, text: string): Promise<void> {
  const target = await realpath(path).catch((error: NodeJS.ErrnoException) => {
    if (error.code === "ENOENT") return path;
    throw error;
  });
  const temporary = join(
    dirname(target),
    `.${basename(target)}.${randomBytes(6).toString("hex")}.tmp`,
  );
  const handle = await open(temporary, "wx", 0o600);
  try {
    try {
      await handle.writeFile(text);
      await handle.sync();
    } finally {
      await handle.close();
    }
    if (process.platform !== "win32") await chmod(temporary, 0o600);
    await rename(temporary, target);
  } catch (error) {
    await unlink(temporary).catch(() => {});
    throw error;
  }
}

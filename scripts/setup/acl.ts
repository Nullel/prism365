import { execFile } from "node:child_process";
import { access, constants } from "node:fs/promises";

// POSIX access control lists on the secrets. On Linux the Docker install's container reads the
// key as user ID 1000, and an administrator whose account has another ID gives it access with
// setfacl (docs/SETUP.md). The setup must not undo that: a chmod of the group bits rewrites the
// ACL's mask, and a rotation's new files would not carry the entries. Node has no ACL API, so this
// runs getfacl and setfacl from /usr/bin or /bin (never whatever PATH finds), with fixed argument
// lists and no shell. Where they are missing, or on other systems, the ACL is unknown.

export interface AccessList {
  // Named entries and the mask on the file itself, as getfacl prints them ("user:1000:r--").
  entries: string[];
  // Default entries a directory hands to new files ("user:1000:r--", without "default:").
  defaults: string[];
}

export interface AclTools {
  // null when the ACL cannot be read here (not Linux, or getfacl missing).
  read(path: string): Promise<AccessList | null>;
  // Adds entries to a file; false when that is not possible here.
  apply(path: string, entries: readonly string[]): Promise<boolean>;
}

const NAMED = /^(?:user|group):[^:]+:[rwx-]{3}$|^mask::[rwx-]{3}$/;

// getfacl's output, keeping only what is more than the plain mode bits.
export function parseGetfacl(output: string): AccessList {
  const entries: string[] = [];
  const defaults: string[] = [];
  for (const raw of output.split("\n")) {
    const line = raw.replace(/\s*#.*$/, "").trim();
    if (!line) continue;
    const entry = line.startsWith("default:") ? line.slice("default:".length) : line;
    if (!NAMED.test(entry)) continue;
    (line.startsWith("default:") ? defaults : entries).push(entry);
  }
  return { entries, defaults };
}

const quote = (value: string) => `'${value.replaceAll("'", "'\\''")}'`;
const named = (entries: readonly string[]) =>
  entries.filter((entry) => !entry.startsWith("mask::"));

// The command an administrator can run by hand when the setup could not re-apply an ACL. Every
// entry and the path are single-quoted: a user or group name can hold shell characters.
export function setfaclCommand(path: string, entries: readonly string[]): string {
  return `setfacl ${named(entries)
    .map((entry) => `-m ${quote(entry)}`)
    .join(" ")} ${quote(path)}`;
}

function run(command: string, args: string[]): Promise<string | null> {
  return new Promise((resolve) => {
    execFile(command, args, { timeout: 10_000, maxBuffer: 1_048_576 }, (error, stdout) =>
      resolve(error ? null : stdout),
    );
  });
}

export const ACL_TOOL_DIRECTORIES = ["/usr/bin", "/bin"] as const;

async function tool(name: "getfacl" | "setfacl"): Promise<string | null> {
  for (const directory of ACL_TOOL_DIRECTORIES) {
    const path = `${directory}/${name}`;
    if (
      await access(path, constants.X_OK).then(
        () => true,
        () => false,
      )
    )
      return path;
  }
  return null;
}

export const systemAcl: AclTools = {
  async read(path) {
    const getfacl = process.platform === "linux" ? await tool("getfacl") : null;
    if (!getfacl) return null;
    const output = await run(getfacl, ["--omit-header", "--absolute-names", "--", path]);
    return output === null ? null : parseGetfacl(output);
  },
  async apply(path, entries) {
    const wanted = named(entries);
    if (wanted.length === 0) return true;
    const setfacl = process.platform === "linux" ? await tool("setfacl") : null;
    if (!setfacl) return false;
    // One -m per entry, so a name holding a comma cannot split an entry in two.
    const args = wanted.flatMap((entry) => ["-m", entry]);
    return (await run(setfacl, [...args, "--", path])) !== null;
  },
};

// The permission bits the setup may set on an existing key or secrets directory. Access for
// everyone else is always removed: no ACL needs it, and clearing it leaves an ACL's mask alone.
// Group bits are removed too, except where they may be an ACL's mask (an ACL was read and has
// entries, or on Linux, where the ACL could not be read). On other systems they are plain group
// permissions. `change` is null when nothing needs changing.
export interface ModeDecision {
  change: number | null;
  kept: "acl" | "group" | null;
}

export function modeDecision(
  mode: number | null,
  acl: AccessList | null,
  platform: NodeJS.Platform,
): ModeDecision {
  if (mode === null || platform === "win32") return { change: null, kept: null };
  let target = mode & ~0o007;
  let kept: ModeDecision["kept"] = null;
  if ((mode & 0o070) !== 0) {
    if (acl && acl.entries.length > 0) kept = "acl";
    else if (!acl && platform === "linux") kept = "group";
    else target &= ~0o070;
  }
  return { change: target === mode ? null : target, kept };
}

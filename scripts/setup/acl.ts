import { execFile } from "node:child_process";

// POSIX access control lists on the secrets. On Linux the Docker install's container reads the
// key as user ID 1000, and an administrator whose account has another ID gives it access with
// setfacl (docs/SETUP.md). The setup must not undo that: a chmod rewrites the ACL's mask, and a
// rotation's new files would not carry the entries. Node has no ACL API, so this asks getfacl and
// setfacl, with fixed argument lists and no shell. Where they are missing, or on other systems,
// the ACL is unknown and callers fall back to leaving group permission bits alone.

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

// The command an administrator can run by hand when the setup could not re-apply an ACL.
export function setfaclCommand(path: string, entries: readonly string[]): string {
  const named = entries.filter((entry) => !entry.startsWith("mask::"));
  return `setfacl -m ${named.join(",")} '${path.replaceAll("'", "'\\''")}'`;
}

function run(command: string, args: string[]): Promise<string | null> {
  return new Promise((resolve) => {
    execFile(command, args, { timeout: 10_000, maxBuffer: 1_048_576 }, (error, stdout) =>
      resolve(error ? null : stdout),
    );
  });
}

export const systemAcl: AclTools = {
  async read(path) {
    if (process.platform !== "linux") return null;
    const output = await run("getfacl", ["--omit-header", "--absolute-names", "--", path]);
    return output === null ? null : parseGetfacl(output);
  },
  async apply(path, entries) {
    const named = entries.filter((entry) => !entry.startsWith("mask::"));
    if (process.platform !== "linux" || named.length === 0) return named.length === 0;
    return (await run("setfacl", ["-m", named.join(","), "--", path])) !== null;
  },
};

// What the setup may do to an existing file or directory's mode. It tightens only what is
// plainly too open, and never what an ACL or a deliberate group setting may depend on.
export type ModeDecision = "fine" | "tighten" | "keep-acl" | "keep-group";

export function modeDecision(mode: number | null, acl: AccessList | null): ModeDecision {
  if (mode === null || (mode & 0o077) === 0) return "fine";
  if (acl && acl.entries.length > 0) return "keep-acl";
  // Without an ACL reading, group bits may be an ACL's mask; only access for everyone else is
  // certainly wrong for a key.
  if (!acl && (mode & 0o007) === 0) return "keep-group";
  return "tighten";
}

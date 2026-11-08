import { useEffect, useRef, useState } from "react";
import { AccountName } from "./common.tsx";
import { lookUpPerson } from "./People.tsx";

// The Mailbox storage report names a mailbox by sign-in name and carries no account id, so a name
// opens its person by looking the name up in the directory. A sign-in name that is a hash is
// Microsoft concealing it (the same test the audits use) and is never looked up.
const HASHED = /^[0-9a-f]{32}$/i;

export const auditsLookedUpByName: Readonly<Record<string, string>> = {
  "mail.mailbox-storage": "mailbox",
};

export function canLookUp(cell: unknown): cell is string {
  return typeof cell === "string" && cell.trim() !== "" && !HASHED.test(cell.trim());
}

// The lookup most recently started anywhere on the page. A slower, older one never navigates.
let latestLookup = 0;

// Opens the account behind a sign-in name. Returns the plain message to show when it cannot: no
// account has that name, or the directory could not be read. A slow answer must not replace a
// page the administrator has since moved on from, so it opens the person only while `stillWanted`
// holds, and only if it is the latest lookup started.
export async function openByName(
  name: string,
  open: (userId: string) => void,
  lookup: (name: string, signal?: AbortSignal) => Promise<string> = lookUpPerson,
  options: { signal?: AbortSignal; stillWanted?: () => boolean } = {},
): Promise<string | null> {
  if (!canLookUp(name)) return "That name cannot be looked up.";
  latestLookup += 1;
  const mine = latestLookup;
  try {
    const userId = await lookup(name, options.signal);
    if (mine === latestLookup && !options.signal?.aborted && (options.stillWanted?.() ?? true))
      open(userId);
    return null;
  } catch (problem) {
    if (options.signal?.aborted) return null;
    return problem instanceof Error ? problem.message : "The account could not be looked up.";
  }
}

export function MailboxName({
  name,
  onOpenPerson,
}: {
  name: string;
  onOpenPerson: (userId: string) => void;
}) {
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState<string | null>(null);
  // Leaving the page cancels a lookup still in flight.
  const controller = useRef<AbortController | null>(null);
  useEffect(() => () => controller.current?.abort(), []);
  return (
    <div className="account-cell">
      <AccountName
        name={name}
        principal={name}
        onOpen={() => {
          if (busy) return;
          setBusy(true);
          setMessage(null);
          // The address at the click: if it has changed when the answer comes, the administrator
          // has gone elsewhere and the answer opens nothing.
          const hash = window.location.hash;
          const mine = new AbortController();
          controller.current = mine;
          void openByName(name, onOpenPerson, lookUpPerson, {
            signal: mine.signal,
            stillWanted: () => window.location.hash === hash,
          })
            .then((result) => {
              if (!mine.signal.aborted) setMessage(result);
            })
            .finally(() => {
              if (!mine.signal.aborted) setBusy(false);
            });
        }}
      />
      {message ? (
        <span className="audit-lookup-note" role="status">
          {message}
        </span>
      ) : null}
    </div>
  );
}

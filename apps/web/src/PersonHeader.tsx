import type { Severity } from "@prism365/contracts";
import type { ReactNode } from "react";
import { ActionsMenu, type MenuItem } from "./ActionsMenu.tsx";
import { canaryExplanation, type PersonCanary } from "./Canary.tsx";
import { CopyButton, type CopyStatus, useCopy } from "./CopyButton.tsx";
import { SeverityDot } from "./common.tsx";

// The top of a person's page: who they are and the state of the account, with every action in one
// menu. Nothing here reads tenant data; the page hands it what it has already loaded.

// Have I Been Pwned's documented deep link: the search runs when the page opens, in the
// administrator's own browser. No API, no key.
export function hibpLink(address: string): string {
  return `https://haveibeenpwned.com/account/${encodeURIComponent(address.trim()).replace(/%40/g, "@")}`;
}

// The person's page in the Entra admin center. The pattern is community-documented, not a
// published Microsoft specification (like the admin center links in the service health module),
// and it opens only for an administrator signed in to this tenant.
export function entraUserLink(userId: string): string {
  return `https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/overview/userId/${encodeURIComponent(userId)}`;
}

// The person's own address, for the breach search and for copying: their mailbox, else their
// sign-in name. Guest sign-in names (#EXT#) are not mail addresses and are skipped.
export function personAddress(
  mail: string | null,
  userPrincipalName: string | null,
): string | null {
  return (
    [mail, userPrincipalName].find(
      (value): value is string => Boolean(value?.includes("@")) && !value?.includes("#EXT#"),
    ) ?? null
  );
}

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// The letters in the avatar: the first letter of the first and last words of the name, or of an
// address's local part. An id, which says nothing about a person, gets a question mark.
export function initials(name: string): string {
  const text = name.trim();
  if (!text || GUID.test(text)) return "?";
  const local = text.includes("@") ? text.slice(0, text.indexOf("@")) : text;
  const letters = local
    .split(/[\s._#-]+/u)
    .map((word) => Array.from(word).find((character) => /\p{L}/u.test(character)))
    .filter((character): character is string => character !== undefined);
  const first = letters[0];
  if (first === undefined) return "?";
  const last = letters.length > 1 ? letters[letters.length - 1] : "";
  return `${first}${last}`.toLocaleUpperCase();
}

export interface PersonAccount {
  enabled: boolean | null;
  userType: string | null;
  synced: boolean | null;
}

export function userTypeLabel(userType: string): string {
  const lower = userType.toLowerCase();
  return lower === "guest" ? "Guest" : lower === "member" ? "Member" : userType;
}

const severityOrder: readonly Severity[] = ["critical", "high", "medium", "low", "info"];

// How many open findings sit at each severity, most serious first, leaving out the empty ones.
export function severityCounts(
  severities: readonly Severity[],
): { severity: Severity; count: number }[] {
  return severityOrder
    .map((severity) => ({
      severity,
      count: severities.filter((found) => found === severity).length,
    }))
    .filter((entry) => entry.count > 0);
}

export function PersonStatusChips({
  account,
  openSeverities,
  canary,
}: {
  // null until the directory has answered.
  account: PersonAccount | null;
  // The severity of each open finding; null until the person's record has loaded.
  openSeverities: readonly Severity[] | null;
  canary: boolean;
}) {
  return (
    <div className="person-tags">
      {account ? (
        <>
          {account.enabled === false ? (
            <span
              className="person-tag person-tag--alert"
              title="Sign-in is blocked for this account"
            >
              Disabled
            </span>
          ) : account.enabled ? (
            <span className="person-tag person-tag--ok">Enabled</span>
          ) : null}
          {account.userType ? (
            <span className="person-tag">{userTypeLabel(account.userType)}</span>
          ) : null}
          <span className="person-tag">
            {account.synced ? "Synced from on-premises" : "Cloud account"}
          </span>
        </>
      ) : null}
      {canary ? (
        <span className="person-tag person-tag--notice" title="Marked as a canary account">
          Canary
        </span>
      ) : null}
      {openSeverities === null ? null : openSeverities.length === 0 ? (
        <span className="person-tag person-tag--ok">Nothing open</span>
      ) : (
        <span className="person-findings">
          <span className="person-findings__label">Open findings</span>
          {severityCounts(openSeverities).map(({ severity, count }) => (
            <SeverityDot key={severity} severity={severity} count={count} />
          ))}
        </span>
      )}
    </div>
  );
}

// The menu's items, in the order an administrator reaches for them. A missing piece takes its
// item away rather than leaving one that cannot work: no address, no breach search or copy; no
// canary list, no canary item.
export function personMenuItems(input: {
  userId: string;
  address: string | null;
  canary: Pick<PersonCanary, "available" | "active" | "busy" | "onMark" | "onUnmark">;
  copy: CopyStatus;
  onCopy: () => void;
}): MenuItem[] {
  const { userId, address, canary, copy, onCopy } = input;
  const items: MenuItem[] = [];
  if (canary.available)
    items.push({
      kind: "button",
      id: "canary",
      label: canary.active ? "Stop treating as canary" : "Mark as canary",
      hint: canaryExplanation,
      disabled: canary.busy,
      onSelect: canary.active ? canary.onUnmark : canary.onMark,
    });
  if (address) {
    items.push({
      kind: "link",
      id: "breach",
      label: "Check Have I Been Pwned",
      hint: `Opens haveibeenpwned.com in a new tab with ${address} already searched.`,
      href: hibpLink(address),
    });
    items.push({
      kind: "button",
      id: "copy",
      label: copy === "copied" ? "Email copied" : copy === "failed" ? "Copy failed" : "Copy email",
      hint:
        copy === "failed"
          ? "The clipboard refused. Select the address and copy it by hand."
          : address,
      keepOpen: true,
      onSelect: onCopy,
    });
  }
  items.push({
    kind: "link",
    id: "entra",
    label: "Open in Microsoft Entra",
    hint: "Opens the Microsoft Entra admin center in a new tab.",
    href: entraUserLink(userId),
  });
  return items;
}

export function PersonActions({
  userId,
  address,
  canary,
}: {
  userId: string;
  address: string | null;
  canary: PersonCanary;
}) {
  const { status, copy } = useCopy(1_500);
  return (
    <>
      <ActionsMenu
        items={personMenuItems({
          userId,
          address,
          canary,
          copy: status,
          onCopy: () => {
            if (address) void copy(address);
          },
        })}
      />
      <span className="visually-hidden" aria-live="polite">
        {status === "copied" ? "Email copied" : status === "failed" ? "Copy failed" : ""}
      </span>
    </>
  );
}

export function PersonHeader({
  titleId,
  name,
  principal,
  jobTitle,
  department,
  account,
  openSeverities,
  canary,
  actions,
  below,
}: {
  titleId: string;
  name: string;
  principal: string | null;
  jobTitle: string | null;
  department: string | null;
  account: PersonAccount | null;
  openSeverities: readonly Severity[] | null;
  canary: boolean;
  // The Actions menu.
  actions: ReactNode;
  // What a choice in the menu asks next, such as the canary confirmation.
  below?: ReactNode;
}) {
  // When the name already is the address, the copy button sits beside the name instead of
  // repeating the address under it.
  const nameIsAddress = principal !== null && principal === name;
  return (
    <div className="person-header">
      <div className="person-header__main">
        <span className="person-avatar" aria-hidden="true">
          {initials(name)}
        </span>
        <div className="person-header__text">
          <div className="person-header__title">
            <h2 id={titleId}>{name}</h2>
            {nameIsAddress && principal ? <CopyButton text={principal} /> : null}
          </div>
          {principal && !nameIsAddress ? (
            <span className="copy-line person-header__principal">
              <span className="copy-line__text">{principal}</span>
              <CopyButton text={principal} />
            </span>
          ) : null}
          {jobTitle || department ? (
            <p className="person-header__role">
              {jobTitle ? <span>{jobTitle}</span> : null}
              {department ? <span>{department}</span> : null}
            </p>
          ) : null}
          {!principal && !jobTitle && !department ? (
            <p className="person-header__role">No directory details</p>
          ) : null}
          <PersonStatusChips account={account} openSeverities={openSeverities} canary={canary} />
        </div>
      </div>
      <div className="person-header__actions">{actions}</div>
      {below ? <div className="person-header__below">{below}</div> : null}
    </div>
  );
}

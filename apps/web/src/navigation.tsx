import { createContext, Fragment, type ReactNode, useContext } from "react";
import type { WorkspaceView } from "./App.tsx";
import { AccountName } from "./common.tsx";
import type { PersonTab } from "./PersonTabs.tsx";

// One way to open a person from anywhere, so a deep component needs no prop chain to do it. The
// App provides it once. A page that opens people from a known place wraps its content in an
// OriginScope, and the person's page then says where you came from. The origin lives in memory
// only: a reload or a pasted link has none, and the address bar holds no trace of it.

export type PersonOrigin =
  | { kind: "view"; view: WorkspaceView; label: string }
  | { kind: "audit"; auditId: string; title: string }
  | { kind: "person"; userId: string; name: string };

export interface Navigation {
  // `from` undefined means "from the page you are on"; null means from nowhere in particular
  // (quick search, which can be opened over any page).
  // `tab` opens the page on that tab; without it the page opens on its Overview.
  openPerson(userId: string, from?: PersonOrigin | null, tab?: PersonTab): void;
  openAudit(auditId: string): void;
  openView(view: WorkspaceView): void;
  // Where the open person's page was reached from, if anywhere.
  origin: PersonOrigin | null;
}

const idle: Navigation = {
  openPerson: () => {},
  openAudit: () => {},
  openView: () => {},
  origin: null,
};

export const NavigationContext = createContext<Navigation>(idle);
const OriginContext = createContext<PersonOrigin | null | undefined>(undefined);

export const useNavigation = (): Navigation => useContext(NavigationContext);

// What a person link opened inside this scope reports as its origin.
export function OriginScope({
  origin,
  children,
}: {
  origin: PersonOrigin | null;
  children: ReactNode;
}) {
  return <OriginContext.Provider value={origin}>{children}</OriginContext.Provider>;
}

export interface Crumb {
  label: string;
  open: () => void;
}

// The parts of a person page's breadcrumb that come before the person's own name. Empty when
// there is no origin, and the page then shows its usual People part.
export function originTrail(
  origin: PersonOrigin | null,
  nav: Pick<Navigation, "openPerson" | "openAudit" | "openView">,
): Crumb[] {
  if (!origin) return [];
  if (origin.kind === "view")
    return [{ label: origin.label, open: () => nav.openView(origin.view) }];
  if (origin.kind === "audit")
    return [
      { label: "Audits", open: () => nav.openView("audits") },
      { label: origin.title, open: () => nav.openAudit(origin.auditId) },
    ];
  return [{ label: origin.name, open: () => nav.openPerson(origin.userId, null) }];
}

// A person's name as a link when the account has an id, and as plain text when it has none (a
// group, contact, device or unresolved trustee). Only user objects belong here.
export function PersonLink({
  userId,
  name,
  principal = null,
  from,
  tab,
}: {
  userId: string | null;
  name: string;
  principal?: string | null;
  from?: PersonOrigin | null;
  tab?: PersonTab;
}) {
  const nav = useNavigation();
  const scope = useContext(OriginContext);
  return (
    <AccountName
      name={name}
      principal={principal}
      onOpen={personOpener(nav, userId, from !== undefined ? from : scope, tab)}
    />
  );
}

// What a person link does when clicked, or null when the name has no account behind it.
export function personOpener(
  nav: Pick<Navigation, "openPerson">,
  userId: string | null,
  origin: PersonOrigin | null | undefined,
  tab?: PersonTab,
): (() => void) | null {
  return userId ? () => nav.openPerson(userId, origin, tab) : null;
}

// The person page's breadcrumb: the places you came from, each returning there, then the
// person's name. Without an origin it is People, which returns to the list.
export function PersonBreadcrumb({
  trail,
  name,
  onBack,
}: {
  trail: readonly Crumb[];
  name: string;
  onBack: () => void;
}) {
  const parts = trail.length > 0 ? trail : [{ label: "People", open: onBack }];
  return (
    <nav className="breadcrumb" aria-label="Breadcrumb">
      {parts.map((crumb) => (
        <Fragment key={crumb.label}>
          <button type="button" onClick={crumb.open}>
            {crumb.label}
          </button>
          <span aria-hidden="true">/</span>
        </Fragment>
      ))}
      <span aria-current="page">{name}</span>
    </nav>
  );
}

// Opens a person from inside the current scope, for code that wants a callback rather than a
// link (a list row that already draws its own button).
export function useOpenPerson(): (userId: string) => void {
  const nav = useNavigation();
  const scope = useContext(OriginContext);
  return (userId) => nav.openPerson(userId, scope);
}

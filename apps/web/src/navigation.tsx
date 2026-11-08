import type { EntityKind } from "@prism365/contracts";
import { createContext, Fragment, type ReactNode, useContext } from "react";
import type { WorkspaceView } from "./App.tsx";
import { AccountName } from "./common.tsx";
import { CopyButton } from "./CopyButton.tsx";
import type { PersonTab } from "./PersonTabs.tsx";

// One way to open a person, or a group, role, policy, device or app, from anywhere, so a deep
// component needs no prop chain to do it. The App provides it once. A page that opens people from
// a known place wraps its content in an OriginScope, and the person's page then says where you came
// from. The origin lives in memory only: a reload or a pasted link has none, and the address bar
// holds no trace of it.

export type PersonOrigin =
  | { kind: "view"; view: WorkspaceView; label: string }
  | { kind: "audit"; auditId: string; title: string }
  | { kind: "person"; userId: string; name: string }
  // A person opened from an entity's panel, which closes when the person's page opens.
  | { kind: "entity"; entityKind: EntityKind; id: string; name: string };

// What an entity is called in a sentence. Lower case; `capital` starts a label with it.
export const capital = (word: string) => word.charAt(0).toUpperCase() + word.slice(1);

export const entityKindLabels: Record<EntityKind, string> = {
  group: "group",
  role: "role",
  policy: "policy",
  device: "device",
  app: "app",
};

export interface Navigation {
  // `from` undefined means "from the page you are on"; null means from nowhere in particular
  // (quick search, which can be opened over any page).
  // `tab` opens the page on that tab; without it the page opens on its Overview.
  openPerson(userId: string, from?: PersonOrigin | null, tab?: PersonTab): void;
  openAudit(auditId: string): void;
  openView(view: WorkspaceView): void;
  // Opens the side panel for a group, role, policy, device or app. Opened from inside the panel it
  // adds a step the panel can go back through; opened from a page it starts the panel afresh.
  openEntity(kind: EntityKind, id: string): void;
  // Where the open person's page was reached from, if anywhere.
  origin: PersonOrigin | null;
}

const idle: Navigation = {
  openPerson: () => {},
  openAudit: () => {},
  openView: () => {},
  openEntity: () => {},
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
  nav: Pick<Navigation, "openPerson" | "openAudit" | "openView" | "openEntity">,
): Crumb[] {
  if (!origin) return [];
  if (origin.kind === "view")
    return [{ label: origin.label, open: () => nav.openView(origin.view) }];
  if (origin.kind === "audit")
    return [
      { label: "Audits", open: () => nav.openView("audits") },
      { label: origin.title, open: () => nav.openAudit(origin.auditId) },
    ];
  if (origin.kind === "entity")
    return [
      {
        label: `${capital(entityKindLabels[origin.entityKind])}: ${origin.name}`,
        open: () => nav.openEntity(origin.entityKind, origin.id),
      },
    ];
  return [{ label: origin.name, open: () => nav.openPerson(origin.userId, null) }];
}

// A group, role, policy, device or app as a link that opens its panel. A thing no stored audit
// names reads "Unknown group" and keeps its id behind a copy button, so the id can be searched
// for elsewhere without taking up the line. With no id there is nothing to open, and the name is
// plain text.
export function EntityLink({
  kind,
  id,
  name,
}: {
  kind: EntityKind;
  id: string | null;
  name: string | null;
}) {
  const nav = useNavigation();
  const label = name?.trim() ? name : `Unknown ${entityKindLabels[kind]}`;
  if (!id) return <span className={name?.trim() ? undefined : "person-muted"}>{label}</span>;
  const link = (
    <button
      className="people-name entity-link"
      type="button"
      onClick={() => nav.openEntity(kind, id)}
    >
      <strong className="people-name__name">{label}</strong>
    </button>
  );
  if (name?.trim()) return link;
  return (
    <span className="copy-line">
      {link}
      <CopyButton text={id} label={`Copy the ID of this ${entityKindLabels[kind]}`} />
    </span>
  );
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

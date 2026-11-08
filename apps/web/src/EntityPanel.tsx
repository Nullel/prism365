import {
  type AuditCell,
  type Entity,
  type EntityList,
  type EntityMention,
  type EntityRow,
  type EntitySection,
  entitySchema,
} from "@prism365/contracts";
import { type ReactNode, type RefObject, useEffect, useId, useMemo, useRef, useState } from "react";
import { day, number, RelativeTime, readJson, SeverityDot, SkeletonRows, when } from "./common.tsx";
import { CopyButton } from "./CopyButton.tsx";
import { type EntityTarget, handlePanelKey, type PanelAction } from "./entityPanelState.ts";
import { CloseIcon } from "./icons.tsx";
import {
  capital,
  EntityLink,
  entityKindLabels,
  OriginScope,
  PersonLink,
  type PersonOrigin,
} from "./navigation.tsx";
import { focusKeeper } from "./QuickSearch.tsx";

// The side panel that opens when a group, role, Conditional Access policy, managed device or
// consented app is clicked: everything the stored audits know about it, one section per audit, each
// saying which audit it is from and how old that is. The people and things it names are links, so
// an administrator can follow a group to its owner to the groups they own without going back to a
// table. Read from stored results only; nothing here asks Microsoft anything.

const entityUrl = (target: EntityTarget) =>
  `/api/v1/entities/${encodeURIComponent(target.kind)}/${encodeURIComponent(target.id)}`;

// A person, or a thing of one of the five kinds, as a link. A name nobody stored reads "Unknown ..."
// and keeps its id behind a copy button.
function Mention({
  mention,
  unknown = "holder",
}: {
  mention: EntityMention;
  // What an object nobody named is called: an owner, a holder.
  unknown?: string;
}) {
  if (mention.kind === "person") {
    if (mention.name || mention.principal)
      return (
        <PersonLink
          userId={mention.id}
          name={mention.name ?? mention.principal ?? "Unknown account"}
          principal={mention.principal}
        />
      );
    return (
      <span className="copy-line">
        <PersonLink userId={mention.id} name="Unknown account" />
        {mention.id ? <CopyButton text={mention.id} label="Copy the ID of this account" /> : null}
      </span>
    );
  }
  if (mention.kind === "other")
    return (
      <span className="copy-line">
        <span className={mention.name ? undefined : "person-muted"}>
          {mention.name ?? `Unknown ${unknown}`}
        </span>
        {mention.id ? (
          <CopyButton text={mention.id} label={`Copy the ID of this ${unknown}`} />
        ) : null}
      </span>
    );
  return <EntityLink kind={mention.kind} id={mention.id} name={mention.name} />;
}

// A cell's value as the panel reads it out. The column's kind says what a number or a date is.
function cellText(kind: string, value: AuditCell): ReactNode {
  if (value === null) return "None";
  if (typeof value === "boolean") return value ? "Yes" : "No";
  if (typeof value === "number")
    return kind === "days"
      ? value === 0
        ? "Today"
        : `${number.format(value)} ${value === 1 ? "day" : "days"}`
      : number.format(value);
  if (Array.isArray(value)) {
    const shown = value.slice(0, 6);
    return (
      <>
        {shown.join(", ")}
        {value.length > shown.length ? (
          <span className="person-muted">
            {" "}
            and {number.format(value.length - shown.length)} more
          </span>
        ) : null}
      </>
    );
  }
  if (typeof value === "object") return value.name ?? value.principal ?? "Unnamed account";
  if (kind === "date")
    return (
      <time dateTime={value} title={when(value)}>
        {day(value)}
      </time>
    );
  return value;
}

function Facts({ cells }: { cells: EntityRow["cells"] }) {
  if (cells.length === 0) return null;
  return (
    <dl className="entity-facts">
      {cells.map((cell) => (
        <div key={cell.key}>
          <dt>{cell.label}</dt>
          <dd>{cellText(cell.kind, cell.value)}</dd>
        </div>
      ))}
    </dl>
  );
}

function Lists({ lists }: { lists: readonly EntityList[] }) {
  return (
    <>
      {lists.map((list) => (
        <div className="entity-list" key={list.label}>
          <h4>{list.label}</h4>
          <ul className="person-access__list">
            {list.items.map((item) => (
              <li key={`${item.kind}:${item.id ?? item.name ?? "unnamed"}`}>
                <Mention mention={item} unknown={list.label === "Owners" ? "owner" : "holder"} />
              </li>
            ))}
          </ul>
          {list.more > 0 ? (
            <p className="signin-coverage">and {number.format(list.more)} more</p>
          ) : null}
        </div>
      ))}
    </>
  );
}

function Row({ row }: { row: EntityRow }) {
  return (
    <li className="entity-row">
      {row.subject ? (
        <div className="entity-row__subject">
          <Mention mention={row.subject} />
          {row.tags.map((tag) => (
            <span className="person-tag" key={tag}>
              {tag}
            </span>
          ))}
        </div>
      ) : null}
      {row.risks.length > 0 ? (
        <ul className="entity-risks">
          {row.risks.map((rule) => (
            <li key={rule.id}>
              <SeverityDot severity={rule.severity} /> {rule.label}
            </li>
          ))}
        </ul>
      ) : null}
      <Facts cells={row.cells} />
      <Lists lists={row.lists} />
    </li>
  );
}

const auditHref = (section: EntitySection) => `#/audits/${encodeURIComponent(section.auditId)}`;

function Source({ section }: { section: EntitySection }) {
  const audit = <a href={auditHref(section)}>{section.auditName}</a>;
  if (section.status === "no_run" || section.status === "failed")
    return <p className="signin-coverage entity-source">From the {audit} audit</p>;
  return (
    <p className="signin-coverage entity-source">
      From the {audit} audit, <RelativeTime value={section.finishedAt} />
      {section.status === "outdated" || section.complete
        ? ""
        : ". That run could not check everything"}
    </p>
  );
}

function SectionBody({ section, kind }: { section: EntitySection; kind: Entity["kind"] }) {
  switch (section.status) {
    case "no_run":
      return (
        <p className="empty-data">
          <a href={auditHref(section)}>Run the {section.auditName} audit</a> to see this here.
        </p>
      );
    case "outdated":
      return (
        <p className="empty-data">
          That run was made before this audit kept what is needed here, so it cannot say.{" "}
          <a href={auditHref(section)}>Run the {section.auditName} audit again</a> to see this here.
        </p>
      );
    case "failed":
      return <p className="empty-data">{section.message}</p>;
    case "not_in_run":
      return (
        <p className="empty-data">
          This {entityKindLabels[kind]} is not in the latest {section.auditName} run.
        </p>
      );
    case "ok":
      return (
        <>
          {section.note ? <p className="empty-data">{section.note}</p> : null}
          {section.rows.length > 0 ? (
            <ul className="entity-rows">
              {section.rows.map((row) => (
                <Row key={row.key} row={row} />
              ))}
            </ul>
          ) : null}
          {section.more ? (
            <p className="signin-coverage">
              Showing the first {number.format(section.rows.length)}. The audit holds more.
            </p>
          ) : null}
        </>
      );
  }
}

export function EntitySectionView({
  section,
  kind,
}: {
  section: EntitySection;
  kind: Entity["kind"];
}) {
  const headingId = useId();
  return (
    <section className="panel person-access entity-section" aria-labelledby={headingId}>
      <h3 id={headingId}>{section.title}</h3>
      <Source section={section} />
      <SectionBody section={section} kind={kind} />
    </section>
  );
}

export interface EntityPanelViewProps {
  target: EntityTarget;
  // null while it loads.
  entity: Entity | null;
  error: string;
  // Whether the panel holds an earlier thing to go back to.
  canGoBack: boolean;
  titleId: string;
  onBack(): void;
  onClose(): void;
  dialogRef?: RefObject<HTMLDivElement | null>;
  // Where focus lands when the panel opens or shows another thing: the heading, which a reader
  // announces by name and which shows no focus ring.
  headingRef?: RefObject<HTMLHeadingElement | null>;
  onBackdrop?: () => void;
}

// The origin a person opened from this panel reports: this thing, by the name it has.
export function panelOrigin(target: EntityTarget, entity: Entity | null): PersonOrigin {
  return {
    kind: "entity",
    entityKind: target.kind,
    id: target.id,
    name: entity?.name ?? `Unknown ${entityKindLabels[target.kind]}`,
  };
}

// The panel as a function of what has loaded. It holds no state, so a test can draw it.
export function EntityPanelView({
  target,
  entity,
  error,
  canGoBack,
  titleId,
  onBack,
  onClose,
  dialogRef,
  headingRef,
  onBackdrop,
}: EntityPanelViewProps) {
  const kindLabel = capital(entityKindLabels[target.kind]);
  const name = entity ? (entity.name ?? `Unknown ${entityKindLabels[target.kind]}`) : kindLabel;
  return (
    // biome-ignore lint/a11y/noStaticElementInteractions: closing on the backdrop is a pointer convenience; Escape and the close button do the same from the keyboard.
    <div
      className="entity-panel__backdrop"
      role="presentation"
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) onBackdrop?.();
      }}
    >
      <div
        className="entity-panel"
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        ref={dialogRef}
      >
        <div className="entity-panel__bar">
          {canGoBack ? (
            <button className="button button--secondary" type="button" onClick={onBack}>
              Back
            </button>
          ) : (
            <span />
          )}
          <button
            className="entity-panel__close"
            type="button"
            aria-label="Close"
            onClick={onClose}
          >
            <CloseIcon />
          </button>
        </div>
        <header className="entity-panel__header">
          <span className="eyebrow">{kindLabel}</span>
          <h2 id={titleId} tabIndex={-1} ref={headingRef}>
            {name}
            {entity && entity.name === null ? (
              <CopyButton
                text={target.id}
                label={`Copy the ID of this ${entityKindLabels[target.kind]}`}
              />
            ) : null}
          </h2>
          {entity && entity.tags.length > 0 ? (
            <div className="person-tags">
              {entity.tags.map((tag) => (
                <span className="person-tag" key={tag}>
                  {tag}
                </span>
              ))}
            </div>
          ) : null}
        </header>
        <div className="entity-panel__body">
          {error ? (
            <p className="empty-data" role="alert">
              {error}
            </p>
          ) : !entity ? (
            <SkeletonRows
              rows={5}
              label={`Reading what is stored about this ${entityKindLabels[target.kind]}`}
            />
          ) : (
            <OriginScope origin={panelOrigin(target, entity)}>
              {entity.notes.map((note) => (
                <p className="signin-coverage" key={note}>
                  {note}
                </p>
              ))}
              {entity.sections.map((section) => (
                <EntitySectionView key={section.id} section={section} kind={entity.kind} />
              ))}
            </OriginScope>
          )}
        </div>
      </div>
    </div>
  );
}

const FOCUSABLE = 'a[href], button:not([disabled]), input:not([disabled]), [tabindex="0"]';

// Loads the thing at the top of the stack and owns the panel's focus and keys: focus moves in when
// the panel opens and each time it shows something else, and goes back to what had it when the
// panel closes.
export function EntityPanel({
  stack,
  dispatch,
}: {
  stack: readonly EntityTarget[];
  dispatch: (action: PanelAction) => void;
}) {
  const titleId = useId();
  const dialog = useRef<HTMLDivElement>(null);
  const heading = useRef<HTMLHeadingElement>(null);
  const keeper = useMemo(
    () => focusKeeper(() => (typeof document === "undefined" ? null : document.activeElement)),
    [],
  );
  const target = stack.at(-1);
  const key = target ? `${target.kind}/${target.id}` : "";
  const [loaded, setLoaded] = useState<{
    key: string;
    entity: Entity | null;
    error: string;
  } | null>(null);

  // Taken once, when the panel opens, before focus moves in.
  useEffect(() => {
    keeper.remember();
    return () => keeper.restore();
  }, [keeper]);

  const kind = target?.kind;
  const targetId = target?.id;
  useEffect(() => {
    if (!kind || !targetId) return;
    const controller = new AbortController();
    readJson(entityUrl({ kind, id: targetId }), (value) => entitySchema.parse(value), {
      signal: controller.signal,
    }).then(
      (entity) => setLoaded({ key: `${kind}/${targetId}`, entity, error: "" }),
      (loadError: unknown) => {
        if (controller.signal.aborted) return;
        setLoaded({
          key: `${kind}/${targetId}`,
          entity: null,
          error:
            loadError instanceof Error && loadError.message
              ? loadError.message
              : "This could not be loaded. Close the panel and try again.",
        });
      },
    );
    return () => controller.abort();
  }, [kind, targetId]);

  // Focus lands on the heading, so a screen reader reads the name and no ring is drawn around the
  // panel, and again when the panel shows another thing.
  useEffect(() => {
    if (key) heading.current?.focus({ preventScroll: true });
  }, [key]);

  // Keys are read on the page, not on the dialog, so Escape works and Tab stays in the panel
  // wherever focus has got to. Quick search, open above the panel, handles its own keys first.
  useEffect(() => {
    const listener = (event: globalThis.KeyboardEvent) => {
      const found = [...(dialog.current?.querySelectorAll<HTMLElement>(FOCUSABLE) ?? [])];
      const active = document.activeElement;
      handlePanelKey(
        event,
        {
          first: found[0] ?? null,
          last: found.at(-1) ?? null,
          active,
          inside: Boolean(active && dialog.current?.contains(active)),
        },
        () => dispatch({ type: "close" }),
      );
    };
    document.addEventListener("keydown", listener);
    return () => document.removeEventListener("keydown", listener);
  }, [dispatch]);

  if (!target) return null;
  const current = loaded && loaded.key === key ? loaded : null;
  return (
    <EntityPanelView
      target={target}
      entity={current?.entity ?? null}
      error={current?.error ?? ""}
      canGoBack={stack.length > 1}
      titleId={titleId}
      dialogRef={dialog}
      headingRef={heading}
      onBack={() => dispatch({ type: "back" })}
      onClose={() => dispatch({ type: "close" })}
      onBackdrop={() => dispatch({ type: "close" })}
    />
  );
}

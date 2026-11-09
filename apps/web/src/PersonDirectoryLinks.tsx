import type { PersonDirectory } from "@prism365/contracts";
import { PersonLink, type PersonOrigin } from "./navigation.tsx";

// The manager and the direct reports in a person's Directory list. Each name that is an account
// opens that person, so an administrator can walk up and down the reporting line.

type Manager = NonNullable<PersonDirectory["manager"]>;
type Reports = NonNullable<PersonDirectory["directReports"]>;

const REPORTS_SHOWN = 50;

const contacts = (count: number) => `${count} ${count === 1 ? "contact" : "contacts"}`;

export function ManagerValue({
  manager,
  from,
}: {
  manager: Manager | null;
  from?: PersonOrigin | null;
}) {
  if (manager?.data) {
    const { id, displayName, userPrincipalName, jobTitle, isUser } = manager.data;
    const name = displayName ?? userPrincipalName ?? "Unnamed";
    // An organizational contact has no account page, so it is text and says what it is.
    if (!isUser) return <>{`${name}${jobTitle ? `, ${jobTitle}` : ""} (contact)`}</>;
    return (
      <>
        <PersonLink userId={id} name={name} {...(from !== undefined ? { from } : {})} />
        {jobTitle ? `, ${jobTitle}` : ""}
      </>
    );
  }
  return <>{manager?.state === "available" ? "None set" : "Unavailable"}</>;
}

// The names flow inline and wrap, each with its job title muted after it, so a manager with many
// reports does not stretch the grid they sit under.
export function DirectReportsValue({
  reports,
  from,
}: {
  reports: Reports | null;
  from?: PersonOrigin | null;
}) {
  if (!reports || reports.state === "unavailable") return <>Unavailable</>;
  if (reports.state === "missingConsent") return <>Microsoft did not allow this read</>;
  if (reports.data.length === 0 && reports.skipped === 0) return <>None</>;
  const shown = reports.data.slice(0, REPORTS_SHOWN);
  const left = reports.data.length - shown.length;
  const tail = left > 0 ? `and ${left} more` : reports.more ? "and more" : null;
  return (
    <ul className="person-reports">
      {shown.map((report) => (
        <li key={report.id}>
          <PersonLink
            userId={report.id}
            name={report.displayName ?? report.userPrincipalName ?? "Unnamed"}
            {...(from !== undefined ? { from } : {})}
          />
          {report.jobTitle ? <span className="person-muted"> {report.jobTitle}</span> : null}
        </li>
      ))}
      {reports.skipped > 0 ? (
        <li className="person-muted">
          {shown.length === 0
            ? `${contacts(reports.skipped)}, no accounts`
            : `and ${contacts(reports.skipped)}`}
        </li>
      ) : null}
      {tail ? <li className="person-muted">{tail}</li> : null}
    </ul>
  );
}

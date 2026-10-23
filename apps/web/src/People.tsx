import {
  findingSchema,
  licenseName,
  type PersonDirectory,
  personDirectorySchema,
  type PersonList,
  personListSchema,
  personLookupSchema,
  type PersonRecord,
  personRecordSchema,
  type PersonSignin,
  personSigninPageSchema,
  type PrismSession,
} from "@prism365/contracts";
import { type FormEvent, useCallback, useEffect, useId, useRef, useState } from "react";
import { day, FindingItem, number, place, post, readJson, when } from "./common.tsx";
import { ShieldIcon } from "./icons.tsx";

// People: every account that signs in, and one page per account with everything Prism365 knows
// about it. Stored sign-ins come from the sign-in audit module; the directory, manager, last
// sign-in and licenses are read live from Microsoft each time the page opens.

const resultLabels: Record<PersonSignin["result"], string> = {
  success: "Success",
  failure: "Failed",
  mfaFailure: "MFA failed",
  blocked: "Blocked",
  interrupted: "Interrupted",
};

// Have I Been Pwned's documented deep link: the search runs when the page opens, in the
// administrator's own browser. No API, no key.
export function hibpLink(address: string): string {
  return `https://haveibeenpwned.com/account/${encodeURIComponent(address.trim()).replace(/%40/g, "@")}`;
}

const yesNo = (value: boolean | null) => (value === null ? "Unknown" : value ? "Yes" : "No");

function Intro({ title, children }: { title: string; children: string }) {
  const id = useId();
  return (
    <section className="people" aria-labelledby={id}>
      <div className="signin-audit__intro">
        <span className="eyebrow">People</span>
        <h2 id={id}>{title}</h2>
        <p>{children}</p>
      </div>
    </section>
  );
}

export function People({
  session,
  selected,
  onSelect,
}: {
  session: PrismSession;
  selected: string | null;
  onSelect: (userId: string | null) => void;
}) {
  if (session.mode === "demo")
    return (
      <Intro title="Everyone who signs in, one page each.">
        Connected to a tenant, this view lists every account that signs in and opens a page for
        each: where and how they sign in, their addresses and devices, their licenses and directory
        details, and anything flagged. Demo mode shows no people, because Prism365 does not invent
        security evidence.
      </Intro>
    );
  if (!session.authenticated)
    return (
      <Intro title="Sign in to see people.">
        Use the tenant connection above. People reads your stored sign-ins and the directory with
        your own delegated access.
      </Intro>
    );
  return selected ? (
    <PersonProfile
      key={selected}
      session={session}
      userId={selected}
      onBack={() => onSelect(null)}
    />
  ) : (
    <PeopleList onSelect={onSelect} />
  );
}

function PeopleList({ onSelect }: { onSelect: (userId: string) => void }) {
  const titleId = useId();
  const searchId = useId();
  const [search, setSearch] = useState("");
  const [query, setQuery] = useState("");
  const [list, setList] = useState<PersonList | null>(null);
  const [error, setError] = useState("");
  const [looking, setLooking] = useState(false);

  useEffect(() => {
    const timer = setTimeout(() => setQuery(search.trim()), 250);
    return () => clearTimeout(timer);
  }, [search]);

  useEffect(() => {
    const controller = new AbortController();
    const url = `/api/v1/people?days=30${query ? `&search=${encodeURIComponent(query)}` : ""}`;
    readJson(url, (value) => personListSchema.parse(value), { signal: controller.signal }).then(
      (value) => {
        setList(value);
        setError("");
      },
      (loadError: unknown) => {
        if (controller.signal.aborted) return;
        setError(loadError instanceof Error ? loadError.message : "People could not be loaded.");
      },
    );
    return () => controller.abort();
  }, [query]);

  // Enter opens the only match, or looks a full work email up in the directory.
  async function lookUp(event: FormEvent) {
    event.preventDefault();
    const value = search.trim();
    const only = list?.items.length === 1 ? list.items[0] : undefined;
    if (only && query === value) return onSelect(only.userId);
    if (!value.includes("@")) return;
    setLooking(true);
    setError("");
    try {
      const found = await readJson(
        `/api/v1/people/lookup?email=${encodeURIComponent(value)}`,
        (body) => personLookupSchema.parse(body),
      );
      onSelect(found.userId);
    } catch (lookupError) {
      setError(lookupError instanceof Error ? lookupError.message : "The lookup failed.");
    } finally {
      setLooking(false);
    }
  }

  const items = list?.items ?? [];
  return (
    <section className="people" aria-labelledby={titleId}>
      <h2 id={titleId} className="visually-hidden">
        People
      </h2>

      <section className="panel signin-table people-table">
        <div className="people-table__header">
          <h3>
            {query ? `Matching “${query}”` : "Accounts that signed in"}, last 30 days
            {list ? ` (${number.format(items.length)}${list.truncated ? "+" : ""})` : ""}
          </h3>
          <form className="people-search" onSubmit={(event) => void lookUp(event)}>
            <label className="visually-hidden" htmlFor={searchId}>
              Find a person
            </label>
            <input
              id={searchId}
              className="people-search__input"
              type="search"
              value={search}
              placeholder="Search by name or email"
              autoComplete="off"
              aria-describedby={`${searchId}-hint`}
              onChange={(event) => setSearch(event.target.value)}
            />
            <button className="button button--secondary" type="submit" disabled={looking}>
              {looking ? "Looking up…" : "Open"}
            </button>
          </form>
        </div>
        <p className="people-search__hint" id={`${searchId}-hint`}>
          A full work email opens anyone in the directory, even someone with no stored sign-ins.
        </p>
        {error ? (
          <p className="form-error" role="alert">
            {error}
          </p>
        ) : null}
        {!list ? (
          <p className="empty-data">Loading…</p>
        ) : items.length === 0 ? (
          <p className="empty-data">
            {query
              ? "No stored sign-ins match. Enter a full work email to look the person up in the directory."
              : "No sign-ins stored yet. Start collecting on Sign-in audit and people appear here."}
          </p>
        ) : (
          <table>
            <caption className="visually-hidden">
              Accounts that signed in during the last 30 days
            </caption>
            <thead>
              <tr>
                <th scope="col">Person</th>
                <th scope="col">Last sign-in</th>
                <th scope="col">Sign-ins</th>
                <th scope="col">Failed</th>
                <th scope="col">Countries</th>
                <th scope="col">Open findings</th>
              </tr>
            </thead>
            <tbody>
              {items.map((item) => (
                <tr key={item.userId}>
                  <th scope="row">
                    <button
                      className="people-name"
                      type="button"
                      onClick={() => onSelect(item.userId)}
                    >
                      <strong className="people-name__name">
                        {item.displayName ?? item.userPrincipalName ?? item.userId}
                      </strong>
                      {item.userPrincipalName ? <span>{item.userPrincipalName}</span> : null}
                    </button>
                  </th>
                  <td>{when(item.lastSignIn)}</td>
                  <td>{number.format(item.signIns)}</td>
                  <td>{number.format(item.failures)}</td>
                  <td>{number.format(item.countries)}</td>
                  <td>
                    {item.openFindings > 0 ? (
                      <span className="person-tag person-tag--alert">{item.openFindings}</span>
                    ) : (
                      "0"
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </section>
    </section>
  );
}

function Status({ usual, known }: { usual: boolean; known: boolean }) {
  if (usual) return <span className="person-tag person-tag--ok">Usual</span>;
  if (known) return <span className="person-tag person-tag--alert">Unfamiliar</span>;
  return <span className="person-muted">Not learned</span>;
}

function PersonProfile({
  session,
  userId,
  onBack,
}: {
  session: PrismSession;
  userId: string;
  onBack: () => void;
}) {
  const titleId = useId();
  const path = `/api/v1/people/${encodeURIComponent(userId)}`;
  const [record, setRecord] = useState<PersonRecord | null>(null);
  const [directory, setDirectory] = useState<PersonDirectory | null>(null);
  const [signins, setSignins] = useState<PersonSignin[] | null>(null);
  const [next, setNext] = useState<string | null>(null);
  const [filter, setFilter] = useState<"all" | "success" | "failed">("all");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState<string | null>(null);

  const loadRecord = useCallback(
    async (signal?: AbortSignal) => {
      const value = await readJson(path, (body) => personRecordSchema.parse(body), {
        ...(signal ? { signal } : {}),
      });
      setRecord(value);
    },
    [path],
  );

  useEffect(() => {
    const controller = new AbortController();
    loadRecord(controller.signal).catch((loadError: unknown) => {
      if (!controller.signal.aborted)
        setError(
          loadError instanceof Error ? loadError.message : "This person could not be loaded.",
        );
    });
    readJson(`${path}/directory`, (body) => personDirectorySchema.parse(body), {
      signal: controller.signal,
    }).then(setDirectory, () => {
      if (!controller.signal.aborted)
        setDirectory({
          state: "unavailable",
          message: "Microsoft's directory could not be read just now.",
          profile: null,
          manager: null,
          signInActivity: null,
          licenses: null,
        });
    });
    return () => controller.abort();
  }, [loadRecord, path]);

  // Each filter change starts a new generation; a page that arrives for an older one is dropped,
  // so "Not successful" can never gain rows loaded under "All".
  const generation = useRef(0);
  const loadSignins = useCallback(
    async (before: string | null, signal?: AbortSignal) => {
      const mine = generation.current;
      const query = new URLSearchParams({ result: filter, limit: "100" });
      if (before) query.set("before", before);
      const page = await readJson(
        `${path}/signins?${query}`,
        (body) => personSigninPageSchema.parse(body),
        {
          ...(signal ? { signal } : {}),
        },
      );
      if (mine !== generation.current) return;
      setSignins((current) => (before && current ? [...current, ...page.items] : page.items));
      setNext(page.next);
    },
    [filter, path],
  );

  useEffect(() => {
    const controller = new AbortController();
    generation.current += 1;
    setSignins(null);
    loadSignins(null, controller.signal).catch((loadError: unknown) => {
      if (!controller.signal.aborted)
        setError(loadError instanceof Error ? loadError.message : "Sign-ins could not be loaded.");
    });
    return () => controller.abort();
  }, [loadSignins]);

  async function act(label: string, operation: () => Promise<unknown>) {
    setBusy(label);
    setError("");
    try {
      await operation();
      await loadRecord();
    } catch (actionError) {
      setError(actionError instanceof Error ? actionError.message : "The request failed.");
    } finally {
      setBusy(null);
    }
  }

  const history = record?.history ?? null;
  const profile = directory?.profile ?? null;
  const name =
    profile?.displayName ??
    history?.displayName ??
    profile?.userPrincipalName ??
    history?.userPrincipalName ??
    userId;
  const upn = profile?.userPrincipalName ?? history?.userPrincipalName ?? null;
  // The person's own address for the breach search: their mailbox, else their sign-in name.
  // Guest sign-in names (#EXT#) are not mail addresses and are skipped.
  const breachAddress =
    [profile?.mail, upn].find(
      (value): value is string => Boolean(value?.includes("@")) && !value?.includes("#EXT#"),
    ) ?? null;
  const learned = new Map(
    (history?.baseline ?? []).map((entry) => [`${entry.kind}|${entry.value}`, entry.trusted]),
  );
  const status = (kind: "ip" | "country" | "state" | "device", value: string | null) => ({
    usual: value ? learned.get(`${kind}|${value}`) === true : false,
    known: value ? learned.has(`${kind}|${value}`) : false,
  });
  const openFindings = (record?.findings ?? []).filter((finding) => finding.state !== "resolved");
  const activity = directory?.signInActivity ?? null;
  const manager = directory?.manager?.data ?? null;

  return (
    <section className="people person" aria-labelledby={titleId}>
      <nav className="breadcrumb" aria-label="Breadcrumb">
        <button type="button" onClick={onBack}>
          People
        </button>
        <span aria-hidden="true">/</span>
        <span aria-current="page">{name}</span>
      </nav>

      {error ? (
        <p className="form-error" role="alert">
          {error}
        </p>
      ) : null}

      <section className="panel person-header">
        <div>
          <span className="eyebrow">Person</span>
          <h2 id={titleId}>{name}</h2>
          <p>
            {[upn, profile?.jobTitle, profile?.department].filter(Boolean).join(" · ") ||
              "No directory details"}
          </p>
          <div className="person-tags">
            {profile ? (
              profile.accountEnabled === false ? (
                <span className="person-tag person-tag--alert">Sign-in blocked</span>
              ) : profile.accountEnabled ? (
                <span className="person-tag person-tag--ok">Account enabled</span>
              ) : null
            ) : null}
            {profile?.userType ? <span className="person-tag">{profile.userType}</span> : null}
            {profile ? (
              <span className="person-tag">
                {profile.onPremisesSyncEnabled ? "Synced from on-premises" : "Cloud account"}
              </span>
            ) : null}
            {openFindings.length > 0 ? (
              <span className="person-tag person-tag--alert">
                {openFindings.length} open {openFindings.length === 1 ? "finding" : "findings"}
              </span>
            ) : record ? (
              <span className="person-tag person-tag--ok">Nothing open</span>
            ) : null}
          </div>
          {breachAddress ? (
            <p className="person-actions">
              <a
                className="button button--secondary"
                href={hibpLink(breachAddress)}
                target="_blank"
                rel="noreferrer noopener"
              >
                Check Have I Been Pwned
              </a>
              <span className="person-muted">
                Opens haveibeenpwned.com in a new tab with {breachAddress} already searched.
              </span>
            </p>
          ) : null}
        </div>
      </section>

      <section className="panel person-facts" aria-label="Directory details">
        <h3>Directory</h3>
        {directory === null ? (
          <p className="empty-data">Reading Microsoft's directory…</p>
        ) : directory.state !== "available" ? (
          <p className="empty-data">{directory.message ?? "Directory details are unavailable."}</p>
        ) : (
          <dl>
            {(
              [
                ["Email", profile?.mail],
                ["Office", profile?.officeLocation],
                ["Company", profile?.companyName],
                ["Employee ID", profile?.employeeId],
                [
                  "Manager",
                  manager
                    ? `${manager.displayName ?? manager.userPrincipalName ?? "Unnamed"}${manager.jobTitle ? `, ${manager.jobTitle}` : ""}`
                    : directory.manager?.state === "available"
                      ? "None set"
                      : "Unavailable",
                ],
                [
                  "Phone",
                  [profile?.mobilePhone, ...(profile?.businessPhones ?? [])]
                    .filter(Boolean)
                    .join(", "),
                ],
                [
                  "Location",
                  [profile?.city, profile?.state, profile?.country].filter(Boolean).join(", "),
                ],
                ["Usage location", profile?.usageLocation],
                ["Account created", profile?.createdDateTime ? day(profile.createdDateTime) : null],
                [
                  "Password last changed",
                  profile?.lastPasswordChangeDateTime
                    ? when(profile.lastPasswordChangeDateTime)
                    : null,
                ],
                [
                  "Last sign-in (Microsoft)",
                  activity?.state === "available"
                    ? when(activity.lastSignIn)
                    : activity?.state === "missingConsent"
                      ? "Needs sign-in log access"
                      : "Unavailable",
                ],
                [
                  "Last successful sign-in",
                  activity?.state === "available" ? when(activity.lastSuccessfulSignIn) : null,
                ],
                [
                  "Last background sign-in",
                  activity?.state === "available" ? when(activity.lastNonInteractiveSignIn) : null,
                ],
              ] as const
            ).map(([label, value]) => (
              <div key={label}>
                <dt>{label}</dt>
                <dd>{value || "Not set"}</dd>
              </div>
            ))}
          </dl>
        )}
      </section>

      {history ? (
        <section className="panel person-stats" aria-label="Stored sign-ins">
          <dl>
            {(
              [
                ["Sign-ins stored", history.totals.signIns],
                ["Successful", history.totals.successes],
                ["Failed", history.totals.failures],
                ["Blocked by policy", history.totals.blocked],
                ["Places", history.places.length],
                ["IP addresses", history.addresses.length],
                ["Devices", history.devices.length],
              ] as const
            ).map(([label, value]) => (
              <div key={label}>
                <dt>{label}</dt>
                <dd>{number.format(value)}</dd>
              </div>
            ))}
          </dl>
          <p>
            Stored from {when(history.totals.firstSeen)} to {when(history.totals.lastSeen)}. Last
            successful sign-in {when(history.totals.lastSuccess)}.
          </p>
        </section>
      ) : record ? (
        <p className="signin-coverage">
          Prism365 has no stored sign-ins for this person. They appear here once collection reads
          one.
        </p>
      ) : null}

      {record && record.findings.length > 0 ? (
        <section className="panel signin-findings" aria-label="Findings">
          <h3>Findings</h3>
          <ul className="finding-list">
            {record.findings.map((finding) => (
              <li key={finding.id}>
                <FindingItem
                  finding={finding}
                  showSubject={false}
                  busy={busy !== null}
                  onState={(state) =>
                    void act(finding.id, () =>
                      readJson(
                        `/api/v1/findings/${finding.id}/state`,
                        (value) => findingSchema.parse(value),
                        post(session, { state }),
                      ),
                    )
                  }
                />
              </li>
            ))}
          </ul>
        </section>
      ) : null}

      <section className="panel signin-table" aria-label="Licenses">
        <h3>Licenses</h3>
        {directory === null ? (
          <p className="empty-data">Reading licenses…</p>
        ) : directory.licenses?.state !== "available" ? (
          <p className="empty-data">
            {directory.licenses?.message ?? directory.message ?? "Licenses are unavailable."}
          </p>
        ) : directory.licenses.data.length === 0 ? (
          <p className="empty-data">No licenses assigned.</p>
        ) : (
          <>
            {directory.licenses.message ? (
              <p className="signin-coverage">{directory.licenses.message}</p>
            ) : null}
            <ul className="person-licenses">
              {directory.licenses.data.map((license) => {
                const active = license.servicePlans.filter(
                  (plan) => plan.provisioningStatus === "Success",
                );
                return (
                  <li key={license.id}>
                    <strong>{licenseName(license.skuPartNumber)}</strong>
                    <span>{license.skuPartNumber}</span>
                    <details>
                      <summary>
                        {active.length} of {license.servicePlans.length} services on
                      </summary>
                      <ul>
                        {license.servicePlans.map((plan) => (
                          <li key={plan.servicePlanId}>
                            {plan.servicePlanName}
                            <span className="person-muted"> {plan.provisioningStatus}</span>
                          </li>
                        ))}
                      </ul>
                    </details>
                  </li>
                );
              })}
            </ul>
          </>
        )}
      </section>

      {history ? (
        <>
          <div className="signin-tables">
            <section className="panel signin-table person-grid">
              <h3>Places</h3>
              <table>
                <caption className="visually-hidden">Places this person signed in from</caption>
                <thead>
                  <tr>
                    <th scope="col">Place</th>
                    <th scope="col" className="num">
                      Sign-ins
                    </th>
                    <th scope="col" className="num">
                      Failed
                    </th>
                    <th scope="col">Last seen</th>
                    <th scope="col">Status</th>
                  </tr>
                </thead>
                <tbody>
                  {history.places.map((row) => {
                    const country = status("country", row.country);
                    const state = status("state", row.state);
                    return (
                      <tr key={`${row.country}|${row.state}|${row.city}`}>
                        <th scope="row">{place(row)}</th>
                        <td className="num">{number.format(row.count)}</td>
                        <td className="num">{number.format(row.failures)}</td>
                        <td className="nowrap">{day(row.lastSeen)}</td>
                        <td>
                          <Status
                            usual={country.usual && (state.usual || !row.state)}
                            known={country.known || state.known}
                          />
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </section>
            <section className="panel signin-table person-grid">
              <h3>Apps</h3>
              <table>
                <caption className="visually-hidden">Apps this person signed in to</caption>
                <thead>
                  <tr>
                    <th scope="col">App</th>
                    <th scope="col" className="num">
                      Sign-ins
                    </th>
                    <th scope="col" className="num">
                      Failed
                    </th>
                    <th scope="col">Last used</th>
                  </tr>
                </thead>
                <tbody>
                  {history.apps.map((row) => (
                    <tr key={row.app ?? "-"}>
                      <th scope="row">{row.app ?? "Not reported"}</th>
                      <td className="num">{number.format(row.count)}</td>
                      <td className="num">{number.format(row.failures)}</td>
                      <td className="nowrap">{day(row.lastSeen)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </section>
          </div>

          <section className="panel signin-table person-grid">
            <h3>IP addresses</h3>
            <table>
              <caption className="visually-hidden">IP addresses this person signed in from</caption>
              <thead>
                <tr>
                  <th scope="col">Address</th>
                  <th scope="col">Location</th>
                  <th scope="col" className="num">
                    Sign-ins
                  </th>
                  <th scope="col" className="num">
                    Failed
                  </th>
                  <th scope="col">First seen</th>
                  <th scope="col">Last seen</th>
                  <th scope="col">Status</th>
                </tr>
              </thead>
              <tbody>
                {history.addresses.map((row) => {
                  const known = status("ip", row.ip);
                  return (
                    <tr key={row.ip}>
                      <th scope="row">
                        <code>{row.ip}</code>
                      </th>
                      <td className="person-text">{place(row)}</td>
                      <td className="num">{number.format(row.count)}</td>
                      <td className="num">{number.format(row.failures)}</td>
                      <td className="nowrap">{day(row.firstSeen)}</td>
                      <td className="nowrap">{when(row.lastSeen)}</td>
                      <td>
                        <Status {...known} />
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </section>

          <section className="panel signin-table person-grid">
            <h3>Devices</h3>
            <table>
              <caption className="visually-hidden">Devices this person signed in with</caption>
              <thead>
                <tr>
                  <th scope="col">Device</th>
                  <th scope="col">Device ID</th>
                  <th scope="col">Compliant</th>
                  <th scope="col">Managed</th>
                  <th scope="col">Join type</th>
                  <th scope="col" className="num">
                    Sign-ins
                  </th>
                  <th scope="col">Last seen</th>
                  <th scope="col">Status</th>
                </tr>
              </thead>
              <tbody>
                {history.devices.map((row) => {
                  const known = status("device", row.deviceId);
                  return (
                    <tr key={`${row.deviceId}|${row.operatingSystem}|${row.browser}`}>
                      <th scope="row">
                        {[row.operatingSystem, row.browser].filter(Boolean).join(" / ") ||
                          "Not reported"}
                      </th>
                      <td className="person-text">
                        {row.deviceId ? <code>{row.deviceId}</code> : "Unregistered"}
                      </td>
                      <td>{yesNo(row.compliant)}</td>
                      <td>{yesNo(row.managed)}</td>
                      <td>{row.trustType ?? "None"}</td>
                      <td className="num">{number.format(row.count)}</td>
                      <td className="nowrap">{when(row.lastSeen)}</td>
                      <td>
                        {row.deviceId ? (
                          <Status {...known} />
                        ) : (
                          <span className="person-muted">No ID</span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            <p className="signin-coverage">
              Sign-in logs do not carry MAC addresses. Those come from Intune device records, which
              Prism365 does not read yet.
            </p>
          </section>
        </>
      ) : null}

      <section className="panel signin-table person-grid person-history">
        <div className="panel__header">
          <h3>Every sign-in</h3>
          <fieldset className="signin-range">
            <legend className="visually-hidden">Sign-in result</legend>
            {(
              [
                ["all", "All"],
                ["success", "Successful"],
                ["failed", "Not successful"],
              ] as const
            ).map(([value, label]) => (
              <button
                key={value}
                type="button"
                aria-pressed={filter === value}
                className={filter === value ? "is-active" : ""}
                onClick={() => setFilter(value)}
              >
                {label}
              </button>
            ))}
          </fieldset>
        </div>
        {signins === null ? (
          <p className="empty-data">Loading sign-ins…</p>
        ) : signins.length === 0 ? (
          <p className="empty-data">No stored sign-ins match.</p>
        ) : (
          <>
            <table>
              <caption className="visually-hidden">Sign-ins, newest first</caption>
              <thead>
                <tr>
                  <th scope="col">When</th>
                  <th scope="col">Result</th>
                  <th scope="col">App</th>
                  <th scope="col">IP address</th>
                  <th scope="col">Location</th>
                  <th scope="col">Device</th>
                  <th scope="col">Client</th>
                  <th scope="col">Conditional Access</th>
                  <th scope="col">Risk</th>
                </tr>
              </thead>
              <tbody>
                {signins.map((row) => (
                  <tr key={row.id}>
                    <th scope="row" className="nowrap">
                      {when(row.createdAt)}
                    </th>
                    <td className="person-text">
                      <span className={`person-result person-result--${row.result}`}>
                        {resultLabels[row.result]}
                      </span>
                      {row.result !== "success" ? (
                        <span className="person-muted person-reason">
                          {row.failureReason ?? "No reason given"} ({row.errorCode})
                        </span>
                      ) : null}
                    </td>
                    <td className="person-text">{row.app ?? "Not reported"}</td>
                    <td className="person-text">
                      {row.ip ? <code>{row.ip}</code> : "Not reported"}
                    </td>
                    <td className="person-text">{place(row)}</td>
                    <td className="person-text">
                      {[row.operatingSystem, row.browser].filter(Boolean).join(" / ") ||
                        "Not reported"}
                      {row.compliant ? <span className="person-muted"> compliant</span> : null}
                    </td>
                    <td className="person-text">{row.clientApp ?? "Not reported"}</td>
                    <td className="person-text">{row.conditionalAccess ?? "Not reported"}</td>
                    <td className="person-text">
                      {row.riskLevel && row.riskLevel !== "none" && row.riskLevel !== "hidden"
                        ? `${row.riskLevel}${row.riskEventTypes.length ? ` (${row.riskEventTypes.join(", ")})` : ""}`
                        : "None reported"}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {next ? (
              <button
                className="button button--secondary person-more"
                type="button"
                disabled={busy !== null}
                onClick={() => {
                  setBusy("more");
                  loadSignins(next)
                    .catch((loadError: unknown) =>
                      setError(
                        loadError instanceof Error
                          ? loadError.message
                          : "Sign-ins could not be loaded.",
                      ),
                    )
                    .finally(() => setBusy(null));
                }}
              >
                Show older sign-ins
              </button>
            ) : null}
          </>
        )}
      </section>

      {record && !history && record.findings.length === 0 ? (
        <div className="signin-findings__empty">
          <ShieldIcon />
          <p>Nothing has been flagged about this person.</p>
        </div>
      ) : null}
    </section>
  );
}

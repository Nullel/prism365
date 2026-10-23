import type { PersonHistory, PersonListItem, PersonSignin } from "@prism365/contracts";
import { moduleTablePrefix } from "@prism365/modules";
import type { SqlClient } from "@prism365/storage";
import { conditionalAccessBlockCode, credentialFailureCodes, mfaFailureCodes } from "./source.ts";

// Per-account views over stored sign-ins: the people list, one account's history, and its
// sign-ins page by page. Unlike the tenant summary these show full addresses and devices, because
// telling whether one account is compromised needs them.

const prefix = moduleTablePrefix("signins.audit");
const events = `"${prefix}event"`;
const seen = `"${prefix}seen"`;
const failureList = [...credentialFailureCodes, ...mfaFailureCodes].join(",");
const iso = (value: Date | string) => new Date(value).toISOString();
const isoOrNull = (value: Date | string | null) => (value ? iso(value) : null);

export function signinResult(errorCode: number): PersonSignin["result"] {
  if (errorCode === 0) return "success";
  if (credentialFailureCodes.has(errorCode)) return "failure";
  if (mfaFailureCodes.has(errorCode)) return "mfaFailure";
  if (errorCode === conditionalAccessBlockCode) return "blocked";
  return "interrupted";
}

const latest = (column: string) =>
  `(array_agg(${column} ORDER BY created_at DESC) FILTER (WHERE ${column} IS NOT NULL))[1]`;

// Columns every grouped table shares: how often, how it went, and when.
const spanColumns = `count(*)::int AS n,
  count(*) FILTER (WHERE error_code = 0)::int AS successes,
  count(*) FILTER (WHERE error_code IN (${failureList}))::int AS failures,
  min(created_at) AS first_seen, max(created_at) AS last_seen`;

interface SpanRow {
  n: number;
  successes: number;
  failures: number;
  first_seen: Date | string;
  last_seen: Date | string;
}

const span = (row: SpanRow) => ({
  count: row.n,
  successes: row.successes,
  failures: row.failures,
  firstSeen: iso(row.first_seen),
  lastSeen: iso(row.last_seen),
});

export async function listPeople(
  sql: SqlClient,
  tenantId: string,
  options: { now: Date; days: number; search?: string; limit?: number },
): Promise<{ items: PersonListItem[]; truncated: boolean }> {
  const limit = Math.min(Math.max(options.limit ?? 500, 1), 2_000);
  const since = new Date(options.now.getTime() - options.days * 86_400_000).toISOString();
  const search = (options.search ?? "").trim().toLowerCase().slice(0, 200);
  const rows = await sql.query<{
    user_id: string;
    upn: string | null;
    display_name: string | null;
    last_seen: Date | string;
    last_success: Date | string | null;
    n: number;
    failures: number;
    countries: number;
  }>(
    `SELECT user_id, ${latest("upn")} AS upn, ${latest("display_name")} AS display_name,
            max(created_at) AS last_seen,
            max(created_at) FILTER (WHERE error_code = 0) AS last_success,
            count(*)::int AS n,
            count(*) FILTER (WHERE error_code IN (${failureList}))::int AS failures,
            count(DISTINCT country)::int AS countries
       FROM ${events}
      WHERE tenant_id = $1 AND created_at >= $2 AND user_id <> ''
        AND ($3 = '' OR strpos(lower(COALESCE(upn, '')), $3) > 0
             OR strpos(lower(COALESCE(display_name, '')), $3) > 0
             OR lower(user_id) = $3)
      GROUP BY user_id
      ORDER BY max(created_at) DESC, user_id
      LIMIT $4`,
    [tenantId, since, search, limit + 1],
  );
  const open = await sql.query<{ subject_id: string; n: number }>(
    `SELECT subject_id, count(*)::int AS n FROM prism365_finding
      WHERE tenant_id = $1 AND module = 'signins.audit' AND subject_kind = 'person'
        AND state <> 'resolved'
      GROUP BY subject_id`,
    [tenantId],
  );
  const openBy = new Map(open.rows.map((row) => [row.subject_id, row.n]));
  return {
    truncated: rows.rows.length > limit,
    items: rows.rows.slice(0, limit).map((row) => ({
      userId: row.user_id,
      userPrincipalName: row.upn,
      displayName: row.display_name,
      lastSignIn: iso(row.last_seen),
      lastSuccess: isoOrNull(row.last_success),
      signIns: row.n,
      failures: row.failures,
      countries: row.countries,
      openFindings: openBy.get(row.user_id) ?? 0,
    })),
  };
}

// Everything stored about one account, grouped. Null when Prism365 has no sign-ins for it.
export async function personHistory(
  sql: SqlClient,
  tenantId: string,
  userId: string,
): Promise<PersonHistory | null> {
  const where = "tenant_id = $1 AND user_id = $2";
  const params = [tenantId, userId];
  const totals = await sql.query<{
    n: number;
    successes: number;
    failures: number;
    blocked: number;
    first_seen: Date | string | null;
    last_seen: Date | string | null;
    last_success: Date | string | null;
    upn: string | null;
    display_name: string | null;
  }>(
    `SELECT count(*)::int AS n,
            count(*) FILTER (WHERE error_code = 0)::int AS successes,
            count(*) FILTER (WHERE error_code IN (${failureList}))::int AS failures,
            count(*) FILTER (WHERE error_code = ${conditionalAccessBlockCode})::int AS blocked,
            min(created_at) AS first_seen, max(created_at) AS last_seen,
            max(created_at) FILTER (WHERE error_code = 0) AS last_success,
            ${latest("upn")} AS upn, ${latest("display_name")} AS display_name
       FROM ${events} WHERE ${where}`,
    params,
  );
  const total = totals.rows[0];
  if (!total || total.n === 0) return null;

  const places = await sql.query<
    SpanRow & { country: string | null; state: string | null; city: string | null }
  >(
    `SELECT country, state, city, ${spanColumns} FROM ${events} WHERE ${where}
      GROUP BY country, state, city ORDER BY max(created_at) DESC LIMIT 200`,
    params,
  );
  const addresses = await sql.query<
    SpanRow & { ip: string; country: string | null; state: string | null; city: string | null }
  >(
    `SELECT ip, ${latest("country")} AS country, ${latest("state")} AS state,
            ${latest("city")} AS city, ${spanColumns}
       FROM ${events} WHERE ${where} AND ip IS NOT NULL
      GROUP BY ip ORDER BY max(created_at) DESC LIMIT 500`,
    params,
  );
  const devices = await sql.query<
    SpanRow & {
      device_id: string | null;
      device_os: string | null;
      device_browser: string | null;
      compliant: boolean | null;
      managed: boolean | null;
      trust_type: string | null;
    }
  >(
    `SELECT device_id, device_os, device_browser, bool_or(is_compliant) AS compliant,
            bool_or(is_managed) AS managed, ${latest("trust_type")} AS trust_type, ${spanColumns}
       FROM ${events} WHERE ${where}
      GROUP BY device_id, device_os, device_browser ORDER BY max(created_at) DESC LIMIT 200`,
    params,
  );
  const apps = await sql.query<SpanRow & { app_name: string | null }>(
    `SELECT app_name, ${spanColumns} FROM ${events} WHERE ${where}
      GROUP BY app_name ORDER BY count(*) DESC, app_name LIMIT 100`,
    params,
  );
  const baseline = await sql.query<{
    kind: "ip" | "country" | "state" | "device";
    value: string;
    trusted: boolean;
    first_seen: Date | string;
    last_seen: Date | string;
  }>(
    `SELECT kind, value, trusted, first_seen, last_seen FROM ${seen} WHERE ${where}
      ORDER BY kind, last_seen DESC LIMIT 500`,
    params,
  );

  return {
    userId,
    userPrincipalName: total.upn,
    displayName: total.display_name,
    totals: {
      signIns: total.n,
      successes: total.successes,
      failures: total.failures,
      blocked: total.blocked,
      firstSeen: isoOrNull(total.first_seen),
      lastSeen: isoOrNull(total.last_seen),
      lastSuccess: isoOrNull(total.last_success),
    },
    places: places.rows.map((row) => ({
      country: row.country,
      state: row.state,
      city: row.city,
      ...span(row),
    })),
    addresses: addresses.rows.map((row) => ({
      ip: row.ip,
      country: row.country,
      state: row.state,
      city: row.city,
      ...span(row),
    })),
    devices: devices.rows.map((row) => ({
      deviceId: row.device_id,
      operatingSystem: row.device_os,
      browser: row.device_browser,
      compliant: row.compliant,
      managed: row.managed,
      trustType: row.trust_type,
      ...span(row),
    })),
    apps: apps.rows.map((row) => ({ app: row.app_name, ...span(row) })),
    baseline: baseline.rows.map((row) => ({
      kind: row.kind,
      value: row.value,
      trusted: row.trusted,
      firstSeen: iso(row.first_seen),
      lastSeen: iso(row.last_seen),
    })),
  };
}

export type SigninResultFilter = "all" | "success" | "failed";

// Newest first. `before` is the `next` value of the previous page.
export async function personSignins(
  sql: SqlClient,
  tenantId: string,
  userId: string,
  options: { before?: string | null; limit?: number; result?: SigninResultFilter } = {},
): Promise<{ items: PersonSignin[]; next: string | null }> {
  const limit = Math.min(Math.max(options.limit ?? 100, 1), 500);
  let beforeAt: string | null = null;
  let beforeId = "";
  if (options.before) {
    const split = options.before.indexOf("|");
    const at = Date.parse(split > 0 ? options.before.slice(0, split) : "");
    if (!Number.isFinite(at)) throw new RangeError("Invalid page position.");
    beforeAt = new Date(at).toISOString();
    beforeId = options.before.slice(split + 1);
  }
  const result =
    options.result === "success"
      ? "AND error_code = 0"
      : options.result === "failed"
        ? "AND error_code <> 0"
        : "";
  const rows = await sql.query<{
    id: string;
    created_at: Date | string;
    error_code: number;
    failure_reason: string | null;
    app_name: string | null;
    resource_name: string | null;
    client_app: string | null;
    ip: string | null;
    city: string | null;
    state: string | null;
    country: string | null;
    lat: number | null;
    lon: number | null;
    device_id: string | null;
    device_os: string | null;
    device_browser: string | null;
    is_compliant: boolean | null;
    is_managed: boolean | null;
    trust_type: string | null;
    ca_status: string | null;
    risk_level: string | null;
    risk_state: string | null;
    risk_event_types: unknown;
    is_interactive: boolean | null;
  }>(
    `SELECT id, created_at, error_code, failure_reason, app_name, resource_name, client_app, ip,
            city, state, country, lat, lon, device_id, device_os, device_browser, is_compliant,
            is_managed, trust_type, ca_status, risk_level, risk_state, risk_event_types,
            is_interactive
       FROM ${events}
      WHERE tenant_id = $1 AND user_id = $2 ${result}
        AND ($3::timestamptz IS NULL OR (created_at, id) < ($3::timestamptz, $4))
      ORDER BY created_at DESC, id DESC
      LIMIT $5`,
    [tenantId, userId, beforeAt, beforeId, limit + 1],
  );
  const page = rows.rows.slice(0, limit);
  const last = page.at(-1);
  return {
    next: rows.rows.length > limit && last ? `${iso(last.created_at)}|${last.id}` : null,
    items: page.map((row) => {
      const types =
        typeof row.risk_event_types === "string"
          ? (JSON.parse(row.risk_event_types) as unknown)
          : row.risk_event_types;
      return {
        id: row.id,
        createdAt: iso(row.created_at),
        result: signinResult(row.error_code),
        errorCode: row.error_code,
        failureReason: row.failure_reason,
        app: row.app_name,
        resource: row.resource_name,
        clientApp: row.client_app,
        ip: row.ip,
        city: row.city,
        state: row.state,
        country: row.country,
        latitude: row.lat === null ? null : Number(row.lat),
        longitude: row.lon === null ? null : Number(row.lon),
        deviceId: row.device_id,
        operatingSystem: row.device_os,
        browser: row.device_browser,
        compliant: row.is_compliant,
        managed: row.is_managed,
        trustType: row.trust_type,
        conditionalAccess: row.ca_status,
        riskLevel: row.risk_level,
        riskState: row.risk_state,
        riskEventTypes: Array.isArray(types)
          ? types.filter((item): item is string => typeof item === "string").slice(0, 50)
          : [],
        interactive: row.is_interactive,
      };
    }),
  };
}

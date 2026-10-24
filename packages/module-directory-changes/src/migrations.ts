import type { ModuleMigration } from "@prism365/modules";

// Normalized directory audit events and the accounts and objects each one involves. Raw Graph
// payloads are never stored. Microsoft keeps these records 7 or 30 days; keeping them longer is
// the point of the module, and its retention rule deletes them after 10 years unless the
// operator chooses less (see docs/PRIVACY.md).
//
// `processed` makes the rules restartable: an event inserted by a run that stopped before its
// rules ran is judged by the next run instead of being skipped as already stored.
// The subject table answers "changes involving this person" with a plain indexed lookup, so no
// query has to look inside the targets JSON. `search_text` is the lower-case text the feed's
// search matches (see searchText in source.ts).

export const directoryChangesMigrations: readonly ModuleMigration[] = [
  {
    id: "001-events",
    sql: `CREATE TABLE IF NOT EXISTS mod_directory_changes_event (
  tenant_id uuid NOT NULL,
  id text NOT NULL,
  activity_at timestamptz NOT NULL,
  category text,
  activity text,
  operation_type text,
  result text,
  result_reason text,
  logged_by_service text,
  correlation_id text,
  initiated_by_kind text NOT NULL CHECK (initiated_by_kind IN ('user','app','unknown')),
  initiated_by_id text,
  initiated_by_name text,
  initiated_by_principal text,
  targets jsonb NOT NULL,
  details jsonb NOT NULL,
  search_text text NOT NULL,
  processed boolean NOT NULL,
  collected_at timestamptz NOT NULL,
  PRIMARY KEY (tenant_id, id)
);
CREATE INDEX IF NOT EXISTS mod_directory_changes_event_time ON mod_directory_changes_event (tenant_id, activity_at DESC, id DESC);
CREATE INDEX IF NOT EXISTS mod_directory_changes_event_pending ON mod_directory_changes_event (tenant_id, processed, activity_at);
SELECT prism365_tenant_table('mod_directory_changes_event');

CREATE TABLE IF NOT EXISTS mod_directory_changes_subject (
  tenant_id uuid NOT NULL,
  event_id text NOT NULL,
  subject_id text NOT NULL,
  subject_kind text NOT NULL CHECK (subject_kind IN ('actor','target')),
  subject_type text,
  activity_at timestamptz NOT NULL,
  PRIMARY KEY (tenant_id, event_id, subject_id, subject_kind)
);
CREATE INDEX IF NOT EXISTS mod_directory_changes_subject_lookup ON mod_directory_changes_subject (tenant_id, subject_id, activity_at DESC, event_id DESC);
SELECT prism365_tenant_table('mod_directory_changes_subject');`,
  },
];

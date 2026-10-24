-- Audit trends and scheduled audit runs. Re-runnable: core migrations run again on every start.
--
-- A trend point is the summary tiles of one successful audit run (inactive accounts, admins
-- without MFA, unused licenses...), copied when the run finishes. Runs are pruned to the newest
-- few per audit; points are a few hundred bytes each and nothing deletes them, so an audit's
-- numbers can be followed for as long as Prism365 runs. A point holds counts and the labels the
-- audit gave them, never an account.

-- How a run was started: by an administrator ('manual') or by the audit's schedule
-- ('schedule'). Added in place, so every run made before schedules existed reads as manual,
-- which it was.
ALTER TABLE prism365_audit_run ADD COLUMN IF NOT EXISTS trigger text NOT NULL DEFAULT 'manual'
  CHECK (trigger IN ('manual','schedule'));

CREATE TABLE IF NOT EXISTS prism365_audit_trend (
  tenant_id uuid NOT NULL,
  run_id uuid NOT NULL,
  audit_id text NOT NULL CHECK (length(audit_id) <= 80),
  audit_version integer NOT NULL CHECK (audit_version > 0),
  finished_at timestamptz NOT NULL,
  complete boolean NOT NULL,
  -- The account domains the run was limited to, as the run stored them; null for every domain.
  run_scope jsonb,
  -- The run's summary tiles: [{ id, label, value, tone }].
  tiles jsonb NOT NULL DEFAULT '[]'::jsonb,
  row_count integer NOT NULL CHECK (row_count >= 0),
  trigger text NOT NULL DEFAULT 'manual' CHECK (trigger IN ('manual','schedule')),
  PRIMARY KEY (tenant_id, run_id)
);
CREATE INDEX IF NOT EXISTS prism365_audit_trend_series
  ON prism365_audit_trend (tenant_id, audit_id, finished_at);
SELECT prism365_tenant_table('prism365_audit_trend');

-- Successful runs already stored become the first points. A tile's id is derived from its label
-- exactly as auditTileId() in packages/contracts/src/insights.ts derives it: every run of
-- characters other than ASCII letters and digits becomes one hyphen, ASCII letters are lowered
-- (translate, not lower(), so no database locale can change the result), the id is cut to 60
-- characters and stripped of hyphens at either end, and an id left empty is "tile". A label
-- whose id repeats an earlier tile's in the same run is dropped, as auditTrendTiles() drops it.
-- ON CONFLICT DO NOTHING makes this a no-op for runs that already have a point, so running it on
-- every start changes nothing.
--
-- The migrator connects as a superuser, which row security does not limit. A role that only
-- owns the tables sees no rows here (the tenant policy is forced), and the backfill then adds
-- nothing: points still start with the next successful run.
INSERT INTO prism365_audit_trend (tenant_id, run_id, audit_id, audit_version, finished_at,
  complete, run_scope, tiles, row_count, trigger)
SELECT run.tenant_id, run.id, run.audit_id, run.audit_version, run.finished_at, run.complete,
       run.run_scope,
       COALESCE((
         SELECT jsonb_agg(
                  jsonb_build_object('id', tile.id, 'label', tile.label, 'value', tile.value,
                    'tone', tile.tone)
                  ORDER BY tile.position)
           FROM (
             SELECT DISTINCT ON (item.id) item.id, item.label, item.value, item.tone,
                    item.position
               FROM (
                 SELECT COALESCE(NULLIF(btrim(left(translate(
                          regexp_replace(element.value->>'label', '[^A-Za-z0-9]+', '-', 'g'),
                          'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),
                          60), '-'), ''), 'tile') AS id,
                        element.value->>'label' AS label,
                        element.value->'value' AS value,
                        element.value->>'tone' AS tone,
                        element.position
                   FROM jsonb_array_elements(
                          CASE WHEN jsonb_typeof(run.summary) = 'array' THEN run.summary
                               ELSE '[]'::jsonb END)
                        WITH ORDINALITY AS element(value, position)
                  WHERE jsonb_typeof(element.value) = 'object'
                    AND jsonb_typeof(element.value->'label') = 'string'
                    AND jsonb_typeof(element.value->'value') = 'number'
               ) AS item
              ORDER BY item.id, item.position
           ) AS tile
       ), '[]'::jsonb),
       run.row_count, run.trigger
  FROM prism365_audit_run AS run
 WHERE run.status = 'succeeded' AND run.finished_at IS NOT NULL
ON CONFLICT (tenant_id, run_id) DO NOTHING;

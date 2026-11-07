-- Shared mailbox access from a read-only Exchange export (ADR-0012). Re-runnable.
--
-- Microsoft has no read-only API for Full Access, Send As or Send on Behalf, so an administrator
-- exports them with Export-MailboxAccess.ps1 in their own Exchange Online session and imports the
-- file. What is kept here is that file's claim, never state Prism365 checked. Only the latest
-- import is kept: a new import replaces the previous one, rows and all, in one transaction, and a
-- clear deletes it. Nothing is ever updated.

-- One row per tenant that is never deleted. An import and a clear both create it if it is missing
-- and then lock it, so they take turns even when no import exists yet to lock. It holds nothing.
CREATE TABLE IF NOT EXISTS prism365_mailbox_access_lock (
  tenant_id uuid NOT NULL,
  PRIMARY KEY (tenant_id)
);
SELECT prism365_tenant_table('prism365_mailbox_access_lock');

-- One row per tenant at most (the unique index on tenant_id): the import, when the export was
-- taken, who imported it, and its counts, so the pages never have to count the grants.
CREATE TABLE IF NOT EXISTS prism365_mailbox_access_import (
  tenant_id uuid NOT NULL,
  id uuid NOT NULL,
  imported_at timestamptz NOT NULL,
  exported_at timestamptz NOT NULL,
  imported_by uuid NOT NULL,
  imported_by_name text CHECK (length(imported_by_name) <= 500),
  include_user_mailboxes boolean NOT NULL,
  mailbox_count integer NOT NULL CHECK (mailbox_count BETWEEN 0 AND 20000),
  grant_count integer NOT NULL CHECK (grant_count BETWEEN 0 AND 200000),
  unresolved_count integer NOT NULL CHECK (unresolved_count BETWEEN 0 AND 200000),
  CHECK (unresolved_count <= grant_count),
  PRIMARY KEY (tenant_id, id)
);
-- Two imports racing for a tenant cannot both be kept: the second to commit fails on this index
-- and keeps nothing.
CREATE UNIQUE INDEX IF NOT EXISTS prism365_mailbox_access_import_latest
  ON prism365_mailbox_access_import (tenant_id);
SELECT prism365_tenant_table('prism365_mailbox_access_import');

-- One row per grant in the export, in the file's order: the mailbox, the right, and the trustee
-- as the script resolved it. A trustee the script could not resolve to exactly one recipient is
-- kept as its raw text, with no id and no address. The foreign key means a grant always belongs
-- to the one import there is, so grants of two imports can never be mixed.
CREATE TABLE IF NOT EXISTS prism365_mailbox_access_grant (
  tenant_id uuid NOT NULL,
  import_id uuid NOT NULL,
  position integer NOT NULL CHECK (position BETWEEN 0 AND 199999),
  mailbox_id uuid NOT NULL,
  mailbox_address text NOT NULL CHECK (length(mailbox_address) BETWEEN 3 AND 320),
  mailbox_name text NOT NULL CHECK (length(mailbox_name) <= 256),
  mailbox_type text NOT NULL CHECK (mailbox_type IN ('SharedMailbox','UserMailbox')),
  access_right text NOT NULL CHECK (access_right IN ('FullAccess','SendAs','SendOnBehalf')),
  trustee_id uuid,
  trustee_address text CHECK (length(trustee_address) BETWEEN 3 AND 320),
  trustee_name text NOT NULL CHECK (length(trustee_name) BETWEEN 1 AND 512),
  trustee_kind text NOT NULL CHECK (trustee_kind IN ('user','group','other','unresolved')),
  CHECK (trustee_kind <> 'unresolved' OR (trustee_id IS NULL AND trustee_address IS NULL)),
  CHECK (trustee_kind NOT IN ('user','group') OR trustee_id IS NOT NULL),
  PRIMARY KEY (tenant_id, import_id, position),
  CONSTRAINT prism365_mailbox_access_grant_import_fkey
    FOREIGN KEY (tenant_id, import_id) REFERENCES prism365_mailbox_access_import (tenant_id, id)
);
CREATE INDEX IF NOT EXISTS prism365_mailbox_access_grant_trustee
  ON prism365_mailbox_access_grant (tenant_id, trustee_id);
CREATE INDEX IF NOT EXISTS prism365_mailbox_access_grant_mailbox
  ON prism365_mailbox_access_grant (tenant_id, mailbox_id);
SELECT prism365_tenant_table('prism365_mailbox_access_grant');

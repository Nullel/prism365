import { createHash, randomUUID } from "node:crypto";
import { auditDefinition, findAudit, getRun } from "@prism365/audits";
import {
  buildHandoff,
  canonicalJson,
  handoffFileName,
  handoffListSchema,
  handoffView,
  hasPersonColumn,
  MAX_RECEIPT_BYTES,
  normalizeUserId,
  personIdSchema,
  personReceiptListSchema,
  queryAuditRows,
  receiptImportSchema,
  receiptItemPageSchema,
  receiptProblem,
  receiptSchema,
} from "@prism365/contracts";
import type { ModuleDatabase } from "@prism365/storage";
import type { FastifyInstance, FastifyReply } from "fastify";
import { z } from "zod";
import { auditExportParamsSchema, auditExportQuery } from "../audit-export.ts";
import {
  findReceipt,
  handoffExists,
  insertHandoff,
  insertReceipt,
  listHandoffs,
  listReceipts,
  personReceiptItems,
  pruneHandoffs,
  receiptItems,
  type StoredReceipt,
  statusCounts,
} from "../handoff-store.ts";
import type { RouteContext } from "./context.ts";

// Handoff lists for action tools, and the receipts they send back (docs/handoff.md).
//
// Prism365 never calls the action tool and shares nothing with it but files. The export writes
// one handoff row and sends the document; the import keeps a receipt exactly as the tool
// reported it. A receipt is the tool's claim: nothing here treats it as checked state, and there
// is no route to change or delete one.

const unavailable = {
  error: "handoffs_unavailable",
  message: "Handoffs need live mode with module storage, and a signed-in administrator.",
};
const signIn = {
  error: "authentication_required",
  message: "Sign in with an authorized administrator account.",
};
const unknownHandoff = {
  error: "unknown_handoff",
  message:
    "handoffId is not a handoff exported from this Prism365. Check that the receipt answers a handoff from this tenant, or send null.",
};
const runParams = z.object({ runId: z.uuid() });
// The columns the page shows, as comma-separated keys: up to 30 keys of up to 40 characters.
const columnsParam = z
  .string()
  .max(1_300)
  .regex(/^[A-Za-z0-9,]*$/)
  .optional();
const receiptParams = z.object({ id: z.guid().transform((value) => value.toLowerCase()) });
const pageQuery = z.object({
  after: z.coerce.number().int().min(0).max(19_999).optional(),
  limit: z.coerce.number().int().min(1).max(500).default(200),
});
const personParams = z.object({ userId: personIdSchema });
// The principal name the person's page shows, for items that name only that.
const personQuery = z.object({ principal: z.string().trim().min(1).max(500).optional() });

// How a receipt's digest is taken: the SHA-256 of the body as received, in canonical JSON (keys
// sorted, no whitespace), before any validation, cleaning or normalisation. Sending the same file
// again always gives the same digest, whatever later versions of Prism365 change in how they
// clean or normalise. Stored with each receipt; a change to this method takes a new version.
export const RECEIPT_DIGEST_VERSION = 1;

export function receiptDigest(body: unknown): string {
  return createHash("sha256").update(canonicalJson(body)).digest("hex");
}

// A receipt id already stored: the same receipt again is answered as before, a different one is
// refused. A kept receipt never changes; a tool corrects itself with a new receipt. A digest
// taken another way cannot be compared, so neither sameness nor difference can be claimed: the
// receipt is refused and the tool asked for a new id.
function answerStored(reply: FastifyReply, stored: StoredReceipt, digest: string) {
  if (stored.digestVersion !== RECEIPT_DIGEST_VERSION)
    return reply.status(409).send({
      error: "receipt_not_comparable",
      message:
        "A receipt with this id was stored by another version of Prism365 and cannot be compared; send it with a new id.",
    });
  if (stored.digest === digest)
    return reply
      .status(200)
      .send(receiptImportSchema.parse({ receipt: stored.summary, imported: false }));
  return reply.status(409).send({
    error: "receipt_conflict",
    message:
      "A different receipt with this id was already imported. A kept receipt is never changed: the tool should send a correction as a new receipt with its own id.",
  });
}

// The kind of failure a retention clean-up met, for the log. Never the driver's message.
function pruneFailure(error: unknown): string {
  const code = (error as { code?: unknown } | null)?.code;
  if (code === "40P01") return "deadlock";
  if (code === "55P03") return "lock";
  if (code === "23503") return "reference";
  if (error instanceof Error && error.name === "TransactionAbortedError") return "timeout";
  return "error";
}

// The foreign key from a receipt to the handoff it answers (006-handoffs.sql).
const HANDOFF_KEY = "prism365_receipt_handoff_fkey";

export function registerHandoffRoutes(app: FastifyInstance, context: RouteContext): void {
  const { store } = context;

  // Deletes a batch of handoffs and receipts past their retention, after a new one is stored and
  // committed, in a short transaction of its own. It never holds up or fails the request it
  // follows: what it cannot delete now (rows another request holds), a later one deletes, and a
  // failure is logged by its kind only.
  const pruneAfterStoring = async (database: ModuleDatabase, tenantId: string) => {
    try {
      const outcome = await database.tenant(
        tenantId,
        (sql) => pruneHandoffs(sql, tenantId, { days: context.settings.handoffRetentionDays }),
        { timeoutMilliseconds: 10_000 },
      );
      if (outcome.skipped) console.warn(`Prism365 handoff retention skipped (${outcome.skipped}).`);
    } catch (error) {
      console.warn(`Prism365 handoff retention skipped (${pruneFailure(error)}).`);
    }
  };

  // The accounts in an audit's current view, with reasons and evidence, as a file for the tool
  // that acts on them. It takes the same parameters as the CSV export and parses them the same
  // way, so both files of one view always hold the same accounts in the same order.
  app.get("/v1/audit-runs/:runId/handoff.json", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(signIn);
    const { modules, audits } = context;
    if (!modules || !audits) return reply.status(409).send(unavailable);
    const params = runParams.safeParse(request.params);
    const { columns, ...rest } = (request.query ?? {}) as Record<string, unknown>;
    const query = auditExportParamsSchema.safeParse(rest);
    const shown = columnsParam.safeParse(columns);
    if (!params.success || !query.success || !shown.success)
      return reply.status(400).send({ error: "invalid_request", message: "Check the filters." });
    const tenantId = session.tenantId;
    await audits.recover(tenantId);
    const run = await modules.database.tenant(tenantId, (sql) =>
      getRun(sql, tenantId, params.data.runId),
    );
    const audit = run ? findAudit(run.auditId) : undefined;
    if (!run || !audit || run.status !== "succeeded")
      return reply
        .status(404)
        .send({ error: "run_not_found", message: "That run has no results." });
    const definition = auditDefinition(audit);
    // Only audits whose rows are accounts a tool could act on say so; a person column alone (a
    // device's owner) is not enough.
    if (!definition.handoff || !hasPersonColumn(definition))
      return reply.status(404).send({
        error: "not_handed_off",
        message: "This audit's rows are not accounts for an action tool, so it has no handoff.",
      });
    // An account is handed off by its directory id. A run whose rows carry none (a usage report
    // that names accounts only by sign-in name, or conceals them) has nothing to hand off.
    if (!run.rows.some((row) => row.personId))
      return reply.status(404).send({
        error: "no_accounts",
        message: "This run names no accounts by their directory id, so it has nothing to hand off.",
      });
    const selection = auditExportQuery(query.data);
    // The CSV export takes an unknown view to mean every row. A handoff never does: it would hand
    // off accounts nobody chose.
    if (!handoffView(definition, selection.view ?? null))
      return reply
        .status(400)
        .send({ error: "unknown_view", message: "This audit has no view with that name." });
    const handoff = buildHandoff({
      id: randomUUID(),
      generatedAt: new Date().toISOString(),
      tenantId,
      generatedBy: { actorId: session.actorId, name: session.displayName },
      definition,
      run,
      query: selection,
      rows: queryAuditRows(definition, run.rows, selection),
      columns: shown.data === undefined ? null : shown.data.split(",").filter(Boolean),
    });
    // The ledger entry comes first, so no list ever leaves unrecorded. It holds ids and a count.
    await store.appendAudit(tenantId, session.actorId, {
      type: "handoff.exported",
      auditId: run.auditId,
      runId: run.id,
      handoffId: handoff.id,
      items: handoff.items.length,
      occurredAt: handoff.generatedAt,
    });
    const { source } = handoff;
    await modules.database.tenant(tenantId, (sql) =>
      insertHandoff(sql, tenantId, {
        id: handoff.id,
        auditId: source.auditId,
        auditTitle: source.auditTitle,
        runId: source.runId,
        view: source.view,
        filtered: source.search !== "" || Object.keys(source.facets).length > 0,
        scope: source.scope,
        complete: source.complete,
        items: handoff.items.length,
        generatedAt: handoff.generatedAt,
        generatedBy: { id: session.actorId, name: handoff.generatedBy.name },
      }),
    );
    await pruneAfterStoring(modules.database, tenantId);
    return reply
      .header("Content-Type", "application/json; charset=utf-8")
      .header(
        "Content-Disposition",
        `attachment; filename="${handoffFileName(run.auditId, run.startedAt, run.scope)}"`,
      )
      .send(`${JSON.stringify(handoff, null, 2)}\n`);
  });

  app.post(
    "/v1/receipts",
    {
      // A receipt may be four times the API's usual body limit, so only a signed-in
      // administrator's request is read that far. The global hook still checks the session and
      // the CSRF token once the body is in.
      bodyLimit: MAX_RECEIPT_BYTES,
      onRequest: async (request, reply) => {
        if (!(await context.sessionFor(request))) return reply.status(401).send(signIn);
      },
      // Problems reading the body, in words that fit a receipt file. The handler answers its own
      // validation problems; anything else is the server's fault and says nothing more.
      errorHandler: (error, _request, reply) => {
        const status =
          error instanceof Error && "statusCode" in error && typeof error.statusCode === "number"
            ? error.statusCode
            : 500;
        if (status === 413)
          return reply.status(413).send({
            error: "request_too_large",
            message:
              "A receipt can be at most 4 MB. Ask the tool to split it into several receipts.",
          });
        if (status === 415)
          return reply
            .status(415)
            .send({ error: "unsupported_media_type", message: "Send the receipt as JSON." });
        if (status >= 400 && status < 500)
          return reply
            .status(400)
            .send({ error: "invalid_receipt", message: "The receipt is not valid JSON." });
        return reply.status(500).send({
          error: "internal_error",
          message:
            "The receipt could not be imported. Try again: a receipt imported twice is kept once.",
        });
      },
    },
    async (request, reply) => {
      const session = await context.sessionFor(request);
      if (!session) return reply.status(401).send(signIn);
      const { modules } = context;
      if (!modules) return reply.status(409).send(unavailable);
      const parsed = receiptSchema.safeParse(request.body);
      if (!parsed.success)
        return reply
          .status(400)
          .send({ error: "invalid_receipt", message: receiptProblem(request.body, parsed.error) });
      const receipt = parsed.data;
      // Of the body as received (RECEIPT_DIGEST_VERSION), not of the cleaned receipt.
      const digest = receiptDigest(request.body);
      const tenantId = session.tenantId;
      const known = await modules.database.tenant(tenantId, async (sql) => ({
        stored: await findReceipt(sql, tenantId, receipt.id),
        handoff: receipt.handoffId ? await handoffExists(sql, tenantId, receipt.handoffId) : true,
      }));
      if (known.stored) return answerStored(reply, known.stored, digest);
      if (!known.handoff) return reply.status(400).send(unknownHandoff);
      const counts = statusCounts(receipt);
      // The ledger entry comes before anything is stored, so no receipt is ever kept unrecorded.
      // Ids and counts only: nothing the tool wrote in its own words.
      await store.appendAudit(tenantId, session.actorId, {
        type: "receipt.imported",
        receiptId: receipt.id,
        handoffId: receipt.handoffId,
        items: receipt.items.length,
        done: counts.done,
        failed: counts.failed,
        skipped: counts.skipped,
        occurredAt: new Date().toISOString(),
      });
      // The receipt and all its items in one transaction: any item that fails keeps nothing.
      let outcome: Awaited<ReturnType<typeof insertReceipt>>;
      try {
        outcome = await modules.database.tenant(tenantId, (sql) =>
          insertReceipt(sql, tenantId, receipt, {
            digest,
            digestVersion: RECEIPT_DIGEST_VERSION,
            importedBy: session.actorId,
          }),
        );
      } catch (error) {
        // The handoff was deleted for age between the check above and the insert. Only that key
        // means so, and only for a receipt that names a handoff.
        const failure = error as { code?: unknown; constraint?: unknown } | null;
        if (receipt.handoffId && failure?.code === "23503" && failure.constraint === HANDOFF_KEY)
          return reply.status(400).send(unknownHandoff);
        throw error;
      }
      // Another request kept the same id a moment earlier.
      if (!outcome.inserted) return answerStored(reply, outcome, digest);
      await pruneAfterStoring(modules.database, tenantId);
      return reply
        .status(201)
        .send(receiptImportSchema.parse({ receipt: outcome.summary, imported: true }));
    },
  );

  // The newest handoffs and receipts. No account appears here, only counts.
  app.get("/v1/handoffs", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(signIn);
    const { modules } = context;
    if (!modules) return reply.status(409).send(unavailable);
    const tenantId = session.tenantId;
    await store.appendAudit(tenantId, session.actorId, {
      type: "handoffs.viewed",
      occurredAt: new Date().toISOString(),
    });
    const list = await modules.database.tenant(tenantId, async (sql) => ({
      handoffs: await listHandoffs(sql, tenantId),
      receipts: await listReceipts(sql, tenantId),
    }));
    return handoffListSchema.parse(list);
  });

  // One receipt's items, as reported, a page at a time.
  app.get("/v1/receipts/:id", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(signIn);
    const { modules } = context;
    if (!modules) return reply.status(409).send(unavailable);
    const params = receiptParams.safeParse(request.params);
    const query = pageQuery.safeParse(request.query ?? {});
    if (!params.success || !query.success)
      return reply.status(400).send({ error: "invalid_request", message: "Check the receipt." });
    const tenantId = session.tenantId;
    // The items name accounts, so reading them is recorded first.
    await store.appendAudit(tenantId, session.actorId, {
      type: "receipt.viewed",
      receiptId: params.data.id,
      occurredAt: new Date().toISOString(),
    });
    const page = await modules.database.tenant(tenantId, async (sql) => {
      const stored = await findReceipt(sql, tenantId, params.data.id);
      if (!stored) return null;
      const items = await receiptItems(sql, tenantId, params.data.id, {
        after: query.data.after ?? null,
        limit: query.data.limit,
      });
      return { receipt: stored.summary, ...items };
    });
    if (!page)
      return reply
        .status(404)
        .send({ error: "receipt_not_found", message: "There is no such receipt." });
    return receiptItemPageSchema.parse(page);
  });

  // What other tools reported doing to one account: items that name its id, and items that name
  // only the principal name the person's page shows.
  app.get("/v1/people/:userId/receipts", async (request, reply) => {
    const session = await context.sessionFor(request);
    if (!session) return reply.status(401).send(signIn);
    const { modules } = context;
    if (!modules) return reply.status(409).send(unavailable);
    const params = personParams.safeParse(request.params);
    const query = personQuery.safeParse(request.query ?? {});
    if (!params.success || !query.success)
      return reply.status(400).send({ error: "invalid_request", message: "Unknown account." });
    const tenantId = session.tenantId;
    await store.appendAudit(tenantId, session.actorId, {
      type: "person.receipts.viewed",
      datasets: ["receipts"],
      occurredAt: new Date().toISOString(),
    });
    const result = await modules.database.tenant(tenantId, (sql) =>
      personReceiptItems(
        sql,
        tenantId,
        normalizeUserId(params.data.userId),
        query.data.principal ?? null,
      ),
    );
    return personReceiptListSchema.parse(result);
  });
}

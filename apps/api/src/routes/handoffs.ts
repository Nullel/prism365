import { createHash, randomUUID } from "node:crypto";
import { auditDefinition, findAudit, getRun } from "@prism365/audits";
import {
  buildHandoff,
  canonicalJson,
  handoffFileName,
  handoffListSchema,
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
  receiptItems,
  type StoredReceipt,
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

function receiptDigest(value: unknown): string {
  return createHash("sha256").update(canonicalJson(value)).digest("hex");
}

// A receipt id already stored: the same receipt again is answered as before, a different one is
// refused. A kept receipt never changes; a tool corrects itself with a new receipt.
function answerStored(reply: FastifyReply, stored: StoredReceipt, digest: string) {
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

export function registerHandoffRoutes(app: FastifyInstance, context: RouteContext): void {
  const { store } = context;

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
    if (!hasPersonColumn(definition))
      return reply.status(404).send({
        error: "no_accounts",
        message: "This audit lists no accounts, so it has nothing to hand off.",
      });
    const selection = auditExportQuery(query.data);
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
        items: handoff.items.length,
        generatedAt: handoff.generatedAt,
        generatedBy: { id: session.actorId, name: handoff.generatedBy.name },
      }),
    );
    return reply
      .header("Content-Type", "application/json; charset=utf-8")
      .header(
        "Content-Disposition",
        `attachment; filename="${handoffFileName(run.auditId, run.startedAt)}"`,
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
      // Of the receipt as validated and cleaned, so key order and spacing never matter.
      const digest = receiptDigest(receipt);
      const tenantId = session.tenantId;
      const known = await modules.database.tenant(tenantId, async (sql) => ({
        stored: await findReceipt(sql, tenantId, receipt.id),
        handoff: receipt.handoffId ? await handoffExists(sql, tenantId, receipt.handoffId) : true,
      }));
      if (known.stored) return answerStored(reply, known.stored, digest);
      if (!known.handoff)
        return reply.status(400).send({
          error: "unknown_handoff",
          message:
            "handoffId is not a handoff exported from this Prism365. Check that the receipt answers a handoff from this tenant, or send null.",
        });
      const importedAt = new Date();
      // The ledger entry comes first, so no receipt is ever kept unrecorded.
      await store.appendAudit(tenantId, session.actorId, {
        type: "receipt.imported",
        receiptId: receipt.id,
        handoffId: receipt.handoffId,
        items: receipt.items.length,
        tool: receipt.tool.name,
        occurredAt: importedAt.toISOString(),
      });
      // The receipt and all its items in one transaction: any item that fails keeps nothing.
      const outcome = await modules.database.tenant(tenantId, (sql) =>
        insertReceipt(sql, tenantId, receipt, {
          digest,
          importedAt,
          importedBy: session.actorId,
        }),
      );
      // Another request kept the same id a moment earlier.
      if (!outcome.inserted) return answerStored(reply, outcome, digest);
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

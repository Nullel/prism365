import { z } from "zod";
import { personIdSchema } from "./people.ts";

// Canary accounts (docs/modules/signins-audit.md): decoy accounts that nobody should ever sign in
// to. An administrator marks them; any sign-in attempt by one becomes a critical finding. The list
// is one setting per tenant in prism365_preference, key "signins.canaries".

export const MAX_CANARY_ACCOUNTS = 100;

export const canaryAccountSchema = z.object({
  userId: personIdSchema,
  // The sign-in name when the account was marked, kept so a sign-in that carries no user id can
  // still be matched, and so the list can say whom it means.
  principal: z.string().trim().min(1).max(320),
  addedAt: z.string().datetime({ offset: true }),
  // The administrator's object id, as in the access ledger. Never a name.
  addedBy: z.string().min(1).max(200),
});
export type CanaryAccount = z.infer<typeof canaryAccountSchema>;

export const canaryListSchema = z.object({
  accounts: z.array(canaryAccountSchema).max(MAX_CANARY_ACCOUNTS),
});
export type CanaryList = z.infer<typeof canaryListSchema>;

const accountIds = z.array(personIdSchema).max(MAX_CANARY_ACCOUNTS);

// A partial update: accounts to mark and accounts to stop marking. Everything else on the list is
// left alone.
export const canaryChangeSchema = z
  .object({ add: accountIds.optional(), remove: accountIds.optional() })
  .refine((change) => (change.add?.length ?? 0) + (change.remove?.length ?? 0) > 0, {
    message: "Name at least one account to add or remove.",
  })
  .refine((change) => !change.add?.some((id) => change.remove?.includes(id)), {
    message: "An account cannot be added and removed in one request.",
  });
export type CanaryChange = z.infer<typeof canaryChangeSchema>;

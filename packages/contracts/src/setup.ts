import { z } from "zod";

// The first-run setup checklist (GET /v1/setup/status, docs/SETUP.md). Without an allow-listed
// session it carries only what /v1/session already says: the mode, whether the installation is
// configured, and the names of missing settings. With one, it adds a check per item.

export const setupCheckStateSchema = z.enum(["ok", "warn", "fail", "unknown"]);
export type SetupCheckState = z.infer<typeof setupCheckStateSchema>;

export const setupCheckSchema = z.object({
  id: z.string().regex(/^[a-z][a-z0-9.-]{0,119}$/),
  state: setupCheckStateSchema,
  title: z.string().min(1).max(200),
  detail: z.string().max(1_000),
  // What to do about it, in plain words; null when nothing needs doing.
  fix: z.string().max(1_000).nullable(),
  // The read pack whose Microsoft consent link (/api/v1/auth/login?feature=) fixes this check.
  grantFeature: z
    .string()
    .regex(/^[a-z][a-z0-9.-]{0,79}$/)
    .nullable(),
});
export type SetupCheck = z.infer<typeof setupCheckSchema>;

export const setupStatusSchema = z.object({
  mode: z.enum(["demo", "live"]),
  configured: z.boolean(),
  missingConfiguration: z.array(z.string().max(200)).max(20),
  checks: z.array(setupCheckSchema).max(100),
  // True when this answer includes a test read of each granted pack (?probe=1).
  probed: z.boolean(),
});
export type SetupStatus = z.infer<typeof setupStatusSchema>;

export const setupStatusQuerySchema = z.object({ probe: z.enum(["0", "1"]).optional() });

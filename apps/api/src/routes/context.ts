import type { Audit } from "@prism365/audits";
import type { AuditAvailability } from "@prism365/contracts";
import type { GraphHttpClient } from "@prism365/graph";
import type { StateStore } from "@prism365/storage";
import type { FastifyRequest } from "fastify";
import type { AuditRunner } from "../audits.ts";
import type { AppConfig } from "../config.ts";
import type { IdentityService, IdentitySession } from "../identity.ts";
import type { DelegatedSessionCredential, ModuleRuntime } from "../modules.ts";

// What a feature's routes get from the app. Each feature keeps its routes in its own file under
// routes/ with one register function, so app.ts stays the place that wires them, not holds them.
// Every route still checks the session itself; this object grants nothing on its own.
export interface RouteContext {
  readonly config: AppConfig;
  readonly store: StateStore;
  readonly identity: IdentityService | null;
  // Live mode with module storage only; null in demo mode.
  readonly modules: ModuleRuntime | null;
  readonly audits: AuditRunner | null;
  sessionFor(request: FastifyRequest): Promise<IdentitySession | null>;
  // The opaque session handle from the cookie, for building a Graph client on that session.
  sessionHandle(request: FastifyRequest): string | undefined;
  // Microsoft permissions this installation enabled (User.Read plus ENABLED_READ_FEATURES).
  enabledPermissions(): string[];
  // A read-only Graph client on one administrator's session, narrowed to the enabled packs that
  // session granted, with the background reader's budgets.
  sessionGraph(
    session: IdentitySession,
    handle: string,
  ): { client: GraphHttpClient; scopes: string[] };
  auditAvailability(audit: Audit, session: IdentitySession): AuditAvailability;
  // The credential background collection runs on, or null when nothing is collecting.
  collectingCredential(): DelegatedSessionCredential | null;
}

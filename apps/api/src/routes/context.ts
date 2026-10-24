import type { Audit } from "@prism365/audits";
import type { AuditAvailability } from "@prism365/contracts";
import type { GraphFeatureId } from "@prism365/graph";
import type { CollectionStatus, ModuleRegistry, ScopedGraph } from "@prism365/modules";
import type { ModuleDatabase, StateStore } from "@prism365/storage";
import type { FastifyRequest } from "fastify";
import type { AuditRunner } from "../audits.ts";
import type { SupportedReadFeature } from "../config.ts";
import type { IdentitySession } from "../identity.ts";

// What a feature's routes get from the app. Each feature keeps its routes in its own file under
// routes/ with one register function, so app.ts stays the place that wires them, not holds them.
// Nothing here is a credential: no client secret, private key, token, encryption key or database
// address, and no way to act as an administrator other than the one making the request. Every
// route still checks the session itself.

// The configuration a route may see. Public facts only.
export interface RouteSettings {
  readonly mode: "demo" | "live";
  readonly origin: string;
  readonly production: boolean;
  // Live sign-in is configured (tenant, client, administrators and a credential).
  readonly configured: boolean;
  // Names of missing settings, as /v1/session already shows them.
  readonly missingConfiguration: readonly string[];
  readonly enabledFeatures: readonly SupportedReadFeature[];
  readonly credentialKind: "certificate" | "secret" | null;
  // The public certificate's fingerprint and expiry, read when the configuration loaded.
  readonly certificate: { readonly thumbprintSha256: string; readonly expiresAt: string } | null;
  // Module data survives a restart (PostgreSQL, or PGlite on disk).
  readonly durableStorage: boolean;
  // Sessions and the access ledger survive a restart (the encrypted PostgreSQL state store). They
  // do not with PRISM_DATA_DIR alone, which only keeps module data.
  readonly durableSessions: boolean;
  // How long handoffs and receipts are kept, in days (HANDOFF_RETENTION_DAYS).
  readonly handoffRetentionDays: number;
}

// Module storage for reading what was collected, and the collection status. Starting, stopping
// and binding collection stay in app.ts.
export interface RouteModules {
  readonly registry: ModuleRegistry;
  readonly database: ModuleDatabase;
  status(): CollectionStatus;
}

// A route asked for a read pack the operator did not enable.
export class RouteBoundaryError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "RouteBoundaryError";
  }
}

export interface RouteContext {
  readonly settings: RouteSettings;
  readonly store: StateStore;
  // Live mode with module storage only; null in demo mode.
  readonly modules: RouteModules | null;
  readonly audits: AuditRunner | null;
  sessionFor(request: FastifyRequest): Promise<IdentitySession | null>;
  // The opaque session handle from the cookie, for sessionGraph on that same session.
  sessionHandle(request: FastifyRequest): string | undefined;
  // Microsoft permissions this installation enabled (User.Read plus ENABLED_READ_FEATURES).
  enabledPermissions(): string[];
  // A read-only Graph client on the requesting administrator's session, narrowed to `packs`
  // (each must be enabled, or this throws RouteBoundaryError), and the permissions among those
  // packs that the session granted. Uses the background reader's budgets.
  sessionGraph(
    session: IdentitySession,
    handle: string,
    packs: readonly GraphFeatureId[],
  ): { graph: ScopedGraph; granted: string[] };
  auditAvailability(audit: Audit, session: IdentitySession): AuditAvailability;
  // Who is collecting, or null when nothing is. Never the credential itself.
  collectingAdmin(): { actorId: string } | null;
}

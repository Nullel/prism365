# ADR-0001: TypeScript workspace with explicit integration boundaries

- Status: accepted for scaffold
- Date: 2026-10-13
- Owners: initial maintainers
- Related issue: foundation milestone

## Context

The project needs a contributor-friendly web UI and API, shared runtime validation, and a
place to enforce Microsoft integration and responsible-use policy. Early contributors and
AI agents should be able to understand boundaries without learning multiple languages.

## Decision

Use npm workspaces with strict TypeScript. Keep deployable surfaces in `apps/` and reusable
boundaries in `packages/`. Use React/Vite for the web shell, Fastify for the API, Zod for
runtime contracts, and the built-in `fetch` API behind `packages/graph` for initial Graph
work. Use Biome, Vitest, and TypeScript as repository checks.

Do not choose the production database, queue, or deployment platform in this ADR.

## Consequences

- Contributors use one runtime and package manager.
- Runtime contracts can be shared without coupling UI to provider payloads.
- The Graph package can enforce cloud hosts, retries, and read-only defaults centrally.
- A later worker can reuse connector and policy packages.
- npm workspaces are less feature-rich than some monorepo orchestrators but add little
  machinery at this size.
- Database and job semantics remain open design work.

## Validation

CI formats, tests, type-checks, and builds all workspaces. Contract and connector tests cover
negative paths. A contributor can run the demo from a clean checkout using the README.

## Revisit when

Build/test time, deployment packaging, or cross-language connector needs make the simple
workspace an obstacle.

# Contributing to TenantLens

Thank you for helping build a safer, open Microsoft 365 administration platform.

## Start with an issue

For anything beyond a typo, open or claim an issue before implementation. A good issue has
one observable outcome, explicit non-goals, acceptance checks, and any permission or data
impact. Integration work must cite the relevant current Microsoft Learn documentation.

## Local setup

```bash
npm install
npm run check
npm run build
```

Run the API and web app in separate terminals with `npm run dev:api` and
`npm run dev:web`.

## Pull requests

- Keep a pull request focused and link its issue.
- Add or update tests and documentation.
- Complete every security, permission, privacy, and data-retention prompt in the template.
- Use synthetic fixtures. Redact screenshots and logs.
- Describe AI-generated or AI-assisted work and review it as carefully as any other code.
- Expect review from a code owner for permissions, connectors, auth, storage, and write paths.

Maintainers may close changes that add employee-ranking features, unnecessarily broad
permissions, copied proprietary material, or unsafe administration shortcuts.

## Design proposals

Open a design issue before adding a new service connector, database, deployment target,
permission, beta Graph endpoint, or write action. Record accepted architectural decisions
in `docs/decisions/` using the included template.

## Developer certificate of origin

By contributing, you certify that you have the right to submit the work under Apache-2.0.
Sign off commits with `git commit -s`. This project uses the
[Developer Certificate of Origin 1.1](https://developercertificate.org/).

## Community standards

Be kind, specific, and evidence-led. Review ideas and code, not people. Harassment,
discrimination, doxxing, and disclosure of tenant data are not tolerated. See
`CODE_OF_CONDUCT.md` for reporting guidance.

# GitHub publishing checklist

The local repository is not public until GitHub returns the repository URL and an anonymous
request can read it. Do not skip that verification.

## Before publishing

- [ ] Review the name, description, Apache-2.0 license, NOTICE, and trademark disclaimer.
- [ ] Review every tracked file with `git status` and `git diff --cached`.
- [ ] Run `npm run check` and `npm run build` from a clean checkout.
- [ ] Confirm fixtures and screenshots are synthetic and no secret scanner reports findings.
- [ ] Confirm the authenticated GitHub login is exactly `nlj3`.
- [ ] Confirm `https://github.com/nlj3/prism365` is still available.

GitHub CLI account access was verified as `nlj3` on 2026-10-14. If access expires, authenticate
interactively and re-verify the login before any publishing action:

```bash
gh auth login -h github.com
gh api user --jq .login
```

Only if the second command prints `nlj3`, create the public repository:

```bash
gh repo create nlj3/prism365 \
  --public \
  --source=. \
  --remote=origin \
  --push \
  --description="Open-source Microsoft 365 administration, reporting, and auditing"
```

## Immediately after publishing

- [ ] Open the repository in a logged-out/private browser and verify README visibility.
- [ ] Enable private vulnerability reporting, secret scanning, push protection, and
  Dependabot alerts where the account supports them.
- [ ] Protect `main`: require a pull request, required CI/CodeQL checks, conversation
  resolution, and no force pushes or deletion.
- [ ] Enable Discussions only if a maintainer will moderate them.
- [ ] Create milestones from `docs/ROADMAP.md`, beginning with “M0 trustworthy foundation.”
- [ ] Create the eight scoped issues in the roadmap and label `good first issue` only where
  the acceptance criteria are self-contained.
- [ ] Add repository topics such as `microsoft-365`, `microsoft-graph`, `entra-id`,
  `open-source`, `security`, and `reporting`.
- [ ] Verify CODEOWNERS resolves to the intended account.
- [ ] Replace the security contact fallback with GitHub's private reporting URL.

## Public verification

Publishing is complete only when all are true:

1. `gh repo view nlj3/prism365 --json isPrivate,url` reports `isPrivate: false`.
2. The returned URL is `https://github.com/nlj3/prism365`.
3. An unauthenticated request can read the repository README.
4. CI has run on `main` and the required checks are visible.

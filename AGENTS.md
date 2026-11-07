# AGENTS.md — Allo Recurring Task Tracker

Working rules for agents and humans in this repository. Read before touching
anything. Sections have stable anchors — cite them in review ("violates §2.1")
instead of pasting prose.

## §0 Branch hygiene (required)

Merged branches are deleted in the same task that merges them — never accumulated.

- **Before creating a branch or PR**, sweep stale ones first:
  ```bash
  git fetch --all --prune
  gh pr list --state all --head <branch> --json number,state
  git ls-remote --heads origin <branch> | wc -l
  ```
- **After a PR merges**, delete both halves (`gh pr merge --delete-branch`).
- **Squash merges make ancestry checks lie** — verify by content:
  `git diff origin/main <branch> -- <files>` (empty = already in main).
- **Never delete**: `main`, a branch with an OPEN PR, or a branch whose content
  genuinely differs from `main`. Ask when unsure.
- Prefer `git-wt switch --create agent/<task-slug>` for isolated work; never
  commit directly to `main`/`master`.

## §1 Background

- **Project:** Google Sheets + Apps Script prototype for the Allo Innoware
  AI Operations assignment — recurring task tracking, employee pending
  reports, manager alerts, leadership escalation. Full spec: `SOLUTION.md`.
- **Runtime:** Google Apps Script (V8). `Code.gs` is the ENTIRE backend —
  pasted into Extensions → Apps Script of the bound sheet. There is no
  build step and no package manifest; do not add one.
- **Package manager:** `pnpm` if ever needed; via `sfw` when on PATH
  (`sfw pnpm ...`). `git`, `rg`, `fd`, `eza`, `bat` run bare.

### §1.1 Map

- `Code.gs` — single-file script. Sections, in order: constants → setup
  (`setupSheets`, `installTriggers`, `onOpen`) → cycle engine
  (`ensureCycles_`) → daily checks (`runDailyChecks`: alerts + escalation)
  → mutations (`completeCycle`, `acknowledge`) → `onEdit` router → renderers
  (`renderMyTasks_`, `renderDashboard_`) → demo helpers → utilities.
- `tests/` — `node:test` regression suite; `harness.mjs` mocks the GAS
  services (SpreadsheetApp, LockService, MailApp, Utilities) with in-memory
  sheets so `Code.gs` runs under plain Node.
- `SOLUTION.md` — sheet structure, SLA reasoning, the 8 demo test scripts,
  submission email draft.
- `.github/` — CI (`node --check` + `node --test` + actionlint), Dependabot
  for actions, OSSF Scorecard, PR template, CODEOWNERS.
- `.coderabbit.yaml` — CodeRabbit PR review config (assertive, free tier).
- `greptile.json` / `repomix.config.json` — Greptile + repomix settings.

## §2 Behaviour

### §2.1 Architecture contract (non-negotiable)

- **Dedupe is the product.** Every mutable event carries a unique key —
  `CycleKey = TaskID|cycleStart`, `MGR_ALERT|key`, `ESC|key`, `ACK|key`,
  `COMPLETE|key`. Never introduce a write path that skips the key check;
  `logKeySet_()`/`cycleKeySet_()` are the guards, `Log` is the audit.
- **All mutations take `LockService.getScriptLock()`** — a time-driven
  trigger and a human edit can collide.
- **No duplicate-work writes:** `completeCycle` must refuse a
  `Status === 'Completed'` row; `acknowledge` must refuse a stamped row.
- **Employee single-point update:** completion data enters ONLY through
  `MyTasks` → `onEdit` → `completeCycle`. Do not add a second completion UI.
- **Statuses:** `Scheduled | Open | Completed`. "Overdue" is COMPUTED
  (`Open && DueDate < TODAY()`), never stored — do not add an Overdue status.
- **Config over code:** emails and SLA knobs (`ALERT_HOUR`,
  `ESCALATE_AFTER_DAYS`, `MGR_ACK_SLA_HOURS`) live in the `Config` tab;
  never hardcode them in logic.
- **No personal data / no real emails.** `example.com` placeholders only;
  `sendMail_()` must keep its placeholder guard (logs `EMAIL_SKIPPED`,
  still stamps for dedupe).
- **Never invent results:** record test runs/demos only when they actually
  happened. Mark everything else planned/TODO.

### §2.2 Verify (before you commit)

```bash
node --check Code.gs        # syntax gate (CI runs this)
node --test tests/          # regression suite (CI runs this)
git diff --check            # no whitespace errors
```

Editing `.github/workflows/` → also run actionlint
(`pnpm dlx actionlint` or CI's lint job).

### §2.3 Push / PR flow

- Branch off `main` → Conventional Commit
  (`feat|fix|docs|test|refactor|ci|chore(scope):`) → push → `gh pr create`
  → wait for CI + CodeRabbit review → fix findings → merge
  (`--squash --delete-branch`).
- Never push straight to `main`; a direct push skips CodeRabbit and CI.
- Bounded waits: `gh pr checks --watch`, ≤2 repolls, 1 rerun for infra
  flakes — then stop and report.

## §3 Output

- Commits: Conventional Commits, focus on why.
- PRs: summary + test plan (see `.github/pull_request_template.md`).
- Behaviour changes to `Code.gs` → update `SOLUTION.md` and the relevant
  test in the same change.

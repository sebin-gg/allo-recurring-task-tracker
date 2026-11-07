# Allo Recurring Task Tracker

Google Sheets + Apps Script prototype for the Allo Innoware AI Operations
assignment: recurring task cycles (daily / 7 / 15 / 30-day), a single-screen
employee pending report, manager dashboard with acknowledgements, same-day
overdue email alerts, and leadership escalation — all deduplicated via unique
keys so re-runs never double-create cycles, completions, or emails.

## Files

| File | Purpose |
|---|---|
| `Code.gs` | Entire backend — paste into Extensions → Apps Script |
| `SOLUTION.md` | Sheet structure, SLA design, all 8 demo test scripts, email draft |
| `AGENTS.md` | Contributor/agent rules (architecture contract in §2.1) |
| `tests/` | `node:test` regression suite — GAS services mocked in-memory |

## Quick start

1. New Google Sheet → Extensions → Apps Script → paste `Code.gs`.
2. Menu **Recurring Tasks → Setup sheets**, authorize, then **Install daily trigger**.
3. Put your own test emails in the `Config` tab (`example.com` is never mailed).
4. Full walkthrough + demo scripts: `SOLUTION.md`.

## Verify

```bash
node --check Code.gs   # copy to .js first on Windows
node --test tests/
```

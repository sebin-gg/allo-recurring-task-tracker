# Allo Innoware — Recurring Task Tracker (Google Sheets + Apps Script)

Full solution for the AI Operations / Founder's Office assignment. One bound
spreadsheet + one `Code.gs` file. No personal data — all emails are your own
test addresses set in `Config`.

---

## 1. Design at a glance

```
TaskMaster ──► Cycles (one row per task-cycle, system-owned)
                  ▲   │
   MyTasks ───────┘   │        employee: sees pending + enters Ticket ID once
   (employee view)    │
                  Dashboard    manager: pending/overdue + acknowledgement
                  Log          dedupe/audit trail for every event & email
                  Config       emails + SLA knobs
```

- **CycleKey** = `TaskID|cycleStartDate` → cycles can never duplicate.
- **LogKey** = `MGR_ALERT|CycleKey`, `ESC|CycleKey`, `ACK|CycleKey`, etc. →
  emails/acks can never duplicate, no matter how many times the job re-runs.
- `LockService` on every mutation → no race between trigger and human edits.

---

## 2. Sheet structure

### Tab 1 — `Config` (key/value)

| Key | Value |
|---|---|
| MANAGER_NAME | Manager |
| MANAGER_EMAIL | *(your test email)* |
| LEADERSHIP_EMAIL | *(your test email, comma-separate for several)* |
| ALERT_HOUR | 18 |
| ESCALATE_AFTER_DAYS | 1 |
| MGR_ACK_SLA_HOURS | 24 |

### Tab 2 — `TaskMaster` (task definitions; the only data-entry for admins)

| TaskID | Task Name | Assignee | Assignee Email | Freq Days | Requires Ticket | Anchor Date | Active | Notes |
|---|---|---|---|---|---|---|---|---|
| T-001 | Daily Operations Update | Employee A | emp-a-test@example.com | 1 | N | 2026-11-07 | Y | WhatsApp group update |
| T-002 | Portal Task A | Employee B | emp-b-test@example.com | 7 | Y | 2026-11-07 | Y | Portal ticket |
| T-003 | Portal Task B | Employee C | emp-c-test@example.com | 15 | Y | 2026-11-07 | Y | Portal ticket |
| T-004 | Portal Task C | Employee D | emp-d-test@example.com | 30 | Y | 2026-11-07 | Y | Portal ticket |

Anchor Date = first cycle's start. Due = start + freq − 1.

### Tab 3 — `Cycles` (auto-managed; one row per occurrence)

| Col | Field | Notes |
|---|---|---|
| A | CycleKey | `T-002|2026-11-07` — unique |
| B–F | TaskID, Task Name, Assignee, Assignee Email, Freq Days | copied from master |
| G | Cycle Start | anchor + k×freq |
| H | Due Date | start + freq − 1 |
| I | Status | Scheduled → Open → Completed (Overdue is computed: Open & due<today) |
| J–L | Completed By / At / Ticket-Proof | written once from `MyTasks` |
| M | Mgr Alert Sent At | set when alert email sent |
| N | Escalated At | set when escalation email sent |
| O–P | Mgr Ack By / At | written from `Dashboard` ack checkbox |

### Tab 4 — `MyTasks` (the employee's ONE screen)

- `B2` dropdown → pick your name; list re-renders automatically.
- Rendered table: **CycleKey | Task | Due | Status | Proof needed | Ticket ID / proof → | Done?**
- Workflow: read the row → type Ticket ID (or a short note for the daily task)
  → tick **Done**. That is the entire employee interaction; the completion is
  stamped into `Cycles` (who/when/proof) and reflected on the Dashboard and Log.

### Tab 5 — `Dashboard` (the manager's ONE screen)

- `B2` manager name (stamped on acknowledgements).
- Section 1: per-employee counts via live formulas:
  - Pending: `=COUNTIFS(Cycles!$D:$D,$A6,Cycles!$I:$I,"Open",Cycles!$H:$H,">="&TODAY())`
  - Overdue: `=COUNTIFS(Cycles!$D:$D,$A6,Cycles!$I:$I,"Open",Cycles!$H:$H,"<"&TODAY())`
  - Completed 30d: `=COUNTIFS(Cycles!$D:$D,$A6,Cycles!$I:$I,"Completed",Cycles!$K:$K,">="&TODAY()-30)`
- Section 2: **Needs acknowledgement** — alerted-but-open items with an `Ack?` checkbox.
- Section 3: **All open & overdue items** — due date, status, alert/escalation/ack flags.

### Tab 6 — `Log` (append-only audit + dedupe store)

`LogKey | Type | CycleKey | At | Detail` — types: `CYCLE_OPEN`, `COMPLETE`,
`MGR_ALERT`, `MGR_ACK`, `ESCALATION`, `EMAIL_SKIPPED`, `ERROR`.

---

## 3. Recurrence logic (R3 — no manual date math)

For a task with frequency **F** and anchor **A**, cycle *k* covers
`A + kF … A + kF + F − 1`; the due date is the last day. `ensureCycles()`
(runs inside every daily check and every completion) creates the **current**
cycle and pre-creates the **next** one as `Scheduled`. Completing a cycle
therefore needs zero date math — the next cycle already exists and flips to
`Open` on its start date. Daily = F1; weekly = F7; etc.

## 4. Daily-task method (T3 — no ticket exists)

For `Requires Ticket = N` tasks the employee simply **ticks Done** on
`MyTasks`; typing a note is optional. The system auto-records
`Self-attested in MyTasks sheet` plus **who** (the selected employee) and
**when** (timestamp). For a WhatsApp-style task the practical proof is the
one-click self-attestation + optional note like `Posted 9:15am`; a manager
can spot-check the group against the timestamp. Minimum manual effort: one
click. Ticket tasks enforce a non-empty Ticket ID — Done is rejected without it.

## 5. SLA / timing logic (T7 — with reasoning)

| Event | Timing | Reasoning |
|---|---|---|
| Cycle opens | start of cycle period (auto) | Employee sees it the whole window |
| Manager alert | daily trigger at `ALERT_HOUR` (18:00): any task due ≤ today & incomplete | End-of-day gives the full working day; "same day" = employee is warned the day it was due, not next morning |
| Escalation (path A) | `today ≥ due + ESCALATE_AFTER_DAYS` (default 1) and still incomplete | One full grace day after the alert to finish the work |
| Escalation (path B) | alert sent ≥ `MGR_ACK_SLA_HOURS` (24h) ago and manager never acknowledged | Covers "the required manager action did not happen" — silence escalates |
| Repeat emails | none — each event is a unique `LogKey` | Avoids alert fatigue; the Log is also the audit trail |

Why this shape: alerts fire once per cycle (not every run), escalation has two
independent triggers (work undone **or** review undone), and every knob lives
in `Config` so the company can tune without touching code.

## 6. Setup steps

1. New Google Sheet → rename it e.g. `Allo Recurring Task Tracker`.
2. `Extensions → Apps Script` → delete stub → paste all of `Code.gs` → Save.
3. Back in the sheet, reload → menu **Recurring Tasks** appears.
4. Run **Setup sheets** (authorize when prompted: Sheets + Mail + Triggers).
   Tabs, headers, validation, demo tasks and the first cycles are created.
5. In `Config`, replace `manager-test@example.com` and
   `leadership-test@example.com` with **your own test addresses**. (Placeholder
   `example.com` addresses are never emailed; the script logs `EMAIL_SKIPPED`
   instead, so demos still work before you set real addresses.)
6. Run **Install daily trigger** (equivalently: Apps Script → Triggers → add
   trigger → `runDailyChecks` → time-driven → daily → 18:00–19:00).
7. Share the sheet with Editor access for submission.

---

## 7. Test scripts — all 8 demos

Each demo lists *action → expected result*. Use the **Recurring Tasks** menu.

### T1 — Recurring cycle auto-reopens
1. `MyTasks` → select **Employee B** → Portal Task A row → enter `TKT-1001` → tick Done.
2. Check `Cycles`.
**Expected:** cycle `T-002|<start>` = Completed with name/timestamp/`TKT-1001`;
a new `T-002|<start+7d>` row already exists (Scheduled/Open). Log shows
`COMPLETE|…` and `CYCLE|…` entries. No date was typed by hand.

### T2 — Single-point completion
1. Same as above but note you never left `MyTasks`.
2. Check `Dashboard` and `Log`.
**Expected:** the item disappears from MyTasks, Dashboard summary counts it
under Completed, and the Ticket ID appears in Cycles col L — all from the
single entry in MyTasks.

### T3 — Daily task without ticket
1. `MyTasks` → **Employee A** → Daily Operations Update → tick Done (optionally
   type `Posted in WA group`).
**Expected:** Completed with proof `Self-attested in MyTasks sheet` (or your
note); tomorrow's daily cycle already exists. Zero ticket needed, one click.

### T4 — Same-day manager alert
1. `Recurring Tasks → DEMO: set first open task due today`.
2. `Recurring Tasks → Run daily checks now`.
**Expected:** one batched email to `MANAGER_EMAIL`, subject `[OVERDUE]…`;
Cycles col M stamped; Log `MGR_ALERT|key`. (With placeholder emails: Log shows
`EMAIL_SKIPPED` with the intended subject — the flow is identical.)

### T5 — Manager acknowledgement
1. After T4, open `Dashboard` → Needs acknowledgement → tick the Ack? checkbox.
2. Now complete the task from MyTasks.
**Expected:** Cycles cols O–P stamped with manager name + time; item leaves
the ack list. Dashboard shows ack happened *and* whether the task was
subsequently completed — both facts independently visible.

### T6 — Leadership escalation
1. `Recurring Tasks → DEMO: age first open task for escalation` (due yesterday,
   alert stamped 25h ago, no ack).
2. `Run daily checks now`.
**Expected:** email to `LEADERSHIP_EMAIL` (cc manager), subject
`[ESCALATION]…`; Cycles col N stamped; Log `ESC|key` with reason
`manager alert unacknowledged`. Also works via path A (task simply left
undone past `ESCALATE_AFTER_DAYS`).

### T7 — SLA logic demonstration
Point the evaluator at `Config` + §5 of this doc: change `ESCALATE_AFTER_DAYS`
to `0`, re-run daily checks on an alerted item → escalation fires immediately,
proving the knobs are live and the timing rule is data-driven.

### T8 — Re-run dedupe
1. `Run daily checks now` three times in a row after T4/T6.
2. Complete an already-Completed cycle again (tick Done twice fast).
3. Re-run `Setup sheets`.
**Expected:** Cycles gains **zero** new rows for existing cycle keys; Log gains
no second `MGR_ALERT|`/`ESC|`/`ACK|` key for the same cycle; second Done is
rejected with "Already completed". Emails: exactly one alert + one escalation
per cycle, ever.

---

## 8. Email write-up (send to hr.ops.allo@gmail.com)

> **Subject: Assignment Submission — Recurring Task Tracker (Google Sheets + Apps Script)**
>
> Hi team,
>
> Please find my completed assignment: a working Google Sheets prototype for
> recurring task tracking, pending reports, and escalation automation.
>
> **Link:** `<your sheet link>` (Editor access granted)
>
> **How it's organized:** `TaskMaster` holds the task definitions; `Cycles`
> stores one auto-generated row per task occurrence; `MyTasks` is the single
> screen each employee uses to see pending work and submit a Ticket ID (or a
> one-click completion for the daily non-ticket task); `Dashboard` gives the
> manager per-employee pending/overdue counts plus an acknowledgement
> checklist; `Log` is the dedupe/audit trail; `Config` holds the SLA settings.
>
> **Automation:** a daily 18:00 trigger sends a same-day alert to the manager
> for anything due and incomplete, and escalates to leadership when the item
> stays unresolved past the configured SLA or the manager alert goes
> unacknowledged for 24h. Unique keys + a log sheet guarantee re-runs never
> duplicate cycles, completions, or emails.
>
> All 8 required test cases (recurring cycle, single-point update, daily
> task, same-day alert, acknowledgement, escalation, SLA rationale, dedupe)
> are demonstrated in the sheet — a step-by-step test guide is included in
> the submission document.
>
> Only test email addresses were used; no personal data is included.
>
> Thanks,
> `<name>` · `<phone>` · `<portfolio link>`

---

## 9. Requirement → solution traceability

| Brief requirement | Implementation |
|---|---|
| Recurring cycle logic, auto re-due | `ensureCycles_()` + CycleKey scheme |
| Employee pending report | `MyTasks` (selector + rendered list) |
| Single-point updating | Ticket/proof + Done checkbox on `MyTasks`; `onEdit` → `completeCycle` |
| Ticket-based completion | `Requires Ticket` enforcement |
| Daily-task completion method | Self-attest checkbox + optional note (§4) |
| Who/when/proof capture; next cycle auto | `Cycles` cols J–L; next cycle pre-created |
| Manager visibility | `Dashboard` summary formulas + open-items table |
| Manager acknowledgement | Ack checkbox → `Cycles` O–P + `ACK` log |
| Same-day manager email | `runDailyChecks` @ `ALERT_HOUR` |
| Leadership escalation | Two-path SLA → `ESCALATION` email + log |
| No duplicate alerts/work | CycleKey + LogKey uniqueness + LockService + status guards |

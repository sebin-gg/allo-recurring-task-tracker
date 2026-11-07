import { test } from 'node:test';
import assert from 'node:assert/strict';
import { loadGas, makeWorld, cycleRows, logKeys, setConfig, clock } from './harness.mjs';

// Anchor date inside the code is startOfDay(clock.now) — keep assertions
// relative (counts, statuses, keys) so the suite is timezone-independent.
clock.now = new Date(2026, 10, 7, 12, 0, 0); // Nov 7 2026, midday LOCAL time

const boot = () => {
  const { api, world } = loadGas(makeWorld());
  api.setupSheets();
  return { api, world };
};

const keyOf = (world, taskPrefix) => cycleRows(world).find((r) => r[0].startsWith(taskPrefix))[0];

test('setup: creates 2 cycles per task (current + next), rerunning adds none (T1/T8)', () => {
  const { api, world } = boot();
  const n0 = cycleRows(world).length;
  assert.equal(n0, 8, '4 tasks x (current + next) cycles');
  api.ensureCycles();
  api.ensureCycles();
  assert.equal(cycleRows(world).length, n0, 'rerun must not duplicate cycles');
});

test('cycle keys are unique and due = start + freq - 1 (T1)', () => {
  const { world } = boot();
  const rows = cycleRows(world);
  assert.equal(new Set(rows.map((r) => r[0])).size, rows.length);
  const t2 = rows.filter((r) => r[0].startsWith('T-002|')).sort((a, b) => a[6] - b[6]);
  assert.equal(t2[1][6] - t2[0][6], 7 * 86400000, 'next 7-day cycle starts +7d');
  assert.equal(t2[0][7] - t2[0][6], 6 * 86400000, 'due = start + 7 - 1');
});

test('ticket task rejects Done without proof, then completes once (T2)', () => {
  const { api, world } = boot();
  const key = keyOf(world, 'T-002|');
  assert.equal(api.completeCycle(key, '', 'Employee B'), false);
  assert.ok(cycleRows(world).find((r) => r[0] === key)[8] !== 'Completed');
  assert.equal(api.completeCycle(key, 'TKT-1001', 'Employee B'), true);
  const row = cycleRows(world).find((r) => r[0] === key);
  assert.equal(row[8], 'Completed');
  assert.equal(row[9], 'Employee B');
  assert.ok(row[10] instanceof Date);
  assert.equal(row[11], 'TKT-1001');
});

test('second completion of same cycle is refused — no duplicate entries (T8)', () => {
  const { api, world } = boot();
  const key = keyOf(world, 'T-003|');
  assert.equal(api.completeCycle(key, 'TKT-2', 'Employee C'), true);
  assert.equal(api.completeCycle(key, 'TKT-3', 'Employee C'), false);
  assert.equal(logKeys(world).filter((k) => k === `COMPLETE|${key}`).length, 1);
});

test('daily no-ticket task completes with self-attestation (T3)', () => {
  const { api, world } = boot();
  const key = keyOf(world, 'T-001|');
  assert.equal(api.completeCycle(key, '', 'Employee A'), true);
  assert.match(String(cycleRows(world).find((r) => r[0] === key)[11]), /Self-attested/);
});

test('same-day manager alert: exactly one email, stamped once, deduped on rerun (T4/T8)', () => {
  const { api, world } = boot();
  setConfig(world, 'MANAGER_EMAIL', 'me+manager@gmail.com');
  api.demoDueToday();
  api.runDailyChecks();
  const mgr = world.mail.filter((m) => /OVERDUE/.test(m.subject));
  assert.equal(mgr.length, 1);
  assert.match(mgr[0].to, /manager/i);
  api.runDailyChecks();
  api.runDailyChecks();
  assert.equal(world.mail.filter((m) => /OVERDUE/.test(m.subject)).length, 1, 'reruns send no second alert');
});

test('leadership escalation fires once when alert goes unacknowledged (T6/T8)', () => {
  const { api, world } = boot();
  setConfig(world, 'MANAGER_EMAIL', 'me+manager@gmail.com');
  setConfig(world, 'LEADERSHIP_EMAIL', 'me+leadership@gmail.com');
  api.demoForceEscalation();
  api.runDailyChecks();
  const esc = world.mail.filter((m) => /ESCALATION/.test(m.subject));
  assert.equal(esc.length, 1);
  assert.match(esc[0].to, /leadership/i);
  api.runDailyChecks();
  assert.equal(world.mail.filter((m) => /ESCALATION/.test(m.subject)).length, 1);
});

test('manager acknowledgement stamped once (T5)', () => {
  const { api, world } = boot();
  const key = keyOf(world, 'T-001|');
  api.acknowledge(key, 'Manager');
  api.acknowledge(key, 'Manager');
  const row = cycleRows(world).find((r) => r[0] === key);
  assert.equal(row[14], 'Manager');
  assert.ok(row[15] instanceof Date);
  assert.equal(logKeys(world).filter((k) => k === `ACK|${key}`).length, 1);
});

test('single-point update: Done tick on MyTasks completes cycle (T2)', () => {
  const { api, world } = boot();
  api.renderMyTasks();
  const my = world.sheets.get('MyTasks');
  const grid = my.getDataRange().getValues();
  const r = grid.findIndex((row) => String(row[0]).startsWith('T-001|')) + 1;
  assert.ok(r > 0, 'daily task rendered in MyTasks');
  my.getRange(r, 6).setValue('Posted 9:15am');
  api.onEdit({ range: my.getRange(r, 7), value: 'TRUE' });
  const key = grid[r - 1][0];
  assert.equal(cycleRows(world).find((x) => x[0] === key)[8], 'Completed');
});

test('placeholder emails are never sent — logged as EMAIL_SKIPPED (privacy)', () => {
  const { api, world } = boot();
  api.demoDueToday();
  api.runDailyChecks(); // Config still has example.com placeholders
  assert.equal(world.mail.length, 0, 'no real email to example.com');
  assert.ok(logKeys(world).some((k) => k.startsWith('EMAIL_SKIPPED|')));
});

test('real test emails DO send when Config holds them', () => {
  const { api, world } = boot();
  setConfig(world, 'MANAGER_EMAIL', 'me+mgr@gmail.com');
  api.demoDueToday();
  api.runDailyChecks();
  assert.equal(world.mail[0].to, 'me+mgr@gmail.com');
});

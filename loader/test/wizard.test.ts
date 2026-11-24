import assert from 'node:assert/strict';
import { STEPS, canVisit, firstBlocked, Ctx } from '../src/wizard';

const ctx = (o: Partial<Ctx> = {}): Ctx => ({ kinds: new Set(), hasWorkpaper: false, gstin: '', nFiles: 0, ...o });

// every tab starts with Add files, which can always be visited
for (const t of Object.keys(STEPS) as (keyof typeof STEPS)[]) {
  assert.equal(STEPS[t][0].id, 'files', t);
  assert.ok(canVisit(STEPS[t], 0, ctx()), t);
}
// nothing loaded: Next is blocked on step 1 with a plain-language reason
const itc = STEPS.itc;
assert.equal(firstBlocked(itc, ctx()), 0);
assert.match(itc[0].missing!(ctx())!, /GSTR-2B/);
assert.ok(!canVisit(itc, 2, ctx()));
// 2B + register: step 2 reachable, but needs a GSTIN
const have = ctx({ kinds: new Set(['GSTR-2B', 'Register']), nFiles: 2 });
assert.equal(itc[0].missing!(have), null);
assert.ok(canVisit(itc, 1, have));
assert.equal(firstBlocked(itc, have), 1);
assert.ok(!canVisit(itc, 2, have));
assert.ok(canVisit(itc, 4, ctx({ ...have, gstin: '32AAACJ1234K1Z5' })));
// a saved workpaper carries its own files
assert.equal(itc[0].missing!(ctx({ hasWorkpaper: true, gstin: 'X', nFiles: 1 })), null);
// flow needs both portal and books ledgers
assert.match(STEPS.flow[0].missing!(ctx({ kinds: new Set(['Portal ledger']) }))!, /Books ledger/);
console.log('wizard ok');

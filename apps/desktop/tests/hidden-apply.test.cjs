const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadScheduler() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/hidden-apply.ts'),
    platform: 'node',
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    new Function('require', 'module', 'exports', output[0].code)(require, mod, mod.exports);
    return mod.exports;
  } finally {
    await bundle.close();
  }
}

test('hidden apply waits for the debounce, skips IME composition, and flushes pending work', async t => {
  const { createHiddenApplyScheduler } = await loadScheduler();
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const calls = [];
  let composing = false;
  const scheduler = createHiddenApplyScheduler({
    delayMs: 80,
    composing: () => composing,
    apply: () => { calls.push('apply'); },
  });
  scheduler.schedule();
  scheduler.schedule();
  assert.deepEqual(calls, []);
  t.mock.timers.tick(79);
  assert.deepEqual(calls, []);
  t.mock.timers.tick(1);
  assert.deepEqual(calls, ['apply']);
  await Promise.resolve();
  composing = true;
  scheduler.schedule();
  t.mock.timers.tick(80);
  assert.deepEqual(calls, ['apply']);
  composing = false;
  scheduler.schedule();
  t.mock.timers.tick(80);
  assert.deepEqual(calls, ['apply', 'apply']);
  await Promise.resolve();
  scheduler.schedule();
  scheduler.cancel();
  t.mock.timers.tick(80);
  assert.deepEqual(calls, ['apply', 'apply']);
  scheduler.schedule();
  await scheduler.flush();
  assert.deepEqual(calls, ['apply', 'apply', 'apply']);
});

test('hidden apply flush waits for an in-flight async apply', async t => {
  const { createHiddenApplyScheduler } = await loadScheduler();
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const calls = [];
  let finish;
  const scheduler = createHiddenApplyScheduler({
    composing: () => false,
    apply: () => new Promise(resolve => { calls.push('start'); finish = resolve; }),
  });
  scheduler.schedule();
  const flushed = scheduler.flush();
  let done = false;
  flushed.then(() => { done = true; });
  await Promise.resolve();
  assert.deepEqual(calls, ['start']);
  assert.equal(done, false);
  finish();
  await flushed;
  assert.equal(done, true);
});

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

test('hidden apply waits for the debounce, skips IME composition, and serializes in-flight work', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const { createHiddenApplyScheduler } = await loadScheduler();
  const calls = [];
  let composing = false;
  let release;
  const scheduler = createHiddenApplyScheduler({
    delayMs: 80,
    composing: () => composing,
    apply: () => {
      calls.push('apply');
      return new Promise(resolve => { release = resolve; });
    },
  });
  scheduler.schedule();
  scheduler.schedule();
  assert.deepEqual(calls, []);
  t.mock.timers.tick(79);
  assert.deepEqual(calls, []);
  t.mock.timers.tick(1);
  assert.deepEqual(calls, ['apply']);
  composing = true;
  scheduler.schedule();
  t.mock.timers.tick(80);
  assert.deepEqual(calls, ['apply']);
  composing = false;
  release();
  await Promise.resolve();
  await Promise.resolve();
  t.mock.timers.tick(80);
  assert.deepEqual(calls, ['apply', 'apply']);
  scheduler.cancel();
  scheduler.schedule();
  scheduler.cancel();
  t.mock.timers.tick(80);
  assert.deepEqual(calls, ['apply', 'apply']);
  scheduler.schedule();
  await scheduler.flush();
  assert.deepEqual(calls, ['apply', 'apply', 'apply']);
  release();
});

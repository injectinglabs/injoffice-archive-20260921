const { test } = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;
async function loadDialog(name = 'default') {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({ input: path.resolve(__dirname, '../src/UpdatesDialog.tsx'), platform: 'node', external: id => /^react(?:\/|$)/.test(id), transform: { jsx: { runtime: 'automatic' } }, plugins: [{ name: 'assets', resolveId(id) { if (/\.(css|png)$/.test(id)) return '\0asset'; }, load(id) { if (id === '\0asset') return 'export default ""'; } }] });
  try { const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false }); const mod = { exports: {} }; new Function('require', 'module', 'exports', output[0].code)(require, mod, mod.exports); return mod.exports[name]; } finally { await bundle.close(); }
}
const initial = { status: 'idle', appVersion: '0.1.0', autoCheck: true };
function deferred() { let resolve; const promise = new Promise(r => resolve = r); return { promise, resolve }; }
test('update events win over delayed hydration and action responses; subscriptions clean up', async () => {
  const Dialog = await loadDialog(); const hydration = deferred(); const check = deferred(); let event, unsubscribed = 0, calls = 0;
  global.window = { injDesktop: { getUpdateState: () => hydration.promise, onUpdateState: fn => { event = fn; return () => unsubscribed++; }, checkForUpdates: () => { calls++; return check.promise; } } };
  let view;
  try {
    await act(async () => { view = create(React.createElement(Dialog, { onClose() {} })); });
    await act(async () => event({ ...initial, status: 'available', version: '0.2.0', releaseNotes: '<script>plain text</script>' }));
    await act(async () => hydration.resolve(initial));
    assert.equal(view.root.findAllByType('h3')[0].children.join(''), 'A new version is available');
    assert.equal(view.root.findAllByType('script').length, 0);
    await act(async () => event(initial));
    await act(async () => view.root.findAllByType('button').find(b => b.children.join('') === 'Check for updates').props.onClick());
    assert.equal(calls, 1);
    await act(async () => event({ ...initial, status: 'downloaded', version: '0.2.0' }));
    await act(async () => check.resolve({ ...initial, status: 'checking' }));
    assert.equal(view.root.findAllByType('h3')[0].children.join(''), 'Ready to restart');
    await act(async () => view.unmount()); view = null;
    assert.equal(unsubscribed, 1);
  } finally { if (view) await act(async () => view.unmount()); delete global.window; }
});
test('download, restart refusal, and automatic preference are explicit user actions', async () => {
  const Dialog = await loadDialog(); let downloads = 0, installs = 0, preference;
  global.window = { injDesktop: { getUpdateState: async () => ({ ...initial, status: 'available', version: '0.2.0' }), onUpdateState: () => () => {}, downloadUpdate: async () => { downloads++; return { ...initial, status: 'downloaded' }; }, installUpdate: async () => { installs++; return { ...initial, status: 'downloaded', message: 'Save your documents before restarting.' }; }, setAutomaticUpdates: async value => { preference = value; return { ...initial, status: 'downloaded', autoCheck: value }; } } };
  let view;
  const click = async label => act(async () => view.root.findAllByType('button').find(b => b.children.join('') === label).props.onClick());
  try {
    await act(async () => { view = create(React.createElement(Dialog, { onClose() {} })); });
    assert.equal(downloads, 0); await click('Download update'); assert.equal(downloads, 1);
    assert.equal(installs, 0); await click('Restart and update'); assert.equal(installs, 1);
    assert.ok(JSON.stringify(view.toJSON()).includes('Save your documents before restarting.'));
    await act(async () => view.root.findByType('input').props.onChange({ target: { checked: false } }));
    assert.equal(preference, false); assert.equal(view.root.findByType('input').props.checked, false);
  } finally { if (view) await act(async () => view.unmount()); delete global.window; }
});

test('background update notice is dismissible without downloading or forcing a dialog', async () => {
  const Notice = await loadDialog('UpdateNotice'); let event, opened = 0, unsubscribed = 0;
  global.window = { injDesktop: { getUpdateState: async () => initial, onUpdateState: fn => { event = fn; return () => unsubscribed++; } } };
  let view;
  try {
    await act(async () => { view = create(React.createElement(Notice, { onOpen() { opened++; } })); });
    assert.equal(view.toJSON(), null);
    await act(async () => event({ ...initial, status: 'available', version: '0.2.0' }));
    assert.equal(opened, 0);
    await act(async () => view.root.findAllByType('button')[0].props.onClick());
    assert.equal(opened, 1);
    await act(async () => view.root.findAllByType('button')[1].props.onClick());
    assert.equal(view.toJSON(), null);
    await act(async () => event({ ...initial, status: 'downloaded', version: '0.2.0' }));
    assert.ok(JSON.stringify(view.toJSON()).includes('update is ready'));
    await act(async () => view.unmount()); view = null;
    assert.equal(unsubscribed, 1);
  } finally { if (view) await act(async () => view.unmount()); delete global.window; }
});

// Opening the dialog used to paint a heavy focus ring on × (autoFocus), and a build without an
// updater said "Updates unavailable" where Office says what it means.
test('the dialog takes its own initial focus and says what an updateless build can do', async () => {
  const Dialog = await loadDialog();
  const focused = [];
  global.window = {};
  let view;
  try {
    await act(async () => { view = create(React.createElement(Dialog, { onClose() {} }), { createNodeMock: element => element.type === 'dialog' ? { showModal() {}, close() {}, focus: options => focused.push(options) } : null }); });
    const close = view.root.findByProps({ 'aria-label': 'Close updates' });
    assert.equal(close.props.autoFocus, undefined, 'the close button is not autofocused');
    assert.deepEqual(focused, [{ preventScroll: true }], 'the dialog itself takes the initial focus');
    assert.equal(view.root.findByType('dialog').props.tabIndex, -1);
    // Without a host bridge the dialog reports the local build honestly.
    assert.match(view.root.findByProps({ className: 'updates-summary' }).findByType('h3').children.join(''), /Updates aren’t available in this build/);
  } finally { if (view) await act(async () => view.unmount()); delete global.window; }
});

test('a checked build with nothing to install says it is up to date', async () => {
  const Dialog = await loadDialog();
  global.window = { injDesktop: { getUpdateState: async () => ({ status: 'not-available', appVersion: '0.1.0', autoCheck: true }), onUpdateState: () => () => {} } };
  let view;
  try {
    await act(async () => { view = create(React.createElement(Dialog, { onClose() {} }), { createNodeMock: element => element.type === 'dialog' ? { showModal() {}, close() {}, focus() {} } : null }); });
    assert.match(view.root.findByProps({ className: 'updates-summary' }).findByType('h3').children.join(''), /You’re up to date/);
  } finally { if (view) await act(async () => view.unmount()); delete global.window; }
});

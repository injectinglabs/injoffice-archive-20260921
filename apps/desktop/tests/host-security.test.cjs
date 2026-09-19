const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const os = require('node:os');
const { pathToFileURL } = require('node:url');
const { allowRequest, resolveAsset, isTrustedSender } = require('../electron/security.cjs');
const root = path.join(os.tmpdir(), 'injoffice-renderer');

test('protocol serves bundled assets and permits local WASM fetches', () => {
  assert.equal(resolveAsset(root, 'injoffice://app/assets/engine.wasm'), path.join(root, 'assets/engine.wasm'));
  assert.equal(allowRequest(root, 'injoffice://app/assets/worker.js'), true);
  assert.equal(allowRequest(root, pathToFileURL(path.join(root, 'assets/engine.wasm')).href), true);
  assert.equal(allowRequest(root, 'data:image/png;base64,AA=='), true);
  assert.equal(allowRequest(root, 'blob:injoffice://app/1234'), true);
});

test('deny internet, loopback, other application hosts and unrelated local files', () => {
  for (const url of ['https://example.com', 'http://127.0.0.1:3000', 'ws://localhost:8080', 'ftp://example.com/a', 'injoffice://other/index.html', 'injoffice://user@app/index.html', 'injoffice://app:8080/index.html', pathToFileURL(path.join(os.tmpdir(), 'private.txt')).href]) {
    assert.equal(allowRequest(root, url), false, url);
  }
});

test('deny encoded traversal, malformed paths and prefix-matching siblings', () => {
  for (const url of ['injoffice://app/..%2fprivate.txt', 'injoffice://app/%2e%2e%2fprivate.txt', 'injoffice://app/%5c..%5cprivate.txt', 'injoffice://app/%00', 'injoffice://app/%ZZ']) {
    assert.throws(() => resolveAsset(root, url), undefined, url);
    assert.equal(allowRequest(root, url), false, url);
  }
  assert.equal(allowRequest(root, pathToFileURL(`${root}-other/private.txt`).href), false);
});

test('IPC accepts only the live application main frame and rejects missing frames without throwing', () => {
  const entryURL = 'injoffice://app/index.html';
  const mainFrame = { url: entryURL };
  const webContents = { mainFrame };
  const window = { webContents, isDestroyed: () => false };
  assert.equal(isTrustedSender({ sender: webContents, senderFrame: mainFrame }, window, entryURL), true);
  for (const event of [{ sender: webContents }, { sender: webContents, senderFrame: { url: entryURL } }, { sender: {}, senderFrame: mainFrame }]) {
    assert.equal(isTrustedSender(event, window, entryURL), false);
  }
  assert.equal(isTrustedSender({ sender: webContents, senderFrame: mainFrame }, undefined, entryURL), false);
  mainFrame.url = 'https://example.com';
  assert.equal(isTrustedSender({ sender: webContents, senderFrame: mainFrame }, window, entryURL), false);
});

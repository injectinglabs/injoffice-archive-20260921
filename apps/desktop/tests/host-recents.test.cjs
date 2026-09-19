const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { RecentFiles, openRecentFile } = require('../electron/recent-files.cjs');
const { FileStore } = require('../electron/file-store.cjs');

async function fixture(t) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'injoffice-recents-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const historyPath = path.join(directory, 'profile', 'recent-files.json');
  const filename = path.join(directory, 'document.docx');
  await fs.writeFile(filename, 'document bytes');
  return { directory, historyPath, filename, history: new RecentFiles(historyPath), store: new FileStore() };
}

test('recent files persist across restart and reopen only through their stored opaque id', async (t) => {
  const { history, historyPath, filename, store } = await fixture(t);
  const recorded = await history.record(filename);
  assert.notEqual(recorded.id, filename);
  assert.equal(recorded.path, await fs.realpath(filename));
  assert.equal(recorded.name, 'document.docx');
  assert.ok(recorded.updatedAt > 0);
  const restarted = new RecentFiles(historyPath);
  assert.deepEqual(await restarted.list(), [recorded]);
  const opened = await openRecentFile(restarted, store, recorded.id);
  assert.equal(opened.bytes.toString(), 'document bytes');
  assert.notEqual(opened.id, recorded.id);
  await assert.rejects(openRecentFile(restarted, store, filename), /no longer in your history/);
  await assert.rejects(openRecentFile(restarted, store, { path: filename }), /no longer in your history/);
});

test('history serializes concurrent updates, deduplicates paths, and keeps only the newest thirty', async (t) => {
  const { directory, history, historyPath } = await fixture(t);
  const files = Array.from({ length: 35 }, (_, index) => path.join(directory, `${index}.xlsx`));
  await Promise.all(files.map(filename => fs.writeFile(filename, 'xlsx bytes')));
  await Promise.all(files.map(filename => history.record(filename)));
  const listed = await history.list();
  assert.equal(listed.length, 30);
  assert.deepEqual(listed.map(entry => entry.name), files.slice(5).reverse().map(filename => path.basename(filename)));
  const last = listed.at(-1);
  const updated = await history.record(last.path);
  assert.equal(updated.id, last.id);
  assert.equal((await history.list())[0].id, last.id);
  assert.equal((await new RecentFiles(historyPath).list()).length, 30);
  listed[0].name = 'renderer mutation';
  assert.notEqual((await history.list())[1].name, 'renderer mutation');
});

test('Save As records the new copy independently; subsequent saves reuse its recent identifier', async (t) => {
  const { directory, filename, history, store } = await fixture(t);
  const opened = await store.open(filename);
  const original = await history.record(store.get(opened.id).filename);
  const copy = path.join(directory, 'copy.docx');
  await store.save(opened.id, Buffer.from('saved copy'), copy);
  const copied = await history.record(store.get(opened.id).filename);
  assert.notEqual(copied.id, original.id);
  await store.save(opened.id, Buffer.from('saved again'));
  assert.equal((await history.record(store.get(opened.id).filename)).id, copied.id);
  assert.deepEqual((await history.list()).map(entry => entry.name), ['copy.docx', 'document.docx']);
  assert.equal(await fs.readFile(filename, 'utf8'), 'document bytes');
});

test('missing recent files give actionable feedback and removal persists without deleting the document', async (t) => {
  const { history, historyPath, filename, store } = await fixture(t);
  const entry = await history.record(filename);
  await fs.unlink(filename);
  await assert.rejects(openRecentFile(history, store, entry.id), /Open file to locate it, or remove it from Recent files/);
  assert.equal((await history.list()).length, 1);
  await fs.writeFile(filename, 'restored');
  await history.remove(entry.id);
  assert.deepEqual(await new RecentFiles(historyPath).list(), []);
  assert.equal(await fs.readFile(filename, 'utf8'), 'restored');
  await assert.rejects(openRecentFile(history, store, entry.id), /no longer in your history/);
});

test('corrupt history is recoverable and malformed persisted entries do not become recent capabilities', async (t) => {
  const { historyPath, filename } = await fixture(t);
  await fs.mkdir(path.dirname(historyPath), { recursive: true });
  await fs.writeFile(historyPath, '{broken');
  const recovered = new RecentFiles(historyPath);
  assert.deepEqual(await recovered.list(), []);
  const entry = await recovered.record(filename);
  await fs.writeFile(historyPath, JSON.stringify({ version: 1, files: [
    { ...entry, path: '../relative.txt' }, { ...entry, id: '../arbitrary-path' },
    { ...entry, updatedAt: 'yesterday' }, { ...entry, name: 'forged name' }, entry,
  ] }));
  assert.deepEqual(await new RecentFiles(historyPath).list(), [entry]);
});

test('failed history persistence does not poison future history updates', async (t) => {
  const { directory, filename } = await fixture(t);
  const blocker = path.join(directory, 'not-a-directory');
  await fs.writeFile(blocker, 'block');
  const history = new RecentFiles(path.join(blocker, 'recent-files.json'));
  await assert.rejects(history.record(filename));
  await fs.unlink(blocker);
  const saved = await history.record(filename);
  assert.deepEqual(await history.list(), [saved]);
});

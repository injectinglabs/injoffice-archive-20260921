const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { FileStore, atomicWrite, normalizeSaveDestination, validateFormat, validateOpenFormat } = require('../electron/file-store.cjs');

async function fixture(t) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'injoffice-host-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const filename = path.join(directory, 'example.xlsx');
  await fs.writeFile(filename, 'original');
  return { directory, filename, store: new FileStore() };
}

test('open exposes opaque identifiers and save replaces the file without leftover temporary files', async (t) => {
  const { directory, filename, store } = await fixture(t);
  const opened = await store.open(filename);
  assert.equal(opened.name, 'example.xlsx');
  assert.equal(opened.bytes.toString(), 'original');
  assert.deepEqual(Object.keys(opened).sort(), ['bytes', 'id', 'name']);
  assert.notEqual(opened.id, filename);
  const saved = await store.save(opened.id, Buffer.from('edited'));
  assert.deepEqual(saved, { id: opened.id, name: opened.name, untitled: false });
  assert.equal(await fs.readFile(filename, 'utf8'), 'edited');
  assert.equal(await store.changed(opened.id), false);
  assert.deepEqual(await fs.readdir(directory), ['example.xlsx']);
});

test('external changes cannot be silently overwritten, even with matching timestamps', async (t) => {
  const { filename, store } = await fixture(t);
  const opened = await store.open(filename);
  const original = await fs.stat(filename);
  await fs.writeFile(filename, 'external');
  await fs.utimes(filename, original.atime, original.mtime);
  await assert.rejects(store.save(opened.id, Buffer.from('mine')), { code: 'EXTERNAL_CHANGE' });
  assert.equal(await fs.readFile(filename, 'utf8'), 'external');
  await store.save(opened.id, Buffer.from('confirmed'), undefined, true);
  assert.equal(await fs.readFile(filename, 'utf8'), 'confirmed');
});

test('Save As leaves the source untouched and future saves follow the new destination', async (t) => {
  const { directory, filename, store } = await fixture(t);
  const opened = await store.open(filename);
  const copy = path.join(directory, 'copy.xlsx');
  assert.equal((await store.save(opened.id, Buffer.from('copy'), copy)).name, 'copy.xlsx');
  await store.save(opened.id, Buffer.from('second edit'));
  assert.equal(await fs.readFile(filename, 'utf8'), 'original');
  assert.equal(await fs.readFile(copy, 'utf8'), 'second edit');
});

test('deleted originals require an explicit overwrite decision', async (t) => {
  const { filename, store } = await fixture(t);
  const opened = await store.open(filename);
  await fs.unlink(filename);
  await assert.rejects(store.save(opened.id, Buffer.from('mine')), { code: 'EXTERNAL_CHANGE' });
});

test('reject arbitrary paths masquerading as ids and non-binary save input', async (t) => {
  const { filename, store } = await fixture(t);
  const opened = await store.open(filename);
  await assert.rejects(store.save(filename, Buffer.from('bad')), /Unknown document/);
  await assert.rejects(store.save(opened.id, 'bad'), /Expected document bytes/);
  assert.equal(await fs.readFile(filename, 'utf8'), 'original');
});

test('atomic replacement refuses symbolic links and preserves file permissions', async (t) => {
  const { directory, filename } = await fixture(t);
  if (process.platform !== 'win32') {
    await fs.chmod(filename, 0o640);
    await atomicWrite(filename, Buffer.from('saved'));
    assert.equal((await fs.stat(filename)).mode & 0o777, 0o640);
    const link = path.join(directory, 'link.xlsx');
    await fs.symlink(filename, link);
    await assert.rejects(atomicWrite(link, Buffer.from('bad')), /regular file/);
    assert.equal(await fs.readFile(filename, 'utf8'), 'saved');
  }
});

test('macro-enabled Word is openable and saveable as .docm but is not a blank-create format', () => {
  assert.equal(validateOpenFormat('docm'), 'docm');
  assert.equal(validateOpenFormat('DOCM'), 'docm');
  assert.throws(() => validateOpenFormat('doc'), /Choose a DOCX/);
  assert.throws(() => validateFormat('docm'), /Choose a DOCX/);
  assert.equal(normalizeSaveDestination('/tmp/macro.docm', 'docm'), '/tmp/macro.docm');
});

test('new Office sessions have opaque ids and independent in-memory bytes without creating files', async (t) => {
  const { directory, store } = await fixture(t);
  const before = await fs.readdir(directory);
  for (const [format, name] of [['docx', 'Untitled.docx'], ['xlsx', 'Workbook.xlsx'], ['pptx', 'Presentation.pptx'], ['pdf', 'Untitled.pdf']]) {
    const bytes = Buffer.from('blank document');
    const created = store.create(format, bytes);
    assert.equal(created.name, name);
    assert.equal(created.untitled, true);
    assert.equal(store.get(created.id).filename, undefined);
    assert.equal(await store.changed(created.id), false);
    assert.notEqual(store.create(format, bytes).id, created.id);
    bytes.fill(0);
    created.bytes.fill(0);
    assert.equal(store.get(created.id).bytes.toString(), 'blank document');
  }
  assert.deepEqual(await fs.readdir(directory), before);
});

test('an untitled session remains usable until a destination is chosen and first save makes it a disk session', async (t) => {
  const { directory, store } = await fixture(t);
  const created = store.create('docx', Buffer.from('blank'));
  // A canceled native Save dialog never calls save; neither the session nor
  // its bytes are consumed. A missing destination is also rejected defensively.
  const initial = store.get(created.id);
  await assert.rejects(store.save(created.id, Buffer.from('draft')), /Choose where to save/);
  assert.equal(store.get(created.id), initial);
  assert.equal(store.get(created.id).untitled, true);
  const destination = path.join(directory, 'my-document.docx');
  assert.deepEqual(await store.save(created.id, Buffer.from('draft'), destination), { id: created.id, name: 'my-document.docx', untitled: false });
  assert.equal(store.get(created.id).bytes, undefined);
  assert.equal(await fs.readFile(destination, 'utf8'), 'draft');
  await store.save(created.id, Buffer.from('second draft'));
  assert.equal((await store.open(destination)).bytes.toString(), 'second draft');
  await fs.writeFile(destination, 'external modification');
  await assert.rejects(store.save(created.id, Buffer.from('third draft')), { code: 'EXTERNAL_CHANGE' });
});

test('failed first-save writes preserve the untitled session for retry', async (t) => {
  const { directory, store } = await fixture(t);
  const created = store.create('xlsx', Buffer.from('blank workbook'));
  await assert.rejects(store.save(created.id, Buffer.from('edited workbook'), path.join(directory, 'missing', 'book.xlsx')), { code: 'ENOENT' });
  assert.equal(store.get(created.id).untitled, true);
  assert.equal(store.get(created.id).filename, undefined);
  const destination = path.join(directory, 'book.xlsx');
  await store.save(created.id, Buffer.from('edited workbook'), destination);
  assert.equal(await fs.readFile(destination, 'utf8'), 'edited workbook');
});

test('creation rejects unsupported formats and arbitrary paths; unsaved ids remain required for saving', async (t) => {
  const { filename, store } = await fixture(t);
  for (const format of ['txt', '../document.docx', filename, 'toString', null, {}]) {
    assert.throws(() => store.create(format, Buffer.from('blank')), /Choose a DOCX/);
  }
  assert.throws(() => store.create('docx', 'not bytes'), /Expected document bytes/);
  const created = store.create('docx', Buffer.from('blank'));
  await assert.rejects(store.save(created.name, Buffer.from('draft'), filename), /Unknown document/);
  await assert.rejects(store.save(created.id, Buffer.from('draft'), '../relative.docx'), /Choose where to save/);
  assert.equal(await fs.readFile(filename, 'utf8'), 'original');
});

test('save destinations add missing extensions, preserve matching names, and reject format conversion by renaming', () => {
  const base = path.join(os.tmpdir(), 'My document');
  for (const format of ['docx', 'xlsx', 'pptx', 'pdf']) {
    assert.equal(normalizeSaveDestination(base, format), `${base}.${format}`);
    assert.equal(normalizeSaveDestination(`${base}.`, format), `${base}.${format}`);
    assert.equal(normalizeSaveDestination(`${base}.${format}`, format), `${base}.${format}`);
    assert.equal(normalizeSaveDestination(`${base}.${format.toUpperCase()}`, format), `${base}.${format.toUpperCase()}`);
  }
  assert.throws(() => normalizeSaveDestination(`${base}.xlsx`, 'docx'), /\.docx extension.*does not convert it to XLSX/);
  assert.throws(() => normalizeSaveDestination(`${base}.pdf`, 'pptx'), /\.pptx extension.*does not convert it to PDF/);
  assert.throws(() => normalizeSaveDestination(`${base}.zip`, 'xlsx'), /\.xlsx extension/);
  assert.throws(() => normalizeSaveDestination('relative.docx', 'docx'), /valid location/);
});

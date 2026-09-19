const test = require('node:test');
const assert = require('node:assert/strict');
const { createBlankDocument } = require('../electron/new-document.cjs');
const { FileStore } = require('../electron/file-store.cjs');

function zipNames(bytes) {
  const names = [];
  let offset = 0;
  while (offset + 30 <= bytes.length && bytes.readUInt32LE(offset) === 0x04034b50) {
    const nameLen = bytes.readUInt16LE(offset + 26);
    const extraLen = bytes.readUInt16LE(offset + 28);
    const size = bytes.readUInt32LE(offset + 18);
    names.push(bytes.subarray(offset + 30, offset + 30 + nameLen).toString());
    offset += 30 + nameLen + extraLen + size;
  }
  return names;
}

test('blank factory rejects unsupported formats', async () => {
  await assert.rejects(createBlankDocument('txt'), /Choose/);
  await assert.rejects(createBlankDocument('../docx'), /Choose/);
  await assert.rejects(createBlankDocument('docm'), /Choose/);
});

test('blank Office packages are stored ZIPs with the required parts and no user text', async () => {
  const expected = {
    docx: ['[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/styles.xml', 'word/numbering.xml', 'word/settings.xml'],
    xlsx: ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/styles.xml'],
    pptx: ['[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml', 'ppt/slides/slide1.xml', 'ppt/slideLayouts/slideLayout1.xml', 'ppt/slideMasters/slideMaster1.xml', 'ppt/theme/theme1.xml'],
  };
  for (const [format, parts] of Object.entries(expected)) {
    const bytes = await createBlankDocument(format);
    assert.equal(bytes.readUInt32LE(0), 0x04034b50);
    const names = zipNames(bytes);
    for (const part of parts) assert.ok(names.includes(part), `${format} missing ${part}`);
    assert.equal(bytes.includes(Buffer.from('Lorem')), false);
  }
});

test('blank PDF is a one-page PDF-1.7 file', async () => {
  const bytes = await createBlankDocument('pdf');
  const text = bytes.toString('latin1');
  assert.ok(text.startsWith('%PDF-1.7\n'));
  assert.ok(text.includes('/MediaBox [0 0 612 792]'));
  assert.ok(text.endsWith('%%EOF\n'));
});

test('blank seeds open as untitled FileStore sessions without writing disk', async () => {
  const store = new FileStore();
  const created = store.create('xlsx', await createBlankDocument('xlsx'));
  assert.equal(created.untitled, true);
  assert.equal(created.name, 'Workbook.xlsx');
  assert.equal(created.bytes.readUInt32LE(0), 0x04034b50);
});

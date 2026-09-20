const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const icons = path.resolve(__dirname, '../build/icons');
const pngSignature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

function readPngSize(bytes) {
  assert.equal(bytes.subarray(0, 8).equals(pngSignature), true);
  assert.equal(bytes.subarray(12, 16).toString(), 'IHDR');
  return { width: bytes.readUInt32BE(16), height: bytes.readUInt32BE(20) };
}

function icoSizes(bytes) {
  assert.equal(bytes.readUInt16LE(0), 0);
  assert.equal(bytes.readUInt16LE(2), 1);
  const count = bytes.readUInt16LE(4);
  const sizes = [];
  for (let i = 0; i < count; i++) {
    const offset = 6 + i * 16;
    sizes.push(bytes[offset] || 256);
  }
  return sizes;
}

function icnsTypes(bytes) {
  assert.equal(bytes.subarray(0, 4).toString(), 'icns');
  assert.equal(bytes.readUInt32BE(4), bytes.length);
  const types = [];
  for (let offset = 8; offset + 8 <= bytes.length; ) {
    const size = bytes.readUInt32BE(offset + 4);
    assert.ok(size >= 8 && offset + size <= bytes.length, `icns ${bytes.subarray(offset, offset + 4)} length`);
    types.push(bytes.subarray(offset, offset + 4).toString());
    offset += size;
  }
  return types;
}

test('desktop icons are nearest-neighbor conversions of the repository logo', () => {
  const png = fs.readFileSync(path.join(icons, 'icon.png'));
  const ico = fs.readFileSync(path.join(icons, 'icon.ico'));
  const icns = fs.readFileSync(path.join(icons, 'icon.icns'));
  assert.deepEqual(readPngSize(png), { width: 1024, height: 1024 });
  assert.deepEqual(icoSizes(ico), [16, 24, 32, 48, 64, 128, 256]);
  const types = icnsTypes(icns);
  for (const type of ['ic04', 'ic05', 'ic07', 'ic08', 'ic09', 'ic10', 'ic11', 'ic12', 'ic13', 'ic14']) {
    assert.ok(types.includes(type), type);
  }
  const generate = fs.readFileSync(path.join(icons, 'generate.py'), 'utf8');
  assert.match(generate, /parents\[3\] \/ "logo\.png"/);
  assert.match(generate, /Image\.Resampling\.NEAREST/);
  assert.equal(fs.existsSync(path.resolve(__dirname, '../../../logo.png')), true);
});

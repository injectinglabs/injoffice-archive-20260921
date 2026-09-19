const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadImage() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/raster-image.ts'),
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

function png(width, height) {
  const bytes = Buffer.alloc(24);
  Buffer.from([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]).copy(bytes);
  bytes.writeUInt32BE(width, 16);
  bytes.writeUInt32BE(height, 20);
  return bytes;
}

test('imageDimensions reads PNG and JPEG headers and rejects empty or huge rasters', async () => {
  const { imageDimensions } = await loadImage();
  assert.deepEqual(imageDimensions(png(8, 4)), { width: 8, height: 4, contentType: 'image/png' });
  const jpeg = Buffer.from([0xff, 0xd8, 0xff, 0xc0, 0x00, 0x08, 0x08, 0x00, 0x10, 0x00, 0x20, 0x00]);
  assert.deepEqual(imageDimensions(jpeg), { width: 32, height: 16, contentType: 'image/jpeg' });
  assert.throws(() => imageDimensions(Buffer.alloc(8)), /valid PNG\/JPEG/);
  assert.throws(() => imageDimensions(png(4001, 4001)), /16 megapixels/);
  assert.throws(() => imageDimensions(Buffer.alloc(2 * 1024 * 1024 + 1)), /2 MiB/);
});

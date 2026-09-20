const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadMedia() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/document-media.ts'),
    platform: 'node',
    external: id => id.includes('packages/docs/src/'),
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

function drawing(id, extra = {}) {
  return {
    id,
    placement: 'inline',
    content_type: 'image/png',
    width_emu: 914400,
    height_emu: 457200,
    raster: { pixel_width: 8, pixel_height: 4, sha256: id },
    ...extra,
  };
}

function document(drawings) {
  return {
    body: { blocks: drawings.map(item => ({ paragraph: { runs: [{ drawing: item }] } })) },
    headers: [],
    footers: [],
    notes: [],
    comment_stories: [],
  };
}

test('loadDocumentImages previews inline PNG and omits unsupported placement', async () => {
  const { loadDocumentImages, documentDrawings } = await loadMedia();
  const png = Buffer.from('png-bytes');
  const reader = { readMedia: async () => png };
  const model = document([drawing('ok'), drawing('float', { placement: 'anchor' })]);
  assert.equal(documentDrawings(model).length, 2);
  const cache = new Map();
  const first = await loadDocumentImages(reader, png, model, cache);
  assert.ok(first.images.ok.startsWith('data:image/png;base64,'));
  assert.equal(first.images.float, undefined);
  assert.match(first.notice, /not shown/);
  let reads = 0;
  const counting = { readMedia: async () => { reads += 1; return png; } };
  await loadDocumentImages(counting, png, document([drawing('ok')]), cache);
  assert.equal(reads, 0);
});

test('loadDocumentImages never splices the declared content type into the data URL', async () => {
  const { loadDocumentImages, rasterDataUrlPrefix, isRasterDataUrl } = await loadMedia();
  const bytes = Buffer.from('raster-bytes');
  const reader = { readMedia: async () => bytes };
  // A hostile or merely unexpected declared type must not reach the URL: the prefix is a lookup, not a template.
  const model = document([
    drawing('png'),
    drawing('jpeg', { content_type: 'image/jpeg' }),
    drawing('svg', { content_type: 'image/svg+xml' }),
    drawing('html', { content_type: 'text/html' }),
    drawing('spoof', { content_type: 'image/png;charset=x' }),
    drawing('none', { content_type: undefined }),
  ]);
  const { images, notice } = await loadDocumentImages(reader, bytes, model, new Map());
  assert.equal(images.png, 'data:image/png;base64,' + bytes.toString('base64'));
  assert.equal(images.jpeg, 'data:image/jpeg;base64,' + bytes.toString('base64'));
  for (const id of ['svg', 'html', 'spoof', 'none']) assert.equal(images[id], undefined, id);
  assert.match(notice, /4 images are preserved but not shown/);
  assert.equal(rasterDataUrlPrefix('image/png'), 'data:image/png;base64,');
  assert.equal(rasterDataUrlPrefix('text/html'), undefined);
  assert.equal(rasterDataUrlPrefix('constructor'), undefined);
  assert.equal(rasterDataUrlPrefix(undefined), undefined);
});

test('isRasterDataUrl admits only our own raster data URLs at the img sink', async () => {
  const { isRasterDataUrl } = await loadMedia();
  assert.equal(isRasterDataUrl('data:image/png;base64,AAAA'), true);
  assert.equal(isRasterDataUrl('data:image/jpeg;base64,AAAA'), true);
  assert.equal(isRasterDataUrl('data:text/html;base64,AAAA'), false);
  assert.equal(isRasterDataUrl('data:image/svg+xml;base64,AAAA'), false);
  assert.equal(isRasterDataUrl('javascript:alert(1)'), false);
  assert.equal(isRasterDataUrl('https://example.test/x.png'), false);
  assert.equal(isRasterDataUrl(undefined), false);
});

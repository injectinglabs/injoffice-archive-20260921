const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadMedia(media = new Map()) {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/document-media.ts'),
    platform: 'node',
    external: id => id.includes('packages/docs/src/') || id.includes('docxPreviewImages'),
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    new Function('require', 'module', 'exports', output[0].code)(id => id.includes('docxPreviewImages') ? {extractDocxPreviewImages: async () => media} : require(id), mod, mod.exports);
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
    media_part: `word/media/${id}.png`,
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
    passthrough_parts: drawings.map(d => ({part_name:d.media_part,sha256:d.id})),
  };
}

test('source images use verified media and cache URLs while refusing unsupported placement', async () => {
  const png = new Uint8Array([1,2,3]);
  const media = new Map([['word/media/ok.png', {bytes:png,mime:'image/png',width:8,height:4}], ['word/media/float.png', {bytes:png,mime:'image/png',width:8,height:4}]]);
  const {loadSourceDocumentImages, documentDrawings, releaseDocumentImages} = await loadMedia(media);
  const model = document([drawing('ok'), drawing('float', {placement:'anchor'}), drawing('missing')]);
  assert.equal(documentDrawings(model).length, 3);
  const cache = new Map();
  const first = await loadSourceDocumentImages(png, model, cache);
  assert.ok(first.images.ok.startsWith('blob:'));
  assert.equal(first.images.float, undefined);
  assert.equal(first.images.missing, undefined);
  assert.match(first.notice, /2 images/);
  const second = await loadSourceDocumentImages(png, model, cache);
  assert.equal(second.images.ok, first.images.ok);
  releaseDocumentImages(cache);
  assert.equal(cache.size, 0);
});

test('source image preview uses fixed raster MIME types and refuses unpainted transforms', async () => {
  const bytes = new Uint8Array([1]);
  const media = new Map(['png','jpeg','svg','crop'].map(id => [`word/media/${id}.png`, {bytes,mime:id==='jpeg'?'image/jpeg':id==='svg'?'image/svg+xml':'image/png',width:8,height:4}]));
  const {loadSourceDocumentImages,rasterContentType,isPreviewImageUrl} = await loadMedia(media);
  const model=document([drawing('png'),drawing('jpeg'),drawing('svg'),drawing('crop',{source_crop:{left:1000}})]);
  const {images,notice}=await loadSourceDocumentImages(bytes,model,new Map());
  assert.ok(isPreviewImageUrl(images.png)&&isPreviewImageUrl(images.jpeg));
  assert.equal(images.svg,undefined);assert.equal(images.crop,undefined);assert.match(notice,/2 images/);
  assert.equal(rasterContentType('image/png'),'image/png');
  for(const bad of ['text/html','image/svg+xml','constructor','image/png;charset=x',undefined])assert.equal(rasterContentType(bad),undefined);
});

test('isPreviewImageUrl admits only blob URLs at the img sink', async () => {
  const { isPreviewImageUrl } = await loadMedia();
  assert.equal(isPreviewImageUrl(URL.createObjectURL(new Blob([new Uint8Array([1])], { type: 'image/png' }))), true);
  for (const bad of ['data:image/png;base64,AAAA', 'data:text/html;base64,AAAA', 'javascript:alert(1)', 'https://example.test/x.png', undefined]) assert.equal(isPreviewImageUrl(bad), false, String(bad));
});

// Node 22+, Go, Chrome, and built workspace packages are required. The test
// starts its own local helper and Vite server; it never uses a running user app.
import { spawn, spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { createServer } from 'node:net'
import { tmpdir } from 'node:os'
import { resolve, sep } from 'node:path'
import { launchChromeForCDP, terminateProcess } from './chrome-cdp-startup.mjs'
import { startShowcaseDevServer } from './showcase-smoke-dev-server.mjs'

const root = resolve(process.env.DOCX_NATIVE_SMOKE_ROOT || resolve(import.meta.dirname, '..'))
const scratch = mkdtempSync(resolve(tmpdir(), 'injoffice-docx-native-smoke-'))
const artifacts = process.env.SHOWCASE_OUTPUT ? resolve(process.env.SHOWCASE_OUTPUT) : mkdtempSync(resolve(tmpdir(), 'injoffice-docx-native-screenshots-'))
mkdirSync(artifacts, { recursive: true })
const profiles = []
const errors = []
const docs = `document.querySelector('[data-demo-surface="docs"]')`
const native = `document.querySelector('[aria-label="Native document pages"]')`
let helper, server, chrome, cdp
let helperLog = ''
const previousApi = process.env.VITE_INJOFFICE_API_BASE
const previousServer = process.env.INJOFFICE_SERVER

try {
  const worker = resolve(root, 'apps/docx-page-paint-worker/dist/worker.js')
  if (!existsSync(worker)) throw new Error('Build the workspace packages and DOCX page-paint worker before running this smoke.')
  const fixture = resolve(scratch, 'native-two-pages.docx')
  await command('go', ['run', './cmd/nativepreviewfixture', '-font', resolve(root, 'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'), '-out', fixture, '-jpeg', '-page-fields', '-scripts', '-underline', 'double', '-floating'], resolve(root, 'go/docxpatch'))
  const originalHash = hash(readFileSync(fixture))
  const binary = resolve(scratch, process.platform === 'win32' ? 'injoffice-server.exe' : 'injoffice-server')
  await command('go', ['build', '-o', binary, './cmd/injoffice-server'], resolve(root, 'go/injoffice-server'))
  const port = await unusedPort()
  helper = spawn(binary, ['-addr', `127.0.0.1:${port}`, '-artifacts', resolve(scratch, 'artifacts'), '-docx-preview-worker', worker], { cwd: root, stdio: ['ignore', 'ignore', 'pipe'] })
  helper.stderr.setEncoding('utf8')
  helper.stderr.on('data', chunk => { helperLog = `${helperLog}${chunk}`.slice(-8000) })
  helper.on('error', error => { helperLog = error.message })
  await poll(async () => { try { return (await fetch(`http://127.0.0.1:${port}/healthz`, { signal: AbortSignal.timeout(1000) })).ok } catch { return false } }, 'isolated helper startup')
  // A nonempty relative base explicitly enables the helper via Vite's local
  // proxy, without introducing cross-origin permissions or an upload fallback.
  process.env.VITE_INJOFFICE_API_BASE = '.'
  process.env.INJOFFICE_SERVER = `http://127.0.0.1:${port}`
  server = await startShowcaseDevServer(resolve(root, 'apps/playground'))
  chrome = await launchChromeForCDP({ executable: findChrome(), createProfile: () => {
    const profile = mkdtempSync(resolve(tmpdir(), 'injoffice-docx-native-chrome-')); profiles.push(profile); return profile
  } })
  cdp = await connectCDP(chrome.target.webSocketDebuggerUrl)
  cdp.on('Runtime.exceptionThrown', ({ exceptionDetails }) => errors.push(exceptionDetails.exception?.description ?? exceptionDetails.text))
  await cdp.send('Runtime.enable')
  await cdp.send('Page.enable')
  await cdp.send('Page.addScriptToEvaluateOnNewDocument', { source: `{
    const originalFetch = window.fetch.bind(window); window.__nativeDocxPosts = [];
    window.fetch = async (input, init) => {
      if (init?.method === 'POST' && String(input).includes('/v1/')) {
        const bytes = init.body instanceof Blob ? await init.body.arrayBuffer() : new TextEncoder().encode(String(init.body));
        const digest = await crypto.subtle.digest('SHA-256', bytes);
        window.__nativeDocxPosts.push({ url: String(input), hash: [...new Uint8Array(digest)].map(n => n.toString(16).padStart(2, '0')).join('') });
      }
      const response = await originalFetch(input, init);
      if (init?.method === 'POST' && String(input).endsWith('/v1/docx/page-preview')) window.__nativeDocxPaint = (await response.clone().json()).page_paint_output;
      return response;
    };
  }` })
  await cdp.send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1100, deviceScaleFactor: 1, mobile: false })
  await cdp.send('Page.navigate', { url: `${server.url}#/docs?feature=editor` })
  await poll(() => evaluate(`Boolean(document.querySelector('[aria-label="Open a DOCX file"]'))`), 'DOCX workbench')
  await upload(fixture)
  await poll(() => evaluate(`${native}?.textContent.includes('Nothing is uploaded')`), 'native preview consent')
  await assert(`window.__nativeDocxPosts.length === 0`, 'opening a real DOCX uploads nothing')
  await click('Upload to helper and render native pages')
  await poll(() => evaluate(`${native}?.querySelector('svg[aria-label="Native document page 1"] path') !== null && ${native}?.textContent.includes('2 native pages')`), 'real font-shaped native pages', 45000)
  await assert(`window.__nativeDocxPosts.length === 1 && window.__nativeDocxPosts[0].hash === ${JSON.stringify(originalHash)}`, 'preview uploads exact original bytes only after consent')
  await assert(`${native}.querySelectorAll('svg').length === 1 && ${native}.querySelector('svg').getBoundingClientRect().height > 100`, 'one bounded native page is mounted')
  // DejaVu Sans maps decimal 1/2/9 to glyphs 20/21/28. The generated source
  // caches 999; its actual native header/footer must instead paint 1 of 2 / 2 of 2.
await assert(`window.__nativeDocxPaint.pages.every((page, index) => ['header','footer'].every(region => { const ids=page.lines.filter(line=>line.region===region).flatMap(line=>line.command_ids); const digits=page.commands.filter(command=>ids.includes(command.id)&&command.kind==='fill_glyph_path'&&[20,21,26,27,28].includes(command.glyph_id)).map(command=>command.glyph_id); return JSON.stringify(digits)===JSON.stringify([26+index,21]); }))`, 'PAGE/NUMPAGES derive per-page glyphs, never stale 999 cache')
  await assert(`${native}.querySelectorAll('svg path').length === window.__nativeDocxPaint.pages[0].commands.filter(command=>command.kind==='fill_glyph_path').length`, 'all native field glyph paths mount on page one')
  await assert(`(() => { const page=window.__nativeDocxPaint.pages[0]; const ids=page.lines.filter(line=>line.region==='body').flatMap(line=>line.command_ids); const scripts=page.commands.filter(command=>ids.includes(command.id)&&command.kind==='fill_glyph_path'&&command.glyph_id===21); const header=page.commands.find(command=>!ids.includes(command.id)&&command.kind==='fill_glyph_path'&&command.glyph_id===21); const bounds=command=>{const points=command.path.filter(point=>'y_millipoints'in point).map(point=>point.y_millipoints);return {top:Math.min(...points),bottom:Math.max(...points)}}; if(scripts.length!==2||!header)return false;const sub=bounds(scripts[0]),superScript=bounds(scripts[1]),normal=bounds(header);return sub.top>superScript.top&&sub.bottom>superScript.bottom&&sub.bottom-sub.top<normal.bottom-normal.top&&superScript.bottom-superScript.top<normal.bottom-normal.top; })()`, 'native subscript and superscript glyphs have smaller outlines and distinct font-metric baselines')
  await poll(async () => await evaluate(`(() => { const image = ${native}?.querySelector('svg image'); return !!image && image.getAttribute('href')?.startsWith('data:image/jpeg;base64,') && image.getAttribute('preserveAspectRatio') === 'none' })()`), 'native JPEG image and source-defined aspect ratio')
  await assert(`(() => { const lines = [...${native}.querySelectorAll('line[data-native-underline]')]; return lines.length > 0 && lines.length % 2 === 0 && lines.every(line => Number(line.getAttribute('x2')) > Number(line.getAttribute('x1')) && Number(line.getAttribute('stroke-width')) > 0 && line.getAttribute('y1') === line.getAttribute('y2')) })()`, 'native double underline uses bounded font-metric strokes')
  await assert(`(async () => { const image = new Image(); image.src = ${native}.querySelector('svg image').getAttribute('href'); await image.decode(); const canvas = document.createElement('canvas'); canvas.width=16; canvas.height=8; const context=canvas.getContext('2d'); context.drawImage(image,0,0); const left=context.getImageData(2,4,1,1).data; const right=context.getImageData(13,4,1,1).data; return image.naturalWidth===16 && image.naturalHeight===8 && left[0]>left[2]+80 && right[2]>right[0]+80 })()`, 'JPEG pixels decode to the generated red/blue pattern')
  await screenshot('docx-native-page-one.png')
  await assert(`(() => { const svg = ${native}.querySelector('svg'); const background = svg.querySelector('rect[data-native-highlight]'); const glyph = [...svg.querySelectorAll('path')].find(node => background && (background.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING)); return background?.getAttribute('fill') === '#FFFF00' && Number(background.getAttribute('width')) > 0 && Number(background.getAttribute('height')) > 0 && !!glyph })()`, 'source-bound native highlight precedes its body glyph painting after the header')
  await click('Next native page')
  await poll(() => evaluate(`${native}?.querySelector('svg[aria-label="Native document page 2"] path') !== null`), 'next native page')
  await assert(`${native}.querySelectorAll('svg').length === 1`, 'navigation keeps a single mounted SVG')
  await assert(`${native}.querySelectorAll('svg path').length === window.__nativeDocxPaint.pages[1].commands.filter(command=>command.kind==='fill_glyph_path').length`, 'all native field glyph paths mount on page two')
  await screenshot('docx-native-page-two.png')
  await assert(`(() => { const page=window.__nativeDocxPaint.pages[1], ids=page.lines.filter(line=>line.region==='body').flatMap(line=>line.command_ids); const digits=page.commands.filter(command=>ids.includes(command.id)&&command.kind==='fill_glyph_path'&&[20,21,26,27,28].includes(command.glyph_id)).map(command=>command.glyph_id); return JSON.stringify(digits)===JSON.stringify([27,21]); })()`, 'body PAGE and NUMPAGES converge to decimal-restarted page eight of two without cached 999')
  await assert(`window.__nativeDocxPaint.pages[0].commands.every(command => command.kind !== 'paint_floating_image') && (() => { const page = window.__nativeDocxPaint.pages[1]; const command = page.commands.at(-1); const image = ${native}.querySelector('svg image'); return command.kind === 'paint_floating_image' && command.layer === 'front' && command.stacking_order === 7 && command.x_millipoints === 288000 && command.y_millipoints === 216000 && image?.getAttribute('x') === '288000' && image?.getAttribute('y') === '216000' && image === ${native}.querySelector('svg').lastElementChild })()`, 'floating image follows its anchor paragraph to page two with exact page coordinates and foreground replay')
  await evaluate(`${native}.querySelector('svg').lastElementChild.scrollIntoView({ block: 'center' })`)
  await screenshot('docx-native-page-two-footer.png')
  await evaluate(`${native}.querySelector('h3').scrollIntoView({ block: 'start' })`)
  await click('Previous native page')
  await poll(() => evaluate(`${native}?.querySelector('svg[aria-label="Native document page 1"] path') !== null`), 'previous native page')
  await assert(`${docs}?.dataset.demoDirty !== 'true' && window.__nativeDocxPosts.every(request => request.url.endsWith('/v1/docx/page-preview'))`, 'preview never mutates the document')
  if (hash(readFileSync(fixture)) !== originalHash) throw new Error('The source fixture changed')
  // Exercise an actual SVG image decode failure after a successful render.
  // The viewer must clear the native success claim, not silently lose pixels.
  await evaluate(`${native}.querySelector('svg image').setAttribute('href', 'data:image/jpeg;base64,AA==')`)
  await poll(() => evaluate(`${native}?.textContent.includes('Native image could not be displayed') && ${native}?.querySelector('svg') === null`), 'image failure clears native pages visibly')
  await assert(`!${native}.textContent.includes('2 native pages') && ${docs}?.dataset.demoDirty !== 'true'`, 'failed image never leaves a misleading success claim or changes source')
  await screenshot('docx-native-image-failure.png')
  // A separate source PNG exercises the same native helper/image decoder,
  // without restarting the helper or falling back to the approximate view.
  const pngFixture = resolve(scratch, 'native-png.docx')
  const exported = spawnSync('go', ['test', '-count=1', '-run', '^TestNativePreviewPNGBrowserFixture$', '.'], { cwd: resolve(root, 'go/docxpatch/cmd/nativepreviewfixture'), env: { ...process.env, INJOFFICE_PNG_FIXTURE_OUTPUT: pngFixture, INJOFFICE_PNG_FIXTURE_FONT: resolve(root, 'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf') }, encoding: 'utf8', timeout: 60000 })
  if (exported.status !== 0) throw new Error(`PNG fixture export failed: ${exported.stderr}\n${exported.stdout}`)
  const pngHash = hash(readFileSync(pngFixture))
  await upload(pngFixture)
  await poll(() => evaluate(`${native}?.textContent.includes('Nothing is uploaded') && ${native}?.querySelector('svg') === null`), 'PNG source replacement clears old pages')
  await assert(`window.__nativeDocxPosts.length === 1`, 'PNG native preview requires new explicit consent')
  await click('Upload to helper and render native pages')
  await poll(() => evaluate(`${native}?.textContent.includes('2 native pages') && ${native}?.querySelector('svg image')?.getAttribute('href')?.startsWith('data:image/png;base64,')`), 'native PNG page paints', 45000)
  await assert(`(async () => { const node=${native}.querySelector('svg image'); const image=new Image(); image.src=node.getAttribute('href'); await image.decode(); const canvas=document.createElement('canvas');canvas.width=16;canvas.height=8;const context=canvas.getContext('2d');context.drawImage(image,0,0);const left=context.getImageData(2,4,1,1).data;const right=context.getImageData(13,4,1,1).data;return image.naturalWidth===16&&image.naturalHeight===8&&left[0]===220&&left[1]===40&&left[2]===40&&right[0]===30&&right[1]===80&&right[2]===220&&node.getAttribute('preserveAspectRatio')==='none' })()`, 'PNG exact source pixels and extents survive native replay')
  await assert(`window.__nativeDocxPosts.length === 2 && window.__nativeDocxPosts[1].hash === ${JSON.stringify(pngHash)} && ${docs}?.dataset.demoDirty !== 'true'`, 'PNG preview sends only unchanged original bytes')
  if (hash(readFileSync(pngFixture)) !== pngHash) throw new Error('PNG source fixture changed')
  await screenshot('docx-native-png.png')
  const tableFixture = resolve(scratch, 'native-repeating-table.docx')
  const tableExport = spawnSync('go', ['test', '-count=1', '-run', '^TestNativePreviewRepeatingTableBrowserFixture$', '.'], { cwd: resolve(root, 'go/docxpatch/cmd/nativepreviewfixture'), env: { ...process.env, INJOFFICE_TABLE_FIXTURE_OUTPUT: tableFixture, INJOFFICE_TABLE_FIXTURE_FONT: resolve(root, 'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf') }, encoding: 'utf8', timeout: 60000 })
  if (tableExport.status !== 0) throw new Error(`Table fixture export failed: ${tableExport.stderr}\n${tableExport.stdout}`)
  const tableHash = hash(readFileSync(tableFixture))
  await upload(tableFixture)
  await poll(() => evaluate(`${native}?.textContent.includes('Nothing is uploaded') && ${native}?.querySelector('svg') === null`), 'table source replacement clears old pages')
  await assert(`window.__nativeDocxPosts.length === 2`, 'table preview needs new explicit consent')
  await click('Upload to helper and render native pages')
  await poll(() => evaluate(`${native}?.textContent.includes('4 native pages') && ${native}?.querySelector('svg path') !== null`), 'real two-column repeating table native pages', 45000)
  const tableGeometry = `(() => { const svg = ${native}.querySelector('svg'); const headers=[...svg.querySelectorAll('rect[fill="#DDEEFF"]')]; return headers.length===2 && headers.every(rect=>Number(rect.getAttribute('y'))===72000&&Number(rect.getAttribute('height'))===36000&&Number(rect.getAttribute('width'))===234000) && Number(headers[0].getAttribute('x'))===72000 && Number(headers[1].getAttribute('x'))===306000 && svg.querySelectorAll('path').length>30 && svg.querySelectorAll('line').length>=8 })()`
  await assert(tableGeometry, 'initial fixed-grid header cells have exact source dimensions and glyphs')
  await screenshot('docx-native-table-first.png')
  for (let page = 2; page <= 4; page += 1) {
    await click('Next native page')
    await poll(() => evaluate(`${native}?.querySelector('svg[aria-label="Native document page ${page}"]') !== null`), `table continuation page ${page}`)
    await assert(tableGeometry, `page ${page} repeats both source header cells and preserves body glyphs`)
  }
  await screenshot('docx-native-table-last.png')
  await assert(`window.__nativeDocxPosts.length===3 && window.__nativeDocxPosts[2].hash===${JSON.stringify(tableHash)} && ${docs}?.dataset.demoDirty !== 'true' && ${native}.querySelectorAll('svg').length===1`, 'table navigation retains exact original source and bounded page mounting')
  if (hash(readFileSync(tableFixture)) !== tableHash) throw new Error('Table source fixture changed')
  const transformedFixture = resolve(scratch, 'native-transformed-image.docx')
  const transformExport = spawnSync('go', ['test', '-count=1', '-run', '^TestNativePreviewTransformedImageBrowserFixture$', '.'], { cwd: resolve(root, 'go/docxpatch/cmd/nativepreviewfixture'), env: { ...process.env, INJOFFICE_TRANSFORM_FIXTURE_OUTPUT: transformedFixture, INJOFFICE_TRANSFORM_FIXTURE_FONT: resolve(root, 'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf') }, encoding: 'utf8', timeout: 60000 })
  if (transformExport.status !== 0) throw new Error(`Transform fixture export failed: ${transformExport.stderr}\n${transformExport.stdout}`)
  const transformedHash = hash(readFileSync(transformedFixture))
  await upload(transformedFixture)
  await poll(() => evaluate(`${native}?.textContent.includes('Nothing is uploaded') && ${native}?.querySelector('svg') === null`), 'image transform source replacement clears old pages')
  await assert(`window.__nativeDocxPosts.length === 3`, 'transformed image preview needs explicit consent')
  await click('Upload to helper and render native pages')
  await poll(() => evaluate(`${native}?.textContent.includes('2 native pages') && ${native}?.querySelector('svg image') !== null`), 'source-transformed native image', 45000)
  await assert(`(async () => {
    const node=${native}.querySelector('svg image'); const x=Number(node.getAttribute('x')), y=Number(node.getAttribute('y')), w=Number(node.getAttribute('width')), h=Number(node.getAttribute('height'));
    if(node.getAttribute('transform')!==('matrix(-1 0 0 1 '+(2*x+w)+' 0)'))return false;
    const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');svg.setAttribute('width','16');svg.setAttribute('height','8');svg.setAttribute('viewBox',[x,y,w,h].join(' '));svg.appendChild(node.cloneNode(true));
    const raster=new Image();raster.src='data:image/svg+xml;base64,'+btoa(new XMLSerializer().serializeToString(svg));await raster.decode();
    const canvas=document.createElement('canvas');canvas.width=16;canvas.height=8;const ctx=canvas.getContext('2d');ctx.drawImage(raster,0,0);const left=ctx.getImageData(2,4,1,1).data,right=ctx.getImageData(13,4,1,1).data;
    return left[2]>left[0]+80&&right[0]>right[2]+80&&left[3]===255&&right[3]===255;
  })()`, 'actual SVG raster pixels apply source half-turn plus vertical flip inside the unchanged box')
  await assert(`window.__nativeDocxPosts.length===4 && window.__nativeDocxPosts[3].hash===${JSON.stringify(transformedHash)} && ${docs}?.dataset.demoDirty !== 'true'`, 'orientation preview sends only unchanged source bytes')
  if (hash(readFileSync(transformedFixture)) !== transformedHash) throw new Error('Transformed image source changed')
  await screenshot('docx-native-transformed-image.png')
  let finalPreviewPosts = 4
  for (const angle of [90, 270]) {
    const quarterFixture = resolve(scratch, `native-quarter-${angle}.docx`)
    const exported = spawnSync('go', ['test', '-count=1', '-run', '^TestNativePreviewTransformedImageBrowserFixture$', '.'], { cwd: resolve(root, 'go/docxpatch/cmd/nativepreviewfixture'), env: { ...process.env, INJOFFICE_TRANSFORM_FIXTURE_OUTPUT: quarterFixture, INJOFFICE_TRANSFORM_FIXTURE_FONT: resolve(root, 'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'), INJOFFICE_TRANSFORM_ANGLE: String(angle) }, encoding: 'utf8', timeout: 60000 })
    if (exported.status !== 0) throw new Error(`Quarter-turn fixture failed: ${exported.stderr}\n${exported.stdout}`)
    const sourceHash = hash(readFileSync(quarterFixture))
    await upload(quarterFixture)
    await poll(() => evaluate(`${native}?.textContent.includes('Nothing is uploaded') && ${native}?.querySelector('svg') === null`), 'quarter-turn source replacement')
    await assert(`window.__nativeDocxPosts.length===${finalPreviewPosts}`, 'quarter-turn preview requires fresh consent')
    await click('Upload to helper and render native pages')
    await poll(() => evaluate(`${native}?.textContent.includes('2 native pages') && ${native}?.querySelector('svg image') !== null`), `native ${angle}-degree image`, 45000)
    await assert(`(async () => {
      const node=${native}.querySelector('svg image'),x=Number(node.getAttribute('x')),y=Number(node.getAttribute('y'));
      if(Number(node.getAttribute('width'))!==144000||Number(node.getAttribute('height'))!==36000)return false;
      const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');svg.setAttribute('width','8');svg.setAttribute('height','16');svg.setAttribute('viewBox',[x,y,36000,144000].join(' '));svg.appendChild(node.cloneNode(true));
      const raster=new Image();raster.src='data:image/svg+xml;base64,'+btoa(new XMLSerializer().serializeToString(svg));await raster.decode();const canvas=document.createElement('canvas');canvas.width=8;canvas.height=16;const ctx=canvas.getContext('2d');ctx.drawImage(raster,0,0);const top=ctx.getImageData(4,2,1,1).data,bottom=ctx.getImageData(4,13,1,1).data;
      return ${angle === 90 ? 'top[0]>top[2]+80&&bottom[2]>bottom[0]+80' : 'top[2]>top[0]+80&&bottom[0]>bottom[2]+80'}&&top[3]===255&&bottom[3]===255;
    })()`, `${angle}-degree source rotation has correct raster colors inside the exact swapped layout bounds`)
    finalPreviewPosts += 1
    await assert(`window.__nativeDocxPosts.length===${finalPreviewPosts} && window.__nativeDocxPosts[${finalPreviewPosts-1}].hash===${JSON.stringify(sourceHash)} && ${docs}?.dataset.demoDirty !== 'true'`, 'quarter-turn preview keeps source bytes unchanged')
    await screenshot(`docx-native-quarter-${angle}.png`)
  }
  const percentFixture = resolve(scratch, 'native-percent-table.docx')
  const percentExport = spawnSync('go', ['test', '-count=1', '-run', '^TestNativePreviewRepeatingTableBrowserFixture$', '.'], { cwd: resolve(root, 'go/docxpatch/cmd/nativepreviewfixture'), env: { ...process.env, INJOFFICE_TABLE_FIXTURE_OUTPUT: percentFixture, INJOFFICE_TABLE_FIXTURE_FONT: resolve(root, 'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'), INJOFFICE_TABLE_PERCENT: '2500' }, encoding: 'utf8', timeout: 60000 })
  if (percentExport.status !== 0) throw new Error(`Percentage fixture failed: ${percentExport.stderr}\n${percentExport.stdout}`)
  const percentHash = hash(readFileSync(percentFixture))
  await upload(percentFixture)
  await poll(() => evaluate(`${native}?.textContent.includes('Nothing is uploaded') && ${native}?.querySelector('svg') === null`), 'percentage table source replacement')
  await assert(`window.__nativeDocxPosts.length===${finalPreviewPosts}`, 'percentage table requires fresh consent')
  await click('Upload to helper and render native pages')
  await poll(() => evaluate(`${native}?.textContent.includes('4 native pages') && ${native}?.querySelector('svg path') !== null`), 'real percentage-width table paints', 45000)
  const percentGeometry = `(() => { const cells=[...${native}.querySelectorAll('rect[fill="#DDEEFF"]')];return cells.length===2&&cells.every(cell=>Number(cell.getAttribute('width'))===117000)&&Number(cells[0].getAttribute('x'))===72000&&Number(cells[1].getAttribute('x'))===189000 })()`
  await assert(percentGeometry, '50-percent table retains proportional columns and exact source section placement')
  await click('Next native page')
  await poll(() => evaluate(`${native}?.querySelector('svg[aria-label="Native document page 2"]')!==null`), 'percentage header continuation')
  await assert(percentGeometry, 'percentage heading repeats with identical qualified widths')
  finalPreviewPosts += 1
  await assert(`window.__nativeDocxPosts.length===${finalPreviewPosts} && window.__nativeDocxPosts[${finalPreviewPosts-1}].hash===${JSON.stringify(percentHash)} && ${docs}?.dataset.demoDirty !== 'true'`, 'percentage preview leaves original source unchanged')
  if(hash(readFileSync(percentFixture))!==percentHash)throw new Error('Percentage source changed')
  await screenshot('docx-native-percent-table.png')
  const splitFixture = resolve(scratch, 'native-split-table.docx')
  const splitExport = spawnSync('go', ['test', '-count=1', '-run', '^TestNativePreviewRepeatingTableBrowserFixture$', '.'], { cwd: resolve(root, 'go/docxpatch/cmd/nativepreviewfixture'), env: { ...process.env, INJOFFICE_TABLE_FIXTURE_OUTPUT: splitFixture, INJOFFICE_TABLE_FIXTURE_FONT: resolve(root, 'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'), INJOFFICE_TABLE_SPLIT: 'true' }, encoding: 'utf8', timeout: 60000 })
  if (splitExport.status !== 0) throw new Error(`Split fixture failed: ${splitExport.stderr}\n${splitExport.stdout}`)
  const splitHash = hash(readFileSync(splitFixture))
  await upload(splitFixture)
  await poll(() => evaluate(`${native}?.textContent.includes('Nothing is uploaded') && ${native}?.querySelector('svg') === null`), 'split table source replacement')
  await assert(`window.__nativeDocxPosts.length===${finalPreviewPosts}`, 'split table requires fresh consent')
  await click('Upload to helper and render native pages')
  await poll(() => evaluate(`${native}?.querySelector('svg path') !== null`), 'natural table row paints across pages', 45000)
  await click('Next native page')
  await poll(() => evaluate(`${native}?.querySelector('svg[aria-label="Native document page 2"]')!==null`), 'split row continuation')
  await assert(`(() => {const cells=[...${native}.querySelectorAll('rect[fill="#EEF5EE"]')];return cells.length===2&&cells.every(cell=>Number(cell.getAttribute('y'))===108000&&Number(cell.getAttribute('height'))<=44000)&&${native}.querySelectorAll('rect[fill="#DDEEFF"]').length===2})()`, 'continuation retains both cell fills and repeated headers within page bounds')
  finalPreviewPosts += 1
  await assert(`window.__nativeDocxPosts.length===${finalPreviewPosts} && window.__nativeDocxPosts[${finalPreviewPosts-1}].hash===${JSON.stringify(splitHash)} && ${docs}?.dataset.demoDirty !== 'true'`, 'row fragmentation preserves original source')
  if(hash(readFileSync(splitFixture))!==splitHash)throw new Error('Split source changed')
  await screenshot('docx-native-split-table.png')
  const splitPageCount = await evaluate(`Number(${native}.textContent.match(/([0-9]+) native pages/)[1])`)
  if (!Number.isInteger(splitPageCount) || splitPageCount < 3 || splitPageCount > 32) throw new Error('Unexpected bounded split-row page count')
  for (let page = 3; page <= splitPageCount; page += 1) {
    await click('Next native page')
    await poll(() => evaluate(`${native}?.querySelector('svg[aria-label="Native document page ${page}"]')!==null`), `split-row page ${page}`)
    await assert(`(() => {const cells=[...${native}.querySelectorAll('rect[fill="#EEF5EE"]')];return cells.length>=2&&cells.every(cell=>Number(cell.getAttribute('y'))>=108000&&Number(cell.getAttribute('y'))+Number(cell.getAttribute('height'))<=152000)&&${native}.querySelectorAll('rect[fill="#DDEEFF"]').length===2&&${native}.querySelector('svg path')!==null})()`, `split table page ${page} retains source geometry within the body`)
  }
  await screenshot('docx-native-split-table-last.png')
  const autofitFixture = resolve(scratch, 'native-content-autofit-table.docx')
  const autofitExport = spawnSync('go', ['test', '-count=1', '-run', '^TestNativePreviewRepeatingTableBrowserFixture$', '.'], { cwd: resolve(root, 'go/docxpatch/cmd/nativepreviewfixture'), env: { ...process.env, INJOFFICE_TABLE_FIXTURE_OUTPUT: autofitFixture, INJOFFICE_TABLE_FIXTURE_FONT: resolve(root, 'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'), INJOFFICE_TABLE_AUTOFIT: 'true' }, encoding: 'utf8', timeout: 60000 })
  if (autofitExport.status !== 0) throw new Error(`Autofit fixture failed: ${autofitExport.stderr}\n${autofitExport.stdout}`)
  const autofitHash = hash(readFileSync(autofitFixture))
  await upload(autofitFixture)
  await poll(() => evaluate(`${native}?.textContent.includes('Nothing is uploaded') && ${native}?.querySelector('svg') === null`), 'autofit source replacement')
  await assert(`window.__nativeDocxPosts.length===${finalPreviewPosts}`, 'autofit requires explicit consent')
  await click('Upload to helper and render native pages')
  await poll(() => evaluate(`${native}?.querySelector('svg path')!==null`), 'real font-shaped content autofit', 45000)
  const autofitGeometry = `(() => {const cells=[...${native}.querySelectorAll('rect[fill="#DDEEFF"]')];if(cells.length!==2)return false;const widths=cells.map(cell=>Number(cell.getAttribute('width')));return widths[0]<widths[1]&&widths[0]+widths[1]===200000&&Number(cells[0].getAttribute('x'))===72000&&Number(cells[1].getAttribute('x'))===72000+widths[0]})()`
  await assert(autofitGeometry, 'equal source grid becomes unequal content-driven column widths at the preferred table width')
  await screenshot('docx-native-content-autofit.png')
  await click('Next native page')
  await poll(() => evaluate(`${native}?.querySelector('svg[aria-label="Native document page 2"]')!==null`), 'autofit continuation')
  await assert(autofitGeometry, 'autofit headers repeat with the same source-qualified intrinsic allocation')
  finalPreviewPosts += 1
  await assert(`window.__nativeDocxPosts.length===${finalPreviewPosts} && window.__nativeDocxPosts[${finalPreviewPosts-1}].hash===${JSON.stringify(autofitHash)} && ${docs}?.dataset.demoDirty !== 'true'`, 'autofit keeps original source bytes unchanged')
  if(hash(readFileSync(autofitFixture))!==autofitHash)throw new Error('Autofit source changed')
  const cropFixture = resolve(scratch, 'native-cropped-quarter-image.docx')
  const cropExport = spawnSync('go', ['test', '-count=1', '-run', '^TestNativePreviewTransformedImageBrowserFixture$', '.'], { cwd: resolve(root, 'go/docxpatch/cmd/nativepreviewfixture'), env: { ...process.env, INJOFFICE_TRANSFORM_FIXTURE_OUTPUT: cropFixture, INJOFFICE_TRANSFORM_FIXTURE_FONT: resolve(root, 'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'), INJOFFICE_TRANSFORM_ANGLE: '90', INJOFFICE_TRANSFORM_CROP: 'left-half' }, encoding: 'utf8', timeout: 60000 })
  if (cropExport.status !== 0) throw new Error(`Crop fixture failed: ${cropExport.stderr}\n${cropExport.stdout}`)
  const cropHash = hash(readFileSync(cropFixture))
  await upload(cropFixture)
  await poll(() => evaluate(`${native}?.textContent.includes('Nothing is uploaded') && ${native}?.querySelector('svg') === null`), 'crop source replacement')
  await assert(`window.__nativeDocxPosts.length===${finalPreviewPosts}`, 'crop preview requires fresh consent')
  await click('Upload to helper and render native pages')
  await poll(() => evaluate(`${native}?.textContent.includes('2 native pages') && ${native}?.querySelector('svg[data-native-crop] image') !== null`), 'native cropped and rotated image', 45000)
  await assert(`(async () => {
    const node=${native}.querySelector('svg[data-native-crop]'),x=Number(node.getAttribute('x')),y=Number(node.getAttribute('y'));
    if(node.getAttribute('viewBox')!=='50000 0 50000 100000'||node.getAttribute('overflow')!=='hidden')return false;
    const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');svg.setAttribute('width','8');svg.setAttribute('height','16');svg.setAttribute('viewBox',[x,y,36000,144000].join(' '));svg.appendChild(node.cloneNode(true));
    const raster=new Image();raster.src='data:image/svg+xml;base64,'+btoa(new XMLSerializer().serializeToString(svg));await raster.decode();const canvas=document.createElement('canvas');canvas.width=8;canvas.height=16;const ctx=canvas.getContext('2d');ctx.drawImage(raster,0,0);
    return [2,6,10,14].every(y=>{const pixel=ctx.getImageData(4,y,1,1).data;return pixel[2]>pixel[0]+80&&pixel[3]===255});
  })()`, 'source crop removes red pixels before quarter-turn rotation without escaping the layout box')
  finalPreviewPosts += 1
  await assert(`window.__nativeDocxPosts.length===${finalPreviewPosts} && window.__nativeDocxPosts[${finalPreviewPosts-1}].hash===${JSON.stringify(cropHash)} && ${docs}?.dataset.demoDirty !== 'true'`, 'crop preview leaves source bytes unchanged')
  if (hash(readFileSync(cropFixture)) !== cropHash) throw new Error('Crop source changed')
  await screenshot('docx-native-cropped-quarter.png')
  const squareFixture=resolve(scratch,'native-square-wrapped-image.docx')
  const squareExport=spawnSync('go',['test','-count=1','-run','^TestNativePreviewSquareWrapBrowserFixture$','.'],{cwd:resolve(root,'go/docxpatch/cmd/nativepreviewfixture'),env:{...process.env,INJOFFICE_SQUARE_FIXTURE_OUTPUT:squareFixture,INJOFFICE_SQUARE_FIXTURE_FONT:resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf')},encoding:'utf8',timeout:60000})
  if(squareExport.status!==0)throw new Error(`Square-wrap fixture: ${squareExport.stderr}\n${squareExport.stdout}`)
  const squareHash=hash(readFileSync(squareFixture));await upload(squareFixture)
  await poll(()=>evaluate(`${native}?.textContent.includes('Nothing is uploaded')&&${native}?.querySelector('svg')===null`),'square-wrap source replacement');await click('Upload to helper and render native pages')
  await poll(()=>evaluate(`window.__nativeDocxPaint?.status==='painted'&&${native}?.querySelector('svg image')!==null`),'source-bound square-wrapped page',45000)
  await assert(`(()=>{const p=window.__nativeDocxPaint.pages[0],image=p.commands.find(c=>c.kind==='paint_floating_image');if(!image||image.x_millipoints!==72000||image.y_millipoints!==72000)return false;const paths=p.commands.filter(c=>c.kind==='fill_glyph_path'&&c.path.length);const points=c=>c.path.filter(p=>'x_millipoints'in p);const beside=paths.filter(c=>points(c).some(p=>p.y_millipoints<172000)),below=paths.filter(c=>points(c).every(p=>p.y_millipoints>=172000));return beside.length>10&&below.length>10&&beside.every(c=>points(c).every(p=>p.x_millipoints>=216000))&&below.some(c=>points(c).some(p=>p.x_millipoints<216000));})()`,'glyphs exclude the square image and return to full width below it')
  finalPreviewPosts+=1
  await assert(`window.__nativeDocxPosts.length===${finalPreviewPosts}&&window.__nativeDocxPosts[${finalPreviewPosts-1}].hash===${JSON.stringify(squareHash)}&&${docs}?.dataset.demoDirty!=='true'`,'square wrapping leaves source bytes unchanged')
  await screenshot('docx-native-square-wrap.png')
  console.log('PASS: native source-bound square wrapping excludes the image and restores full line width below it')
  // This existing real DOCX has no embedded qualified font assets. It must
  // retain its approximate content view rather than invent native glyphs.
  const unsupported = resolve(scratch, 'unsupported-font.docx')
  writeFileSync(unsupported, Buffer.from(readFileSync(resolve(root, 'apps/playground/public/native-docx/northstar-launch-brief.docx.b64'), 'utf8').trim(), 'base64'))
  await upload(unsupported)
  await poll(() => evaluate(`${native}?.textContent.includes('Nothing is uploaded') && ${native}?.querySelector('svg') === null`), 'source replacement clears stale pages')
  await assert(`window.__nativeDocxPosts.length === ${finalPreviewPosts}`, 'replacement document also requires explicit consent')
  await click('Upload to helper and render native pages')
  await poll(() => evaluate(`${native}?.textContent.includes('original file is unchanged')`), 'unsupported document explicitly refused', 45000)
  await assert(`${native}.querySelector('svg') === null && document.querySelectorAll('.docx-editable-run').length > 0 && ${docs}?.dataset.demoDirty !== 'true'`, 'refusal retains approximate editable content and original source')
  await screenshot('docx-native-refusal.png')
  if (errors.length) throw new Error(`Browser exceptions: ${errors.join('\n')}`)
  console.log(JSON.stringify({ result: 'PASS', checks: ['explicit upload consent', 'real embedded-font shaping and pagination', 'paragraph-mark formatting and empty paragraph', 'native SVG glyphs', 'native PAGE/NUMPAGES in body, header and footer with decimal restart, stale cache ignored', 'native subscript/superscript outlines and baselines', 'native font-metric double underline', 'native text highlight behind glyphs', 'native JPEG pixels and source extents', 'native PNG pixels and source extents', 'page-relative floating image follows source paragraph to page two and paints in front', 'real two-column repeating table across four pages', 'exact percentage-width table geometry', 'line-safe natural table row fragments across every page', 'source image flips and all quarter turns verified in raster pixels', 'source crop before rotation verified in raster pixels', 'image decode failure clears native success', 'bounded page navigation', 'original source unchanged', 'source replacement clears stale output', 'unsupported rendering refusal'], screenshots: artifacts }, null, 2))
} catch (error) {
  console.error(`Native DOCX screenshots: ${artifacts}\nHelper diagnostics: ${helperLog}`)
  if (cdp) {
    await screenshot('failure.png').catch(() => undefined)
    console.error(await evaluate(`document.body.innerText.slice(-6000)`).catch(() => 'No browser diagnostic'))
  }
  throw error
} finally {
  cdp?.close()
  try { if (chrome) await terminateProcess(chrome.child) }
  finally {
    try { await server?.close() }
    finally {
      if (helper) await terminateProcess(helper)
      for (const profile of profiles) rmSync(profile, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 })
      rmSync(scratch, { recursive: true, force: true, maxRetries: 10, retryDelay: 100 })
      if (previousApi === undefined) delete process.env.VITE_INJOFFICE_API_BASE; else process.env.VITE_INJOFFICE_API_BASE = previousApi
      if (previousServer === undefined) delete process.env.INJOFFICE_SERVER; else process.env.INJOFFICE_SERVER = previousServer
    }
  }
}

function hash(bytes) { return createHash('sha256').update(bytes).digest('hex') }
async function command(executable, args, cwd) {
  await new Promise((resolvePromise, reject) => {
    const child = spawn(executable, args, { cwd, stdio: ['ignore', 'ignore', 'pipe'] }); let stderr = ''
    const timer = setTimeout(() => { child.kill('SIGKILL'); reject(new Error(`${executable} timed out`)) }, 120000)
    child.stderr.on('data', chunk => { stderr = `${stderr}${chunk}`.slice(-8000) })
    child.on('error', error => { clearTimeout(timer); reject(error) })
    child.on('exit', code => { clearTimeout(timer); code === 0 ? resolvePromise() : reject(new Error(`${executable} exited ${code}: ${stderr}`)) })
  })
}
async function unusedPort() {
  const socket = createServer()
  await new Promise((done, reject) => { socket.once('error', reject); socket.listen(0, '127.0.0.1', done) })
  const port = socket.address().port
  await new Promise(done => socket.close(done)); return port
}
async function upload(path) {
  const handle = await cdp.send('Runtime.evaluate', { expression: `document.querySelector('[aria-label="Open a DOCX file"]')` })
  const node = await cdp.send('DOM.describeNode', { objectId: handle.result.objectId })
  await cdp.send('DOM.setFileInputFiles', { backendNodeId: node.node.backendNodeId, files: [path] })
  await cdp.send('Runtime.releaseObject', { objectId: handle.result.objectId })
}
async function evaluate(expression) {
  const value = await cdp.send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true })
  if (value.exceptionDetails) throw new Error(value.exceptionDetails.exception?.description ?? value.exceptionDetails.text)
  return value.result.value
}
async function assert(expression, label) { if (!await evaluate(expression)) throw new Error(`Failed: ${label}`) }
async function poll(check, label, timeout = 30000) {
  const end = Date.now() + timeout
  while (Date.now() < end) { if (await check()) return; await new Promise(done => setTimeout(done, 75)) }
  throw new Error(`Timed out: ${label}`)
}
async function click(label) {
  const handle = await cdp.send('Runtime.evaluate', { expression: 'globalThis' })
  try {
    await poll(async () => {
      const value = await cdp.send('Runtime.callFunctionOn', {
        functionDeclaration: `function (label) {
          const section = document.querySelector('[aria-label="Native document pages"]');
          const button = [...(section?.querySelectorAll('button') ?? [])].find(button => button.textContent.trim() === label);
          if (!button || button.disabled) return false;
          button.click();
          return true;
        }`,
        objectId: handle.result.objectId,
        arguments: [{ value: label }],
        returnByValue: true,
      })
      if (value.exceptionDetails) throw new Error(value.exceptionDetails.exception?.description ?? value.exceptionDetails.text)
      return value.result.value
    }, label)
  } finally { await cdp.send('Runtime.releaseObject', { objectId: handle.result.objectId }) }
}
async function screenshot(name) {
  await evaluate(`${native}?.scrollIntoView({ block: 'start' }); window.scrollBy(0, -80)`)
  const result = await cdp.send('Page.captureScreenshot', { format: 'png' })
  writeFileSync(resolve(artifacts, name), Buffer.from(result.data, 'base64'))
}
function findChrome() {
  for (const candidate of [process.env.CHROME_PATH, process.env.CHROME_BIN, '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/Applications/Chromium.app/Contents/MacOS/Chromium', 'google-chrome', 'google-chrome-stable', 'chromium', 'chromium-browser'].filter(Boolean)) {
    if (candidate.includes(sep) && !existsSync(candidate)) continue
    if (spawnSync(candidate, ['--version'], { stdio: 'ignore', timeout: 5000 }).status === 0) return candidate
  }
  throw new Error('Chrome/Chromium is required; set CHROME_PATH')
}
async function connectCDP(url) {
  const socket = new WebSocket(url)
  await new Promise((done, reject) => { socket.addEventListener('open', done, { once: true }); socket.addEventListener('error', reject, { once: true }) })
  let next = 0
  const pending = new Map(), listeners = new Map()
  socket.addEventListener('message', event => {
    const message = JSON.parse(event.data)
    if (message.id !== undefined) {
      const request = pending.get(message.id); if (!request) return
      pending.delete(message.id); clearTimeout(request.timer)
      if (message.error) request.reject(new Error(message.error.message)); else request.resolve(message.result)
    } else for (const callback of listeners.get(message.method) ?? []) callback(message.params)
  })
  return {
    send(method, params = {}) { return new Promise((resolvePromise, reject) => { const id = ++next; const timer = setTimeout(() => { pending.delete(id); reject(new Error(`CDP timeout: ${method}`)) }, 30000); pending.set(id, { resolve: resolvePromise, reject, timer }); socket.send(JSON.stringify({ id, method, params })) }) },
    on(method, callback) { listeners.set(method, [...listeners.get(method) ?? [], callback]) },
    close() { for (const request of pending.values()) { clearTimeout(request.timer); request.reject(new Error('CDP closed')) } pending.clear(); socket.close() },
  }
}

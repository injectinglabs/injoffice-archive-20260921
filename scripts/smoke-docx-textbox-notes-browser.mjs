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
const scratch = mkdtempSync(resolve(tmpdir(), 'injoffice-docx-textbox-notes-smoke-'))
const artifacts = process.env.SHOWCASE_OUTPUT ? resolve(process.env.SHOWCASE_OUTPUT) : mkdtempSync(resolve(tmpdir(), 'injoffice-docx-native-screenshots-'))
mkdirSync(artifacts, { recursive: true })
const profiles = []
const errors = []
const docs = `document.querySelector('[data-demo-surface="docs"]')`
const native = `document.querySelector('[aria-label="Native document pages"]')`
let helper, server, chrome, cdp
let helperLog = ''
let checks=0
const previousApi = process.env.VITE_INJOFFICE_API_BASE
const previousServer = process.env.INJOFFICE_SERVER

try {
  const worker = resolve(root, 'apps/docx-page-paint-worker/dist/worker.js')
  if (!existsSync(worker)) throw new Error('Build the workspace packages and DOCX page-paint worker before running this smoke.')
  const textboxDir=resolve(scratch,'textbox'),footnoteDir=resolve(scratch,'footnote'),lineDir=resolve(scratch,'footnote-lines'),flowDir=resolve(scratch,'footnote-flow'),multiDir=resolve(scratch,'multiple-textboxes'),positionDir=resolve(scratch,'positioned-textboxes'),stackDir=resolve(scratch,'stacked-textboxes'),parityDir=resolve(scratch,'parity-textboxes')
  for(const [dir,test,env] of [[textboxDir,'TestNativeTextboxPageSource','INJOFFICE_TEXTBOX_PAGE_EVIDENCE_DIR'],[footnoteDir,'TestNativeFootnoteContinuationSource','INJOFFICE_FOOTNOTE_EVIDENCE_DIR'],[lineDir,'TestNativeFootnoteLineContinuationSource','INJOFFICE_FOOTNOTE_LINE_EVIDENCE_DIR'],[flowDir,'TestNativeFootnoteSharedFlowSource','INJOFFICE_FOOTNOTE_FLOW_EVIDENCE_DIR'],[multiDir,'TestNativeMultipleTextboxPagesSource','INJOFFICE_TEXTBOX_PAGES_EVIDENCE_DIR'],[positionDir,'TestNativeRelativeTextboxPagesSource','INJOFFICE_TEXTBOX_POSITION_EVIDENCE_DIR'],[stackDir,'TestNativeStackedTextboxPagesSource','INJOFFICE_TEXTBOX_STACK_EVIDENCE_DIR'],[parityDir,'TestNativeParityTextboxPagesSource','INJOFFICE_TEXTBOX_PARITY_EVIDENCE_DIR']]){
    const result=spawnSync('go',['test','-count=1','-run','^'+test+'$','./cmd/nativepreviewfixture'],{cwd:resolve(root,'go/docxpatch'),env:{...process.env,[env]:dir},encoding:'utf8',timeout:120000})
    if(result.status!==0)throw new Error('Source fixture export failed: '+result.stderr+' '+result.stdout)
  }
  const textboxFixture=resolve(textboxDir,'page-textbox.docx'),footnoteFixture=resolve(footnoteDir,'footnote-continuation.docx')
  const textboxHash=hash(readFileSync(textboxFixture)),footnoteHash=hash(readFileSync(footnoteFixture))
  const noteSource=JSON.parse(readFileSync(resolve(footnoteDir,'source.json'),'utf8'))
  const label=noteSource.document.notes.flatMap(n=>n.blocks).flatMap(b=>b.paragraph?.runs??[]).find(r=>r.reference?.role==='label').id
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
      if(init?.method==='POST'&&String(input).endsWith('/v1/docx/page-preview-textboxes')) {
        const value=await response.clone().json();window.__nativeDocxTextbox=value;
        if(window.__corruptTextbox&&value.preview){const box=value.preview.version===2?value.preview.textboxes.at(-1):value.preview.textbox;box.x_millipoints++;return new Response(JSON.stringify(value),{status:response.status,headers:response.headers});}
        if(window.__corruptStack&&value.preview?.textboxes?.[0]?.stacking){value.preview.textboxes[0].stacking.behind_doc=false;return new Response(JSON.stringify(value),{status:response.status,headers:response.headers});}
        if(window.__dropTextbox&&value.preview?.version===2){value.preview.textboxes.pop();return new Response(JSON.stringify(value),{status:response.status,headers:response.headers});}
      }
      if (init?.method === 'POST' && String(input).endsWith('/v1/docx/page-preview')) window.__nativeDocxPaint = (await response.clone().json()).page_paint_output;
      return response;
    };
  }` })
  await cdp.send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1100, deviceScaleFactor: 1, mobile: false })
  await cdp.send('Page.navigate', { url: `${server.url}#/docs?feature=editor` })
  await poll(() => evaluate(`Boolean(document.querySelector('[aria-label="Open a DOCX file"]'))`), 'DOCX workbench')
  await setFixtureData({textboxHash,footnoteHash,label})
  await upload(textboxFixture)
  await poll(()=>evaluate(`${native}?.textContent.includes('Nothing is uploaded')`),'textbox consent')
  await assert(`window.__nativeDocxPosts.length===0`,'opening textbox document uploads nothing')
  await click('Upload to helper and preview page-placed textboxes')
  await poll(()=>evaluate(`${native}?.querySelector('[data-native-textbox]')!==null&&${native}?.textContent.includes('1 approximate, read-only pages')`),'source rectangle page',45000)
  await assert(`window.__nativeDocxPosts.length===1&&window.__nativeDocxPosts[0].hash===window.__nativeDocxFixture.textboxHash`,'textbox upload preserves exact original bytes')
  await assert(`(() => {const box=${native}.querySelector('[data-native-textbox]'),rect=box.querySelector('rect');return box.getAttribute('transform')==='translate(72000 144000)'&&rect.getAttribute('width')==='216000'&&rect.getAttribute('height')==='72000'&&rect.getAttribute('stroke-width')==='1000'&&rect.getAttribute('fill')==='#FFF2CC'&&rect.getAttribute('stroke')==='#204060'})()`,'page offsets, bounds, colors and stroke match source')
  await assert(`${native}.querySelectorAll('svg').length===1&&${native}.querySelectorAll('svg path').length===window.__nativeDocxTextbox.preview.body_paint.pages[0].commands.filter(c=>c.kind==='fill_glyph_path').length+window.__nativeDocxTextbox.preview.textbox.paint.paths.length`,'all body and textbox glyph outlines share one mounted page')
  await assert(`${native}.textContent.includes(window.__nativeDocxTextbox.preview.source_diagnostics[0].message)&&${docs}?.dataset.demoDirty!=='true'`,'source drawing warning remains visible and source stays clean')
  await assert(`(async()=>{const svg=${native}.querySelector('svg').cloneNode(true);svg.setAttribute('width','612');svg.setAttribute('height','792');const image=new Image();image.src='data:image/svg+xml;base64,'+btoa(new XMLSerializer().serializeToString(svg));await image.decode();const canvas=document.createElement('canvas');canvas.width=612;canvas.height=792;const context=canvas.getContext('2d');context.drawImage(image,0,0,612,792);const pixel=context.getImageData(80,205,1,1).data;return pixel[0]===255&&pixel[1]===242&&pixel[2]===204&&pixel[3]===255})()`,'actual SVG raster fills the authored physical-page rectangle')
  await screenshot('docx-page-textbox.png')
  await evaluate('window.__corruptTextbox=true')
  await click('Upload to helper and preview page-placed textboxes')
  await poll(()=>evaluate(`${native}?.textContent.includes('Textbox placement does not fit')&&${native}?.querySelector('svg')===null`),'forged coordinate refusal',45000)
  await assert(`window.__nativeDocxPosts.length===2`,'failed preview requires a new explicit upload')
  await evaluate('window.__corruptTextbox=false')
  await upload(footnoteFixture)
  await poll(()=>evaluate(`${native}?.textContent.includes('Nothing is uploaded')&&${native}?.querySelector('svg')===null`),'footnote replacement clears prior pages')
  await assert(`window.__nativeDocxPosts.length===2`,'replacement does not reuse upload consent')
  await click('Upload to helper and render native pages')
  await poll(()=>evaluate(`${native}?.textContent.includes('4 native pages')&&${native}?.querySelector('svg path')!==null`),'footnote continuation pages',45000)
  await assert(`window.__nativeDocxPosts.length===3&&window.__nativeDocxPosts[2].hash===window.__nativeDocxFixture.footnoteHash`,'footnote upload preserves exact original bytes')
  await assert(`window.__nativeDocxPaint.pages.flatMap(p=>p.commands).filter(c=>c.kind==='fill_glyph_path'&&c.source_id===window.__nativeDocxFixture.label).length===1`,'continued footnote emits its source label exactly once')
  for(let ordinal=0;ordinal<4;ordinal++){
    if(ordinal){await click('Next native page');await poll(()=>evaluate(`${native}?.querySelector('svg[aria-label="Native document page ${ordinal+1}"]')!==null`),'footnote page '+(ordinal+1))}
    await assert(`(() => {const page=window.__nativeDocxPaint.pages[${ordinal}],rule=page.commands.find(c=>c.kind==='stroke_note_separator'),line=${native}.querySelector('svg line');return !!rule&&!!line&&rule.x2_millipoints-rule.x1_millipoints===${ordinal?468000:144000}&&Number(line.getAttribute('x2'))-Number(line.getAttribute('x1'))===${ordinal?468000:144000}&&${native}.querySelectorAll('svg').length===1&&page.lines.some(l=>l.region==='footnote')&&${ordinal?'page.lines.every(l=>l.region!=="body")':'page.lines.some(l=>l.region==="body")'}})()`,'source separator and note content on page '+(ordinal+1))
    await screenshot('docx-footnote-page-'+(ordinal+1)+'.png')
  }
  await assert(`${docs}?.dataset.demoDirty!=='true'&&window.__nativeDocxPosts.length===3`,'note navigation preserves original source without more uploads')
  const lineFixture=resolve(lineDir,'footnote-continuation.docx'),lineHash=hash(readFileSync(lineFixture))
  const lineSource=JSON.parse(readFileSync(resolve(lineDir,'source.json'),'utf8'))
  const lineNote=lineSource.document.notes.find(n=>n.note_role==='content'),lineParagraph=lineNote.blocks[0].id
  const lineLabel=lineNote.blocks[0].paragraph.runs.find(r=>r.reference?.role==='label').id
  await setFixtureData({lineHash,lineParagraph,lineLabel})
  await upload(lineFixture)
  await poll(()=>evaluate(`${native}?.textContent.includes('Nothing is uploaded')&&${native}?.querySelector('svg')===null`),'single-paragraph replacement')
  await assert(`window.__nativeDocxPosts.length===3`,'line continuation needs fresh upload consent')
  await click('Upload to helper and render native pages')
  await poll(()=>evaluate(`${native}?.textContent.includes('native pages')&&${native}?.querySelector('svg path')!==null`),'single-paragraph footnote paint',45000)
  const linePages=await evaluate('window.__nativeDocxPaint.pages.length')
  await assert(`window.__nativeDocxPaint.pages.length>=3&&window.__nativeDocxPosts.length===4&&window.__nativeDocxPosts[3].hash===window.__nativeDocxFixture.lineHash`,'one long footnote paragraph spans several pages without changing the package')
  await assert(`(() => {const slices=window.__nativeDocxPaint.pages.map(p=>p.lines.filter(l=>l.region==='footnote'&&l.paragraph_id===window.__nativeDocxFixture.lineParagraph)),lines=slices.flat();return slices.every(s=>s.length>=2)&&lines.every((l,i)=>l.source_line_ordinal===i)&&new Set(lines.map(l=>l.line_id)).size===lines.length})()`,'line continuation preserves exact source order and widow/orphan pairs')
  await assert(`window.__nativeDocxPaint.pages.flatMap(p=>p.commands).filter(c=>c.kind==='fill_glyph_path'&&c.source_id===window.__nativeDocxFixture.lineLabel).length===1`,'single-paragraph note label is painted once')
  for(let ordinal=1;ordinal<linePages;ordinal++){
    await click('Next native page')
    await poll(()=>evaluate(`${native}?.querySelector('svg[aria-label="Native document page ${ordinal+1}"]')!==null`),'line continuation page '+(ordinal+1))
  }
  await assert(`${native}.querySelectorAll('svg').length===1&&${native}.querySelector('svg path')!==null&&${docs}?.dataset.demoDirty!=='true'`,'final split paragraph page remains painted and source stays clean')
  await screenshot('docx-footnote-line-last.png')
  const flowFixture=resolve(flowDir,'footnote-continuation.docx'),flowHash=hash(readFileSync(flowFixture))
  const flowSource=JSON.parse(readFileSync(resolve(flowDir,'source.json'),'utf8')).document
  const flowBodyIDs=flowSource.body.blocks.map(b=>b.id)
  const flowNotes=flowSource.notes.filter(n=>n.note_role==='content').map(n=>({paragraphID:n.blocks[0].id,labelID:n.blocks[0].paragraph.runs.find(r=>r.reference?.role==='label').id,referenceParagraphID:flowSource.body.blocks.find(b=>b.paragraph.runs.some(r=>r.reference?.target_id===n.id)).id}))
  await setFixtureData({flowHash,flowBodyIDs,flowNotes})
  await upload(flowFixture)
  await poll(()=>evaluate(`${native}?.textContent.includes('Nothing is uploaded')&&${native}?.querySelector('svg')===null`),'shared-flow replacement')
  await assert('window.__nativeDocxPosts.length===4','shared-flow document needs new upload consent')
  await click('Upload to helper and render native pages')
  await poll(()=>evaluate(`${native}?.textContent.includes('native pages')&&${native}?.querySelector('svg path')!==null`),'shared body and note flow',45000)
  const flowPages=await evaluate('window.__nativeDocxPaint.pages.length')
  await assert('window.__nativeDocxPosts.length===5&&window.__nativeDocxPosts[4].hash===window.__nativeDocxFixture.flowHash','shared-flow upload preserves source bytes')
  await assert(`window.__nativeDocxPaint.pages.slice(1).some(p=>p.lines.some(l=>l.region==='body')&&p.lines.some(l=>l.region==='footnote')&&p.commands.some(c=>c.kind==='stroke_note_separator'&&c.x2_millipoints-c.x1_millipoints===468000))`,'continued note and later body text share a page')
  await assert(`JSON.stringify(window.__nativeDocxPaint.pages.flatMap(p=>p.lines.filter(l=>l.region==='body').map(l=>l.paragraph_id)))===JSON.stringify(window.__nativeDocxFixture.flowBodyIDs)`,'all later body paragraphs retain source order exactly once')
  await assert(`window.__nativeDocxFixture.flowNotes.every(n=>{const pages=window.__nativeDocxPaint.pages,lines=pages.flatMap(p=>p.lines.filter(l=>l.paragraph_id===n.paragraphID));return lines.length>0&&lines.every((l,i)=>l.source_line_ordinal===i)&&new Set(lines.map(l=>l.line_id)).size===lines.length&&pages.findIndex(p=>p.lines.some(l=>l.paragraph_id===n.paragraphID))===pages.findIndex(p=>p.lines.some(l=>l.paragraph_id===n.referenceParagraphID))&&pages.flatMap(p=>p.commands).filter(c=>c.kind==='fill_glyph_path'&&c.source_id===n.labelID).length===1})`,'all three notes start on their reference page and preserve source lines and labels')
  await assert(`window.__nativeDocxPaint.pages.every(p=>{const body=p.lines.filter(l=>l.region==='body'),notes=p.lines.filter(l=>l.region==='footnote');return !body.length||!notes.length||Math.max(...body.map(l=>l.y_millipoints+l.height_millipoints))<=Math.min(...notes.map(l=>l.y_millipoints))})`,'body and note regions never overlap')
  for(let ordinal=0;ordinal<flowPages;ordinal++){
    if(ordinal){await click('Next native page');await poll(()=>evaluate(`${native}?.querySelector('svg[aria-label="Native document page ${ordinal+1}"]')!==null`),'shared-flow page '+(ordinal+1))}
    await assert(`${native}.querySelectorAll('svg').length===1&&${native}.querySelectorAll('svg path').length===window.__nativeDocxPaint.pages[${ordinal}].commands.filter(c=>c.kind==='fill_glyph_path').length`,'all shared-flow glyphs mount on page '+(ordinal+1))
    if(ordinal===1||ordinal===flowPages-1)await screenshot('docx-footnote-flow-page-'+(ordinal+1)+'.png')
  }
  await assert(`${docs}?.dataset.demoDirty!=='true'&&window.__nativeDocxPosts.length===5`,'shared-flow navigation preserves source without further uploads')
  const multiFixture=resolve(multiDir,'page-textbox.docx'),multiHash=hash(readFileSync(multiFixture))
  const multiSource=JSON.parse(readFileSync(resolve(multiDir,'source.json'),'utf8'))
  await setFixtureData({multiHash,multiParagraphIDs:multiSource.textbox_geometry.items.map(i=>i.owner.paragraph_id),multiDiagnosticIDs:multiSource.textbox_geometry.items.map(i=>i.owner.diagnostic_id)})
  await upload(multiFixture)
  await poll(()=>evaluate(`${native}?.textContent.includes('Nothing is uploaded')&&${native}?.querySelector('svg')===null`),'multiple-textbox replacement')
  await assert('window.__nativeDocxPosts.length===5','multiple textboxes require fresh upload consent')
  await click('Upload to helper and preview page-placed textboxes')
  await poll(()=>evaluate(`${native}?.textContent.includes('2 approximate, read-only pages')&&${native}?.querySelector('[data-native-textbox]')!==null`),'multiple textbox pages',45000)
  const multiPages=await evaluate('window.__nativeDocxTextbox.preview.body_paint.pages.length')
  await assert('window.__nativeDocxPosts.length===6&&window.__nativeDocxPosts[5].hash===window.__nativeDocxFixture.multiHash&&window.__nativeDocxTextbox.preview.version===2&&window.__nativeDocxTextbox.preview.textboxes.length===2','two source rectangles bind the exact uploaded document')
  await assert('window.__nativeDocxFixture.multiDiagnosticIDs.every(id=>window.__nativeDocxTextbox.preview.source_diagnostics.some(d=>d.id===id))','every original drawing restriction remains in the multiple preview')
  for(let ordinal=0;ordinal<multiPages;ordinal++){
    if(ordinal){await click('Next approximate page');await poll(()=>evaluate(`${native}?.querySelector('svg[aria-label="Approximate document page ${ordinal+1}"]')!==null`),'multiple-textbox page '+(ordinal+1))}
    await assert(`(()=>{const result=window.__nativeDocxTextbox.preview,page=result.body_paint.pages[${ordinal}],box=result.textboxes[${ordinal}],mounted=${native}.querySelector('[data-native-textbox]');return ${native}.querySelectorAll('[data-native-textbox]').length===1&&mounted.getAttribute('data-native-textbox-id')===window.__nativeDocxFixture.multiDiagnosticIDs[${ordinal}]&&box.page_id===page.id&&page.lines.some(l=>l.paragraph_id===window.__nativeDocxFixture.multiParagraphIDs[${ordinal}])&&${native}.querySelectorAll('svg path').length===page.commands.filter(c=>c.kind==='fill_glyph_path').length+box.paint.paths.length})()`,'only the source-bound textbox and all glyphs mount on page '+(ordinal+1))
    await assert(`(()=>{const box=${native}.querySelector('[data-native-textbox]'),rect=box.querySelector('rect');return box.getAttribute('transform')==='translate(${ordinal?288000:72000} ${ordinal?216000:144000})'&&rect.getAttribute('fill')==='${ordinal?'#DDEEFF':'#FFF2CC'}'})()`,'authored coordinates and color on textbox page '+(ordinal+1))
    await screenshot('docx-multiple-textbox-page-'+(ordinal+1)+'.png')
  }
  await evaluate('window.__corruptTextbox=true')
  await click('Upload to helper and preview page-placed textboxes')
  await poll(()=>evaluate(`${native}?.textContent.includes('Textbox placement does not fit')&&${native}?.querySelector('svg')===null`),'corrupt second textbox refusal',45000)
  await assert('window.__nativeDocxPosts.length===7','a forged second textbox clears the entire preview')
  await evaluate('window.__corruptTextbox=false;window.__dropTextbox=true')
  await click('Upload to helper and preview page-placed textboxes')
  await poll(()=>evaluate(`${native}?.textContent.includes('cover every source rectangle')&&${native}?.querySelector('svg')===null`),'missing second textbox refusal',45000)
  await assert('window.__nativeDocxPosts.length===8','an omitted textbox cannot produce a partial preview')
  await evaluate('window.__dropTextbox=false')
  await assert(`${docs}?.dataset.demoDirty!=='true'`,'multiple-textbox rendering and refusals preserve source')
  const positionFixture=resolve(positionDir,'page-textbox.docx'),positionHash=hash(readFileSync(positionFixture))
  await setFixtureData({positionHash})
  await upload(positionFixture)
  await poll(()=>evaluate(`${native}?.textContent.includes('Nothing is uploaded')&&${native}?.querySelector('svg')===null`),'positioned-textbox replacement')
  await assert('window.__nativeDocxPosts.length===8','relative textbox positions require a fresh explicit upload')
  await click('Upload to helper and preview page-placed textboxes')
  await poll(()=>evaluate(`${native}?.textContent.includes('2 approximate, read-only pages')&&${native}?.querySelector('[data-native-textbox]')!==null`),'relative textbox pages',45000)
  await assert('window.__nativeDocxPosts.length===9&&window.__nativeDocxPosts[8].hash===window.__nativeDocxFixture.positionHash&&window.__nativeDocxTextbox.preview.version===2','relative positions bind exact source bytes')
  for(let ordinal=0;ordinal<2;ordinal++){
    if(ordinal){await click('Next approximate page');await poll(()=>evaluate(`${native}?.querySelector('svg[aria-label="Approximate document page 2"]')!==null`),'relative textbox page 2')}
    const x=ordinal?36000:198000,y=ordinal?144000:360000
    await assert(`(()=>{const preview=window.__nativeDocxTextbox.preview,box=preview.textboxes[${ordinal}],page=preview.body_paint.pages[${ordinal}],mounted=${native}.querySelector('[data-native-textbox]');return box.x_millipoints===${x}&&box.y_millipoints===${y}&&box.page_id===page.id&&mounted.getAttribute('transform')==='translate(${x} ${y})'&&${native}.querySelectorAll('[data-native-textbox]').length===1&&getComputedStyle(${native}.querySelector('svg')).overflow==='hidden'})()`,'source alignment and signed offsets on page '+(ordinal+1))
    await screenshot('docx-relative-textbox-page-'+(ordinal+1)+'.png')
  }
  await evaluate('window.__corruptTextbox=true')
  await click('Upload to helper and preview page-placed textboxes')
  await poll(()=>evaluate(`${native}?.textContent.includes('Textbox placement does not fit')&&${native}?.querySelector('svg')===null`),'forged relative textbox refusal',45000)
  await assert(`window.__nativeDocxPosts.length===10&&${docs}?.dataset.demoDirty!=='true'`,'relative coordinate forgery rejects the preview and preserves source')
  await evaluate('window.__corruptTextbox=false')
  const stackFixture=resolve(stackDir,'page-textbox.docx'),stackHash=hash(readFileSync(stackFixture))
  await setFixtureData({stackHash})
  await upload(stackFixture)
  await poll(()=>evaluate(`${native}?.textContent.includes('Nothing is uploaded')&&${native}?.querySelector('svg')===null`),'stacked source replacement')
  await assert('window.__nativeDocxPosts.length===10','stacked preview requires its own upload')
  await click('Upload to helper and preview page-placed textboxes')
  await poll(()=>evaluate(`${native}?.querySelectorAll('[data-native-textbox]').length===4`),'stacked textbox page',45000)
  await assert('window.__nativeDocxPosts.length===11&&window.__nativeDocxPosts[10].hash===window.__nativeDocxFixture.stackHash&&window.__nativeDocxTextbox.preview.source_diagnostics.length===4','stacked composition preserves uploaded bytes and all source restrictions')
  await assert(`(()=>{const svg=${native}.querySelector('svg'),boxes=[...svg.querySelectorAll('[data-native-textbox]')],p=window.__nativeDocxTextbox.preview,ids=p.textboxes.map(t=>t.paint.diagnostic_id),body=svg.querySelector(':scope > path');return JSON.stringify(boxes.map(b=>b.getAttribute('data-native-textbox-id')))===JSON.stringify([ids[0],ids[2],ids[1],ids[3]])&&!!(boxes[0].compareDocumentPosition(body)&4)&&!!(body.compareDocumentPosition(boxes[1])&4)&&svg.querySelectorAll('path').length===p.body_paint.pages[0].commands.filter(c=>c.kind==='fill_glyph_path').length+60})()`,'body lies between source-ranked behind and foreground textboxes, with stable equal-rank order')
  await assert(`(async()=>{const svg=${native}.querySelector('svg').cloneNode(true);svg.removeAttribute('style');svg.setAttribute('width','612');svg.setAttribute('height','792');const image=new Image();image.src='data:image/svg+xml;base64,'+btoa(new XMLSerializer().serializeToString(svg));await image.decode();const canvas=document.createElement('canvas');canvas.width=612;canvas.height=792;const ctx=canvas.getContext('2d');ctx.drawImage(image,0,0);return [[280,120,'255,221,238'],[280,160,'221,238,255'],[280,200,'255,242,204']].every(([x,y,rgb])=>Array.from(ctx.getImageData(x,y,1,1).data).slice(0,3).join(',')===rgb)})()`,'raster pixels confirm authored stacking in overlapping rectangles')
  await screenshot('docx-stacked-textboxes.png')
  await evaluate('window.__corruptStack=true')
  await click('Upload to helper and preview page-placed textboxes')
  await poll(()=>evaluate(`${native}?.textContent.includes('Textbox stacking does not match source')&&${native}?.querySelector('svg')===null`),'forged textbox layer refusal',45000)
  await assert(`window.__nativeDocxPosts.length===12&&${docs}?.dataset.demoDirty!=='true'`,'forged stacking refuses all pages and preserves source')
  await evaluate('window.__corruptStack=false')
  const parityFixture=resolve(parityDir,'page-textbox.docx'),parityHash=hash(readFileSync(parityFixture))
  await setFixtureData({parityHash})
  await upload(parityFixture)
  await poll(()=>evaluate(`${native}?.textContent.includes('Nothing is uploaded')&&${native}?.querySelector('svg')===null`),'parity-textbox replacement')
  await assert('window.__nativeDocxPosts.length===12','parity textbox positions require a fresh explicit upload')
  await click('Upload to helper and preview page-placed textboxes')
  await poll(()=>evaluate(`${native}?.textContent.includes('2 approximate, read-only pages')&&${native}?.querySelector('[data-native-textbox]')!==null`),'parity textbox pages',45000)
  await assert('window.__nativeDocxPosts.length===13&&window.__nativeDocxPosts[12].hash===window.__nativeDocxFixture.parityHash&&window.__nativeDocxTextbox.preview.version===2','parity positions bind exact source bytes')
  for(let ordinal=0;ordinal<2;ordinal++){
    if(ordinal){await click('Next approximate page');await poll(()=>evaluate(`${native}?.querySelector('svg[aria-label="Approximate document page 2"]')!==null`),'parity textbox page 2')}
    const x=ordinal?552000:12000,y=ordinal?738000:18000
    await assert(`(()=>{const preview=window.__nativeDocxTextbox.preview,box=preview.textboxes[${ordinal}],page=preview.body_paint.pages[${ordinal}],mounted=${native}.querySelector('[data-native-textbox]');return box.x_millipoints===${x}&&box.y_millipoints===${y}&&box.page_id===page.id&&mounted.getAttribute('transform')==='translate(${x} ${y})'&&${native}.querySelectorAll('[data-native-textbox]').length===1&&getComputedStyle(${native}.querySelector('svg')).overflow==='hidden'})()`,'inside margin parity coordinates on page '+(ordinal+1))
    await screenshot('docx-parity-textbox-page-'+(ordinal+1)+'.png')
  }
  await evaluate('window.__corruptTextbox=true')
  await click('Upload to helper and preview page-placed textboxes')
  await poll(()=>evaluate(`${native}?.textContent.includes('Textbox placement does not fit')&&${native}?.querySelector('svg')===null`),'forged parity textbox refusal',45000)
  await assert(`window.__nativeDocxPosts.length===14&&${docs}?.dataset.demoDirty!=='true'`,'parity coordinate forgery rejects the preview and preserves source')
  if(hash(readFileSync(parityFixture))!==parityHash)throw Error('Parity source changed')
  if(hash(readFileSync(stackFixture))!==stackHash)throw Error('Stacked-textbox source changed')
  if(hash(readFileSync(positionFixture))!==positionHash)throw Error('Relative-textbox source changed')
  if(hash(readFileSync(multiFixture))!==multiHash)throw Error('Multiple-textbox source changed')
  if(hash(readFileSync(flowFixture))!==flowHash)throw Error('Shared-flow source changed')
  if(hash(readFileSync(lineFixture))!==lineHash)throw Error('Line continuation source changed')
  if(hash(readFileSync(textboxFixture))!==textboxHash||hash(readFileSync(footnoteFixture))!==footnoteHash)throw Error('Source fixture changed')
  if(errors.length)throw Error(errors.join('\n'))
  writeFileSync(resolve(artifacts,'summary.json'),JSON.stringify({status:'passed',cases:checks,textboxPages:1,footnotePages:4,lineContinuationPages:linePages,sharedFlowPages:flowPages,multipleTextboxPages:multiPages,relativeTextboxPages:2,stackedTextboxes:4,parityTextboxPages:2,uploads:14},null,2))
  console.log(`DOCX textbox and footnote browser checks passed: ${checks} cases`)
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
// Fixture identifiers travel as CDP data, never as generated JavaScript source.
async function setFixtureData(data) {
  const handle=await cdp.send('Runtime.evaluate',{expression:'globalThis'})
  try {
    const result=await cdp.send('Runtime.callFunctionOn',{
      objectId:handle.result.objectId,
      functionDeclaration:'function (data) { window.__nativeDocxFixture = {...window.__nativeDocxFixture, ...data}; }',
      arguments:[{value:data}],returnByValue:true,
    })
    if(result.exceptionDetails)throw new Error(result.exceptionDetails.exception?.description??result.exceptionDetails.text)
  } finally { await cdp.send('Runtime.releaseObject',{objectId:handle.result.objectId}) }
}
async function evaluate(expression) {
  const value = await cdp.send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true })
  if (value.exceptionDetails) throw new Error(value.exceptionDetails.exception?.description ?? value.exceptionDetails.text)
  return value.result.value
}
async function assert(expression, label) { if (!await evaluate(expression)) throw new Error(`Failed: ${label}`);checks++ }
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
  await evaluate(`${native}?.querySelector('svg')?.scrollIntoView({ block: 'center' })`)
  const raster=await evaluate(`(async()=>{const svg=${native}.querySelector('svg').cloneNode(true);svg.removeAttribute('style');svg.setAttribute('width','816');svg.setAttribute('height','1056');const image=new Image();image.src='data:image/svg+xml;base64,'+btoa(new XMLSerializer().serializeToString(svg));await image.decode();const canvas=document.createElement('canvas');canvas.width=816;canvas.height=1056;const context=canvas.getContext('2d');context.fillStyle='#fff';context.fillRect(0,0,816,1056);context.drawImage(image,0,0,816,1056);return canvas.toDataURL('image/png').split(',')[1]})()`)
  writeFileSync(resolve(artifacts,name.replace('.png','-page.png')),Buffer.from(raster,'base64'))
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

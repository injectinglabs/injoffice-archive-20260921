import { describe, expect, it } from 'vitest'
import type { NativeAsset, NativeElement, NativePptxDeck } from '@injoffice/pptx-native'
import { previewColor, previewImage, previewIssue, previewSlideSize } from './pptxPreview'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import PptxFilePreview,{PptxFilePreviewVector} from './components/PptxFilePreview'
import {compileFilePreviewGeometry} from './filePreviewGeometry'

const asset = (overrides: Partial<NativeAsset> = {}): NativeAsset => ({ id: 'image', provenance: 'parsed', contentType: 'image/png', sha256: 'a'.repeat(64), byteLength: 3, dataBase64: 'YWJj', passthrough: [], ...overrides })
const element = (overrides: object): NativeElement => ({ kind: 'picture', id: 'picture', provenance: 'parsed', assetId: 'image', transform: { x: 0, y: 0, cx: 100, cy: 100 }, compatibility: { status: 'editable', diagnostics: [] }, passthrough: [], ...overrides } as NativeElement)

describe('PPTX approximate preview policy', () => {
  it('accepts only literal RGB colors, not document-controlled CSS', () => {
    expect(previewColor('aBcDeF')).toBe('#aBcDeF')
    expect(previewColor('url(https://example.com/image)')).toBe('#ffffff')
    expect(previewColor(undefined, 'none')).toBe('none')
  })
  it('embeds bounded raster bytes and refuses active or remote resources', () => {
    expect(previewImage(asset())).toBe('data:image/png;base64,YWJj')
    expect(previewImage(asset({ contentType: 'image/jpeg' }))).toBe('data:image/jpeg;base64,YWJj')
    expect(previewImage(asset({ contentType: 'image/svg+xml' }))).toBeUndefined()
    expect(previewImage(asset({ dataBase64: 'https://example.com/image' }))).toBeUndefined()
    expect(previewImage(asset({ dataBase64: 'a'.repeat(4_000_001) }))).toBeUndefined()
    expect(previewImage(undefined)).toBeUndefined()
  })
  it('keeps missing content explicit and never relaxes refused object policy', () => {
    expect(previewIssue(element({}), [asset()])).toBeUndefined()
    expect(previewIssue(element({ compatibility: { status: 'preserveOnly', diagnostics: [{ severity: 'warning', code: 'pptx.picture-crop-unavailable', message: 'outset crop' }] } }), [asset()])).toContain('crop cannot be previewed')
    expect(previewIssue(element({}), [])).toContain('unavailable')
    expect(previewIssue(element({ kind: 'chart', chart: {} }), [])).toContain('no embedded preview')
    expect(previewIssue(element({ kind: 'group', children: [] }), [])).toBeUndefined()
    expect(previewIssue(element({ compatibility: { status: 'refused', diagnostics: [] } }), [asset()])).toContain('Unsupported')
  })
  it('preserves slide aspect ratio and rejects invalid or excessive dimensions', () => {
    const deck = (cx: number, cy: number) => ({ size: { cx, cy } } as NativePptxDeck)
    expect(previewSlideSize(deck(12_192_000, 6_858_000))).toMatchObject({ width: 960, height: 540 })
    for (const [cx, cy] of [[0, 1], [Infinity, 1], [1, NaN], [1, 1000], [-1, 1]]) expect(previewSlideSize(deck(cx, cy))).toBeUndefined()
  })
  it('keeps the fidelity label while asynchronous geometry is prepared',()=>{
    const deck=validDeck([]),html=renderToStaticMarkup(createElement(PptxFilePreview,{deck}))
    expect(html).toContain('Approximate file preview');expect(html).toContain('Not PowerPoint-equivalent rendering');expect(html).toContain('Preparing local geometry preview')
  })
  it('renders original text escaped and applies insets once inside the compiled text rectangle',async()=>{
    const text:NativeElement={kind:'text',id:'text',provenance:'authored',transform:{x:0,y:0,cx:1270000,cy:635000},paragraphs:[{runs:[{text:'<script>alert(1)</script>',fontSizeHundredthPt:1200}]}],textBody:{leftInsetEmu:12700,rightInsetEmu:25400,topInsetEmu:38100,bottomInsetEmu:50800,wrap:'none',verticalAnchor:'top',autoFit:'none',horizontalOverflow:'overflow',verticalOverflow:'overflow'},compatibility:{status:'editable',diagnostics:[]},passthrough:[]}
    const deck=validDeck([text]),before=JSON.stringify(deck),geometry=await compileFilePreviewGeometry(deck,0),html=renderToStaticMarkup(createElement(PptxFilePreviewVector,{deck,geometry}))
    expect(html).toContain('viewBox="0 0 960 540"');expect(html).toContain('&lt;script&gt;');expect(html).not.toContain('<script>');expect(html).toContain('white-space:pre;');expect(html).toContain('overflow:visible');expect(html).toContain('x="1" y="3" width="97" height="43"');expect(html).not.toContain('overflow-wrap:anywhere');expect(JSON.stringify(deck)).toBe(before)
  })
  it('crops the original raster through an unchanged percentage viewport',async()=>{
    const picture:NativeElement={kind:'picture',id:'picture',provenance:'authored',assetId:'image',transform:{x:12700,y:25400,cx:1270000,cy:635000},crop:{left:12500,top:25000,right:37500,bottom:0},compatibility:{status:'editable',diagnostics:[]},passthrough:[]}
    const deck=validDeck([picture]);deck.assets=[asset({provenance:'authored',sha256:'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad'})]
    const before=JSON.stringify(deck),geometry=await compileFilePreviewGeometry(deck,0),html=renderToStaticMarkup(createElement(PptxFilePreviewVector,{deck,geometry}))
    expect(html).toContain('viewBox="12500 25000 50000 75000"');expect(html).toContain('overflow="hidden"');expect(html).toContain('width="100000" height="100000" preserveAspectRatio="none"');expect(html).toContain('href="data:image/png;base64,YWJj"');expect(html).toContain('width="100" height="50"');expect(JSON.stringify(deck)).toBe(before)
  })
  it('preserves authored bullets, font styles and paragraph offsets in the approximate overlay',async()=>{
    const text:NativeElement={kind:'text',id:'text',provenance:'authored',transform:{x:0,y:0,cx:1270000,cy:635000},paragraphs:[{bullet:true,bulletCharacter:'▪',marginLeftEmu:152400,indentEmu:-50800,runs:[{text:'Author list',fontFamily:'DejaVu Sans',fontSizeHundredthPt:1200,color:'2F6FED',bold:true,italic:true}]}],compatibility:{status:'editable',diagnostics:[]},passthrough:[]}
    const deck=validDeck([text]),geometry=await compileFilePreviewGeometry(deck,0),html=renderToStaticMarkup(createElement(PptxFilePreviewVector,{deck,geometry}))
    expect(html).toContain('▪ ');expect(html).not.toContain('• ');expect(html).toContain('padding-left:12px');expect(html).toContain('text-indent:-4px');expect(html).toContain('font-family:DejaVu Sans');expect(html).toContain('font-weight:700');expect(html).toContain('font-style:italic')
  })
  it('refuses malformed source instead of hiding it through the geometry projection',async()=>{
    const deck=validDeck([{kind:'shape',id:'invalid',provenance:'authored',transform:{x:0,y:0,cx:0,cy:100},preset:'rect',paragraphs:[],passthrough:[],compatibility:{status:'editable',diagnostics:[]}}])
    await expect(compileFilePreviewGeometry(deck,0)).rejects.toThrow('validated')
  })
})
function validDeck(elements:NativeElement[]):NativePptxDeck{return {contractVersion:'pptx-native/v1',documentId:'file-preview',origin:'authored',size:{cx:12192000,cy:6858000},assets:[],compatibility:{status:'editable',diagnostics:[]},slides:[{id:'slide',provenance:'authored',passthrough:[],compatibility:{status:'editable',diagnostics:[]},elements}]}}

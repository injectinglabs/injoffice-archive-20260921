import {describe,it,expect} from 'vitest'
import {createElement} from 'react'
import {renderToStaticMarkup} from 'react-dom/server'
import type {NativeElement,NativePptxDeck} from '@injoffice/pptx-native'
import type {PptxPreview,PreviewNode} from '../../pptx-page-paint-worker/src/contract'
import {nativePptxPreviewStatus} from './nativePptxPreviewStatus'
import {NativePptxPreviewResult} from './components/NativePptxSlides'

const hash='a'.repeat(64)
const rectangle:PreviewNode={kind:'rect',rect:{x:10,y:10,cx:30,cy:20},radius:0,fill:'00AA00'}
const placeholder:PreviewNode={kind:'placeholder',rect:{x:60,y:10,cx:30,cy:20},label:'Unsupported text'}
function preview(nodes:PreviewNode[]=[rectangle]):PptxPreview{return {version:1,package_sha256:hash,slide_index:0,slide_count:1,width:100,height:60,background:'FFFFFF',policy:'max-run-natural-v1',nodes,diagnostics:[],font_digests:[],resources:[]}}
function source():NativePptxDeck{return {contractVersion:'pptx-native/v1',documentId:'deck',origin:'parsed',sourceRevision:`rev-${hash}`,size:{cx:100,cy:60},assets:[],compatibility:{status:'preserveOnly',diagnostics:[]},slides:[{id:'slide-1',provenance:'parsed',elements:[],passthrough:[],compatibility:{status:'preserveOnly',diagnostics:[]}}]}}
function refused():NativeElement{return {kind:'text',id:'text-2',name:'Budget notes',provenance:'parsed',transform:{x:60,y:10,cx:30,cy:20},paragraphs:[],passthrough:[],compatibility:{status:'refused',diagnostics:[{severity:'refusal',code:'pptx.text-layout-unavailable',message:'Shape autofit is not supported',scope:{slideId:'slide-1',elementId:'text-2'}}]}}}

describe('native PPTX preview coverage reporting',()=>{
 it('shows source and selected fonts outside details and never calls substitution complete',()=>{
  const result:PptxPreview={...preview(),font_substitutions:[{source_id:'text',paragraph_index:0,run_index:0,source_family:'Authored',selected_family:'Selected',face_id:'font-0',font_digest:`sha256:${hash}`}],font_substitution_policy:'explicit-whole-run-font-substitution-v1',font_substitution_policy_sha256:`sha256:${hash}`}
  expect(nativePptxPreviewStatus(result).status).toBe('partial')
  const html=renderToStaticMarkup(createElement(NativePptxPreviewResult,{preview:result}))
  expect(html).toContain('Authored → Selected');expect(html).toContain('Font substitution preview');expect(html.indexOf('Font substitution preview')).toBeLessThan(html.indexOf('<details'))
 })
 it('keeps inherited text approximation visible and partial',()=>{
  const result={...preview(),inherited_text_preview_count:1,inherited_text_policy:'source-latin-inheritance-approximate-v1' as const}
  expect(nativePptxPreviewStatus(result,source()).status).toBe('partial')
  expect(nativePptxPreviewStatus(result,source()).reasons.join(' ')).not.toContain('source-frame autofit')
  const html=renderToStaticMarkup(createElement(NativePptxPreviewResult,{preview:result,source:source()}))
  expect(html).toContain('data-inherited-text-approximation');expect(html).toContain('Kerning is disabled');expect(html.indexOf('Approximate inherited text preview.')).toBeLessThan(html.indexOf('<details'))
 })
 it('keeps the source-frame warning visible outside collapsed diagnostics and never calls it complete',()=>{
  const result={...preview(),source_frame_autofit_count:1}
  expect(nativePptxPreviewStatus(result,source()).status).toBe('partial')
  const html=renderToStaticMarkup(createElement(NativePptxPreviewResult,{preview:result,source:source()}))
  expect(html).toContain('data-autofit-approximation')
  expect(html).toContain('Approximate autofit preview.')
  expect(html.indexOf('Approximate autofit preview.')).toBeLessThan(html.indexOf('<details'))
  expect(html).toContain('without resizing')
  expect(html).toContain('editing permissions are unchanged')
 })
 it('keeps supported paint visible while identifying refused objects and placeholder regions separately',()=>{
  const deck=source();deck.slides[0]!.elements=[refused()]
  deck.compatibility.diagnostics=[...refused().compatibility.diagnostics]
  const result=preview([rectangle,{kind:'group',transform:[1,0,0,1,0,0],children:[placeholder]}])
  const before=JSON.stringify({deck,result})
  const status=nativePptxPreviewStatus(result,deck)
  expect(status).toMatchObject({status:'partial',sourceBound:true,paintPrimitives:1,glyphRecords:0,placeholderRegions:1,knownRefusedObjects:1})
  expect(status.reasons.filter(reason=>reason.includes('Budget notes'))).toEqual(['Budget notes: Shape autofit is not supported'])
  const html=renderToStaticMarkup(createElement(NativePptxPreviewResult,{preview:result,source:deck}))
  expect(html).toContain('Partial native preview')
  expect(html).toContain('fill="#00AA00"')
  expect(html).toContain('Budget notes: Shape autofit is not supported')
  expect(html).toContain('1 known refused source objects')
  expect(html).toContain('data-native-preview-status="partial"')
  expect(JSON.stringify({deck,result})).toBe(before)
 })
 it('never reports a placeholder-only or empty/transparent response as rendered',()=>{
  for(const nodes of [[],[placeholder],[{...rectangle,fill:'none'} as PreviewNode],[{kind:'path',d:'',fill:'000000'} as PreviewNode]]){
   expect(nativePptxPreviewStatus(preview(nodes),source()).status).toBe('unavailable')
  }
 })
 it('does not equate zero text with failure on an image/shape-only slide or absence of gaps with full fidelity',()=>{
  const status=nativePptxPreviewStatus(preview(),source())
  expect(status).toMatchObject({status:'unverified',label:'Native preview — coverage unverified',glyphRecords:0})
  expect(status.label).not.toContain('Complete')
 })
 it('does not turn preserve-only checking/policy warnings into missing content',()=>{
  const deck=source();deck.compatibility.diagnostics=[{severity:'warning',code:'pptx.text-checking-metadata-preserved',message:'Checking flags preserved'}]
  const result=preview();result.diagnostics=['native.compatibility: Checking flags preserved','render.preserveOnly: preview remains read-only','text.deterministicLayout: uses explicit policy']
  expect(nativePptxPreviewStatus(result,deck)).toMatchObject({status:'unverified',knownRefusedObjects:0,reasons:[]})
 })
 it('refuses to join stale source metadata or a different slide inventory',()=>{
  for(const change of [(deck:NativePptxDeck)=>{deck.sourceRevision=`rev-${'b'.repeat(64)}`},(deck:NativePptxDeck)=>{deck.slides.push({...deck.slides[0]!,id:'other'})}]){
   const deck=source();deck.slides[0]!.elements=[refused()];change(deck)
   const status=nativePptxPreviewStatus(preview(),deck)
   expect(status).toMatchObject({sourceBound:false,knownRefusedObjects:0,status:'unverified'})
   expect(status.reasons.join(' ')).toContain('could not be joined')
  }
 })
 it('ignores other-slide diagnostics and surfaces runtime text omissions without assigning false object counts',()=>{
  const deck=source();deck.compatibility.diagnostics=[{severity:'refusal',code:'pptx.text-refused',message:'Other slide',scope:{slideId:'other'}}]
  expect(nativePptxPreviewStatus(preview(),deck).status).toBe('unverified')
  const result=preview();result.diagnostics=['text.wrapUnavailable: line wrapping unavailable']
  expect(nativePptxPreviewStatus(result,deck)).toMatchObject({status:'partial',knownRefusedObjects:0,placeholderRegions:0})
 })
 it('counts nested text glyph records independently from source objects',()=>{
  const result=preview([{kind:'group',sourceRole:'contentRun',transform:[1,0,0,1,0,0],children:[{kind:'path',d:'M0 0 L1 1 Z',fill:'000000'}]}])
  expect(nativePptxPreviewStatus(result,source())).toMatchObject({glyphRecords:1,paintPrimitives:1,knownRefusedObjects:0,status:'unverified'})
 })
 it('bounds displayed coverage and diagnostic rows without hiding their full counts',()=>{
  const result=preview();result.diagnostics=Array.from({length:75},(_,i)=>`text.refused: reason ${i}`)
  const html=renderToStaticMarkup(createElement(NativePptxPreviewResult,{preview:result,source:source()}))
  expect(html).toContain('coverage limits (75)')
  expect(html).toContain('Native diagnostics (75)')
  expect(html).toContain('25 additional coverage reasons are not displayed')
  expect(html).toContain('25 additional diagnostics are not displayed')
  expect(html.match(/<li>/g)).toHaveLength(100)
  expect(html).not.toContain('reason 74')
 })
})

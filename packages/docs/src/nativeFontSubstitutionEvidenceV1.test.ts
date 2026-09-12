import {describe,it,expect} from 'vitest'
import {nativeDocxFontSubstitutionDiagnosticV1,qualifyNativeDocxFontSubstitutionsV1} from './nativeFontSubstitutionEvidenceV1.js'

// Original synthetic resolver/shaper projections, not external Office exports.
function fixture(role:'run'|'paragraph-mark'|'list-marker'='run'){
 const digest=`sha256:${'a'.repeat(64)}`
 const record={source_id:role==='run'?'r1':'p1',source_role:role,source_family:'Missing',selected_family:'Selected',face_id:'face1',font_digest:digest,weight:400,style:'normal' as const}
 const props={font_family:'Missing',font_size_half_points:20}
 const resolved:any={document_id:'d1',revision:'rev1',diagnostics:[],paragraphs:[{paragraph_id:'p1',properties:{},paragraph_mark_properties:props,...(role==='list-marker'?{numbering:{format:'decimal',resolved_text:'1.',marker_properties:props}}:{})}],runs:[{run_id:'r1',paragraph_id:'p1',properties:props}]}
 const document:any={document_id:'d1',revision:'rev1',body:{blocks:[{paragraph:{id:'p1',runs:[{id:'r1',kind:'text',text:'Actual source'}]}}]},headers:[],footers:[],notes:[],comment_stories:[]}
 const fragment={source_id:record.source_id,source_kind:role,face_id:'face1',direction:'ltr',script:'Latn'}
 const shaped:any={document_id:'d1',revision:'rev1',paragraphs:[{paragraph_id:'p1',lines:[{fragments:role==='paragraph-mark'?[]:[fragment]}]}],font_substitutions:[record],diagnostics:[nativeDocxFontSubstitutionDiagnosticV1(record)]}
 const manifest:any={version:1,manifestId:'m1',revision:'rev1',faces:[{faceId:'face1',family:'Selected',weight:400,style:'normal',stretch:100,source:{kind:'host',resourceId:'font1',contentDigest:digest}}],fallbackChains:[]}
 const policy={version:1,mappings:[{sourceFamily:'Missing',targetFamily:'Selected',weight:400,style:'normal'}]}
 return {shaped,resolved,document,manifest,policy,record}
}
const qualify=(v:ReturnType<typeof fixture>)=>qualifyNativeDocxFontSubstitutionsV1(v.shaped,v.resolved,v.manifest,v.policy,v.document)
describe('source-bound DOCX substitution qualification',()=>{
 it.each(['run','paragraph-mark','list-marker'] as const)('joins %s actual source, selected bytes and derived diagnostic',role=>{
  const v=fixture(role),before=structuredClone(v)
  expect(qualify(v)).toEqual([v.record]);expect(v).toEqual(before)
 })
 it.each(['run','paragraph-mark','list-marker'] as const)('does not accept missing %s evidence',role=>{
  const v=fixture(role);v.shaped.font_substitutions=[];v.shaped.diagnostics=[]
  expect(()=>qualify(v)).toThrow(/evidence/)
 })
 it.each([
  (v:ReturnType<typeof fixture>)=>{v.document.body.blocks[0].paragraph.runs[0].text='漢'},
  (v:ReturnType<typeof fixture>)=>{v.document.body.blocks[0].paragraph.runs[0].text='א'},
  (v:ReturnType<typeof fixture>)=>{v.resolved.runs[0].properties.rtl=true},
  (v:ReturnType<typeof fixture>)=>{v.resolved.paragraphs[0].properties.bidi=true},
  (v:ReturnType<typeof fixture>)=>{v.shaped.paragraphs[0].lines[0].fragments[0].script='Hani'},
  (v:ReturnType<typeof fixture>)=>{v.shaped.paragraphs[0].lines[0].fragments[0].direction='rtl'},
  (v:ReturnType<typeof fixture>)=>{v.document.revision='stale'},
  (v:ReturnType<typeof fixture>)=>{v.shaped.font_substitutions[0].font_digest=`sha256:${'b'.repeat(64)}`},
  (v:ReturnType<typeof fixture>)=>{v.shaped.font_substitutions.push(structuredClone(v.record))},
  (v:ReturnType<typeof fixture>)=>{v.shaped.diagnostics[0].message+=' forged'},
  (v:ReturnType<typeof fixture>)=>{v.resolved.diagnostics.push({code:'FONT_MATCHING_METADATA_PRESERVED'})},
 ])('refuses forged or unsupported source evidence %#',mutate=>{const v=fixture();mutate(v);expect(()=>qualify(v)).toThrow()})
 it('keeps authored symbol bullets exact',()=>{const v=fixture('list-marker');v.resolved.paragraphs[0].numbering.format='bullet';expect(()=>qualify(v)).toThrow(/Latin|eligible/)})
})

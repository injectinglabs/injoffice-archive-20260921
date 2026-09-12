import {preflightWire,canonicalWireSha256} from './nativePagePaintWireV1.js'
import type {NativeDocxShapingDiagnosticV1,NativeDocxShapedLinesV1} from './nativeShapingLines.js'
import type {NativeDocxResolvedLayoutInputV1} from './nativeResolvedLayout.js'
import type {NativeDocxDocumentV1,NativeDocxParagraphV1} from './nativeContract.js'
import {decodeExplicitFontPolicyV1,selectExplicitFontV1,type NativeFontManifest,type TextRunInput} from '@injoffice/font-metrics/layout'

export interface NativeDocxFontSubstitutionV1 {
 source_id:string
 source_role:'run'|'list-marker'|'paragraph-mark'
 source_family:string
 selected_family:string
 face_id:string
 font_digest:string
 weight:number
 style:'normal'|'italic'
}
export const DOCX_FONT_SUBSTITUTION_CODE='font-substitution-approximate' as const
export function nativeDocxFontSubstitutionDiagnosticV1(record:NativeDocxFontSubstitutionV1):NativeDocxShapingDiagnosticV1{
 return {code:DOCX_FONT_SUBSTITUTION_CODE,severity:'deferred',scope_id:record.source_id,source_id:record.source_id,message:`Read-only substituted ${record.source_role}: ${record.source_family} -> ${record.selected_family}; evidence ${canonicalWireSha256(record)}. Font metrics and layout may differ.`}
}
export function decodeNativeDocxFontSubstitutionsV1(value:unknown):NativeDocxFontSubstitutionV1[]{
 if(preflightWire(value,'font substitutions',100_000,10_000).length||!Array.isArray(value)||value.length>10000)throw new TypeError('Invalid bounded DOCX font substitution evidence')
 const records=structuredClone(value) as NativeDocxFontSubstitutionV1[],seen=new Set<string>()
 for(const r of records){
  if(!r||Object.keys(r).sort().join(',')!=='face_id,font_digest,selected_family,source_family,source_id,source_role,style,weight'||!['run','list-marker','paragraph-mark'].includes(r.source_role)||![400,700].includes(r.weight)||!['normal','italic'].includes(r.style)||!/^sha256:[a-f0-9]{64}$/.test(r.font_digest))throw new TypeError('Invalid DOCX font substitution record')
  for(const s of [r.source_id,r.face_id])if(typeof s!=='string'||s.length===0||s.length>256||!/^[A-Za-z0-9._:/-]+$/.test(s))throw new TypeError('Invalid font substitution identity')
  for(const s of [r.source_family,r.selected_family])if(typeof s!=='string'||s.length===0||s.length>128||s.trim()!==s||/[\x00-\x1f\x7f]/.test(s))throw new TypeError('Invalid font substitution family')
  const key=JSON.stringify([r.source_id,r.source_role]);if(seen.has(key))throw new TypeError('Duplicate font substitution source');seen.add(key)
 }
 return records
}
/** Validate derived evidence against original source properties, actual selected
 * manifest faces and the explicit consumer policy; no source projection. */
export function qualifyNativeDocxFontSubstitutionsV1(shaped:NativeDocxShapedLinesV1,resolved:NativeDocxResolvedLayoutInputV1,manifest:NativeFontManifest,policyValue:unknown,document:NativeDocxDocumentV1):NativeDocxFontSubstitutionV1[]{
 if(shaped.document_id!==document.document_id||shaped.revision!==document.revision||resolved.document_id!==document.document_id||resolved.revision!==document.revision)throw new TypeError('Font evidence source revision mismatch')
 const policy=decodeExplicitFontPolicyV1(policyValue),records=decodeNativeDocxFontSubstitutionsV1(shaped.font_substitutions??[])
 if(records.length&&resolved.diagnostics.some(d=>d.code==='FONT_MATCHING_METADATA_PRESERVED'))throw new TypeError('Font matching metadata requires exact faces')
 const expectedDiagnostics=new Set<string>()
 const paragraphs=new Map(resolved.paragraphs.map(p=>[p.paragraph_id,p])),runs=new Map(resolved.runs.map(r=>[r.run_id,r]))
 const nativeParagraphs=new Map<string,NativeDocxParagraphV1>()
 for(const story of [document.body,...document.headers,...document.footers,...document.notes,...document.comment_stories])for(const block of story.blocks){if(block.paragraph)nativeParagraphs.set(block.paragraph.id,block.paragraph);for(const row of block.table?.rows??[])for(const cell of row.cells)for(const p of cell.paragraphs)nativeParagraphs.set(p.id,p)}
 const nativeRuns=new Map([...nativeParagraphs.values()].flatMap(p=>p.runs.map(r=>[r.id,{run:r,paragraph:p.id}] as const)))
 const fragments=new Map<string,NativeDocxShapedLinesV1['paragraphs'][number]['lines'][number]['fragments']>()
 for(const p of shaped.paragraphs)for(const line of p.lines)for(const f of line.fragments){const key=JSON.stringify([f.source_id,f.source_kind]);const list=fragments.get(key)??[];list.push(f);fragments.set(key,list)}
 const propertiesFor=(id:string,role:NativeDocxFontSubstitutionV1['source_role'])=>role==='run'?runs.get(id)?.properties:role==='paragraph-mark'?paragraphs.get(id)?.paragraph_mark_properties:paragraphs.get(id)?.numbering?.marker_properties
 const recordKeys=new Set(records.map(r=>JSON.stringify([r.source_id,r.source_role])))
 const eligibleSource=(id:string,role:NativeDocxFontSubstitutionV1['source_role'])=>{
  const props=propertiesFor(id,role),p=paragraphs.get(role==='run'?nativeRuns.get(id)?.paragraph??'':id)
  if(!props||!p||props.rtl||p.properties.bidi)throw new TypeError('Font preview requires source-qualified left-to-right properties')
  if(role==='run'){const source=nativeRuns.get(id)?.run;if(source?.kind!=='text'||typeof source.text!=='string'||!/^[\x20-\x7e]*$/.test(source.text))throw new TypeError('Font substitution currently requires original graphic ASCII Latin runs')}
  if(role==='list-marker'&&(!p.numbering||p.numbering.format==='bullet'||!/^[\x20-\x7e]+$/.test(p.numbering.resolved_text)))throw new TypeError('Only source-qualified Latin numbering markers may substitute')
 }
 for(const r of records){
  const paragraph=paragraphs.get(r.source_id)
  const properties=propertiesFor(r.source_id,r.source_role)
  eligibleSource(r.source_id,r.source_role)
  if(!properties||properties.font_family!==r.source_family||r.source_role==='list-marker'&&paragraph?.numbering?.format==='bullet')throw new TypeError('Substitution does not join an eligible source font')
  const run:TextRunInput={version:1,text:'A',fontSizeMilliPoints:1000,font:{families:[r.source_family],weight:properties.bold?700:400,style:properties.italic?'italic':'normal',stretch:100},script:'Latn',language:'und',direction:'ltr'}
  const selected=selectExplicitFontV1(manifest,run,policy)
  if(!selected||selected.face.resolution!=='substitute'||selected.face.family!==r.selected_family||selected.face.faceId!==r.face_id||selected.face.contentDigest!==r.font_digest||selected.face.weight!==r.weight||selected.face.style!==r.style)throw new TypeError('Substitution does not join operator mapping and selected bytes')
  const expected=nativeDocxFontSubstitutionDiagnosticV1(r)
  if(!shaped.diagnostics.some(d=>JSON.stringify(d)===JSON.stringify(expected)))throw new TypeError('Substitution lacks its exact derived diagnostic')
  expectedDiagnostics.add(JSON.stringify(expected))
  const sourceFragments=fragments.get(JSON.stringify([r.source_id,r.source_role]))??[]
  if(r.source_role!=='paragraph-mark'&&(!sourceFragments.length||sourceFragments.some(f=>f.face_id!==r.face_id)))throw new TypeError('Substitution has no matching shaped source fragments')
  if(r.source_role==='paragraph-mark'&&!shaped.paragraphs.some(p=>p.paragraph_id===r.source_id&&p.lines.some(l=>l.fragments.length===0||l.fragments.every(f=>f.source_kind==='tab'))))throw new TypeError('Paragraph-mark evidence has no blank-line metric consumer')
 }
 for(const [key,fs]of fragments){const f=fs[0]!;if(f.source_kind!=='run'&&f.source_kind!=='list-marker')continue;const props=propertiesFor(f.source_id,f.source_kind);if(!props?.font_family)throw new TypeError('Shaped font has no source family');const selected=selectExplicitFontV1(manifest,{version:1,text:'A',fontSizeMilliPoints:1000,font:{families:[props.font_family],weight:props.bold?700:400,style:props.italic?'italic':'normal',stretch:100},script:'Latn',language:'und',direction:'ltr'},policy);if(!selected||fs.some(f=>f.face_id!==selected.face.faceId)||selected.face.resolution==='substitute'&&!recordKeys.has(key))throw new TypeError('Shaped source font selection lacks complete substitution evidence');if(selected.face.resolution==='substitute'){eligibleSource(f.source_id,f.source_kind);if(fs.some(f=>f.direction!=='ltr'||!['Latn','Zyyy'].includes(f.script)))throw new TypeError('Substituted fragments lack source-qualified Latin direction')}}
 for(const p of shaped.paragraphs)if(p.lines.some(l=>l.fragments.length===0||l.fragments.every(f=>f.source_kind==='tab'))){const props=propertiesFor(p.paragraph_id,'paragraph-mark');if(!props?.font_family)throw new TypeError('Blank-line font source is missing');const selected=selectExplicitFontV1(manifest,{version:1,text:'',fontSizeMilliPoints:1000,font:{families:[props.font_family],weight:props.bold?700:400,style:props.italic?'italic':'normal',stretch:100},script:'Zyyy',language:'und',direction:'ltr'},policy);if(!selected||selected.face.resolution==='substitute'&&!recordKeys.has(JSON.stringify([p.paragraph_id,'paragraph-mark'])))throw new TypeError('Blank-line font substitution evidence is incomplete')}
 for(const d of shaped.diagnostics)if(d.code===DOCX_FONT_SUBSTITUTION_CODE&&!expectedDiagnostics.has(JSON.stringify(d)))throw new TypeError('Unbound substitution diagnostic')
 return records
}
export function isQualifiedNativeDocxFontDiagnosticV1(d:NativeDocxShapingDiagnosticV1,records:readonly NativeDocxFontSubstitutionV1[]):boolean{
 return records.some(r=>JSON.stringify(nativeDocxFontSubstitutionDiagnosticV1(r))===JSON.stringify(d))
}

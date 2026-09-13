import {decodeNativeDocxDocument,type NativeDocxDocumentV1,type NativeDocxSourceAnchorV1} from './nativeContract.js'
import {createNativeDocxPartialContentPreviewV1,type NativeDocxPartialParagraphV1,type NativeDocxPartialSegmentV1} from './nativePartialContentV1.js'
export interface NativeDocxReviewChangeV1 {package_sha256:string;paragraph_id:string;diagnostic_id:string;anchor:NativeDocxSourceAnchorV1;kind:'insertion'|'deletion'|'move-to'|'move-from';revision_id:string;author:string;created_at:string;run_ids:string[]}
export interface NativeDocxReviewEvidenceV1 {items:NativeDocxReviewChangeV1[];omitted_count:number}
function record(v:unknown,keys:string[]):Record<string,unknown>{
 if(!v||typeof v!=='object'||Object.getPrototypeOf(v)!==Object.prototype||Reflect.ownKeys(v).some(k=>typeof k!=='string')||Object.keys(v).sort().join(',')!==[...keys].sort().join(','))throw new TypeError('Invalid review evidence record')
 const out:Record<string,unknown>={};for(const key of keys){const d=Object.getOwnPropertyDescriptor(v,key);if(!d||!('value'in d))throw new TypeError('Review accessors are forbidden');out[key]=d.value}return out
}
function array(v:unknown,max:number):unknown[]{
 if(!Array.isArray(v)||Object.getPrototypeOf(v)!==Array.prototype||v.length>max||Reflect.ownKeys(v).length!==v.length+1)throw new TypeError('Invalid review array')
 return Array.from({length:v.length},(_,i)=>{const d=Object.getOwnPropertyDescriptor(v,String(i));if(!d||!('value'in d))throw new TypeError('Review arrays must be dense plain data');return d.value})
}
const paragraphs=(d:NativeDocxDocumentV1)=>d.body.blocks.flatMap(b=>b.paragraph?[b.paragraph]:b.table!.rows.flatMap(r=>r.cells.flatMap(c=>c.paragraphs)))
/** Same-byte native producer evidence. This joins original diagnostic anchors;
 * callers must not treat arbitrary hand-authored models as XML proof. */
export function decodeNativeDocxReviewEvidenceV1(source:unknown,input:unknown):NativeDocxReviewEvidenceV1 {
 const decoded=decodeNativeDocxDocument(source);if(!decoded.ok)throw new TypeError('Invalid review source')
 if(input===undefined)return {items:[],omitted_count:0}
 const d=decoded.value,outer=record(input,['items','omitted_count']),ps=new Map(paragraphs(d).map(p=>[p.id,p]))
 if(!Number.isSafeInteger(outer.omitted_count)||Number(outer.omitted_count)<0||Number(outer.omitted_count)>100000)throw new TypeError('Review omission budget exceeded')
 const seen=new Set<string>(),items:NativeDocxReviewChangeV1[]=[]
 for(const raw of array(outer.items,128)){
  const v=record(raw,['package_sha256','paragraph_id','diagnostic_id','anchor','kind','revision_id','author','created_at','run_ids']),a=record(v.anchor,['part_name','path','start_byte','end_byte','xml_sha256'])
  if(v.package_sha256!==d.source.package_sha256||typeof v.paragraph_id!=='string'||typeof v.diagnostic_id!=='string'||seen.has(v.diagnostic_id)||typeof v.kind!=='string'||!['insertion','deletion','move-to','move-from'].includes(v.kind))throw new TypeError('Invalid review identity')
  for(const [key,max]of [['revision_id',128],['author',1024],['created_at',128]] as const)if(typeof v[key]!=='string'||v[key].length>max||key==='revision_id'&&!v[key].length)throw new TypeError('Invalid review metadata')
  const p=ps.get(v.paragraph_id),diag=d.unsupported.find(x=>x.id===v.diagnostic_id),tag={insertion:'ins',deletion:'del','move-to':'moveTo','move-from':'moveFrom'}[v.kind as NativeDocxReviewChangeV1['kind']]
  if(!p||!diag?.anchor||diag.scope_id!==p.id||diag.code!==(v.kind==='insertion'||v.kind==='move-to'?'WRAPPED_RUN_MARKUP':'UNMODELED_PARAGRAPH_CONTENT')||diag.capability!=='run-structure'||diag.preservation!=='refuse-mutation'||Object.entries(diag.anchor).some(([k,x])=>a[k]!==x)||a.part_name!==d.source.main_part||a.part_name!==p.anchor.part_name||typeof a.path!=='string'||!new RegExp('^/w:'+tag+'\\[[1-9][0-9]*\\]$').test(a.path.slice(p.anchor.path.length))||!a.path.startsWith(p.anchor.path)||Number(a.start_byte)<p.anchor.start_byte||Number(a.end_byte)>p.anchor.end_byte)throw new TypeError('Review evidence does not join a direct original wrapper')
  const runIDs=array(v.run_ids,128),seenRuns=new Set<string>()
  if(v.kind!=='insertion'&&runIDs.length)throw new TypeError('Only insertion text can be qualified')
  for(const id of runIDs){
   const r=p.runs.find(r=>r.id===id)
   if(typeof id!=='string'||seenRuns.has(id)||!r||r.kind!=='text'||r.page_field||r.anchor.part_name!==a.part_name||r.anchor.start_byte<Number(a.start_byte)||r.anchor.end_byte>Number(a.end_byte)||!r.anchor.path.startsWith(a.path+'/')||!/^\/w:r\[[1-9][0-9]*\]\/w:t\[[1-9][0-9]*\]$/.test(r.anchor.path.slice(a.path.length)))throw new TypeError('Invalid direct insertion run join')
   seenRuns.add(id)
  }
  const contained=p.runs.filter(r=>r.anchor.path.startsWith(a.path+'/')).map(r=>r.id)
  if(runIDs.length&&JSON.stringify(contained)!==JSON.stringify(runIDs))throw new TypeError('Insertion run inventory is incomplete or reordered')
  seen.add(v.diagnostic_id);items.push({...v,anchor:a,run_ids:runIDs} as unknown as NativeDocxReviewChangeV1)
 }
 return {items,omitted_count:Number(outer.omitted_count)}
}
export interface NativeDocxReviewInventoryV1 {
 policy:'source-review-inventory-v1';read_only:true;source:{document_id:string;revision:string;package_sha256:string}
 items:Array<NativeDocxReviewChangeV1&{segments:NativeDocxPartialSegmentV1[];text_status:'qualified-insertion'|'omitted';retained_diagnostic_ids:string[]}>
 omitted_count:number;source_diagnostics:ReturnType<typeof createNativeDocxPartialContentPreviewV1>['source_diagnostics']
}
export function createNativeDocxReviewInventoryV1(source:unknown,options:{policy:'source-review-inventory-v1';read_only:true},resolved:unknown,evidence:unknown):NativeDocxReviewInventoryV1 {
 const opts=record(options,['policy','read_only']);if(opts.policy!=='source-review-inventory-v1'||opts.read_only!==true)throw new TypeError('Explicit read-only review policy required')
 const decoded=decodeNativeDocxDocument(source);if(!decoded.ok)throw new TypeError('Invalid review source')
 const d=structuredClone(decoded.value),review=decodeNativeDocxReviewEvidenceV1(d,evidence),qualified=new Set(review.items.filter(i=>i.kind==='insertion'&&i.run_ids.length).map(i=>i.diagnostic_id))
 // The exact producer-qualified insertion boundary is the only diagnostic
 // exception. Every other source, ancestor and resolved visibility gate runs.
 // The original source/diagnostics are retained below; this is never native paint.
 const projected=createNativeDocxPartialContentPreviewV1({...d,unsupported:d.unsupported.filter(x=>!qualified.has(x.id))},{policy:'source-text-with-omissions-v1',read_only:true},resolved)
 const ps=new Map<string,NativeDocxPartialParagraphV1>()
 for(const b of projected.blocks){if(b.kind==='paragraph')ps.set(b.source.scope_id,b);if(b.kind==='table-source')for(const c of b.cells)for(const p of c.paragraphs)if(p.kind==='paragraph')ps.set(p.source.scope_id,p)}
 let units=0,omitted=review.omitted_count
 const items:NativeDocxReviewInventoryV1['items']=[]
 for(const item of review.items){
  const metadata=item.revision_id.length+item.author.length+item.created_at.length
  if(units+metadata>32768){omitted++;continue}units+=metadata
  const ids=new Set(item.run_ids),segments=(ps.get(item.paragraph_id)?.segments??[]).filter(s=>ids.has(s.source.scope_id))
  const allText=item.kind==='insertion'&&ids.size>0&&segments.length===ids.size&&segments.every(s=>s.kind==='text')
  const textUnits=allText?segments.reduce((n,s)=>n+(s.kind==='text'?s.text.length:0),0):0
  const show=allText&&units+textUnits<=32768;if(show)units+=textUnits
  items.push({...item,segments:show?segments:[],text_status:show?'qualified-insertion':'omitted',retained_diagnostic_ids:[item.diagnostic_id]})
 }
 return {policy:'source-review-inventory-v1',read_only:true,source:{document_id:d.document_id,revision:d.revision,package_sha256:d.source.package_sha256},items,omitted_count:omitted,source_diagnostics:{document:d.unsupported,resolved:projected.source_diagnostics.resolved}}
}

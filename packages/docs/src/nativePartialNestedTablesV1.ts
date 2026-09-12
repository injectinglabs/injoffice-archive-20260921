import {decodeNativeDocxDocument,type NativeDocxSourceAnchorV1} from './nativeContract.js'

export interface NativeDocxNestedTableOmissionV1 {package_sha256:string;part_sha256:string;table_id:string;cell_id:string;diagnostic_id:string;anchor:NativeDocxSourceAnchorV1}
export interface NativeDocxNestedTableOmissionsV1 {items:NativeDocxNestedTableOmissionV1[];omitted_count:number}
function record(v:unknown,keys:string[]):Record<string,unknown>{
 if(!v||typeof v!=='object'||Object.getPrototypeOf(v)!==Object.prototype||Reflect.ownKeys(v).some(k=>typeof k!=='string')||Object.keys(v).sort().join(',')!==[...keys].sort().join(','))throw new TypeError('Invalid nested table evidence record')
 const out:Record<string,unknown>={}
 for(const key of keys){const d=Object.getOwnPropertyDescriptor(v,key);if(!d||!('value'in d))throw new TypeError('Nested table evidence accessors are forbidden');out[key]=d.value}
 return out
}
/** Trusted same-byte inspector evidence, not independent XML parsing. Main-part
 * SHA is producer evidence; the package and exact diagnostic slice join remain
 * mandatory. This authorizes an omission boundary only, never nested content. */
export function decodeNativeDocxNestedTableOmissionsV1(source:unknown,input:unknown):NativeDocxNestedTableOmissionsV1 {
 const d=decodeNativeDocxDocument(source);if(!d.ok)throw new TypeError('Invalid source for nested omissions')
 if(input===undefined)return {items:[],omitted_count:0}
 const outer=record(input,['items','omitted_count']),raw=outer.items
 if(!Number.isSafeInteger(outer.omitted_count)||Number(outer.omitted_count)<0||Number(outer.omitted_count)>100000||!Array.isArray(raw)||Object.getPrototypeOf(raw)!==Array.prototype||raw.length>64||Reflect.ownKeys(raw).length!==raw.length+1)throw new TypeError('Nested omission budget exceeded')
 const items:NativeDocxNestedTableOmissionV1[]=[],seen=new Set<string>(),hash=/^sha256:[0-9a-f]{64}$/
 for(let i=0;i<raw.length;i++){
  const descriptor=Object.getOwnPropertyDescriptor(raw,String(i));if(!descriptor||!('value'in descriptor))throw new TypeError('Nested omission arrays must be dense plain data')
  const v=record(descriptor.value,['package_sha256','part_sha256','table_id','cell_id','diagnostic_id','anchor']),a=record(v.anchor,['part_name','path','start_byte','end_byte','xml_sha256'])
  if(v.package_sha256!==d.value.source.package_sha256||typeof v.part_sha256!=='string'||!hash.test(v.part_sha256)||items.some(x=>x.part_sha256!==v.part_sha256)||typeof v.table_id!=='string'||typeof v.cell_id!=='string'||typeof v.diagnostic_id!=='string'||seen.has(v.diagnostic_id))throw new TypeError('Invalid nested source identity')
  const table=d.value.body.blocks.find(b=>b.table?.id===v.table_id)?.table,cell=table?.rows.flatMap(r=>r.cells).find(c=>c.id===v.cell_id)
  const diagnostics=d.value.unsupported.filter(x=>x.id===v.diagnostic_id),diag=diagnostics[0]
  if(!table||!cell||diagnostics.length!==1||diag?.code!=='NESTED_TABLE_OR_CELL_MARKUP'||diag.scope_id!==table.id||diag.capability!=='table-structure'||diag.preservation!=='refuse-mutation'||!diag.anchor||Object.entries(diag.anchor).some(([k,x])=>a[k]!==x))throw new TypeError('Nested omission does not join original diagnostic')
  if(a.part_name!==d.value.source.main_part||a.part_name!==cell.anchor.part_name||typeof a.path!=='string'||!a.path.startsWith(cell.anchor.path+'/w:tbl[')||!/^\/w:tbl\[[1-9][0-9]*\]$/.test(a.path.slice(cell.anchor.path.length))||!Number.isSafeInteger(a.start_byte)||!Number.isSafeInteger(a.end_byte)||Number(a.start_byte)<cell.anchor.start_byte||Number(a.end_byte)>cell.anchor.end_byte||Number(a.end_byte)<=Number(a.start_byte)||typeof a.xml_sha256!=='string'||!hash.test(a.xml_sha256))throw new TypeError('Invalid direct nested table containment')
  if(cell.paragraphs.some(p=>p.anchor.part_name!==a.part_name||p.anchor.start_byte<Number(a.end_byte)&&p.anchor.end_byte>Number(a.start_byte))||items.some(x=>x.cell_id===cell.id&&x.anchor.start_byte<Number(a.end_byte)&&x.anchor.end_byte>Number(a.start_byte)))throw new TypeError('Ambiguous nested table boundaries')
  seen.add(v.diagnostic_id);items.push({...v,anchor:a} as unknown as NativeDocxNestedTableOmissionV1)
 }
 return {items,omitted_count:Number(outer.omitted_count)}
}

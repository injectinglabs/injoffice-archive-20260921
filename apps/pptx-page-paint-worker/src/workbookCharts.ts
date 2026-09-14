import {createHash} from 'node:crypto'
import {decodeNativePptxChartWorkbookInspection,resolveNativePptxWorkbookCharts,type NativePptxDeck,type NativeWorkbookChartResolution} from '@injoffice/pptx-native'

const MAX_PAYLOAD=8*1024*1024
function object(value:unknown,keys:readonly string[]):Record<string,unknown>{
 if(!value||typeof value!=='object'||Array.isArray(value)||Object.keys(value).some(k=>!keys.includes(k)))throw new TypeError('Invalid workbook chart engine payload')
 return value as Record<string,unknown>
}
/** Operator-local engine payload, never resolved geometry. Both engine outputs
 * are decoded and source-joined again before admission to the renderer. */
export async function workbookChartsForPreview(raw:unknown,deck:NativePptxDeck,packageSHA:string):Promise<NativeWorkbookChartResolution>{
 const payload=object(raw,['inspection_json','workbooks'])
 if(Buffer.byteLength(JSON.stringify(payload),'utf8')>MAX_PAYLOAD||typeof payload.inspection_json!=='string'||!Array.isArray(payload.workbooks)||payload.workbooks.length>8)throw new RangeError('Workbook chart payload exceeds aggregate UTF-8 budget')
 const inspection=await decodeNativePptxChartWorkbookInspection(payload.inspection_json,deck,packageSHA)
 if(payload.workbooks.length!==inspection.workbooks.length)throw new TypeError('Workbook engine payload resource closure mismatch')
 const records=new Map<string,{sha:string;contract?:string;refusal?:string}>()
 for(const item of payload.workbooks){
  const value=object(item,['part','sha256','contract_json','refusal'])
  const source=inspection.workbooks.find(r=>r.part===value.part)
  if(!source||source.sha256!==value.sha256||records.has(source.part))throw new TypeError('Workbook engine payload source identity mismatch')
  const contract=value.contract_json,refusal=value.refusal
  if((typeof contract==='string')===(typeof refusal==='string')||contract!==undefined&&(typeof contract!=='string'||contract.length===0)||refusal!==undefined&&(typeof refusal!=='string'||!refusal||refusal.length>2048))throw new TypeError('Workbook engine payload requires exactly one contract or refusal')
  records.set(source.part,{sha:source.sha256,...(typeof contract==='string'?{contract}:{}),...(typeof refusal==='string'?{refusal}:{})})
 }
 return resolveNativePptxWorkbookCharts(inspection,async bytes=>{
  const sha=createHash('sha256').update(bytes).digest('hex')
  const matching=[...records.values()].filter(r=>r.sha===sha)
  if(!matching.length||matching.some(r=>r.contract!==matching[0]!.contract||r.refusal!==matching[0]!.refusal))throw new TypeError('Workbook engine payload digest is absent or ambiguous')
  const record=matching[0]!
  if(record.refusal)throw new RangeError(record.refusal)
  return record.contract!
 })
}

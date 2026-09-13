import {decodeNativeWorkbookV2} from '@injoffice/sheets/browser'
import {resolveChartWorkbookReferences} from './chartWorkbookResolve.js'
import type {ChartWorkbookBinding,ChartWorkbookReference,ChartWorkbookResolvedValues} from './chartWorkbookTypes.js'

/** The caller explicitly supplies its trusted XLSX engine. Validation checks the
 * returned contract and identity; it cannot authenticate an arbitrary callback's
 * computation. No chart cache or formula cache is an extraction input. */
export type ChartWorkbookExtractor=(bytes:Uint8Array)=>Promise<string>

export async function extractChartWorkbookReferences(
 binding:ChartWorkbookBinding,references:readonly ChartWorkbookReference[],
 embeddedBytes:Uint8Array,extract:ChartWorkbookExtractor,
):Promise<readonly ChartWorkbookResolvedValues[]> {
 if(!(embeddedBytes instanceof Uint8Array)||embeddedBytes.byteLength<1||embeddedBytes.byteLength>8*1024*1024||embeddedBytes.byteLength!==binding.byteLength||references.length<1||references.length>64)throw new RangeError('embedded workbook byte budget or length mismatch')
 // Freeze metadata and snapshot bytes before the first asynchronous boundary.
 const bound=Object.freeze({...binding})
 const refs=references.map(reference=>Object.freeze({...reference,range:Object.freeze({...reference.range})}))
 const snapshot=Uint8Array.from(embeddedBytes)
 const digest=Array.from(new Uint8Array(await globalThis.crypto.subtle.digest('SHA-256',snapshot)),b=>b.toString(16).padStart(2,'0')).join('')
 if(bound.sha256!==digest)throw new RangeError('embedded workbook bytes do not match source binding')
 // The callback may transfer or mutate its input. It never owns our snapshot.
 const serialized=await extract(Uint8Array.from(snapshot))
 if(typeof serialized!=='string'||serialized.length>16*1024*1024||new TextEncoder().encode(serialized).byteLength>16*1024*1024)throw new RangeError('extracted workbook JSON budget exceeded')
 const decoded=decodeNativeWorkbookV2(serialized)
 if(!decoded.ok)throw new RangeError('injected XLSX extractor returned an invalid native V2 contract')
 return resolveChartWorkbookReferences(bound,refs,decoded.value)
}

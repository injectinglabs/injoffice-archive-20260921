import type {NativeElement,NativeLiteralBarAxis,NativePptxDeck} from './types.js'
import {assertNativePptx} from './validate.js'
import {nativeChartDecimal} from './chartDecimalValidation.js'
import {validNativeChartAxisLabels} from './chartAxisLabelsValidation.js'
import {parseChartWorkbookRange} from './chartWorkbookRange.js'
import {parseChartWorkbookJson} from './chartWorkbookJson.js'
import type {ChartWorkbookReference} from './chartWorkbookTypes.js'
import type {NativePptxChartWorkbookInspection,NativePptxWorkbookChartSource} from './chartWorkbookInspectionTypes.js'

const fail=():never=>{throw new TypeError('invalid or source-mismatched chart workbook inspection')}
function object(value:unknown,required:string,optional=''):Record<string,unknown>{
 if(!value||typeof value!=='object'||Array.isArray(value))return fail()
 const result=value as Record<string,unknown>,allowed=new Set([...required.split(' '),...optional.split(' ')])
 if(Object.keys(result).some(key=>!allowed.has(key))||required.split(' ').some(key=>!Object.hasOwn(result,key)))return fail()
 return result
}
function text(value:unknown,max=1024):string{if(typeof value!=='string'||value.length<1||value.length>max||/[\p{Cc}]/u.test(value))return fail();return value}
function hash(value:unknown):string{const s=text(value,64);if(/^[a-f0-9]{64}$/.exec(s)?.[0]!==s)return fail();return s}
function part(value:unknown):string{const s=text(value);if(s.startsWith('/')||/[\\?#]/.test(s)||s.split('/').some(x=>x===''||x==='.'||x==='..'))return fail();return s}
function integer(value:unknown,min:number,max:number):number{if(!Number.isSafeInteger(value)||Object.is(value,-0)||(value as number)<min||(value as number)>max)return fail();return value as number}
function array(value:unknown,max:number):unknown[]{if(!Array.isArray(value)||value.length>max)return fail();return value}
function bool(value:unknown):boolean{if(typeof value!=='boolean')return fail();return value}
function rgb(value:unknown):string{const s=text(value,7);if(/^#[A-F0-9]{6}$/.exec(s)?.[0]!==s)return fail();return s}
function decimal(value:unknown):string{const s=text(value,128);if(!nativeChartDecimal(s))return fail();return s}
function compare(a:string,b:string):number{const x=nativeChartDecimal(a)!,y=nativeChartDecimal(b)!,e=Math.min(x.exponent,y.exponent),d=x.coefficient*10n**BigInt(x.exponent-e)-y.coefficient*10n**BigInt(y.exponent-e);return d<0n?-1:d>0n?1:0}
function reference(value:unknown,kind:'strRef'|'numRef'):ChartWorkbookReference{
 const r=object(value,'kind formula range cachePresent');if(r.kind!==kind)return fail()
 const formula=text(r.formula),expected=parseChartWorkbookRange(formula),range=object(r.range,'sheet startRow startColumn endRow endColumn count')
 for(const key of ['sheet','startRow','startColumn','endRow','endColumn','count'] as const)if(range[key]!==expected[key])return fail()
 bool(r.cachePresent);return r as unknown as ChartWorkbookReference
}
function axis(value:unknown,numeric:boolean):NativeLiteralBarAxis{
 const a=object(value,'id crossAxisId orientation position deleted','color widthEmu min max crossesAt labels')
 integer(a.id,0,4294967295);integer(a.crossAxisId,0,4294967295)
 if(!['minMax','maxMin'].includes(a.orientation as string)||!['b','l'].includes(a.position as string))return fail()
 if(bool(a.deleted)){if(a.color!==undefined||a.widthEmu!==undefined||a.labels!==undefined)return fail()}
 else{rgb(a.color);integer(a.widthEmu,1,20116800)}
 if(numeric){decimal(a.min);decimal(a.max);decimal(a.crossesAt);if(compare(a.min as string,a.max as string)>=0||compare(a.crossesAt as string,'0')!==0||compare(a.min as string,'0')>0||compare(a.max as string,'0')<0)return fail()}
 else if(a.min!==undefined||a.max!==undefined||a.crossesAt!==undefined)return fail()
 if(a.labels!==undefined){
  const l=object(a.labels,'profile position majorTickMark style','majorUnit numberFormat'),s=object(l.style,'fontFamily fontSize color bold italic language')
  if(l.profile!=='explicit-axis-labels-v1'||!['low','high'].includes(l.position as string)||!['none','out'].includes(l.majorTickMark as string))return fail()
  text(s.fontFamily,128);integer(s.fontSize,1,400000);rgb(s.color);bool(s.bold);bool(s.italic);text(s.language,128)
  if(numeric){decimal(l.majorUnit);if(typeof l.numberFormat!=='string'||/^0(?:\.0{1,6})?$/.exec(l.numberFormat)?.[0]!==l.numberFormat)return fail()}
  else if(l.majorUnit!==undefined||l.numberFormat!==undefined)return fail()
 }
 return a as unknown as NativeLiteralBarAxis
}
function source(value:unknown):NativePptxWorkbookChartSource{
 const s=object(value,'family xAxis yAxis series plotVisibleOnly','barDirection gapWidth dispBlanksAs')
 if(!['bar','line','scatter'].includes(s.family as string)||s.plotVisibleOnly!==false||s.dispBlanksAs!==undefined&&!['gap','zero','span'].includes(s.dispBlanksAs as string))return fail()
 const bar=s.family==='bar',scatter=s.family==='scatter',horizontal=bar&&s.barDirection==='bar'
 if(bar){if(!['col','bar'].includes(s.barDirection as string))return fail();integer(s.gapWidth,0,500)}else if(s.barDirection!==undefined||s.gapWidth!==undefined)return fail()
 const x=axis(s.xAxis,scatter||horizontal),y=axis(s.yAxis,!horizontal)
 if(x.position!=='b'||y.position!=='l'||x.id===y.id||x.crossAxisId!==y.id||y.crossAxisId!==x.id||!validNativeChartAxisLabels(x,y,scatter||horizontal)||!validNativeChartAxisLabels(y,x,!horizontal))return fail()
 const series=array(s.series,16),ids=new Set<number>();let count:number|undefined,totalText=0,refs=0
 if(series.length<1)return fail()
 series.forEach((value,i)=>{
  const item=object(value,'index order valueReference','title titleReference categoryReference xReference colors color widthEmu')
  const index=integer(item.index,0,4294967295);if(ids.has(index)||item.order!==i)return fail();ids.add(index)
  if(item.title!==undefined){if(typeof item.title!=='string'||item.title.length>1024||item.titleReference!==undefined)return fail();totalText+=item.title.length}
  if(item.titleReference!==undefined){if(reference(item.titleReference,'strRef').range.count!==1)return fail();refs++}
  const values=reference(item.valueReference,'numRef');refs++
  if(scatter){if(item.categoryReference!==undefined||reference(item.xReference,'numRef').range.count!==values.range.count)return fail()}
  else{if(item.xReference!==undefined||reference(item.categoryReference,'strRef').range.count!==values.range.count||count!==undefined&&count!==values.range.count)return fail();count=values.range.count}
  refs++
  if(bar){const colors=array(item.colors,256);if(colors.length!==values.range.count||item.color!==undefined||item.widthEmu!==undefined)return fail();colors.forEach(rgb)}
  else{if(item.colors!==undefined)return fail();rgb(item.color);integer(item.widthEmu,1,20116800)}
 })
 if(refs>64||totalText>32768)return fail()
 return s as unknown as NativePptxWorkbookChartSource
}
function freeze<T>(value:T):T{if(value&&typeof value==='object'){for(const child of Object.values(value))freeze(child);Object.freeze(value)}return value}
const admitted=new WeakSet<NativePptxChartWorkbookInspection>()
export function assertAdmittedChartWorkbookInspection(value:NativePptxChartWorkbookInspection):void{if(!admitted.has(value))throw new TypeError('chart workbook inspection must pass the source-bound decoder')}

/** Trust the explicitly selected PPTX engine, validate its complete transport,
 * and bind it to the caller's private source snapshot and native deck. */
export async function decodeNativePptxChartWorkbookInspection(json:string,deck:NativePptxDeck,packageSHA256:string):Promise<NativePptxChartWorkbookInspection>{
 assertNativePptx(deck);hash(packageSHA256)
 const input=object(parseChartWorkbookJson(json),'protocol package_sha256 source_revision charts workbooks omissions')
 if(input.protocol!=='pptx-chart-workbook-inspection-v1'||input.package_sha256!==packageSHA256||input.source_revision!==`rev-${packageSHA256}`||deck.sourceRevision!==input.source_revision)return fail()
 const charts=array(input.charts,64),omissions=array(input.omissions,64),resources=array(input.workbooks,8)
 if(charts.length+omissions.length>64)return fail()
 const expected=new Map<string,{element:Extract<NativeElement,{kind:'chart'}>;slideIndex:number}>()
 deck.slides.forEach((slide,slideIndex)=>{const walk=(elements:readonly NativeElement[])=>{for(const element of elements){if(element.kind==='group')walk(element.children);else if(element.kind==='chart'){if(!element.source)return fail();expected.set(`${slide.id}\0${element.source.objectId}`,{element,slideIndex})}}};walk(slide.elements)})
 const resourceMap=new Map<string,Record<string,unknown>>();let totalBytes=0
 for(const raw of resources){
  const r=object(raw,'part sha256 byteLength bytesBase64'),name=part(r.part),sha=hash(r.sha256),length=integer(r.byteLength,1,8388608)
  totalBytes+=length;if(totalBytes>16777216||resourceMap.has(name)||typeof r.bytesBase64!=='string'||r.bytesBase64.length!==4*Math.ceil(length/3))return fail()
  const encoded=r.bytesBase64,padding=(3-length%3)%3
  for(let i=0;i<encoded.length;i++){const c=encoded.charCodeAt(i);if(i>=encoded.length-padding?c!==61:!(c>=65&&c<=90||c>=97&&c<=122||c>=48&&c<=57||c===43||c===47))return fail()}
  const decoded=Uint8Array.from(atob(encoded),c=>c.charCodeAt(0));if(decoded.length!==length||btoa(String.fromCharCode(...decoded.subarray(length-(length%3||3)))).slice(-4)!==encoded.slice(-4))return fail()
  const digest=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',decoded)),b=>b.toString(16).padStart(2,'0')).join('');if(digest!==sha)return fail()
  resourceMap.set(name,r)
 }
 const seen=new Set<string>(),used=new Set<string>()
 for(const raw of charts){
  const c=object(raw,'slide_id slide_index slide_part slide_sha256 element_id object_id frame_sha256 chart_relationship_id chart_part chart_sha256 source workbook')
  const key=`${text(c.slide_id,256)}\0${text(c.object_id,64)}`,entry=expected.get(key);if(!entry||seen.has(key))return fail();seen.add(key)
  const slide=deck.slides[entry.slideIndex]!,element=entry.element
  if(c.slide_index!==entry.slideIndex||c.slide_part!==slide.source?.partName||c.slide_sha256!==slide.source?.fingerprintSha256||c.element_id!==element.id||c.frame_sha256!==element.source?.fingerprintSha256||c.chart_relationship_id!==element.chart.relationshipId||c.chart_part!==element.chart.chartPart||c.chart_sha256!==element.chart.opaqueRef.fingerprintSha256)return fail()
  source(c.source)
  const binding=object(c.workbook,'relationshipId part sha256 byteLength','autoUpdate');text(binding.relationshipId);if(binding.autoUpdate!==undefined)bool(binding.autoUpdate)
  const name=part(binding.part),r=resourceMap.get(name);if(!r||binding.sha256!==r.sha256||binding.byteLength!==r.byteLength)return fail();used.add(name)
 }
 for(const raw of omissions){const o=object(raw,'slide_id object_id reason'),key=`${text(o.slide_id,256)}\0${text(o.object_id,64)}`;text(o.reason,4096);if(!expected.has(key)||seen.has(key))return fail();seen.add(key)}
 if(seen.size!==expected.size||used.size!==resourceMap.size)return fail()
 const result=freeze(input as unknown as NativePptxChartWorkbookInspection);admitted.add(result);return result
}

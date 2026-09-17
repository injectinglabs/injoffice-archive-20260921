import {isNativePreviewPartPathV1} from './nativePreviewPartPathV1.js'
import {snapshotNativePlainData} from './nativePlainData.js'

export interface NativeSheetPageConfigV1 {
 paper:'Letter'|'A4'; orientation:'portrait'|'landscape'; scale:number;
 left_inches:number;right_inches:number;top_inches:number;bottom_inches:number;
 /**
  * Authored header and footer margins, measured from the paper edge like the
  * other four (ECMA-376 §18.3.1.62), not added to them. Excel prints the body
  * between max(top, header) and max(bottom, footer). Omission reserves no
  * header or footer band, which is what an explicit host choice that states
  * neither means; a source-projected setting always states both.
  */
 header_inches?:number;footer_inches?:number;
 /** Omission retains down-then-over ordering. */
 page_order?:'downThenOver'|'overThenDown';
 /** Read-only shrink-to-fit target; zero leaves that dimension unconstrained. */
 fit_to_page?:{width:number;height:number};
}
/** Authored margins for a worksheet that declares no pageSetup at all. */
export interface NativeSheetPageMarginsV1 {
 left_inches:number;right_inches:number;top_inches:number;bottom_inches:number;
 /** Edge-measured header and footer margins; see NativeSheetPageConfigV1. */
 header_inches?:number;footer_inches?:number;
}
/** ECMA-376 CT_PageSetup attributes whose schema default a source may omit. */
export type NativeSheetPageDefaultedFactV1='paper'|'orientation'|'scale'
export interface NativeSheetPageSettingsV1 {
 sheet_id:string;sheet_part:string;status:'available'|'margins-only'|'unavailable';
 settings?:NativeSheetPageConfigV1;
 /** Present only for 'margins-only': paper, orientation and scale are unauthored. */
 margins?:NativeSheetPageMarginsV1;
 /**
  * Facts in `settings` that came from the ECMA-376 §18.3.1.63 attribute default
  * because the worksheet omitted the attribute, never from a printer. Present
  * only for 'available', and only for the attributes actually omitted, so a
  * consumer can label the geometry a host default rather than an authored one.
  */
 defaulted?:NativeSheetPageDefaultedFactV1[];
 warnings:string[];
}
/** A standalone bounded copy; package identity is joined by its enclosing object response. */
export function decodeNativeSheetPageSettingsV1(input:unknown):NativeSheetPageSettingsV1[]{
 const value=snapshotNativePlainData(input,{maxDepth:6,maxNodes:4096})
 const fail=():never=>{throw new TypeError('Invalid native worksheet page settings')}
 const exact=(v:unknown,keys:string[])=>{
  if(!v||typeof v!=='object'||Array.isArray(v))return fail()
  const o=v as Record<string,unknown>
  if(Object.keys(o).length!==keys.length||keys.some(k=>!Object.hasOwn(o,k)))return fail()
  return o
 }
 if(!Array.isArray(value)||value.length>64)return fail()
 const ids=new Set<string>(),parts=new Set<string>()
 return value.map(v=>{
  const status=(v as Record<string,unknown>)?.status
  const hasDefaulted=!!v&&typeof v==='object'&&Object.hasOwn(v,'defaulted')
  const o=exact(v,['sheet_id','sheet_part','status','warnings',...(status==='available'?['settings']:[]),...(status==='margins-only'?['margins']:[]),...(hasDefaulted?['defaulted']:[])])
  if(typeof o.sheet_id!=='string'||!/^[1-9][0-9]{0,9}$/.test(o.sheet_id)||Number(o.sheet_id)>0xffffffff||ids.has(o.sheet_id)||!isNativePreviewPartPathV1(o.sheet_part)||parts.has(o.sheet_part)||(status!=='available'&&status!=='margins-only'&&status!=='unavailable'))return fail()
  ids.add(o.sheet_id);parts.add(o.sheet_part)
  if(!Array.isArray(o.warnings)||o.warnings.length<1||o.warnings.length>8||o.warnings.some(w=>typeof w!=='string'||w.length>4096))return fail()
  let settings:NativeSheetPageConfigV1|undefined
  if(status==='available'){
   const has=(key:string)=>!!o.settings&&typeof o.settings==='object'&&Object.hasOwn(o.settings,key)
   const hasOrder=has('page_order'),hasFit=has('fit_to_page'),hasBands=['header_inches','footer_inches'].filter(has)
   const s=exact(o.settings,['paper','orientation','scale','left_inches','right_inches','top_inches','bottom_inches',...hasBands,...(hasOrder?['page_order']:[]),...(hasFit?['fit_to_page']:[])])
   if(hasFit){
    const fit=exact(s.fit_to_page,['width','height'])
    if(['width','height'].some(k=>!Number.isInteger(fit[k])||Object.is(fit[k],-0)||Number(fit[k])<0||Number(fit[k])>100)||!(Number(fit.width)>0||Number(fit.height)>0))return fail()
   }
   if(hasOrder&&s.page_order!=='downThenOver'&&s.page_order!=='overThenDown')return fail()
   if((s.paper!=='Letter'&&s.paper!=='A4')||(s.orientation!=='portrait'&&s.orientation!=='landscape')||!Number.isInteger(s.scale)||Number(s.scale)<10||Number(s.scale)>400)return fail()
   for(const key of ['left_inches','right_inches','top_inches','bottom_inches',...hasBands])if(typeof s[key]!=='number'||!Number.isFinite(s[key])||Number(s[key])<0||Number(s[key])>20)return fail()
   settings=s as unknown as NativeSheetPageConfigV1
  }
  let margins:NativeSheetPageMarginsV1|undefined
  if(status==='margins-only'){
   const hasBands=['header_inches','footer_inches'].filter(key=>!!o.margins&&typeof o.margins==='object'&&Object.hasOwn(o.margins,key))
   const m=exact(o.margins,['left_inches','right_inches','top_inches','bottom_inches',...hasBands])
   for(const key of ['left_inches','right_inches','top_inches','bottom_inches',...hasBands])if(typeof m[key]!=='number'||!Number.isFinite(m[key])||Number(m[key])<0||Number(m[key])>20)return fail()
   margins=m as unknown as NativeSheetPageMarginsV1
  }
  let defaulted:NativeSheetPageDefaultedFactV1[]|undefined
  if(hasDefaulted){
   // A defaulted fact only means anything beside the settings it names, and a
   // repeat would let one attribute be defaulted twice over.
   const list=o.defaulted
   if(status!=='available'||!Array.isArray(list)||list.length<1||list.length>3||new Set(list).size!==list.length)return fail()
   if(list.some(fact=>fact!=='paper'&&fact!=='orientation'&&fact!=='scale'))return fail()
   defaulted=list as NativeSheetPageDefaultedFactV1[]
  }
  return {sheet_id:o.sheet_id,sheet_part:o.sheet_part,status,warnings:o.warnings as string[],...(settings?{settings}:{}),...(margins?{margins}:{}),...(defaulted?{defaulted}:{})}
 })
}

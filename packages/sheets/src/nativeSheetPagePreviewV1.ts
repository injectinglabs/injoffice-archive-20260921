import {compiledNativeSheetGeometrySourcePart,isCompiledNativeSheetGeometryV2,isCompiledNativeStoredRowSheetGeometryV1,type NativeSheetGeometryV2,type NativeSheetGeometryRectV2} from './nativeSheetGeometryV2.js'
import {decodeNativeWorkbookObjectsV1,type NativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'
import {decodeNativeSheetPageSettingsV1,type NativeSheetPageConfigV1} from './nativeSheetPageSettingsV1.js'
import {snapshotNativePlainData} from './nativePlainData.js'

export interface NativeSheetHostPagePolicyV1 extends NativeSheetPageConfigV1 {
 kind:'explicit-host-page-policy-v1'
}
export interface NativeSheetPagePreviewOptionsV1 {repeat_print_titles:true}
export interface NativeSheetPreviewRegionV1 {
 kind:'body'|'repeat-rows'|'repeat-columns'|'repeat-corner';
 source_clip:NativeSheetGeometryRectV2;
 translate_x_emu:number;translate_y_emu:number;
 rows:{start:number;end:number};columns:{start:number;end:number};
}
export interface NativeSheetPreviewPageV1 {
 number:number;
 width_emu:number;height_emu:number;
 content_clip:NativeSheetGeometryRectV2;
 source_clip:NativeSheetGeometryRectV2;
 /** Map viewport-local native paint into this page: x * scale + translate_x. */
 scale:number;translate_x_emu:number;translate_y_emu:number;
 rows:{start:number;end:number};columns:{start:number;end:number};
 /** Explicit source-title repetition only. Paint each region once with this page's scale. */
 regions?:NativeSheetPreviewRegionV1[];
}
export interface NativeSheetPagePreviewV1 {
 protocol:'injoffice.xlsx.selected-range-pages';version:1;
 fidelity:'approximate';read_only:true;
 document_id:string;sheet_id:string;source_revision:string;source_package_sha256:string;
 geometry_sha256:string;
 policy:'whole-bands-down-then-over-v1'|'whole-bands-over-then-down-v1';
 settings_origin:'source'|'explicit-host';settings:NativeSheetPageConfigV1;
 warnings:string[];pages:NativeSheetPreviewPageV1[];
}

/** Read-only pagination of an explicitly selected viewport, not Excel's print engine.
 * Geometry must originate from exact-font, source-bound native compilation.
 * Host overrides are explicit choices and never treated as authored settings.
 */
export function compileNativeSheetPagePreviewV1(
 geometry:NativeSheetGeometryV2,objects:NativeWorkbookObjectsV1,hostPolicy?:NativeSheetHostPagePolicyV1,
 options?:NativeSheetPagePreviewOptionsV1,
):NativeSheetPagePreviewV1 {
 if(!isCompiledNativeSheetGeometryV2(geometry)&&!isCompiledNativeStoredRowSheetGeometryV1(geometry))throw new TypeError('Page preview requires compiled source-qualified sheet geometry')
 const source=decodeNativeWorkbookObjectsV1(objects,geometry.source_package_sha256)
 if(options!==undefined){
  const value=snapshotNativePlainData(options,{maxDepth:2,maxNodes:4}) as Record<string,unknown>
  if(!value||Array.isArray(value)||Object.keys(value).length!==1||value.repeat_print_titles!==true)throw new TypeError('Repeated print titles require the explicit repeat_print_titles: true option')
 }
 const titles=options===undefined?undefined:source.print_titles?.find(s=>s.sheet_id===geometry.sheet_id)
 if(options!==undefined&&(!titles||titles.sheet_part!==compiledNativeSheetGeometrySourcePart(geometry)||titles.status!=='available'))throw new TypeError('Saved print titles unavailable or do not join the source worksheet part')
 const candidates=source.page_settings?.filter(s=>s.sheet_id===geometry.sheet_id)??[]
 if(candidates.length!==1)throw new TypeError('Page settings do not join the source worksheet')
 const pageSettings=candidates[0]!
 if(pageSettings.sheet_part!==compiledNativeSheetGeometrySourcePart(geometry))throw new TypeError('Page settings do not join the source worksheet part')
 let settings:NativeSheetPageConfigV1
 if(hostPolicy!==undefined){
  hostPolicy=snapshotNativePlainData(hostPolicy,{maxDepth:4,maxNodes:64}) as NativeSheetHostPagePolicyV1
  const copy=Object.getOwnPropertyDescriptors(hostPolicy)
  if(!copy.kind||!('value'in copy.kind)||copy.kind.value!=='explicit-host-page-policy-v1'||Object.keys(copy).length!==8+Number(Object.hasOwn(copy,'page_order'))+Number(Object.hasOwn(copy,'fit_to_page'))||Object.values(copy).some(d=>!('value'in d)))throw new TypeError('Host page choices must be explicit plain data')
  const {kind:_,...config}=hostPolicy
  settings=decodeNativeSheetPageSettingsV1([{sheet_id:pageSettings.sheet_id,sheet_part:pageSettings.sheet_part,status:'available',settings:config,warnings:['Explicit host choices']}])[0]!.settings!
 }else{
  if(pageSettings.status!=='available'||!pageSettings.settings)throw new TypeError('Source page settings unavailable; an explicit host page policy is required')
  settings=pageSettings.settings
 }
 // A4 is exactly 210 by 297 mm; Letter is exactly 8.5 by 11 inches.
 const size=settings.paper==='A4'?[7560000,10692000]:[7772400,10058400]
 const [width,height]=settings.orientation==='landscape'?[size[1]!,size[0]!]:[size[0]!,size[1]!]
 const inch=(n:number)=>Math.round(n*914400)
 const left=inch(settings.left_inches),right=inch(settings.right_inches),top=inch(settings.top_inches),bottom=inch(settings.bottom_inches)
 const cw=width-left-right,ch=height-top-bottom
 let scale=settings.scale/100
 if(cw<=0||ch<=0)throw new RangeError('Page margins leave no printable area')
 const split=(bands:readonly {index:number;at:number;length:number}[],capacity:number)=>{
  const result:{start:number;end:number;at:number;length:number}[]=[]
  let current:typeof result[number]|undefined
  for(const b of bands){
   if(b.length===0)continue
   if(b.length>capacity)throw new RangeError('A source row or column exceeds one page; no silent clipping or rescaling')
   if(!current||b.at+b.length-current.at>capacity){current={start:b.index,end:b.index,at:b.at,length:b.length};result.push(current)}
   else {current.end=b.index;current.length=b.at+b.length-current.at}
  }
  return result
 }
 const allRows=geometry.rows.map(r=>({index:r.row,at:r.y_emu,length:r.height_emu}))
 const allColumns=geometry.columns.map(c=>({index:c.column,at:c.x_emu,length:c.width_emu}))
 const partition=(bands:typeof allRows,range:{start:number;end:number}|undefined)=>{
  if(!range)return {body:bands,title:undefined}
  if(range.start!==bands[0]?.index||range.end>=bands[bands.length-1]!.index)throw new RangeError('Saved print titles must lead the selected range and leave body rows or columns')
  const selected=bands.filter(b=>b.index<=range.end),body=bands.filter(b=>b.index>range.end)
  const visible=selected.filter(b=>b.length>0)
  if(!visible.length||!body.some(b=>b.length>0))throw new RangeError('Saved print titles and body must each include visible rows or columns')
  const first=visible[0]!,last=visible[visible.length-1]!
  return {body,title:{start:range.start,end:range.end,at:first.at,length:last.at+last.length-first.at}}
 }
 const rp=partition(allRows,titles?.rows),cp=partition(allColumns,titles?.columns)
 const rowBands=rp.body,columnBands=cp.body,tr=rp.title,tc=cp.title
 const titleHeight=tr?.length??0,titleWidth=tc?.length??0
 if(titles)for(const {rect:r} of geometry.merged_ranges){
  if((tr&&r.y_emu<tr.at+tr.length&&r.y_emu+r.height_emu>tr.at+tr.length)||(tc&&r.x_emu<tc.at+tc.length&&r.x_emu+r.width_emu>tc.at+tc.length))throw new RangeError('A merged cell crosses a repeated-title region boundary')
 }
 const mergesFit=(rows:ReturnType<typeof split>,columns:ReturnType<typeof split>)=>geometry.merged_ranges.every(({rect:r})=>
  [...rows,...(tr?[tr]:[])].some(b=>r.y_emu>=b.at&&r.y_emu+r.height_emu<=b.at+b.length)&&[...columns,...(tc?[tc]:[])].some(b=>r.x_emu>=b.at&&r.x_emu+r.width_emu<=b.at+b.length))
 const fit=settings.fit_to_page
 if(fit){
  // Bounded, explicit approximation: greatest whole-percent shrink satisfying
  // actual whole-band pagination. Source percentage is retained but not applied.
  let found=false
  for(let percent=100;percent>=10;percent--){
   const candidate=percent/100,rc=Math.floor(ch/candidate)-titleHeight,cc=Math.floor(cw/candidate)-titleWidth
   if(rowBands.some(b=>b.length>rc)||columnBands.some(b=>b.length>cc))continue
   const r=split(rowBands,rc),c=split(columnBands,cc)
   if((fit.height===0||r.length<=fit.height)&&(fit.width===0||c.length<=fit.width)&&r.length*c.length<=100&&mergesFit(r,c)){scale=candidate;found=true;break}
  }
  if(!found)throw new RangeError('Fit-to-page target cannot be met between 10% and 100% within the 100-page preview budget without splitting merged cells')
 }
 const rows=split(rowBands,Math.floor(ch/scale)-titleHeight)
 const columns=split(columnBands,Math.floor(cw/scale)-titleWidth)
 if(rows.length*columns.length>100)throw new RangeError('Worksheet page preview exceeds 100 pages')
 if(!mergesFit(rows,columns))throw new RangeError('A merged cell crosses a preview page boundary')
 const pages:NativeSheetPreviewPageV1[]=[]
 const addPage=(c:typeof columns[number],r:typeof rows[number])=>{
  const region=(kind:NativeSheetPreviewRegionV1['kind'],x:typeof c,y:typeof r,dx:number,dy:number):NativeSheetPreviewRegionV1=>({kind,source_clip:{x_emu:x.at,y_emu:y.at,width_emu:x.length,height_emu:y.length},translate_x_emu:left+(dx-x.at)*scale,translate_y_emu:top+(dy-y.at)*scale,rows:{start:y.start,end:y.end},columns:{start:x.start,end:x.end}})
  const {kind:_,...body}=region('body',c,r,titleWidth,titleHeight)
  pages.push({number:pages.length+1,width_emu:width,height_emu:height,content_clip:{x_emu:left,y_emu:top,width_emu:cw,height_emu:ch},source_clip:body.source_clip,scale,translate_x_emu:body.translate_x_emu,translate_y_emu:body.translate_y_emu,rows:body.rows,columns:body.columns,...(titles?{regions:[{kind:'body' as const,...body},...(tr?[region('repeat-rows',c,tr,titleWidth,0)]:[]),...(tc?[region('repeat-columns',tc,r,0,titleHeight)]:[]),...(tr&&tc?[region('repeat-corner',tc,tr,0,0)]:[])]}: {})})
 }
 if(settings.page_order==='overThenDown'){for(const r of rows)for(const c of columns)addPage(c,r)}
 else {for(const c of columns)for(const r of rows)addPage(c,r)}
 const policy=settings.page_order==='overThenDown'?'whole-bands-over-then-down-v1':'whole-bands-down-then-over-v1'
 return {protocol:'injoffice.xlsx.selected-range-pages',version:1,fidelity:'approximate',read_only:true,document_id:geometry.document_id,sheet_id:geometry.sheet_id,source_revision:geometry.source_revision,source_package_sha256:geometry.source_package_sha256,geometry_sha256:geometry.geometry_sha256,policy,settings_origin:hostPolicy?'explicit-host':'source',settings,warnings:[...pageSettings.warnings,...(fit?[`Approximate fit-to-page: greatest whole-percent shrink from 100% to 10% meeting the selected-range whole-band targets. Effective scale is ${Math.round(scale*100)}%; stored percentage is not applied. This is not Excel's fit algorithm.`]:[]),...(isCompiledNativeStoredRowSheetGeometryV1(geometry)?['Stored row-height approximation: source descender metadata does not alter row boxes. Automatic text fitting and baselines are not qualified.']:[]),...(titles?[...titles.warnings,'Explicit source-title repetition reserves leading row and column bands on every page. Hosts must paint each returned region once, including the corner. Drawing repetition is not supplied. This is not Excel print fidelity.']:[]),(titles?'Only the supplied range is paginated. Saved leading print titles are repeated; headers and printer-specific layout are not reproduced. Chart and drawing paint is supplied separately by the host. Whole source rows/columns are kept together; this is not Excel pagination fidelity.':'Only the supplied range is paginated; saved print-area selection is a separate source-bound step. Chart and drawing paint is supplied separately by the host. Headers, repeated print titles and printer-specific layout are not reproduced. Whole source rows/columns are kept together; this is not Excel pagination fidelity.'),...(hostPolicy?['Paper, margins and scale are explicit host choices, not authored workbook settings.']:[])],pages}
}

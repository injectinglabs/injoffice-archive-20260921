import {isCompiledNativeSheetGeometryV2,isCompiledNativeStoredRowSheetGeometryV1,type NativeSheetGeometryV2,type NativeSheetGeometryRectV2} from './nativeSheetGeometryV2.js'
import {decodeNativeWorkbookObjectsV1,type NativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'
import {decodeNativeSheetPageSettingsV1,type NativeSheetPageConfigV1} from './nativeSheetPageSettingsV1.js'
import {snapshotNativePlainData} from './nativePlainData.js'

export interface NativeSheetHostPagePolicyV1 extends NativeSheetPageConfigV1 {
 kind:'explicit-host-page-policy-v1'
}
export interface NativeSheetPreviewPageV1 {
 number:number;
 width_emu:number;height_emu:number;
 content_clip:NativeSheetGeometryRectV2;
 source_clip:NativeSheetGeometryRectV2;
 /** Map viewport-local native paint into this page: x * scale + translate_x. */
 scale:number;translate_x_emu:number;translate_y_emu:number;
 rows:{start:number;end:number};columns:{start:number;end:number};
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
):NativeSheetPagePreviewV1 {
 if(!isCompiledNativeSheetGeometryV2(geometry)&&!isCompiledNativeStoredRowSheetGeometryV1(geometry))throw new TypeError('Page preview requires compiled source-qualified sheet geometry')
 const source=decodeNativeWorkbookObjectsV1(objects,geometry.source_package_sha256)
 const candidates=source.page_settings?.filter(s=>s.sheet_id===geometry.sheet_id)??[]
 if(candidates.length!==1)throw new TypeError('Page settings do not join the source worksheet')
 const pageSettings=candidates[0]!
 let settings:NativeSheetPageConfigV1
 if(hostPolicy!==undefined){
  hostPolicy=snapshotNativePlainData(hostPolicy,{maxDepth:4,maxNodes:64}) as NativeSheetHostPagePolicyV1
  const copy=Object.getOwnPropertyDescriptors(hostPolicy)
  if(!copy.kind||!('value'in copy.kind)||copy.kind.value!=='explicit-host-page-policy-v1'||Object.keys(copy).length!==(Object.hasOwn(copy,'page_order')?9:8)||Object.values(copy).some(d=>!('value'in d)))throw new TypeError('Host page choices must be explicit plain data')
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
 const cw=width-left-right,ch=height-top-bottom,scale=settings.scale/100
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
 const rows=split(geometry.rows.map(r=>({index:r.row,at:r.y_emu,length:r.height_emu})),Math.floor(ch/scale))
 const columns=split(geometry.columns.map(c=>({index:c.column,at:c.x_emu,length:c.width_emu})),Math.floor(cw/scale))
 if(rows.length*columns.length>100)throw new RangeError('Worksheet page preview exceeds 100 pages')
 for(const merge of geometry.merged_ranges){
  const r=merge.rect
  if(!rows.some(b=>r.y_emu>=b.at&&r.y_emu+r.height_emu<=b.at+b.length)||!columns.some(b=>r.x_emu>=b.at&&r.x_emu+r.width_emu<=b.at+b.length))throw new RangeError('A merged cell crosses a preview page boundary')
 }
 const pages:NativeSheetPreviewPageV1[]=[]
 const addPage=(c:typeof columns[number],r:typeof rows[number])=>pages.push({number:pages.length+1,width_emu:width,height_emu:height,content_clip:{x_emu:left,y_emu:top,width_emu:cw,height_emu:ch},source_clip:{x_emu:c.at,y_emu:r.at,width_emu:c.length,height_emu:r.length},scale,translate_x_emu:left-c.at*scale,translate_y_emu:top-r.at*scale,rows:{start:r.start,end:r.end},columns:{start:c.start,end:c.end}})
 if(settings.page_order==='overThenDown'){for(const r of rows)for(const c of columns)addPage(c,r)}
 else {for(const c of columns)for(const r of rows)addPage(c,r)}
 const policy=settings.page_order==='overThenDown'?'whole-bands-over-then-down-v1':'whole-bands-down-then-over-v1'
 return {protocol:'injoffice.xlsx.selected-range-pages',version:1,fidelity:'approximate',read_only:true,document_id:geometry.document_id,sheet_id:geometry.sheet_id,source_revision:geometry.source_revision,source_package_sha256:geometry.source_package_sha256,geometry_sha256:geometry.geometry_sha256,policy,settings_origin:hostPolicy?'explicit-host':'source',settings,warnings:[...pageSettings.warnings,...(isCompiledNativeStoredRowSheetGeometryV1(geometry)?['Stored row-height approximation: source descender metadata does not alter row boxes. Automatic text fitting and baselines are not qualified.']:[]),'Only the explicitly selected range is paginated. Charts, drawings, headers, print areas/titles and printer-specific layout are not reproduced. Whole source rows/columns are kept together; this is not Excel pagination fidelity.',...(hostPolicy?['Paper, margins and scale are explicit host choices, not authored workbook settings.']:[])],pages}
}

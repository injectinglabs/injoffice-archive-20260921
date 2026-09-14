import type {NativeDocxDocumentV1,NativeDocxDrawingV1} from './nativeContract.js'
export interface NativeDocxQualifiedInlineTextboxV1 { drawing_id:string; run_id:string; width_millipoints:number; height_millipoints:number; text:string; fill_rgb:string; line_rgb:string }
type InlineTextboxDrawing = NativeDocxDrawingV1 & {textbox_text?:string;textbox_fill_rgb?:string;textbox_line_rgb?:string}
export function qualifyNativeDocxInlineTextboxV1(_document:NativeDocxDocumentV1,runID:string,drawing:NativeDocxDrawingV1):{ok:true;value:NativeDocxQualifiedInlineTextboxV1}|{ok:false;message:string}{
 const source=drawing as InlineTextboxDrawing
 if(source.placement!=='inline'||source.textbox_text===undefined||source.width_emu<=0||source.height_emu<=0)return{ok:false,message:'Drawing is not a bounded inline textbox'}
 const fill=source.textbox_fill_rgb??'FFFFFF',line=source.textbox_line_rgb??'000000'
 if(!/^[0-9A-F]{6}$/.test(fill)||!/^[0-9A-F]{6}$/.test(line)||source.textbox_text.length>4096)return{ok:false,message:'Inline textbox paint metadata is invalid'}
 return{ok:true,value:{drawing_id:source.id,run_id:runID,width_millipoints:Math.round(source.width_emu*10000/9144),height_millipoints:Math.round(source.height_emu*10000/9144),text:source.textbox_text,fill_rgb:fill,line_rgb:line}}
}

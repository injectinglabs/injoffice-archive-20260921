import {decodeTextboxHardBreaks,decodeTextboxLinePaint,type NativeTextboxHardBreakLayoutV1,type NativeTextboxLinePaintV1} from './nativeTextboxHardBreaksV1.js'
import {decodeNativeDocxTextboxEvidenceV1,type NativeDocxTextboxV1} from './nativeTextboxInventoryV1.js'
import {sha256} from '@noble/hashes/sha2.js'
import {bytesToHex} from '@noble/hashes/utils.js'

export interface NativeDocxTextboxGeometryV1 {
 width_emu:number;height_emu:number;insets_emu:[number,number,number,number];fill_rgb:string;line_rgb:string;line_width_emu:number;font_family:string;font_size_half_points:number;text_rgb:string
}
export interface NativeDocxTextboxGeometryItemV1 {owner:NativeDocxTextboxV1;geometry:NativeDocxTextboxGeometryV1|null;hard_break_layout?:NativeTextboxHardBreakLayoutV1}
export interface NativeDocxTextboxGeometryEvidenceV1 {items:NativeDocxTextboxGeometryItemV1[];omitted_count:number}

/** Copy only bounded plain own data, without invoking accessors or toJSON. */
export function nativeTextboxGeometryPlainData(value:unknown):unknown {
 let nodes=0,bytes=0
 const visit=(value:unknown,depth=0):unknown=>{
 if(++nodes>50000)throw new TypeError("Textbox traversal budget")
 if(typeof value==='string'){bytes+=value.length;if(bytes>8000000)throw new TypeError("Textbox text budget")}
 if(depth>16)throw new TypeError('Textbox data nesting limit')
 if(value===null||typeof value==='string'||typeof value==='boolean'||typeof value==='number')return value
 if(!value||typeof value!=='object')throw new TypeError('Invalid textbox data')
 if(Array.isArray(value)){
  if(Object.getPrototypeOf(value)!==Array.prototype||value.length>100000||Reflect.ownKeys(value).length!==value.length+1)throw new TypeError('Invalid textbox array')
  return Array.from({length:value.length},(_,i)=>{const d=Object.getOwnPropertyDescriptor(value,String(i));if(!d||!('value'in d))throw new TypeError('Textbox accessors forbidden');return visit(d.value,depth+1)})
 }
 if(Object.getPrototypeOf(value)!==Object.prototype||Reflect.ownKeys(value).length>64)throw new TypeError('Invalid textbox record')
 const out:Record<string,unknown>={}
 for(const key of Reflect.ownKeys(value)){if(typeof key!=='string'||key==='__proto__')throw new TypeError('Invalid textbox key');const d=Object.getOwnPropertyDescriptor(value,key);if(!d||!('value'in d))throw new TypeError('Textbox accessors forbidden');out[key]=visit(d.value,depth+1)}
 return out
 }
 return visit(value)
}
function keys(value:unknown,wanted:string[]):asserts value is Record<string,unknown>{if(!value||typeof value!=='object'||Array.isArray(value)||Object.keys(value).sort().join(',')!==[...wanted].sort().join(','))throw new TypeError('Unexpected textbox geometry fields')}
export function nativeTextboxGeometryDigestV1(value:unknown):string{
 const sort=(v:unknown):unknown=>Array.isArray(v)?v.map(sort):v&&typeof v==='object'?Object.fromEntries(Object.entries(v).sort(([a],[b])=>a<b?-1:a>b?1:0).map(([k,x])=>[k,sort(x)])):v
 return 'sha256:'+bytesToHex(sha256(new TextEncoder().encode(JSON.stringify(sort(nativeTextboxGeometryPlainData(value))))))
}
export function decodeNativeDocxTextboxGeometryV1(source:unknown,input:unknown):NativeDocxTextboxGeometryEvidenceV1{
 if(input===undefined){decodeNativeDocxTextboxEvidenceV1(source,undefined);return {items:[],omitted_count:0}}
 const data=nativeTextboxGeometryPlainData(input);keys(data,['items','omitted_count'])
 if(!Array.isArray(data.items)||data.items.length>64)throw new TypeError('Textbox geometry budget')
 const raw=data.items.map(i=>{keys(i,['owner','geometry',...(i&&typeof i==='object'&&Object.hasOwn(i,'hard_break_layout')?['hard_break_layout']:[])]);return i})
 const joined=decodeNativeDocxTextboxEvidenceV1(source,{items:raw.map(i=>i.owner),omitted_count:data.omitted_count})
 return {items:raw.map((item,i)=>{
  const owner=joined.items[i]!,g=item.geometry
  if(owner.kind!=='drawingml')throw new TypeError('Geometry requires DrawingML')
  if(owner.status==='omitted'){if(g!==null||Object.hasOwn(item,'hard_break_layout'))throw new TypeError('Omitted textbox has geometry');return {owner,geometry:null}}
  keys(g,['width_emu','height_emu','insets_emu','fill_rgb','line_rgb','line_width_emu','font_family','font_size_half_points','text_rgb'])
  const integer=(v:unknown,min:number,max:number):v is number=>Number.isSafeInteger(v)&&Number(v)>=min&&Number(v)<=max
  if(!integer(g.width_emu,1,127000000)||!integer(g.height_emu,1,127000000)||!Array.isArray(g.insets_emu)||g.insets_emu.length!==4||!g.insets_emu.every(v=>integer(v,0,127000000)))throw new TypeError('Invalid shape bounds')
  if([g.width_emu,g.height_emu,...g.insets_emu].some(v=>v%127!==0))throw new TypeError('Geometry must be exactly representable')
  if(g.insets_emu[0]+g.insets_emu[2]>=g.width_emu||g.insets_emu[1]+g.insets_emu[3]>=g.height_emu)throw new TypeError('Empty textbox inset box')
  if(typeof g.fill_rgb!=='string'||!/^(none|[0-9A-F]{6})$/.test(g.fill_rgb)||typeof g.line_rgb!=='string'||!/^(none|[0-9A-F]{6})$/.test(g.line_rgb)||!integer(g.line_width_emu,0,127000)||(g.line_rgb==='none'?g.line_width_emu!==0:g.line_width_emu===0)||g.line_width_emu%127!==0||g.line_width_emu>=g.width_emu||g.line_width_emu>=g.height_emu)throw new TypeError('Invalid textbox paint')
  const hard=Object.hasOwn(item,'hard_break_layout')?decodeTextboxHardBreaks(item.hard_break_layout,owner):undefined
  if(typeof g.font_family!=='string'||!g.font_family.length||g.font_family.length>128||g.font_family.trim()!==g.font_family||!integer(g.font_size_half_points,1,400)||typeof g.text_rgb!=='string'||!/^[0-9A-F]{6}$/.test(g.text_rgb)||owner.paragraphs.length!==1||(!hard&&!/^[\x20-\x7e]{1,4096}$/.test(owner.paragraphs[0]!)))throw new TypeError('Invalid qualified textbox run')
  return {owner,geometry:g as unknown as NativeDocxTextboxGeometryV1,...(hard?{hard_break_layout:hard}:{})}
 }),omitted_count:joined.omitted_count}
}

export interface NativeDocxTextboxShapePaintV1 {
 protocol:'injoffice.docx.textbox-shape-paint';version:1;package_sha256:string;diagnostic_id:string;input_sha256:string;font_sha256:string;status:'supported'|'omitted';reason:string;paths:string[];width_millipoints:number;height_millipoints:number;line_width_millipoints:number;fill_rgb:string;line_rgb:string;text_rgb:string;shaper_revision:string;line_layout?:NativeTextboxLinePaintV1
}
/** Decode the Node helper's numeric/path-only output against this exact source. */
export function decodeNativeDocxTextboxShapePaintV1(source:unknown,evidence:unknown,index:number,input:unknown,expectedFontSHA256:string):NativeDocxTextboxShapePaintV1 {
 const joined=decodeNativeDocxTextboxGeometryV1(source,evidence),item=joined.items[index]
 if(!item||!Number.isSafeInteger(index))throw new TypeError('Unknown textbox index')
 const v=nativeTextboxGeometryPlainData(input);keys(v,['protocol','version','package_sha256','diagnostic_id','input_sha256','font_sha256','status','reason','paths','width_millipoints','height_millipoints','line_width_millipoints','fill_rgb','line_rgb','text_rgb','shaper_revision',...(v&&typeof v==='object'&&Object.hasOwn(v,'line_layout')?['line_layout']:[])])
 if(v.protocol!=='injoffice.docx.textbox-shape-paint'||v.version!==1||v.package_sha256!==item.owner.package_sha256||v.diagnostic_id!==item.owner.diagnostic_id||v.input_sha256!==nativeTextboxGeometryDigestV1(item)||v.font_sha256!==expectedFontSHA256||typeof v.font_sha256!=='string'||!/^sha256:[0-9a-f]{64}$/.test(v.font_sha256)||typeof v.reason!=='string'||v.reason.length>512||typeof v.shaper_revision!=='string'||v.shaper_revision.length>256||!Array.isArray(v.paths)||v.paths.length>16384)throw new TypeError('Textbox paint source mismatch')
 if(v.status==='omitted'){if(Object.hasOwn(v,'line_layout')||!v.reason||v.paths.length||v.width_millipoints!==0||v.height_millipoints!==0||v.line_width_millipoints!==0||v.fill_rgb!=='none'||v.line_rgb!=='none'||v.text_rgb!=='000000')throw new TypeError('Invalid textbox refusal')}
 else {
  const g=item.geometry;if(v.status!=='supported'||!g||v.reason!==''||v.width_millipoints!==Math.round(g.width_emu*10/127)||v.height_millipoints!==Math.round(g.height_emu*10/127)||v.line_width_millipoints!==Math.round(g.line_width_emu*10/127)||v.fill_rgb!==g.fill_rgb||v.line_rgb!==g.line_rgb||v.text_rgb!==g.text_rgb)throw new TypeError('Textbox paint geometry mismatch')
  const lines=item.hard_break_layout?decodeTextboxLinePaint(v.line_layout,item.hard_break_layout,g.insets_emu[1]*10/127,(g.height_emu-g.insets_emu[3])*10/127,v.paths.length,item.owner.paragraphs[0]!):undefined
  if(!lines&&Object.hasOwn(v,'line_layout'))throw new TypeError('Unexpected line layout')
  let pathIndex=0,budget=0;for(const path of v.paths){if(typeof path!=='string'||path.length===0||path.length>1000000||!/^[MLQCZ0-9 .,-]*$/.test(path))throw new TypeError('Invalid glyph path');const line=lines?.lines.find(l=>pathIndex>=l.path_start&&pathIndex<l.path_end);const top=line?g.insets_emu[1]*10/127+line.ordinal*item.hard_break_layout!.line_step_twips*50:undefined;validateTextboxPath(path,g,top,top===undefined?undefined:top+item.hard_break_layout!.line_step_twips*50);pathIndex++;budget+=path.length}if(budget>8000000)throw new TypeError('Glyph path budget')
 }
 return v as unknown as NativeDocxTextboxShapePaintV1
}

function validateTextboxPath(path:string,g:NativeDocxTextboxGeometryV1,lineTop?:number,lineBottom?:number):void {
 const tokens=path.match(/[MLQCZ]|-?\d+(?:\.\d+)?/g)??[]
 if(!tokens.some(t=>t==='L'||t==='Q'||t==='C')||tokens.join('')!==path.replace(/[ ,]/g,'')||tokens.length>200000)throw new TypeError('Invalid glyph path tokens')
 let i=0,started=false
 const [l,t,r,b]=g.insets_emu.map(n=>n*10/127),w=g.width_emu*10/127,h=g.height_emu*10/127
 while(i<tokens.length){const kind=tokens[i++]!,count=({M:2,L:2,Q:4,C:6,Z:0} as Record<string,number>)[kind];if(count===undefined||(!started&&kind!=='M'))throw new TypeError('Invalid glyph command');started=true
 for(let n=0;n<count;n+=2){const x=Number(tokens[i++]),y=Number(tokens[i++]);if(!Number.isSafeInteger(x)||!Number.isSafeInteger(y)||x<l!||x>w-r!||y<(lineTop??t!)||y>(lineBottom??(h-b!)))throw new TypeError('Glyph coordinates outside source inset box')}
 }
}
export function nativeTextboxFontDigestV1(bytes:Uint8Array):string{return 'sha256:'+bytesToHex(sha256(bytes))}

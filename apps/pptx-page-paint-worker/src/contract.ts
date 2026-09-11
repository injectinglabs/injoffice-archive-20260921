// Browser-safe transport. Never accepts HTML, CSS, remote URLs or SVG markup.
import {decodeNativeDocxPagePaintResourceListV1,type NativeDocxPagePaintMediaAssetV1} from '@injoffice/docs/native-raster'
export type PreviewMatrix = [number,number,number,number,number,number]
export type PreviewRect = {x:number;y:number;cx:number;cy:number}
export type PreviewStroke = {stroke?:string;strokeWidth?:number;strokeLinecap?:'butt'|'round'|'square';strokeLinejoin?:'round'|'bevel'|'miter';strokeMiterlimit?:number}
export type PreviewNode =
 | {kind:'group';transform:PreviewMatrix;clip?:PreviewRect;children:PreviewNode[];sourceRole?:'paragraphBullet'|'contentRun'|'connectorArrow'}
 | ({kind:'path';d:string;fill:string}&PreviewStroke)
 | ({kind:'rect';rect:PreviewRect;radius:number;fill:string}&PreviewStroke)
 | ({kind:'ellipse';rect:PreviewRect;fill:string}&PreviewStroke)
 | {kind:'placeholder';rect:PreviewRect;label:string}
 | {kind:'image';rect:PreviewRect;resourceId:string;crop?:{left:number;top:number;right:number;bottom:number}}
export interface PptxPreview {
 version:1; package_sha256:string;slide_index:number;slide_count:number;
 width:number;height:number;background:string;policy:'max-run-natural-v1';
 nodes:PreviewNode[];diagnostics:string[];font_digests:string[];resources:NativeDocxPagePaintMediaAssetV1[]
}
export function decodePptxPreview(value:unknown):PptxPreview {
 const fail=()=>{throw new TypeError('Native slide preview failed bounded validation.')}
 const record=(v:unknown):Record<string,unknown>=>{if(!v||typeof v!=='object'||Array.isArray(v))return fail();return v as Record<string,unknown>}
 const number=(v:unknown,min=-1e9,max=1e9)=>{if(typeof v!=='number'||!Number.isFinite(v)||v<min||v>max)fail()}
 const color=(v:unknown)=>{if(typeof v!=='string'||!/^([0-9a-f]{6}|none)$/i.test(v))fail()}
 const rect=(v:unknown)=>{const r=record(v);number(r.x);number(r.y);number(r.cx,0);number(r.cy,0)}
 let count=0,pathBytes=0
 const resources=decodeNativeDocxPagePaintResourceListV1(record(value).resources)
 const path=(d:string)=>{
  const tokens=d.match(/[MLQCZ]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g)??[]
  if(tokens.join('')!==d.replace(/[\s,]/g,''))fail()
  for(let i=0;i<tokens.length;){const arity=({M:2,L:2,Q:4,C:6,Z:0} as Record<string,number>)[tokens[i++]!];if(arity===undefined||i+arity>tokens.length)fail();for(let j=0;j<arity;j++)number(Number(tokens[i++]))}
 }
 const node=(v:unknown,depth:number)=>{if(++count>20000||depth>32)fail();const n=record(v)
  if(n.strokeLinecap!==undefined&&!['butt','round','square'].includes(String(n.strokeLinecap)))fail()
  if(n.strokeLinejoin!==undefined&&!['round','bevel','miter'].includes(String(n.strokeLinejoin)))fail()
  if(n.strokeMiterlimit!==undefined)number(n.strokeMiterlimit,1)
  switch(n.kind){
   case 'group':if(!Array.isArray(n.transform)||n.transform.length!==6||!Array.isArray(n.children)||n.sourceRole!==undefined&&!['paragraphBullet','contentRun','connectorArrow'].includes(String(n.sourceRole)))return fail();n.transform.forEach(v=>number(v));if(n.clip!==undefined)rect(n.clip);n.children.forEach(v=>node(v,depth+1));break
   case 'path':if(typeof n.d!=='string'||n.d.length>200000||!/^[MLQCZ0-9eE+.,\s-]*$/.test(n.d))return fail();path(n.d);pathBytes+=n.d.length;if(pathBytes>8e6)fail();color(n.fill);if(n.stroke!==undefined)color(n.stroke);if(n.strokeWidth!==undefined)number(n.strokeWidth,0);break
   case 'rect':number(n.radius,0);
   case 'ellipse':rect(n.rect);color(n.fill);if(n.stroke!==undefined)color(n.stroke);if(n.strokeWidth!==undefined)number(n.strokeWidth,0);break
   case 'placeholder':rect(n.rect);if(typeof n.label!=='string'||n.label.length>1024)fail();break
   case 'image':rect(n.rect);if(!resources.some(r=>r.id===n.resourceId))fail();if(n.crop!==undefined){const c=record(n.crop);for(const edge of ['left','right','top','bottom']){number(c[edge],0,100000);if(!Number.isInteger(c[edge]))fail()}if(Number(c.left)+Number(c.right)>=100000||Number(c.top)+Number(c.bottom)>=100000)fail()}break
   default:fail()
  }
 }
 const v=record(value)
 if(v.version!==1||typeof v.package_sha256!=='string'||!/^([a-f0-9]{64})$/.test(v.package_sha256)||v.policy!=='max-run-natural-v1')fail()
 number(v.width,1);number(v.height,1);number(v.slide_count,1,10000);number(v.slide_index,0,Number(v.slide_count)-1)
 if(!Number.isInteger(v.slide_count)||!Number.isInteger(v.slide_index))fail()
 color(v.background)
 if(!Array.isArray(v.nodes)||!Array.isArray(v.diagnostics)||v.diagnostics.length>2000||v.diagnostics.some(s=>typeof s!=='string'||s.length>2048)||!Array.isArray(v.font_digests)||v.font_digests.length>256||v.font_digests.some(s=>typeof s!=='string'||!/^sha256:[a-f0-9]{64}$/.test(s)))return fail()
 v.nodes.forEach(n=>node(n,0))
 return value as PptxPreview
}

import type {NativeDocxSourceAnchorV1} from './nativeContract.js'
import type {NativeDocxTextboxV1} from './nativeTextboxInventoryV1.js'

/** Authored page-relative coordinates, not a physical page assignment. */
export interface NativeTextboxPageAnchorV1 {
 policy:'page-offset-no-wrap-v1'
 source_anchor:NativeDocxSourceAnchorV1
 horizontal_anchor:NativeDocxSourceAnchorV1
 vertical_anchor:NativeDocxSourceAnchorV1
 x_emu:number
 y_emu:number
}
function keys(v:unknown,wanted:string[]):asserts v is Record<string,unknown>{
 if(!v||typeof v!=='object'||Array.isArray(v)||Object.keys(v).sort().join(',')!==wanted.sort().join(','))throw new TypeError('Invalid textbox page anchor fields')
}
function anchor(v:unknown,part:string):NativeDocxSourceAnchorV1{
 keys(v,['part_name','path','start_byte','end_byte','xml_sha256'])
 if(v.part_name!==part||typeof v.path!=='string'||!Number.isSafeInteger(v.start_byte)||!Number.isSafeInteger(v.end_byte)||Number(v.start_byte)<0||Number(v.end_byte)<=Number(v.start_byte)||typeof v.xml_sha256!=='string'||!/^sha256:[0-9a-f]{64}$/.test(v.xml_sha256))throw new TypeError('Invalid textbox position source')
 return v as unknown as NativeDocxSourceAnchorV1
}
/** Called after bounded own-data copying and source diagnostic joins. As with
 * geometry, hashes bind trusted inspection output; they do not verify raw XML. */
export function decodeTextboxPageAnchor(value:unknown,owner:NativeDocxTextboxV1):NativeTextboxPageAnchorV1{
 keys(value,['policy','source_anchor','horizontal_anchor','vertical_anchor','x_emu','y_emu'])
 if(value.policy!=='page-offset-no-wrap-v1'||[value.x_emu,value.y_emu].some(v=>!Number.isSafeInteger(v)||Number(v)<0||Number(v)>127000000||Number(v)%127!==0))throw new TypeError('Invalid textbox page offsets')
 const o=owner.anchor,s=anchor(value.source_anchor,o.part_name),h=anchor(value.horizontal_anchor,o.part_name),v=anchor(value.vertical_anchor,o.part_name)
 const root=o.path.match(/^(.*\/w:drawing\[[1-9][0-9]*\])/u)?.[1]
 const match=root&&s.path.slice(root.length).match(/^\/(wp|ns[0-9a-f]{8}):anchor\[1\]$/)
 if(!root||!s.path.startsWith(root)||!match)throw new TypeError('Invalid textbox floating container')
 const contains=(a:NativeDocxSourceAnchorV1,b:NativeDocxSourceAnchorV1)=>b.path.startsWith(a.path+'/')&&b.start_byte>a.start_byte&&b.end_byte<a.end_byte
 if(o.path===s.path?o.start_byte!==s.start_byte||o.end_byte!==s.end_byte||o.xml_sha256!==s.xml_sha256:!contains(o,s)&&!contains(s,o))throw new TypeError('Textbox position diagnostic mismatch')
 const prefix=match[1]!
 if(h.path!==s.path+`/${prefix}:positionH[1]/${prefix}:posOffset[1]`||v.path!==s.path+`/${prefix}:positionV[1]/${prefix}:posOffset[1]`||!contains(s,h)||!contains(s,v)||h.end_byte>=v.start_byte||(contains(s,o)&&v.end_byte>=o.start_byte))throw new TypeError('Invalid textbox position source order')
 return {policy:'page-offset-no-wrap-v1',source_anchor:s,horizontal_anchor:h,vertical_anchor:v,x_emu:Number(value.x_emu),y_emu:Number(value.y_emu)}
}

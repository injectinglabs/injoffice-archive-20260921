import {expect,it} from 'vitest'
import {qualifySymbolBullet} from './symbolBullet.js'

// Original tiny table-directory evidence fixture, not a bundled font.
function evidence() {
 const bytes=new Uint8Array(700), v=new DataView(bytes.buffer)
 const u16=(p:number,n:number)=>v.setUint16(p,n),u32=(p:number,n:number)=>v.setUint32(p,n)
 u32(0,0x10000);u16(4,6)
 for(const [i,tag,offset,length] of [[0,'OS/2',108,68],[1,'maxp',176,6],[2,'cmap',182,314],[3,'head',496,54],[4,'hhea',550,36],[5,'hmtx',586,40]] as const){bytes.set([...tag].map(c=>c.charCodeAt(0)),12+i*16);u32(20+i*16,offset);u32(24+i*16,length)}
 bytes[140]=5;u16(172,0xf020);u16(174,0xf0ff);u16(180,10)
 u16(184,2);u16(186,1);u16(188,0);u32(190,20);u16(194,3);u16(196,0);u32(198,282)
 u16(202,0);u16(204,262);bytes[208+0x71]=7
 const s=464;u16(s,4);u16(s+2,32);u16(s+6,4);u16(s+14,0xf071);u16(s+16,0xffff);u16(s+20,0xf071);u16(s+22,0xffff);u16(s+24,(7-0xf071)&0xffff);u16(s+26,1)
 u16(514,1000);u16(584,10);u16(614,600)
 u32(176,0x10000);u32(496,0x10000);u32(550,0x10000)
 return bytes
}
it('qualifies exact Mac byte and Windows symbol glyph agreement without changing the source character',()=>{
 expect(qualifySymbolBullet(evidence(),'q')).toEqual({policy:'windows-symbol-byte-v1',sourceByte:113,transportCodePoint:0xf071,glyphId:7,unitsPerEm:1000,advanceWidth:600,ascender:0,descender:0,lineGap:0})
})
it('refuses wrong first-character, glyph disagreement, missing/zero glyphs and malformed bounded tables',()=>{
 for(const mutate of [(b:Uint8Array)=>b[173]=0x21,(b:Uint8Array)=>b[321]=8,(b:Uint8Array)=>b[321]=0,(b:Uint8Array)=>b[185]=255,(b:Uint8Array)=>b[181]=7,(b:Uint8Array)=>b[140]=2,(b:Uint8Array)=>b[465]=12,(b:Uint8Array)=>b[493]=2,(b:Uint8Array)=>b[493]=3,(b:Uint8Array)=>b[23]=0,(b:Uint8Array)=>b[39]=108,(b:Uint8Array)=>b[585]=0,(b:Uint8Array)=>b[107]=1]){
  const bytes=evidence();mutate(bytes);expect(()=>qualifySymbolBullet(bytes,'q')).toThrow()
 }
 for(const text of ['', 'qq', '\u0001','🙂'])expect(()=>qualifySymbolBullet(evidence(),text)).toThrow()
})
it('uses the last long horizontal metric and requires every trailing bearing byte',()=>{
 const bytes=evidence(),v=new DataView(bytes.buffer)
 v.setUint16(584,2);v.setUint16(590,600)
 expect(qualifySymbolBullet(bytes,'q').advanceWidth).toBe(600)
 v.setUint32(104,23)
 expect(()=>qualifySymbolBullet(bytes,'q')).toThrow()
})
it('rejects directory aliases, overlapping cmap ranges and contradictory OS2 range/version',()=>{
 for(const mutate of [(v:DataView)=>v.setUint32(190,4),(v:DataView)=>v.setUint32(198,20),(v:DataView)=>v.setUint16(174,0xf020),(v:DataView)=>v.setUint16(108,6),(v:DataView)=>v.setUint32(550,0)]){
  const bytes=evidence();mutate(new DataView(bytes.buffer));expect(()=>qualifySymbolBullet(bytes,'q')).toThrow()
 }
})
it('refuses fonts advertising shaping or variation tables rather than bypassing their semantics',()=>{
 for(const tag of ['GSUB','GPOS','kerx','morx','mort','fvar','gvar','HVAR']) {
  const original=evidence(), bytes=new Uint8Array(original.length+16)
  bytes.set(original.subarray(0,108));bytes.set(original.subarray(108),124)
  const v=new DataView(bytes.buffer);v.setUint16(4,7)
  for(let i=0;i<6;i++)v.setUint32(20+i*16,v.getUint32(20+i*16)+16)
  bytes.set([...tag].map(c=>c.charCodeAt(0)),108);v.setUint32(116,bytes.length)
  expect(()=>qualifySymbolBullet(bytes,'q')).toThrow()
 }
})

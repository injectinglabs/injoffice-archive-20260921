/** Microsoft OpenType cmap symbol encoding, not Unicode substitution.
 * https://learn.microsoft.com/en-us/typography/opentype/spec/cmap
 * Only the measured legacy Mac-byte / Windows F020 convention is qualified. */
export function qualifySymbolBullet(bytes: Uint8Array, character: string) {
  const fail = (): never => { throw new Error('symbol bullet font encoding is not qualified') }
  if (character.length !== 1 || character.charCodeAt(0) < 0x20 || character.charCodeAt(0) > 0x7e || bytes.length < 12 || bytes.length > 16 * 1024 * 1024) fail()
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const u16 = (offset: number) => { if (offset < 0 || offset + 2 > bytes.length) fail(); return view.getUint16(offset) }
  const u32 = (offset: number) => { if (offset < 0 || offset + 4 > bytes.length) fail(); return view.getUint32(offset) }
  if (u32(0) !== 0x00010000) fail()
  const count = u16(4), tables = new Map<string, {offset:number,length:number}>()
  if (count < 1 || count > 128 || 12 + count * 16 > bytes.length) fail()
  for (let i = 0; i < count; i++) {
    const p = 12 + i * 16, tag = String.fromCharCode(...bytes.subarray(p,p+4)), offset=u32(p+8), length=u32(p+12)
    if (tables.has(tag) || offset < 12+count*16 || offset > bytes.length || length > bytes.length-offset) fail()
    for (const prior of tables.values()) if (length && prior.length && offset < prior.offset+prior.length && prior.offset < offset+length) fail()
    tables.set(tag,{offset,length})
  }
  const cmap=tables.get('cmap'), os2=tables.get('OS/2'), maxp=tables.get('maxp'), head=tables.get('head'), hhea=tables.get('hhea'), hmtx=tables.get('hmtx')
  if (!cmap || !os2 || !maxp || !head || !hhea || !hmtx || os2.length<68 || maxp.length<6 || cmap.length<4 || head.length<54 || hhea.length<36) return fail()
  if (['fvar','gvar','HVAR','GSUB','GPOS','kerx','morx','mort'].some(tag=>tables.has(tag))) fail()
  // Do not generalize the first-character offset to arbitrary symbol layouts.
  if (u16(os2.offset)>5 || u32(head.offset)!==0x10000 || u32(hhea.offset)!==0x10000 || u32(maxp.offset)!==0x10000 || u16(os2.offset+64)!==0xf020 || u16(os2.offset+66)>0xf0ff || u16(os2.offset+66)<0xf000+character.charCodeAt(0) || bytes[os2.offset+32]!==5) fail()
  const numGlyphs=u16(maxp.offset+4), encodingCount=u16(cmap.offset+2)
  if (u16(cmap.offset)!==0 || encodingCount<1 || encodingCount>128 || 4+encodingCount*8>cmap.length) fail()
  let macGlyph: number|undefined, symbolGlyph: number|undefined
  const recognizedRanges: {start:number,end:number}[]=[]
  const range = (start:number,length:number) => { if(recognizedRanges.some(r=>start<r.end&&r.start<start+length))fail();recognizedRanges.push({start,end:start+length}) }
  const sourceByte=character.charCodeAt(0), transportCodePoint=0xf000+sourceByte
  for (let i=0;i<encodingCount;i++) {
    const p=cmap.offset+4+i*8, platform=u16(p), encoding=u16(p+2), relative=u32(p+4)
    if (relative<4+encodingCount*8 || relative>cmap.length-2) fail()
    const s=cmap.offset+relative, format=u16(s)
    if (platform===1 && encoding===0) {
      if (macGlyph!==undefined || format!==0 || relative+262>cmap.length || u16(s+2)!==262) fail()
      range(s,262)
      macGlyph=bytes[s+6+sourceByte]
    }
    if (platform===3 && encoding===0) {
      if (symbolGlyph!==undefined || format!==4 || relative+16>cmap.length) fail()
      const length=u16(s+2), countX2=u16(s+6), n=countX2/2
      if (length<24 || relative+length>cmap.length || countX2<2 || countX2%2 || 16+8*n>length || u16(s+14+2*n)!==0) fail()
      range(s,length)
      let prior=-1; symbolGlyph=0
      for(let j=0;j<n;j++) {
        const end=u16(s+14+2*j), start=u16(s+16+2*n+2*j), delta=u16(s+16+4*n+2*j), at=s+16+6*n+2*j, range=u16(at)
        if(start>end || start<=prior) fail(); prior=end
        if(range && (range%2 || at+range<s+16+8*n || at+range+2*(end-start)+2>s+length)) fail()
        if(transportCodePoint<start || transportCodePoint>end) continue
        if(range===0) symbolGlyph=(transportCodePoint+delta)&0xffff
        else { const g=at+range+2*(transportCodePoint-start); if(range%2 || g<s+16+8*n || g+2>s+length) fail(); const id=u16(g); symbolGlyph=id===0?0:(id+delta)&0xffff }
      }
      if(prior!==0xffff) fail()
    }
  }
  if (!macGlyph || !symbolGlyph || macGlyph!==symbolGlyph || symbolGlyph>=numGlyphs) return fail()
  const unitsPerEm=u16(head.offset+18), numberOfHMetrics=u16(hhea.offset+34)
  if(unitsPerEm<16 || unitsPerEm>16384 || numberOfHMetrics<1 || numberOfHMetrics>numGlyphs || hmtx.length<4*numberOfHMetrics+2*(numGlyphs-numberOfHMetrics)) fail()
  const advanceWidth=u16(hmtx.offset+4*Math.min(symbolGlyph,numberOfHMetrics-1))
  if(!advanceWidth) fail()
  const ascender=view.getInt16(hhea.offset+4),descender=view.getInt16(hhea.offset+6),lineGap=view.getInt16(hhea.offset+8)
  return {policy:'windows-symbol-byte-v1' as const,sourceByte,transportCodePoint,glyphId:symbolGlyph,unitsPerEm,advanceWidth,ascender,descender,lineGap}
}

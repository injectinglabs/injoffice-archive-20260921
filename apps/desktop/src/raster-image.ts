/** Bounded header sizing only; native image writers fully validate and decode before accepting bytes. */
export function imageDimensions(bytes: Uint8Array): { width: number; height: number; contentType: 'image/png' | 'image/jpeg' } {
  if (bytes.byteLength > 2 * 1024 * 1024) throw new Error('Choose a PNG or JPEG no larger than 2 MiB for this version.');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength); let width = 0, height = 0; let contentType: 'image/png' | 'image/jpeg' = 'image/png';
  if (bytes.length >= 24 && [137,80,78,71,13,10,26,10].every((n, i) => bytes[i] === n) && view.getUint32(12) === 0x49484452) { width = view.getUint32(16); height = view.getUint32(20); }
  else if (bytes.length >= 4 && bytes[0] === 255 && bytes[1] === 216) {
    contentType = 'image/jpeg'; let offset = 2;
    while (offset + 4 <= bytes.length) {
      if (bytes[offset] !== 255) break;
      while (bytes[offset] === 255) offset++;
      const marker = bytes[offset++]!; if (marker === 0xda || marker === 0xd9 || offset + 2 > bytes.length) break;
      const length = view.getUint16(offset); if (length < 2 || offset + length > bytes.length) break;
      if ([0xc0, 0xc1, 0xc2].includes(marker) && length >= 8) { height = view.getUint16(offset + 3); width = view.getUint16(offset + 5); break; }
      offset += length;
    }
  }
  if (width < 1 || height < 1 || width * height > 16_000_000) throw new Error('Choose a valid PNG/JPEG image up to 16 megapixels.');
  return { width, height, contentType };
}

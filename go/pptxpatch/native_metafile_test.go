package pptxpatch

import (
	"bytes"
	"encoding/binary"
	"errors"
	"image"
	"image/png"
	"strings"
	"testing"
)

// metafileBuilder assembles the byte-exact MS-WMF record stream these tests
// play, so every fixture states its own records instead of hiding them in a
// binary blob.
type metafileBuilder struct {
	objects int
	records []byte
}

func (b *metafileBuilder) record(function uint16, payload []byte) *metafileBuilder {
	if len(payload)%2 != 0 {
		payload = append(append([]byte{}, payload...), 0)
	}
	var head [6]byte
	binary.LittleEndian.PutUint32(head[0:], uint32(len(payload)/2+3))
	binary.LittleEndian.PutUint16(head[4:], function)
	b.records = append(b.records, head[:]...)
	b.records = append(b.records, payload...)
	return b
}

func words(values ...uint16) []byte {
	out := make([]byte, len(values)*2)
	for index, value := range values {
		binary.LittleEndian.PutUint16(out[index*2:], value)
	}
	return out
}

func dword(value uint32) []byte {
	out := make([]byte, 4)
	binary.LittleEndian.PutUint32(out, value)
	return out
}

func (b *metafileBuilder) mapMode(mode uint16) *metafileBuilder {
	return b.record(nativeMetafileRecordSetMapMode, words(mode))
}

func (b *metafileBuilder) windowOrg(x, y int16) *metafileBuilder {
	return b.record(nativeMetafileRecordSetWindowOrg, words(uint16(y), uint16(x)))
}

func (b *metafileBuilder) windowExt(width, height int16) *metafileBuilder {
	return b.record(nativeMetafileRecordSetWindowExt, words(uint16(height), uint16(width)))
}

// brush creates a solid brush and takes the next free object handle.
func (b *metafileBuilder) brush(red, green, blue byte) *metafileBuilder {
	b.objects++
	payload := append(words(nativeMetafileBrushSolid), dword(uint32(red)|uint32(green)<<8|uint32(blue)<<16)...)
	return b.record(nativeMetafileRecordCreateBrushIndirect, append(payload, words(0)...))
}

func (b *metafileBuilder) selectObject(handle uint16) *metafileBuilder {
	return b.record(nativeMetafileRecordSelectObject, words(handle))
}

func (b *metafileBuilder) patBlt(x, y, width, height int16, rop uint32) *metafileBuilder {
	return b.record(nativeMetafileRecordPatBlt, append(dword(rop), words(uint16(height), uint16(width), uint16(y), uint16(x))...))
}

// dib24 builds an uncompressed bottom-up 24bpp BITMAPINFOHEADER bitmap from
// top-down RGB triples.
func dib24(width, height int, topDownRGB []byte) []byte {
	out := make([]byte, 40)
	binary.LittleEndian.PutUint32(out[0:], 40)
	binary.LittleEndian.PutUint32(out[4:], uint32(width))
	binary.LittleEndian.PutUint32(out[8:], uint32(height))
	binary.LittleEndian.PutUint16(out[12:], 1)
	binary.LittleEndian.PutUint16(out[14:], 24)
	stride := ((width*24 + 31) / 32) * 4
	pixels := make([]byte, stride*height)
	for row := 0; row < height; row++ {
		line := pixels[(height-1-row)*stride:]
		for column := 0; column < width; column++ {
			source := (row*width + column) * 3
			line[column*3] = topDownRGB[source+2]
			line[column*3+1] = topDownRGB[source+1]
			line[column*3+2] = topDownRGB[source]
		}
	}
	return append(out, pixels...)
}

// dib1 builds an uncompressed bottom-up monochrome bitmap; bits selects colour
// zero or colour one per pixel, top-down.
func dib1(width, height int, bits []byte, zero, one [3]byte) []byte {
	out := make([]byte, 40)
	binary.LittleEndian.PutUint32(out[0:], 40)
	binary.LittleEndian.PutUint32(out[4:], uint32(width))
	binary.LittleEndian.PutUint32(out[8:], uint32(height))
	binary.LittleEndian.PutUint16(out[12:], 1)
	binary.LittleEndian.PutUint16(out[14:], 1)
	out = append(out, zero[2], zero[1], zero[0], 0, one[2], one[1], one[0], 0)
	stride := ((width*1 + 31) / 32) * 4
	pixels := make([]byte, stride*height)
	for row := 0; row < height; row++ {
		line := pixels[(height-1-row)*stride:]
		for column := 0; column < width; column++ {
			if bits[row*width+column] != 0 {
				line[column/8] |= 1 << (7 - uint(column%8))
			}
		}
	}
	return append(out, pixels...)
}

func (b *metafileBuilder) stretchBlt(rop uint32, srcX, srcY, srcW, srcH, dstX, dstY, dstW, dstH int16, dib []byte) *metafileBuilder {
	payload := append(dword(rop), words(uint16(srcH), uint16(srcW), uint16(srcY), uint16(srcX), uint16(dstH), uint16(dstW), uint16(dstY), uint16(dstX))...)
	return b.record(nativeMetafileRecordDIBStretchBlt, append(payload, dib...))
}

func (b *metafileBuilder) bitBlt(rop uint32, srcX, srcY, dstX, dstY, width, height int16, dib []byte) *metafileBuilder {
	payload := append(dword(rop), words(uint16(srcY), uint16(srcX), uint16(height), uint16(width), uint16(dstY), uint16(dstX))...)
	return b.record(nativeMetafileRecordDIBBitBlt, append(payload, dib...))
}

func (b *metafileBuilder) build() []byte {
	b.record(nativeMetafileRecordEOF, nil)
	header := make([]byte, 18)
	binary.LittleEndian.PutUint16(header[0:], 1) // memory metafile
	binary.LittleEndian.PutUint16(header[2:], 9) // header size in words
	binary.LittleEndian.PutUint16(header[4:], 0x0300)
	binary.LittleEndian.PutUint32(header[6:], uint32(18+len(b.records))/2)
	binary.LittleEndian.PutUint16(header[10:], uint16(b.objects))
	return append(header, b.records...)
}

// baseline is the state prologue every control snapshot opens with.
func baseline(width, height int16) *metafileBuilder {
	b := &metafileBuilder{}
	return b.mapMode(nativeMetafileMapModeAnisotropic).windowOrg(0, 0).windowExt(width, height)
}

func pixelAt(t *testing.T, raster nativeMetafileRaster, x, y int) [3]byte {
	t.Helper()
	if x < 0 || y < 0 || x >= raster.Width || y >= raster.Height {
		t.Fatalf("pixel %d,%d is outside the %dx%d raster", x, y, raster.Width, raster.Height)
	}
	at := (y*raster.Width + x) * 3
	return [3]byte{raster.Pix[at], raster.Pix[at+1], raster.Pix[at+2]}
}

func decodeOrFatal(t *testing.T, data []byte) nativeMetafileRaster {
	t.Helper()
	raster, err := decodeNativeMetafile(data)
	if err != nil {
		t.Fatalf("decode refused unexpectedly: %v", err)
	}
	return raster
}

func TestNativeMetafileSizesFromWindowExtent(t *testing.T) {
	raster := decodeOrFatal(t, baseline(203, 54).build())
	if raster.Width != 203 || raster.Height != 54 {
		t.Fatalf("raster is %dx%d, want 203x54", raster.Width, raster.Height)
	}
	// An untouched snapshot is opaque white, not transparent: a metafile has
	// no alpha and the picture beneath it must not show through.
	if got := pixelAt(t, raster, 0, 0); got != [3]byte{0xFF, 0xFF, 0xFF} {
		t.Fatalf("empty raster pixel is %v, want opaque white", got)
	}
}

func TestNativeMetafilePatBltPaintsSelectedBrush(t *testing.T) {
	data := baseline(20, 10).
		brush(0x20, 0x40, 0x60).
		selectObject(0).
		patBlt(4, 2, 6, 3, nativeMetafileROPPatCopy).
		build()
	raster := decodeOrFatal(t, data)
	if got := pixelAt(t, raster, 4, 2); got != [3]byte{0x20, 0x40, 0x60} {
		t.Fatalf("filled pixel is %v, want the brush colour", got)
	}
	if got := pixelAt(t, raster, 9, 4); got != [3]byte{0x20, 0x40, 0x60} {
		t.Fatalf("last filled pixel is %v, want the brush colour", got)
	}
	// The rectangle is half-open: x+width and y+height are outside it.
	if got := pixelAt(t, raster, 10, 4); got != [3]byte{0xFF, 0xFF, 0xFF} {
		t.Fatalf("pixel past the fill width is %v, want white", got)
	}
	if got := pixelAt(t, raster, 4, 5); got != [3]byte{0xFF, 0xFF, 0xFF} {
		t.Fatalf("pixel past the fill height is %v, want white", got)
	}
}

func TestNativeMetafilePatBltHonoursWindowOrigin(t *testing.T) {
	builder := baseline(20, 10).brush(0x11, 0x22, 0x33).selectObject(0)
	builder.windowOrg(3, 1).patBlt(3, 1, 2, 2, nativeMetafileROPPatCopy)
	raster := decodeOrFatal(t, builder.build())
	if got := pixelAt(t, raster, 0, 0); got != [3]byte{0x11, 0x22, 0x33} {
		t.Fatalf("origin-shifted fill landed at %v, want it at the raster origin", got)
	}
}

func TestNativeMetafileSaveRestoreRewindsBrush(t *testing.T) {
	builder := baseline(8, 4).brush(0x00, 0x00, 0x00).selectObject(0)
	builder.record(nativeMetafileRecordSaveDC, nil)
	builder.brush(0xFF, 0x00, 0x00).selectObject(1)
	builder.record(nativeMetafileRecordRestoreDC, words(0xFFFF))
	builder.patBlt(0, 0, 2, 2, nativeMetafileROPPatCopy)
	raster := decodeOrFatal(t, builder.build())
	if got := pixelAt(t, raster, 0, 0); got != [3]byte{0x00, 0x00, 0x00} {
		t.Fatalf("restored brush painted %v, want the brush selected before the save", got)
	}
}

func TestNativeMetafileBitBltCopiesTrueColourBitmap(t *testing.T) {
	source := []byte{
		10, 20, 30, 40, 50, 60,
		70, 80, 90, 100, 110, 120,
	}
	data := baseline(8, 4).bitBlt(nativeMetafileROPSrcCopy, 0, 0, 1, 1, 2, 2, dib24(2, 2, source)).build()
	raster := decodeOrFatal(t, data)
	if got := pixelAt(t, raster, 1, 1); got != [3]byte{10, 20, 30} {
		t.Fatalf("top-left blitted pixel is %v, want 10,20,30", got)
	}
	if got := pixelAt(t, raster, 2, 2); got != [3]byte{100, 110, 120} {
		t.Fatalf("bottom-right blitted pixel is %v, want 100,110,120", got)
	}
	if got := pixelAt(t, raster, 0, 0); got != [3]byte{0xFF, 0xFF, 0xFF} {
		t.Fatalf("pixel outside the destination is %v, want white", got)
	}
}

// The masked-icon pattern PowerPoint writes: AND a monochrome mask down, then
// OR the colour image in. Where the mask bit is one the background survives;
// where it is zero the colour image lands.
func TestNativeMetafileMaskedIconComposites(t *testing.T) {
	mask := []byte{1, 0, 1, 0}
	colour := []byte{0, 0, 0, 200, 100, 50, 0, 0, 0, 0, 0, 0}
	builder := baseline(4, 4).
		brush(0x80, 0x80, 0x80).selectObject(0).patBlt(0, 0, 4, 4, nativeMetafileROPPatCopy).
		stretchBlt(nativeMetafileROPSrcAnd, 0, 0, 2, 2, 0, 0, 2, 2, dib1(2, 2, mask, [3]byte{0, 0, 0}, [3]byte{255, 255, 255})).
		stretchBlt(nativeMetafileROPSrcPaint, 0, 0, 2, 2, 0, 0, 2, 2, dib24(2, 2, colour))
	raster := decodeOrFatal(t, builder.build())
	// Mask bit one -> white AND grey = grey, then OR black leaves it grey.
	if got := pixelAt(t, raster, 0, 0); got != [3]byte{0x80, 0x80, 0x80} {
		t.Fatalf("masked-through pixel is %v, want the untouched background", got)
	}
	// Mask bit zero -> black AND grey = black, then OR the colour paints it.
	if got := pixelAt(t, raster, 1, 0); got != [3]byte{200, 100, 50} {
		t.Fatalf("masked-in pixel is %v, want the icon colour", got)
	}
}

func TestNativeMetafileStretchBltScalesNearest(t *testing.T) {
	source := []byte{1, 1, 1, 250, 250, 250, 1, 1, 1, 250, 250, 250}
	data := baseline(8, 8).stretchBlt(nativeMetafileROPSrcCopy, 0, 0, 2, 2, 0, 0, 4, 4, dib24(2, 2, source)).build()
	raster := decodeOrFatal(t, data)
	for _, probe := range []struct {
		x, y int
		want [3]byte
	}{{0, 0, [3]byte{1, 1, 1}}, {1, 1, [3]byte{1, 1, 1}}, {2, 0, [3]byte{250, 250, 250}}, {3, 3, [3]byte{250, 250, 250}}} {
		if got := pixelAt(t, raster, probe.x, probe.y); got != probe.want {
			t.Fatalf("stretched pixel %d,%d is %v, want %v", probe.x, probe.y, got, probe.want)
		}
	}
}

func TestNativeMetafileOpaqueTextRectangleIsPaintedAndGlyphsAreDisclosed(t *testing.T) {
	builder := baseline(20, 10)
	builder.record(nativeMetafileRecordSetBkColor, dword(0x0000FF))
	builder.record(nativeMetafileRecordExtTextOut, append(
		words(0, 0, 2, nativeMetafileETOOpaque, 2, 2, 8, 6), []byte("hi")...))
	raster := decodeOrFatal(t, builder.build())
	if got := pixelAt(t, raster, 3, 3); got != [3]byte{0xFF, 0x00, 0x00} {
		t.Fatalf("opaque text background is %v, want the background colour", got)
	}
	if raster.TextOmitted != 1 {
		t.Fatalf("TextOmitted is %d, want 1 so the caller can disclose the missing glyphs", raster.TextOmitted)
	}
}

// A text record with no characters carries no glyph limit to disclose; this is
// how a snapshot states a plain interior fill.
func TestNativeMetafileEmptyTextRecordDisclosesNothing(t *testing.T) {
	builder := baseline(20, 10)
	builder.record(nativeMetafileRecordExtTextOut, words(0, 0, 0, nativeMetafileETOOpaque, 2, 2, 8, 6))
	raster := decodeOrFatal(t, builder.build())
	if raster.TextOmitted != 0 {
		t.Fatalf("TextOmitted is %d for a record with no characters, want 0", raster.TextOmitted)
	}
}

func TestNativeMetafileRefusesUnmodeledConstructs(t *testing.T) {
	polygon := baseline(8, 8)
	polygon.record(0x0324, words(3, 0, 0, 4, 0, 2, 4))

	hatchBrush := &metafileBuilder{objects: 1}
	hatchBrush.mapMode(nativeMetafileMapModeAnisotropic).windowExt(8, 8)
	hatchBrush.record(nativeMetafileRecordCreateBrushIndirect, append(append(words(2), dword(0)...), words(3)...))

	badROP := baseline(8, 8).brush(0, 0, 0).selectObject(0).patBlt(0, 0, 2, 2, 0x005A0049)

	compressed := dib24(2, 2, make([]byte, 12))
	binary.LittleEndian.PutUint32(compressed[16:], 1) // BI_RLE8
	rle := baseline(8, 8).bitBlt(nativeMetafileROPSrcCopy, 0, 0, 0, 0, 2, 2, compressed)

	paletted := dib24(2, 2, make([]byte, 12))
	binary.LittleEndian.PutUint16(paletted[14:], 8)
	eightBit := baseline(8, 8).bitBlt(nativeMetafileROPSrcCopy, 0, 0, 0, 0, 2, 2, paletted)

	isotropic := (&metafileBuilder{}).mapMode(1).windowExt(8, 8)

	rescale := baseline(8, 8)
	rescale.windowExt(16, 16)

	unselected := baseline(8, 8).selectObject(0)

	unfilled := baseline(8, 8).patBlt(0, 0, 2, 2, nativeMetafileROPPatCopy)

	rotated := baseline(8, 8)
	rotatedFont := make([]byte, 18)
	binary.LittleEndian.PutUint16(rotatedFont[4:], 900)
	rotated.record(nativeMetafileRecordCreateFontIndirect, append(rotatedFont, "Arial\x00"...))
	rotated.objects++

	blitRop := baseline(8, 8).bitBlt(0x00330008, 0, 0, 0, 0, 2, 2, dib24(2, 2, make([]byte, 12)))

	noExtent := (&metafileBuilder{}).mapMode(nativeMetafileMapModeAnisotropic)

	oversized := (&metafileBuilder{}).mapMode(nativeMetafileMapModeAnisotropic).windowExt(4097, 8)

	for _, tc := range []struct {
		name string
		data []byte
		want string
	}{
		{"polygon record", polygon.build(), "0x0324"},
		{"hatched brush", hatchBrush.build(), "BS_SOLID"},
		{"unmodeled pattern rop", badROP.build(), "PATCOPY"},
		{"compressed bitmap", rle.build(), "BI_RGB"},
		{"paletted bitmap", eightBit.build(), "bits per pixel"},
		{"non-anisotropic map mode", isotropic.build(), "MM_ANISOTROPIC"},
		{"window rescale", rescale.build(), "rescales"},
		{"unknown object handle", unselected.build(), "never created"},
		{"fill without a brush", unfilled.build(), "before selecting a brush"},
		{"rotated font", rotated.build(), "rotates text"},
		{"unmodeled bitmap rop", blitRop.build(), "SRCCOPY"},
		{"no window extent", noExtent.build(), "no window extent"},
		{"oversized extent", oversized.build(), "exceeds"},
		{"truncated header", []byte{1, 0}, "shorter than its header"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			_, err := decodeNativeMetafile(tc.data)
			if err == nil {
				t.Fatalf("decode accepted %s, want a refusal", tc.name)
			}
			var refusal nativeMetafileRefusal
			if !errors.As(err, &refusal) {
				t.Fatalf("error is %T, want nativeMetafileRefusal so the caller can disclose it", err)
			}
			if !strings.Contains(refusal.reason, tc.want) {
				t.Fatalf("refusal %q does not name %q", refusal.reason, tc.want)
			}
		})
	}
}

func TestNativeMetafileRefusesRecordStreamWithoutEOF(t *testing.T) {
	builder := baseline(8, 8)
	data := append(make([]byte, 18), builder.records...)
	binary.LittleEndian.PutUint16(data[0:], 1)
	binary.LittleEndian.PutUint16(data[2:], 9)
	if _, err := decodeNativeMetafile(data); err == nil || !strings.Contains(err.Error(), "META_EOF") {
		t.Fatalf("decode accepted a stream with no META_EOF: %v", err)
	}
}

func TestNativeMetafilePNGRoundTripsThePixels(t *testing.T) {
	data := baseline(5, 3).
		brush(0x12, 0x34, 0x56).selectObject(0).patBlt(1, 1, 3, 2, nativeMetafileROPPatCopy).
		build()
	raster := decodeOrFatal(t, data)
	encoded := encodeNativeMetafilePNG(raster)
	decoded, err := png.Decode(bytes.NewReader(encoded))
	if err != nil {
		t.Fatalf("the encoder produced a PNG no decoder accepts: %v", err)
	}
	if decoded.Bounds() != (image.Rectangle{Max: image.Point{X: 5, Y: 3}}) {
		t.Fatalf("decoded PNG is %v, want 5x3", decoded.Bounds())
	}
	for y := 0; y < raster.Height; y++ {
		for x := 0; x < raster.Width; x++ {
			red, green, blue, alpha := decoded.At(x, y).RGBA()
			want := pixelAt(t, raster, x, y)
			if byte(red>>8) != want[0] || byte(green>>8) != want[1] || byte(blue>>8) != want[2] || alpha>>8 != 0xFF {
				t.Fatalf("PNG pixel %d,%d is %d,%d,%d alpha %d, want %v opaque", x, y, red>>8, green>>8, blue>>8, alpha>>8, want)
			}
		}
	}
}

// The extractor states this asset's digest and the server re-derives the bytes
// before serving them, so the encoding has to be a fixed function of the pixels.
func TestNativeMetafilePNGEncodingIsDeterministic(t *testing.T) {
	raster := decodeOrFatal(t, baseline(64, 64).brush(1, 2, 3).selectObject(0).patBlt(0, 0, 64, 64, nativeMetafileROPPatCopy).build())
	first := encodeNativeMetafilePNG(raster)
	if !bytes.Equal(first, encodeNativeMetafilePNG(raster)) {
		t.Fatal("two encodings of one raster differ, so the stated digest cannot be re-derived")
	}
	if !bytes.HasPrefix(first, []byte{0x89, 'P', 'N', 'G', '\r', '\n', 0x1A, '\n'}) {
		t.Fatal("encoded bytes do not open with the PNG signature")
	}
	// A raster wider than one stored DEFLATE block must still decode, which is
	// the case the block-splitting loop exists for.
	if _, err := png.Decode(bytes.NewReader(first)); err != nil {
		t.Fatalf("multi-block PNG does not decode: %v", err)
	}
}

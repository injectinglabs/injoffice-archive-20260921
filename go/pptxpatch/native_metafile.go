package pptxpatch

import (
	"encoding/binary"
	"fmt"
	"hash/crc32"
)

// PowerPoint snapshots an embedded ActiveX control as a Windows Metafile
// (MS-WMF) and stores it as the mc:Fallback picture for that control. The
// native preview contract admits PNG and JPEG only, so a metafile reaches the
// painter only if this package turns it into pixels first.
//
// MS-WMF is the whole GDI drawing surface, and this decoder deliberately
// implements one closed corner of it: the records PowerPoint actually writes
// into a control snapshot. That corner is device-independent bitmap blits,
// solid pattern fills, and the device-context state those two consume. Every
// other record, and every parameter outside the modeled range, refuses by name
// rather than painting an approximation, because a control snapshot that is
// quietly missing a drawing operation is indistinguishable from a correct one.
//
// What is modeled:
//
//	META_SETMAPMODE (MM_ANISOTROPIC only), META_SETWINDOWORG, META_SETWINDOWEXT,
//	META_SAVEDC, META_RESTOREDC, META_SETBKMODE, META_SETBKCOLOR,
//	META_SETTEXTCOLOR, META_SETTEXTALIGN, META_SETSTRETCHBLTMODE,
//	META_CREATEBRUSHINDIRECT (BS_SOLID only), META_CREATEFONTINDIRECT,
//	META_SELECTOBJECT, META_DELETEOBJECT, META_PATBLT (PATCOPY only),
//	META_DIBBITBLT and META_DIBSTRETCHBLT (SRCCOPY, SRCAND, SRCPAINT over
//	uncompressed 1bpp and 24bpp DIBs), META_EXTTEXTOUT (its ETO_OPAQUE
//	background rectangle only), META_EOF.
//
// What is not modeled, and therefore refuses: every other record type, pens and
// lines, curves and regions, clipping, palettes, escapes, raster operations
// outside the three above, compressed or paletted-beyond-monochrome DIBs, and
// rotated or sheared text. Glyphs are a separate, disclosed limit — see
// nativeMetafileRaster.TextOmitted.

const (
	nativeMetafileRecordEOF                 = 0x0000
	nativeMetafileRecordSaveDC              = 0x001E
	nativeMetafileRecordSetBkMode           = 0x0102
	nativeMetafileRecordSetMapMode          = 0x0103
	nativeMetafileRecordSetStretchBltMode   = 0x0107
	nativeMetafileRecordRestoreDC           = 0x0127
	nativeMetafileRecordSelectObject        = 0x012D
	nativeMetafileRecordSetTextAlign        = 0x012E
	nativeMetafileRecordDeleteObject        = 0x01F0
	nativeMetafileRecordSetBkColor          = 0x0201
	nativeMetafileRecordSetTextColor        = 0x0209
	nativeMetafileRecordSetWindowOrg        = 0x020B
	nativeMetafileRecordSetWindowExt        = 0x020C
	nativeMetafileRecordCreateFontIndirect  = 0x02FB
	nativeMetafileRecordCreateBrushIndirect = 0x02FC
	nativeMetafileRecordPatBlt              = 0x061D
	nativeMetafileRecordDIBBitBlt           = 0x0940
	nativeMetafileRecordExtTextOut          = 0x0A32
	nativeMetafileRecordDIBStretchBlt       = 0x0B41
)

// The three ternary raster operations a control snapshot uses. PATCOPY paints
// the selected brush; SRCAND then SRCPAINT is the classic two-pass masked icon
// (AND the monochrome mask down, OR the colour image in).
const (
	nativeMetafileROPPatCopy  = 0x00F00021
	nativeMetafileROPSrcCopy  = 0x00CC0020
	nativeMetafileROPSrcAnd   = 0x008800C6
	nativeMetafileROPSrcPaint = 0x00EE0086
)

const (
	nativeMetafileMapModeAnisotropic = 8
	nativeMetafileBrushSolid         = 0
	nativeMetafileETOOpaque          = 0x0002

	// A control snapshot is a widget, not a photograph. These bounds keep a
	// hostile or corrupt metafile from turning into an unbounded allocation
	// while leaving every real snapshot far inside them.
	nativeMetafileMaxExtent  = 4096
	nativeMetafileMaxRecords = 65536
	nativeMetafileMaxObjects = 1024
)

// nativeMetafileRefusal names a metafile construct this decoder declines. It is
// a refusal, never a fallback: the caller reports the reason and paints nothing
// rather than painting a snapshot with a record missing from it.
type nativeMetafileRefusal struct {
	reason string
}

func (refusal nativeMetafileRefusal) Error() string { return refusal.reason }

func refuseNativeMetafile(format string, args ...any) error {
	return nativeMetafileRefusal{reason: fmt.Sprintf(format, args...)}
}

// nativeMetafileRaster is a decoded control snapshot: opaque RGB pixels at the
// metafile's own logical size, plus the glyph limit that applies to it.
type nativeMetafileRaster struct {
	Width  int
	Height int
	// Pix is Width*Height*3 bytes, red then green then blue, row-major from
	// the top-left. A metafile has no alpha channel, and a control snapshot
	// always paints its own background, so an opaque buffer loses nothing.
	Pix []byte
	// TextOmitted counts META_EXTTEXTOUT records that carried characters. This
	// decoder paints a text record's opaque background but not its glyphs:
	// rasterising glyphs needs a font rasteriser and font bytes, neither of
	// which exists on this side of the contract. The count exists so the
	// caller can disclose the omission on the element instead of shipping a
	// caption-less control as if it were complete.
	TextOmitted int
}

type nativeMetafileObject struct {
	kind  int // 0 unused, 1 brush, 2 font
	color [3]byte
}

type nativeMetafileState struct {
	brush                  *[3]byte
	bkColor                [3]byte
	textColor              [3]byte
	windowOrgX, windowOrgY int32
}

// decodeNativeMetafile rasterises a WMF control snapshot. It returns a refusal
// (errors.As against nativeMetafileRefusal) for anything outside the modeled
// corner described above.
func decodeNativeMetafile(data []byte) (nativeMetafileRaster, error) {
	body, err := nativeMetafileBody(data)
	if err != nil {
		return nativeMetafileRaster{}, err
	}
	records, objectCount, err := nativeMetafileRecords(body)
	if err != nil {
		return nativeMetafileRaster{}, err
	}
	width, height, err := nativeMetafileExtent(records)
	if err != nil {
		return nativeMetafileRaster{}, err
	}
	raster := nativeMetafileRaster{Width: width, Height: height, Pix: make([]byte, width*height*3)}
	for index := range raster.Pix {
		raster.Pix[index] = 0xFF
	}
	if objectCount > nativeMetafileMaxObjects {
		return nativeMetafileRaster{}, refuseNativeMetafile("metafile declares %d objects, beyond the modeled table", objectCount)
	}
	objects := make([]nativeMetafileObject, objectCount)
	state := nativeMetafileState{bkColor: [3]byte{0xFF, 0xFF, 0xFF}}
	var saved []nativeMetafileState
	for _, record := range records {
		if err := nativeMetafilePlay(&raster, record, &state, &saved, objects); err != nil {
			return nativeMetafileRaster{}, err
		}
	}
	return raster, nil
}

// nativeMetafileBody strips an Aldus placeable header when one is present and
// validates the metafile header that follows.
func nativeMetafileBody(data []byte) ([]byte, error) {
	offset := 0
	if len(data) >= 22 && binary.LittleEndian.Uint32(data) == 0x9AC6CDD7 {
		offset = 22
	}
	if len(data) < offset+18 {
		return nil, refuseNativeMetafile("metafile is shorter than its header")
	}
	header := data[offset:]
	metafileType := binary.LittleEndian.Uint16(header[0:])
	headerSize := binary.LittleEndian.Uint16(header[2:])
	if metafileType != 1 && metafileType != 2 {
		return nil, refuseNativeMetafile("metafile type %d is not a memory or disk metafile", metafileType)
	}
	if headerSize != 9 {
		return nil, refuseNativeMetafile("metafile header is %d words, not the 9 MS-WMF requires", headerSize)
	}
	return data[offset:], nil
}

type nativeMetafileRecord struct {
	function uint16
	payload  []byte
}

func nativeMetafileRecords(header []byte) ([]nativeMetafileRecord, int, error) {
	objectCount := int(binary.LittleEndian.Uint16(header[10:]))
	at := 18
	records := make([]nativeMetafileRecord, 0, 32)
	for {
		if at+6 > len(header) {
			return nil, 0, refuseNativeMetafile("metafile record stream ends without META_EOF")
		}
		size := binary.LittleEndian.Uint32(header[at:])
		function := binary.LittleEndian.Uint16(header[at+4:])
		if size < 3 {
			return nil, 0, refuseNativeMetafile("metafile record declares %d words, fewer than a record header", size)
		}
		if size > uint32(len(header)/2) || at+int(size)*2 > len(header) {
			return nil, 0, refuseNativeMetafile("metafile record runs past the end of the metafile")
		}
		records = append(records, nativeMetafileRecord{function: function, payload: header[at+6 : at+int(size)*2]})
		if len(records) > nativeMetafileMaxRecords {
			return nil, 0, refuseNativeMetafile("metafile carries more than %d records", nativeMetafileMaxRecords)
		}
		if function == nativeMetafileRecordEOF {
			return records, objectCount, nil
		}
		at += int(size) * 2
	}
}

// nativeMetafileExtent reads the logical size the metafile paints into. The
// snapshot is rasterised at exactly that size, so the picture's own a:xfrm
// does all the scaling, the same way it would for an authored PNG.
func nativeMetafileExtent(records []nativeMetafileRecord) (int, int, error) {
	for _, record := range records {
		if record.function != nativeMetafileRecordSetWindowExt {
			continue
		}
		if len(record.payload) < 4 {
			return 0, 0, refuseNativeMetafile("META_SETWINDOWEXT is truncated")
		}
		height := int(int16(binary.LittleEndian.Uint16(record.payload[0:])))
		width := int(int16(binary.LittleEndian.Uint16(record.payload[2:])))
		if width <= 0 || height <= 0 {
			return 0, 0, refuseNativeMetafile("metafile window extent %dx%d is empty or mirrored", width, height)
		}
		if width > nativeMetafileMaxExtent || height > nativeMetafileMaxExtent {
			return 0, 0, refuseNativeMetafile("metafile window extent %dx%d exceeds the %d pixel bound", width, height, nativeMetafileMaxExtent)
		}
		return width, height, nil
	}
	return 0, 0, refuseNativeMetafile("metafile states no window extent")
}

func nativeMetafileColor(value uint32) [3]byte {
	return [3]byte{byte(value), byte(value >> 8), byte(value >> 16)}
}

func nativeMetafilePoint(payload []byte, offset int) (int32, int32, bool) {
	if offset+4 > len(payload) {
		return 0, 0, false
	}
	y := int32(int16(binary.LittleEndian.Uint16(payload[offset:])))
	x := int32(int16(binary.LittleEndian.Uint16(payload[offset+2:])))
	return x, y, true
}

func nativeMetafileAllocate(objects []nativeMetafileObject, object nativeMetafileObject) error {
	for index := range objects {
		if objects[index].kind == 0 {
			objects[index] = object
			return nil
		}
	}
	return refuseNativeMetafile("metafile creates more objects than its header reserves")
}

func nativeMetafilePlay(raster *nativeMetafileRaster, record nativeMetafileRecord, state *nativeMetafileState, saved *[]nativeMetafileState, objects []nativeMetafileObject) error {
	payload := record.payload
	switch record.function {
	case nativeMetafileRecordEOF:
		return nil
	case nativeMetafileRecordSaveDC:
		if len(*saved) > nativeMetafileMaxObjects {
			return refuseNativeMetafile("metafile nests META_SAVEDC beyond the modeled depth")
		}
		*saved = append(*saved, *state)
		return nil
	case nativeMetafileRecordRestoreDC:
		if len(*saved) == 0 {
			return refuseNativeMetafile("metafile restores a device context it never saved")
		}
		*state = (*saved)[len(*saved)-1]
		*saved = (*saved)[:len(*saved)-1]
		return nil
	case nativeMetafileRecordSetMapMode:
		if len(payload) < 2 {
			return refuseNativeMetafile("META_SETMAPMODE is truncated")
		}
		mode := int16(binary.LittleEndian.Uint16(payload))
		if mode != nativeMetafileMapModeAnisotropic {
			return refuseNativeMetafile("metafile map mode %d is outside the modeled MM_ANISOTROPIC", mode)
		}
		return nil
	case nativeMetafileRecordSetWindowOrg:
		x, y, ok := nativeMetafilePoint(payload, 0)
		if !ok {
			return refuseNativeMetafile("META_SETWINDOWORG is truncated")
		}
		state.windowOrgX, state.windowOrgY = x, y
		return nil
	case nativeMetafileRecordSetWindowExt:
		if len(payload) < 4 {
			return refuseNativeMetafile("META_SETWINDOWEXT is truncated")
		}
		height := int(int16(binary.LittleEndian.Uint16(payload[0:])))
		width := int(int16(binary.LittleEndian.Uint16(payload[2:])))
		// The raster is sized from the first extent. A metafile that rescales
		// its window mid-stream is a different drawing model than the one
		// modeled here, so it refuses rather than painting at the wrong scale.
		if width != raster.Width || height != raster.Height {
			return refuseNativeMetafile("metafile rescales its window to %dx%d mid-stream", width, height)
		}
		return nil
	case nativeMetafileRecordSetBkMode, nativeMetafileRecordSetStretchBltMode:
		// Background mode only affects glyph backgrounds, which this decoder
		// does not paint; stretch mode only affects resampling, and every
		// modeled blit is sampled nearest-neighbour.
		return nil
	case nativeMetafileRecordSetTextAlign:
		return nil
	case nativeMetafileRecordSetBkColor:
		if len(payload) < 4 {
			return refuseNativeMetafile("META_SETBKCOLOR is truncated")
		}
		state.bkColor = nativeMetafileColor(binary.LittleEndian.Uint32(payload))
		return nil
	case nativeMetafileRecordSetTextColor:
		if len(payload) < 4 {
			return refuseNativeMetafile("META_SETTEXTCOLOR is truncated")
		}
		state.textColor = nativeMetafileColor(binary.LittleEndian.Uint32(payload))
		return nil
	case nativeMetafileRecordCreateBrushIndirect:
		if len(payload) < 8 {
			return refuseNativeMetafile("META_CREATEBRUSHINDIRECT is truncated")
		}
		style := binary.LittleEndian.Uint16(payload)
		if style != nativeMetafileBrushSolid {
			return refuseNativeMetafile("metafile brush style %d is outside the modeled BS_SOLID", style)
		}
		return nativeMetafileAllocate(objects, nativeMetafileObject{kind: 1, color: nativeMetafileColor(binary.LittleEndian.Uint32(payload[2:]))})
	case nativeMetafileRecordCreateFontIndirect:
		if len(payload) < 18 {
			return refuseNativeMetafile("META_CREATEFONTINDIRECT is truncated")
		}
		escapement := int16(binary.LittleEndian.Uint16(payload[4:]))
		orientation := int16(binary.LittleEndian.Uint16(payload[6:]))
		if escapement != 0 || orientation != 0 {
			return refuseNativeMetafile("metafile rotates text by %d/%d tenths of a degree", escapement, orientation)
		}
		// The font is tracked so its handle slot stays in step with the
		// metafile's own object table. Its glyphs are never painted.
		return nativeMetafileAllocate(objects, nativeMetafileObject{kind: 2})
	case nativeMetafileRecordSelectObject:
		if len(payload) < 2 {
			return refuseNativeMetafile("META_SELECTOBJECT is truncated")
		}
		index := int(binary.LittleEndian.Uint16(payload))
		if index >= len(objects) || objects[index].kind == 0 {
			return refuseNativeMetafile("metafile selects object %d, which it never created", index)
		}
		if objects[index].kind == 1 {
			color := objects[index].color
			state.brush = &color
		}
		return nil
	case nativeMetafileRecordDeleteObject:
		if len(payload) < 2 {
			return refuseNativeMetafile("META_DELETEOBJECT is truncated")
		}
		index := int(binary.LittleEndian.Uint16(payload))
		if index >= len(objects) {
			return refuseNativeMetafile("metafile deletes object %d, which is outside its table", index)
		}
		objects[index] = nativeMetafileObject{}
		return nil
	case nativeMetafileRecordPatBlt:
		return nativeMetafilePatBlt(raster, payload, state)
	case nativeMetafileRecordExtTextOut:
		return nativeMetafileExtTextOut(raster, payload, state)
	case nativeMetafileRecordDIBBitBlt:
		return nativeMetafileBlt(raster, payload, false)
	case nativeMetafileRecordDIBStretchBlt:
		return nativeMetafileBlt(raster, payload, true)
	default:
		return refuseNativeMetafile("metafile record 0x%04X is outside the modeled control-snapshot subset", record.function)
	}
}

func nativeMetafilePatBlt(raster *nativeMetafileRaster, payload []byte, state *nativeMetafileState) error {
	if len(payload) < 12 {
		return refuseNativeMetafile("META_PATBLT is truncated")
	}
	rop := binary.LittleEndian.Uint32(payload)
	if rop != nativeMetafileROPPatCopy {
		return refuseNativeMetafile("metafile pattern raster operation 0x%08X is outside the modeled PATCOPY", rop)
	}
	height := int32(int16(binary.LittleEndian.Uint16(payload[4:])))
	width := int32(int16(binary.LittleEndian.Uint16(payload[6:])))
	y := int32(int16(binary.LittleEndian.Uint16(payload[8:])))
	x := int32(int16(binary.LittleEndian.Uint16(payload[10:])))
	if state.brush == nil {
		return refuseNativeMetafile("metafile fills a pattern before selecting a brush")
	}
	nativeMetafileFill(raster, x-state.windowOrgX, y-state.windowOrgY, width, height, *state.brush)
	return nil
}

// nativeMetafileExtTextOut paints the opaque background rectangle a text record
// asks for and counts, but does not paint, its glyphs.
func nativeMetafileExtTextOut(raster *nativeMetafileRaster, payload []byte, state *nativeMetafileState) error {
	if len(payload) < 8 {
		return refuseNativeMetafile("META_EXTTEXTOUT is truncated")
	}
	count := int(binary.LittleEndian.Uint16(payload[4:]))
	options := binary.LittleEndian.Uint16(payload[6:])
	if options&nativeMetafileETOOpaque != 0 {
		if len(payload) < 16 {
			return refuseNativeMetafile("META_EXTTEXTOUT states ETO_OPAQUE without a rectangle")
		}
		left := int32(int16(binary.LittleEndian.Uint16(payload[8:])))
		top := int32(int16(binary.LittleEndian.Uint16(payload[10:])))
		right := int32(int16(binary.LittleEndian.Uint16(payload[12:])))
		bottom := int32(int16(binary.LittleEndian.Uint16(payload[14:])))
		nativeMetafileFill(raster, left-state.windowOrgX, top-state.windowOrgY, right-left, bottom-top, state.bkColor)
	}
	if count > 0 {
		raster.TextOmitted++
	}
	return nil
}

func nativeMetafileFill(raster *nativeMetafileRaster, x, y, width, height int32, color [3]byte) {
	x0, y0 := int(x), int(y)
	x1, y1 := x0+int(width), y0+int(height)
	if x0 < 0 {
		x0 = 0
	}
	if y0 < 0 {
		y0 = 0
	}
	if x1 > raster.Width {
		x1 = raster.Width
	}
	if y1 > raster.Height {
		y1 = raster.Height
	}
	for row := y0; row < y1; row++ {
		base := (row*raster.Width + x0) * 3
		for column := x0; column < x1; column++ {
			raster.Pix[base] = color[0]
			raster.Pix[base+1] = color[1]
			raster.Pix[base+2] = color[2]
			base += 3
		}
	}
}

type nativeMetafileDIB struct {
	width, height int
	pix           []byte // width*height*3, top-down
}

// nativeMetafileBlt plays META_DIBBITBLT and META_DIBSTRETCHBLT.
func nativeMetafileBlt(raster *nativeMetafileRaster, payload []byte, stretch bool) error {
	header := 16
	if stretch {
		header = 20
	}
	if len(payload) < header {
		return refuseNativeMetafile("metafile bitmap transfer is truncated")
	}
	rop := binary.LittleEndian.Uint32(payload)
	var srcWidth, srcHeight, srcX, srcY, dstHeight, dstWidth, dstY, dstX int32
	read := func(offset int) int32 { return int32(int16(binary.LittleEndian.Uint16(payload[offset:]))) }
	if stretch {
		srcHeight, srcWidth = read(4), read(6)
		srcY, srcX = read(8), read(10)
		dstHeight, dstWidth = read(12), read(14)
		dstY, dstX = read(16), read(18)
	} else {
		srcY, srcX = read(4), read(6)
		dstHeight, dstWidth = read(8), read(10)
		dstY, dstX = read(12), read(14)
		srcWidth, srcHeight = dstWidth, dstHeight
	}
	// A record whose payload stops at the header is the "no source bitmap"
	// form, which only makes sense for the pattern-only raster operations this
	// decoder does not accept here.
	if len(payload) <= header {
		return refuseNativeMetafile("metafile bitmap transfer carries no device-independent bitmap")
	}
	dib, err := nativeMetafileDecodeDIB(payload[header:])
	if err != nil {
		return err
	}
	if srcWidth <= 0 || srcHeight <= 0 || dstWidth <= 0 || dstHeight <= 0 {
		return refuseNativeMetafile("metafile bitmap transfer states an empty or mirrored rectangle")
	}
	switch rop {
	case nativeMetafileROPSrcCopy, nativeMetafileROPSrcAnd, nativeMetafileROPSrcPaint:
	default:
		return refuseNativeMetafile("metafile bitmap raster operation 0x%08X is outside the modeled SRCCOPY, SRCAND and SRCPAINT", rop)
	}
	for row := int32(0); row < dstHeight; row++ {
		destinationRow := int(dstY + row)
		if destinationRow < 0 || destinationRow >= raster.Height {
			continue
		}
		sourceRow := int(srcY + row*srcHeight/dstHeight)
		if sourceRow < 0 || sourceRow >= dib.height {
			continue
		}
		for column := int32(0); column < dstWidth; column++ {
			destinationColumn := int(dstX + column)
			if destinationColumn < 0 || destinationColumn >= raster.Width {
				continue
			}
			sourceColumn := int(srcX + column*srcWidth/dstWidth)
			if sourceColumn < 0 || sourceColumn >= dib.width {
				continue
			}
			source := (sourceRow*dib.width + sourceColumn) * 3
			destination := (destinationRow*raster.Width + destinationColumn) * 3
			for channel := 0; channel < 3; channel++ {
				value := dib.pix[source+channel]
				switch rop {
				case nativeMetafileROPSrcAnd:
					value &= raster.Pix[destination+channel]
				case nativeMetafileROPSrcPaint:
					value |= raster.Pix[destination+channel]
				}
				raster.Pix[destination+channel] = value
			}
		}
	}
	return nil
}

// nativeMetafileDecodeDIB reads an uncompressed BITMAPINFOHEADER DIB into
// top-down RGB. Only the 1bpp and 24bpp BI_RGB forms a control snapshot uses
// are modeled.
func nativeMetafileDecodeDIB(data []byte) (nativeMetafileDIB, error) {
	if len(data) < 40 {
		return nativeMetafileDIB{}, refuseNativeMetafile("device-independent bitmap header is truncated")
	}
	headerSize := binary.LittleEndian.Uint32(data)
	if headerSize != 40 {
		return nativeMetafileDIB{}, refuseNativeMetafile("device-independent bitmap header is %d bytes, not the modeled BITMAPINFOHEADER", headerSize)
	}
	width := int(int32(binary.LittleEndian.Uint32(data[4:])))
	rawHeight := int(int32(binary.LittleEndian.Uint32(data[8:])))
	planes := binary.LittleEndian.Uint16(data[12:])
	bitCount := binary.LittleEndian.Uint16(data[14:])
	compression := binary.LittleEndian.Uint32(data[16:])
	colorsUsed := binary.LittleEndian.Uint32(data[32:])
	if compression != 0 {
		return nativeMetafileDIB{}, refuseNativeMetafile("device-independent bitmap compression %d is outside the modeled BI_RGB", compression)
	}
	if planes != 1 {
		return nativeMetafileDIB{}, refuseNativeMetafile("device-independent bitmap declares %d colour planes", planes)
	}
	topDown := rawHeight < 0
	height := rawHeight
	if topDown {
		height = -height
	}
	if width <= 0 || height <= 0 || width > nativeMetafileMaxExtent || height > nativeMetafileMaxExtent {
		return nativeMetafileDIB{}, refuseNativeMetafile("device-independent bitmap is %dx%d, outside the modeled bound", width, height)
	}
	var palette [][3]byte
	offset := 40
	switch bitCount {
	case 24:
	case 1:
		entries := int(colorsUsed)
		if entries == 0 {
			entries = 2
		}
		if entries != 2 {
			return nativeMetafileDIB{}, refuseNativeMetafile("monochrome bitmap declares %d palette entries", entries)
		}
		if len(data) < offset+8 {
			return nativeMetafileDIB{}, refuseNativeMetafile("monochrome bitmap palette is truncated")
		}
		for entry := 0; entry < 2; entry++ {
			at := offset + entry*4
			palette = append(palette, [3]byte{data[at+2], data[at+1], data[at]})
		}
		offset += 8
	default:
		return nativeMetafileDIB{}, refuseNativeMetafile("device-independent bitmap is %d bits per pixel, outside the modeled 1 and 24", bitCount)
	}
	stride := ((width*int(bitCount) + 31) / 32) * 4
	if len(data)-offset < stride*height {
		return nativeMetafileDIB{}, refuseNativeMetafile("device-independent bitmap pixel data is truncated")
	}
	dib := nativeMetafileDIB{width: width, height: height, pix: make([]byte, width*height*3)}
	for row := 0; row < height; row++ {
		sourceRow := row
		if !topDown {
			sourceRow = height - 1 - row
		}
		line := data[offset+sourceRow*stride:]
		destination := row * width * 3
		if bitCount == 24 {
			for column := 0; column < width; column++ {
				dib.pix[destination] = line[column*3+2]
				dib.pix[destination+1] = line[column*3+1]
				dib.pix[destination+2] = line[column*3]
				destination += 3
			}
			continue
		}
		for column := 0; column < width; column++ {
			bit := (line[column/8] >> (7 - uint(column%8))) & 1
			color := palette[bit]
			dib.pix[destination] = color[0]
			dib.pix[destination+1] = color[1]
			dib.pix[destination+2] = color[2]
			destination += 3
		}
	}
	return dib, nil
}

// encodeNativeMetafilePNG writes the raster as a PNG.
//
// The bytes are produced here rather than by image/png on purpose. The
// extractor states this asset's SHA-256 and the server re-derives the same
// bytes before serving them, so the encoding has to be a fixed function of the
// pixels and nothing else — not of a compressor's heuristics, which are free to
// change between toolchain releases. Stored (uncompressed) DEFLATE blocks make
// that guarantee exactly, at the cost of size: a control snapshot is a few tens
// of kilobytes, comfortably inside the preview byte budget.
func encodeNativeMetafilePNG(raster nativeMetafileRaster) []byte {
	raw := make([]byte, 0, raster.Height*(raster.Width*3+1))
	for row := 0; row < raster.Height; row++ {
		raw = append(raw, 0) // filter type 0 (None), so the bytes stay a pure function of the pixels
		raw = append(raw, raster.Pix[row*raster.Width*3:(row+1)*raster.Width*3]...)
	}
	var header [13]byte
	binary.BigEndian.PutUint32(header[0:], uint32(raster.Width))
	binary.BigEndian.PutUint32(header[4:], uint32(raster.Height))
	header[8] = 8 // bit depth
	header[9] = 2 // colour type: truecolour
	out := []byte{0x89, 'P', 'N', 'G', '\r', '\n', 0x1A, '\n'}
	out = appendNativePNGChunk(out, "IHDR", header[:])
	out = appendNativePNGChunk(out, "IDAT", nativeStoredZlib(raw))
	out = appendNativePNGChunk(out, "IEND", nil)
	return out
}

func appendNativePNGChunk(out []byte, kind string, payload []byte) []byte {
	var length [4]byte
	binary.BigEndian.PutUint32(length[:], uint32(len(payload)))
	out = append(out, length[:]...)
	start := len(out)
	out = append(out, kind...)
	out = append(out, payload...)
	var sum [4]byte
	binary.BigEndian.PutUint32(sum[:], crc32.ChecksumIEEE(out[start:]))
	return append(out, sum[:]...)
}

// nativeStoredZlib wraps payload in a zlib stream of stored DEFLATE blocks.
func nativeStoredZlib(payload []byte) []byte {
	out := make([]byte, 0, len(payload)+len(payload)/65535*5+11)
	out = append(out, 0x78, 0x01) // 32KiB window, fastest compression level
	for at := 0; ; {
		block := len(payload) - at
		if block > 65535 {
			block = 65535
		}
		final := byte(0)
		if at+block == len(payload) {
			final = 1
		}
		out = append(out, final, byte(block), byte(block>>8), byte(^uint16(block)), byte(^uint16(block)>>8))
		out = append(out, payload[at:at+block]...)
		at += block
		if at == len(payload) {
			break
		}
	}
	var sum [4]byte
	binary.BigEndian.PutUint32(sum[:], nativeAdler32(payload))
	return append(out, sum[:]...)
}

func nativeAdler32(payload []byte) uint32 {
	var a, b uint32 = 1, 0
	for _, value := range payload {
		a = (a + uint32(value)) % 65521
		b = (b + a) % 65521
	}
	return b<<16 | a
}

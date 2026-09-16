package pptxpatch

// A refused p:graphicFrame must still occupy its authored box.
//
// A refused p:sp keeps its element: extractAutoShape returns a shape with no
// preset, no geometry and compatibility "refused", which the preview draws as
// an empty outlined region rather than painted content (#259). A refused
// p:graphicFrame used to append no element at all, so the identical refusal
// produced a different shape set than the slide declares: every following
// element shifted one z-index down and the frame's box disappeared from the
// element list that downstream layout and origin computation read.
//
// The region below is that same refused-shape projection: the frame's own
// transform and nothing else. It states no preset, no geometry, no fill and no
// paragraphs, so it can never be mistaken for content we rendered. The
// preserved subtree stays on the slide passthrough issued by
// markSlideUnsupported; the region carries no passthrough of its own.
func (extractor *nativeExtractor) refusedGraphicFrameRegion(node *nativeXMLNode, slidePart, slideID, objectID, fingerprint, code, message string, dialect nativeExtractDialect) (NativeElement, bool) {
	if objectID == "" {
		return NativeElement{}, false
	}
	transform, ok := nativeGraphicFrameRegionTransform(node, dialect)
	if !ok {
		// An inherited, absent or unreadable frame box is not a box we may
		// claim. Nothing is emitted rather than a fabricated extent.
		return NativeElement{}, false
	}
	elementID := extractor.identities.elements[nativeIdentityKey(slidePart, objectID)]
	if elementID == "" {
		elementID = stableNativeID("element", extractor.documentID+"\x00"+slidePart, objectID)
	}
	paragraphs := []NativeParagraph{}
	partName := slidePart
	scopeSlideID := slideID
	return NativeElement{
		Kind: NativeElementKindShape, ID: elementID, Provenance: NativeProvenanceParsed,
		Transform: transform, Paragraphs: &paragraphs,
		Passthrough: []NativePassthroughRef{}, Children: nil,
		Source: &NativeSourceAnchor{PartName: slidePart, ObjectID: objectID, FingerprintSHA256: fingerprint},
		Compatibility: NativeCompatibility{
			Status: NativeCompatibilityStatusRefused,
			Diagnostics: []NativeDiagnostic{{
				Severity: NativeDiagnosticSeverityRefusal, Code: code, Message: message,
				Scope: &NativeDiagnosticScope{SlideID: &scopeSlideID, ElementID: &elementID, PartName: &partName},
			}},
		},
	}, true
}

// nativeGraphicFrameRegionTransform reads only the frame's own p:xfrm box. It
// never consults a layout, a master or a placeholder: an extent the source does
// not state exactly is reported as unavailable.
func nativeGraphicFrameRegionTransform(node *nativeXMLNode, dialect nativeExtractDialect) (NativeTransform, bool) {
	frame := nativeChild(node, dialect.presentation, "xfrm")
	if frame == nil {
		return NativeTransform{}, false
	}
	off := nativeChild(frame, dialect.drawing, "off")
	ext := nativeChild(frame, dialect.drawing, "ext")
	x, xErr := requiredNativeInt64(off, "", "x")
	y, yErr := requiredNativeInt64(off, "", "y")
	cx, cxErr := requiredNativePositiveInt64(ext, "", "cx")
	cy, cyErr := requiredNativePositiveInt64(ext, "", "cy")
	if xErr != nil || yErr != nil || cxErr != nil || cyErr != nil {
		return NativeTransform{}, false
	}
	transform := NativeTransform{X: int64Pointer(x), Y: int64Pointer(y), Cx: int64Pointer(cx), Cy: int64Pointer(cy)}
	// Orientation travels only when the frame states it exactly. An unmodeled
	// rot/flip leaves the region axis-aligned instead of dropping the region.
	if orientation, err := parseNativeSourceAffine(frame); err == nil {
		if orientation.Rotation != 0 {
			transform.RotationAngle = int64Pointer(orientation.Rotation)
		}
		if orientation.FlipH {
			flipH := true
			transform.FlipH = &flipH
		}
		if orientation.FlipV {
			flipV := true
			transform.FlipV = &flipV
		}
	}
	return transform, true
}

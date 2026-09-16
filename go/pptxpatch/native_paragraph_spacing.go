package pptxpatch

import "strconv"

// nativeParagraphSpacingCode discloses that the read-only approximate preview
// carried authored a:pPr paragraph spacing into native v1 layout. The contract
// fields it authorizes may never travel without it.
const nativeParagraphSpacingCode = "pptx.paragraph-spacing-approximate"

// nativeEMUPerHundredthPoint is exact: one point is 12700 EMU, so one
// hundredth of a point is 127 EMU.
const nativeEMUPerHundredthPoint = int64(127)

// nativeMaxParagraphSpacingEMU bounds an emitted gap and an absolute line
// pitch at the same maximum the contract uses for every other EMU offset.
const nativeMaxParagraphSpacingEMU = int64(51206400)

// nativeParagraphSpacingSource is one paragraph's authored a:lnSpc/a:spcBef/
// a:spcAft after the inherited-text preview resolved the style cascade. The
// percentage forms stay unresolved here because ECMA-376 21.1.2.2.7/.9 define
// them against the paragraph's largest authored text size, which is known only
// once the runs (and any authored a:normAutofit fontScale) are extracted.
type nativeParagraphSpacingSource struct {
	present                bool
	lineSpacingPercent1000 *int64
	lineSpacingHundredthPt *int64
	spaceBeforePercent1000 *int64
	spaceBeforeHundredthPt *int64
	spaceAfterPercent1000  *int64
	spaceAfterHundredthPt  *int64
}

// nativeReadParagraphSpacing reads the a:lnSpc/a:spcBef/a:spcAft children of a
// merged a:pPr. The values were already validated by
// nativeInheritedSpacingValue while the cascade layers were sanitized, so a
// malformed value can no longer appear here; a value that somehow does not
// parse is reported as unmodeled instead of being guessed at.
func nativeReadParagraphSpacing(properties *nativeXMLNode, d nativeExtractDialect) (nativeParagraphSpacingSource, error) {
	source := nativeParagraphSpacingSource{}
	if properties == nil {
		return source, nil
	}
	for _, slot := range []struct {
		name    string
		percent **int64
		points  **int64
	}{
		{"lnSpc", &source.lineSpacingPercent1000, &source.lineSpacingHundredthPt},
		{"spcBef", &source.spaceBeforePercent1000, &source.spaceBeforeHundredthPt},
		{"spcAft", &source.spaceAfterPercent1000, &source.spaceAfterHundredthPt},
	} {
		node := nativeChild(properties, d.drawing, slot.name)
		if node == nil {
			continue
		}
		if err := nativeInheritedSpacingValue(node, d); err != nil {
			return nativeParagraphSpacingSource{}, err
		}
		value := node.Children[0]
		raw, _ := exactNativeAttr(value, "", "val")
		parsed, err := strconv.ParseInt(raw, 10, 64)
		if err != nil {
			return nativeParagraphSpacingSource{}, unsupportedNativeTextContent("paragraph spacing value is not canonical")
		}
		source.present = true
		if value.Name.Local == "spcPts" {
			*slot.points = int64Pointer(parsed)
			continue
		}
		*slot.percent = int64Pointer(parsed)
	}
	return source, nil
}

// nativeWithoutParagraphSpacing returns an owned copy of a merged a:pPr with
// the spacing children removed, so the projected paint XML stays exactly the
// node shape the exact paragraph extractor already accepts.
func nativeWithoutParagraphSpacing(node *nativeXMLNode, d nativeExtractDialect) *nativeXMLNode {
	if node == nil {
		return nil
	}
	copied := *node
	copied.Children = make([]*nativeXMLNode, 0, len(node.Children))
	for _, child := range node.Children {
		if nativeIsParagraphSpacingNode(child, d) {
			continue
		}
		copied.Children = append(copied.Children, child)
	}
	return &copied
}

func nativeIsParagraphSpacingNode(node *nativeXMLNode, d nativeExtractDialect) bool {
	return node != nil && node.Name.Space == d.drawing &&
		(node.Name.Local == "lnSpc" || node.Name.Local == "spcBef" || node.Name.Local == "spcAft")
}

// nativeParagraphSpacingBasisHundredthPt is the paragraph's largest authored
// run size, the basis ECMA-376 21.1.2.2.7/.9 give the percentage spacing forms.
// Zero means no run declared an explicit size.
func nativeParagraphSpacingBasisHundredthPt(paragraph NativeParagraph) int64 {
	largest := int64(0)
	for _, run := range paragraph.Runs {
		if run.FontSizeHundredthPt != nil && *run.FontSizeHundredthPt > largest {
			largest = *run.FontSizeHundredthPt
		}
	}
	return largest
}

// nativeParagraphGapEMU resolves one a:spcBef/a:spcAft slot to EMU. A
// percentage resolves against the paragraph's largest authored text size; an
// absolute value converts exactly. ok is false when the slot is authored but
// cannot be modeled, so the caller keeps disclosing it as an omission.
func nativeParagraphGapEMU(percent1000, hundredthPt *int64, basisHundredthPt int64) (gap int64, ok bool) {
	if hundredthPt != nil {
		emu := *hundredthPt * nativeEMUPerHundredthPoint
		if emu < 0 || emu > nativeMaxParagraphSpacingEMU {
			return 0, false
		}
		return emu, true
	}
	if percent1000 == nil {
		return 0, true
	}
	if *percent1000 == 0 {
		return 0, true
	}
	if basisHundredthPt <= 0 {
		return 0, false
	}
	// Half-up rounding of basisEMU * percent / 100000, in exact integers.
	numerator := basisHundredthPt*nativeEMUPerHundredthPoint**percent1000*2 + 100000
	emu := numerator / 200000
	if emu < 0 || emu > nativeMaxParagraphSpacingEMU {
		return 0, false
	}
	return emu, true
}

// nativeApplyParagraphSpacing resolves every paragraph's authored spacing onto
// the extracted contract paragraphs. Identity values (100% line spacing, a zero
// gap) are left absent because they are exactly what absence already means.
// Anything that cannot be modeled stays an omission. It reports whether any
// paragraph now carries spacing, which is what authorizes the disclosure.
func nativeApplyParagraphSpacing(paragraphs []NativeParagraph, sources []nativeParagraphSpacingSource, omit *nativeInheritedTextOmissions) bool {
	if len(sources) == 0 || len(paragraphs) != len(sources) {
		// Without a one-to-one projection the spacing cannot be attributed to a
		// paragraph, so every authored value stays disclosed as an omission.
		for _, source := range sources {
			nativeDiscloseParagraphSpacing(source, omit)
		}
		return false
	}
	applied := false
	for index := range paragraphs {
		source := sources[index]
		if !source.present {
			continue
		}
		basis := nativeParagraphSpacingBasisHundredthPt(paragraphs[index])
		if source.lineSpacingHundredthPt != nil {
			emu := *source.lineSpacingHundredthPt * nativeEMUPerHundredthPoint
			if emu >= 1 && emu <= nativeMaxParagraphSpacingEMU {
				paragraphs[index].LineSpacingEmu = int64Pointer(emu)
				applied = true
			} else {
				omit.add("a:lnSpc")
			}
		} else if source.lineSpacingPercent1000 != nil && *source.lineSpacingPercent1000 != 100000 {
			if *source.lineSpacingPercent1000 >= 1 && *source.lineSpacingPercent1000 <= 13200000 {
				paragraphs[index].LineSpacingPercent1000 = int64Pointer(*source.lineSpacingPercent1000)
				applied = true
			} else {
				omit.add("a:lnSpc")
			}
		}
		for _, slot := range []struct {
			name    string
			percent *int64
			points  *int64
			target  **int64
		}{
			{"a:spcBef", source.spaceBeforePercent1000, source.spaceBeforeHundredthPt, &paragraphs[index].SpaceBeforeEmu},
			{"a:spcAft", source.spaceAfterPercent1000, source.spaceAfterHundredthPt, &paragraphs[index].SpaceAfterEmu},
		} {
			if slot.percent == nil && slot.points == nil {
				continue
			}
			gap, ok := nativeParagraphGapEMU(slot.percent, slot.points, basis)
			if !ok {
				omit.add(slot.name)
				continue
			}
			if gap > 0 {
				*slot.target = int64Pointer(gap)
				applied = true
			}
		}
	}
	return applied
}

// nativeDiscloseParagraphSpacing records every authored slot of one source as
// an omission, for the paths where the spacing could not be carried at all.
func nativeDiscloseParagraphSpacing(source nativeParagraphSpacingSource, omit *nativeInheritedTextOmissions) {
	if source.lineSpacingPercent1000 != nil || source.lineSpacingHundredthPt != nil {
		omit.add("a:lnSpc")
	}
	if source.spaceBeforePercent1000 != nil || source.spaceBeforeHundredthPt != nil {
		omit.add("a:spcBef")
	}
	if source.spaceAfterPercent1000 != nil || source.spaceAfterHundredthPt != nil {
		omit.add("a:spcAft")
	}
}

// nativeMarkParagraphSpacing discloses the declared spacing rule. The element
// becomes preserve-only: PowerPoint resolves this cascade itself on edit.
func nativeMarkParagraphSpacing(element *NativeElement) {
	if element == nil {
		return
	}
	element.Compatibility.Status = worseNativeStatus(element.Compatibility.Status, NativeCompatibilityStatusPreserveOnly)
	element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
		Severity: NativeDiagnosticSeverityWarning,
		Code:     nativeParagraphSpacingCode,
		Message: "Read-only approximate preview resolves authored a:lnSpc, a:spcBef and a:spcAft through the declared style cascade (direct paragraph properties over the body list style over the placeholder/master text style over presentation defaults). " +
			"a:spcPts converts exactly to EMU; a:spcBef/a:spcAft a:spcPct resolves against the paragraph's largest authored run size per ECMA-376 21.1.2.2.7/21.1.2.2.9; a:lnSpc a:spcPct scales the measured natural line height, which is a declared approximation of the renderer's line box, not PowerPoint's line-spacing model. " +
			"Space before is suppressed on the first paragraph of the text body and space after on the last, so authored spacing only ever appears between paragraphs. An authored a:normAutofit lnSpcReduction still reduces the resulting line pitch and never the paragraph gaps. Line pitch, wrapping and overflow may differ from PowerPoint.",
	})
}

// nativeParagraphsCarrySpacing reports whether any paragraph carries an
// authored spacing projection, which is what the disclosure has to cover.
func nativeParagraphsCarrySpacing(paragraphs []NativeParagraph) bool {
	for _, paragraph := range paragraphs {
		if paragraph.LineSpacingPercent1000 != nil || paragraph.LineSpacingEmu != nil || paragraph.SpaceBeforeEmu != nil || paragraph.SpaceAfterEmu != nil {
			return true
		}
	}
	return false
}

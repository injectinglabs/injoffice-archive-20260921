package docxpatch

import (
	"strings"
	"testing"
)

func floatingPositionParts(tblpPr string) map[string]string {
	parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:tbl><w:tblPr>` + tblpPr +
		`</w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4000"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:p/><w:sectPr/></w:body></w:document>`
	return parts
}

func extractedFloatingPosition(t *testing.T, tblpPr string) (*NativeTableFloatingPositionV1, []NativeUnsupportedCapabilityV1) {
	t.Helper()
	document, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(floatingPositionParts(tblpPr))))
	if err != nil {
		t.Fatal(err)
	}
	for _, block := range document.Body.Blocks {
		if block.Table != nil {
			return block.Table.FloatingPosition, document.Unsupported
		}
	}
	t.Fatal("document has no table block")
	return nil, nil
}

func TestNativeTableFloatingPositionProjectsQualifiedFrames(t *testing.T) {
	// The three hard-v2 packages that carry a w:tblpPr Word itself wrote.
	for _, testCase := range []struct {
		name   string
		markup string
		want   NativeTableFloatingPositionV1
	}{
		{
			"negative offset against the margin",
			`<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="-87"/>`,
			NativeTableFloatingPositionV1{HorizontalAnchor: "margin", VerticalAnchor: "text", YTwips: nativeInt64(-87), LeftFromTextTwips: 180, RightFromTextTwips: 180},
		},
		{
			"centred horizontal alignment",
			`<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="center" w:tblpY="36"/>`,
			NativeTableFloatingPositionV1{HorizontalAnchor: "margin", VerticalAnchor: "text", XAlignment: nativeString("center"), YTwips: nativeInt64(36), LeftFromTextTwips: 180, RightFromTextTwips: 180},
		},
		{
			"absent anchors default to the surrounding text",
			`<w:tblpPr w:tblpX="200" w:tblpY="69" w:topFromText="12" w:bottomFromText="13"/>`,
			NativeTableFloatingPositionV1{HorizontalAnchor: "text", VerticalAnchor: "text", XTwips: nativeInt64(200), YTwips: nativeInt64(69), TopFromTextTwips: 12, BottomFromTextTwips: 13},
		},
	} {
		t.Run(testCase.name, func(t *testing.T) {
			position, unsupported := extractedFloatingPosition(t, testCase.markup)
			if position == nil {
				t.Fatal("qualified w:tblpPr was not projected")
			}
			if position.HorizontalAnchor != testCase.want.HorizontalAnchor || position.VerticalAnchor != testCase.want.VerticalAnchor ||
				!nativeSameOptionalInt64(position.XTwips, testCase.want.XTwips) || !nativeSameOptionalInt64(position.YTwips, testCase.want.YTwips) ||
				!nativeSameOptionalString(position.XAlignment, testCase.want.XAlignment) || !nativeSameOptionalString(position.YAlignment, testCase.want.YAlignment) ||
				position.LeftFromTextTwips != testCase.want.LeftFromTextTwips || position.RightFromTextTwips != testCase.want.RightFromTextTwips ||
				position.TopFromTextTwips != testCase.want.TopFromTextTwips || position.BottomFromTextTwips != testCase.want.BottomFromTextTwips {
				t.Fatalf("projected %#v, want %#v", *position, testCase.want)
			}
			// Page paint still lays a floating table out inline, so the frame
			// is projected and the verbatim disclosure is still published.
			found := false
			for _, entry := range unsupported {
				if entry.Code == "UNMODELED_TABLE_PROPERTY" && strings.HasSuffix(entry.Anchor.Path, "/w:tblpPr[1]") {
					found = true
				}
			}
			if !found {
				t.Fatal("projection dropped the verbatim disclosure before placement applies the frame")
			}
		})
	}
}

func TestNativeTableFloatingPositionRefusesUnqualifiedFrames(t *testing.T) {
	// `tblppr-shape.docx` in hard-v2 carries the empty tblpXSpec/tblpYSpec of
	// the first case: not a value of ST_XAlign/ST_YAlign, so nothing is stated.
	for _, testCase := range []struct{ name, markup string }{
		{"empty alignment", `<w:tblpPr w:bottomFromText="0" w:horzAnchor="text" w:leftFromText="180" w:rightFromText="180" w:tblpX="0" w:tblpXSpec="" w:tblpY="1" w:tblpYSpec="" w:topFromText="0" w:vertAnchor="text"/>`},
		{"unknown anchor", `<w:tblpPr w:vertAnchor="paragraph" w:tblpY="10"/>`},
		{"offset and alignment on one axis", `<w:tblpPr w:tblpY="10" w:tblpYSpec="top"/>`},
		{"negative distance from text", `<w:tblpPr w:leftFromText="-1" w:tblpY="10"/>`},
		{"non-numeric offset", `<w:tblpPr w:tblpY="up"/>`},
		{"unknown attribute", `<w:tblpPr w:tblpY="10" w:tblpZ="1"/>`},
		{"child content", `<w:tblpPr w:tblpY="10"><w:extra/></w:tblpPr>`},
	} {
		t.Run(testCase.name, func(t *testing.T) {
			position, unsupported := extractedFloatingPosition(t, testCase.markup)
			if position != nil {
				t.Fatalf("unqualified w:tblpPr was projected as %#v", *position)
			}
			for _, entry := range unsupported {
				if entry.Code == "UNMODELED_TABLE_PROPERTY" && strings.HasSuffix(entry.Anchor.Path, "/w:tblpPr[1]") {
					return
				}
			}
			t.Fatal("unqualified w:tblpPr published no verbatim disclosure")
		})
	}
}

func nativeSameOptionalInt64(got, want *int64) bool {
	return (got == nil) == (want == nil) && (got == nil || *got == *want)
}

func nativeSameOptionalString(got, want *string) bool {
	return (got == nil) == (want == nil) && (got == nil || *got == *want)
}

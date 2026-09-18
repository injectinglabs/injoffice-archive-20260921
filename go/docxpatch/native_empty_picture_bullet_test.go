package docxpatch

import "testing"

const nativeEmptyPictureBulletNSDecl = ` xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="` + relNSTransitional + `" xmlns:a="` + drawingMLTransitional + `"`

// ECMA-376 17.9.21: a w:numPicBullet is the picture a level draws as its
// marker, and the picture itself is a relationship named inside the VML shape.
// A subtree that names no relationship anywhere carries no picture, so the
// level defines a marker with no image in it and no marker ink is lost by
// painting the paragraph without one -- Word's own PDF export of
// `lvlPicBulletId.docx` paints that document's eight bulleted markers as a
// single ArialMT glyph whose ToUnicode maps it to U+0020. A picture bullet that
// does name a relationship is a graphic this tier cannot paint and must keep
// refusing rather than silently dropping it.
func TestNativeEmptyPictureBulletIsUnpaintedAndDisclosed(t *testing.T) {
	for _, tc := range []struct {
		name    string
		bullet  string
		levelID string
		empty   bool
	}{
		{"empty stock picture frame", `<w:pict><v:shape id="_x0000_i1026" type="#_x0000_t75" style="width:3in;height:3in" o:bullet="t"/></w:pict>`, "0", true},
		{"no pict at all", ``, "0", true},
		{"VML image data", `<w:pict><v:shape o:bullet="t"><v:imagedata r:id="rId1" o:title="bullet"/></v:shape></w:pict>`, "0", false},
		{"relationship id anywhere in the subtree", `<w:pict><v:shape o:bullet="t"><v:fill r:id="rId4"/></v:shape></w:pict>`, "0", false},
		{"OLE relationship id", `<w:pict><v:shape o:bullet="t"><o:OLEObject o:relid="rId7"/></v:shape></w:pict>`, "0", false},
		{"DrawingML blip", `<w:drawing><a:blip/></w:drawing>`, "0", false},
		{"unnamed image relationship", `<w:pict><v:shape o:bullet="t"><v:imagedata/></v:shape></w:pict>`, "0", false},
		{"level names an undeclared picture bullet", `<w:pict><v:shape o:bullet="t"/></w:pict>`, "4", false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `"` + nativeEmptyPictureBulletNSDecl + `>` +
				`<w:numPicBullet w:numPicBulletId="0">` + tc.bullet + `</w:numPicBullet>` +
				`<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0A7;"/><w:lvlPicBulletId w:val="` + tc.levelID + `"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>` +
				`<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
			resolved := resolveNumberingFixture(t, numbering, numberedParagraph("2", 0, "bulleted"))
			if resolved.Paragraphs[0].Numbering != nil {
				t.Fatalf("a picture-bullet level resolved a text marker: %#v", resolved.Paragraphs[0].Numbering)
			}
			if got := hasResolutionDiagnostic(resolved, "EMPTY_PICTURE_BULLET_UNPAINTED"); got != tc.empty {
				t.Fatalf("EMPTY_PICTURE_BULLET_UNPAINTED=%v, want %v: %#v", got, tc.empty, resolved.Diagnostics)
			}
			if got := hasResolutionDiagnostic(resolved, "PICTURE_BULLET_PRESERVED"); got == tc.empty {
				t.Fatalf("PICTURE_BULLET_PRESERVED=%v, want %v: %#v", got, !tc.empty, resolved.Diagnostics)
			}
		})
	}
}

// Two w:numPicBullet elements sharing one id state a resource this reading
// cannot identify, so the level keeps refusing rather than picking one.
func TestNativeDuplicatePictureBulletIDKeepsRefusing(t *testing.T) {
	numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `"` + nativeEmptyPictureBulletNSDecl + `>` +
		`<w:numPicBullet w:numPicBulletId="0"><w:pict><v:shape o:bullet="t"/></w:pict></w:numPicBullet>` +
		`<w:numPicBullet w:numPicBulletId="0"><w:pict><v:shape o:bullet="t"><v:imagedata r:id="rId1"/></v:shape></w:pict></w:numPicBullet>` +
		`<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0A7;"/><w:lvlPicBulletId w:val="0"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>` +
		`<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
	resolved := resolveNumberingFixture(t, numbering, numberedParagraph("2", 0, "bulleted"))
	if hasResolutionDiagnostic(resolved, "EMPTY_PICTURE_BULLET_UNPAINTED") || !hasResolutionDiagnostic(resolved, "PICTURE_BULLET_PRESERVED") {
		t.Fatalf("a duplicated picture-bullet id was resolved: %#v", resolved.Diagnostics)
	}
}

// The helper reads the subtree, not one element, and a relationship named
// anywhere in it disqualifies the whole resource.
func TestNativeAbsentPictureBulletHelper(t *testing.T) {
	for _, tc := range []struct {
		markup string
		absent bool
	}{
		{`<w:numPicBullet/>`, true},
		{`<w:numPicBullet><w:pict><v:shape o:bullet="t"/></w:pict></w:numPicBullet>`, true},
		{`<w:numPicBullet><w:pict><v:shape><v:imagedata/></v:shape></w:pict></w:numPicBullet>`, false},
		{`<w:numPicBullet><w:pict><v:shape r:id="rId1"/></w:pict></w:numPicBullet>`, false},
		{`<w:numPicBullet><w:pict><v:shape><o:OLEObject o:relid="rId9"/></v:shape></w:pict></w:numPicBullet>`, false},
		{`<w:numPicBullet><w:pict><v:shape id="rId3"/></w:pict></w:numPicBullet>`, false},
		{`<w:numPicBullet><w:pict><v:shape id="_x0000_i1026"/></w:pict></w:numPicBullet>`, true},
	} {
		root, err := parseNativeXML("numbering.xml", []byte(`<w:root xmlns:w="`+wordMLTransitional+`"`+nativeEmptyPictureBulletNSDecl+`>`+tc.markup+`</w:root>`))
		if err != nil {
			t.Fatal(err)
		}
		if got := nativeAbsentPictureBullet(root.Children[0]); got != tc.absent {
			t.Fatalf("%s: absent=%v, want %v", tc.markup, got, tc.absent)
		}
	}
}

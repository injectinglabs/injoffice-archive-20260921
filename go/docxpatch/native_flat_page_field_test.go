package docxpatch

import (
	"strings"
	"testing"
)

func TestNativeFlatPageField(t *testing.T) {
	field := `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>999 stale</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`
	for _, tc := range []struct {
		name, from, to string
		pass           bool
	}{
		{name: "PAGE", pass: true}, {"NUMPAGES", " PAGE ", " NUMPAGES ", true},
		{"mergeformat", " PAGE ", " PAGE \\* MERGEFORMAT ", true},
		{"arabic", " PAGE ", " PAGE \\* Arabic ", true},
		{"both-switches", " PAGE ", " PAGE \\*Arabic \\*MERGEFORMAT ", true},
		{"unsupported-switch", " PAGE ", " PAGE \\* CHARFORMAT ", false},
		{"unknown", " PAGE ", " DATE ", false},
		{"locked", `w:fldCharType="begin"`, `w:fldCharType="begin" w:fldLock="true"`, false},
		{"dirty", `w:fldCharType="begin"`, `w:fldCharType="begin" w:dirty="true"`, false},
		{"nested", `w:fldCharType="separate"`, `w:fldCharType="begin"`, false},
		{"missing-end", `w:fldCharType="end"`, `w:fldCharType="separate"`, false},
		{"mixed-result", `<w:t>999 stale</w:t>`, `<w:t>999 stale</w:t><w:tab/>`, false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			markup := field
			if tc.from != "" {
				markup = strings.Replace(markup, tc.from, tc.to, 1)
			}
			parts := transitionalNativeParts()
			parts["Custom/Stories/HeaderA.XML"] = `<w:hdr xmlns:w="` + testW + `"><w:p>` + markup + `</w:p></w:hdr>`
			doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			paragraph := doc.Headers[0].Blocks[0].Paragraph
			fields := 0
			for _, run := range paragraph.Runs {
				if run.PageField != "" {
					fields++
					if run.Text == nil || *run.Text != "" {
						t.Fatal("stale cache escaped")
					}
				}
			}
			if tc.pass && fields != 1 || !tc.pass && fields != 0 {
				t.Fatalf("field count %d", fields)
			}
			if paragraph.EditPolicy.Mode != "read-only" {
				t.Fatal("field edit authority widened")
			}
			if paragraph.Anchor.XMLSHA256 == "" || doc.Headers[0].Anchor.XMLSHA256 == "" {
				t.Fatal("raw instruction source digest missing")
			}
		})
	}
}

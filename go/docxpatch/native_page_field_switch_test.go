package docxpatch

import "testing"

func TestNativePageFieldSwitches(t *testing.T) {
	for _, tc := range []struct {
		instruction string
		pass        bool
	}{
		{`PAGE \* MERGEFORMAT`, true}, {`NUMPAGES \* Arabic`, true},
		{`PAGE \* Arabic \* MERGEFORMAT`, true}, {`PAGE \*MERGEFORMAT \*Arabic`, true},
		{`PAGE \* ArabicDash`, false}, {`PAGE \* CHARFORMAT`, false},
		{`PAGE \* Arabic \* Arabic`, false}, {`PAGE \* MERGEFORMAT \* MERGEFORMAT`, false},
		{`PAGE \*`, false}, {`PAGE \# 000`, false}, {`PAGE \* Arabic junk`, false},
		{`PAGE \* "Arabic"`, false}, {`DATE \* Arabic`, false},
	} {
		t.Run(tc.instruction, func(t *testing.T) {
			parts := transitionalNativeParts()
			parts["Custom/Stories/HeaderA.XML"] = `<w:hdr xmlns:w="` + testW + `"><w:p><w:fldSimple w:instr="` + tc.instruction + `"><w:r><w:rPr><w:b/><w:color w:val="123456"/></w:rPr><w:t>999</w:t></w:r></w:fldSimple></w:p></w:hdr>`
			// Quoted instruction is tested at parser boundary rather than invalid XML.
			if tc.instruction == `PAGE \* "Arabic"` {
				if _, ok := nativePageFieldInstruction(tc.instruction); ok {
					t.Fatal("quoted format admitted")
				}
				return
			}
			doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			p := doc.Headers[0].Blocks[0].Paragraph
			if tc.pass {
				if len(p.Runs) != 1 || p.Runs[0].PageField == "" || p.Runs[0].Text == nil || *p.Runs[0].Text != "" {
					t.Fatal("field cache was not discarded")
				}
				properties := p.Runs[0].Properties
				if properties == nil || properties.Bold == nil || !*properties.Bold || properties.Color == nil || *properties.Color != "123456" {
					t.Fatalf("result formatting lost: %+v", properties)
				}
				if p.Anchor.XMLSHA256 == "" || doc.Headers[0].Anchor.XMLSHA256 == "" {
					t.Fatal("instruction source binding absent")
				}
			} else if len(p.Runs) != 0 {
				t.Fatal("unsupported switch admitted")
			}
			if p.EditPolicy.Mode != "read-only" {
				t.Fatal("field result became editable")
			}
		})
	}
}

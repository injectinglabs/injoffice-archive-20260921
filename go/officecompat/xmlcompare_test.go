package officecompat_test

import (
	"reflect"
	"testing"

	"github.com/injectinglabs/injoffice/go/officecompat"
)

func TestStructuralCompareXMLSeparatesLexicalAndStructuralEquality(t *testing.T) {
	before := []byte(`<a:root xmlns:a="urn:test" b="2" a="1"><a:item>hi&amp;</a:item><!-- comment --></a:root>`)
	after := []byte(`<q:root a='1' xmlns:q='urn:test' b='2'><q:item>hi&#38;</q:item></q:root>`)
	comparison := officecompat.CompareXML(before, after)
	if comparison.LexicalEqual {
		t.Fatal("differently serialized XML reported lexical equality")
	}
	if comparison.StructuralStatus != officecompat.XMLStructuralEqual {
		t.Fatalf("namespace/attribute/entity-equivalent XML = %+v", comparison)
	}
	if comparison.BeforeSHA256 == comparison.AfterSHA256 {
		t.Fatal("lexical digests unexpectedly match")
	}
	if comparison.BeforeStructuralSHA256 != comparison.AfterStructuralSHA256 {
		t.Fatal("structural digests unexpectedly differ")
	}

	changed := officecompat.CompareXML(before, []byte(`<a:root xmlns:a="urn:test" a="1" b="2"><a:item>bye&amp;</a:item></a:root>`))
	if changed.StructuralStatus != officecompat.XMLStructuralDifferent {
		t.Fatalf("changed character data = %+v", changed)
	}
}

func TestStructuralCompareXMLCoalescesLogicalCharacterData(t *testing.T) {
	before := []byte(`<root>left&amp;<!-- ignored --><![CDATA[middle]]>&#47;right</root>`)
	after := []byte(`<root>left&amp;middle/right</root>`)
	comparison := officecompat.CompareXML(before, after)
	if comparison.LexicalEqual || comparison.StructuralStatus != officecompat.XMLStructuralEqual {
		t.Fatalf("comment/CDATA/entity-equivalent XML = %+v", comparison)
	}

	changed := officecompat.CompareXML(before, []byte(`<root>left&amp;middle-right</root>`))
	if changed.StructuralStatus != officecompat.XMLStructuralDifferent {
		t.Fatalf("changed coalesced character data = %+v", changed)
	}

	acrossElementBoundary := officecompat.CompareXML([]byte(`<root>left<child/>right</root>`), []byte(`<root>leftright<child/></root>`))
	if acrossElementBoundary.StructuralStatus != officecompat.XMLStructuralDifferent {
		t.Fatalf("character data crossed an element boundary = %+v", acrossElementBoundary)
	}
}

func TestStructuralCompareXMLValidatesDeclaration(t *testing.T) {
	without := []byte(`<root/>`)
	valid := [][]byte{
		[]byte(`<?xml version='1.0'   ?><root></root>`),
		[]byte(`<?xml version='1.0' encoding='utf-8'   ?><root></root>`),
		[]byte(`<?xml version='1.0' encoding='utf-8' standalone='yes'   ?><root></root>`),
	}
	for _, declaration := range valid {
		comparison := officecompat.CompareXML(declaration, without)
		if comparison.LexicalEqual || comparison.StructuralStatus != officecompat.XMLStructuralEqual {
			t.Fatalf("valid declaration changed structure = %+v", comparison)
		}
	}

	tests := []struct {
		name string
		xml  []byte
		want officecompat.XMLFailureCode
	}{
		{name: "malformed pseudoattribute", xml: []byte(`<?xml version='1.0' garbage?><root/>`), want: officecompat.XMLFailureInvalidXML},
		{name: "duplicate declaration", xml: []byte(`<?xml version='1.0'?><?xml version='1.0'?><root/>`), want: officecompat.XMLFailureProcessingInstruction},
		{name: "unsupported version", xml: []byte(`<?xml version='1.1'?><root/>`), want: officecompat.XMLFailureInvalidXML},
		{name: "unsupported encoding", xml: []byte(`<?xml version='1.0' encoding='UTF-16'?><root/>`), want: officecompat.XMLFailureInvalidXML},
		{name: "invalid standalone", xml: []byte(`<?xml version='1.0' standalone='maybe'?><root/>`), want: officecompat.XMLFailureInvalidXML},
		{name: "declaration after whitespace", xml: []byte(" \n<?xml version='1.0'?><root/>"), want: officecompat.XMLFailureProcessingInstruction},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			comparison := officecompat.CompareXML(test.xml, test.xml)
			if comparison.StructuralStatus != officecompat.XMLStructuralUnavailable || comparison.BeforeStructuralFailure != test.want || comparison.AfterStructuralFailure != test.want {
				t.Fatalf("invalid declaration comparison = %+v", comparison)
			}
		})
	}
}

func TestStructuralCompareXMLAcceptsLeadingUTF8BOM(t *testing.T) {
	without := []byte(`<root a="1"/>`)
	withBOM := append([]byte{0xef, 0xbb, 0xbf}, without...)
	comparison := officecompat.CompareXML(withBOM, without)
	if comparison.LexicalEqual || comparison.StructuralStatus != officecompat.XMLStructuralEqual {
		t.Fatalf("leading UTF-8 BOM comparison = %+v", comparison)
	}

	declared := append([]byte{0xef, 0xbb, 0xbf}, []byte(`<?xml version="1.0" encoding="UTF-8"?><root a="1"/>`)...)
	comparison = officecompat.CompareXML(declared, without)
	if comparison.StructuralStatus != officecompat.XMLStructuralEqual {
		t.Fatalf("BOM plus declaration comparison = %+v", comparison)
	}

	misplaced := []byte(" \xef\xbb\xbf<root/>")
	comparison = officecompat.CompareXML(misplaced, misplaced)
	if comparison.StructuralStatus != officecompat.XMLStructuralUnavailable || comparison.BeforeStructuralFailure != officecompat.XMLFailureInvalidXML {
		t.Fatalf("misplaced BOM comparison = %+v", comparison)
	}
}

func TestStructuralCompareXMLNormalizesLiteralAttributeWhitespaceOnly(t *testing.T) {
	literal := []byte("<root a=\"left\tcenter\r\nright\rdown\nend\"/>")
	spaces := []byte(`<root a="left center right down end"/>`)
	comparison := officecompat.CompareXML(literal, spaces)
	if comparison.LexicalEqual || comparison.StructuralStatus != officecompat.XMLStructuralEqual {
		t.Fatalf("XML 1.0 literal attribute whitespace comparison = %+v", comparison)
	}

	characterReferences := []byte(`<root a="left&#x9;center&#xA;right&#xD;down"/>`)
	literalEquivalent := []byte("<root a=\"left\tcenter\nright\rdown\"/>")
	comparison = officecompat.CompareXML(characterReferences, literalEquivalent)
	if comparison.StructuralStatus != officecompat.XMLStructuralDifferent {
		t.Fatalf("character-reference whitespace was collapsed with literals: %+v", comparison)
	}
	comparison = officecompat.CompareXML(characterReferences, []byte("<root a=\"left&#9;center&#10;right&#13;down\"/>"))
	if comparison.StructuralStatus != officecompat.XMLStructuralEqual {
		t.Fatalf("equivalent attribute character references = %+v", comparison)
	}
}

func TestStructuralCompareXMLFailsClosedOnMalformedOrUnsupportedInput(t *testing.T) {
	tests := []struct {
		name string
		xml  []byte
		want officecompat.XMLFailureCode
	}{
		{name: "malformed", xml: []byte(`<root><child></root>`), want: officecompat.XMLFailureInvalidXML},
		{name: "doctype", xml: []byte(`<!DOCTYPE root [<!ENTITY x "boom">]><root>&x;</root>`), want: officecompat.XMLFailureDirective},
		{name: "processing instruction", xml: []byte(`<?target value?><root/>`), want: officecompat.XMLFailureProcessingInstruction},
		{name: "undeclared namespace", xml: []byte(`<w:root/>`), want: officecompat.XMLFailureInvalidXML},
		{name: "reserved namespace", xml: []byte(`<root xmlns:xml="urn:spoof"/>`), want: officecompat.XMLFailureInvalidXML},
		{name: "multiple name colons", xml: []byte(`<a:b:c xmlns:a="urn:a"/>`), want: officecompat.XMLFailureInvalidXML},
		{name: "non-XML whitespace outside root", xml: []byte("\u00a0<root/>"), want: officecompat.XMLFailureInvalidXML},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			comparison := officecompat.CompareXML(test.xml, test.xml)
			if !comparison.LexicalEqual {
				t.Fatal("identical raw bytes lost lexical equality")
			}
			if comparison.StructuralStatus != officecompat.XMLStructuralUnavailable || comparison.BeforeStructuralFailure != test.want || comparison.AfterStructuralFailure != test.want {
				t.Fatalf("fail-closed comparison = %+v", comparison)
			}
		})
	}
}

func TestStructuralCompareXMLResourceLimits(t *testing.T) {
	tests := []struct {
		name   string
		input  []byte
		adjust func(*officecompat.Limits)
	}{
		{
			name:  "bytes",
			input: []byte(`<root/>`),
			adjust: func(limits *officecompat.Limits) {
				limits.MaxXMLBytes = 6
			},
		},
		{
			name:  "depth",
			input: []byte(`<a><b><c/></b></a>`),
			adjust: func(limits *officecompat.Limits) {
				limits.MaxXMLDepth = 2
			},
		},
		{
			name:  "tokens",
			input: []byte(`<a><b/></a>`),
			adjust: func(limits *officecompat.Limits) {
				limits.MaxXMLTokens = 2
			},
		},
		{
			name:  "attributes",
			input: []byte(`<a x="1" y="2"/>`),
			adjust: func(limits *officecompat.Limits) {
				limits.MaxXMLAttributes = 1
			},
		},
		{
			name:  "namespace attributes",
			input: []byte(`<a xmlns:x="urn:x" xmlns:y="urn:y"/>`),
			adjust: func(limits *officecompat.Limits) {
				limits.MaxXMLAttributes = 1
			},
		},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			limits := officecompat.DefaultLimits()
			test.adjust(&limits)
			comparison, err := officecompat.CompareXMLWithLimits(test.input, test.input, limits)
			if err != nil {
				t.Fatal(err)
			}
			if comparison.StructuralStatus != officecompat.XMLStructuralUnavailable || comparison.BeforeStructuralFailure != officecompat.XMLFailureResourceLimit {
				t.Fatalf("resource-bound comparison = %+v", comparison)
			}
		})
	}
}

func TestStructuralCompareXMLOversizedInputsRetainLexicalAuthority(t *testing.T) {
	limits := officecompat.DefaultLimits()
	limits.MaxXMLBytes = 6
	identical := []byte(`<root/>`)
	comparison, err := officecompat.CompareXMLWithLimits(identical, identical, limits)
	if err != nil {
		t.Fatal(err)
	}
	if !comparison.LexicalEqual || comparison.BeforeSHA256 == "" || comparison.BeforeSHA256 != comparison.AfterSHA256 || comparison.StructuralStatus != officecompat.XMLStructuralUnavailable || comparison.BeforeStructuralFailure != officecompat.XMLFailureResourceLimit || comparison.AfterStructuralFailure != officecompat.XMLFailureResourceLimit {
		t.Fatalf("identical oversized lexical result = %+v", comparison)
	}

	changed := []byte(`<roots/>`)
	comparison, err = officecompat.CompareXMLWithLimits(identical, changed, limits)
	if err != nil {
		t.Fatal(err)
	}
	if comparison.LexicalEqual || comparison.BeforeSHA256 == "" || comparison.AfterSHA256 == "" || comparison.BeforeSHA256 == comparison.AfterSHA256 {
		t.Fatalf("changed oversized lexical result = %+v", comparison)
	}
}

func TestStructuralComparePackagesReportsEqualityAndOpaqueChanges(t *testing.T) {
	beforeEntries := minimalOPCRoots()
	beforeEntries["word/document.xml"] = []byte(`<w:document xmlns:w="urn:word" a="1" b="2"><w:body/></w:document>`)
	beforeEntries["custom/data.bin"] = []byte{0, 1, 2}
	afterEntries := cloneEntries(beforeEntries)
	afterEntries["word/document.xml"] = []byte(`<x:document b='2' xmlns:x='urn:word' a='1'><x:body></x:body></x:document>`)
	afterEntries["custom/data.bin"] = []byte{0, 1, 3}

	report, err := officecompat.ComparePackages(
		buildPackage(t, beforeEntries, false, true),
		buildPackage(t, afterEntries, true, false),
	)
	if err != nil {
		t.Fatal(err)
	}
	if len(report.Changed) != 2 {
		t.Fatalf("changed parts = %+v", report.Changed)
	}
	if report.Changed[0].Before.Name != "custom/data.bin" || report.Changed[0].XML != nil {
		t.Fatalf("opaque part was interpreted as XML: %+v", report.Changed[0])
	}
	if report.Changed[1].Before.Name != "word/document.xml" || report.Changed[1].XML == nil || report.Changed[1].XML.StructuralStatus != officecompat.XMLStructuralEqual {
		t.Fatalf("XML structural result = %+v", report.Changed[1])
	}

	repeated, err := officecompat.ComparePackages(
		buildPackage(t, beforeEntries, false, true),
		buildPackage(t, afterEntries, true, false),
	)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(report, repeated) {
		t.Fatalf("package reports are nondeterministic:\nfirst:  %+v\nsecond: %+v", report, repeated)
	}
}

func TestStructuralComparePackagesFallsBackToLexicalForInvalidXML(t *testing.T) {
	beforeEntries := minimalOPCRoots()
	beforeEntries["word/document.xml"] = []byte(`<root><broken></root>`)
	afterEntries := cloneEntries(beforeEntries)
	afterEntries["word/document.xml"] = []byte(`<root><different></root>`)
	report, err := officecompat.ComparePackages(
		buildPackage(t, beforeEntries, false, true),
		buildPackage(t, afterEntries, false, true),
	)
	if err != nil {
		t.Fatal(err)
	}
	if len(report.Changed) != 1 || report.Changed[0].XML == nil {
		t.Fatalf("missing XML fallback report: %+v", report)
	}
	comparison := report.Changed[0].XML
	if comparison.LexicalEqual || comparison.StructuralStatus != officecompat.XMLStructuralUnavailable || comparison.BeforeStructuralFailure != officecompat.XMLFailureInvalidXML || comparison.AfterStructuralFailure != officecompat.XMLFailureInvalidXML {
		t.Fatalf("invalid XML was not a lexical-only difference: %+v", comparison)
	}
}

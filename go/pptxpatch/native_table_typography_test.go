package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func tableTypographyFixture(t *testing.T, strict bool, change func(map[string]string)) []byte {
	return paintTestFixture(t, strict, false, func(parts map[string]string) {
		d, p := nsDrawingTransitional, nsPresentationTransitional
		if strict {
			d, p = nsDrawingStrict, nsPresentationStrict
		}
		level := `<a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>`
		parts["relocated/deck.xml"] = strings.Replace(parts["relocated/deck.xml"], `</p:presentation>`, `<p:defaultTextStyle xmlns:a="`+d+`">`+level+`</p:defaultTextStyle></p:presentation>`, 1)
		parts["relocated/masters/master.xml"] = strings.Replace(nativeExactMasterWithColorMapXML(p), `</p:sldMaster>`, `<p:txStyles xmlns:a="`+d+`"><p:otherStyle>`+level+`</p:otherStyle></p:txStyles></p:sldMaster>`, 1)
		parts["relocated/themes/theme.xml"] = nativeExactThemeXML(d)
		if change != nil {
			change(parts)
		}
	})
}
func TestSourceNoBorderTableTypography(t *testing.T) {
	for _, strict := range []bool{false, true} {
		source := tableTypographyFixture(t, strict, nil)
		copySource := append([]byte(nil), source...)
		deck, err := ExtractNativePPTX(source, nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		found := false
		for _, e := range deck.Slides[0].Elements {
			if e.Table == nil {
				continue
			}
			found = true
			cell := e.Table.Rows[0][0]
			run := (*cell.Paragraphs)[0].Runs[0]
			if *run.FontFamily != "Calibri" || *run.FontSizeHundredthPt != 1800 || *run.KerningThresholdHundredthPt != 1200 || *cell.Text != "Readable & safe" {
				t.Fatalf("source typography changed %+v", run)
			}
			if cell.TextBody.HorizontalOverflow != "clip" || *cell.TextBody.LeftInsetEMU != 91440 || *cell.TextBody.TopInsetEMU != 45720 || e.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
				t.Fatal("source cell defaults/status lost")
			}
		}
		if !found {
			t.Fatal("qualified source table refused")
		}
		if !bytes.Equal(source, copySource) {
			t.Fatal("source mutated")
		}
	}
}
func TestSourceNoBorderTypographyRefusesUnresolvedSources(t *testing.T) {
	changes := map[string]func(map[string]string){
		"conflicting-size": func(p map[string]string) {
			p["relocated/masters/master.xml"] = strings.Replace(p["relocated/masters/master.xml"], `sz="1800"`, `sz="1900"`, 1)
		},
		"conflicting-font": func(p map[string]string) {
			p["relocated/deck.xml"] = strings.Replace(p["relocated/deck.xml"], `+mn-lt`, `+mj-lt`, 1)
		},
		"conflicting-color": func(p map[string]string) {
			p["relocated/deck.xml"] = strings.Replace(p["relocated/deck.xml"], `val="tx1"`, `val="accent1"`, 1)
		},
		"missing-kern": func(p map[string]string) {
			p["relocated/deck.xml"] = strings.Replace(p["relocated/deck.xml"], ` kern="1200"`, ``, 1)
		},
		"active-default-run": func(p map[string]string) {
			p["relocated/deck.xml"] = strings.Replace(p["relocated/deck.xml"], `<a:defRPr sz`, `<a:defRPr b="1" sz`, 1)
		},
		"stray-default-run-text": func(p map[string]string) {
			p["relocated/deck.xml"] = strings.Replace(p["relocated/deck.xml"], `kern="1200">`, `kern="1200">unqualified`, 1)
		},
		"active-terminal": func(p map[string]string) {
			p["relocated/slides/slide-a.xml"] = strings.Replace(p["relocated/slides/slide-a.xml"], `<a:endParaRPr lang="en-US"/>`, `<a:endParaRPr lang="en-US" sz="2400"/>`, 1)
		},
		"active-baseline": func(p map[string]string) {
			p["relocated/slides/slide-a.xml"] = strings.Replace(p["relocated/slides/slide-a.xml"], `lang="en-US" dirty="0"`, `lang="en-US" baseline="1000" dirty="0"`, 1)
		},
		"cell-default-override": func(p map[string]string) {
			p["relocated/slides/slide-a.xml"] = strings.Replace(p["relocated/slides/slide-a.xml"], `<a:tcPr>`, `<a:tcPr marL="1">`, 1)
		},
		"body-autofit": func(p map[string]string) {
			p["relocated/slides/slide-a.xml"] = strings.Replace(p["relocated/slides/slide-a.xml"], `<a:txBody><a:bodyPr/>`, `<a:txBody><a:bodyPr><a:spAutoFit/></a:bodyPr>`, 1)
		},
		"run-unknown": func(p map[string]string) {
			p["relocated/slides/slide-a.xml"] = strings.Replace(p["relocated/slides/slide-a.xml"], `lang="en-US" dirty="0"`, `lang="en-US" unknown="1" dirty="0"`, 1)
		},
		"hidden-frame": func(p map[string]string) {
			p["relocated/slides/slide-a.xml"] = strings.Replace(p["relocated/slides/slide-a.xml"], `name="Inspection"`, `name="Inspection" hidden="1"`, 1)
		},
	}
	for name, change := range changes {
		t.Run(name, func(t *testing.T) {
			deck, err := ExtractNativePPTX(tableTypographyFixture(t, false, change), nativeTestExtractOptions())
			if err != nil {
				return
			}
			for _, e := range deck.Slides[0].Elements {
				if e.Table != nil {
					t.Fatal("unqualified source table admitted")
				}
			}
		})
	}
}

func TestKerningThresholdCannotBeSilentlyDroppedByMutation(t *testing.T) {
	paragraphs := nativeMutationParagraphs("AV")
	paragraphs[0].Runs[0].KerningThresholdHundredthPt = int64Pointer(1200)
	if err := validateNativeMutationParagraphs(paragraphs, &nativeMutationBudget{}); err == nil {
		t.Fatal("mutation accepted unsupported kerning serialization")
	}
	plain := nativeMutationParagraphs("AV")
	if nativeParagraphsEqual(paragraphs, plain) {
		t.Fatal("paragraph equality ignored authored kerning")
	}
}

func TestSourceTypographyCoalescesEquivalentRunsForKerning(t *testing.T) {
	input := tableTypographyFixture(t, false, func(p map[string]string) {
		p["relocated/slides/slide-a.xml"] = strings.Replace(p["relocated/slides/slide-a.xml"], `<a:t>Readable &amp; safe</a:t>`, `<a:t>A</a:t></a:r><a:r><a:rPr lang="en-US" baseline="0" dirty="1"/><a:t>V</a:t>`, 1)
	})
	deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	found := false
	for _, e := range deck.Slides[0].Elements {
		if e.Table == nil {
			continue
		}
		found = true
		runs := (*e.Table.Rows[0][0].Paragraphs)[0].Runs
		if len(runs) != 1 || *runs[0].Text != "AV" {
			t.Fatal("equivalent source span lost cross-run kerning")
		}
	}
	if !found {
		t.Fatal("split source refused")
	}
}

func TestSourceTableTypographyMutationAndPreservation(t *testing.T) {
	source := tableTypographyFixture(t, false, nil)
	deck, err := ExtractNativePPTX(source, nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	var table NativeElement
	for _, e := range deck.Slides[0].Elements {
		if e.Table != nil {
			table = e
		}
	}
	if table.Table == nil {
		t.Fatal("source table missing")
	}
	paragraphs := nativeMutationParagraphs("changed")
	request := NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "refuse-table", Kind: NativePPTXReplaceText, ElementID: table.ID, ExpectedFingerprintSHA256: table.Source.FingerprintSHA256, Paragraphs: &paragraphs}}}
	if output, err := ApplyNativePPTXMutations(source, request); err == nil || output != nil {
		t.Fatal("table replacement was not atomically refused")
	}
	title := deck.Slides[0].Elements[0]
	request.Operations[0].ElementID = title.ID
	request.Operations[0].ExpectedFingerprintSHA256 = title.Source.FingerprintSHA256
	output, err := ApplyNativePPTXMutations(source, request)
	if err != nil {
		t.Fatal(err)
	}
	after, err := ExtractNativePPTX(output, nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	found := false
	for _, e := range after.Slides[0].Elements {
		if e.ID == table.ID {
			found = true
			if e.Source.FingerprintSHA256 != table.Source.FingerprintSHA256 {
				t.Fatal("untouched table bytes changed")
			}
		}
	}
	if !found {
		t.Fatal("table lost after unrelated edit")
	}
	for _, part := range []string{"relocated/styles/table.xml", "relocated/themes/theme.xml", "relocated/masters/master.xml", "relocated/deck.xml"} {
		if !bytes.Equal(chartZipEntry(t, source, part), chartZipEntry(t, output, part)) {
			t.Fatal("untouched source context changed", part)
		}
	}
	for _, rotation := range []int64{0, 1} {
		changed := cloneNativePPTXDeckForMutationTest(t, deck)
		for i := range changed.Slides[0].Elements {
			if changed.Slides[0].Elements[i].ID == table.ID {
				changed.Slides[0].Elements[i].Transform.QuarterTurns = &rotation
			}
		}
		found := false
		for _, issue := range ValidateNativePPTX(changed) {
			if issue.Code == "native.horizontalClip" {
				found = true
			}
		}
		if !found {
			t.Fatal("rotated clipped table not rejected")
		}
	}
}

func TestGroupedSourceTableClipContractRefused(t *testing.T) {
	deck, err := ExtractNativePPTX(tableTypographyFixture(t, false, nil), nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	for i, e := range deck.Slides[0].Elements {
		if e.Table != nil {
			group := e
			group.Kind = NativeElementKindGroup
			group.ID = "group-clip-test"
			group.Table = nil
			group.Children = []NativeElement{e}
			group.ChildTransform = &e.Transform
			deck.Slides[0].Elements[i] = group
			break
		}
	}
	for _, issue := range ValidateNativePPTX(deck) {
		if issue.Code == "native.horizontalClip" {
			return
		}
	}
	t.Fatal("grouped clipped table did not receive explicit contract refusal")
}

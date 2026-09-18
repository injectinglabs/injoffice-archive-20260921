package pptxpatch

import (
	"fmt"
	"strings"
	"testing"
)

// The synthetic fixture mirrors smartart-org-chart.pptx: six data nodes
// (Manager with two paragraphs, Employee, Employee2, Manager2 and the
// assistant), an organization-chart hierarchy layout using composite,
// hierRoot, hierChild, sp, tx and conn, a quick style with matrix references
// and a color transform, but an EMPTY drawing fallback part.
func nativeDiagramLayoutPointXML(id, kind, text string, paragraphs ...string) string {
	attrs := `modelId="` + id + `"`
	if kind != "" {
		attrs += ` type="` + kind + `"`
	}
	body := `<a:bodyPr/><a:lstStyle/>`
	if len(paragraphs) == 0 {
		paragraphs = []string{text}
	}
	for _, paragraph := range paragraphs {
		if paragraph == "" {
			body += `<a:p><a:endParaRPr lang="en-US"/></a:p>`
			continue
		}
		body += `<a:p><a:r><a:rPr lang="en-US" dirty="0" smtClean="0"/><a:t>` + paragraph + `</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p>`
	}
	prSet := `<dgm:prSet phldrT="[Text]"/>`
	if kind != "" {
		prSet = `<dgm:prSet/>`
	}
	return `<dgm:pt ` + attrs + `>` + prSet + `<dgm:spPr/><dgm:t>` + body + `</dgm:t></dgm:pt>`
}

func nativeDiagramLayoutPresPointXML(id, assoc, name, label string) string {
	return `<dgm:pt modelId="` + id + `" type="pres"><dgm:prSet presAssocID="` + assoc + `" presName="` + name + `" presStyleLbl="` + label + `" presStyleIdx="0" presStyleCnt="2"/><dgm:spPr/></dgm:pt>`
}

// nativeDiagramLayoutDataXML returns the data model; extraCxn is appended to
// the connection list and extraPts to the point list.
func nativeDiagramLayoutDataXML(diagramNS, drawingNS, modelExt, extraPts, extraCxn string) string {
	points := nativeDiagramLayoutPointXML("{DOC}", "doc", "") +
		nativeDiagramLayoutPointXML("{MGR}", "", "", "Manager", "Second para") +
		nativeDiagramLayoutPointXML("{MGR-PT}", "parTrans", "") + nativeDiagramLayoutPointXML("{MGR-ST}", "sibTrans", "") +
		nativeDiagramLayoutPointXML("{EMP}", "", "Employee") +
		nativeDiagramLayoutPointXML("{EMP-PT}", "parTrans", "") + nativeDiagramLayoutPointXML("{EMP-ST}", "sibTrans", "") +
		nativeDiagramLayoutPointXML("{EMP2}", "", "Employee2") +
		nativeDiagramLayoutPointXML("{EMP2-PT}", "parTrans", "") + nativeDiagramLayoutPointXML("{EMP2-ST}", "sibTrans", "") +
		nativeDiagramLayoutPointXML("{MGR2}", "", "Manager2") +
		nativeDiagramLayoutPointXML("{MGR2-PT}", "parTrans", "") + nativeDiagramLayoutPointXML("{MGR2-ST}", "sibTrans", "") +
		nativeDiagramLayoutPointXML("{ASST}", "asst", "Assistant") +
		nativeDiagramLayoutPointXML("{ASST-PT}", "parTrans", "") + nativeDiagramLayoutPointXML("{ASST-ST}", "sibTrans", "") +
		// PowerPoint records the style label of layout nodes without styleLbl
		// on the matching presentation point.
		nativeDiagramLayoutPresPointXML("{P-EMP}", "{EMP}", "rootText", "node2") +
		nativeDiagramLayoutPresPointXML("{P-EMP2}", "{EMP2}", "rootText", "node2") +
		extraPts
	connections := `<dgm:cxn modelId="{C1}" srcId="{DOC}" destId="{MGR}" srcOrd="0" destOrd="0" parTransId="{MGR-PT}" sibTransId="{MGR-ST}"/>` +
		`<dgm:cxn modelId="{C2}" srcId="{MGR}" destId="{EMP}" srcOrd="0" destOrd="0" parTransId="{EMP-PT}" sibTransId="{EMP-ST}"/>` +
		`<dgm:cxn modelId="{C3}" srcId="{MGR}" destId="{EMP2}" srcOrd="1" destOrd="0" parTransId="{EMP2-PT}" sibTransId="{EMP2-ST}"/>` +
		`<dgm:cxn modelId="{C4}" srcId="{MGR}" destId="{ASST}" srcOrd="2" destOrd="0" parTransId="{ASST-PT}" sibTransId="{ASST-ST}"/>` +
		`<dgm:cxn modelId="{C5}" srcId="{DOC}" destId="{MGR2}" srcOrd="1" destOrd="0" parTransId="{MGR2-PT}" sibTransId="{MGR2-ST}"/>` +
		`<dgm:cxn modelId="{C6}" type="presOf" srcId="{EMP}" destId="{P-EMP}" srcOrd="0" destOrd="0" presId="urn:test/orgchart"/>` +
		extraCxn
	return `<dgm:dataModel xmlns:dgm="` + diagramNS + `" xmlns:a="` + drawingNS + `"><dgm:ptLst>` + points + `</dgm:ptLst><dgm:cxnLst>` + connections + `</dgm:cxnLst><dgm:bg/><dgm:whole/>` + modelExt + `</dgm:dataModel>`
}

func nativeDiagramLayoutConnXML(name, endPts, bendPt string) string {
	bend := ""
	if bendPt != "" {
		bend = `<dgm:param type="bendPt" val="` + bendPt + `"/>`
	}
	return `<dgm:layoutNode name="` + name + `" styleLbl="parChTrans1D2"><dgm:alg type="conn"><dgm:param type="connRout" val="bend"/><dgm:param type="dim" val="1D"/><dgm:param type="endSty" val="noArr"/><dgm:param type="begPts" val="bCtr"/><dgm:param type="endPts" val="` + endPts + `"/>` + bend + `</dgm:alg>` +
		`<dgm:shape type="conn" zOrderOff="-99999"><dgm:adjLst/></dgm:shape><dgm:presOf axis="self"/><dgm:constrLst><dgm:constr type="begPad"/><dgm:constr type="endPad"/></dgm:constrLst><dgm:ruleLst/></dgm:layoutNode>`
}

// nativeDiagramLayoutCompositeXML mirrors rootComposite: a text rectangle
// filling the composite plus a hidden connector anchor one fifth as wide.
func nativeDiagramLayoutCompositeXML(suffix, textLabel, connectorLabel string) string {
	textAttr := ""
	if textLabel != "" {
		textAttr = ` styleLbl="` + textLabel + `"`
	}
	return `<dgm:layoutNode name="rootComposite` + suffix + `"><dgm:alg type="composite"/><dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf axis="self" ptType="node" cnt="1"/>` +
		`<dgm:constrLst><dgm:constr type="l" for="ch" forName="rootText` + suffix + `"/><dgm:constr type="t" for="ch" forName="rootText` + suffix + `"/><dgm:constr type="w" for="ch" forName="rootText` + suffix + `" refType="w"/><dgm:constr type="h" for="ch" forName="rootText` + suffix + `" refType="h"/>` +
		`<dgm:constr type="l" for="ch" forName="rootConnector` + suffix + `"/><dgm:constr type="t" for="ch" forName="rootConnector` + suffix + `"/><dgm:constr type="w" for="ch" forName="rootConnector` + suffix + `" refType="w" refFor="ch" refForName="rootText` + suffix + `" fact="0.2"/><dgm:constr type="h" for="ch" forName="rootConnector` + suffix + `" refType="h" refFor="ch" refForName="rootText` + suffix + `"/></dgm:constrLst><dgm:ruleLst/>` +
		`<dgm:layoutNode name="rootText` + suffix + `"` + textAttr + `><dgm:varLst><dgm:chPref val="3"/></dgm:varLst><dgm:alg type="tx"/><dgm:shape type="rect"><dgm:adjLst/></dgm:shape><dgm:presOf axis="self" ptType="node" cnt="1"/>` +
		`<dgm:constrLst><dgm:constr type="primFontSz" val="65"/><dgm:constr type="lMarg" refType="primFontSz" fact="0.05"/><dgm:constr type="rMarg" refType="primFontSz" fact="0.05"/><dgm:constr type="tMarg" refType="primFontSz" fact="0.05"/><dgm:constr type="bMarg" refType="primFontSz" fact="0.05"/></dgm:constrLst>` +
		`<dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst></dgm:layoutNode>` +
		`<dgm:layoutNode name="rootConnector` + suffix + `" styleLbl="` + connectorLabel + `" moveWith="rootText` + suffix + `"><dgm:alg type="sp"/><dgm:shape type="rect" hideGeom="1"><dgm:adjLst/></dgm:shape><dgm:presOf axis="self" ptType="node" cnt="1"/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode>` +
		`</dgm:layoutNode>`
}

// nativeDiagramLayoutHierRootXML mirrors hierRoot2/hierRoot3: leaf-style
// hanging alignment when the subtree is shallow, standard otherwise.
func nativeDiagramLayoutHierRootXML(name, suffix, textLabel, connectorLabel, childA, childB, loopA, loopB string) string {
	return `<dgm:layoutNode name="` + name + `"><dgm:varLst><dgm:hierBranch val="init"/></dgm:varLst>` +
		`<dgm:choose name="ch-` + name + `"><dgm:if name="if-` + name + `" axis="des" func="maxDepth" op="lte" val="1"><dgm:alg type="hierRoot"><dgm:param type="hierAlign" val="tL"/></dgm:alg><dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst><dgm:constr type="alignOff" val="0.25"/></dgm:constrLst></dgm:if>` +
		`<dgm:else name="else-` + name + `"><dgm:alg type="hierRoot"/><dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst><dgm:constr type="alignOff"/><dgm:constr type="bendDist" for="des" ptType="parTrans" refType="sp" fact="0.5"/></dgm:constrLst></dgm:else></dgm:choose><dgm:ruleLst/>` +
		nativeDiagramLayoutCompositeXML(suffix, textLabel, connectorLabel) +
		`<dgm:layoutNode name="` + childA + `"><dgm:choose name="ch-` + childA + `"><dgm:if name="if-` + childA + `" axis="des" func="maxDepth" op="lte" val="1"><dgm:alg type="hierChild"><dgm:param type="chAlign" val="l"/><dgm:param type="linDir" val="fromT"/></dgm:alg></dgm:if><dgm:else name="else-` + childA + `"><dgm:alg type="hierChild"/></dgm:else></dgm:choose><dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/><dgm:forEach name="loop-` + childA + `" ref="` + loopA + `"/></dgm:layoutNode>` +
		`<dgm:layoutNode name="` + childB + `"><dgm:alg type="hierChild"><dgm:param type="chAlign" val="l"/><dgm:param type="linDir" val="fromL"/><dgm:param type="secChAlign" val="t"/><dgm:param type="secLinDir" val="fromT"/></dgm:alg><dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/><dgm:forEach name="loop-` + childB + `" ref="` + loopB + `"/></dgm:layoutNode>` +
		`</dgm:layoutNode>`
}

func nativeDiagramLayoutLayoutXML(diagramNS string) string {
	rootConstraints := ""
	for _, composite := range []string{"rootComposite1", "rootComposite", "rootComposite3"} {
		rootConstraints += `<dgm:constr type="w" for="des" forName="` + composite + `" refType="w" fact="10"/><dgm:constr type="h" for="des" forName="` + composite + `" refType="w" refFor="des" refForName="rootComposite1" fact="0.5"/>`
	}
	rootConstraints += `<dgm:constr type="primFontSz" for="des" ptType="node" op="equ"/><dgm:constr type="sp" for="des" op="equ"/>` +
		`<dgm:constr type="sp" for="des" forName="hierRoot1" refType="w" refFor="des" refForName="rootComposite1" fact="0.21"/>` +
		`<dgm:constr type="sp" for="des" forName="hierRoot2" refType="sp" refFor="des" refForName="hierRoot1"/><dgm:constr type="sp" for="des" forName="hierRoot3" refType="sp" refFor="des" refForName="hierRoot1"/>` +
		`<dgm:constr type="sibSp" refType="w" refFor="des" refForName="rootComposite1" fact="0.21"/><dgm:constr type="secSibSp" refType="w" refFor="des" refForName="rootComposite1" fact="0.21"/>`
	for _, child := range []string{"hierChild2", "hierChild3", "hierChild4", "hierChild5", "hierChild6", "hierChild7"} {
		rootConstraints += `<dgm:constr type="sibSp" for="des" forName="` + child + `" refType="sibSp"/><dgm:constr type="secSibSp" for="des" forName="` + child + `" refType="secSibSp"/>`
	}
	return `<dgm:layoutDef xmlns:dgm="` + diagramNS + `" uniqueId="urn:test/orgchart"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="hierarchy" pri="1000"/></dgm:catLst>` +
		`<dgm:layoutNode name="hierChild1"><dgm:varLst><dgm:orgChart val="1"/><dgm:chPref val="1"/><dgm:dir/><dgm:animOne val="branch"/><dgm:animLvl val="lvl"/><dgm:resizeHandles/></dgm:varLst>` +
		`<dgm:choose name="Name0"><dgm:if name="Name1" func="var" arg="dir" op="equ" val="norm"><dgm:alg type="hierChild"><dgm:param type="linDir" val="fromL"/></dgm:alg></dgm:if><dgm:else name="Name2"><dgm:alg type="hierChild"><dgm:param type="linDir" val="fromR"/></dgm:alg></dgm:else></dgm:choose>` +
		`<dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst>` + rootConstraints + `</dgm:constrLst><dgm:ruleLst/>` +
		`<dgm:forEach name="Name3" axis="ch"><dgm:forEach name="Name4" axis="self" ptType="node">` +
		`<dgm:layoutNode name="hierRoot1"><dgm:varLst><dgm:hierBranch val="init"/></dgm:varLst>` +
		`<dgm:choose name="Name5"><dgm:if name="Name6" func="var" arg="hierBranch" op="equ" val="hang"><dgm:alg type="hierRoot"/><dgm:constrLst><dgm:constr type="alignOff" val="0.65"/></dgm:constrLst></dgm:if><dgm:else name="Name7"><dgm:alg type="hierRoot"/><dgm:constrLst><dgm:constr type="alignOff"/><dgm:constr type="bendDist" for="des" ptType="parTrans" refType="sp" fact="0.5"/></dgm:constrLst></dgm:else></dgm:choose>` +
		`<dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:ruleLst/>` +
		nativeDiagramLayoutCompositeXML("1", "node0", "node1") +
		`<dgm:layoutNode name="hierChild2"><dgm:choose name="Name21"><dgm:if name="Name22" func="var" arg="dir" op="equ" val="norm"><dgm:alg type="hierChild"/></dgm:if><dgm:else name="Name23"><dgm:alg type="hierChild"><dgm:param type="linDir" val="fromR"/></dgm:alg></dgm:else></dgm:choose><dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/>` +
		`<dgm:forEach name="rep2a" axis="ch" ptType="nonAsst"><dgm:forEach name="Name32" axis="precedSib" ptType="parTrans" st="-1" cnt="1">` +
		`<dgm:choose name="Name38"><dgm:if name="Name39" axis="self" func="depth" op="lte" val="2">` + nativeDiagramLayoutConnXML("Name37", "tCtr", "end") + `</dgm:if><dgm:else name="Name40">` + nativeDiagramLayoutConnXML("Name41", "midL midR", "") + `</dgm:else></dgm:choose>` +
		`</dgm:forEach>` +
		nativeDiagramLayoutHierRootXML("hierRoot2", "", "", "node2", "hierChild4", "hierChild5", "rep2a", "rep2b") +
		`</dgm:forEach></dgm:layoutNode>` +
		`<dgm:layoutNode name="hierChild3"><dgm:alg type="hierChild"><dgm:param type="chAlign" val="l"/><dgm:param type="linDir" val="fromL"/><dgm:param type="secChAlign" val="t"/><dgm:param type="secLinDir" val="fromT"/></dgm:alg><dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/>` +
		`<dgm:forEach name="rep2b" axis="ch" ptType="asst"><dgm:forEach name="Name110" axis="precedSib" ptType="parTrans" st="-1" cnt="1">` + nativeDiagramLayoutConnXML("Name111", "midL midR", "") + `</dgm:forEach>` +
		nativeDiagramLayoutHierRootXML("hierRoot3", "3", "asst1", "asst1", "hierChild6", "hierChild7", "rep2a", "rep2b") +
		`</dgm:forEach></dgm:layoutNode>` +
		`</dgm:layoutNode></dgm:forEach></dgm:forEach></dgm:layoutNode></dgm:layoutDef>`
}

func nativeDiagramLayoutQuickStyleXML(diagramNS, drawingNS string) string {
	label := func(name, fillIdx, font string) string {
		return `<dgm:styleLbl name="` + name + `"><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:sp3d/><dgm:txPr/><dgm:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="` + fillIdx + `"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor">` + font + `</a:fontRef></dgm:style></dgm:styleLbl>`
	}
	return `<dgm:styleDef xmlns:dgm="` + diagramNS + `" xmlns:a="` + drawingNS + `" uniqueId="urn:test/style"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="simple" pri="10100"/></dgm:catLst>` +
		label("node0", "1", `<a:schemeClr val="lt1"/>`) + label("node1", "1", `<a:schemeClr val="lt1"/>`) + label("node2", "1", `<a:schemeClr val="lt1"/>`) + label("asst1", "1", `<a:schemeClr val="lt1"/>`) + label("parChTrans1D2", "0", "") +
		`</dgm:styleDef>`
}

func nativeDiagramLayoutColorsXML(diagramNS, drawingNS string) string {
	label := func(name, line, text string) string {
		return `<dgm:styleLbl name="` + name + `"><dgm:fillClrLst meth="repeat"><a:schemeClr val="accent1"/></dgm:fillClrLst><dgm:linClrLst meth="repeat">` + line + `</dgm:linClrLst><dgm:effectClrLst/><dgm:txLinClrLst/><dgm:txFillClrLst>` + text + `</dgm:txFillClrLst><dgm:txEffectClrLst/></dgm:styleLbl>`
	}
	return `<dgm:colorsDef xmlns:dgm="` + diagramNS + `" xmlns:a="` + drawingNS + `" uniqueId="urn:test/colors"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="accent1" pri="11200"/></dgm:catLst>` +
		label("node0", `<a:schemeClr val="lt1"/>`, "") + label("node1", `<a:schemeClr val="lt1"/>`, "") + label("node2", `<a:schemeClr val="lt1"/>`, "") + label("asst1", `<a:schemeClr val="lt1"/>`, "") +
		label("parChTrans1D2", `<a:schemeClr val="accent1"><a:shade val="60000"/></a:schemeClr>`, `<a:schemeClr val="dk1"/>`) +
		`</dgm:colorsDef>`
}

// nativeDiagramLayoutThemeXML carries the Office format scheme subset the
// quick style references: one solid fill entry and three exact solid lines.
func nativeDiagramLayoutThemeXML(drawingNS string) string {
	lines := ""
	for _, width := range []string{"9525", "25400", "38100"} {
		lines += `<a:ln w="` + width + `" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>`
	}
	theme := strings.Replace(nativeExactThemeXML(drawingNS),
		`<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>`,
		`<a:lnStyleLst>`+lines+`</a:lnStyleLst>`, 1)
	if !strings.Contains(theme, `w="25400"`) {
		panic("theme fixture drifted")
	}
	return theme
}

type nativeDiagramLayoutFixtureOptions struct {
	strict   bool
	extraPts string
	extraCxn string
	layout   string
	style    string
	colors   string
	dataXML  string
}

func nativeDiagramLayoutFixture(t *testing.T, options nativeDiagramLayoutFixtureOptions) []byte {
	t.Helper()
	diagramNS, drawingNS := nativeDiagramURITransitional, nsDrawingTransitional
	if options.strict {
		diagramNS, drawingNS = nativeDiagramURIStrict, nsDrawingStrict
	}
	modelExt := `<dgm:extLst><a:ext uri="` + nsDiagramDrawing + `"><dsp:dataModelExt xmlns:dsp="` + nsDiagramDrawing + `" relId="rIdDrawing" minVer="` + diagramNS + `"/></a:ext></dgm:extLst>`
	data := options.dataXML
	if data == "" {
		data = nativeDiagramLayoutDataXML(diagramNS, drawingNS, modelExt, options.extraPts, options.extraCxn)
	}
	layout := options.layout
	if layout == "" {
		layout = nativeDiagramLayoutLayoutXML(diagramNS)
	}
	style := options.style
	if style == "" {
		style = nativeDiagramLayoutQuickStyleXML(diagramNS, drawingNS)
	}
	colors := options.colors
	if colors == "" {
		colors = nativeDiagramLayoutColorsXML(diagramNS, drawingNS)
	}
	empty := `<dsp:drawing xmlns:dgm="` + diagramNS + `" xmlns:dsp="` + nsDiagramDrawing + `" xmlns:a="` + drawingNS + `"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/></dsp:spTree></dsp:drawing>`
	return nativeDiagramFixture(t, nativeDiagramFixtureOptions{
		strict: options.strict, drawingXML: empty, themeXML: nativeDiagramLayoutThemeXML(drawingNS),
		dataXML: data, layoutXML: layout, styleXML: style, colorsXML: colors,
	})
}

func nativeDiagramLayoutApproximateOptions() NativePPTXExtractOptions {
	options := nativeTestExtractOptions()
	options.AllowInheritedTextPreview = true
	return options
}

func nativeDiagramLayoutChildText(child NativeElement) string {
	if child.Paragraphs == nil {
		return ""
	}
	parts := []string{}
	for _, paragraph := range *child.Paragraphs {
		for _, run := range paragraph.Runs {
			if run.Text != nil {
				parts = append(parts, *run.Text)
			}
		}
	}
	return strings.Join(parts, "/")
}

func nativeDiagramLayoutCodes(element NativeElement) string {
	codes := []string{}
	for _, diagnostic := range element.Compatibility.Diagnostics {
		codes = append(codes, diagnostic.Code)
	}
	return strings.Join(codes, ",")
}

func TestExtractNativePPTXDiagramLayoutApproximatesOrgChart(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			t.Parallel()
			deck, err := ExtractNativePPTX(nativeDiagramLayoutFixture(t, nativeDiagramLayoutFixtureOptions{strict: strict}), nativeDiagramLayoutApproximateOptions())
			if err != nil {
				t.Fatalf("extract laid-out diagram: %v", err)
			}
			slide := deck.Slides[0]
			group := nativeFixtureDiagramGroup(t, slide)
			if group.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || len(group.Children) != 8 || group.ChildTransform == nil {
				t.Fatalf("diagram was not laid out as one preserve-only group of 3 connectors and 5 shapes: %d children", len(group.Children))
			}
			if *group.Transform.X != 1_524_000 || *group.Transform.Y != 1_397_000 || *group.Transform.Cx != 6_096_000 || *group.Transform.Cy != 4_064_000 {
				t.Fatalf("frame EMU changed: %#v", group.Transform)
			}
			if nativeDiagramLayoutCodes(group) != nativeDiagramLayoutPreviewCode || !strings.Contains(group.Compatibility.Diagnostics[0].Message, nativeDiagramLayoutPolicy) || len(group.Passthrough) != 1 {
				t.Fatalf("layout policy was not declared on the group: %#v", group.Compatibility.Diagnostics)
			}
			frameW, frameH := *group.Transform.Cx, *group.Transform.Cy
			shapes := map[string]NativeElement{}
			connectors := []NativeElement{}
			expectedLine := applyNativeShade(mustParseNativeSRGB(t, "2F6FED"), 60000).hex()
			for index, child := range group.Children {
				if child.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || child.Source == nil || !strings.HasPrefix(child.Source.ObjectID, "cNvPr-4/dgm/") || !strings.HasPrefix(nativeDiagramLayoutCodes(child), nativeDiagramLayoutPreviewCode) {
					t.Fatalf("child %d is not a labeled read-only layout element: %#v", index, child)
				}
				if *child.Transform.X < 0 || *child.Transform.Y < 0 || *child.Transform.X+*child.Transform.Cx > frameW || *child.Transform.Y+*child.Transform.Cy > frameH {
					t.Fatalf("child %d escapes the frame: %#v", index, child.Transform)
				}
				switch child.Kind {
				case NativeElementKindConnector:
					if len(shapes) != 0 {
						t.Fatalf("connector %d must paint behind the shapes (zOrderOff)", index)
					}
					if child.Geometry == nil || len(child.Geometry.Paths) != 1 || child.Geometry.Paths[0].FillMode != "none" || !child.Geometry.Paths[0].Stroke || child.Stroke == nil || child.Stroke.Color != expectedLine || *child.Stroke.WidthEMU != 25_400 {
						t.Fatalf("connector %d paint or geometry drifted from the color transform and theme line: %#v", index, child)
					}
					connectors = append(connectors, child)
				case NativeElementKindShape:
					if child.Geometry == nil || child.Preset != nil || child.Fill == nil || *child.Fill != "2F6FED" || child.Stroke == nil || child.Stroke.Color != "FFFFFF" || *child.Stroke.WidthEMU != 25_400 {
						t.Fatalf("shape %d paint was not resolved from quick style, colors and theme: %#v", index, child)
					}
					// The uniform fit scale is not a whole number of EMU, so the
					// file's h = 0.5 w constraint is honored down to the EMU
					// grid (1 EMU = 1/914400 inch), not below it.
					if ratio := *child.Transform.Cx - 2**child.Transform.Cy; ratio < -1 || ratio > 1 {
						t.Fatalf("shape %d does not honor the h = 0.5 w constraint: %#v", index, child.Transform)
					}
					if !strings.Contains(nativeDiagramLayoutCodes(child), nativeInheritedTextPreviewCode) || child.TextBody == nil || child.TextBody.VerticalAnchor != NativeTextVerticalAnchorCenter {
						t.Fatalf("shape %d text is not the declared inherited preview centered in the box: %s %#v", index, nativeDiagramLayoutCodes(child), child.TextBody)
					}
					shapes[nativeDiagramLayoutChildText(child)] = child
				default:
					t.Fatalf("unexpected child kind %s", child.Kind)
				}
			}
			if len(connectors) != 3 || len(shapes) != 5 {
				t.Fatalf("expected 3 connectors and 5 shapes: %d %v", len(connectors), shapes)
			}
			manager, manager2, employee, employee2, assistant := shapes["Manager/Second para"], shapes["Manager2"], shapes["Employee"], shapes["Employee2"], shapes["Assistant"]
			for label, shape := range shapes {
				if shape.ID == "" {
					t.Fatalf("shape %q missing", label)
				}
			}
			// Sizes follow the constraint ratios: box 10 x 5 units, spacing 2.1
			// units. Manager2 has no descendants, so contour packing only has
			// to clear the manager box on the root row (10 + 2.1 units past the
			// manager's left edge) instead of the whole first subtree; the
			// chart is 28.15 x 19.2 units, fitted to the 4064000 EMU frame
			// height and centered horizontally.
			unit := 4_064_000.0 / 19.2
			originX := (6_096_000 - 28.15*unit) / 2
			near := func(name string, got int64, want float64) {
				if diff := float64(got) - want; diff < -1.5 || diff > 1.5 {
					t.Fatalf("%s = %d, want %.0f", name, got, want)
				}
			}
			near("box width", *manager.Transform.Cx, 10*unit)
			near("manager x", *manager.Transform.X, originX+6.05*unit)
			near("manager y", *manager.Transform.Y, 0)
			near("manager2 x", *manager2.Transform.X, originX+18.15*unit)
			near("employee x", *employee.Transform.X, originX)
			near("employee2 x", *employee2.Transform.X, originX+12.1*unit)
			near("employee y", *employee.Transform.Y, 14.2*unit)
			near("assistant y", *assistant.Transform.Y, 7.1*unit)
			if *manager2.Transform.Y != *manager.Transform.Y || *employee2.Transform.Y != *employee.Transform.Y {
				t.Fatalf("hierarchy rows are not aligned: %#v %#v", manager.Transform, employee.Transform)
			}
			managerCenter := *manager.Transform.X + *manager.Transform.Cx/2
			if right := *assistant.Transform.X + *assistant.Transform.Cx; right >= managerCenter || managerCenter-right > int64(1.1*unit) {
				t.Fatalf("assistant must hang left of the manager trunk by half the sibling spacing: right=%d center=%d", right, managerCenter)
			}
			childrenCenter := (*employee.Transform.X + *employee2.Transform.X + *employee2.Transform.Cx) / 2
			if childrenCenter < managerCenter-1 || childrenCenter > managerCenter+1 {
				t.Fatalf("children row is not centered under the manager: %d vs %d", childrenCenter, managerCenter)
			}
			// Contour packing: the childless second root only clears the root
			// row, so it starts left of the first root's own envelope, yet no
			// two painted boxes overlap anywhere.
			if *manager2.Transform.X >= *employee2.Transform.X+*employee2.Transform.Cx {
				t.Fatalf("second root was pushed past the first subtree envelope instead of packing against its contour: %d", *manager2.Transform.X)
			}
			for _, left := range shapes {
				for _, right := range shapes {
					if left.ID == right.ID {
						continue
					}
					if *left.Transform.X < *right.Transform.X+*right.Transform.Cx && *right.Transform.X < *left.Transform.X+*left.Transform.Cx &&
						*left.Transform.Y < *right.Transform.Y+*right.Transform.Cy && *right.Transform.Y < *left.Transform.Y+*left.Transform.Cy {
						t.Fatalf("contour packing overlapped two painted boxes: %#v %#v", left.Transform, right.Transform)
					}
				}
			}
			// Connector 1: Manager bottom center -> bend -> Employee top center,
			// with the horizontal bar half a spacing above the children.
			absolute := func(connector NativeElement, index int) (int64, int64) {
				command := connector.Geometry.Paths[0].Commands[index]
				return *connector.Transform.X + *command.X, *connector.Transform.Y + *command.Y
			}
			first := connectors[0]
			if len(first.Geometry.Paths[0].Commands) != 4 {
				t.Fatalf("parent-child connector must be a three-segment bend: %#v", first.Geometry.Paths[0].Commands)
			}
			if x, y := absolute(first, 0); x != managerCenter || y != *manager.Transform.Y+*manager.Transform.Cy {
				t.Fatalf("connector does not start at the manager bottom center: %d,%d", x, y)
			}
			if x, y := absolute(first, 3); x != *employee.Transform.X+*employee.Transform.Cx/2 || y != *employee.Transform.Y {
				t.Fatalf("connector does not end at the employee top center: %d,%d", x, y)
			}
			if _, barY := absolute(first, 1); float64(*employee.Transform.Y-barY) < 1.04*unit || float64(*employee.Transform.Y-barY) > 1.06*unit {
				t.Fatalf("bend bar is not bendDist = 0.5 sp above the child row: %d", *employee.Transform.Y-barY)
			}
			third := connectors[2]
			if len(third.Geometry.Paths[0].Commands) != 3 {
				t.Fatalf("assistant connector must be a two-segment bend: %#v", third.Geometry.Paths[0].Commands)
			}
			// Connector ends and shape edges are rounded to EMU from the same
			// unrounded layout, so they agree on the EMU grid (1 EMU =
			// 1/914400 inch) rather than below it.
			onGrid := func(got, want int64) bool { return got-want >= -1 && got-want <= 1 }
			if x, y := absolute(third, 2); !onGrid(x, *assistant.Transform.X+*assistant.Transform.Cx) || !onGrid(y, *assistant.Transform.Y+*assistant.Transform.Cy/2) {
				t.Fatalf("assistant connector does not end at the assistant middle right: %d,%d", x, y)
			}
			// tx: one fitted size shared by the primFontSz equalization group,
			// white minor-font runs from the quick style font reference.
			var size int64
			for label, shape := range shapes {
				for _, paragraph := range *shape.Paragraphs {
					for _, run := range paragraph.Runs {
						if run.FontSizeHundredthPt == nil || run.Color == nil || *run.Color != "FFFFFF" || run.FontFamily == nil || *run.FontFamily != "Calibri" {
							t.Fatalf("%s run did not resolve the diagram font reference: %#v", label, run)
						}
						if size == 0 {
							size = *run.FontSizeHundredthPt
						}
						if *run.FontSizeHundredthPt != size || size < 500 || size > 6500 {
							t.Fatalf("%s font size %d is not the equalized fitted size %d", label, *run.FontSizeHundredthPt, size)
						}
					}
				}
			}
			if len(*manager.Paragraphs) != 2 || (*manager.Paragraphs)[0].Align == nil || *(*manager.Paragraphs)[0].Align != NativeTextAlignCenter {
				t.Fatalf("manager paragraphs were not projected centered: %#v", manager.Paragraphs)
			}
			for _, diagnostic := range slide.Compatibility.Diagnostics {
				if strings.HasPrefix(diagnostic.Code, "pptx.diagram-") && diagnostic.Code != nativeDiagramLayoutPreviewCode {
					t.Fatalf("laid-out diagram must not also be refused at slide level: %#v", diagnostic)
				}
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("invalid laid-out diagram deck: %#v", issues)
			}
			encoded, err := MarshalNativePPTXJSON(deck)
			if err != nil {
				t.Fatalf("encode: %v", err)
			}
			if _, err := DecodeNativePPTXJSON(encoded); err != nil {
				t.Fatalf("decode: %v", err)
			}
		})
	}
}

func TestExtractNativePPTXDiagramLayoutExactTierStillRefusesEmptyFallback(t *testing.T) {
	t.Parallel()
	deck, err := ExtractNativePPTX(nativeDiagramLayoutFixture(t, nativeDiagramLayoutFixtureOptions{}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract exact diagram: %v", err)
	}
	slide := deck.Slides[0]
	for _, element := range slide.Elements {
		if element.Kind == NativeElementKindGroup {
			t.Fatalf("exact tier must not lay out the diagram: %#v", element)
		}
	}
	found := false
	for _, diagnostic := range slide.Compatibility.Diagnostics {
		found = found || diagnostic.Code == nativeDiagramDrawingEmptyCode
	}
	if !found {
		t.Fatalf("exact tier lost the empty-fallback refusal: %v", nativeDiagramLayoutCodes(NativeElement{Compatibility: slide.Compatibility}))
	}
}

func TestExtractNativePPTXDiagramLayoutElementsRemainPreviewOnlyForMutation(t *testing.T) {
	t.Parallel()
	original := nativeDiagramLayoutFixture(t, nativeDiagramLayoutFixtureOptions{})
	deck, err := ExtractNativePPTX(original, nativeDiagramLayoutApproximateOptions())
	if err != nil {
		t.Fatalf("extract laid-out diagram: %v", err)
	}
	group := nativeFixtureDiagramGroup(t, deck.Slides[0])
	text := "Renamed"
	align := NativeTextAlignCenter
	paragraphs := []NativeParagraph{{Runs: []NativeTextRun{{Text: &text}}, Align: &align, Level: int64Pointer(0), Bullet: boolPointer(false)}}
	for _, target := range []NativeElement{group.Children[len(group.Children)-1], group} {
		request := NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{
			OperationID: "op-1", Kind: NativePPTXReplaceText, ElementID: target.ID, ExpectedFingerprintSHA256: target.Source.FingerprintSHA256, Paragraphs: &paragraphs,
		}}}
		if produced, err := ApplyNativePPTXMutations(original, request); err == nil || produced != nil {
			t.Fatalf("laid-out diagram element %q accepted a mutation", target.ID)
		}
	}
}

func assertNativeDiagramLayoutRefused(t *testing.T, options nativeDiagramLayoutFixtureOptions, code string) {
	t.Helper()
	deck, err := ExtractNativePPTX(nativeDiagramLayoutFixture(t, options), nativeDiagramLayoutApproximateOptions())
	if err != nil {
		t.Fatalf("extract refused diagram (%s): %v", code, err)
	}
	slide := deck.Slides[0]
	for _, element := range slide.Elements {
		if element.Kind == NativeElementKindGroup {
			t.Fatalf("refused diagram leaked a partial layout (%s): %#v", code, element)
		}
	}
	codes := []string{}
	for _, diagnostic := range slide.Compatibility.Diagnostics {
		codes = append(codes, diagnostic.Code)
	}
	// The refused frame keeps an empty region, so the slide reports a refusal.
	if !strings.Contains(strings.Join(codes, ","), code) || slide.Compatibility.Status != NativeCompatibilityStatusRefused {
		t.Fatalf("diagram frame was not refused with %s: %v", code, codes)
	}
	if len(slide.Elements) != 2 || slide.Elements[0].Kind != NativeElementKindText {
		t.Fatalf("sibling text element was lost by the diagram refusal: %#v", slide.Elements)
	}
	if regions := nativeRefusedFrameRegions(t, slide.Elements, code); len(regions) != 1 {
		t.Fatalf("refused diagram frame did not keep an empty region: %#v", slide.Elements)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("invalid refused diagram deck: %#v", issues)
	}
}

func TestExtractNativePPTXDiagramLayoutRefusesOutsideTheSubset(t *testing.T) {
	t.Parallel()
	layout := nativeDiagramLayoutLayoutXML(nativeDiagramURITransitional)
	// Algorithms other than the hierarchy subset never produce a partial chart.
	cycle := strings.Replace(layout, `<dgm:alg type="hierChild"><dgm:param type="linDir" val="fromL"/></dgm:alg>`, `<dgm:alg type="cycle"/>`, 1)
	if cycle == layout {
		t.Fatal("layout fixture drifted")
	}
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{layout: cycle}, nativeDiagramLayoutAlgorithmCode)
	// Unknown constraint operators and malformed values.
	sibSp := `<dgm:constr type="sibSp" refType="w" refFor="des" refForName="rootComposite1" fact="0.21"/>`
	if !strings.Contains(layout, sibSp) || !strings.Contains(layout, `<dgm:constr type="primFontSz" val="65"/>`) {
		t.Fatal("layout fixture drifted")
	}
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{layout: strings.Replace(layout, sibSp, strings.Replace(sibSp, `/>`, ` op="mul"/>`, 1), 1)}, nativeDiagramLayoutConstraintCode)
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{layout: strings.Replace(layout, `<dgm:constr type="primFontSz" val="65"/>`, `<dgm:constr type="primFontSz" val="wide"/>`, 1)}, nativeDiagramLayoutConstraintCode)
	// A reference to a constraint that was never defined is not defaulted.
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{layout: strings.Replace(layout, `refType="sp" fact="0.5"/>`, `refType="diam" fact="0.5"/>`, -1)}, nativeDiagramLayoutConstraintCode)
	// Unmodeled condition functions and axes.
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{layout: strings.Replace(layout, `func="depth"`, `func="pos"`, 1)}, nativeDiagramLayoutDefinitionCode)
	// A parOf cycle in the data model.
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{extraCxn: `<dgm:cxn modelId="{C9}" srcId="{EMP}" destId="{MGR}" srcOrd="0" destOrd="0"/>`}, nativeDiagramLayoutDataCode)
	// Style labels come from the layout or the authored presentation points only.
	data := nativeDiagramLayoutDataXML(nativeDiagramURITransitional, nsDrawingTransitional, "", "", "")
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{dataXML: strings.Replace(data, `presStyleLbl="node2"`, `presStyleLbl=""`, -1)}, nativeDiagramLayoutStyleCode)
	// Authored geometry customizations are not applied silently.
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{dataXML: strings.Replace(data, `<dgm:prSet phldrT="[Text]"/>`, `<dgm:prSet phldrT="[Text]" custScaleX="150000"/>`, 1)}, nativeDiagramLayoutDataCode)
}

func TestExtractNativePPTXDiagramLayoutBudgetsAreBounded(t *testing.T) {
	t.Parallel()
	var points, connections strings.Builder
	for index := 0; index < nativeMaxDiagramLayoutPoints; index++ {
		id := fmt.Sprintf("{X%03d}", index)
		points.WriteString(nativeDiagramLayoutPointXML(id, "", "N"))
		connections.WriteString(`<dgm:cxn modelId="{CX` + id + `}" srcId="{MGR2}" destId="` + id + `" srcOrd="` + fmt.Sprint(index) + `" destOrd="0"/>`)
	}
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{extraPts: points.String(), extraCxn: connections.String()}, nativeDiagramLayoutBudgetCode)
}

func TestExtractNativePPTXDiagramLayoutRefusesReviewedAdversarialInputs(t *testing.T) {
	t.Parallel()
	layout := nativeDiagramLayoutLayoutXML(nativeDiagramURITransitional)
	// A gigantic primFontSz must refuse instead of spinning the fitter.
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{layout: strings.Replace(layout, `<dgm:constr type="primFontSz" val="65"/>`, `<dgm:constr type="primFontSz" val="100000000000"/>`, -1)}, nativeDiagramLayoutConstraintCode)
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{layout: strings.Replace(layout, `<dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/>`, `<dgm:rule type="primFontSz" val="900" fact="NaN" max="NaN"/>`, -1)}, nativeDiagramLayoutConstraintCode)
	// A closed parOf cycle disconnected from the document point.
	cycle := nativeDiagramLayoutPointXML("{CA}", "", "Loop A") + nativeDiagramLayoutPointXML("{CB}", "", "Loop B")
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{extraPts: cycle, extraCxn: `<dgm:cxn modelId="{CC1}" srcId="{CA}" destId="{CB}" srcOrd="0" destOrd="0"/><dgm:cxn modelId="{CC2}" srcId="{CB}" destId="{CA}" srcOrd="0" destOrd="0"/>`}, nativeDiagramLayoutDataCode)
	// Shapes nested under sp/tx nodes are refused, not dropped.
	nested := strings.Replace(layout, `<dgm:layoutNode name="rootConnector1" styleLbl="node1" moveWith="rootText1"><dgm:alg type="sp"/><dgm:shape type="rect" hideGeom="1"><dgm:adjLst/></dgm:shape><dgm:presOf axis="self" ptType="node" cnt="1"/><dgm:constrLst/><dgm:ruleLst/>`,
		`<dgm:layoutNode name="rootConnector1" styleLbl="node1" moveWith="rootText1"><dgm:alg type="sp"/><dgm:shape type="rect" hideGeom="1"><dgm:adjLst/></dgm:shape><dgm:presOf axis="self" ptType="node" cnt="1"/><dgm:constrLst/><dgm:ruleLst/><dgm:layoutNode name="hidden" styleLbl="node1"><dgm:alg type="sp"/><dgm:shape type="ellipse"><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode>`, 1)
	if nested == layout {
		t.Fatal("layout fixture drifted")
	}
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{layout: nested}, nativeDiagramLayoutAlgorithmCode)
	// Constraint types the subset does not consume, and rules aimed at other nodes.
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{layout: strings.Replace(layout, `<dgm:constr type="alignOff"/>`, `<dgm:constr type="alignOff"/><dgm:constr type="wOff" val="10"/>`, 1)}, nativeDiagramLayoutConstraintCode)
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{layout: strings.Replace(layout, `<dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/>`, `<dgm:rule type="primFontSz" for="ch" forName="rootText1" val="5" fact="NaN" max="NaN"/>`, 1)}, nativeDiagramLayoutConstraintCode)
}

func TestExtractNativePPTXDiagramLayoutHiddenConnectorsAreNotEmitted(t *testing.T) {
	t.Parallel()
	layout := nativeDiagramLayoutLayoutXML(nativeDiagramURITransitional)
	marker := `<dgm:layoutNode name="Name111" styleLbl="parChTrans1D2">`
	index := strings.Index(layout, marker)
	if index < 0 {
		t.Fatal("layout fixture drifted")
	}
	tail := strings.Replace(layout[index:], `<dgm:shape type="conn" zOrderOff="-99999">`, `<dgm:shape type="conn" zOrderOff="-99999" hideGeom="1">`, 1)
	deck, err := ExtractNativePPTX(nativeDiagramLayoutFixture(t, nativeDiagramLayoutFixtureOptions{layout: layout[:index] + tail}), nativeDiagramLayoutApproximateOptions())
	if err != nil {
		t.Fatalf("extract: %v", err)
	}
	group := nativeFixtureDiagramGroup(t, deck.Slides[0])
	connectors := 0
	for _, child := range group.Children {
		if child.Kind == NativeElementKindConnector {
			connectors++
		}
	}
	if len(group.Children) != 7 || connectors != 2 {
		t.Fatalf("hidden assistant connector must not paint: %d children, %d connectors", len(group.Children), connectors)
	}
}

// hierAlign tL is laid out as a hanging block below the root (declared
// deviation from the §21.4.7.36 above-the-parent wording): a grandchild under
// Employee hangs left-aligned at alignOff x width below its parent.
func TestExtractNativePPTXDiagramLayoutHangingLeafChildrenDeclareTheTLDeviation(t *testing.T) {
	t.Parallel()
	extraPts := nativeDiagramLayoutPointXML("{GC}", "", "Intern") + nativeDiagramLayoutPointXML("{GC-PT}", "parTrans", "") + nativeDiagramLayoutPointXML("{GC-ST}", "sibTrans", "") +
		nativeDiagramLayoutPresPointXML("{P-GC}", "{GC}", "rootText", "node2")
	extraCxn := `<dgm:cxn modelId="{C7}" srcId="{EMP}" destId="{GC}" srcOrd="0" destOrd="0" parTransId="{GC-PT}" sibTransId="{GC-ST}"/>`
	deck, err := ExtractNativePPTX(nativeDiagramLayoutFixture(t, nativeDiagramLayoutFixtureOptions{extraPts: extraPts, extraCxn: extraCxn}), nativeDiagramLayoutApproximateOptions())
	if err != nil {
		t.Fatalf("extract: %v", err)
	}
	group := nativeFixtureDiagramGroup(t, deck.Slides[0])
	if !strings.Contains(group.Compatibility.Diagnostics[0].Message, "tL/tR laid out as hanging blocks below the root") {
		t.Fatalf("tL deviation is not disclosed: %s", group.Compatibility.Diagnostics[0].Message)
	}
	shapes := map[string]NativeElement{}
	connectors := 0
	for _, child := range group.Children {
		if child.Kind == NativeElementKindConnector {
			connectors++
			continue
		}
		shapes[nativeDiagramLayoutChildText(child)] = child
	}
	employee, intern := shapes["Employee"], shapes["Intern"]
	if len(shapes) != 6 || connectors != 4 || employee.ID == "" || intern.ID == "" {
		t.Fatalf("expected six shapes and four connectors: %d shapes, %d connectors", len(shapes), connectors)
	}
	width := float64(*employee.Transform.Cx)
	if dx := float64(*intern.Transform.X - *employee.Transform.X); dx < 0.25*width-2 || dx > 0.25*width+2 {
		t.Fatalf("hanging child left edge must sit alignOff x width right of its parent: dx=%v width=%v", dx, width)
	}
	if dy := float64(*intern.Transform.Y - (*employee.Transform.Y + *employee.Transform.Cy)); dy < 0.21*width-2 || dy > 0.21*width+2 {
		t.Fatalf("hanging child must sit sp below its parent: dy=%v width=%v", dy, width)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("invalid deck: %#v", issues)
	}
}

// dgm:forEach and dgm:presOf declare axis (ST_AxisTypes), ptType
// (ST_ElementTypes), st and step (ST_Ints) and cnt (ST_UnsignedInts) — every
// one of them a LIST whose entries pair up with the axis steps (ECMA-376
// Part 1 §21.4.7.6, §21.4.7.40, §21.4.7.63). st, cnt and step used to be
// parsed as single integers and applied once after the whole chain, so a
// perfectly ordinary two-axis selection refused the diagram outright. Two
// SmartArt decks in the hard-v2 corpus are written this way.
func TestExtractNativePPTXDiagramLayoutPairsSelectionListsWithAxes(t *testing.T) {
	t.Parallel()
	layout := nativeDiagramLayoutLayoutXML(nativeDiagramURITransitional)
	const single = `<dgm:forEach name="rep2a" axis="ch" ptType="nonAsst">`
	if !strings.Contains(layout, single) {
		t.Fatal("layout fixture drifted: the rep2a forEach is gone")
	}
	children := func(t *testing.T, selection string) int {
		t.Helper()
		fixture := nativeDiagramLayoutFixture(t, nativeDiagramLayoutFixtureOptions{
			layout: strings.Replace(layout, single, selection, 1),
		})
		deck, err := ExtractNativePPTX(fixture, nativeDiagramLayoutApproximateOptions())
		if err != nil {
			t.Fatalf("extract diagram with selection %q: %v", selection, err)
		}
		group := nativeFixtureDiagramGroup(t, deck.Slides[0])
		if group.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
			t.Fatalf("selection %q refused the diagram: %s", selection, nativeDiagramLayoutCodes(group))
		}
		return len(group.Children)
	}

	// "ch self" with a per-axis window that trims nothing selects exactly what
	// the one-axis form does.
	all := children(t, `<dgm:forEach name="rep2a" axis="ch self" ptType="nonAsst node" st="1 1" cnt="0 0" step="1 1">`)
	if all != 8 {
		t.Fatalf("a no-op two-axis window changed the selection: %d children, want 8", all)
	}

	// cnt="1 0" keeps one point at the FIRST axis step and everything at the
	// second. Applying the window once after the chain — or reading only the
	// last entry — would leave the selection untrimmed at 8.
	trimmed := children(t, `<dgm:forEach name="rep2a" axis="ch self" ptType="nonAsst node" st="1 1" cnt="1 0">`)
	if trimmed >= all {
		t.Fatalf("cnt=\"1 0\" did not trim the first axis step: %d children, want fewer than %d", trimmed, all)
	}
}

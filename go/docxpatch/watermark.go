// D11 breadth: watermarks. A real docx watermark is NOT body content — it
// lives in a HEADER part as a VML shape (a <w:pict>/<v:shape> using the
// same "_x0000_t136" text-path shapetype Word's own Insert > Watermark
// feature has emitted for 20+ years, kept for broad compatibility even in
// modern .docx). This is genuinely new territory for this package: every
// prior D11 extension (images, notes, charts) only ever touched
// word/document.xml plus a new content part; a watermark requires finding
// or creating a HEADER part and wiring it into the section properties
// (word/document.xml's <w:sectPr>) — a part of the schema this patcher had
// not needed to reach into before.
//
// Scope, stated plainly:
//   - TEXT watermarks only. An image watermark would reuse imagewrite.go's
//     media-part machinery with a <v:imagedata> shape instead of
//     <v:textpath> — a real, smaller follow-up once this lands, not
//     attempted here.
//   - Single-section documents only (exactly one <w:sectPr> in
//     word/document.xml). A multi-section document (different headers per
//     section — a section-break mid-body) is refused with a clear error
//     rather than silently watermarking only one section or corrupting the
//     others; handling every section is a real follow-up, not attempted
//     here.
//   - Only the DEFAULT header (w:type="default") is targeted — first-page-
//     different and even/odd distinct headers are a Word feature this
//     writer doesn't create or touch.
//
// If the document already has a default header, the watermark paragraph
// is APPENDED to it, preserving whatever header content already existed;
// if not, a fresh header part is created carrying only the watermark.
//
// Honest validation limit: python-docx has no VML/watermark object model,
// so validation here (like themes/charts) uses low-level docx.oxml/lxml —
// which proves the header part/relationship/content-type wiring and the
// VML shape's structure and watermark text are all real and correct. It
// does NOT prove visual rendering (rotation, opacity, position) — that
// needs an actual Word/LibreOffice render, and LibreOffice headless isn't
// installed on this machine (checked, not assumed). Flagged rather than
// silently skipped.
package docxpatch

import (
	"archive/zip"
	"bytes"
	"fmt"
	"regexp"
	"strings"
)

const (
	relTypeHeader = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
	ctHeader      = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
)

// WatermarkSpec describes a text watermark to insert into the document's
// default header.
type WatermarkSpec struct {
	Text string
	// FontFamily defaults to "Calibri" when empty.
	FontFamily string
	// ColorHex defaults to "808080" (Word's own watermark gray) when empty. No '#'.
	ColorHex string
	// Horizontal false (the default) rotates the text diagonally like
	// Word's own default watermark (rotation:315); set true to lay it out
	// flat (rotation:0) instead.
	Horizontal bool
}

func (s WatermarkSpec) fontOrDefault() string {
	if s.FontFamily != "" {
		return s.FontFamily
	}
	return "Calibri"
}

func (s WatermarkSpec) colorOrDefault() string {
	if s.ColorHex != "" {
		return strings.ToUpper(strings.TrimPrefix(s.ColorHex, "#"))
	}
	return "808080"
}

// InsertWatermark inserts a text watermark into the document's default
// header, creating the header if none exists yet.
func InsertWatermark(docx []byte, spec WatermarkSpec) ([]byte, error) {
	if strings.TrimSpace(spec.Text) == "" {
		return nil, fmt.Errorf("docxpatch: empty watermark text")
	}
	if !hexColorRe.MatchString(spec.colorOrDefault()) {
		return nil, fmt.Errorf("docxpatch: watermark color %q is not a 6-digit hex color", spec.ColorHex)
	}

	zr, err := zip.NewReader(bytes.NewReader(docx), int64(len(docx)))
	if err != nil {
		return nil, fmt.Errorf("docxpatch: not a readable .docx: %w", err)
	}
	docRaw, err := readPart(zr, docPart)
	if err != nil {
		return nil, fmt.Errorf("docxpatch: %s not found: %w", docPart, err)
	}
	docXML := string(docRaw)

	if n := strings.Count(docXML, "<w:sectPr"); n != 1 {
		return nil, fmt.Errorf("docxpatch: watermark requires exactly one section (<w:sectPr>), found %d — multi-section documents are not supported", n)
	}

	relsXML, hasRels := readOptionalPart(zr, docRelsPart)
	if !hasRels {
		relsXML = emptyRelsXML
	}
	ctXML, err := readPart(zr, contentTypes)
	if err != nil {
		return nil, fmt.Errorf("docxpatch: %s not found: %w", contentTypes, err)
	}

	watermarkPara := watermarkParagraphXML(spec)

	if existingRelID, ok := defaultHeaderRelID(docXML); ok {
		// Append to the existing default header.
		headerPart, ok := resolveDocRelTarget(relsXML, existingRelID)
		if !ok {
			return nil, fmt.Errorf("docxpatch: sectPr references header relationship %q but it's not in %s", existingRelID, docRelsPart)
		}
		headerRaw, err := readPart(zr, headerPart)
		if err != nil {
			return nil, fmt.Errorf("docxpatch: header part %q referenced but not found: %w", headerPart, err)
		}
		newHeaderXML, err := appendParagraphToHeader(string(headerRaw), watermarkPara)
		if err != nil {
			return nil, err
		}
		return ApplyPatch(docx, Patch{Replace: map[string][]byte{headerPart: []byte(newHeaderXML)}})
	}

	// No default header yet: create one.
	headerPart := nextFreeHeaderPart(zr)
	relID := nextFreeRelID(relsXML)
	newRels, err := appendRelationship(relsXML, relID, relTypeHeader, headerPart[len("word/"):])
	if err != nil {
		return nil, err
	}
	newCT, err := overridePartWith(string(ctXML), "/"+headerPart, ctHeader)
	if err != nil {
		return nil, err
	}
	newDocXML, err := insertDefaultHeaderReference(docXML, relID)
	if err != nil {
		return nil, err
	}

	patch := Patch{
		Replace: map[string][]byte{
			docPart:      []byte(newDocXML),
			contentTypes: []byte(newCT),
		},
		Add: map[string][]byte{
			headerPart: []byte(headerDocXML(watermarkPara)),
		},
	}
	if hasRels {
		patch.Replace[docRelsPart] = []byte(newRels)
	} else {
		patch.Add[docRelsPart] = []byte(newRels)
	}
	return ApplyPatch(docx, patch)
}

func readOptionalPart(zr *zip.Reader, name string) (string, bool) {
	raw, err := readPart(zr, name)
	if err != nil {
		return "", false
	}
	return string(raw), true
}

var defaultHeaderRefRe = regexp.MustCompile(`<w:headerReference[^>]*w:type="default"[^>]*r:id="([^"]+)"`)

// defaultHeaderRelID looks for an EXISTING <w:headerReference w:type="default" .../>
// anywhere in document.xml (there is exactly one sectPr per the caller's
// check, so this is unambiguous) and returns its relationship id.
func defaultHeaderRelID(docXML string) (string, bool) {
	m := defaultHeaderRefRe.FindStringSubmatch(docXML)
	if m == nil {
		return "", false
	}
	return m[1], true
}

func resolveDocRelTarget(relsXML, relID string) (string, bool) {
	re := regexp.MustCompile(`<Relationship[^>]*Id="` + regexp.QuoteMeta(relID) + `"[^>]*Target="([^"]+)"`)
	m := re.FindStringSubmatch(relsXML)
	if m == nil {
		// attribute order can vary — try Target-before-Id too
		re2 := regexp.MustCompile(`<Relationship[^>]*Target="([^"]+)"[^>]*Id="` + regexp.QuoteMeta(relID) + `"`)
		m = re2.FindStringSubmatch(relsXML)
		if m == nil {
			return "", false
		}
	}
	return resolveWordRelTarget(m[1]), true
}

var headerPartRe = regexp.MustCompile(`^word/header(\d+)\.xml$`)

func nextFreeHeaderPart(zr *zip.Reader) string {
	n := 1
	for _, f := range zr.File {
		if m := headerPartRe.FindStringSubmatch(f.Name); m != nil {
			var existing int
			fmt.Sscanf(m[1], "%d", &existing) //nolint:errcheck
			if existing >= n {
				n = existing + 1
			}
		}
	}
	return fmt.Sprintf("word/header%d.xml", n)
}

// insertDefaultHeaderReference adds <w:headerReference w:type="default"
// r:id="X"/> to the document's single <w:sectPr>, as the FIRST child
// (schema-safe regardless of what else the sectPr already has —
// CT_SectPr's sequence requires header/footer references to precede
// pgSz/pgMar/etc, and inserting first always satisfies "precedes").
func insertDefaultHeaderReference(docXML, relID string) (string, error) {
	// xmlns:r declared LOCALLY on this element — never trust that the
	// document root already bound the r: prefix (the same defensive
	// posture imageParagraphXML/chartParagraphXML already use for their
	// own r:embed/r:id attributes). A real test caught this: a fixture
	// whose root only declared xmlns:w produced a headerReference with an
	// UNDEFINED namespace prefix, invalid XML that python-docx correctly
	// refused to parse.
	ref := fmt.Sprintf(`<w:headerReference w:type="default" r:id=%q xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`, relID)
	// Self-closing sectPr (the common case for a simple/agent-generated doc): <w:sectPr .../> or <w:sectPr/>
	selfClosingRe := regexp.MustCompile(`<w:sectPr([^>]*)/>`)
	if loc := selfClosingRe.FindStringSubmatchIndex(docXML); loc != nil {
		attrs := docXML[loc[2]:loc[3]]
		open := "<w:sectPr" + attrs + ">"
		return docXML[:loc[0]] + open + ref + "</w:sectPr>" + docXML[loc[1]:], nil
	}
	// Expanded sectPr: <w:sectPr ...>...</w:sectPr> — insert right after the opening tag.
	openRe := regexp.MustCompile(`<w:sectPr[^>]*>`)
	loc := openRe.FindStringIndex(docXML)
	if loc == nil {
		return "", fmt.Errorf("docxpatch: malformed <w:sectPr>")
	}
	return docXML[:loc[1]] + ref + docXML[loc[1]:], nil
}

// appendParagraphToHeader adds a paragraph at the end of an existing
// header part's content, before </w:hdr>.
func appendParagraphToHeader(headerXML, paraXML string) (string, error) {
	idx := strings.LastIndex(headerXML, "</w:hdr>")
	if idx < 0 {
		return "", fmt.Errorf("docxpatch: malformed header part (no </w:hdr>)")
	}
	return headerXML[:idx] + paraXML + headerXML[idx:], nil
}

// headerDocXML wraps a paragraph in a fresh, minimal header part.
func headerDocXML(paraXML string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n" +
		`<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
		`xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">` +
		paraXML + `</w:hdr>`
}

// watermarkParagraphXML builds the paragraph carrying the VML watermark
// shape — the same "_x0000_t136" text-path shapetype Word's own Insert >
// Watermark feature emits, kept exactly as Word/LibreOffice expect it
// (deviating from this well-established boilerplate risks a shape that
// parses but doesn't render the curved text-path correctly).
func watermarkParagraphXML(spec WatermarkSpec) string {
	rotation := "315" // Word's default: bottom-left to top-right diagonal
	if spec.Horizontal {
		rotation = "0"
	}
	return `<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:pict>` +
		`<v:shapetype id="_x0000_t136" coordsize="1600,21600" o:spt="136" adj="10800" path="m@7,0l@8,0m@5,21600l@6,21600e">` +
		`<v:formulas>` +
		`<v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/>` +
		`<v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/>` +
		`<v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/>` +
		`<v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/>` +
		`</v:formulas>` +
		`<v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>` +
		`<v:textpath on="t" fitshape="t"/>` +
		`<v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>` +
		`<o:lock v:ext="edit" text="t" shapetype="t"/>` +
		`</v:shapetype>` +
		fmt.Sprintf(
			`<v:shape id="WordprocessingWatermark" o:spid="_x0000_s2049" type="#_x0000_t136" `+
				`style="position:absolute;margin-left:0;margin-top:0;width:415pt;height:207.5pt;`+
				`rotation:%s;z-index:-251654144;mso-position-horizontal:center;`+
				`mso-position-horizontal-relative:margin;mso-position-vertical:center;`+
				`mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="#%s" stroked="f">`,
			rotation, spec.colorOrDefault(),
		) +
		`<v:fill opacity=".5"/>` +
		fmt.Sprintf(`<v:textpath style="font-family:&quot;%s&quot;;font-size:1pt" string=%q/>`, spec.fontOrDefault(), xmlEscape(spec.Text)) +
		`</v:shape>` +
		`</w:pict></w:r></w:p>`
}

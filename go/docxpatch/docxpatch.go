// Package docxpatch edits real Word documents SURGICALLY: only the
// paragraphs an edit touches are rewritten, every other byte of the .docx —
// untouched paragraphs, styles.xml, numbering, themes, images, headers,
// footers, settings, custom parts — passes through verbatim. That is the
// property a full parse-regenerate library cannot give: a document that
// round-trips through docxpatch differs from the original ONLY where the
// text changed.
//
// The model is paragraph-level. Extract flattens every <w:p> in
// word/document.xml (body and table cells alike) to {index, style, text} in
// document order; Apply rewrites chosen paragraphs' runs while preserving
// each paragraph's own properties (<w:pPr>) and the formatting of its first
// run (<w:rPr>). Paragraphs carrying non-text content (drawings, fields,
// footnote refs) refuse a rewrite rather than silently destroying it.
//
// OOXML paragraphs never nest — <w:p> may appear inside table cells but
// never inside another <w:p> — so a linear scan over document.xml is exact,
// no XML tree required. This keeps the package dependency-free (stdlib
// only), the same discipline as xlsxpatch.
package docxpatch

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"regexp"
	"strings"
)

// Paragraph is one <w:p> of the document, in document order.
type Paragraph struct {
	// Index is the paragraph's position (0-based) across the whole
	// document, table-cell paragraphs included.
	Index int
	// Style is the paragraph style id (<w:pStyle w:val>), "" for default.
	Style string
	// Text is the concatenated run text; tabs are \t, line breaks \n.
	Text string
	// InTable is true when the paragraph lives inside a table cell.
	InTable bool
	// HasNonText is true when the paragraph carries content a text
	// rewrite would destroy (images/drawings, fields, footnotes…);
	// such paragraphs cannot be Set.
	HasNonText bool
}

// Edit is one paragraph-level operation for Apply.
type Edit struct {
	// Op: "set" (replace paragraph Index's text), "insert_after" (new
	// paragraph after Index, -1 prepends at the start), "delete".
	Op    string
	Index int
	// Text for set/insert_after. \n becomes a line break within the
	// paragraph, \t a tab.
	Text string
}

const docPart = "word/document.xml"

// --- zip plumbing (verbatim copy of every part but document.xml) ---

func readPart(zr *zip.Reader, name string) ([]byte, error) {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, fmt.Errorf("%s not found", name)
}

// rewriteZip replaces exactly one existing part, verbatim-copying every
// other entry. A thin convenience wrapper over the general ApplyPatch (see
// patch.go) for the common single-part-replace case Apply/ReplaceText use.
func rewriteZip(orig []byte, replaced string, content []byte) ([]byte, error) {
	return ApplyPatch(orig, Patch{Replace: map[string][]byte{replaced: content}})
}

// --- paragraph scanning ---

// para is one scanned paragraph with its byte extent in document.xml.
type para struct {
	start, end int // [start, end) of the whole <w:p …>…</w:p> (or <w:p/>)
	openEnd    int // end of the opening tag (0 for self-closing)
	selfClosed bool
	inTable    bool
}

var (
	rePStart = regexp.MustCompile(`<w:p(?:[ >/])`)
	rePStyle = regexp.MustCompile(`<w:pStyle [^>]*w:val="([^"]*)"`)
	reT      = regexp.MustCompile(`(?s)<w:t(?: [^>]*)?>(.*?)</w:t>|<w:t(?: [^>]*)?/>|<w:tab/>|<w:br/>|<w:cr/>`)
	rePPr    = regexp.MustCompile(`(?s)^<w:pPr(?: [^>]*)?>.*?</w:pPr>|^<w:pPr(?: [^>]*)?/>`)
	reRunPr  = regexp.MustCompile(`(?s)<w:r(?: [^>]*)?>(<w:rPr(?: [^>]*)?>.*?</w:rPr>|<w:rPr(?: [^>]*)?/>)`)
)

// nonTextMarkers — content inside a paragraph a run rewrite would destroy.
var nonTextMarkers = []string{
	"<w:drawing", "<pic:pic", "<w:object", "<w:pict", "<w:fldChar",
	"<w:fldSimple", "<w:footnoteReference", "<w:endnoteReference",
	"<w:commentReference", "<m:oMath", "<w:hyperlink", "<w:sdt",
}

// scanParas walks document.xml and returns every paragraph's extent.
// <w:p> cannot nest, so the first </w:p> after an opening tag closes it.
func scanParas(doc string) []para {
	// Table extents, precomputed: a paragraph is in a table when more
	// <w:tbl> than </w:tbl> occur before it.
	var tblOpens, tblCloses []int
	for i := 0; ; {
		j := strings.Index(doc[i:], "<w:tbl>")
		if j < 0 {
			break
		}
		tblOpens = append(tblOpens, i+j)
		i += j + 1
	}
	for i := 0; ; {
		j := strings.Index(doc[i:], "</w:tbl>")
		if j < 0 {
			break
		}
		tblCloses = append(tblCloses, i+j)
		i += j + 1
	}
	countBefore := func(positions []int, at int) int {
		n := 0
		for _, p := range positions {
			if p < at {
				n++
			}
		}
		return n
	}

	var out []para
	i := 0
	for i < len(doc) {
		loc := rePStart.FindStringIndex(doc[i:])
		if loc == nil {
			break
		}
		pAt := i + loc[0]
		inTable := countBefore(tblOpens, pAt) > countBefore(tblCloses, pAt)
		gt := strings.IndexByte(doc[pAt:], '>')
		if gt < 0 {
			break
		}
		openEnd := pAt + gt + 1
		if doc[openEnd-2] == '/' { // <w:p/> or <w:p …/>
			out = append(out, para{start: pAt, end: openEnd, selfClosed: true, inTable: inTable})
			i = openEnd
			continue
		}
		close := strings.Index(doc[openEnd:], "</w:p>")
		if close < 0 {
			break
		}
		end := openEnd + close + len("</w:p>")
		out = append(out, para{start: pAt, end: end, openEnd: openEnd, inTable: inTable})
		i = end
	}
	return out
}

func paraText(inner string) string {
	var b strings.Builder
	for _, m := range reT.FindAllString(inner, -1) {
		switch {
		case strings.HasPrefix(m, "<w:tab/>"):
			b.WriteByte('\t')
		case strings.HasPrefix(m, "<w:br/>"), strings.HasPrefix(m, "<w:cr/>"):
			b.WriteByte('\n')
		default:
			sub := reT.FindStringSubmatch(m)
			if len(sub) > 1 {
				b.WriteString(xmlUnescape(sub[1]))
			}
		}
	}
	return b.String()
}

func hasNonText(inner string) bool {
	for _, m := range nonTextMarkers {
		if strings.Contains(inner, m) {
			return true
		}
	}
	return false
}

func xmlEscape(s string) string {
	r := strings.NewReplacer("&", "&amp;", "<", "&lt;", ">", "&gt;")
	return r.Replace(s)
}

func xmlUnescape(s string) string {
	r := strings.NewReplacer("&lt;", "<", "&gt;", ">", "&quot;", `"`, "&apos;", "'", "&amp;", "&")
	return r.Replace(s)
}

// Extract lists the document's paragraphs in order.
func Extract(docx []byte) ([]Paragraph, error) {
	zr, err := zip.NewReader(bytes.NewReader(docx), int64(len(docx)))
	if err != nil {
		return nil, fmt.Errorf("not a readable .docx: %w", err)
	}
	doc, err := readPart(zr, docPart)
	if err != nil {
		return nil, err
	}
	s := string(doc)
	var out []Paragraph
	for idx, p := range scanParas(s) {
		pg := Paragraph{Index: idx, InTable: p.inTable}
		if !p.selfClosed {
			inner := s[p.openEnd : p.end-len("</w:p>")]
			pg.Text = paraText(inner)
			pg.HasNonText = hasNonText(inner)
			if m := rePStyle.FindStringSubmatch(inner); m != nil {
				pg.Style = m[1]
			}
		}
		out = append(out, pg)
	}
	return out, nil
}

// buildRuns renders text as runs, honoring \n (line break) and \t (tab),
// with rPr applied to every text run.
func buildRuns(text, rPr string) string {
	var b strings.Builder
	b.WriteString("<w:r>")
	b.WriteString(rPr)
	flushText := func(t string) {
		if t == "" {
			return
		}
		b.WriteString(`<w:t xml:space="preserve">`)
		b.WriteString(xmlEscape(t))
		b.WriteString("</w:t>")
	}
	cur := strings.Builder{}
	for _, r := range text {
		switch r {
		case '\n':
			flushText(cur.String())
			cur.Reset()
			b.WriteString("<w:br/>")
		case '\t':
			flushText(cur.String())
			cur.Reset()
			b.WriteString("<w:tab/>")
		default:
			cur.WriteRune(r)
		}
	}
	flushText(cur.String())
	b.WriteString("</w:r>")
	return b.String()
}

// rewriteParagraph rebuilds one paragraph's XML with new text: opening tag
// and <w:pPr> verbatim, first run's <w:rPr> carried onto the new runs.
func rewriteParagraph(s string, p para, text string) (string, error) {
	if p.selfClosed {
		// An empty <w:p/> has no pPr/rPr to preserve.
		open := strings.TrimSuffix(s[p.start:p.end], "/>") + ">"
		return open + buildRuns(text, "") + "</w:p>", nil
	}
	inner := s[p.openEnd : p.end-len("</w:p>")]
	if hasNonText(inner) {
		return "", fmt.Errorf("paragraph contains non-text content (image, field, link or similar) — edit it in Word")
	}
	pPr := rePPr.FindString(inner)
	rPr := ""
	if m := reRunPr.FindStringSubmatch(inner); m != nil {
		rPr = m[1]
	}
	return s[p.start:p.openEnd] + pPr + buildRuns(text, rPr) + "</w:p>", nil
}

// Apply performs the edits and returns the new .docx. Indices refer to the
// ORIGINAL document (as returned by Extract); edits are applied bottom-up so
// they never invalidate each other. One failing edit fails the whole Apply —
// no partial writes.
func Apply(docx []byte, edits []Edit) ([]byte, error) {
	if len(edits) == 0 {
		return docx, nil
	}
	zr, err := zip.NewReader(bytes.NewReader(docx), int64(len(docx)))
	if err != nil {
		return nil, fmt.Errorf("not a readable .docx: %w", err)
	}
	doc, err := readPart(zr, docPart)
	if err != nil {
		return nil, err
	}
	s := string(doc)
	paras := scanParas(s)

	// Sort bottom-up by index (stable across ops at the same index:
	// keep given order, applied last-first).
	ordered := make([]Edit, len(edits))
	copy(ordered, edits)
	for i := 0; i < len(ordered); i++ {
		for j := i + 1; j < len(ordered); j++ {
			if ordered[j].Index > ordered[i].Index {
				ordered[i], ordered[j] = ordered[j], ordered[i]
			}
		}
	}

	for _, e := range ordered {
		switch e.Op {
		case "set":
			if e.Index < 0 || e.Index >= len(paras) {
				return nil, fmt.Errorf("set: paragraph %d does not exist (document has %d)", e.Index, len(paras))
			}
			p := paras[e.Index]
			nx, err := rewriteParagraph(s, p, e.Text)
			if err != nil {
				return nil, fmt.Errorf("set paragraph %d: %w", e.Index, err)
			}
			s = s[:p.start] + nx + s[p.end:]
		case "insert_after":
			// Clone paragraph formatting from the anchor (or the first
			// paragraph when prepending at -1).
			src := e.Index
			if src < 0 {
				src = 0
			}
			if len(paras) == 0 || src >= len(paras) {
				return nil, fmt.Errorf("insert_after: paragraph %d does not exist", e.Index)
			}
			p := paras[src]
			var open, pPr, rPr string
			if p.selfClosed {
				open = strings.TrimSuffix(s[p.start:p.end], "/>") + ">"
			} else {
				open = s[p.start:p.openEnd]
				inner := s[p.openEnd : p.end-len("</w:p>")]
				pPr = rePPr.FindString(inner)
				if m := reRunPr.FindStringSubmatch(inner); m != nil {
					rPr = m[1]
				}
			}
			np := open + pPr + buildRuns(e.Text, rPr) + "</w:p>"
			at := p.end
			if e.Index < 0 {
				at = p.start
			}
			s = s[:at] + np + s[at:]
		case "delete":
			if e.Index < 0 || e.Index >= len(paras) {
				return nil, fmt.Errorf("delete: paragraph %d does not exist", e.Index)
			}
			p := paras[e.Index]
			s = s[:p.start] + s[p.end:]
		default:
			return nil, fmt.Errorf("unknown op %q (use set, insert_after, delete)", e.Op)
		}
	}
	return rewriteZip(docx, docPart, []byte(s))
}

// ReplaceText replaces every occurrence of old with new across the document
// (matching on each paragraph's concatenated text, so a phrase split across
// runs still matches). Returns the new bytes and how many paragraphs
// changed. Paragraphs with non-text content are skipped and counted in
// skipped.
func ReplaceText(docx []byte, old, new string) (out []byte, changed, skipped int, err error) {
	if old == "" {
		return docx, 0, 0, fmt.Errorf("empty search text")
	}
	paras, err := Extract(docx)
	if err != nil {
		return nil, 0, 0, err
	}
	var edits []Edit
	for _, p := range paras {
		if !strings.Contains(p.Text, old) {
			continue
		}
		if p.HasNonText {
			skipped++
			continue
		}
		edits = append(edits, Edit{Op: "set", Index: p.Index, Text: strings.ReplaceAll(p.Text, old, new)})
	}
	if len(edits) == 0 {
		return docx, 0, skipped, nil
	}
	out, err = Apply(docx, edits)
	return out, len(edits), skipped, err
}

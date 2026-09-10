package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"net/url"
	"regexp"
	"sort"
	"strings"
)

// Chart WRITING: turn an InjOffice ChartSpec into a real OOXML chart inside
// the workbook, so charts made in the browser exist in the .xlsx itself and
// open in Excel / Google Sheets / LibreOffice.
//
// AddChart computes a Patch (new chart part + new-or-extended drawing part +
// rels + content types + worksheet hookup) and runs it through Apply — which
// means every fidelity guarantee holds by construction: nothing outside the
// explicitly patched parts can change, verified byte-for-byte.
//
// Supported types mirror the categorical/simple family the reader maps:
// column, bar, line, area, pie, doughnut, scatter. Specialty renderer types
// (heatmap, sankey, ...) have no OOXML chart equivalent and stay browser-side.

// WriteSeries is one series to write: display name plus the cell references
// its data binds to, in "Sheet!$A$1:$A$5" form.
type WriteSeries struct {
	Name                                 string
	NameRef                              string // optional; when set, wins over Name as the tx strRef
	CategoriesRef                        string // optional
	ValuesRef                            string
	nameCache, categoryCache, valueCache string // derived only from source workbook bytes
}

// ChartAnchor places the chart over the grid in cell coordinates (0-indexed,
// "from" inclusive top-left, "to" exclusive-ish bottom-right per OOXML
// twoCellAnchor semantics).
type ChartAnchor struct {
	FromCol, FromRow, ToCol, ToRow int
}

// ChartWriteSpec is what AddChart needs. Types follow ChartSpec vocabulary.
type ChartWriteSpec struct {
	SheetName string
	Type      string
	Title     string
	Series    []WriteSeries
	Anchor    ChartAnchor
}

var writableTypes = map[string]bool{
	"column": true, "bar": true, "line": true, "area": true,
	"pie": true, "doughnut": true, "scatter": true,
}

// AddChart returns new workbook bytes containing the chart. The original is
// untouched on any error.
func AddChart(orig []byte, spec ChartWriteSpec) ([]byte, error) {
	if err := validateChartWriteSpec(spec); err != nil {
		return nil, fmt.Errorf("xlsxpatch: add chart: %w", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: add chart: %w", err)
	}
	read := func(name string) (string, bool) {
		for _, f := range zr.File {
			if f.Name == name {
				rc, err := f.Open()
				if err != nil {
					return "", false
				}
				defer rc.Close()
				b, err := io.ReadAll(rc)
				if err != nil {
					return "", false
				}
				return string(b), true
			}
		}
		return "", false
	}

	sheetPart, err := worksheetPartFor(read, spec.SheetName)
	if err != nil {
		return nil, err
	}
	sheetXML, ok := read(sheetPart)
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: worksheet part %q unreadable", sheetPart)
	}

	spec = chartSpecWithCaches(read, spec)
	patch := Patch{Replace: map[string][]byte{}, Add: map[string][]byte{}}

	chartPart := nextFreePart(zr, "xl/charts/chart", ".xml")
	patch.Add[chartPart] = []byte(chartSpaceXML(spec))

	sheetRelsPart := relsPartFor(sheetPart)
	sheetRels, hasSheetRels := read(sheetRelsPart)

	// One drawing element per worksheet is the schema's rule: when the sheet
	// already has a drawing (agent-generated chart, images...), extend that
	// drawing part; otherwise create a fresh one and hook it up.
	existingDrawingRelID := drawingRelIDInSheet(sheetXML)
	if existingDrawingRelID != "" && hasSheetRels {
		drawingPart, err := relTarget(sheetRels, existingDrawingRelID, "xl/worksheets")
		if err != nil {
			return nil, err
		}
		drawingXML, ok := read(drawingPart)
		if !ok {
			return nil, fmt.Errorf("xlsxpatch: drawing part %q unreadable", drawingPart)
		}
		index, err := indexShapeDrawing([]byte(drawingXML))
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: add chart: drawing part %q: %w", drawingPart, err)
		}
		if index.maxObjectID == ^uint32(0) {
			return nil, fmt.Errorf("xlsxpatch: add chart: drawing part %q has exhausted DrawingML object ids", drawingPart)
		}
		drawingRelsPart := relsPartFor(drawingPart)
		drawingRels, hasDrawingRels := read(drawingRelsPart)
		if !hasDrawingRels {
			drawingRels = emptyRelsXML
		}
		chartRelID := nextFreeRelID(drawingRels)
		newDrawingRels, err := appendRelationship(drawingRels, chartRelID, relTypeChart, partRelTargetFrom(drawingPart, chartPart))
		if err != nil {
			return nil, err
		}
		newDrawingXML, err := appendAnchorToDrawing(drawingXML, anchorXMLWithObjectID(spec, chartRelID, index.maxObjectID+1))
		if err != nil {
			return nil, err
		}
		patch.Replace[drawingPart] = []byte(newDrawingXML)
		if hasDrawingRels {
			patch.Replace[drawingRelsPart] = []byte(newDrawingRels)
		} else {
			patch.Add[drawingRelsPart] = []byte(newDrawingRels)
		}
		ct, ok := read("[Content_Types].xml")
		if !ok {
			return nil, fmt.Errorf("xlsxpatch: missing [Content_Types].xml")
		}
		newCT, err := contentTypesWith(ct, map[string]string{"/" + chartPart: ctChart})
		if err != nil {
			return nil, err
		}
		patch.Replace["[Content_Types].xml"] = []byte(newCT)
		return Apply(orig, patch)
	}

	// Fresh drawing path.
	drawingPart := nextFreePart(zr, "xl/drawings/drawing", ".xml")
	patch.Add[drawingPart] = []byte(drawingXMLDoc(anchorXMLWithObjectID(spec, "rId1", 1)))
	patch.Add[relsPartFor(drawingPart)] = []byte(relsDoc(relationshipXML("rId1", relTypeChart, partRelTargetFrom(drawingPart, chartPart))))

	if !hasSheetRels {
		sheetRels = emptyRelsXML
	}
	drawingRelID := nextFreeRelID(sheetRels)
	newSheetRels, err := appendRelationship(sheetRels, drawingRelID, relTypeDrawing, partRelTargetFrom(sheetPart, drawingPart))
	if err != nil {
		return nil, err
	}
	if hasSheetRels {
		patch.Replace[sheetRelsPart] = []byte(newSheetRels)
	} else {
		patch.Add[sheetRelsPart] = []byte(newSheetRels)
	}

	newSheetXML, err := insertDrawingElement(sheetXML, drawingRelID)
	if err != nil {
		return nil, err
	}
	patch.Replace[sheetPart] = []byte(newSheetXML)

	ct, ok := read("[Content_Types].xml")
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: missing [Content_Types].xml")
	}
	newCT, err := contentTypesWith(ct, map[string]string{
		"/" + chartPart:   ctChart,
		"/" + drawingPart: ctDrawing,
	})
	if err != nil {
		return nil, err
	}
	patch.Replace["[Content_Types].xml"] = []byte(newCT)

	return Apply(orig, patch)
}

func validateChartWriteSpec(spec ChartWriteSpec) error {
	if !writableTypes[spec.Type] {
		return fmt.Errorf("chart type %q has no OOXML mapping (browser-only type)", spec.Type)
	}
	if len(spec.Series) == 0 {
		return fmt.Errorf("chart needs at least one series")
	}
	if spec.Anchor.FromCol < 0 || spec.Anchor.FromRow < 0 || spec.Anchor.ToCol <= spec.Anchor.FromCol || spec.Anchor.ToRow <= spec.Anchor.FromRow {
		return fmt.Errorf("degenerate chart anchor")
	}
	return nil
}

const (
	relTypeChart                 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
	relTypeDrawing               = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
	relTypeWorksheetTransitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
	relTypeWorksheetStrict       = "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"
	ctChart                      = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
	ctDrawing                    = "application/vnd.openxmlformats-officedocument.drawing+xml"
	emptyRelsXML                 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n" +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`
)

func esc(s string) string {
	var b bytes.Buffer
	_ = xml.EscapeText(&b, []byte(s))
	return b.String()
}

// worksheetPartFor resolves a sheet name to its part path via workbook.xml +
// its rels.
func worksheetPartFor(read func(string) (string, bool), sheetName string) (string, error) {
	wb, ok := read("xl/workbook.xml")
	if !ok {
		return "", fmt.Errorf("xlsxpatch: missing xl/workbook.xml")
	}
	relID := ""
	dec := xml.NewDecoder(strings.NewReader(wb))
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return "", fmt.Errorf("xlsxpatch: parse workbook.xml: %w", err)
		}
		if se, isStart := tok.(xml.StartElement); isStart && se.Name.Local == "sheet" {
			name, rid := "", ""
			for _, a := range se.Attr {
				switch a.Name.Local {
				case "name":
					name = a.Value
				case "id": // r:id
					rid = a.Value
				}
			}
			if name == sheetName {
				relID = rid
				break
			}
		}
	}
	if relID == "" {
		return "", fmt.Errorf("xlsxpatch: sheet %q not found in workbook", sheetName)
	}
	rels, ok := read("xl/_rels/workbook.xml.rels")
	if !ok {
		return "", fmt.Errorf("xlsxpatch: missing workbook rels")
	}
	return relTargetOfType(rels, relID, "xl", relTypeWorksheetTransitional, relTypeWorksheetStrict)
}

// relTarget resolves a relationship id to a zip part path, relative to the
// rels owner's base directory.
func relTarget(relsXML, relID, baseDir string) (string, error) {
	return relTargetOfType(relsXML, relID, baseDir)
}

func relTargetOfType(relsXML, relID, baseDir string, allowedTypes ...string) (string, error) {
	relationships, err := parseRoutingRelationships([]byte(relsXML))
	if err != nil {
		return "", fmt.Errorf("xlsxpatch: parse rels: %w", err)
	}
	var match *routingRelationship
	for index := range relationships {
		if relationships[index].id == relID {
			match = &relationships[index]
			break
		}
	}
	if match == nil {
		return "", fmt.Errorf("xlsxpatch: relationship %q not found", relID)
	}
	switch match.targetMode {
	case "", "Internal":
	case "External":
		return "", fmt.Errorf("xlsxpatch: relationship %q has an external target", relID)
	default:
		return "", fmt.Errorf("xlsxpatch: relationship %q has unsupported target mode %q", relID, match.targetMode)
	}
	if len(allowedTypes) > 0 {
		allowed := false
		for _, candidate := range allowedTypes {
			if match.relType == candidate {
				allowed = true
				break
			}
		}
		if !allowed {
			return "", fmt.Errorf("xlsxpatch: relationship %q has unsupported type %q", relID, match.relType)
		}
	}
	resolved, err := resolveRelPath(baseDir, match.target)
	if err != nil {
		return "", fmt.Errorf("xlsxpatch: relationship %q target %q: %w", relID, match.target, err)
	}
	return resolved, nil
}

// resolveRelPath turns an internal OPC relationship target into the exact
// canonical ZIP part name. Relative targets start at the relationship owner's
// directory, root-relative targets start at the package root, and traversal
// above that root is refused rather than silently redirected to another part.
func resolveRelPath(baseDir, target string) (string, error) {
	if target == "" {
		return "", fmt.Errorf("empty internal relationship target")
	}
	reference, err := url.Parse(target)
	if err != nil {
		return "", fmt.Errorf("invalid internal relationship target: %w", err)
	}
	if reference.IsAbs() || reference.Host != "" || reference.Opaque != "" {
		return "", fmt.Errorf("target is not an internal package path")
	}
	if reference.RawQuery != "" || reference.ForceQuery {
		return "", fmt.Errorf("internal relationship target must not contain a query")
	}
	if reference.Path == "" {
		return "", fmt.Errorf("internal relationship target has no resolvable part path")
	}
	lowerEscapedPath := strings.ToLower(reference.EscapedPath())
	if strings.Contains(lowerEscapedPath, "%2f") || strings.Contains(lowerEscapedPath, "%5c") {
		return "", fmt.Errorf("internal relationship target contains an encoded path separator")
	}
	if err := validateDecodedPackagePath(reference.Path); err != nil {
		return "", err
	}
	targetPath := strings.TrimPrefix(reference.Path, "/")
	if strings.Contains(targetPath, "//") || strings.HasSuffix(targetPath, "/") {
		return "", fmt.Errorf("internal relationship target contains an empty path segment")
	}
	if baseDir != "" {
		if strings.HasPrefix(baseDir, "/") || strings.Contains(baseDir, "//") || strings.HasSuffix(baseDir, "/") {
			return "", fmt.Errorf("invalid relationship base directory %q: non-canonical path separators", baseDir)
		}
		if err := validateDecodedPackagePath(baseDir); err != nil {
			return "", fmt.Errorf("invalid relationship base directory %q: %w", baseDir, err)
		}
	}

	parts := make([]string, 0, 8)
	if !strings.HasPrefix(reference.Path, "/") {
		var baseErr error
		parts, baseErr = appendCanonicalPackageSegments(parts, baseDir)
		if baseErr != nil {
			return "", fmt.Errorf("invalid relationship base directory %q: %w", baseDir, baseErr)
		}
	}
	parts, err = appendCanonicalPackageSegments(parts, reference.Path)
	if err != nil {
		return "", err
	}
	if len(parts) == 0 {
		return "", fmt.Errorf("internal relationship target resolves to the package root")
	}
	return strings.Join(parts, "/"), nil
}

func validateDecodedPackagePath(value string) error {
	for _, character := range value {
		if character == '\\' {
			return fmt.Errorf("internal relationship target contains a backslash")
		}
		if character < 0x20 || character == 0x7f {
			return fmt.Errorf("internal relationship target contains control character U+%04X", character)
		}
	}
	return nil
}

func appendCanonicalPackageSegments(parts []string, value string) ([]string, error) {
	for _, segment := range strings.Split(value, "/") {
		switch segment {
		case "", ".":
			continue
		case "..":
			if len(parts) == 0 {
				return nil, fmt.Errorf("internal relationship target traverses above the package root")
			}
			parts = parts[:len(parts)-1]
		default:
			if strings.HasSuffix(segment, ".") {
				return nil, fmt.Errorf("internal relationship path segment %q ends with a dot", segment)
			}
			parts = append(parts, segment)
		}
	}
	return parts, nil
}

// partRelTargetFrom builds the relative Target from a part's rels to another
// part (e.g. drawing → chart is "../charts/chart2.xml").
func partRelTargetFrom(fromPart, toPart string) string {
	fromDir := fromPart[:strings.LastIndex(fromPart, "/")]
	fromSegs := strings.Split(fromDir, "/")
	toSegs := strings.Split(toPart, "/")
	common := 0
	for common < len(fromSegs) && common < len(toSegs)-1 && fromSegs[common] == toSegs[common] {
		common++
	}
	var b strings.Builder
	for i := common; i < len(fromSegs); i++ {
		b.WriteString("../")
	}
	b.WriteString(strings.Join(toSegs[common:], "/"))
	return b.String()
}

func relsPartFor(part string) string {
	idx := strings.LastIndex(part, "/")
	if idx < 0 {
		return "_rels/" + part + ".rels"
	}
	return part[:idx] + "/_rels/" + part[idx+1:] + ".rels"
}

func nextFreePart(zr *zip.Reader, prefix, suffix string) string {
	taken := map[string]bool{}
	for _, f := range zr.File {
		taken[f.Name] = true
	}
	for i := 1; ; i++ {
		name := fmt.Sprintf("%s%d%s", prefix, i, suffix)
		if !taken[name] {
			return name
		}
	}
}

var relIDRe = regexp.MustCompile(`Id="rId(\d+)"`)

func nextFreeRelID(relsXML string) string {
	max := 0
	for _, m := range relIDRe.FindAllStringSubmatch(relsXML, -1) {
		n := 0
		fmt.Sscanf(m[1], "%d", &n) //nolint:errcheck
		if n > max {
			max = n
		}
	}
	return fmt.Sprintf("rId%d", max+1)
}

func relationshipXML(id, relType, target string) string {
	return fmt.Sprintf(`<Relationship Id=%q Type=%q Target=%q/>`, id, relType, esc(target))
}

func relsDoc(inner string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n" +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` + inner + `</Relationships>`
}

func appendRelationship(relsXML, id, relType, target string) (string, error) {
	idx := strings.LastIndex(relsXML, "</Relationships>")
	if idx < 0 {
		return "", fmt.Errorf("xlsxpatch: malformed rels part")
	}
	return relsXML[:idx] + relationshipXML(id, relType, target) + relsXML[idx:], nil
}

var drawingElRe = regexp.MustCompile(`<drawing[^>]*r:id="([^"]+)"`)

func drawingRelIDInSheet(sheetXML string) string {
	if m := drawingElRe.FindStringSubmatch(sheetXML); m != nil {
		return m[1]
	}
	return ""
}

// insertDrawingElement places <drawing r:id=.../> at a schema-valid position:
// before the first element that must follow it, else before </worksheet>.
func insertDrawingElement(sheetXML, relID string) (string, error) {
	// Local xmlns:r for the same reason as workbookWithPivotCache: never
	// trust where (or whether) the document bound the prefix.
	el := fmt.Sprintf(`<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id=%q/>`, relID)
	insertAt := strings.LastIndex(sheetXML, "</worksheet>")
	if insertAt < 0 {
		return "", fmt.Errorf("xlsxpatch: malformed worksheet part")
	}
	for _, follower := range []string{"<legacyDrawing", "<legacyDrawingHF", "<picture", "<oleObjects", "<controls", "<webPublishItems", "<tableParts", "<extLst"} {
		if i := strings.Index(sheetXML, follower); i >= 0 && i < insertAt {
			insertAt = i
		}
	}
	return sheetXML[:insertAt] + el + sheetXML[insertAt:], nil
}

func appendAnchorToDrawing(drawingXML, anchor string) (string, error) {
	idx := strings.LastIndex(drawingXML, "</xdr:wsDr>")
	if idx < 0 {
		return "", fmt.Errorf("xlsxpatch: unsupported drawing part layout (no xdr:wsDr close)")
	}
	return drawingXML[:idx] + anchor + drawingXML[idx:], nil
}

func contentTypesWith(ct string, overrides map[string]string) (string, error) {
	idx := strings.LastIndex(ct, "</Types>")
	if idx < 0 {
		return "", fmt.Errorf("xlsxpatch: malformed [Content_Types].xml")
	}
	var b strings.Builder
	parts := make([]string, 0, len(overrides))
	for part := range overrides {
		parts = append(parts, part)
	}
	sort.Strings(parts)
	for _, part := range parts {
		typ := overrides[part]
		if strings.Contains(ct, `PartName="`+part+`"`) {
			continue
		}
		fmt.Fprintf(&b, `<Override PartName=%q ContentType=%q/>`, part, typ)
	}
	return ct[:idx] + b.String() + ct[idx:], nil
}

// ---- XML generation ----

func anchorXMLWithObjectID(spec ChartWriteSpec, chartRelID string, objectID uint32) string {
	a := spec.Anchor
	return fmt.Sprintf(`<xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`+
		`<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>`+
		`<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`+
		`<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="%d" name=%q/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`+
		`<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>`+
		`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">`+
		`<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id=%q/>`+
		`</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`,
		a.FromCol, a.FromRow, a.ToCol, a.ToRow, objectID, esc(chartTitleOr(spec, "InjOffice chart")), chartRelID)
}

func chartTitleOr(spec ChartWriteSpec, fallback string) string {
	if spec.Title != "" {
		return spec.Title
	}
	return fallback
}

func drawingXMLDoc(anchors string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n" +
		`<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
		anchors + `</xdr:wsDr>`
}

func serTxXML(s WriteSeries, idx int) string {
	name := s.Name
	if name == "" {
		name = fmt.Sprintf("Series %d", idx+1)
	}
	if s.NameRef != "" {
		return fmt.Sprintf(`<c:tx><c:strRef><c:f>%s</c:f>%s</c:strRef></c:tx>`, esc(s.NameRef), s.nameCache)
	}
	return fmt.Sprintf(`<c:tx><c:v>%s</c:v></c:tx>`, esc(name))
}

func serXML(s WriteSeries, idx int, scatter bool) string {
	var b strings.Builder
	b.WriteString(`<c:ser>`)
	fmt.Fprintf(&b, `<c:idx val="%d"/><c:order val="%d"/>`, idx, idx)
	b.WriteString(serTxXML(s, idx))
	if scatter {
		if s.CategoriesRef != "" {
			fmt.Fprintf(&b, `<c:xVal><c:numRef><c:f>%s</c:f>%s</c:numRef></c:xVal>`, esc(s.CategoriesRef), s.categoryCache)
		}
		fmt.Fprintf(&b, `<c:yVal><c:numRef><c:f>%s</c:f>%s</c:numRef></c:yVal>`, esc(s.ValuesRef), s.valueCache)
	} else {
		if s.CategoriesRef != "" {
			fmt.Fprintf(&b, `<c:cat><c:strRef><c:f>%s</c:f>%s</c:strRef></c:cat>`, esc(s.CategoriesRef), s.categoryCache)
		}
		fmt.Fprintf(&b, `<c:val><c:numRef><c:f>%s</c:f>%s</c:numRef></c:val>`, esc(s.ValuesRef), s.valueCache)
	}
	b.WriteString(`</c:ser>`)
	return b.String()
}

func plotXML(spec ChartWriteSpec) string {
	var sers strings.Builder
	scatter := spec.Type == "scatter"
	for i, s := range spec.Series {
		sers.WriteString(serXML(s, i, scatter))
	}
	axes := `<c:axId val="111111111"/><c:axId val="222222222"/>`
	switch spec.Type {
	case "column", "bar":
		dir := "col"
		if spec.Type == "bar" {
			dir = "bar"
		}
		return `<c:barChart><c:barDir val="` + dir + `"/><c:grouping val="clustered"/><c:varyColors val="0"/>` + sers.String() + axes + `</c:barChart>` + catValAxesXML()
	case "line":
		return `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>` + sers.String() + `<c:marker val="1"/>` + axes + `</c:lineChart>` + catValAxesXML()
	case "area":
		return `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>` + sers.String() + axes + `</c:areaChart>` + catValAxesXML()
	case "scatter":
		return `<c:scatterChart><c:scatterStyle val="marker"/><c:varyColors val="0"/>` + sers.String() + axes + `</c:scatterChart>` + valValAxesXML()
	case "pie":
		return `<c:pieChart><c:varyColors val="1"/>` + sers.String() + `</c:pieChart>`
	case "doughnut":
		return `<c:doughnutChart><c:varyColors val="1"/>` + sers.String() + `<c:holeSize val="50"/></c:doughnutChart>`
	}
	return ""
}

func catValAxesXML() string {
	return `<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>` +
		`<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>`
}

func valValAxesXML() string {
	return `<c:valAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:valAx>` +
		`<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>`
}

func chartSpaceXML(spec ChartWriteSpec) string {
	title := ""
	if spec.Title != "" {
		title = `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>` + esc(spec.Title) + `</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
	}
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n" +
		`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<c:chart>` + title + `<c:plotArea><c:layout/>` + plotXML(spec) + `</c:plotArea>` +
		`<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>` +
		`<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`
}

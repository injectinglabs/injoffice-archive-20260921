package docxpatch

import (
	"encoding/xml"
	"fmt"
	"math"
	"regexp"
	"strconv"
	"strings"
	"unicode/utf16"
)

// Read-only chart sidecar for the explicitly labeled approximate page preview.
//
// The strict extractor keeps refusing DrawingML charts (PICTURE_GRAPHIC_REQUIRED
// / UNMODELED_DRAWING). This inspection joins those retained diagnostics back to
// their w:drawing (same discipline as InspectNativeApproximateDrawingShapesV1:
// package digest, owning body paragraph, diagnostic anchors inside the drawing
// run) and describes a bounded chart model read only from cached values in the
// chart part: clustered bar/column charts with c:strCache / c:numCache (or
// literal) points, series names, title text, legend position, axis presence and
// paint. Nothing is recalculated: no workbook cells are read, no scale is
// derived here, and every colour that comes from the package theme or an Office
// default is disclosed as an approximation. Every other chart type, grouping or
// missing cache refuses per chart with a declared reason. The reader mirrors the
// cache handling of the PPTX literal chart extractors (go/pptxpatch/
// native_chart_bar.go, native_chart_literal_series.go); docxpatch and pptxpatch
// are independent Go modules, so the minimal subset is restated here rather than
// imported. Source bytes, native extraction, editing authority and pagination
// are unchanged.
const NativeApproximateDrawingChartsProtocol = "injoffice.docx.approximate-drawing-charts"
const NativeApproximateDrawingChartPolicy = "docx.approximate-drawing-chart-preview-v1"
const nativeApproximateDrawingChartLimit = 16
const nativeApproximateChartMaxSeries = 16
const nativeApproximateChartMaxCategories = 256
const nativeApproximateChartMaxTextUnits = 32768
const nativeApproximateChartMaxTitleUnits = 1024
const nativeChartNSTransitional = "http://schemas.openxmlformats.org/drawingml/2006/chart"
const nativeChartNSStrict = "http://purl.oclc.org/ooxml/drawingml/chart"
const nativeChartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"

var nativeApproximateChartDecimal = regexp.MustCompile(`^-?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][-+]?[0-9]{1,3})?$`)

type NativeApproximateChartFontV1 struct {
	Family          string `json:"family"`
	SizeHundredthPt int64  `json:"size_hundredth_pt"`
	RGB             string `json:"rgb"`
	Bold            bool   `json:"bold"`
	Italic          bool   `json:"italic"`
}

type NativeApproximateChartSeriesV1 struct {
	Index   int64                         `json:"index"`
	Order   int64                         `json:"order"`
	Title   *string                       `json:"title,omitempty"`
	Values  []string                      `json:"values"`
	FillRGB string                        `json:"fill_rgb"`
	Line    *NativeApproximateShapeLineV1 `json:"line,omitempty"`
}

type NativeApproximateChartAxisV1 struct {
	Deleted        bool                          `json:"deleted"`
	Orientation    string                        `json:"orientation"`
	Line           *NativeApproximateShapeLineV1 `json:"line,omitempty"`
	Labels         *NativeApproximateChartFontV1 `json:"labels,omitempty"`
	MajorGridlines *NativeApproximateShapeLineV1 `json:"major_gridlines,omitempty"`
	NumberFormat   string                        `json:"number_format"`
	Min            *string                       `json:"min,omitempty"`
	Max            *string                       `json:"max,omitempty"`
	MajorUnit      *string                       `json:"major_unit,omitempty"`
}

type NativeApproximateChartTitleV1 struct {
	Text    *string                      `json:"text,omitempty"`
	Font    NativeApproximateChartFontV1 `json:"font"`
	Overlay bool                         `json:"overlay"`
}

type NativeApproximateChartLegendV1 struct {
	Position string                       `json:"position"`
	Font     NativeApproximateChartFontV1 `json:"font"`
	Overlay  bool                         `json:"overlay"`
}

type NativeApproximateChartModelV1 struct {
	Kind            string                           `json:"kind"`
	BarDirection    string                           `json:"bar_direction"`
	Grouping        string                           `json:"grouping"`
	GapWidthPercent int64                            `json:"gap_width_percent"`
	OverlapPercent  int64                            `json:"overlap_percent"`
	Categories      []string                         `json:"categories"`
	Series          []NativeApproximateChartSeriesV1 `json:"series"`
	CategoryAxis    NativeApproximateChartAxisV1     `json:"category_axis"`
	ValueAxis       NativeApproximateChartAxisV1     `json:"value_axis"`
	Title           *NativeApproximateChartTitleV1   `json:"title,omitempty"`
	Legend          *NativeApproximateChartLegendV1  `json:"legend,omitempty"`
	AreaFillRGB     *string                          `json:"area_fill_rgb,omitempty"`
	AreaLine        *NativeApproximateShapeLineV1    `json:"area_line,omitempty"`
	PlotFillRGB     *string                          `json:"plot_fill_rgb,omitempty"`
	PlotLine        *NativeApproximateShapeLineV1    `json:"plot_line,omitempty"`
}

type NativeApproximateDrawingChartV1 struct {
	ID              string                         `json:"id"`
	ParagraphID     string                         `json:"paragraph_id"`
	DiagnosticIDs   []string                       `json:"diagnostic_ids"`
	Anchor          NativeSourceAnchorV1           `json:"anchor"`
	RunAnchor       NativeSourceAnchorV1           `json:"run_anchor"`
	Status          string                         `json:"status"`
	Reason          string                         `json:"reason,omitempty"`
	Placement       string                         `json:"placement,omitempty"`
	WidthEMU        int64                          `json:"width_emu"`
	HeightEMU       int64                          `json:"height_emu"`
	PageAnchor      *NativeTextboxPageAnchorV1     `json:"page_anchor,omitempty"`
	Wrap            string                         `json:"wrap,omitempty"`
	ChartPart       string                         `json:"chart_part,omitempty"`
	ChartPartSHA256 string                         `json:"chart_part_sha256,omitempty"`
	Chart           *NativeApproximateChartModelV1 `json:"chart,omitempty"`
	Notes           []string                       `json:"notes,omitempty"`
}

type NativeApproximateDrawingChartsV1 struct {
	Protocol      string                            `json:"protocol"`
	Version       int                               `json:"version"`
	Policy        string                            `json:"policy"`
	PackageSHA256 string                            `json:"package_sha256"`
	PartSHA256    string                            `json:"part_sha256"`
	Items         []NativeApproximateDrawingChartV1 `json:"items"`
	OmittedCount  int                               `json:"omitted_count"`
}

type nativeApproximateChartContext struct {
	shapes    *nativeApproximateShapeContext
	relNS     string
	relBase   string
	c         string
	minorFont string
	majorFont string
	seen      map[string]int
}

// InspectNativeApproximateDrawingChartsV1 returns nil when the body has no
// refused c:chart drawings. Callers must treat every item as approximate
// evidence and re-validate the joins against the same-bytes document.
func InspectNativeApproximateDrawingChartsV1(data []byte) (*NativeApproximateDrawingChartsV1, error) {
	if len(data) == 0 || len(data) > NativeDOCXMaxPackageBytes {
		return nil, fmt.Errorf("approximate drawing charts package size must be 1..%d bytes", NativeDOCXMaxPackageBytes)
	}
	data = append([]byte(nil), data...)
	resolver, err := newNativeLayoutResolver(data, NativeExtractionOptions{})
	if err != nil {
		return nil, err
	}
	doc := resolver.doc
	ns, main := resolver.wordNS, resolver.mainPart
	wp, a, relNS, relBase, c := wordDrawingTransitional, drawingMLTransitional, relNSTransitional, relBaseTransitional, nativeChartNSTransitional
	if ns == wordMLStrict {
		wp, a, relNS, relBase, c = wordDrawingStrict, drawingMLStrict, relNSStrict, relBaseStrict, nativeChartNSStrict
	}
	raw := resolver.pkg.files[main]
	shapes := &nativeApproximateShapeContext{resolver: resolver, ns: ns, wp: wp, a: a, main: main, raw: raw, theme: nativeApproximateThemeFromPackage(resolver.pkg, main, ns), shapeSeen: map[string]int{}}
	context := &nativeApproximateChartContext{shapes: shapes, relNS: relNS, relBase: relBase, c: c, seen: map[string]int{}}
	context.minorFont, context.majorFont = nativeApproximateThemeFonts(resolver.pkg, main, ns)
	nodes := map[string]*nativeXMLNode{}
	var visit func(*nativeXMLNode)
	visit = func(n *nativeXMLNode) {
		nodes[n.Path] = n
		for _, child := range n.Children {
			visit(child)
		}
	}
	visit(resolver.mainRoot)
	diagnostics := map[string][]NativeUnsupportedCapabilityV1{}
	for _, d := range doc.Unsupported {
		if (d.Capability == "drawings" || d.Capability == "runs") && d.Anchor != nil && d.Anchor.PartName == main {
			diagnostics[d.ScopeID] = append(diagnostics[d.ScopeID], d)
		}
	}
	out := &NativeApproximateDrawingChartsV1{Protocol: NativeApproximateDrawingChartsProtocol, Version: 1, Policy: NativeApproximateDrawingChartPolicy, PackageSHA256: doc.Source.PackageSHA256, PartSHA256: nativeSHA(raw), Items: []NativeApproximateDrawingChartV1{}}
	for _, block := range doc.Body.Blocks {
		p := block.Paragraph
		if p == nil || len(diagnostics[p.ID]) == 0 {
			continue
		}
		owner := nodes[p.Anchor.Path]
		if owner == nil || owner.Name != (xml.Name{Space: ns, Local: "p"}) {
			continue
		}
		for _, run := range owner.Children {
			if run.Name != (xml.Name{Space: ns, Local: "r"}) {
				continue
			}
			for _, child := range run.Children {
				drawing := shapes.drawingNode(child)
				if drawing == nil || !context.isChartDrawing(drawing) {
					// Pictures, wps shapes and groups belong to their own sidecars.
					continue
				}
				joined := []string{}
				for _, d := range diagnostics[p.ID] {
					if d.Anchor.Path != child.Path && !strings.HasPrefix(d.Anchor.Path, child.Path+"/") {
						continue
					}
					n := nodes[d.Anchor.Path]
					if n == nil || nativeSHA(raw[n.Start:n.End]) != d.Anchor.XMLSHA256 {
						continue
					}
					joined = append(joined, d.ID)
				}
				if len(joined) == 0 {
					continue
				}
				if len(out.Items) >= nativeApproximateDrawingChartLimit {
					out.OmittedCount++
					continue
				}
				out.Items = append(out.Items, context.describe(p.ID, joined, run, drawing))
			}
		}
	}
	if len(out.Items) == 0 && out.OmittedCount == 0 {
		return nil, nil
	}
	return out, nil
}

// isChartDrawing reports whether the drawing's graphicData carries the
// DrawingML chart URI. Anything else is left to other sidecars.
func (context *nativeApproximateChartContext) isChartDrawing(drawing *nativeXMLNode) bool {
	return nativeApproximateGraphicURI(drawing, context.shapes.wp, context.shapes.a) == context.c
}

// nativeApproximateGraphicURI returns the uri of the first graphicData under
// the drawing's inline/anchor container, or "" when there is none. The shape
// and chart sidecars use it to hand each drawing to exactly one describer.
func nativeApproximateGraphicURI(drawing *nativeXMLNode, wp, a string) string {
	for _, container := range drawing.Children {
		if container.Name.Space != wp || (container.Name.Local != "inline" && container.Name.Local != "anchor") {
			continue
		}
		graphic := firstDirectNativeChild(container, a, "graphic")
		if graphic == nil {
			continue
		}
		graphicData := firstDirectNativeChild(graphic, a, "graphicData")
		if graphicData == nil {
			continue
		}
		uri, _ := nativeUnqualifiedAttr(graphicData, "uri")
		return uri
	}
	return ""
}

func (context *nativeApproximateChartContext) describe(paragraphID string, diagnosticIDs []string, run, drawing *nativeXMLNode) NativeApproximateDrawingChartV1 {
	shapes := context.shapes
	digest := strings.TrimPrefix(nativeSHA(shapes.raw[drawing.Start:drawing.End]), "sha256:")[:16]
	context.seen[digest]++
	item := NativeApproximateDrawingChartV1{ID: fmt.Sprintf("approximate-drawing-chart:%s:%d", digest, context.seen[digest]), ParagraphID: paragraphID, DiagnosticIDs: diagnosticIDs, Anchor: shapes.anchor(drawing), RunAnchor: shapes.anchor(run), Status: "omitted"}
	omit := func(reason string) NativeApproximateDrawingChartV1 {
		item.Reason = reason
		item.Chart = nil
		return item
	}
	// Only the drawing may share its w:r with w:rPr; a run that also carries
	// text would leave modeled runs nested inside the chart's run anchor.
	for _, sibling := range run.Children {
		if sibling.Name == (xml.Name{Space: shapes.ns, Local: "rPr"}) || sibling == drawing || (sibling.Name == (xml.Name{Space: nativeMarkupCompatibilityNS, Local: "AlternateContent"}) && shapes.drawingNode(sibling) == drawing) {
			continue
		}
		return omit("shared-run")
	}
	wp, a := shapes.wp, shapes.a
	var container *nativeXMLNode
	for _, c := range drawing.Children {
		if c.Name.Space == wp && (c.Name.Local == "inline" || c.Name.Local == "anchor") {
			if container != nil {
				return omit("ambiguous-drawing-container")
			}
			container = c
		}
	}
	if container == nil {
		return omit("missing-drawing-container")
	}
	extent := firstDirectNativeChild(container, wp, "extent")
	if extent == nil {
		return omit("missing-extent")
	}
	width, okW := nativePositiveInt64Attr(extent, "", "cx")
	height, okH := nativePositiveInt64Attr(extent, "", "cy")
	if !okW || !okH || width > 127000000 || height > 127000000 {
		return omit("invalid-extent")
	}
	item.WidthEMU, item.HeightEMU = width, height
	graphic := firstDirectNativeChild(container, a, "graphic")
	if graphic == nil {
		return omit("missing-graphic")
	}
	graphicData := firstDirectNativeChild(graphic, a, "graphicData")
	if graphicData == nil {
		return omit("missing-graphic-data")
	}
	charts := directNativeChildren(graphicData, context.c, "chart")
	if len(charts) != 1 || len(graphicData.Children) != 1 {
		return omit("missing-chart-reference")
	}
	relID, ok := nativeAttr(charts[0], context.relNS, "id")
	if !ok || relID == "" || len(relID) > 256 {
		return omit("missing-chart-reference")
	}
	var rel *nativeRelationship
	for i := range shapes.resolver.pkg.rels[shapes.main] {
		candidate := &shapes.resolver.pkg.rels[shapes.main][i]
		if candidate.ID != relID {
			continue
		}
		if rel != nil {
			return omit("ambiguous-chart-relationship")
		}
		rel = candidate
	}
	if rel == nil || rel.Type != context.relBase+"chart" {
		return omit("missing-chart-relationship")
	}
	if rel.External || rel.PartName == "" {
		return omit("external-chart-relationship")
	}
	if !nativeASCIIEqual(shapes.resolver.pkg.contentTypes[rel.PartName], nativeChartContentType) {
		return omit("chart-content-type-mismatch")
	}
	payload := shapes.resolver.pkg.files[rel.PartName]
	if len(payload) == 0 {
		return omit("missing-chart-part")
	}
	root, err := parseNativeXML(rel.PartName, payload)
	if err != nil {
		return omit("invalid-chart-xml")
	}
	if root.Name != (xml.Name{Space: context.c, Local: "chartSpace"}) {
		return omit("not-chart-space")
	}
	item.ChartPart, item.ChartPartSHA256 = rel.PartName, nativeSHA(payload)
	if container.Name.Local == "inline" {
		item.Placement = "inline"
	} else {
		item.Placement = "anchored"
		pageAnchor, wrap, reason := shapes.pageAnchor(container)
		if reason != "" {
			return omit(reason)
		}
		item.PageAnchor, item.Wrap = pageAnchor, wrap
		if wrap != "none" {
			item.Notes = append(item.Notes, "body text wrapping around the chart is not applied")
		}
	}
	model, notes, reason := context.chartModel(root)
	if reason != "" {
		return omit(reason)
	}
	item.Chart = model
	item.Notes = append(item.Notes, notes...)
	item.Status = "supported"
	return item
}

// chartModel reads a clustered bar/column chart from the chartSpace root.
// Unknown siblings are tolerated (the preview is approximate), but every value
// that is painted must be cached in the part.
func (context *nativeApproximateChartContext) chartModel(root *nativeXMLNode) (*NativeApproximateChartModelV1, []string, string) {
	c := context.c
	notes := []string{}
	chart := firstDirectNativeChild(root, c, "chart")
	if chart == nil {
		return nil, nil, "missing-chart"
	}
	plot := firstDirectNativeChild(chart, c, "plotArea")
	if plot == nil {
		return nil, nil, "missing-plot-area"
	}
	defaultFont := context.font(firstDirectNativeChild(root, c, "txPr"), NativeApproximateChartFontV1{Family: context.minorFont, SizeHundredthPt: 1000, RGB: "000000"}, &notes)
	var plotChart *nativeXMLNode
	for _, child := range plot.Children {
		if child.Name.Space != c || !strings.HasSuffix(child.Name.Local, "Chart") {
			continue
		}
		if plotChart != nil {
			return nil, nil, "combination-chart"
		}
		plotChart = child
	}
	if plotChart == nil {
		return nil, nil, "missing-chart-type"
	}
	if plotChart.Name.Local != "barChart" {
		return nil, nil, "unsupported-chart-type:" + plotChart.Name.Local
	}
	if layout := firstDirectNativeChild(plot, c, "layout"); layout != nil && len(layout.Children) > 0 {
		notes = append(notes, "manual plot area layout is not applied")
	}
	model := &NativeApproximateChartModelV1{Kind: "bar", BarDirection: "column", Grouping: "clustered", GapWidthPercent: 150, Categories: []string{}, Series: []NativeApproximateChartSeriesV1{}}
	if direction, ok := nativeApproximateChartVal(plotChart, c, "barDir"); ok {
		switch direction {
		case "col":
			model.BarDirection = "column"
		case "bar":
			model.BarDirection = "bar"
		default:
			return nil, nil, "unsupported-bar-direction:" + direction
		}
	}
	if grouping, ok := nativeApproximateChartVal(plotChart, c, "grouping"); ok && grouping != "clustered" {
		return nil, nil, "unsupported-grouping:" + grouping
	}
	varyColors := nativeApproximateChartBool(plotChart, c, "varyColors")
	if gap, ok := nativeApproximateChartVal(plotChart, c, "gapWidth"); ok {
		value, err := strconv.ParseInt(strings.TrimSuffix(gap, "%"), 10, 64)
		if err != nil || value < 0 || value > 500 {
			return nil, nil, "invalid-gap-width"
		}
		model.GapWidthPercent = value
	}
	if overlap, ok := nativeApproximateChartVal(plotChart, c, "overlap"); ok {
		value, err := strconv.ParseInt(strings.TrimSuffix(overlap, "%"), 10, 64)
		if err != nil || value < -100 || value > 100 {
			return nil, nil, "invalid-overlap"
		}
		model.OverlapPercent = value
	}
	if nativeApproximateChartLabelsShown(firstDirectNativeChild(plotChart, c, "dLbls"), c) {
		notes = append(notes, "data labels are omitted")
	}
	indices, orders := map[int64]bool{}, map[int64]bool{}
	units := 0
	categoriesSeen := false
	for _, ser := range directNativeChildren(plotChart, c, "ser") {
		if len(model.Series) >= nativeApproximateChartMaxSeries {
			return nil, nil, "too-many-series"
		}
		series, categories, seriesNotes, reason := context.series(ser, len(model.Series), varyColors, &units)
		if reason != "" {
			return nil, nil, reason
		}
		if indices[series.Index] || orders[series.Order] {
			return nil, nil, "duplicate-series-order"
		}
		indices[series.Index], orders[series.Order] = true, true
		notes = append(notes, seriesNotes...)
		if categories != nil {
			if !categoriesSeen {
				model.Categories, categoriesSeen = categories, true
			} else if len(categories) != len(model.Categories) {
				notes = append(notes, "series category counts differ; the first series' categories are used")
			}
		}
		model.Series = append(model.Series, *series)
	}
	if len(model.Series) == 0 {
		return nil, nil, "no-series"
	}
	count := len(model.Categories)
	for _, series := range model.Series {
		if len(series.Values) > count {
			count = len(series.Values)
		}
	}
	if count == 0 {
		return nil, nil, "no-cached-points"
	}
	if count > nativeApproximateChartMaxCategories {
		return nil, nil, "too-many-categories"
	}
	for len(model.Categories) < count {
		// Office numbers categories without labels; the numbers are not cached
		// so the preview leaves them blank rather than inventing them.
		model.Categories = append(model.Categories, "")
	}
	if !categoriesSeen {
		notes = append(notes, "category labels are not cached; left blank")
	}
	for i := range model.Series {
		for len(model.Series[i].Values) < count {
			model.Series[i].Values = append(model.Series[i].Values, "")
		}
	}
	axisIDs := []string{}
	for _, node := range directNativeChildren(plotChart, c, "axId") {
		if id, ok := nativeUnqualifiedAttr(node, "val"); ok {
			axisIDs = append(axisIDs, id)
		}
	}
	var category, value *NativeApproximateChartAxisV1
	for _, child := range plot.Children {
		if child.Name.Space != c {
			continue
		}
		switch child.Name.Local {
		case "catAx", "dateAx":
			if !nativeApproximateChartAxisListed(child, c, axisIDs) {
				continue
			}
			if category != nil {
				return nil, nil, "ambiguous-category-axis"
			}
			axis, axisNotes, ok := context.axis(child, defaultFont, false)
			if !ok {
				return nil, nil, "unsupported-axis-paint"
			}
			if child.Name.Local == "dateAx" {
				axisNotes = append(axisNotes, "date axis is treated as a category axis")
			}
			notes = append(notes, axisNotes...)
			category = axis
		case "valAx":
			if !nativeApproximateChartAxisListed(child, c, axisIDs) {
				continue
			}
			if value != nil {
				return nil, nil, "ambiguous-value-axis"
			}
			axis, axisNotes, ok := context.axis(child, defaultFont, true)
			if !ok {
				return nil, nil, "unsupported-axis-paint"
			}
			notes = append(notes, axisNotes...)
			value = axis
		case "serAx":
			return nil, nil, "unsupported-axis:serAx"
		}
	}
	if category == nil || value == nil {
		return nil, nil, "missing-axis"
	}
	// An authored scale must be a finite, increasing range with a positive
	// unit; anything else refuses this chart only (the preview never derives
	// around authored bounds).
	if value.Min != nil && value.Max != nil {
		low, errLow := strconv.ParseFloat(*value.Min, 64)
		high, errHigh := strconv.ParseFloat(*value.Max, 64)
		if errLow != nil || errHigh != nil || !(low < high) || math.IsInf(high-low, 0) {
			return nil, nil, "invalid-axis-scale"
		}
	}
	if value.MajorUnit != nil {
		unit, err := strconv.ParseFloat(*value.MajorUnit, 64)
		if err != nil || !(unit > 0) || math.IsInf(unit, 0) {
			return nil, nil, "invalid-axis-scale"
		}
	}
	model.CategoryAxis, model.ValueAxis = *category, *value
	if value.Min == nil || value.Max == nil {
		notes = append(notes, "value axis scale is not authored; the preview derives a host scale from the cached values")
	}
	if value.MajorUnit == nil {
		notes = append(notes, "value axis major unit is not authored; the preview derives a host unit")
	}
	if plotPr := firstDirectNativeChild(plot, c, "spPr"); plotPr != nil {
		fill, fillNotes, ok := context.shapes.shapeFill(plotPr, nil)
		if !ok {
			return nil, nil, "unsupported-plot-fill"
		}
		line, lineNotes, ok := context.shapes.shapeLine(plotPr, nil)
		if !ok {
			return nil, nil, "unsupported-plot-outline"
		}
		model.PlotFillRGB, model.PlotLine = fill.RGB, line
		notes = append(append(notes, fillNotes...), lineNotes...)
	}
	if title := firstDirectNativeChild(chart, c, "title"); title != nil && !nativeApproximateChartBool(chart, c, "autoTitleDeleted") {
		model.Title = context.title(title, defaultFont, &notes)
		if model.Title == nil {
			return nil, nil, "unsupported-title"
		}
	}
	if legend := firstDirectNativeChild(chart, c, "legend"); legend != nil {
		position := "r"
		if value, ok := nativeApproximateChartVal(legend, c, "legendPos"); ok {
			switch value {
			case "b", "t", "l", "r", "tr":
				position = value
			default:
				return nil, nil, "unsupported-legend-position:" + value
			}
		}
		model.Legend = &NativeApproximateChartLegendV1{Position: position, Overlay: nativeApproximateChartBool(legend, c, "overlay"), Font: context.font(firstDirectNativeChild(legend, c, "txPr"), defaultFont, &notes)}
		if len(directNativeChildren(legend, c, "legendEntry")) > 0 {
			notes = append(notes, "legend entry overrides are not applied")
		}
	}
	if areaPr := firstDirectNativeChild(root, c, "spPr"); areaPr != nil {
		fill, fillNotes, ok := context.shapes.shapeFill(areaPr, nil)
		if !ok {
			return nil, nil, "unsupported-area-fill"
		}
		line, lineNotes, ok := context.shapes.shapeLine(areaPr, nil)
		if !ok {
			return nil, nil, "unsupported-area-outline"
		}
		model.AreaFillRGB, model.AreaLine = fill.RGB, line
		notes = append(append(notes, fillNotes...), lineNotes...)
	}
	if nativeApproximateChartBool(root, c, "roundedCorners") {
		notes = append(notes, "rounded chart corners are painted square")
	}
	return model, nativeApproximateUniqueNotes(notes), ""
}

func nativeApproximateUniqueNotes(notes []string) []string {
	seen := map[string]bool{}
	out := []string{}
	for _, note := range notes {
		if seen[note] || len(out) >= 32 {
			continue
		}
		seen[note] = true
		out = append(out, note)
	}
	return out
}

func nativeApproximateChartAxisListed(axis *nativeXMLNode, c string, ids []string) bool {
	node := firstDirectNativeChild(axis, c, "axId")
	if node == nil {
		return false
	}
	id, _ := nativeUnqualifiedAttr(node, "val")
	for _, candidate := range ids {
		if candidate == id {
			return true
		}
	}
	return false
}

func nativeApproximateChartVal(parent *nativeXMLNode, c, local string) (string, bool) {
	node := firstDirectNativeChild(parent, c, local)
	if node == nil {
		return "", false
	}
	value, ok := nativeUnqualifiedAttr(node, "val")
	if !ok || len(value) > 64 {
		return "", false
	}
	return strings.TrimSpace(value), true
}

// nativeApproximateChartBool reads a CT_Boolean child; an element without val
// means true per the schema default.
func nativeApproximateChartBool(parent *nativeXMLNode, c, local string) bool {
	node := firstDirectNativeChild(parent, c, local)
	if node == nil {
		return false
	}
	value, ok := nativeUnqualifiedAttr(node, "val")
	return !ok || value == "1" || value == "true"
}

func nativeApproximateChartLabelsShown(dLbls *nativeXMLNode, c string) bool {
	if dLbls == nil {
		return false
	}
	for _, key := range []string{"showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize"} {
		if nativeApproximateChartBool(dLbls, c, key) {
			return true
		}
	}
	return len(directNativeChildren(dLbls, c, "dLbl")) > 0
}

// series reads one c:ser: cached name, cached categories and values, explicit
// or defaulted paint. Nil categories means the series carries none.
func (context *nativeApproximateChartContext) series(ser *nativeXMLNode, ordinal int, varyColors bool, units *int) (*NativeApproximateChartSeriesV1, []string, []string, string) {
	c := context.c
	notes := []string{}
	series := &NativeApproximateChartSeriesV1{Index: int64(ordinal), Order: int64(ordinal), Values: []string{}}
	if raw, ok := nativeApproximateChartVal(ser, c, "idx"); ok {
		value, err := strconv.ParseInt(raw, 10, 64)
		if err != nil || value < 0 || value > 4294967295 {
			return nil, nil, nil, "invalid-series-index"
		}
		series.Index = value
	}
	if raw, ok := nativeApproximateChartVal(ser, c, "order"); ok {
		value, err := strconv.ParseInt(raw, 10, 64)
		if err != nil || value < 0 || value > 4294967295 {
			return nil, nil, nil, "invalid-series-order"
		}
		series.Order = value
	}
	if tx := firstDirectNativeChild(ser, c, "tx"); tx != nil {
		title, ok := context.cachedText(tx, units)
		if !ok {
			return nil, nil, nil, "invalid-series-title"
		}
		if title != nil {
			series.Title = title
		} else {
			notes = append(notes, "series name is not cached; legend entry left blank")
		}
	}
	var categories []string
	if cat := firstDirectNativeChild(ser, c, "cat"); cat != nil {
		points, ok := context.cachedPoints(cat, false, units)
		if !ok {
			return nil, nil, nil, "missing-category-cache"
		}
		categories = points
	}
	val := firstDirectNativeChild(ser, c, "val")
	if val == nil {
		return nil, nil, nil, "missing-values"
	}
	values, ok := context.cachedPoints(val, true, units)
	if !ok {
		return nil, nil, nil, "missing-value-cache"
	}
	series.Values = values
	spPr := firstDirectNativeChild(ser, c, "spPr")
	var fill *string
	if spPr != nil {
		resolved, fillNotes, ok := context.shapes.shapeFill(spPr, nil)
		if !ok {
			return nil, nil, nil, "unsupported-series-fill"
		}
		notes = append(notes, fillNotes...)
		fill = resolved.RGB
		if fill == nil {
			// An explicit noFill paints nothing, and gradient/pattern/picture fills
			// are not approximated: keep the bar white so outlines still show
			// rather than inventing an accent colour.
			for _, child := range spPr.Children {
				if child.Name.Space != context.shapes.a {
					continue
				}
				switch child.Name.Local {
				case "noFill", "gradFill", "pattFill", "blipFill", "grpFill":
					fill = nativeString("FFFFFF")
					notes = append(notes, "series "+child.Name.Local+" is not approximated; bar painted white")
				}
			}
		}
		line, lineNotes, ok := context.shapes.shapeLine(spPr, nil)
		if !ok {
			return nil, nil, nil, "unsupported-series-outline"
		}
		notes = append(notes, lineNotes...)
		series.Line = line
	}
	if fill == nil {
		// Office's default style cycles the theme accents by series index; the
		// chart style/color parts are not evaluated here.
		theme := context.shapes.theme
		if theme == nil {
			return nil, nil, nil, "series-fill-unresolvable"
		}
		accent := theme.colors[fmt.Sprintf("accent%d", series.Index%6+1)]
		if accent == "" {
			return nil, nil, nil, "series-fill-unresolvable"
		}
		fill = nativeString(accent)
		notes = append(notes, "series fill defaulted to the theme accent cycle")
	}
	series.FillRGB = *fill
	if varyColors {
		notes = append(notes, "per-point colour variation is not applied")
	}
	if len(directNativeChildren(ser, c, "dPt")) > 0 {
		notes = append(notes, "data point overrides are not applied")
	}
	if nativeApproximateChartLabelsShown(firstDirectNativeChild(ser, c, "dLbls"), c) {
		notes = append(notes, "data labels are omitted")
	}
	if nativeApproximateChartBool(ser, c, "invertIfNegative") {
		notes = append(notes, "invert-if-negative is not applied")
	}
	return series, categories, notes, ""
}

// cachedText reads a c:tx: a strRef with strCache (one point) or a literal c:v.
// A reference without a cache returns nil, true: nothing is invented.
func (context *nativeApproximateChartContext) cachedText(tx *nativeXMLNode, units *int) (*string, bool) {
	c := context.c
	if v := firstDirectNativeChild(tx, c, "v"); v != nil {
		return context.boundedText(strings.TrimSpace(v.Text), units, nativeApproximateChartMaxTitleUnits)
	}
	ref := firstDirectNativeChild(tx, c, "strRef")
	if ref == nil {
		return nil, true
	}
	cache := firstDirectNativeChild(ref, c, "strCache")
	if cache == nil {
		return nil, true
	}
	points, ok := context.cachedPointList(cache, false, units)
	if !ok {
		return nil, false
	}
	parts := []string{}
	for _, point := range points {
		if point != "" {
			parts = append(parts, point)
		}
	}
	if len(parts) == 0 {
		return nil, true
	}
	return context.boundedText(strings.Join(parts, " "), units, nativeApproximateChartMaxTitleUnits)
}

func (context *nativeApproximateChartContext) boundedText(text string, units *int, max int) (*string, bool) {
	length := len(utf16.Encode([]rune(text)))
	if length > max {
		return nil, false
	}
	*units += length
	if *units > nativeApproximateChartMaxTextUnits {
		return nil, false
	}
	for _, r := range text {
		if r < 0x20 && r != '\t' {
			return nil, false
		}
	}
	return nativeString(text), true
}

// cachedPoints reads c:cat / c:val: numRef/strRef caches or numLit/strLit.
// Multi-level category caches use their leaf level.
func (context *nativeApproximateChartContext) cachedPoints(parent *nativeXMLNode, numeric bool, units *int) ([]string, bool) {
	c := context.c
	for _, child := range parent.Children {
		if child.Name.Space != c {
			continue
		}
		switch child.Name.Local {
		case "numRef":
			cache := firstDirectNativeChild(child, c, "numCache")
			if cache == nil {
				return nil, false
			}
			return context.cachedPointList(cache, true, units)
		case "strRef":
			cache := firstDirectNativeChild(child, c, "strCache")
			if cache == nil {
				return nil, false
			}
			return context.cachedPointList(cache, numeric, units)
		case "numLit":
			return context.cachedPointList(child, true, units)
		case "strLit":
			return context.cachedPointList(child, numeric, units)
		case "multiLvlStrRef":
			cache := firstDirectNativeChild(child, c, "multiLvlStrCache")
			if cache == nil {
				return nil, false
			}
			levels := directNativeChildren(cache, c, "lvl")
			if len(levels) == 0 {
				return nil, false
			}
			count := firstDirectNativeChild(cache, c, "ptCount")
			leaf := &nativeXMLNode{Name: cache.Name, Children: []*nativeXMLNode{}}
			if count != nil {
				leaf.Children = append(leaf.Children, count)
			}
			leaf.Children = append(leaf.Children, levels[0].Children...)
			return context.cachedPointList(leaf, numeric, units)
		}
	}
	return nil, false
}

func (context *nativeApproximateChartContext) cachedPointList(cache *nativeXMLNode, numeric bool, units *int) ([]string, bool) {
	c := context.c
	count := int64(-1)
	if raw, ok := nativeApproximateChartVal(cache, c, "ptCount"); ok {
		value, err := strconv.ParseInt(raw, 10, 64)
		if err != nil || value < 0 || value > nativeApproximateChartMaxCategories {
			return nil, false
		}
		count = value
	}
	points := directNativeChildren(cache, c, "pt")
	if count < 0 {
		count = int64(len(points))
	}
	if int64(len(points)) > count || count > nativeApproximateChartMaxCategories {
		return nil, false
	}
	values := make([]string, count)
	seen := make([]bool, count)
	for _, pt := range points {
		raw, ok := nativeUnqualifiedAttr(pt, "idx")
		if !ok {
			return nil, false
		}
		index, err := strconv.ParseInt(strings.TrimSpace(raw), 10, 64)
		if err != nil || index < 0 || index >= count || seen[index] {
			return nil, false
		}
		v := firstDirectNativeChild(pt, c, "v")
		if v == nil {
			return nil, false
		}
		text := strings.TrimSpace(v.Text)
		if numeric {
			if len(text) > 64 || !nativeApproximateChartDecimal.MatchString(text) {
				return nil, false
			}
			*units += len(text)
			if *units > nativeApproximateChartMaxTextUnits {
				return nil, false
			}
		} else {
			bounded, ok := context.boundedText(text, units, nativeApproximateChartMaxTitleUnits)
			if !ok {
				return nil, false
			}
			text = *bounded
		}
		seen[index] = true
		values[index] = text
	}
	return values, true
}

func (context *nativeApproximateChartContext) axis(node *nativeXMLNode, defaultFont NativeApproximateChartFontV1, value bool) (*NativeApproximateChartAxisV1, []string, bool) {
	c := context.c
	notes := []string{}
	axis := &NativeApproximateChartAxisV1{Orientation: "minMax", NumberFormat: "General"}
	axis.Deleted = nativeApproximateChartBool(node, c, "delete")
	if scaling := firstDirectNativeChild(node, c, "scaling"); scaling != nil {
		if orientation, ok := nativeApproximateChartVal(scaling, c, "orientation"); ok && orientation == "maxMin" {
			axis.Orientation = "maxMin"
		}
		if value {
			for _, key := range []string{"min", "max"} {
				raw, ok := nativeApproximateChartVal(scaling, c, key)
				if !ok {
					continue
				}
				if !nativeApproximateChartDecimal.MatchString(raw) {
					return nil, nil, false
				}
				if key == "min" {
					axis.Min = nativeString(raw)
				} else {
					axis.Max = nativeString(raw)
				}
			}
			if firstDirectNativeChild(scaling, c, "logBase") != nil {
				return nil, nil, false
			}
		}
	}
	if value {
		if raw, ok := nativeApproximateChartVal(node, c, "majorUnit"); ok {
			if !nativeApproximateChartDecimal.MatchString(raw) {
				return nil, nil, false
			}
			axis.MajorUnit = nativeString(raw)
		}
	}
	if format := firstDirectNativeChild(node, c, "numFmt"); format != nil {
		if code, ok := nativeUnqualifiedAttr(format, "formatCode"); ok && len(code) <= 64 {
			axis.NumberFormat = code
		}
	}
	if axis.Deleted {
		return axis, notes, true
	}
	if spPr := firstDirectNativeChild(node, c, "spPr"); spPr != nil {
		line, lineNotes, ok := context.shapes.shapeLine(spPr, nil)
		if !ok {
			return nil, nil, false
		}
		notes = append(notes, lineNotes...)
		axis.Line = line
	} else {
		axis.Line = &NativeApproximateShapeLineV1{RGB: "D9D9D9", WidthEMU: 9525, Dash: "solid"}
		notes = append(notes, "axis line defaulted to the Office chart style (D9D9D9, 0.75pt)")
	}
	if gridlines := firstDirectNativeChild(node, c, "majorGridlines"); gridlines != nil {
		if spPr := firstDirectNativeChild(gridlines, c, "spPr"); spPr != nil {
			line, lineNotes, ok := context.shapes.shapeLine(spPr, nil)
			if !ok {
				return nil, nil, false
			}
			notes = append(notes, lineNotes...)
			axis.MajorGridlines = line
		} else {
			axis.MajorGridlines = &NativeApproximateShapeLineV1{RGB: "D9D9D9", WidthEMU: 9525, Dash: "solid"}
			notes = append(notes, "gridline style defaulted to the Office chart style (D9D9D9, 0.75pt)")
		}
	}
	if firstDirectNativeChild(node, c, "minorGridlines") != nil {
		notes = append(notes, "minor gridlines are omitted")
	}
	if position, ok := nativeApproximateChartVal(node, c, "tickLblPos"); !ok || position != "none" {
		font := context.font(firstDirectNativeChild(node, c, "txPr"), defaultFont, &notes)
		axis.Labels = &font
	}
	if firstDirectNativeChild(node, c, "title") != nil {
		notes = append(notes, "axis titles are omitted")
	}
	if firstDirectNativeChild(node, c, "crossesAt") != nil {
		notes = append(notes, "authored axis crossing value is not applied")
	}
	return axis, notes, true
}

// title reads c:title. Rich or cached text is retained; an automatic title
// (no c:tx) keeps its band without text because the string Office displays is
// a UI default, not a cached value.
func (context *nativeApproximateChartContext) title(node *nativeXMLNode, defaultFont NativeApproximateChartFontV1, notes *[]string) *NativeApproximateChartTitleV1 {
	c, a := context.c, context.shapes.a
	title := &NativeApproximateChartTitleV1{Overlay: nativeApproximateChartBool(node, c, "overlay")}
	font := NativeApproximateChartFontV1{Family: defaultFont.Family, SizeHundredthPt: 1400, RGB: defaultFont.RGB}
	units := 0
	if tx := firstDirectNativeChild(node, c, "tx"); tx != nil {
		if rich := firstDirectNativeChild(tx, c, "rich"); rich != nil {
			parts := []string{}
			for _, p := range directNativeChildren(rich, a, "p") {
				if pPr := firstDirectNativeChild(p, a, "pPr"); pPr != nil {
					font = context.runFont(firstDirectNativeChild(pPr, a, "defRPr"), font, notes)
				}
				text := ""
				for _, r := range directNativeChildren(p, a, "r") {
					if rPr := firstDirectNativeChild(r, a, "rPr"); rPr != nil {
						font = context.runFont(rPr, font, notes)
					}
					if t := firstDirectNativeChild(r, a, "t"); t != nil {
						text += t.Text
					}
				}
				if len(directNativeChildren(p, a, "fld")) > 0 {
					*notes = append(*notes, "title fields are omitted")
				}
				if text != "" {
					parts = append(parts, text)
				}
			}
			if len(parts) > 0 {
				bounded, ok := context.boundedText(strings.Join(parts, "\n"), &units, nativeApproximateChartMaxTitleUnits)
				if !ok {
					return nil
				}
				title.Text = bounded
			}
		} else {
			text, ok := context.cachedText(tx, &units)
			if !ok {
				return nil
			}
			title.Text = text
		}
	}
	if txPr := firstDirectNativeChild(node, c, "txPr"); txPr != nil {
		font = context.font(txPr, font, notes)
	}
	title.Font = font
	if title.Text == nil {
		*notes = append(*notes, "automatic title text is an Office UI default and not cached; the title band is reserved without text")
	}
	return title
}

// font reads c:txPr/a:p/a:pPr/a:defRPr over a default.
func (context *nativeApproximateChartContext) font(txPr *nativeXMLNode, base NativeApproximateChartFontV1, notes *[]string) NativeApproximateChartFontV1 {
	if txPr == nil {
		return base
	}
	a := context.shapes.a
	for _, p := range directNativeChildren(txPr, a, "p") {
		if pPr := firstDirectNativeChild(p, a, "pPr"); pPr != nil {
			return context.runFont(firstDirectNativeChild(pPr, a, "defRPr"), base, notes)
		}
	}
	return base
}

func (context *nativeApproximateChartContext) runFont(rPr *nativeXMLNode, base NativeApproximateChartFontV1, notes *[]string) NativeApproximateChartFontV1 {
	if rPr == nil {
		return base
	}
	a := context.shapes.a
	font := base
	if size, ok := nativeNonnegativeInt64Attr(rPr, "", "sz"); ok && size >= 100 && size <= 400000 {
		font.SizeHundredthPt = size
	}
	if value, ok := nativeUnqualifiedAttr(rPr, "b"); ok {
		font.Bold = value == "1" || value == "true"
	}
	if value, ok := nativeUnqualifiedAttr(rPr, "i"); ok {
		font.Italic = value == "1" || value == "true"
	}
	if fill := firstDirectNativeChild(rPr, a, "solidFill"); fill != nil {
		if rgb, colorNotes, ok := context.shapes.solidColor(fill, ""); ok {
			font.RGB = rgb
			*notes = append(*notes, colorNotes...)
		} else {
			*notes = append(*notes, "chart text colour could not be resolved; default kept")
		}
	}
	if latin := firstDirectNativeChild(rPr, a, "latin"); latin != nil {
		if typeface, ok := nativeUnqualifiedAttr(latin, "typeface"); ok && len(typeface) <= 128 {
			switch {
			case typeface == "+mn-lt":
				if context.minorFont != "" {
					font.Family = context.minorFont
					*notes = append(*notes, "chart text uses the theme minor font")
				}
			case typeface == "+mj-lt":
				if context.majorFont != "" {
					font.Family = context.majorFont
					*notes = append(*notes, "chart text uses the theme major font")
				}
			case strings.HasPrefix(typeface, "+"):
			default:
				font.Family = typeface
			}
		}
	}
	return font
}

// nativeApproximateThemeFonts returns the theme minor and major Latin typefaces
// through the same single theme relationship the shape sidecar uses.
func nativeApproximateThemeFonts(pkg *nativePackage, mainPart, wordNS string) (string, string) {
	relBase, drawingNS := relBaseTransitional, drawingMLTransitional
	if wordNS == wordMLStrict {
		relBase, drawingNS = relBaseStrict, drawingMLStrict
	}
	partName, count := "", 0
	for _, rel := range pkg.rels[mainPart] {
		if rel.Type != relBase+"theme" {
			continue
		}
		if rel.External || rel.PartName == "" {
			return "", ""
		}
		count++
		partName = rel.PartName
	}
	if count != 1 || !nativeASCIIEqual(pkg.contentTypes[partName], "application/vnd.openxmlformats-officedocument.theme+xml") {
		return "", ""
	}
	root, err := parseNativeXML(partName, pkg.files[partName])
	if err != nil || root.Name != (xml.Name{Space: drawingNS, Local: "theme"}) {
		return "", ""
	}
	elements := firstDirectNativeChild(root, drawingNS, "themeElements")
	if elements == nil {
		return "", ""
	}
	scheme := firstDirectNativeChild(elements, drawingNS, "fontScheme")
	if scheme == nil {
		return "", ""
	}
	read := func(local string) string {
		group := firstDirectNativeChild(scheme, drawingNS, local)
		if group == nil {
			return ""
		}
		latin := firstDirectNativeChild(group, drawingNS, "latin")
		if latin == nil {
			return ""
		}
		typeface, _ := nativeUnqualifiedAttr(latin, "typeface")
		if len(typeface) > 128 {
			return ""
		}
		return typeface
	}
	return read("minorFont"), read("majorFont")
}

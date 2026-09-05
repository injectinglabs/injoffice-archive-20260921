package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strings"
)

// ChartInfo is a DrawingML chart summarized into InjOffice's spec vocabulary.
// It is the Go-side mirror of @injoffice/charts' ChartSpec: enough to render
// an agent-generated (openpyxl, Excel, ...) chart in the browser layer, and
// the shape the future chart WRITER will produce. Cell references stay in
// "SheetName!$A$1:$B$5" form — the TS host resolves sheet names to sheet ids.
type ChartInfo struct {
	// Part is the zip entry this chart came from (xl/charts/chart1.xml).
	Part string `json:"part"`
	// Identity is present when the chart has one unambiguous top-level drawing
	// anchor. Orphaned chart parts remain inspectable but cannot be mutated.
	Identity *ChartIdentity `json:"identity,omitempty"`
	// Type in ChartSpec vocabulary: column, bar, line, area, pie, doughnut,
	// scatter, radar — or the raw OOXML element name (e.g. "bubbleChart")
	// when unmapped, so callers can still say "there is a chart here".
	Type  string `json:"type"`
	Title string `json:"title,omitempty"`
	// Stacked is true for stacked/percentStacked groupings (rendered but not
	// yet round-tripped into SeriesOverride — carried for fidelity).
	Stacked bool        `json:"stacked,omitempty"`
	Series  []SeriesRef `json:"series"`
}

// SeriesRef is one c:ser: its display name (from the cached value when
// present) and the formula references its data binds to.
type SeriesRef struct {
	Name          string `json:"name,omitempty"`
	NameRef       string `json:"nameRef,omitempty"`
	CategoriesRef string `json:"categoriesRef,omitempty"`
	ValuesRef     string `json:"valuesRef,omitempty"`
}

// plot-type element local names → ChartSpec types ("" means resolve via barDir).
var chartTypeByElement = map[string]string{
	"barChart":      "", // column/bar via <c:barDir val=.../>
	"bar3DChart":    "",
	"lineChart":     "line",
	"line3DChart":   "line",
	"areaChart":     "area",
	"area3DChart":   "area",
	"pieChart":      "pie",
	"pie3DChart":    "pie",
	"doughnutChart": "doughnut",
	"scatterChart":  "scatter",
	"radarChart":    "radar",
}

// ReadCharts parses every chart part in the workbook. Unparseable individual
// charts are skipped with their part still reported (Type "unknown") rather
// than failing the whole read — the caller's question is "what charts exist
// and what can we render", not "is every chart pristine".
func ReadCharts(data []byte) ([]ChartInfo, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: read charts: %w", err)
	}
	ins, err := Inspect(data)
	if err != nil {
		return nil, err
	}
	parts := append([]string(nil), ins.ChartParts...)
	sort.Strings(parts)
	identities, err := ReadChartIdentities(data)
	if err != nil {
		return nil, err
	}
	identityByPart := make(map[string]ChartIdentity, len(identities))
	for _, identity := range identities {
		identityByPart[identity.Part] = identity
	}

	var out []ChartInfo
	for _, part := range parts {
		var file *zip.File
		for _, f := range zr.File {
			if f.Name == part {
				file = f
				break
			}
		}
		if file == nil {
			continue
		}
		rc, err := file.Open()
		if err != nil {
			out = append(out, ChartInfo{Part: part, Type: "unknown"})
			continue
		}
		info, perr := parseChartXML(rc)
		rc.Close() //nolint:errcheck
		if perr != nil {
			out = append(out, ChartInfo{Part: part, Type: "unknown"})
			continue
		}
		info.Part = part
		if identity, ok := identityByPart[part]; ok {
			copy := identity
			info.Identity = &copy
		}
		out = append(out, info)
	}
	return out, nil
}

// parseChartXML walks one chartSpace document with a namespace-agnostic
// element-stack tokenizer — resilient to prefix choices (c:, mc fallbacks)
// and to elements we don't model.
func parseChartXML(r io.Reader) (ChartInfo, error) {
	dec := xml.NewDecoder(r)
	var info ChartInfo
	var stack []string
	var curSer *SeriesRef
	inTitle := false

	has := func(names ...string) bool {
		// true when the innermost stack elements end with names...
		if len(stack) < len(names) {
			return false
		}
		for i, n := range names {
			if stack[len(stack)-len(names)+i] != n {
				return false
			}
		}
		return true
	}
	inside := func(name string) bool {
		for _, s := range stack {
			if s == name {
				return true
			}
		}
		return false
	}

	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return ChartInfo{}, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			local := t.Name.Local
			stack = append(stack, local)

			if mapped, isPlot := chartTypeByElement[local]; isPlot && inside("plotArea") && info.Type == "" {
				if mapped != "" {
					info.Type = mapped
				} else {
					info.Type = "column" // default barDir per OOXML is "col"
				}
			}
			switch local {
			case "barDir":
				if (info.Type == "column" || info.Type == "") && inside("plotArea") {
					if attrVal(t, "val") == "bar" {
						info.Type = "bar"
					}
				}
			case "grouping":
				if inside("plotArea") {
					v := attrVal(t, "val")
					if v == "stacked" || v == "percentStacked" {
						info.Stacked = true
					}
				}
			case "ser":
				info.Series = append(info.Series, SeriesRef{})
				curSer = &info.Series[len(info.Series)-1]
			case "title":
				if has("chart", "title") {
					inTitle = true
				}
			}
		case xml.EndElement:
			if t.Name.Local == "ser" {
				curSer = nil
			}
			if t.Name.Local == "title" && inTitle && !inside("ser") {
				inTitle = false
			}
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		case xml.CharData:
			text := string(t)
			if strings.TrimSpace(text) == "" {
				continue
			}
			switch {
			case inTitle && has("t"):
				info.Title += text
			case curSer != nil && has("f"):
				// Which reference is this formula for? Look upward for the
				// nearest of tx/cat/val (also xVal/yVal for scatter).
				switch nearestOf(stack, "tx", "cat", "val", "xVal", "yVal") {
				case "tx":
					curSer.NameRef = strings.TrimSpace(text)
				case "cat", "xVal":
					curSer.CategoriesRef = strings.TrimSpace(text)
				case "val", "yVal":
					curSer.ValuesRef = strings.TrimSpace(text)
				}
			case curSer != nil && curSer.Name == "" && has("tx", "strRef", "strCache", "pt", "v"):
				curSer.Name = text
			}
		}
	}
	if info.Type == "" {
		return ChartInfo{}, fmt.Errorf("no recognized plot type")
	}
	return info, nil
}

func nearestOf(stack []string, names ...string) string {
	for i := len(stack) - 1; i >= 0; i-- {
		for _, n := range names {
			if stack[i] == n {
				return n
			}
		}
	}
	return ""
}

func attrVal(el xml.StartElement, name string) string {
	for _, a := range el.Attr {
		if a.Name.Local == name {
			return a.Value
		}
	}
	return ""
}

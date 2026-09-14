package pptxpatch

import "slices"

// Dedicated preparation records do not attach to the public deck until the
// coordinated schema/compiler handoff. Existing literal validators remain intact.
type nativeStackedLineRecord struct {
	NativeLiteralConnected
	Grouping string `json:"grouping"`
}

func validNativeStackedBarRecord(c *NativeLiteralBar) bool {
	if c == nil || c.Profile != "literal-stacked-bar-v1" || (c.Grouping != "stacked" && c.Grouping != "percentStacked") || c.Overlap != 100 {
		return false
	}
	copy := *c
	copy.Profile = "literal-bar-v1"
	copy.Grouping = "clustered"
	copy.Overlap = 0
	copy.Series = slices.Clone(c.Series)
	orders := map[int64]bool{}
	for i := range copy.Series {
		order := copy.Series[i].Order
		if order < 0 || order >= int64(len(copy.Series)) || orders[order] {
			return false
		}
		orders[order] = true
		copy.Series[i].Order = int64(i)
	}
	// Original literal validation already accepts signed exact decimal points,
	// all source paint/labels and source-ordered aligned series. Group semantics
	// alone differ; the shallow local value is never written back to its caller.
	return validNativeLiteralBar(&copy)
}
func validNativeStackedLineRecord(c *nativeStackedLineRecord) bool {
	if c == nil || c.Profile != "literal-stacked-line-v1" || (c.Grouping != "stacked" && c.Grouping != "percentStacked") {
		return false
	}
	copy := c.NativeLiteralConnected
	copy.Profile = "literal-line-v1"
	copy.Series = slices.Clone(c.Series)
	orders := map[int64]bool{}
	for i := range copy.Series {
		order := copy.Series[i].Order
		if order < 0 || order >= int64(len(copy.Series)) || orders[order] {
			return false
		}
		orders[order] = true
		copy.Series[i].Order = int64(i)
	}
	return validNativeLiteralConnected(&copy)
}

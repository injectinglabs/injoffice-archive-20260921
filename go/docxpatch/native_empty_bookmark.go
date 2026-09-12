package docxpatch

import (
	"encoding/xml"
	"regexp"
	"strconv"
)

var nativeEmptyBookmarkName = regexp.MustCompile(`^[A-Za-z_][A-Za-z0-9_]{0,39}$`)

// Only adjacent empty markers are qualified here. Ranged bookmarks and field
// references keep their existing diagnostics; no navigation/edit semantics are
// synthesized from this read-only zero-width projection.
func nativeExactEmptyBookmark(nodes []*nativeXMLNode, ns string) bool {
	if len(nodes) < 2 {
		return false
	}
	start, end := nodes[0], nodes[1]
	if start.Name != (xml.Name{Space: ns, Local: "bookmarkStart"}) || end.Name != (xml.Name{Space: ns, Local: "bookmarkEnd"}) ||
		!nativeExactLeaf(start, xml.Name{Space: ns, Local: "id"}, xml.Name{Space: ns, Local: "name"}) ||
		!nativeExactLeaf(end, xml.Name{Space: ns, Local: "id"}) {
		return false
	}
	id, hasID := nativeAttr(start, ns, "id")
	endID, hasEndID := nativeAttr(end, ns, "id")
	name, hasName := nativeAttr(start, ns, "name")
	value, err := strconv.ParseUint(id, 10, 31)
	return hasID && hasEndID && id == endID && err == nil && strconv.FormatUint(value, 10) == id && hasName && nativeEmptyBookmarkName.MatchString(name)
}

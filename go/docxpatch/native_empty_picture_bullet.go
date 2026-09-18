package docxpatch

import "strings"

// A picture-bullet resource that carries no picture.
//
// ECMA-376 17.9.21 defines w:numPicBullet as the picture a level referencing it
// through w:lvlPicBulletId draws as its marker, and the picture itself is a
// relationship: the VML shape inside the w:pict names it with r:id/r:embed/
// r:link on a v:imagedata, or o:relid on an o:OLEObject. A w:numPicBullet whose
// whole subtree names no relationship at all therefore states a marker with no
// image in it -- Word draws nothing for it.
//
// Word's own PDF export of `lvlPicBulletId.docx` is the measurement. That
// package's only w:numPicBullet is an empty `v:shape` of the stock
// `#_x0000_t75` picture type, `filled="f" stroked="f"`, with no v:imagedata and
// no r:id -- and `word/numbering.xml` has no relationship part at all, so no
// relationship could be resolved even if one were named. Word paints the
// markers of that document's eight bulleted paragraphs as a single glyph in
// ArialMT whose ToUnicode maps it to U+0020: whitespace. There is no bullet ink
// on the page.
//
// Only a subtree that names nothing qualifies. A picture bullet that does carry
// a relationship is a visible graphic this tier has no marker-image input for,
// and stays PICTURE_BULLET_PRESERVED so that it keeps refusing rather than
// silently dropping the graphic.
func nativeAbsentPictureBullet(node *nativeXMLNode) bool {
	if node == nil {
		return false
	}
	for _, attr := range node.Attrs {
		if attr.Name.Space == relNSTransitional || attr.Name.Space == relNSStrict {
			return false
		}
		// o:relid on an embedded object names a relationship without the
		// relationships namespace, so the local name is checked too.
		if strings.EqualFold(attr.Name.Local, "relid") || strings.EqualFold(attr.Name.Local, "id") && strings.HasPrefix(attr.Value, "rId") {
			return false
		}
	}
	if strings.EqualFold(node.Name.Local, "imagedata") || strings.EqualFold(node.Name.Local, "blip") {
		return false
	}
	for _, child := range node.Children {
		if !nativeAbsentPictureBullet(child) {
			return false
		}
	}
	return true
}

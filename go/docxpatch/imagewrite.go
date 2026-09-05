// D11 breadth: inline images. Extends the surgical (never-regenerate)
// patcher to insert a real image as its own paragraph, immediately after a
// chosen paragraph. Unlike text edits (Apply's "set"), an image touches
// FOUR parts: word/document.xml (the new paragraph), word/_rels/
// document.xml.rels (a new relationship), [Content_Types].xml (a Default
// extension entry, if this image type hasn't appeared yet), and a brand new
// word/media/imageN.ext part — which is exactly why patch.go's general
// Patch{Replace/Add} exists: the original single-file rewriteZip could only
// replace one existing part.
//
// The OOXML chart-writing pattern in go/xlsxpatch/chartwrite.go (part
// naming, relationship-id allocation, content-types patching) is the same
// shape of problem one level up (drawingML wrapping a chart part instead of
// a picture) — this file's helpers are a docx-local rewrite of that same
// idea, not a shared import (docxpatch stays stdlib-only and
// dependency-free from xlsxpatch, same discipline both packages already
// follow independently).
package docxpatch

import (
	"archive/zip"
	"bytes"
	"fmt"
	"regexp"
	"strings"
)

const (
	docRelsPart  = "word/_rels/document.xml.rels"
	contentTypes = "[Content_Types].xml"
	relTypeImage = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	emptyRelsXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n" +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`
	// EMUPerPixel96DPI: OOXML measures drawings in EMUs (914400 per inch);
	// at the standard 96 CSS px/inch this is the px->EMU factor.
	EMUPerPixel96DPI = 914400 / 96
)

var imageContentTypes = map[string]string{
	"png":  "image/png",
	"jpg":  "image/jpeg",
	"jpeg": "image/jpeg",
	"gif":  "image/gif",
	"bmp":  "image/bmp",
	"tiff": "image/tiff",
}

// PixelsToEMU converts a CSS-pixel dimension (96 dpi) to EMUs for InsertImageAfter.
func PixelsToEMU(px int) int64 {
	return int64(px) * EMUPerPixel96DPI
}

// InsertImageAfter inserts a new paragraph containing an inline image
// immediately after paragraph afterIndex (-1 prepends before the first
// paragraph, matching Apply's insert_after convention). widthEMU/heightEMU
// are the display size in EMUs (see PixelsToEMU). ext must be one of
// png/jpg/jpeg/gif/bmp/tiff.
func InsertImageAfter(docx []byte, afterIndex int, image []byte, ext string, widthEMU, heightEMU int64) ([]byte, error) {
	ext = strings.ToLower(strings.TrimPrefix(ext, "."))
	ct, ok := imageContentTypes[ext]
	if !ok {
		return nil, fmt.Errorf("docxpatch: unsupported image extension %q (want png/jpg/jpeg/gif/bmp/tiff)", ext)
	}
	if widthEMU <= 0 || heightEMU <= 0 {
		return nil, fmt.Errorf("docxpatch: image dimensions must be positive (got %dx%d EMU)", widthEMU, heightEMU)
	}
	if len(image) == 0 {
		return nil, fmt.Errorf("docxpatch: empty image data")
	}

	zr, err := zip.NewReader(bytes.NewReader(docx), int64(len(docx)))
	if err != nil {
		return nil, fmt.Errorf("docxpatch: not a readable .docx: %w", err)
	}
	read := func(name string) ([]byte, bool) {
		for _, f := range zr.File {
			if f.Name == name {
				rc, err := f.Open()
				if err != nil {
					return nil, false
				}
				defer rc.Close()
				var buf bytes.Buffer
				if _, err := buf.ReadFrom(rc); err != nil {
					return nil, false
				}
				return buf.Bytes(), true
			}
		}
		return nil, false
	}

	doc, ok := read(docPart)
	if !ok {
		return nil, fmt.Errorf("docxpatch: %s not found", docPart)
	}
	s := string(doc)
	paras := scanParas(s)
	src := afterIndex
	if src < 0 {
		src = 0
	}
	if len(paras) == 0 || src >= len(paras) {
		return nil, fmt.Errorf("docxpatch: insert image after paragraph %d: paragraph does not exist", afterIndex)
	}
	anchor := paras[src]

	// Media part naming: next free imageN.ext across ANY existing media
	// (not just this extension) so two images never collide even with
	// mixed types.
	mediaN := 1
	mediaRe := regexp.MustCompile(`^word/media/image(\d+)\.`)
	for _, f := range zr.File {
		if m := mediaRe.FindStringSubmatch(f.Name); m != nil {
			var n int
			fmt.Sscanf(m[1], "%d", &n) //nolint:errcheck
			if n >= mediaN {
				mediaN = n + 1
			}
		}
	}
	mediaPart := fmt.Sprintf("word/media/image%d.%s", mediaN, ext)

	relsXML, hasRels := read(docRelsPart)
	relsStr := emptyRelsXML
	if hasRels {
		relsStr = string(relsXML)
	}
	relID := nextFreeRelID(relsStr)
	newRels, err := appendRelationship(relsStr, relID, relTypeImage, "media/"+mediaPart[len("word/media/"):])
	if err != nil {
		return nil, err
	}

	ctXML, ok := read(contentTypes)
	if !ok {
		return nil, fmt.Errorf("docxpatch: %s not found", contentTypes)
	}
	newCT, err := defaultExtensionWith(string(ctXML), ext, ct)
	if err != nil {
		return nil, err
	}

	// docPr ids must be unique across the WHOLE document (images, charts,
	// any other drawing) — NOT the same counter as media file naming, which
	// is scoped to images only. Scanning the live document.xml for every
	// existing wp:docPr id (see nextFreeDocPrID in chartwrite.go) is what
	// keeps this correct once InsertChart exists too.
	picID := nextFreeDocPrID(s)
	drawingPara := imageParagraphXML(relID, picID, widthEMU, heightEMU)
	newDoc := s[:anchor.end] + drawingPara + s[anchor.end:]
	if afterIndex < 0 {
		newDoc = s[:anchor.start] + drawingPara + s[anchor.start:]
	}

	patch := Patch{
		Replace: map[string][]byte{
			docPart:      []byte(newDoc),
			contentTypes: []byte(newCT),
		},
		Add: map[string][]byte{
			mediaPart: image,
		},
	}
	if hasRels {
		patch.Replace[docRelsPart] = []byte(newRels)
	} else {
		patch.Add[docRelsPart] = []byte(newRels)
	}

	return ApplyPatch(docx, patch)
}

// imageParagraphXML builds a standalone <w:p> containing one inline
// drawing. All namespaces used (wp/a/pic/r) are declared locally on the
// elements that need them rather than assumed present at the document
// root — the same defensive posture as scanParas' non-nesting assumption:
// never trust what the source document declared elsewhere.
func imageParagraphXML(relID string, picID int, cx, cy int64) string {
	return fmt.Sprintf(
		`<w:p><w:r><w:drawing>`+
			`<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">`+
			`<wp:extent cx="%d" cy="%d"/>`+
			`<wp:effectExtent l="0" t="0" r="0" b="0"/>`+
			`<wp:docPr id="%d" name="Picture %d"/>`+
			`<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>`+
			`<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`+
			`<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`+
			`<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">`+
			`<pic:nvPicPr><pic:cNvPr id="%d" name="Picture %d"/><pic:cNvPicPr/></pic:nvPicPr>`+
			`<pic:blipFill><a:blip r:embed="%s" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`+
			`<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`+
			`</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
		cx, cy, picID, picID, picID, picID, relID, cx, cy,
	)
}

// ---- relationship / content-type plumbing (docx-local; see file header) ----

var relIDRe = regexp.MustCompile(`Id="rId(\d+)"`)

func nextFreeRelID(relsXML string) string {
	max := 0
	for _, m := range relIDRe.FindAllStringSubmatch(relsXML, -1) {
		var n int
		fmt.Sscanf(m[1], "%d", &n) //nolint:errcheck
		if n > max {
			max = n
		}
	}
	return fmt.Sprintf("rId%d", max+1)
}

func relationshipXML(id, relType, target string) string {
	return fmt.Sprintf(`<Relationship Id=%q Type=%q Target=%q/>`, id, relType, target)
}

func appendRelationship(relsXML, id, relType, target string) (string, error) {
	idx := strings.LastIndex(relsXML, "</Relationships>")
	if idx < 0 {
		return "", fmt.Errorf("docxpatch: malformed rels part")
	}
	return relsXML[:idx] + relationshipXML(id, relType, target) + relsXML[idx:], nil
}

// defaultExtensionWith adds <Default Extension="ext" ContentType="ct"/> to
// [Content_Types].xml unless that extension is already registered
// (multiple images of the same type share one Default entry).
func defaultExtensionWith(ct, ext, contentType string) (string, error) {
	if strings.Contains(ct, `Extension="`+ext+`"`) {
		return ct, nil
	}
	idx := strings.LastIndex(ct, "</Types>")
	if idx < 0 {
		return "", fmt.Errorf("docxpatch: malformed [Content_Types].xml")
	}
	entry := fmt.Sprintf(`<Default Extension=%q ContentType=%q/>`, ext, contentType)
	return ct[:idx] + entry + ct[idx:], nil
}

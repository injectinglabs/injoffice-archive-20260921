package officehttp

import (
	"archive/zip"
	"bytes"
	"context"
	"crypto/sha256"
	"fmt"
	"testing"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

// A picture-filled DrawingML rectangle is refused by the strict extractor, so
// the strict document names no media part for it. The approximate drawing-shape
// sidecar does, and its part must reach the compiler once, alongside whatever
// the strict document already supplied.
func TestDOCXPreviewMediaAssetsIncludeShapePictureFills(t *testing.T) {
	const wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
	image := []byte("\xff\xd8\xff\xe0 picture bytes the compiler verifies")
	drawing := `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:wps="` + wps + `" Requires="wps"><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="` + wps + `" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>2399665</wp:posOffset></wp:positionV><wp:extent cx="3197829" cy="1934210"/><wp:effectExtent l="0" t="0" r="3175" b="8890"/><wp:wrapNone/><wp:docPr id="89" name="Rectangle 89"/><wp:cNvGraphicFramePr/>` +
		`<a:graphic><a:graphicData uri="` + wps + `"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3197829" cy="1934210"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId10"/><a:srcRect/><a:stretch><a:fillRect t="-115561" b="-13789"/></a:stretch></a:blipFill><a:ln><a:noFill/></a:ln></wps:spPr><wps:bodyPr rot="0" vert="horz" wrap="square" anchor="ctr"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>` +
		`<mc:Fallback><w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml" style="position:absolute;width:100pt;height:50pt" fillcolor="#ff0000"/></w:pict></mc:Fallback></mc:AlternateContent>`
	parts := map[string][]byte{
		"[Content_Types].xml":          []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`),
		"_rels/.rels":                  []byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
		"word/document.xml":            []byte(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>A</w:t></w:r><w:r>` + drawing + `</w:r></w:p><w:sectPr/></w:body></w:document>`),
		"word/_rels/document.xml.rels": []byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="settings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.jpg"/></Relationships>`),
		"word/settings.xml":            []byte(`<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`),
		"word/media/image1.jpg":        image,
	}
	var buffer bytes.Buffer
	writer := zip.NewWriter(&buffer)
	for name, content := range parts {
		entry, err := writer.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err = entry.Write(content); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	data := buffer.Bytes()
	input, err := docxPreviewInput(context.Background(), data)
	if err != nil {
		t.Fatal(err)
	}
	strict, _ := input["media_assets"].([]map[string]any)
	if len(strict) != 0 {
		t.Fatalf("the strict document names no picture for a refused shape: %#v", strict)
	}
	shapes, err := docxpatch.InspectNativeApproximateDrawingShapesV1(data)
	if err != nil {
		t.Fatal(err)
	}
	if shapes == nil || len(shapes.Items) != 1 || shapes.Items[0].BlipFill == nil {
		t.Fatalf("sidecar must resolve the picture fill: %#v", shapes)
	}
	media, err := docxPreviewMediaAssets(context.Background(), data, shapes, strict)
	if err != nil {
		t.Fatal(err)
	}
	if len(media) != 1 || media[0]["part_name"] != "word/media/image1.jpg" || media[0]["content_type"] != "image/jpeg" || media[0]["content_digest"] != fmt.Sprintf("sha256:%x", sha256.Sum256(image)) {
		t.Fatalf("shape picture part not supplied: %#v", media)
	}
	// A part the strict document already supplied is not read twice.
	again, err := docxPreviewMediaAssets(context.Background(), data, shapes, media)
	if err != nil {
		t.Fatal(err)
	}
	if len(again) != 1 || again[0]["content_digest"] != media[0]["content_digest"] {
		t.Fatalf("existing media must be kept once: %#v", again)
	}
	// Only image/png and image/jpeg parts are read, whatever the sidecar names.
	other, err := docxPreviewMediaAssets(context.Background(), data, map[string]any{"media_part": "word/media/image1.jpg", "content_type": "image/gif"}, nil)
	if err != nil || len(other) != 0 {
		t.Fatalf("unexpected media for an unsupported content type: %#v %v", other, err)
	}
}

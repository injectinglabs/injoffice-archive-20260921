package docxpatch

// Insertions splice source XML; they never regenerate an existing paragraph.
import (
	"bytes"
	"encoding/base64"
	"fmt"
	"image"
	_ "image/jpeg"
	_ "image/png"
	"strconv"
	"strings"
)

type nativeInsertImage struct {
	Data        string `json:"data_base64"`
	ContentType string `json:"content_type"`
	Width       int64  `json:"width_emu"`
	Height      int64  `json:"height_emu"`
	Alt         string `json:"alt_text"`
}
type nativeInsertMutation struct {
	TargetKind string             `json:"target_kind"`
	TargetID   string             `json:"target_id"`
	SHA        string             `json:"expected_xml_sha256"`
	Operation  string             `json:"operation"`
	Text       *string            `json:"text,omitempty"`
	Image      *nativeInsertImage `json:"image,omitempty"`
	RunID      string             `json:"run_id,omitempty"`
	Offset     *int               `json:"offset_utf16,omitempty"`
}

func init() {
	nativeOperations["page_break.insert"] = true
	nativeParagraphOperations["page_break.insert"] = true
}

func advertiseNativeInsert(doc *NativeDocumentV1) {
	if doc.Source.MainPart != docPart {
		return
	}
	for _, block := range doc.Body.Blocks {
		p := block.Paragraph
		if p == nil || p.EditPolicy.Mode != "read-write" {
			continue
		}
		p.EditPolicy.AllowedOperations = append(p.EditPolicy.AllowedOperations, "page_break.insert")
	}
}

func nativeDecodeInsertImage(raw []byte) (*nativeInsertImage, error) {
	fields, err := nativeFlatJSONObject(raw)
	if err != nil || len(fields) != 5 {
		return nil, fmt.Errorf("invalid image fields")
	}
	im := &nativeInsertImage{}
	for name, to := range map[string]*string{"data_base64": &im.Data, "content_type": &im.ContentType, "alt_text": &im.Alt} {
		value, err := decodeNativeMutationJSONString(fields[name])
		if err != nil {
			return nil, err
		}
		*to = value
	}
	width, ok := nativeJSONInt(fields["width_emu"])
	height, ok2 := nativeJSONInt(fields["height_emu"])
	if !ok || !ok2 {
		return nil, fmt.Errorf("invalid image extents")
	}
	im.Width, im.Height = int64(width), int64(height)
	return im, nil
}

func applyNativeInsert(source []byte, revision string, mutation nativeDOCXMutationV1) (*NativeDOCXMutationResultV1, error) {
	fail := func(code, message string) (*NativeDOCXMutationResultV1, error) {
		return nil, nativeMutationError(code, mutation.TargetID, message)
	}
	m := nativeInsertMutation{TargetKind: mutation.TargetKind, TargetID: mutation.TargetID, SHA: mutation.ExpectedXMLSHA256, Operation: mutation.Operation, Image: mutation.Image, RunID: mutation.SplitRunID, Offset: &mutation.SplitOffset}
	if mutation.Operation == "block.insert_after" {
		m.Offset = nil
	}
	if revision != nativeSHA(source) {
		return fail("STALE_REVISION", "package revision changed")
	}
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		return nil, err
	}
	pkg, err := openNativeDOCXPackage(source)
	if err != nil {
		return nil, err
	}
	if nativeDOCXPackageHasDigitalSignature(pkg) {
		return fail("SIGNED_PACKAGE", "signed packages cannot be modified")
	}
	index := -1
	var p *NativeParagraphV1
	for i, b := range doc.Body.Blocks {
		if b.Paragraph != nil && b.Paragraph.ID == m.TargetID {
			index = i
			p = b.Paragraph
		}
	}
	if p == nil || m.TargetKind != "paragraph" || !nativePolicyAllows(p.EditPolicy, m.Operation) {
		return fail("UNSUPPORTED_TARGET", "choose a writable body paragraph")
	}
	if p.Anchor.XMLSHA256 != m.SHA {
		return fail("STALE_TARGET", "paragraph anchor changed")
	}
	part := pkg.files[p.Anchor.PartName]
	root, err := parseNativeXML(p.Anchor.PartName, part)
	if err != nil {
		return nil, err
	}
	node := nativeNodeByPath(root, p.Anchor.Path)
	if node == nil || node.Name.Space != wordMLTransitional {
		return fail("UNSUPPORTED_CONSTRUCT", "insertion requires transitional WordprocessingML")
	}
	empty := `<w:p xmlns:w="` + wordMLTransitional + `"><w:r><w:t></w:t></w:r></w:p>`
	start, end := node.End, node.End
	inserted := empty
	patch := Patch{Replace: map[string][]byte{}, Add: map[string][]byte{}}
	media := ""
	if m.Operation == "block.insert_after" {
		if m.Image == nil {
			return fail("INVALID_PAYLOAD", "insert an empty paragraph or an image")
		}
		if m.Image != nil {
			im := m.Image
			if im.Width < 1 || im.Height < 1 || im.Width > 20116800 || im.Height > 20116800 || len(im.Alt) > 1024 || !nativeMutationXMLTextValid(im.Alt) || strings.ContainsAny(im.Alt, "\x00\r\n\t") {
				return fail("INVALID_IMAGE", "image extents or alternative text are invalid")
			}
			data, err := base64.StdEncoding.Strict().DecodeString(im.Data)
			if err != nil || len(data) == 0 || len(data) > 4*1024*1024 {
				return fail("INVALID_IMAGE", "image must be bounded base64 PNG or JPEG")
			}
			format, width, height := nativeInsertImageHeader(data)
			if format == "" || "image/"+format != im.ContentType || width < 1 || height < 1 || width*height > 16000000 {
				return fail("INVALID_IMAGE", "image bytes must match a bounded PNG or JPEG")
			}
			ext := format
			if ext == "jpeg" {
				ext = "jpg"
			}
			for n := 1; ; n++ {
				media = fmt.Sprintf("word/media/insert%d.%s", n, ext)
				collision := false
				for name := range pkg.files {
					if strings.EqualFold(name, media) {
						collision = true
					}
				}
				if !collision {
					break
				}
			}
			rels := pkg.files[docRelsPart]
			if rels == nil {
				rels = []byte(emptyRelsXML)
			}
			relRoot, err := parseNativeXML(docRelsPart, rels)
			if err != nil {
				return nil, err
			}
			ids := map[string]bool{}
			for _, child := range relRoot.Children {
				for _, a := range child.Attrs {
					if a.Name.Local == "Id" {
						ids[a.Value] = true
					}
				}
			}
			relID := ""
			for n := 1; ; n++ {
				relID = fmt.Sprintf("rIdInsert%d", n)
				if !ids[relID] {
					break
				}
			}
			// Namespace declarations are local so source prefix choices stay untouched.
			rel := `<Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships" Id="` + relID + `" Type="` + relTypeImage + `" Target="media/` + strings.TrimPrefix(media, "word/media/") + `"/>`
			newRels, err := nativeInsertChild(rels, relRoot, rel)
			if err != nil {
				return fail("UNSUPPORTED_LEXICAL_FORM", err.Error())
			}
			if pkg.files[docRelsPart] == nil {
				patch.Add[docRelsPart] = newRels
			} else {
				patch.Replace[docRelsPart] = newRels
			}
			ct := pkg.files[contentTypes]
			ctRoot, err := parseNativeXML(contentTypes, ct)
			if err != nil {
				return nil, err
			}
			newCT, err := nativeInsertChild(ct, ctRoot, `<Override xmlns="http://schemas.openxmlformats.org/package/2006/content-types" PartName="/`+media+`" ContentType="`+im.ContentType+`"/>`)
			if err != nil {
				return nil, err
			}
			patch.Replace[contentTypes] = newCT
			patch.Add[media] = data
			maxID := 0
			var walk func(*nativeXMLNode)
			walk = func(n *nativeXMLNode) {
				if n.Name.Local == "docPr" {
					for _, a := range n.Attrs {
						if a.Name.Local == "id" {
							id, _ := strconv.Atoi(a.Value)
							if id > maxID {
								maxID = id
							}
						}
					}
				}
				for _, c := range n.Children {
					walk(c)
				}
			}
			walk(root)
			inserted = imageParagraphXML(relID, maxID+1, im.Width, im.Height)
			inserted = strings.Replace(inserted, "<w:p>", `<w:p xmlns:w="`+wordMLTransitional+`">`, 1)
			inserted = strings.Replace(inserted, "<wp:docPr ", `<wp:docPr descr="`+nativeMutationEscapeAttribute(im.Alt)+`" `, 1) + empty
		}
	} else if m.Operation == "page_break.insert" {
		if m.Image != nil || m.Text != nil || m.Offset == nil || *m.Offset < 0 {
			return fail("INVALID_PAYLOAD", "page break requires a run and caret offset")
		}
		var run *NativeRunV1
		for i := range p.Runs {
			if p.Runs[i].ID == m.RunID {
				run = &p.Runs[i]
			}
		}
		if run == nil || run.Kind != "text" || run.Text == nil {
			return fail("UNSUPPORTED_TARGET", "page break requires a text run")
		}
		pos, ok := nativeUTF16ByteOffset(*run.Text, *m.Offset)
		if !ok {
			return fail("INVALID_RANGE", "caret is outside text or splits a surrogate pair")
		}
		textNode := nativeNodeByPath(node, run.Anchor.Path)
		if textNode == nil || textNode.parent == nil || textNode.parent.parent != node {
			return fail("UNSUPPORTED_CONSTRUCT", "caret must be in a direct paragraph run")
		}
		owner := textNode.parent
		for _, c := range owner.Children {
			if c != textNode && c.Name.Local != "rPr" {
				return fail("UNSUPPORTED_CONSTRUCT", "caret run contains other content")
			}
		}
		left, err := nativeFormatTextElement(part[textNode.Start:textNode.End], (*run.Text)[:pos])
		if err != nil {
			return nil, err
		}
		right, err := nativeFormatTextElement(part[textNode.Start:textNode.End], (*run.Text)[pos:])
		if err != nil {
			return nil, err
		}
		start, end = textNode.Start, textNode.End
		inserted = string(left) + `<w:br xmlns:w="` + wordMLTransitional + `" w:type="page"/>` + string(right)
	} else {
		return fail("INVALID_PAYLOAD", "unsupported insertion operation")
	}
	output := append([]byte(nil), part[:start]...)
	output = append(output, inserted...)
	output = append(output, part[end:]...)
	patch.Replace[p.Anchor.PartName] = output
	produced, err := ApplyPatch(source, patch)
	if err != nil {
		return nil, err
	}
	after, err := ExtractNativeDocumentV1WithOptions(produced, NativeExtractionOptions{Previous: doc})
	if err != nil {
		return nil, err
	}
	if m.Operation == "page_break.insert" {
		beforeText, afterText := "", ""
		beforeBreaks, afterBreaks := 0, 0
		for _, r := range p.Runs {
			if r.Text != nil {
				beforeText += *r.Text
			}
			if r.Control == "page-break" {
				beforeBreaks++
			}
		}
		actual := after.Body.Blocks[index].Paragraph
		if actual == nil {
			return fail("POST_WRITE_MISMATCH", "paragraph missing")
		}
		for _, r := range actual.Runs {
			if r.Text != nil {
				afterText += *r.Text
			}
			if r.Control == "page-break" {
				afterBreaks++
			}
		}
		if beforeText != afterText || afterBreaks != beforeBreaks+1 {
			return fail("POST_WRITE_MISMATCH", "page break did not round trip")
		}
	} else {
		actual := after.Body.Blocks[index+1].Paragraph
		if actual == nil {
			return fail("POST_WRITE_MISMATCH", "inserted paragraph missing")
		}
		if m.Image != nil {
			if len(actual.Runs) != 1 || actual.Runs[0].Drawing == nil {
				return fail("POST_WRITE_MISMATCH", "inserted image missing")
			}
			d := actual.Runs[0].Drawing
			if d.MediaPart == nil || *d.MediaPart != media || d.WidthEMU == nil || *d.WidthEMU != m.Image.Width || d.HeightEMU == nil || *d.HeightEMU != m.Image.Height {
				return fail("POST_WRITE_MISMATCH", "image relationship or extents changed")
			}
		} else if len(actual.Runs) != 1 || actual.Runs[0].Text == nil || *actual.Runs[0].Text != "" {
			return fail("POST_WRITE_MISMATCH", "empty paragraph did not round trip")
		}
	}
	return &NativeDOCXMutationResultV1{Package: produced, Document: after, Evidence: NativeDOCXMutationEvidenceV1{SourceRevision: revision, ResultRevision: nativeSHA(produced), UntouchedPartsVerified: len(pkg.files) - len(patch.Replace)}}, nil
}

func nativeInsertChild(part []byte, root *nativeXMLNode, child string) ([]byte, error) {
	end := int(root.End)
	start := bytes.LastIndex(part[:end], []byte("</"))
	if start < int(root.Start) {
		return nil, fmt.Errorf("container must have an explicit closing tag")
	}
	result := append([]byte(nil), part[:start]...)
	result = append(result, child...)
	result = append(result, part[start:]...)
	return result, nil
}

// Validate dimensions before decoding so malformed rasters cannot force
// unbounded pixel allocations. The original compressed bytes remain authority.
func nativeInsertImageHeader(data []byte) (string, int64, int64) {
	config, format, err := image.DecodeConfig(bytes.NewReader(data))
	if err != nil || (format != "png" && format != "jpeg") || config.Width < 1 || config.Height < 1 || int64(config.Width)*int64(config.Height) > 16000000 {
		return "", 0, 0
	}
	decoded, decodedFormat, err := image.Decode(bytes.NewReader(data))
	if err != nil || decodedFormat != format || decoded.Bounds().Dx() != config.Width || decoded.Bounds().Dy() != config.Height {
		return "", 0, 0
	}
	return format, int64(config.Width), int64(config.Height)
}

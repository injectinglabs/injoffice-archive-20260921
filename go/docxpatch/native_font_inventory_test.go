package docxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"encoding/hex"
	"errors"
	"os"
	"strings"
	"testing"
)

const nativeFontTestKey = "{00112233-4455-6677-8899-AABBCCDDEEFF}"

func TestNativeDOCXFontInventoryTransitionalContentAddressedResolver(t *testing.T) {
	sfnt := nativeFontTestSFNT(0x0008)
	source := nativeFontTestPackage(t, false, sfnt, "false", nil)
	first, err := ExtractNativeDOCXFontInventoryV1(source)
	if err != nil {
		t.Fatal(err)
	}
	second, err := ExtractNativeDOCXFontInventoryV1(source)
	if err != nil {
		t.Fatal(err)
	}
	firstJSON, err := EncodeNativeDOCXFontInventoryV1(first)
	if err != nil {
		t.Fatal(err)
	}
	secondJSON, err := EncodeNativeDOCXFontInventoryV1(second)
	if err != nil || !bytes.Equal(firstJSON, secondJSON) {
		t.Fatalf("inventory is nondeterministic: %v\n%s\n%s", err, firstJSON, secondJSON)
	}
	golden, err := os.ReadFile("testdata/font-inventory-v1.json")
	if err != nil {
		t.Fatal(err)
	}
	goldenInventory, err := ExtractNativeDOCXFontInventoryV1(nativeFontTestPackage(t, false, nativeFontTestSFNT(0), "false", nil))
	if err != nil {
		t.Fatal(err)
	}
	goldenJSON, err := EncodeNativeDOCXFontInventoryV1(goldenInventory)
	if err != nil || !bytes.Equal(bytes.TrimSpace(golden), goldenJSON) {
		t.Fatalf("canonical Go/Node inventory golden drifted: %v\n%s\n%s", err, golden, goldenJSON)
	}
	decoded, err := DecodeNativeDOCXFontInventoryV1(firstJSON)
	if err != nil {
		t.Fatal(err)
	}
	pkg, err := openNativeDOCXPackage(source)
	if err != nil {
		t.Fatal(err)
	}
	if decoded.InventorySHA256 != first.InventorySHA256 || first.Revision != "rev:"+strings.TrimPrefix(first.PackageSHA256, "sha256:")[:32] || first.MainSHA256 != nativeSHA(pkg.files[first.MainPart]) {
		t.Fatalf("inventory identity is not bound to the full package: %#v", first)
	}
	if first.FontTable == nil || first.FontTable.PartName != "Word/FontTable.XML" || first.FontTable.MainRelationshipsPart != "Word/_rels/Document.XML.rels" || first.FontTable.FontRelationshipsPart == nil || *first.FontTable.FontRelationshipsPart != "Word/_rels/FontTable.XML.rels" || first.FontTable.RelationshipID != "fontTable" || first.FontTable.RelationshipTarget != "FontTable.XML" {
		t.Fatalf("font-table provenance was not exact: %#v", first.FontTable)
	}
	if len(first.Families) != 1 || len(first.Families[0].Faces) != 1 || len(first.References) != 1 {
		t.Fatalf("unexpected inventory: %#v", first)
	}
	face := first.Families[0].Faces[0]
	if face.Family != "Fixture Sans" || face.AltName == nil || *face.AltName != "Fixture Alias" || face.Weight != 400 || face.Style != "normal" || face.Source.FaceSlot != "embedRegular" || face.Source.RelationshipTarget != "Fonts/Face.ODTTF" || face.Source.AssetPart != "Word/Fonts/Face.ODTTF" || face.Source.StoredSHA256 == face.Source.ContentSHA256 || face.Source.CollectionIndex != nil {
		t.Fatalf("face metadata/content addressing is incomplete: %#v", face)
	}
	if face.Source.Licensing.RightsStatus != "verified" || face.Source.Licensing.EmbeddingRights != "editable" || face.Source.Licensing.AllowedScope != "document-only" {
		t.Fatalf("licensing was not enforced: %#v", face.Source.Licensing)
	}
	if first.NativeTextManifest == nil || first.NativeTextManifestSHA256 == nil || *first.NativeTextManifestSHA256 != nativeDOCXCanonicalWireSHA256(first.NativeTextManifest) || len(first.NativeTextManifest.Faces) != 1 || len(first.NativeTextManifest.FallbackChains) != 0 || first.NativeTextManifest.Faces[0].Source.Kind != "document" || first.NativeTextManifest.Faces[0].Source.ContentDigest != face.Source.ContentSHA256 {
		t.Fatalf("generic manifest projection is not fail-closed: %#v", first.NativeTextManifest)
	}
	resolved, err := ResolveNativeDOCXFontAssetV1(source, first, NativeDOCXFontResolutionRequestV1{Family: "fixture alias", Weight: 400, Style: "normal", Stretch: 100})
	if err != nil || !bytes.Equal(resolved.Bytes, sfnt) || resolved.Face.FaceID != face.FaceID {
		t.Fatalf("exact document font resolution failed: asset=%#v err=%v", resolved, err)
	}
	workerAssets, err := ResolveNativeDOCXPagePaintFontAssetsV1(source, first)
	if err != nil || len(workerAssets) != 1 || workerAssets[0].FaceID != face.FaceID || workerAssets[0].FaceSlot != face.Source.FaceSlot || workerAssets[0].ResourceID != face.Source.ResourceID || workerAssets[0].ContentDigest != face.Source.ContentSHA256 || workerAssets[0].CollectionIndex != nil || workerAssets[0].BytesBase64 != base64.StdEncoding.EncodeToString(sfnt) {
		t.Fatalf("canonical worker asset bridge is incomplete: assets=%#v err=%v", workerAssets, err)
	}
	for _, request := range []NativeDOCXFontResolutionRequestV1{
		{Family: "System Font", Weight: 400, Style: "normal", Stretch: 100},
		{Family: "Fixture Sans", Weight: 700, Style: "normal", Stretch: 100},
		{Family: "Fixture Sans", Weight: 400, Style: "oblique", Stretch: 100},
		{Family: "Fixture Sans", Weight: 400, Style: "normal", Stretch: 90},
	} {
		if _, err := ResolveNativeDOCXFontAssetV1(source, first, request); err == nil {
			t.Fatalf("non-exact/system/substituted request was accepted: %#v", request)
		}
	}
}

func TestNativeDOCXFontInventoryStrictRelocatedParts(t *testing.T) {
	source := nativeFontTestPackage(t, true, nativeFontTestSFNT(0), "false", nil)
	inventory, err := ExtractNativeDOCXFontInventoryV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if inventory.MainPart != "Word/Document.XML" || inventory.FontTable == nil || inventory.FontTable.RelationshipType != relBaseStrict+"fontTable" || inventory.Families[0].Faces[0].Source.RelationshipType != relBaseStrict+"font" || inventory.Families[0].Faces[0].Source.Licensing.EmbeddingRights != "installable" {
		t.Fatalf("Strict font graph was not preserved exactly: %#v", inventory)
	}
}

func TestNativeDOCXFontInventoryRejectsAdversarialRelationshipsPathsFontsAndLicenses(t *testing.T) {
	valid := nativeFontTestSFNT(0x0004)
	tests := []struct {
		name      string
		strict    bool
		sfnt      []byte
		subsetted string
		mutate    func(map[string]string)
		want      string
	}{
		{"wrong dialect font relationship", false, valid, "false", func(parts map[string]string) {
			parts["Word/_rels/FontTable.XML.rels"] = strings.Replace(parts["Word/_rels/FontTable.XML.rels"], relBaseTransitional+"font", relBaseStrict+"font", 1)
		}, "not an exact internal"},
		{"external font relationship", false, valid, "false", func(parts map[string]string) {
			parts["Word/_rels/FontTable.XML.rels"] = strings.Replace(parts["Word/_rels/FontTable.XML.rels"], `Target="Fonts/Face.ODTTF"`, `Target="https://example.invalid/font" TargetMode="External"`, 1)
		}, "not an exact internal"},
		{"wrong asset content type", false, valid, "false", func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], nativeDOCXObfuscatedFontType, "application/octet-stream", 1)
		}, "unsupported content type"},
		{"duplicate face slot", false, valid, "false", func(parts map[string]string) {
			marker := `<w:embedRegular r:id="fontRegular" w:fontKey="` + nativeFontTestKey + `" w:subsetted="false"/>`
			parts["Word/FontTable.XML"] = strings.Replace(parts["Word/FontTable.XML"], marker, marker+marker, 1)
		}, "duplicate embedRegular"},
		{"duplicate embedded resource", false, valid, "false", func(parts map[string]string) {
			parts["Word/FontTable.XML"] = strings.Replace(parts["Word/FontTable.XML"], `</w:font>`, `<w:embedBold r:id="fontBold" w:fontKey="`+nativeFontTestKey+`" w:subsetted="false"/></w:font>`, 1)
			parts["Word/_rels/FontTable.XML.rels"] = strings.Replace(parts["Word/_rels/FontTable.XML.rels"], `</Relationships>`, `<Relationship Id="fontBold" Type="`+relBaseTransitional+`font" Target="Fonts/Face.ODTTF"/></Relationships>`, 1)
		}, "duplicate font resource"},
		{"missing relationship id", false, valid, "false", func(parts map[string]string) {
			parts["Word/FontTable.XML"] = strings.Replace(parts["Word/FontTable.XML"], ` r:id="fontRegular"`, "", 1)
		}, "invalid relationship id"},
		{"noncanonical lowercase key", false, valid, "false", func(parts map[string]string) {
			parts["Word/FontTable.XML"] = strings.Replace(parts["Word/FontTable.XML"], nativeFontTestKey, strings.ToLower(nativeFontTestKey), 1)
		}, "invalid canonical fontKey"},
		{"zero key", false, valid, "false", func(parts map[string]string) {
			parts["Word/FontTable.XML"] = strings.Replace(parts["Word/FontTable.XML"], nativeFontTestKey, nativeDOCXZeroFontKey, 1)
		}, "invalid canonical fontKey"},
		{"short font", false, []byte{0, 1, 2}, "false", nil, "at least 32 bytes"},
		{"malformed sfnt table range", false, nativeFontTestMalformedRange(), "false", nil, "malformed sfnt table"},
		{"restricted license", false, nativeFontTestSFNT(0x0002), "false", nil, "restricted-license"},
		{"reserved license bit", false, nativeFontTestSFNT(0x0001), "false", nil, "reserved or conflicting"},
		{"bitmap only", false, nativeFontTestSFNT(0x0200), "false", nil, "bitmap-only"},
		{"no subsetting conflict", false, nativeFontTestSFNT(0x0100), "true", nil, "no-subsetting"},
		{"wrong main font-table dialect", true, valid, "false", func(parts map[string]string) {
			parts["Word/_rels/Document.XML.rels"] = strings.Replace(parts["Word/_rels/Document.XML.rels"], relBaseStrict+"fontTable", relBaseTransitional+"fontTable", 1)
		}, "wrong Strict/Transitional"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			_, err := ExtractNativeDOCXFontInventoryV1(nativeFontTestPackage(t, test.strict, test.sfnt, test.subsetted, test.mutate))
			if err == nil || !strings.Contains(err.Error(), test.want) {
				t.Fatalf("error = %v, want %q", err, test.want)
			}
		})
	}
	aliasEntries := nativeFontTestEntries(t, false, valid, "false", nil)
	aliasEntries = append(aliasEntries, nativeZipEntry{name: "word/fonts/face.odttf", data: nativeFontObfuscate(valid), method: zip.Deflate})
	if _, err := ExtractNativeDOCXFontInventoryV1(buildNativeDOCX(t, aliasEntries)); err == nil || !strings.Contains(err.Error(), "ambiguous ZIP entries") {
		t.Fatalf("case-aliased font path was accepted: %v", err)
	}
	parts := nativeFontTestParts(false, "false")
	parts["Word/_rels/FontTable.XML.rels"] = strings.Replace(parts["Word/_rels/FontTable.XML.rels"], `Target="Fonts/Face.ODTTF"`, `Target="../%2E%2E/evil.odttf"`, 1)
	if _, err := ExtractNativeDOCXFontInventoryV1(nativeFontTestPackage(t, false, valid, "false", func(target map[string]string) {
		target["Word/_rels/FontTable.XML.rels"] = parts["Word/_rels/FontTable.XML.rels"]
	})); err == nil {
		t.Fatal("encoded traversal target was accepted")
	}
}

func TestNativeDOCXFontInventoryManifestTamperRevisionAndStrictJSON(t *testing.T) {
	source := nativeFontTestPackage(t, false, nativeFontTestSFNT(0), "false", nil)
	inventory, err := ExtractNativeDOCXFontInventoryV1(source)
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := EncodeNativeDOCXFontInventoryV1(inventory)
	if err != nil {
		t.Fatal(err)
	}
	unknown := bytes.Replace(encoded, []byte(`"protocol":`), []byte(`"unknown":true,"protocol":`), 1)
	if _, err := DecodeNativeDOCXFontInventoryV1(unknown); err == nil || !strings.Contains(err.Error(), "unknown field") {
		t.Fatalf("unknown field was accepted: %v", err)
	}
	duplicate := bytes.Replace(encoded, []byte(`"protocol":`), []byte(`"protocol":"injoffice.docx.font-inventory","protocol":`), 1)
	if _, err := DecodeNativeDOCXFontInventoryV1(duplicate); err == nil || !strings.Contains(err.Error(), "duplicated") {
		t.Fatalf("duplicate attestation field was accepted: %v", err)
	}
	incomplete := bytes.Replace(encoded, []byte(`,"relationship_target":"FontTable.XML"`), nil, 1)
	if _, err := DecodeNativeDOCXFontInventoryV1(incomplete); err == nil {
		t.Fatalf("incomplete relationship binding was accepted: %v", err)
	}
	missingFalse := bytes.Replace(encoded, []byte(`,"no_subsetting":false`), nil, 1)
	if _, err := DecodeNativeDOCXFontInventoryV1(missingFalse); err == nil || !strings.Contains(err.Error(), "canonical") {
		t.Fatalf("missing zero-valued attestation field was accepted: %v", err)
	}
	forged := *inventory
	forged.Families = append([]NativeDOCXFontFamilyV1(nil), inventory.Families...)
	forged.Families[0].Faces = append([]NativeDOCXFontFaceV1(nil), inventory.Families[0].Faces...)
	forged.Families[0].Faces[0].Source.StoredSHA256 = "sha256:" + strings.Repeat("a", 64)
	forged.InventorySHA256 = nativeDOCXFontInventoryDigest(&forged)
	if _, err := ResolveNativeDOCXFontAssetV1(source, &forged, NativeDOCXFontResolutionRequestV1{Family: "Fixture Sans", Weight: 400, Style: "normal", Stretch: 100}); err == nil {
		t.Fatal("forged signed path/digest binding was accepted")
	}
	stale := nativeFontTestPackage(t, false, nativeFontTestSFNT(0), "false", func(parts map[string]string) {
		parts["custom/unrelated.bin"] = "changed"
		parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/custom/unrelated.bin" ContentType="application/octet-stream"/></Types>`, 1)
	})
	_, err = ResolveNativeDOCXFontAssetV1(stale, inventory, NativeDOCXFontResolutionRequestV1{Family: "Fixture Sans", Weight: 400, Style: "normal", Stretch: 100})
	var resolutionErr *NativeDOCXFontResolutionError
	if !errors.As(err, &resolutionErr) || resolutionErr.Code != "REVISION_MISMATCH" {
		t.Fatalf("stale package error = %v, want REVISION_MISMATCH", err)
	}
	_, err = ResolveNativeDOCXPagePaintFontAssetsV1(stale, inventory)
	if !errors.As(err, &resolutionErr) || resolutionErr.Code != "REVISION_MISMATCH" {
		t.Fatalf("stale package worker-asset bridge error = %v, want REVISION_MISMATCH", err)
	}
}

func TestNativeDOCXFontInventoryPreservesPackagePartsAcrossMutation(t *testing.T) {
	source := nativeFontTestPackage(t, false, nativeFontTestSFNT(0), "false", nil)
	before, err := ExtractNativeDOCXFontInventoryV1(source)
	if err != nil {
		t.Fatal(err)
	}
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	run := doc.Body.Blocks[0].Paragraph.Runs[0]
	result, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "After"}})
	if err != nil {
		t.Fatal(err)
	}
	for _, part := range []string{"Word/FontTable.XML", "Word/_rels/FontTable.XML.rels", "Word/Fonts/Face.ODTTF"} {
		assertNativeRawPartPreserved(t, source, result.Package, part)
	}
	after, err := ExtractNativeDOCXFontInventoryV1WithOptions(result.Package, NativeExtractionOptions{Previous: doc})
	if err != nil {
		t.Fatal(err)
	}
	if before.PackageSHA256 == after.PackageSHA256 || before.Revision == after.Revision || before.Families[0].Faces[0].Source.ContentSHA256 != after.Families[0].Faces[0].Source.ContentSHA256 {
		t.Fatalf("package/font revision behavior is incorrect: before=%#v after=%#v", before, after)
	}
}

func nativeFontTestPackage(t *testing.T, strict bool, sfnt []byte, subsetted string, mutate func(map[string]string)) []byte {
	t.Helper()
	return buildNativeDOCX(t, nativeFontTestEntries(t, strict, sfnt, subsetted, mutate))
}

func nativeFontTestEntries(t *testing.T, strict bool, sfnt []byte, subsetted string, mutate func(map[string]string)) []nativeZipEntry {
	t.Helper()
	parts := nativeFontTestParts(strict, subsetted)
	if mutate != nil {
		mutate(parts)
	}
	font := nativeFontObfuscate(sfnt)
	entries := nativeEntries(parts)
	for index := range entries {
		if entries[index].name == "Word/Fonts/Face.ODTTF" {
			entries[index].data = font
		}
	}
	return entries
}

func nativeFontTestParts(strict bool, subsetted string) map[string]string {
	w, r, base := wordMLTransitional, relNSTransitional, relBaseTransitional
	if strict {
		w, r, base = wordMLStrict, relNSStrict, relBaseStrict
	}
	return map[string]string{
		"[Content_Types].xml":           `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/Word/Document.XML" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/Word/FontTable.XML" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Override PartName="/Word/Fonts/Face.ODTTF" ContentType="` + nativeDOCXObfuscatedFontType + `"/></Types>`,
		"_rels/.rels":                   `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + base + `officeDocument" Target="Word/Document.XML"/></Relationships>`,
		"Word/Document.XML":             `<w:document xmlns:w="` + w + `"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Fixture Sans" w:hAnsi="Fixture Sans"/></w:rPr><w:t>Before</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		"Word/_rels/Document.XML.rels":  `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="fontTable" Type="` + base + `fontTable" Target="FontTable.XML"/></Relationships>`,
		"Word/FontTable.XML":            `<w:fonts xmlns:w="` + w + `" xmlns:r="` + r + `"><w:font w:name="Fixture Sans"><w:altName w:val="Fixture Alias"/><w:embedRegular r:id="fontRegular" w:fontKey="` + nativeFontTestKey + `" w:subsetted="` + subsetted + `"/></w:font></w:fonts>`,
		"Word/_rels/FontTable.XML.rels": `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="fontRegular" Type="` + base + `font" Target="Fonts/Face.ODTTF"/></Relationships>`,
		"Word/Fonts/Face.ODTTF":         "placeholder-font-bytes-overwritten",
	}
}

func nativeFontTestSFNT(fsType uint16) []byte {
	data := make([]byte, 38)
	copy(data[:4], []byte{0, 1, 0, 0})
	data[5] = 1
	copy(data[12:16], "OS/2")
	data[23] = 28
	data[27] = 10
	data[36], data[37] = byte(fsType>>8), byte(fsType)
	return data
}

func nativeFontTestMalformedRange() []byte {
	data := nativeFontTestSFNT(0)
	data[23] = 250
	return data
}

func nativeFontObfuscate(sfnt []byte) []byte {
	if len(sfnt) < 32 {
		return append([]byte(nil), sfnt...)
	}
	key, _ := hexDecodeNativeFontTestKey()
	stored := append([]byte(nil), sfnt...)
	for index := 0; index < 32; index++ {
		stored[index] ^= key[15-index%16]
	}
	return stored
}

func hexDecodeNativeFontTestKey() ([]byte, error) {
	return hex.DecodeString(strings.ReplaceAll(strings.Trim(nativeFontTestKey, "{}"), "-", ""))
}

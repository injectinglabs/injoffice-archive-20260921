package pptxpatch

import (
	"encoding/base64"
	"fmt"
	"io"
	"mime"
	"net/url"
	"reflect"
	"regexp"
	"strings"
	"unicode"
	"unicode/utf8"
)

const nativeMaxSafeInteger int64 = 9007199254740991

var nativeLanguagePattern = regexp.MustCompile(`^[A-Za-z]{2,8}(-[A-Za-z0-9]{1,8})*$`)

func validNativeLanguage(value string) bool {
	return len(value) <= 63 && nativeLanguagePattern.MatchString(value)
}

func containsNativeString(values []string, value string) bool {
	for _, candidate := range values {
		if candidate == value {
			return true
		}
	}
	return false
}

var (
	nativeIDPattern      = regexp.MustCompile(`^[A-Za-z0-9][A-Za-z0-9._~:/-]{0,255}$`)
	sha256Pattern        = regexp.MustCompile(`^[0-9a-f]{64}$`)
	colorPattern         = regexp.MustCompile(`^[0-9A-F]{6}$`)
	diagnosticPattern    = regexp.MustCompile(`^[a-z][a-z0-9]*(?:[.-][a-z0-9]+)*$`)
	partCharacterPattern = regexp.MustCompile(`^(?:[A-Za-z0-9._~!$&'()*+,;=:@/-]|%[0-9A-F]{2})+$`)
)

// NativeContractIssue is stable, JSON-safe validation output. Paths use a
// JSONPath-like notation shared with @injoffice/pptx-native.
type NativeContractIssue struct {
	Path    string `json:"path"`
	Code    string `json:"code"`
	Message string `json:"message"`
}

// ValidateNativePPTX validates both the schema-level and semantic invariants
// that JSON Schema cannot express: durable-id uniqueness, reference integrity,
// compatibility aggregation, source-revision rules, and strict union fields.
// It returns at most 100 deterministic issues.
func ValidateNativePPTX(deck NativePPTXDeck) []NativeContractIssue {
	valueBudget := nativeGoValueBudget{}
	if issue := validateNativeGoValueBudget(reflect.ValueOf(deck), 0, &valueBudget); issue != nil {
		return []NativeContractIssue{*issue}
	}
	v := nativeValidator{
		ids:      map[string]bool{},
		assets:   map[string]NativeAsset{},
		slides:   map[string]bool{},
		elements: map[string]string{},
	}
	if deck.ContractVersion != NativePPTXContractVersion {
		v.add("$.contractVersion", "schema.const", fmt.Sprintf("must equal %q", NativePPTXContractVersion))
	}
	v.id(deck.DocumentID, "$.documentId")
	if deck.Origin != NativeOriginAuthored && deck.Origin != NativeOriginParsed {
		v.add("$.origin", "schema.enum", "must be authored or parsed")
	}
	if deck.Origin == NativeOriginParsed && deck.SourceRevision == nil {
		v.add("$.sourceRevision", "native.sourceRevision", "is required for a parsed deck")
	}
	if deck.Origin == NativeOriginAuthored && deck.SourceRevision != nil {
		v.add("$.sourceRevision", "native.sourceRevision", "is not allowed for an authored deck")
	}
	if deck.SourceRevision != nil && !nativeIDPattern.MatchString(*deck.SourceRevision) {
		v.add("$.sourceRevision", "schema.pattern", "has an invalid durable-id format")
	}
	v.requiredPositive(deck.Size.Cx, "$.size.cx")
	v.requiredPositive(deck.Size.Cy, "$.size.cy")
	if deck.Assets == nil {
		v.add("$.assets", "schema.required", "must be an array")
	}
	if deck.Slides == nil {
		v.add("$.slides", "schema.required", "must be an array")
	}
	if len(deck.Assets) > nativeMaxAssets {
		v.add("$.assets", "schema.maxItems", fmt.Sprintf("must contain at most %d assets", nativeMaxAssets))
	}
	if len(deck.Slides) > nativeMaxSlides {
		v.add("$.slides", "schema.maxItems", fmt.Sprintf("must contain at most %d slides", nativeMaxSlides))
	}
	v.compatibility(deck.Compatibility, "$.compatibility")

	for index, asset := range deck.Assets[:boundedLength(len(deck.Assets), nativeMaxAssets)] {
		p := fmt.Sprintf("$.assets[%d]", index)
		v.id(asset.ID, p+".id")
		v.assets[asset.ID] = asset
		v.provenance(deck.Origin, asset.Provenance, asset.Source, asset.Passthrough, p)
		if len(asset.ContentType) > 256 {
			v.add(p+".contentType", "schema.maxLength", "must contain at most 256 characters")
		} else if mediaType, _, err := mime.ParseMediaType(asset.ContentType); err != nil || mediaType != strings.ToLower(asset.ContentType) {
			v.add(p+".contentType", "schema.pattern", "must be a lowercase MIME content type without parameters")
		}
		v.hash(asset.SHA256, p+".sha256")
		v.requiredBoundedNonnegative(asset.ByteLength, nativeMaxAssetBytes, p+".byteLength")
		if asset.Passthrough == nil {
			v.add(p+".passthrough", "schema.required", "must be an array")
		}
		v.sourceState(asset.Source, asset.Passthrough, p)
		if asset.Provenance == NativeProvenanceParsed && asset.DataBase64 == nil {
			readCapability := false
			if asset.Source != nil {
				for _, ref := range asset.Passthrough {
					if ref.OwnerPart == asset.Source.PartName && ref.FingerprintSHA256 == asset.SHA256 {
						readCapability = true
						break
					}
				}
			}
			if !readCapability {
				v.add(p+".passthrough", "native.assetReadCapability", "parsed source-only assets require a capability bound to the source part and asset digest")
			}
		}
		if asset.DataBase64 != nil {
			v.inlineAssetBase64CodeUnits += int64(len(*asset.DataBase64))
			if len(*asset.DataBase64) > nativeMaxInlineAssetBase64CodeUnits {
				v.add(p+".dataBase64", "schema.maxLength", fmt.Sprintf("must contain at most %d code units", nativeMaxInlineAssetBase64CodeUnits))
				continue
			}
			if v.inlineAssetBase64CodeUnits > nativeMaxTotalInlineAssetBase64CodeUnits {
				continue
			}
			decoded, err := io.Copy(io.Discard, base64.NewDecoder(base64.StdEncoding.Strict(), strings.NewReader(*asset.DataBase64)))
			if err != nil {
				v.add(p+".dataBase64", "schema.contentEncoding", "must be padded standard base64")
			} else if asset.ByteLength != nil {
				if decoded != *asset.ByteLength {
					v.add(p+".dataBase64", "native.assetLength", "decoded length does not match byteLength")
				}
			}
		}
	}
	if v.inlineAssetBase64CodeUnits > nativeMaxTotalInlineAssetBase64CodeUnits {
		v.add("$.assets", "native.resourceBudget", fmt.Sprintf("inline asset data exceeds %d code units", nativeMaxTotalInlineAssetBase64CodeUnits))
	}

	worst := NativeCompatibilityStatusEditable
	for slideIndex, slide := range deck.Slides[:boundedLength(len(deck.Slides), nativeMaxSlides)] {
		p := fmt.Sprintf("$.slides[%d]", slideIndex)
		v.id(slide.ID, p+".id")
		v.slides[slide.ID] = true
		if slide.Background != nil {
			v.color(*slide.Background, p+".background")
		}
		if slide.Elements == nil {
			v.add(p+".elements", "schema.required", "must be an array")
		}
		if slide.Passthrough == nil {
			v.add(p+".passthrough", "schema.required", "must be an array")
		}
		v.provenance(deck.Origin, slide.Provenance, slide.Source, slide.Passthrough, p)
		v.sourceState(slide.Source, slide.Passthrough, p)
		v.compatibility(slide.Compatibility, p+".compatibility")
		slideWorst := slide.Compatibility.Status
		if slide.Transition != nil {
			v.transition(*slide.Transition, p+".transition")
		}
		if len(slide.Elements) > nativeMaxElementsPerContainer {
			v.add(p+".elements", "schema.maxItems", fmt.Sprintf("must contain at most %d elements", nativeMaxElementsPerContainer))
		}
		slidePart := ""
		if slide.Source != nil {
			slidePart = slide.Source.PartName
		}
		for elementIndex, element := range slide.Elements[:boundedLength(len(slide.Elements), nativeMaxElementsPerContainer)] {
			elementWorst := v.element(deck.Origin, element, fmt.Sprintf("%s.elements[%d]", p, elementIndex), 1, slide.ID, slidePart)
			slideWorst = worseNativeStatus(slideWorst, elementWorst)
		}
		if nativeStatusRank(slide.Compatibility.Status) < nativeStatusRank(slideWorst) {
			v.add(p+".compatibility.status", "native.compatibilityAggregate", fmt.Sprintf("must be at least %s because an element has that status", slideWorst))
		}
		worst = worseNativeStatus(worst, slideWorst)
	}
	if nativeStatusRank(deck.Compatibility.Status) < nativeStatusRank(worst) {
		v.add("$.compatibility.status", "native.compatibilityAggregate", fmt.Sprintf("must be at least %s because a descendant has that status", worst))
	}
	if v.textCodeUnits > nativeMaxTotalTextCodeUnits {
		v.add("$", "native.resourceBudget", fmt.Sprintf("text exceeds %d code units", nativeMaxTotalTextCodeUnits))
	}
	if v.tableCells > nativeMaxTableCells {
		v.add("$", "native.resourceBudget", fmt.Sprintf("tables exceed %d cells", nativeMaxTableCells))
	}

	v.validateDiagnosticReferences(deck)
	return v.issues
}

type nativeGoValueBudget struct{ nodes int }

func validateNativeGoValueBudget(value reflect.Value, depth int, budget *nativeGoValueBudget) *NativeContractIssue {
	for value.IsValid() && (value.Kind() == reflect.Pointer || value.Kind() == reflect.Interface) {
		if value.IsNil() {
			break
		}
		value = value.Elem()
	}
	if depth > nativeMaxDepth {
		return &NativeContractIssue{Path: "$", Code: "native.resourceDepth", Message: fmt.Sprintf("contract nesting exceeds %d", nativeMaxDepth)}
	}
	budget.nodes++
	if budget.nodes > nativeMaxNodes {
		return &NativeContractIssue{Path: "$", Code: "native.resourceBudget", Message: fmt.Sprintf("contract exceeds %d JSON nodes", nativeMaxNodes)}
	}
	if !value.IsValid() || (value.Kind() == reflect.Pointer && value.IsNil()) || (value.Kind() == reflect.Interface && value.IsNil()) {
		return nil
	}
	switch value.Kind() {
	case reflect.Struct:
		valueType := value.Type()
		for index := 0; index < value.NumField(); index++ {
			field := value.Field(index)
			tag := strings.Split(valueType.Field(index).Tag.Get("json"), ",")
			if tag[0] == "-" || tag[0] == "" {
				continue
			}
			if len(tag) > 1 && tag[1] == "omitempty" && nativeJSONEmptyValue(field) {
				continue
			}
			if issue := validateNativeGoValueBudget(field, depth+1, budget); issue != nil {
				return issue
			}
		}
	case reflect.Slice, reflect.Array:
		for index := 0; index < value.Len(); index++ {
			if issue := validateNativeGoValueBudget(value.Index(index), depth+1, budget); issue != nil {
				return issue
			}
		}
	}
	return nil
}

func nativeJSONEmptyValue(value reflect.Value) bool {
	switch value.Kind() {
	case reflect.Array, reflect.Map, reflect.Slice, reflect.String:
		return value.Len() == 0
	case reflect.Bool, reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64,
		reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr,
		reflect.Float32, reflect.Float64, reflect.Interface, reflect.Pointer:
		return value.IsZero()
	}
	return false
}

type nativeValidator struct {
	issues                     []NativeContractIssue
	ids                        map[string]bool
	assets                     map[string]NativeAsset
	slides                     map[string]bool
	elements                   map[string]string
	elementCount               int
	textCodeUnits              int64
	tableCells                 int
	inlineAssetBase64CodeUnits int64
}

func (v *nativeValidator) element(origin NativeOrigin, element NativeElement, p string, depth int, slideID, slidePart string) NativeCompatibilityStatus {
	v.elementCount++
	if v.elementCount > nativeMaxTotalElements {
		v.add(p, "native.resourceBudget", fmt.Sprintf("deck exceeds %d elements", nativeMaxTotalElements))
		return NativeCompatibilityStatusRefused
	}
	if depth > nativeMaxDepth {
		v.add(p, "native.resourceDepth", fmt.Sprintf("group depth exceeds %d", nativeMaxDepth))
		return NativeCompatibilityStatusRefused
	}
	v.id(element.ID, p+".id")
	v.elements[element.ID] = slideID
	v.transform(element.Transform, p+".transform")
	if element.Transform.QuarterTurns != nil && element.Kind != NativeElementKindText && element.Kind != NativeElementKindShape {
		v.add(p+".transform.quarterTurns", "native.rotation", "quarter turns are supported only for text and shapes")
	}
	if element.ChildTransform != nil && element.ChildTransform.QuarterTurns != nil {
		v.add(p+".childTransform.quarterTurns", "native.rotation", "child coordinate systems cannot carry quarter turns")
	}
	if element.Name != nil && utf16CodeUnitLengthBounded(*element.Name, 1024) > 1024 {
		v.add(p+".name", "schema.maxLength", "must contain at most 1024 characters")
	}
	if element.Passthrough == nil {
		v.add(p+".passthrough", "schema.required", "must be an array")
	}
	v.provenance(origin, element.Provenance, element.Source, element.Passthrough, p)
	v.sourceState(element.Source, element.Passthrough, p)
	if element.TextBody != nil && element.TextBody.AutoFit == "shape-source-frame" {
		warning := false
		for _, diagnostic := range element.Compatibility.Diagnostics {
			if diagnostic.Code == "pptx.autofit-source-frame-approximate" && diagnostic.Severity == NativeDiagnosticSeverityWarning {
				warning = true
			}
		}
		if element.Provenance != NativeProvenanceParsed || element.Source == nil || element.Compatibility.Status == NativeCompatibilityStatusEditable || !warning {
			v.add(p+".textBody.autoFit", "native.autofitApproximation", "source-frame autofit requires a parsed source, non-editable status and explicit approximation warning")
		}
	}
	if element.TextBody != nil && element.TextBody.WritingMode != nil && element.Provenance == NativeProvenanceParsed && element.Compatibility.Status == NativeCompatibilityStatusEditable {
		v.add(p+".textBody.writingMode", "native.verticalPreview", "parsed vertical text must remain read-only")
	}
	if element.Provenance == NativeProvenanceParsed && slidePart == "" {
		v.add(p+".provenance", "native.sourceOwnership", "parsed elements require an owning parsed slide with a source anchor")
	} else if element.Provenance == NativeProvenanceParsed && element.Source != nil && element.Source.PartName != slidePart {
		v.add(p+".source.partName", "native.sourcePart", "parsed elements must be anchored to their owning slide part")
	}
	v.compatibility(element.Compatibility, p+".compatibility")
	if element.Animation != nil {
		v.animation(*element.Animation, p+".animation")
	}
	if element.Kind != NativeElementKindText && element.Kind != NativeElementKindShape && element.TextBody != nil {
		v.add(p+".textBody", "native.elementUnion", "is not allowed for this element kind")
	}

	commonForbidden := func(paragraphs, preset, placeholder, paint, connector, asset, table, chart, childTransform, children bool) {
		if !paragraphs && element.Paragraphs != nil {
			v.add(p+".paragraphs", "native.elementUnion", "is not allowed for this element kind")
		}
		if !preset && element.Preset != nil {
			v.add(p+".preset", "native.elementUnion", "is not allowed for this element kind")
		}
		if !placeholder && element.Placeholder != nil {
			v.add(p+".placeholder", "native.elementUnion", "is not allowed for this element kind")
		}
		if !paint && (element.Fill != nil || element.Stroke != nil) {
			v.add(p+".fill", "native.elementUnion", "fill or stroke is not allowed for this element kind")
		}
		if !connector && (element.HeadArrow != nil || element.TailArrow != nil || element.FlipH != nil || element.HeadEnd != nil || element.TailEnd != nil) {
			v.add(p+".headArrow", "native.elementUnion", "connector flags are not allowed for this element kind")
		}
		if !asset && element.AssetID != nil {
			v.add(p+".assetId", "native.elementUnion", "is not allowed for this element kind")
		}
		if !asset && element.Crop != nil {
			v.add(p+".crop", "native.elementUnion", "is not allowed for this element kind")
		}
		if !asset && element.Clip != nil {
			v.add(p+".clip", "native.elementUnion", "is not allowed for this element kind")
		}
		if !table && element.Table != nil {
			v.add(p+".table", "native.elementUnion", "is not allowed for this element kind")
		}
		if !chart && element.Chart != nil {
			v.add(p+".chart", "native.elementUnion", "is not allowed for this element kind")
		}
		if !childTransform && element.ChildTransform != nil {
			v.add(p+".childTransform", "native.elementUnion", "is not allowed for this element kind")
		}
		if !children && element.Children != nil {
			v.add(p+".children", "native.elementUnion", "is not allowed for this element kind")
		}
	}

	switch element.Kind {
	case NativeElementKindText:
		commonForbidden(true, false, true, false, false, false, false, false, false, false)
		if element.Paragraphs == nil {
			v.add(p+".paragraphs", "schema.required", "must be an array")
		} else {
			v.paragraphs(*element.Paragraphs, p+".paragraphs")
		}
		v.placeholder(element.Placeholder, p+".placeholder")
		if element.TextBody != nil {
			v.textBody(*element.TextBody, element.Transform, p+".textBody")
		}
	case NativeElementKindShape:
		commonForbidden(true, true, true, true, false, false, false, false, false, false)
		if element.Preset == nil {
			if element.Compatibility.Status != NativeCompatibilityStatusRefused {
				v.add(p+".preset", "native.shapePreset", "is required unless the shape is explicitly refused")
			}
		} else {
			v.shapePreset(element.Preset, p+".preset")
		}
		v.placeholder(element.Placeholder, p+".placeholder")
		if element.Paragraphs == nil {
			v.add(p+".paragraphs", "schema.required", "must be an array")
		} else {
			v.paragraphs(*element.Paragraphs, p+".paragraphs")
		}
		if element.Fill != nil {
			v.color(*element.Fill, p+".fill")
		}
		if element.Stroke != nil {
			v.stroke(*element.Stroke, p+".stroke")
		}
		if element.TextBody != nil {
			v.textBody(*element.TextBody, element.Transform, p+".textBody")
		}
	case NativeElementKindConnector:
		commonForbidden(false, false, false, true, true, false, false, false, false, false)
		for _, end := range []struct {
			name  string
			value *NativeArrowEnd
			flag  *bool
		}{{"headEnd", element.HeadEnd, element.HeadArrow}, {"tailEnd", element.TailEnd, element.TailArrow}} {
			if end.value == nil {
				continue
			}
			if !containsNativeString([]string{"none", "triangle", "arrow", "stealth", "diamond", "oval"}, end.value.Type) {
				v.add(p+"."+end.name+".type", "schema.enum", "unknown arrow type")
			}
			for _, size := range []*string{end.value.W, end.value.Len} {
				if size != nil && !containsNativeString([]string{"sm", "med", "lg"}, *size) {
					v.add(p+"."+end.name, "schema.enum", "unknown arrow size")
				}
			}
			if end.flag != nil && *end.flag != (end.value.Type != "none") {
				v.add(p+"."+end.name, "native.arrowPresence", "typed endpoint conflicts with legacy presence flag")
			}
		}
		if element.Stroke != nil {
			v.stroke(*element.Stroke, p+".stroke")
		}
	case NativeElementKindPicture:
		commonForbidden(false, false, false, false, false, true, false, false, false, false)
		if element.Clip != nil && *element.Clip != "roundRect" {
			v.add(p+".clip", "schema.const", "only the default roundRect picture clip is supported")
		}
		if crop := element.Crop; crop != nil {
			valid := true
			for _, side := range []struct {
				name  string
				value *int64
			}{{"left", crop.Left}, {"top", crop.Top}, {"right", crop.Right}, {"bottom", crop.Bottom}} {
				if side.value == nil || *side.value < 0 || *side.value >= 100_000 {
					v.add(p+".crop."+side.name, "native.pictureCrop", "requires an integer inset from 0 through 99999")
					valid = false
				}
			}
			if valid && (*crop.Left+*crop.Right >= 100_000 || *crop.Top+*crop.Bottom >= 100_000) {
				v.add(p+".crop", "native.pictureCrop", "opposing crop insets must leave a positive source rectangle")
			}
		}
		if element.AssetID == nil {
			v.add(p+".assetId", "schema.required", "is required")
		} else if asset, ok := v.assets[*element.AssetID]; !ok {
			v.add(p+".assetId", "native.assetReference", "references an unknown asset id")
		} else if !strings.HasPrefix(asset.ContentType, "image/") {
			v.add(p+".assetId", "native.assetType", "picture assets must have an image content type")
		}
	case NativeElementKindTable:
		commonForbidden(false, false, false, false, false, false, true, false, false, false)
		if element.Table == nil {
			v.add(p+".table", "schema.required", "is required")
		} else {
			v.table(*element.Table, element.Transform, p+".table")
		}
	case NativeElementKindChart:
		commonForbidden(false, false, false, false, false, false, false, true, false, false)
		if element.Chart == nil {
			v.add(p+".chart", "schema.required", "is required")
		} else {
			v.chart(*element.Chart, p+".chart")
		}
		if element.Source == nil {
			v.add(p+".source", "schema.required", "opaque charts require a source anchor")
		} else if element.Chart != nil && (element.Source.RelationshipID == nil || *element.Source.RelationshipID != element.Chart.RelationshipID) {
			v.add(p+".chart.relationshipId", "native.chartReference", "must equal source.relationshipId")
		}
		if element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
			v.add(p+".compatibility.status", "native.opaqueChart", "opaque charts must be preserveOnly")
		}
	case NativeElementKindGroup:
		commonForbidden(false, false, false, false, false, false, false, false, true, true)
		if element.Provenance == NativeProvenanceParsed && element.ChildTransform == nil {
			v.add(p+".childTransform", "native.groupTransform", "parsed groups require an authoritative DrawingML child coordinate transform")
		}
		if element.ChildTransform != nil {
			v.transform(*element.ChildTransform, p+".childTransform")
			if _, _, _, _, err := nativeGroupAffineComponents(element.Transform, *element.ChildTransform); err != nil {
				v.add(p+".childTransform", "native.groupTransform", "must compose to exact safe integer-PPM scale and integer-EMU translation")
			}
		}
		if len(element.Children) == 0 {
			v.add(p+".children", "schema.minItems", "must contain at least one child")
		}
		if len(element.Children) > nativeMaxElementsPerContainer {
			v.add(p+".children", "schema.maxItems", fmt.Sprintf("must contain at most %d children", nativeMaxElementsPerContainer))
		}
		worst := element.Compatibility.Status
		for index, child := range element.Children[:boundedLength(len(element.Children), nativeMaxElementsPerContainer)] {
			worst = worseNativeStatus(worst, v.element(origin, child, fmt.Sprintf("%s.children[%d]", p, index), depth+1, slideID, slidePart))
		}
		if nativeStatusRank(element.Compatibility.Status) < nativeStatusRank(worst) {
			v.add(p+".compatibility.status", "native.compatibilityAggregate", fmt.Sprintf("must be at least %s because a child has that status", worst))
		}
		return worst
	default:
		v.add(p+".kind", "schema.enum", "has an unsupported element kind")
	}
	return element.Compatibility.Status
}

func (v *nativeValidator) textBody(body NativeTextBodyLayout, transform NativeTransform, p string) {
	for _, field := range []struct {
		name  string
		value *int64
	}{
		{name: "leftInsetEmu", value: body.LeftInsetEMU},
		{name: "rightInsetEmu", value: body.RightInsetEMU},
		{name: "topInsetEmu", value: body.TopInsetEMU},
		{name: "bottomInsetEmu", value: body.BottomInsetEMU},
	} {
		if field.value == nil {
			v.add(p+"."+field.name, "schema.required", "is required")
		} else if *field.value < 0 || *field.value > 2147483647 {
			v.add(p+"."+field.name, "schema.range", "must be a nonnegative signed 32-bit DrawingML coordinate")
		}
	}
	if body.Wrap != NativeTextWrapSquare && body.Wrap != NativeTextWrapNone {
		v.add(p+".wrap", "schema.enum", "must be square or none")
	}
	if body.VerticalAnchor != NativeTextVerticalAnchorTop && body.VerticalAnchor != NativeTextVerticalAnchorCenter && body.VerticalAnchor != NativeTextVerticalAnchorBottom {
		v.add(p+".verticalAnchor", "schema.enum", "must be top, center, or bottom")
	}
	if body.AutoFit != "none" && body.AutoFit != "shape-source-frame" {
		v.add(p+".autoFit", "schema.enum", "must be none or shape-source-frame")
	}
	if body.WritingMode != nil && *body.WritingMode != "vertical-clockwise" {
		v.add(p+".writingMode", "schema.enum", "must equal vertical-clockwise")
	}
	if body.HorizontalOverflow != "overflow" {
		v.add(p+".horizontalOverflow", "schema.const", "must equal overflow")
	}
	if body.VerticalOverflow != "overflow" {
		v.add(p+".verticalOverflow", "schema.const", "must equal overflow")
	}
	if transform.Cx != nil && body.LeftInsetEMU != nil && body.RightInsetEMU != nil {
		width := *transform.Cx - *body.LeftInsetEMU - *body.RightInsetEMU
		if width <= 0 {
			v.add(p, "native.textBodyBounds", "horizontal insets must leave a positive text-body width")
		}
	}
	if transform.Cy != nil && body.TopInsetEMU != nil && body.BottomInsetEMU != nil {
		height := *transform.Cy - *body.TopInsetEMU - *body.BottomInsetEMU
		if height <= 0 {
			v.add(p, "native.textBodyBounds", "vertical insets must leave a positive text-body height")
		}
	}
}

func (v *nativeValidator) provenance(deckOrigin NativeOrigin, provenance NativeProvenance, source *NativeSourceAnchor, passthrough []NativePassthroughRef, p string) {
	if deckOrigin == NativeOriginAuthored && provenance != NativeProvenanceAuthored {
		v.add(p+".provenance", "native.provenance", "objects in an authored deck must be authored")
	}
	if provenance == NativeProvenanceAuthored && (source != nil || len(passthrough) != 0) {
		v.add(p, "native.authoredSource", "authored objects cannot claim source anchors or passthrough tokens")
	}
	if provenance == NativeProvenanceParsed && deckOrigin != NativeOriginParsed {
		v.add(p+".provenance", "native.provenance", "parsed objects require a parsed deck")
	}
	if provenance == NativeProvenanceParsed && source == nil {
		v.add(p+".source", "native.parsedSource", "parsed objects require a source anchor")
	}
	if provenance != NativeProvenanceAuthored && provenance != NativeProvenanceParsed {
		v.add(p+".provenance", "schema.enum", "must be authored or parsed")
	}
}

func (v *nativeValidator) sourceState(source *NativeSourceAnchor, passthrough []NativePassthroughRef, p string) {
	if source != nil {
		v.source(*source, p+".source")
	}
	if len(passthrough) > nativeMaxPassthroughPerObject {
		v.add(p+".passthrough", "schema.maxItems", fmt.Sprintf("must contain at most %d entries", nativeMaxPassthroughPerObject))
	}
	for index, ref := range passthrough[:boundedLength(len(passthrough), nativeMaxPassthroughPerObject)] {
		v.passthrough(ref, fmt.Sprintf("%s.passthrough[%d]", p, index))
	}
}

func (v *nativeValidator) source(anchor NativeSourceAnchor, p string) {
	v.part(anchor.PartName, p+".partName")
	if !nativeIDPattern.MatchString(anchor.ObjectID) {
		v.add(p+".objectId", "schema.pattern", "has an invalid durable-id format")
	}
	if anchor.RelationshipID != nil && !nativeIDPattern.MatchString(*anchor.RelationshipID) {
		v.add(p+".relationshipId", "schema.pattern", "has an invalid durable-id format")
	}
	v.hash(anchor.FingerprintSHA256, p+".fingerprintSha256")
}

func (v *nativeValidator) passthrough(ref NativePassthroughRef, p string) {
	if !nativeIDPattern.MatchString(ref.Token) {
		v.add(p+".token", "schema.pattern", "has an invalid opaque-token format")
	}
	v.part(ref.OwnerPart, p+".ownerPart")
	v.hash(ref.FingerprintSHA256, p+".fingerprintSha256")
	if ref.Disposition != NativePassthroughDispositionPreserve && ref.Disposition != NativePassthroughDispositionReplaceOnEdit {
		v.add(p+".disposition", "schema.enum", "has an unsupported passthrough disposition")
	}
}

func (v *nativeValidator) compatibility(value NativeCompatibility, p string) {
	if value.Diagnostics == nil {
		v.add(p+".diagnostics", "schema.required", "must be an array")
	}
	if len(value.Diagnostics) > nativeMaxDiagnosticsPerScope {
		v.add(p+".diagnostics", "schema.maxItems", fmt.Sprintf("must contain at most %d diagnostics", nativeMaxDiagnosticsPerScope))
	}
	refusals, warnings := 0, 0
	for index, diagnostic := range value.Diagnostics[:boundedLength(len(value.Diagnostics), nativeMaxDiagnosticsPerScope)] {
		dp := fmt.Sprintf("%s.diagnostics[%d]", p, index)
		switch diagnostic.Severity {
		case NativeDiagnosticSeverityInfo:
		case NativeDiagnosticSeverityWarning:
			warnings++
		case NativeDiagnosticSeverityRefusal:
			refusals++
		default:
			v.add(dp+".severity", "schema.enum", "has an unsupported diagnostic severity")
		}
		if len(diagnostic.Code) == 0 || len(diagnostic.Code) > 128 || !diagnosticPattern.MatchString(diagnostic.Code) {
			v.add(dp+".code", "schema.pattern", "has an invalid namespaced diagnostic code")
		}
		messageLength := utf16CodeUnitLengthBounded(diagnostic.Message, 2048)
		if messageLength == 0 || messageLength > 2048 {
			v.add(dp+".message", "schema.length", "must contain 1 to 2048 characters")
		}
		if diagnostic.Scope != nil {
			scope := diagnostic.Scope
			if scope.SlideID == nil && scope.ElementID == nil && scope.PartName == nil {
				v.add(dp+".scope", "schema.minProperties", "must identify at least one scope")
			}
			if scope.SlideID != nil && !nativeIDPattern.MatchString(*scope.SlideID) {
				v.add(dp+".scope.slideId", "schema.pattern", "has an invalid durable-id format")
			}
			if scope.ElementID != nil && !nativeIDPattern.MatchString(*scope.ElementID) {
				v.add(dp+".scope.elementId", "schema.pattern", "has an invalid durable-id format")
			}
			if scope.PartName != nil {
				v.part(*scope.PartName, dp+".scope.partName")
			}
		}
	}
	switch value.Status {
	case NativeCompatibilityStatusEditable:
		if refusals != 0 {
			v.add(p, "native.compatibility", "editable content cannot contain refusal diagnostics")
		}
	case NativeCompatibilityStatusPreserveOnly:
		if warnings == 0 {
			v.add(p, "native.compatibility", "preserveOnly content requires a warning diagnostic")
		}
		if refusals != 0 {
			v.add(p, "native.compatibility", "preserveOnly content cannot contain refusal diagnostics")
		}
	case NativeCompatibilityStatusRefused:
		if refusals == 0 {
			v.add(p, "native.compatibility", "refused content requires a refusal diagnostic")
		}
	default:
		v.add(p+".status", "schema.enum", "has an unsupported compatibility status")
	}
}

func (v *nativeValidator) paragraphs(paragraphs []NativeParagraph, p string) {
	if len(paragraphs) > nativeMaxParagraphsPerElement {
		v.add(p, "schema.maxItems", fmt.Sprintf("must contain at most %d paragraphs", nativeMaxParagraphsPerElement))
	}
	for paragraphIndex, paragraph := range paragraphs[:boundedLength(len(paragraphs), nativeMaxParagraphsPerElement)] {
		pp := fmt.Sprintf("%s[%d]", p, paragraphIndex)
		if paragraph.Runs == nil {
			v.add(pp+".runs", "schema.required", "must be an array")
		}
		if paragraph.Align != nil && *paragraph.Align != NativeTextAlignLeft && *paragraph.Align != NativeTextAlignCenter && *paragraph.Align != NativeTextAlignRight {
			v.add(pp+".align", "schema.enum", "has an unsupported text alignment")
		}
		if paragraph.Level != nil && (*paragraph.Level < 0 || *paragraph.Level > 8) {
			v.add(pp+".level", "schema.range", "must be between 0 and 8")
		}
		if paragraph.BulletCharacter != nil {
			if paragraph.Bullet == nil || !*paragraph.Bullet || utf8.RuneCountInString(*paragraph.BulletCharacter) != 1 || strings.IndexFunc(*paragraph.BulletCharacter, unicode.IsControl) >= 0 {
				v.add(pp+".bulletCharacter", "native.bulletCharacter", "requires one authored character and bullet=true")
			}
		}
		if paragraph.MarginLeftEmu != nil && (*paragraph.MarginLeftEmu < 0 || *paragraph.MarginLeftEmu > 51206400) {
			v.add(pp+".marginLeftEmu", "schema.range", "invalid paragraph margin")
		}
		if paragraph.IndentEmu != nil && (*paragraph.IndentEmu < -51206400 || *paragraph.IndentEmu > 51206400) {
			v.add(pp+".indentEmu", "schema.range", "invalid paragraph indent")
		}
		if len(paragraph.Runs) > nativeMaxRunsPerParagraph {
			v.add(pp+".runs", "schema.maxItems", fmt.Sprintf("must contain at most %d runs", nativeMaxRunsPerParagraph))
		}
		for runIndex, run := range paragraph.Runs[:boundedLength(len(paragraph.Runs), nativeMaxRunsPerParagraph)] {
			rp := fmt.Sprintf("%s.runs[%d]", pp, runIndex)
			if run.Text == nil {
				v.add(rp+".text", "schema.required", "is required")
			} else {
				length := v.addTextCodeUnits(*run.Text)
				if length > nativeMaxTextCodeUnits {
					v.add(rp+".text", "schema.maxLength", fmt.Sprintf("must contain at most %d code units", nativeMaxTextCodeUnits))
				}
			}
			if run.FontSizeHundredthPt != nil && (*run.FontSizeHundredthPt < 1 || *run.FontSizeHundredthPt > 400000) {
				v.add(rp+".fontSizeHundredthPt", "schema.range", "must be between 1 and 400000")
			}
			if run.Color != nil {
				v.color(*run.Color, rp+".color")
			}
			if run.Language != nil && !validNativeLanguage(*run.Language) {
				v.add(rp+".language", "schema.pattern", "must be a bounded supported language tag")
			}
			if run.FontFamily != nil {
				fontFamilyLength := utf16CodeUnitLengthBounded(*run.FontFamily, 256)
				if fontFamilyLength == 0 || fontFamilyLength > 256 {
					v.add(rp+".fontFamily", "schema.maxLength", "must contain 1 to 256 characters")
				}
			}
		}
	}
}

func (v *nativeValidator) table(table NativeTable, transform NativeTransform, p string) {
	if len(table.ColumnWidths) == 0 {
		v.add(p+".columnWidths", "schema.minItems", "must contain at least one width")
	}
	if table.RowHeights == nil {
		v.add(p+".rowHeights", "schema.required", "must be an array")
	}
	if len(table.Rows) == 0 {
		v.add(p+".rows", "schema.minItems", "must contain at least one row")
	}
	if len(table.ColumnWidths) > nativeMaxTableColumns {
		v.add(p+".columnWidths", "schema.maxItems", fmt.Sprintf("must contain at most %d columns", nativeMaxTableColumns))
	}
	if len(table.RowHeights) > nativeMaxTableRows {
		v.add(p+".rowHeights", "schema.maxItems", fmt.Sprintf("must contain at most %d heights", nativeMaxTableRows))
	}
	if len(table.Rows) > nativeMaxTableRows {
		v.add(p+".rows", "schema.maxItems", fmt.Sprintf("must contain at most %d rows", nativeMaxTableRows))
	}
	if len(table.RowHeights) != 0 && len(table.RowHeights) != len(table.Rows) {
		v.add(p+".rowHeights", "native.tableDimensions", "must be empty or contain one height per row")
	}
	for index, width := range table.ColumnWidths[:boundedLength(len(table.ColumnWidths), nativeMaxTableColumns)] {
		v.positive(width, fmt.Sprintf("%s.columnWidths[%d]", p, index))
	}
	for index, height := range table.RowHeights[:boundedLength(len(table.RowHeights), nativeMaxTableRows)] {
		v.positive(height, fmt.Sprintf("%s.rowHeights[%d]", p, index))
	}
	authoritativeCells := 0
	legacyCells := 0
	for rowIndex, row := range table.Rows[:boundedLength(len(table.Rows), nativeMaxTableRows)] {
		rp := fmt.Sprintf("%s.rows[%d]", p, rowIndex)
		if len(row) != len(table.ColumnWidths) {
			v.add(rp, "native.tableDimensions", "must contain one cell per column")
		}
		if len(row) > nativeMaxTableColumns {
			v.add(rp, "schema.maxItems", fmt.Sprintf("must contain at most %d cells", nativeMaxTableColumns))
		}
		v.tableCells += boundedLength(len(row), nativeMaxTableColumns)
		for cellIndex, cell := range row[:boundedLength(len(row), nativeMaxTableColumns)] {
			cp := fmt.Sprintf("%s[%d]", rp, cellIndex)
			if cell.Text == nil {
				v.add(cp+".text", "schema.required", "is required")
			} else {
				length := utf16CodeUnitLengthBounded(*cell.Text, nativeMaxTextCodeUnits+1)
				if length > nativeMaxTextCodeUnits {
					v.add(cp+".text", "schema.maxLength", fmt.Sprintf("must contain at most %d code units", nativeMaxTextCodeUnits))
				}
			}
			hasParagraphs := cell.Paragraphs != nil
			hasTextBody := cell.TextBody != nil
			if hasParagraphs != hasTextBody {
				v.add(cp, "native.tableTextAuthority", "paragraphs and textBody must be supplied together")
			}
			if hasParagraphs && hasTextBody {
				if cell.TextBody.AutoFit != "none" {
					v.add(cp+".textBody.autoFit", "native.autofitApproximation", "table cell autofit preview is not supported")
				}
				if cell.TextBody.WritingMode != nil {
					v.add(cp+".textBody.writingMode", "native.verticalPreview", "vertical table cells are not supported")
				}
				authoritativeCells++
				v.paragraphs(*cell.Paragraphs, cp+".paragraphs")
				if cell.Align != nil {
					v.add(cp+".align", "native.tableTextAuthority", "legacy align is not allowed with authoritative cell paragraphs")
				}
				if cell.Text != nil && !nativeTableCellTextMatches(*cell.Text, *cell.Paragraphs) {
					v.add(cp+".text", "native.tableTextAuthority", "must equal the newline-joined authoritative paragraph text")
				}
				if rowIndex < len(table.RowHeights) && cellIndex < len(table.ColumnWidths) {
					width, height := table.ColumnWidths[cellIndex], table.RowHeights[rowIndex]
					v.textBody(*cell.TextBody, NativeTransform{X: int64Pointer(0), Y: int64Pointer(0), Cx: &width, Cy: &height}, cp+".textBody")
				}
			} else {
				legacyCells++
				if cell.Text != nil {
					v.addTextCodeUnits(*cell.Text)
				}
			}
			if cell.Fill != nil {
				v.color(*cell.Fill, cp+".fill")
			}
			if cell.Border != nil {
				v.color(cell.Border.Color, cp+".border.color")
				v.requiredBoundedNonnegative(cell.Border.WidthEMU, nativeMaxLineWidthEmu, cp+".border.widthEmu")
			}
			if cell.Align != nil && *cell.Align != NativeTextAlignLeft && *cell.Align != NativeTextAlignCenter && *cell.Align != NativeTextAlignRight {
				v.add(cp+".align", "schema.enum", "has an unsupported text alignment")
			}
		}
	}
	if authoritativeCells != 0 && legacyCells != 0 {
		v.add(p+".rows", "native.tableTextAuthority", "authoritative and legacy table cells cannot be mixed")
	}
	if authoritativeCells != 0 {
		if len(table.RowHeights) != len(table.Rows) {
			v.add(p+".rowHeights", "native.tableGeometry", "authoritative tables require one exact height per row")
		}
		columnTotal, columnOK := nativeExactTableTrackTotal(table.ColumnWidths)
		rowTotal, rowOK := nativeExactTableTrackTotal(table.RowHeights)
		if transform.Cx == nil || !columnOK || columnTotal != *transform.Cx {
			v.add(p+".columnWidths", "native.tableGeometry", "authoritative column tracks must sum exactly to the table frame width")
		}
		if transform.Cy == nil || !rowOK || rowTotal != *transform.Cy {
			v.add(p+".rowHeights", "native.tableGeometry", "authoritative row tracks must sum exactly to the table frame height")
		}
	}
}

func nativeExactTableTrackTotal(tracks []int64) (int64, bool) {
	var total int64
	for _, track := range tracks {
		if track <= 0 || track > nativeMaxSafeInteger || total > nativeMaxSafeInteger-track {
			return 0, false
		}
		total += track
	}
	return total, true
}

func nativeTableCellTextMatches(text string, paragraphs []NativeParagraph) bool {
	cursor := 0
	for paragraphIndex, paragraph := range paragraphs {
		if paragraphIndex != 0 {
			if cursor >= len(text) || text[cursor] != '\n' {
				return false
			}
			cursor++
		}
		for _, run := range paragraph.Runs {
			if run.Text == nil || len(text)-cursor < len(*run.Text) || text[cursor:cursor+len(*run.Text)] != *run.Text {
				return false
			}
			cursor += len(*run.Text)
		}
	}
	return cursor == len(text)
}

func (v *nativeValidator) chart(chart NativeOpaqueChart, p string) {
	v.part(chart.ChartPart, p+".chartPart")
	if !nativeIDPattern.MatchString(chart.RelationshipID) {
		v.add(p+".relationshipId", "schema.pattern", "has an invalid relationship id")
	}
	v.passthrough(chart.OpaqueRef, p+".opaqueRef")
	if chart.OpaqueRef.OwnerPart != chart.ChartPart {
		v.add(p+".opaqueRef.ownerPart", "native.chartReference", "must equal chartPart")
	}
	if chart.PreviewAssetID != nil {
		if asset, ok := v.assets[*chart.PreviewAssetID]; !ok {
			v.add(p+".previewAssetId", "native.assetReference", "references an unknown asset id")
		} else if !strings.HasPrefix(asset.ContentType, "image/") {
			v.add(p+".previewAssetId", "native.assetType", "chart previews must reference an image asset")
		}
	}
}

func (v *nativeValidator) animation(animation NativeAnimation, p string) {
	if animation.Effect != NativeAnimationEffectFade && animation.Effect != NativeAnimationEffectFlyIn {
		v.add(p+".effect", "schema.enum", "has an unsupported animation effect")
	}
	if animation.Direction != nil && !validNativeDirection(*animation.Direction) {
		v.add(p+".direction", "schema.enum", "has an unsupported direction")
	}
	if animation.DelayMS != nil && (*animation.DelayMS < 0 || *animation.DelayMS > 86400000) {
		v.add(p+".delayMs", "schema.range", "must be between 0 and 86400000")
	}
	if animation.DurationMS != nil && (*animation.DurationMS < 1 || *animation.DurationMS > 86400000) {
		v.add(p+".durationMs", "schema.range", "must be between 1 and 86400000")
	}
	if animation.DistancePPM != nil && (*animation.DistancePPM < 0 || *animation.DistancePPM > 1000000) {
		v.add(p+".distancePpm", "schema.range", "must be between 0 and 1000000")
	}
	if animation.Effect == NativeAnimationEffectFade && (animation.Direction != nil || animation.DistancePPM != nil) {
		v.add(p, "native.animation", "fade cannot declare direction or distancePpm")
	}
}

func (v *nativeValidator) transition(transition NativeTransition, p string) {
	if transition.Type != NativeTransitionTypeFade && transition.Type != NativeTransitionTypePush && transition.Type != NativeTransitionTypeWipe {
		v.add(p+".type", "schema.enum", "has an unsupported transition type")
	}
	if transition.Direction != nil && !validNativeDirection(*transition.Direction) {
		v.add(p+".direction", "schema.enum", "has an unsupported direction")
	}
	if transition.Type == NativeTransitionTypeFade && transition.Direction != nil {
		v.add(p, "native.transition", "fade cannot declare a direction")
	}
}

func (v *nativeValidator) stroke(stroke NativeStroke, p string) {
	v.color(stroke.Color, p+".color")
	v.requiredBoundedNonnegative(stroke.WidthEMU, nativeMaxLineWidthEmu, p+".widthEmu")
	present := 0
	if stroke.Cap != nil {
		present++
	}
	if stroke.Join != nil {
		present++
	}
	if stroke.Dash != nil {
		present++
	}
	if present != 0 && present != 3 {
		v.add(p, "native.strokeMetadata", "cap, join, and dash must be supplied together")
	}
	if stroke.Cap != nil && *stroke.Cap != NativeStrokeCapFlat && *stroke.Cap != NativeStrokeCapRound && *stroke.Cap != NativeStrokeCapSquare {
		v.add(p+".cap", "schema.enum", "has an unsupported line cap")
	}
	if stroke.Join != nil && *stroke.Join != NativeStrokeJoinRound && *stroke.Join != NativeStrokeJoinBevel && *stroke.Join != NativeStrokeJoinMiter {
		v.add(p+".join", "schema.enum", "has an unsupported line join")
	}
	if stroke.Dash != nil && *stroke.Dash != NativeStrokeDashSolid {
		v.add(p+".dash", "schema.enum", "has an unsupported line dash")
	}
	if stroke.Join != nil && *stroke.Join == NativeStrokeJoinMiter {
		v.requiredBoundedNonnegative(stroke.MiterLimit, nativeMaxDrawingPercentage, p+".miterLimit")
	} else if stroke.MiterLimit != nil {
		v.add(p+".miterLimit", "native.stroke", "is allowed only for a miter join")
	}
}
func (v *nativeValidator) transform(transform NativeTransform, p string) {
	v.requiredInteger(transform.X, p+".x")
	v.requiredInteger(transform.Y, p+".y")
	v.requiredPositive(transform.Cx, p+".cx")
	v.requiredPositive(transform.Cy, p+".cy")
	if transform.QuarterTurns != nil {
		q := *transform.QuarterTurns
		if q < 1 || q > 3 {
			v.add(p+".quarterTurns", "native.rotation", "must be 1, 2, or 3")
		}
		if q%2 != 0 && transform.Cx != nil && transform.Cy != nil && (*transform.Cx%2 != *transform.Cy%2) {
			v.add(p+".quarterTurns", "native.rotation", "quarter-turn center must remain exact integer EMU")
		}
	}
}
func (v *nativeValidator) placeholder(value *NativePlaceholderType, p string) {
	if value != nil && *value != NativePlaceholderTypeTitle && *value != NativePlaceholderTypeCtrTitle && *value != NativePlaceholderTypeSubTitle && *value != NativePlaceholderTypeBody {
		v.add(p, "schema.enum", "has an unsupported placeholder type")
	}
}
func (v *nativeValidator) shapePreset(value *NativeShapePreset, p string) {
	if value == nil {
		return
	}
	switch *value {
	case NativeShapePresetRect, NativeShapePresetRoundRect, NativeShapePresetEllipse, NativeShapePresetTriangle, NativeShapePresetDiamond, NativeShapePresetRightArrow, NativeShapePresetPentagon, NativeShapePresetHexagon, NativeShapePresetStar5:
	default:
		v.add(p, "schema.enum", "has an unsupported shape preset")
	}
}

func (v *nativeValidator) validateDiagnosticReferences(deck NativePPTXDeck) {
	check := func(diagnostics []NativeDiagnostic, p string) {
		for index, diagnostic := range diagnostics[:boundedLength(len(diagnostics), nativeMaxDiagnosticsPerScope)] {
			if diagnostic.Scope == nil {
				continue
			}
			dp := fmt.Sprintf("%s[%d].scope", p, index)
			if diagnostic.Scope.SlideID != nil && !v.slides[*diagnostic.Scope.SlideID] {
				v.add(dp+".slideId", "native.scopeReference", "references an unknown slide id")
			}
			if diagnostic.Scope.ElementID != nil {
				if _, ok := v.elements[*diagnostic.Scope.ElementID]; !ok {
					v.add(dp+".elementId", "native.scopeReference", "references an unknown element id")
				}
			}
			if diagnostic.Scope.SlideID != nil && diagnostic.Scope.ElementID != nil {
				if ownerSlide, ok := v.elements[*diagnostic.Scope.ElementID]; ok && ownerSlide != *diagnostic.Scope.SlideID {
					v.add(dp+".elementId", "native.scopeOwnership", "element does not belong to the scoped slide")
				}
			}
		}
	}
	check(deck.Compatibility.Diagnostics, "$.compatibility.diagnostics")
	for slideIndex, slide := range deck.Slides[:boundedLength(len(deck.Slides), nativeMaxSlides)] {
		check(slide.Compatibility.Diagnostics, fmt.Sprintf("$.slides[%d].compatibility.diagnostics", slideIndex))
		var walk func([]NativeElement, string, int)
		walk = func(elements []NativeElement, p string, depth int) {
			if depth > nativeMaxDepth {
				return
			}
			for index, element := range elements[:boundedLength(len(elements), nativeMaxElementsPerContainer)] {
				ep := fmt.Sprintf("%s[%d]", p, index)
				check(element.Compatibility.Diagnostics, ep+".compatibility.diagnostics")
				walk(element.Children, ep+".children", depth+1)
			}
		}
		walk(slide.Elements, fmt.Sprintf("$.slides[%d].elements", slideIndex), 1)
	}
}

func (v *nativeValidator) id(value, p string) {
	if !nativeIDPattern.MatchString(value) {
		v.add(p, "schema.pattern", "has an invalid durable-id format")
	}
	if v.ids[value] {
		v.add(p, "native.duplicateId", fmt.Sprintf("duplicate durable id %s", value))
	}
	v.ids[value] = true
}
func (v *nativeValidator) hash(value, p string) {
	if !sha256Pattern.MatchString(value) {
		v.add(p, "schema.pattern", "must be a lowercase SHA-256 digest")
	}
}
func (v *nativeValidator) color(value, p string) {
	if !colorPattern.MatchString(value) {
		v.add(p, "schema.pattern", "must be six uppercase hexadecimal digits")
	}
}
func (v *nativeValidator) part(value, p string) {
	if len(value) == 0 || len(value) > 1024 || strings.HasPrefix(value, "/") || !partCharacterPattern.MatchString(value) {
		v.add(p, "schema.pattern", "must be a safe canonical package-relative part name")
		return
	}
	components := strings.Split(value, "/")
	for _, component := range components {
		decoded, err := url.PathUnescape(component)
		if err != nil || !utf8.ValidString(decoded) || decoded == "" || decoded == "." || decoded == ".." || strings.HasSuffix(decoded, ".") || strings.ContainsAny(decoded, "/\\") {
			v.add(p, "schema.pattern", "must be a safe canonical package-relative part name")
			return
		}
		for _, codePoint := range decoded {
			if codePoint < 0x20 || codePoint == 0x7f {
				v.add(p, "schema.pattern", "must be a safe canonical package-relative part name")
				return
			}
		}
		for index := 0; index < len(component); index++ {
			if component[index] != '%' {
				continue
			}
			if index+2 >= len(component) || !isUpperHex(component[index+1]) || !isUpperHex(component[index+2]) {
				v.add(p, "schema.pattern", "must be a safe canonical package-relative part name")
				return
			}
			index += 2
		}
	}
}
func isUpperHex(value byte) bool {
	return value >= '0' && value <= '9' || value >= 'A' && value <= 'F'
}
func (v *nativeValidator) positive(value int64, p string) {
	if value <= 0 || value > nativeMaxSafeInteger {
		v.add(p, "schema.range", "must be a positive safe integer")
	}
}
func (v *nativeValidator) nonnegative(value int64, p string) {
	if value < 0 || value > nativeMaxSafeInteger {
		v.add(p, "schema.range", "must be a non-negative safe integer")
	}
}
func (v *nativeValidator) integer(value int64, p string) {
	if value < -nativeMaxSafeInteger || value > nativeMaxSafeInteger {
		v.add(p, "schema.range", "must be a safe integer")
	}
}
func (v *nativeValidator) requiredInteger(value *int64, p string) {
	if value == nil {
		v.add(p, "schema.required", "is required")
		return
	}
	v.integer(*value, p)
}
func (v *nativeValidator) requiredPositive(value *int64, p string) {
	if value == nil {
		v.add(p, "schema.required", "is required")
		return
	}
	v.positive(*value, p)
}
func (v *nativeValidator) requiredNonnegative(value *int64, p string) {
	if value == nil {
		v.add(p, "schema.required", "is required")
		return
	}
	v.nonnegative(*value, p)
}
func (v *nativeValidator) requiredBoundedNonnegative(value *int64, maximum int64, p string) {
	if value == nil {
		v.add(p, "schema.required", "is required")
		return
	}
	if *value < 0 || *value > maximum {
		v.add(p, "schema.range", fmt.Sprintf("must be between 0 and %d", maximum))
	}
}
func boundedLength(length int, maximum int) int {
	if length > maximum {
		return maximum
	}
	return length
}
func (v *nativeValidator) addTextCodeUnits(value string) int {
	if v.textCodeUnits > nativeMaxTotalTextCodeUnits {
		return 0
	}
	length := utf16CodeUnitLengthBounded(value, nativeMaxTextCodeUnits)
	v.textCodeUnits += int64(length)
	return length
}
func utf16CodeUnitLengthBounded(value string, maximum int) int {
	length := 0
	for _, codePoint := range value {
		if codePoint > 0xffff {
			length += 2
		} else {
			length++
		}
		if length > maximum {
			break
		}
	}
	return length
}
func (v *nativeValidator) add(p, code, message string) {
	if len(v.issues) < 100 {
		v.issues = append(v.issues, NativeContractIssue{Path: p, Code: code, Message: message})
	}
}

func validNativeDirection(value NativeDirection) bool {
	return value == NativeDirectionLeft || value == NativeDirectionRight || value == NativeDirectionUp || value == NativeDirectionDown
}
func worseNativeStatus(left, right NativeCompatibilityStatus) NativeCompatibilityStatus {
	if nativeStatusRank(left) >= nativeStatusRank(right) {
		return left
	}
	return right
}
func nativeStatusRank(value NativeCompatibilityStatus) int {
	switch value {
	case NativeCompatibilityStatusEditable:
		return 0
	case NativeCompatibilityStatusPreserveOnly:
		return 1
	case NativeCompatibilityStatusRefused:
		return 2
	default:
		return 3
	}
}

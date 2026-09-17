package docxpatch

import (
	"bytes"
	"encoding/json"
	"fmt"
	"sort"
	"strconv"
	"strings"
	"unicode/utf8"
)

// This file owns the bounded, source-order WordprocessingML list state machine.
// OPC discovery and style application remain registered in native_style_resolve.go;
// marker counters and text are resolved here before any renderer sees them.

type nativeNumberingState struct {
	values          map[string]map[int]int
	instanceStarted map[string]map[int]bool
}

func newNativeNumberingState() *nativeNumberingState {
	return &nativeNumberingState{values: map[string]map[int]int{}, instanceStarted: map[string]map[int]bool{}}
}

func nativeEffectiveNumberingLevel(abstract *nativeNumberingLevel, override *nativeNumberingOverride) *nativeNumberingLevel {
	if override == nil || override.level == nil {
		return abstract
	}
	copyLevel := *override.level
	// Word applies only the concrete lvlOverride/startOverride to counter starts,
	// and ignores replacement-level start and lvlRestart. Preserve those abstract
	// policies when present; an override-only level uses the ordinary defaults.
	if abstract != nil {
		copyLevel.start = abstract.start
		copyLevel.restart = abstract.restart
	} else {
		copyLevel.start = nil
		copyLevel.restart = nil
	}
	return &copyLevel
}

func (resolver *nativeLayoutResolver) effectiveNumberingLevel(instance *nativeNumberingInstance, abstract *nativeAbstractNumbering, level int) *nativeNumberingLevel {
	base := abstract.levels[level]
	return nativeEffectiveNumberingLevel(base, instance.overrides[level])
}

func nativeNumberingDefaults(level *nativeNumberingLevel) (start int, format, text, suffix, alignment string) {
	start = 0
	format = "decimal"
	suffix = "tab"
	alignment = "start"
	if level.start != nil {
		start = *level.start
	}
	if level.format != nil {
		format = *level.format
	}
	if level.text != nil {
		text = *level.text
	}
	if level.suffix != nil {
		suffix = *level.suffix
	}
	if level.alignment != nil {
		alignment = *level.alignment
	}
	return
}

func nativeRestartPolicy(level int, definition *nativeNumberingLevel) (after *int, never bool, valid bool) {
	if level == 0 {
		if definition.restart != nil && *definition.restart != 0 {
			return nil, false, false
		}
		return nil, true, true
	}
	if definition.restart == nil {
		value := level - 1
		return &value, false, true
	}
	if *definition.restart == 0 {
		return nil, true, true
	}
	value := *definition.restart - 1
	if value < 0 || value >= level {
		return nil, false, false
	}
	return &value, false, true
}

func (resolver *nativeLayoutResolver) advanceNumberingState(state *nativeNumberingState, instance *nativeNumberingInstance, abstract *nativeAbstractNumbering, level int, start int, hasStartOverride bool) (int, bool) {
	values := state.values[instance.id]
	if values == nil {
		values = map[int]int{}
		state.values[instance.id] = values
	}
	started := state.instanceStarted[instance.id]
	if started == nil {
		started = map[int]bool{}
		state.instanceStarted[instance.id] = started
	}
	for deeper := range values {
		if deeper <= level {
			continue
		}
		definition := resolver.effectiveNumberingLevel(instance, abstract, deeper)
		if definition == nil {
			return 0, false
		}
		restartAfter, never, valid := nativeRestartPolicy(deeper, definition)
		if !valid {
			return 0, false
		}
		if !never && restartAfter != nil && level <= *restartAfter {
			delete(values, deeper)
		}
	}
	value, present := values[level]
	if !present || (hasStartOverride && !started[level]) {
		value = start
	} else {
		value++
	}
	if value < 0 || value > 2_147_483_647 {
		return 0, false
	}
	values[level] = value
	started[level] = true
	return value, true
}

func nativeFormatAlphabetic(value int, upper bool) (string, bool) {
	if value <= 0 || value > 806 {
		return "", false
	}
	base := byte('a')
	if upper {
		base = 'A'
	}
	letter := base + byte((value-1)%26)
	repeat := ((value - 1) / 26) + 1
	return strings.Repeat(string([]byte{letter}), repeat), true
}

func nativeFormatRoman(value int, upper bool) (string, bool) {
	if value <= 0 || value > 31_000 {
		return "", false
	}
	entries := []struct {
		value int
		text  string
	}{{1000, "M"}, {900, "CM"}, {500, "D"}, {400, "CD"}, {100, "C"}, {90, "XC"}, {50, "L"}, {40, "XL"}, {10, "X"}, {9, "IX"}, {5, "V"}, {4, "IV"}, {1, "I"}}
	var output strings.Builder
	for _, entry := range entries {
		for value >= entry.value {
			output.WriteString(entry.text)
			value -= entry.value
		}
	}
	result := output.String()
	if !upper {
		result = strings.ToLower(result)
	}
	return result, true
}

// The ideographic numbering systems this tier models. Each row was read off
// Word's own rendering of the benchmark documents rather than inferred from
// the format name.
//
//   - cycle is a fixed sequence: the Heavenly Stems and the Earthly Branches.
//     Word exhausts the sequence once and then falls back to decimal instead
//     of restarting it, so its length is a real boundary, not a modulus.
//   - digits with a unit is a counting system: a tens unit sits between two
//     digits. leadingUnit keeps the explicit "one" before that unit, as the
//     legal/financial numerals require (ten is one-ten) and the plain counting
//     numerals forbid (ten is ten). Hundreds and above need the
//     zero-insertion rules, which these documents do not attest, so the
//     modelled range stops at ninety-nine.
//   - digits without a unit is a positional system: every decimal digit maps
//     to one ideograph and no unit appears at all (ten is one-zero).
//
// Every ideograph below is in U+0800..U+FFFF, so it occupies exactly three
// UTF-8 bytes and the tables are indexed by byte triple rather than decoded.
var nativeIdeographicNumberSystems = [...]struct {
	format, cycle, digits, unit string
	leadingUnit                 bool
}{
	{format: "ideographTraditional", cycle: "\u7532\u4e59\u4e19\u4e01\u620a\u5df1\u5e9a\u8f9b\u58ec\u7678"},
	{format: "ideographZodiac", cycle: "\u5b50\u4e11\u5bc5\u536f\u8fb0\u5df3\u5348\u672a\u7533\u9149\u620c\u4ea5"},
	{format: "taiwaneseCountingThousand", digits: "\u3007\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d", unit: "\u5341"},
	{format: "ideographLegalTraditional", digits: "\u96f6\u58f9\u8cb3\u53c3\u8086\u4f0d\u9678\u67d2\u634c\u7396", unit: "\u62fe", leadingUnit: true},
	{format: "koreanDigital2", digits: "\u96f6\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d"},
}

func nativeFormatIdeographicCounter(value int, format string) (string, bool) {
	for _, system := range nativeIdeographicNumberSystems {
		if system.format != format {
			continue
		}
		if value < 1 {
			return "", false
		}
		if system.cycle != "" {
			if value*3 > len(system.cycle) {
				return strconv.Itoa(value), true
			}
			return system.cycle[(value-1)*3 : value*3], true
		}
		if system.unit == "" {
			output := ""
			for _, digit := range strconv.Itoa(value) {
				output += system.digits[(digit-'0')*3 : (digit-'0')*3+3]
			}
			return output, true
		}
		if value > 99 {
			return "", false
		}
		tens, units, output := value/10, value%10, ""
		if tens > 0 {
			if tens > 1 || system.leadingUnit {
				output = system.digits[tens*3 : tens*3+3]
			}
			output += system.unit
		}
		if units > 0 || tens == 0 {
			output += system.digits[units*3 : units*3+3]
		}
		return output, true
	}
	return "", false
}

// nativeFormatEnclosedCircleCounter renders ECMA-376 §17.18.59
// `decimalEnclosedCircle`: the counter as a decimal digit drawn inside a
// circle. Unicode encodes that series as one precomposed character per value,
// CIRCLED DIGIT ONE (U+2460) through CIRCLED NUMBER TWENTY (U+2473), and stops
// there: twenty-one has no enclosed form in that block. The modelled range is
// therefore one through twenty, the exact set Unicode supplies. A larger
// counter is left unformatted rather than approximated with a composed circle
// or a bare digit, because neither is the character the format names.
func nativeFormatEnclosedCircleCounter(value int) (string, bool) {
	if value < 1 || value > 20 {
		return "", false
	}
	return string(rune(0x2460 + value - 1)), true
}

func nativeFormatNumberingCounter(value int, format string) (string, bool) {
	switch format {
	case "decimal":
		return strconv.Itoa(value), true
	case "decimalEnclosedCircle":
		return nativeFormatEnclosedCircleCounter(value)
	case "lowerLetter":
		return nativeFormatAlphabetic(value, false)
	case "upperLetter":
		return nativeFormatAlphabetic(value, true)
	case "lowerRoman":
		return nativeFormatRoman(value, false)
	case "upperRoman":
		return nativeFormatRoman(value, true)
	default:
		return nativeFormatIdeographicCounter(value, format)
	}
}

func (resolver *nativeLayoutResolver) resolveLevelText(instance *nativeNumberingInstance, abstract *nativeAbstractNumbering, currentLevel int, format, template string, values map[int]int) (string, []NativeResolvedCounterValueV1, error) {
	if template == "" || !utf8.ValidString(template) || strings.ContainsAny(template, "\x00\r\n") {
		return "", nil, fmt.Errorf("missing or invalid lvlText")
	}
	if format == "bullet" {
		if utf8.RuneCountInString(template) > 31 {
			return "", nil, fmt.Errorf("expanded bullet lvlText exceeds 31 characters")
		}
		return template, []NativeResolvedCounterValueV1{}, nil
	}
	var output strings.Builder
	countersByLevel := map[int]NativeResolvedCounterValueV1{}
	tokens := 0
	for index := 0; index < len(template); {
		if template[index] != '%' {
			r, size := utf8.DecodeRuneInString(template[index:])
			if r == utf8.RuneError && size == 1 {
				return "", nil, fmt.Errorf("invalid UTF-8 in lvlText")
			}
			output.WriteRune(r)
			index += size
			continue
		}
		if index+1 >= len(template) || template[index+1] < '1' || template[index+1] > '9' {
			output.WriteByte('%')
			index++
			continue
		}
		if index+2 < len(template) && template[index+2] >= '0' && template[index+2] <= '9' {
			return "", nil, fmt.Errorf("multi-digit lvlText placeholder is ambiguous")
		}
		referenced := int(template[index+1] - '1')
		if referenced > currentLevel {
			return "", nil, fmt.Errorf("lvlText references a deeper numbering level")
		}
		value, present := values[referenced]
		if !present {
			return "", nil, fmt.Errorf("lvlText references an uninitialized parent counter")
		}
		definition := resolver.effectiveNumberingLevel(instance, abstract, referenced)
		if definition == nil {
			return "", nil, fmt.Errorf("lvlText references a missing numbering level")
		}
		_, referencedFormat, _, _, _ := nativeNumberingDefaults(definition)
		formatted, ok := nativeFormatNumberingCounter(value, referencedFormat)
		if !ok {
			return "", nil, fmt.Errorf("lvlText references an unsupported or out-of-range numbering format")
		}
		output.WriteString(formatted)
		countersByLevel[referenced] = NativeResolvedCounterValueV1{Level: referenced, Value: value, Format: referencedFormat}
		tokens++
		if tokens > 9 {
			return "", nil, fmt.Errorf("lvlText exceeds nine placeholders")
		}
		index += 2
	}
	resolved := output.String()
	if utf8.RuneCountInString(resolved) > 31 {
		return "", nil, fmt.Errorf("expanded lvlText exceeds 31 characters")
	}
	levels := make([]int, 0, len(countersByLevel))
	for level := range countersByLevel {
		levels = append(levels, level)
	}
	sort.Ints(levels)
	counters := make([]NativeResolvedCounterValueV1, 0, len(levels))
	for _, level := range levels {
		counters = append(counters, countersByLevel[level])
	}
	return resolved, counters, nil
}

func (resolver *nativeLayoutResolver) materializeNativeNumbering(instance *nativeNumberingInstance, abstract *nativeAbstractNumbering, level *nativeNumberingLevel, override *nativeNumberingOverride, levelIndex int, scopeID string, state *nativeNumberingState) *NativeResolvedNumberingV1 {
	start, format, template, suffix, alignment := nativeNumberingDefaults(level)
	if override != nil && override.start != nil {
		start = *override.start
	}
	if !nativeOrdinaryNumberFormat(format) {
		resolver.addDiagnostic("UNSUPPORTED_NUMBER_FORMAT", scopeID, level.partName, level.node, "This numbering format is outside the bounded decimal/letter/Roman/bullet/ideographic set")
		return nil
	}
	if level.picture {
		resolver.addDiagnostic("PICTURE_BULLET_PRESERVED", scopeID, level.partName, level.node, "Picture-bullet marker semantics are preserved and not guessed")
		return nil
	}
	if alignment != "left" && alignment != "right" && alignment != "center" && alignment != "start" && alignment != "end" {
		resolver.addDiagnostic("UNSUPPORTED_NUMBER_ALIGNMENT", scopeID, level.partName, level.node, "Numbering label alignment must be left, right, center, start, or end")
		return nil
	}
	restartAfter, never, validRestart := nativeRestartPolicy(levelIndex, level)
	if !validRestart {
		resolver.addDiagnostic("UNSUPPORTED_NUMBER_RESTART", scopeID, level.partName, level.node, "lvlRestart is outside the bounded one-based ancestor range")
		return nil
	}
	previousValues := map[int]int{}
	for key, value := range state.values[instance.id] {
		previousValues[key] = value
	}
	previousStarted := map[int]bool{}
	for key, value := range state.instanceStarted[instance.id] {
		previousStarted[key] = value
	}
	hasInstanceStart := override != nil && override.start != nil
	counterValue, advanced := resolver.advanceNumberingState(state, instance, abstract, levelIndex, start, hasInstanceStart)
	if !advanced {
		state.values[instance.id] = previousValues
		state.instanceStarted[instance.id] = previousStarted
		resolver.addDiagnostic("NUMBERING_COUNTER_OVERFLOW", scopeID, level.partName, level.node, "Numbering counter state is missing, invalid, or outside the bounded 31-bit range")
		return nil
	}
	values := state.values[instance.id]
	// A w:lvlText that is present and empty is a label with no text, not a
	// missing one. Word advances the counter and draws nothing. The counter
	// vector is empty for the same reason: no placeholder was expanded.
	if level.text != nil && template == "" {
		return &NativeResolvedNumberingV1{
			MarkerID: nativeStableID("marker", level.partName, scopeID, instance.id+":"+strconv.Itoa(levelIndex)),
			NumID:    instance.id, AbstractNumID: instance.abstractID, Level: levelIndex,
			Start: start, Format: format, Text: template, Suffix: suffix, Alignment: alignment,
			RestartAfterLevel: restartAfter, NeverRestart: never, CounterValue: counterValue,
			CounterValues: []NativeResolvedCounterValueV1{}, ResolvedText: "",
			Marker: NativeResolvedRunPropertiesV1{}, NumberingTabTwips: level.numTab,
			alignmentDefaulted: level.alignment == nil,
		}
	}
	resolvedText, counterValues, err := resolver.resolveLevelText(instance, abstract, levelIndex, format, template, values)
	if err != nil {
		state.values[instance.id] = previousValues
		state.instanceStarted[instance.id] = previousStarted
		resolver.addDiagnostic("MALFORMED_NUMBERING_TEXT", scopeID, level.partName, level.node, err.Error())
		return nil
	}
	result := &NativeResolvedNumberingV1{
		MarkerID: nativeStableID("marker", level.partName, scopeID, instance.id+":"+strconv.Itoa(levelIndex)),
		NumID:    instance.id, AbstractNumID: instance.abstractID, Level: levelIndex,
		Start: start, Format: format, Text: template, Suffix: suffix, Alignment: alignment,
		RestartAfterLevel: restartAfter, NeverRestart: never, CounterValue: counterValue, CounterValues: counterValues,
		ResolvedText: resolvedText, Marker: NativeResolvedRunPropertiesV1{},
		NumberingTabTwips:  level.numTab,
		alignmentDefaulted: level.alignment == nil,
	}
	return result
}

func nativeResolvedNumberingDefinitionSHA256(numbering *NativeResolvedNumberingV1, partSHA256 string) string {
	payload := struct {
		Protocol          string  `json:"protocol"`
		PartSHA256        string  `json:"part_sha256"`
		NumID             string  `json:"num_id"`
		AbstractNumID     string  `json:"abstract_num_id"`
		Level             int     `json:"level"`
		LevelStyleID      *string `json:"level_style_id,omitempty"`
		Start             int     `json:"start"`
		Format            string  `json:"format"`
		Text              string  `json:"text"`
		Suffix            string  `json:"suffix"`
		Alignment         string  `json:"alignment"`
		RestartAfterLevel *int    `json:"restart_after_level,omitempty"`
		NeverRestart      bool    `json:"never_restart"`
		NumberingTabTwips *int64  `json:"numbering_tab_twips,omitempty"`
		LabelStartTwips   int64   `json:"label_start_twips"`
		LabelEndTwips     int64   `json:"label_end_twips"`
		TextStartTwips    int64   `json:"text_start_twips"`
	}{
		"injoffice.docx.numbering-definition/v1", partSHA256, numbering.NumID, numbering.AbstractNumID,
		numbering.Level, numbering.LevelStyleID, numbering.Start, numbering.Format, numbering.Text, numbering.Suffix, numbering.Alignment,
		numbering.RestartAfterLevel, numbering.NeverRestart, numbering.NumberingTabTwips,
		numbering.LabelStartTwips, numbering.LabelEndTwips, numbering.TextStartTwips,
	}
	return nativeNumberingCanonicalSHA256V1(payload)
}

const nativeNumberingCanonicalSerializerV1 = "injoffice.canonical-json/utf8-v1"

// nativeNumberingCanonicalSHA256V1 is deliberately shared with the TypeScript
// numbering contract. It recursively orders object keys, emits UTF-8 JSON with
// HTML escaping disabled, retains Go's explicit U+2028/U+2029 escapes, and
// domain-separates the bytes with a versioned serializer identifier.
func nativeNumberingCanonicalSHA256V1(value any) string {
	raw, err := json.Marshal(value)
	if err != nil {
		return ""
	}
	decoder := json.NewDecoder(bytes.NewReader(raw))
	decoder.UseNumber()
	var canonical any
	if err := decoder.Decode(&canonical); err != nil {
		return ""
	}
	var encoded bytes.Buffer
	encoded.WriteString(nativeNumberingCanonicalSerializerV1)
	encoded.WriteByte(0)
	encoder := json.NewEncoder(&encoded)
	encoder.SetEscapeHTML(false)
	if err := encoder.Encode(canonical); err != nil {
		return ""
	}
	return nativeSHA(bytes.TrimSuffix(encoded.Bytes(), []byte{'\n'}))
}

func (resolver *nativeLayoutResolver) resolveNumberingGeometry(numbering *NativeResolvedNumberingV1, properties nativeParagraphProperties, scopeID string) bool {
	if numbering.ResolvedText == "" {
		// An empty label paints nothing. Without a hanging indent it also
		// displaces nothing: the numbering suffix cannot advance past a
		// text margin the label already sits on, so the paragraph lays out at
		// exactly the indents it resolved, and carrying a marker that paints
		// no glyph would only add an empty label region to the wire. Under a
		// hanging indent the empty label still owns that region and the suffix
		// still moves the first line, which is not modelled here, so that
		// shape keeps refusing rather than guessing a first-line offset.
		if properties.hanging == nil || *properties.hanging == 0 {
			return false
		}
		resolver.addDiagnostic("MALFORMED_NUMBERING_TEXT", scopeID, resolver.partsValue(resolver.parts.NumberingPart), nil, "An empty lvlText under a hanging indent leaves the label region and suffix unmodeled")
		return false
	}
	if numbering.alignmentDefaulted {
		numbering.Alignment = "left"
		if properties.bidi.present && properties.bidi.value {
			numbering.Alignment = "right"
		}
	}
	start := properties.indentStart
	if start == nil {
		if properties.bidi.present && properties.bidi.value {
			start = properties.indentRight
		} else {
			start = properties.indentLeft
		}
	}
	if start == nil || properties.hanging == nil || *properties.hanging <= 0 || *start < *properties.hanging {
		resolver.addDiagnostic("UNSUPPORTED_NUMBERING_GEOMETRY", scopeID, resolver.partsValue(resolver.parts.NumberingPart), nil, "Exact list-marker layout requires a nonnegative start indent and positive hanging indent")
		return false
	}
	numbering.LabelStartTwips = *start - *properties.hanging
	numbering.LabelEndTwips = *start
	numbering.TextStartTwips = *start
	numbering.DefinitionSHA256 = nativeResolvedNumberingDefinitionSHA256(numbering, resolver.numberingSource.PartSHA256)
	if numbering.Format == "lowerLetter" || numbering.Format == "upperLetter" {
		language := "und"
		if numbering.Marker.Language != nil {
			language = strings.ToLower(*numbering.Marker.Language)
		}
		if language != "und" && language != "en" && !strings.HasPrefix(language, "en-") {
			resolver.addDiagnostic("UNSUPPORTED_NUMBERING_LANGUAGE", scopeID, resolver.partsValue(resolver.parts.NumberingPart), nil, "Letter numbering is exact only for the attested English/und ASCII alphabet")
		}
	}
	return true
}

func nativeResolvedNumberingModelSHA256(input *NativeResolvedLayoutInputV1) string {
	markers := make([]NativeResolvedNumberingV1, 0)
	for _, paragraph := range input.Paragraphs {
		if paragraph.Numbering != nil {
			markers = append(markers, *paragraph.Numbering)
		}
	}
	type digestSource struct {
		RelationshipsPart   string `json:"relationships_part"`
		RelationshipsSHA256 string `json:"relationships_sha256"`
		RelationshipID      string `json:"relationship_id"`
		RelationshipType    string `json:"relationship_type"`
		RelationshipTarget  string `json:"relationship_target"`
		PartName            string `json:"part_name"`
		ContentType         string `json:"content_type"`
		PartSHA256          string `json:"part_sha256"`
	}
	type digestPayload struct {
		Protocol string                      `json:"protocol"`
		Source   digestSource                `json:"source"`
		Markers  []NativeResolvedNumberingV1 `json:"markers"`
	}
	payload := digestPayload{Protocol: "injoffice.docx.numbering-model/v1", Markers: markers}
	if input.NumberingSource != nil {
		source := input.NumberingSource
		payload.Source = digestSource{source.RelationshipsPart, source.RelationshipsSHA256, source.RelationshipID, source.RelationshipType, source.RelationshipTarget, source.PartName, source.ContentType, source.PartSHA256}
	}
	return nativeNumberingCanonicalSHA256V1(payload)
}

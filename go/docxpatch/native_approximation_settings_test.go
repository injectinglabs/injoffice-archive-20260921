package docxpatch

import (
	"reflect"
	"strings"
	"testing"
)

func TestNativeApproximationKnownSettingsRetainStrictRefusal(t *testing.T) {
	math := `<m:mathPr xmlns:m="` + nativeMathNamespace + `"><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr>`
	shape := `<w:shapeDefaults xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml"><o:shapedefaults v:ext="edit" spidmax="1026"/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout></w:shapeDefaults>`
	for _, test := range []struct {
		name, markup string
		eligible     bool
	}{
		{"language", `<w:themeFontLang w:val="en-US" w:eastAsia="x-none"/>`, true},
		{"locale", `<w:decimalSymbol w:val=","/><w:listSeparator w:val=";"/>`, true},
		{"math", math, true}, {"shape", shape, true},
		{"combined", math + shape + `<w:themeFontLang w:val="fr-FR"/><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/>`, true},
		{"compatibility flags", `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat>`, true},
		{"malformed language", `<w:themeFontLang w:val="en_US"/>`, false},
		{"unknown language", `<w:themeFontLang w:val="ar-SA"/>`, false},
		{"duplicate locale", `<w:decimalSymbol w:val="."/><w:decimalSymbol w:val="."/>`, false},
		{"unknown decimal", `<w:decimalSymbol w:val="unknown"/>`, false},
		{"nested locale", `<w:decimalSymbol w:val="."><w:foo/></w:decimalSymbol>`, false},
		{"unknown math", strings.Replace(math, `m:val="1440"`, `m:val="999"`, 1), false},
		{"duplicate math", strings.Replace(math, `</m:mathPr>`, `<m:intLim m:val="subSup"/></m:mathPr>`, 1), false},
		{"unknown math attribute", strings.Replace(math, `<m:mathPr `, `<m:mathPr foo="1" `, 1), false},
		{"malformed shape id", strings.Replace(shape, `spidmax="1026"`, `spidmax="01026"`, 1), false},
		{"shape content", strings.Replace(shape, `data="1"/>`, `data="1"><o:unknown/></o:idmap>`, 1), false},
		{"unknown flag value", `<w:compat><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="oops"/></w:compat>`, false},
		{"flag first without mode", `<w:compat><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat>`, false},
		{"flag before mode", `<w:compat><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>`, false},
	} {
		t.Run(test.name, func(t *testing.T) {
			data := buildNativeDOCX(t, nativeEntries(nativePaginationSettingsParts(`<w:settings xmlns:w="`+wordMLTransitional+`">`+test.markup+`</w:settings>`)))
			strict, err := ExtractNativePaginationSettingsV1(data)
			if err != nil {
				t.Fatal(err)
			}
			approx, err := ExtractNativeDocxApproximationEligibilityV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if (approx.Status == "eligible") != test.eligible {
				t.Fatalf("wrong eligibility %#v", approx)
			}
			if test.eligible && (len(approx.ApproximatedSettings) == 0 || len(approx.Reasons) <= len(strict.Diagnostics)) {
				t.Fatalf("missing source facts/warnings %#v", approx)
			}
			after, err := ExtractNativePaginationSettingsV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if strict.Profile != "unsupported" || !reflect.DeepEqual(strict, after) {
				t.Fatal("strict settings changed")
			}
		})
	}
}

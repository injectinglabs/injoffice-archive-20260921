// D11 breadth: docx themes. Surfaces (Extract) and surgically patches
// (Apply) the DrawingML theme part (word/theme/theme1.xml) — the 12-color
// scheme plus major/minor Latin fonts every real Word document carries,
// driving default colors/fonts across the whole document via theme
// references (though body text in THIS patcher's scope, per docxpatch.go's
// own model, sets explicit run properties rather than theme references —
// changing the theme changes what a reader sees for anything that DOES
// reference it: headings, accent-colored elements, charts inserted by
// chartwrite.go, which all default to theme colors/fonts per their own
// generated XML).
//
// Same fail-closed, single-part-touched discipline as everything else in
// this package: ApplyTheme replaces ONLY the specific named color slots
// and/or font typefaces requested, as a targeted string splice within the
// EXISTING theme1.xml (never a full regenerate) — the large fmtScheme
// block (fill/line/effect styles) and everything else in the part passes
// through completely untouched, verified via the same ApplyPatch fail-
// closed machinery.
//
// Scope, stated plainly: a document with NO theme part at all (rare in
// practice — every real Word-authored .docx has one; only a hand-built or
// third-party-generated file might lack it) is out of scope for Apply —
// generating a complete DEFAULT theme from nothing is a different task
// (a real design choice: dozens of fmtScheme fill/line/effect values) from
// "change an existing theme's colors/fonts," which is the actual agent
// use case this serves. ExtractTheme and ApplyTheme both return a clear
// error rather than silently fabricating a theme.
package docxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"regexp"
	"strings"
)

// ThemeColorSlot names the 12 canonical theme color slots, in the order
// OOXML's a:clrScheme schema declares them.
type ThemeColorSlot string

const (
	ThemeDark1             ThemeColorSlot = "dk1"
	ThemeLight1            ThemeColorSlot = "lt1"
	ThemeDark2             ThemeColorSlot = "dk2"
	ThemeLight2            ThemeColorSlot = "lt2"
	ThemeAccent1           ThemeColorSlot = "accent1"
	ThemeAccent2           ThemeColorSlot = "accent2"
	ThemeAccent3           ThemeColorSlot = "accent3"
	ThemeAccent4           ThemeColorSlot = "accent4"
	ThemeAccent5           ThemeColorSlot = "accent5"
	ThemeAccent6           ThemeColorSlot = "accent6"
	ThemeHyperlink         ThemeColorSlot = "hlink"
	ThemeFollowedHyperlink ThemeColorSlot = "folHlink"
)

var themeColorSlotOrder = []ThemeColorSlot{
	ThemeDark1, ThemeLight1, ThemeDark2, ThemeLight2,
	ThemeAccent1, ThemeAccent2, ThemeAccent3, ThemeAccent4, ThemeAccent5, ThemeAccent6,
	ThemeHyperlink, ThemeFollowedHyperlink,
}

// Theme is what ExtractTheme reports: the effective hex color (no '#') for
// each of the 12 scheme slots, and the major/minor Latin typefaces.
type Theme struct {
	Name       string
	Colors     map[ThemeColorSlot]string // hex RRGGBB, uppercase, no '#'
	MajorLatin string
	MinorLatin string
}

// ThemePatch is what ApplyTheme changes. Colors maps a subset of the 12
// slots to new hex values (RRGGBB, '#' optional) — slots not present are
// left untouched. MajorLatinFont/MinorLatinFont, when non-empty, replace
// the heading/body font; leave empty to not change that font.
type ThemePatch struct {
	Colors         map[ThemeColorSlot]string
	MajorLatinFont string
	MinorLatinFont string
}

const relTypeThemeSuffix = "/relationships/theme"

// themePartFor resolves the theme part's path via document.xml.rels'
// theme relationship (never hardcode "theme1.xml" — resolve it the same
// defensive way xlsxpatch resolves worksheet parts).
func themePartFor(zr *zip.Reader) (string, error) {
	relsRaw, err := readPart(zr, docRelsPart)
	if err != nil {
		return "", fmt.Errorf("docxpatch: no theme relationship (missing %s): %w", docRelsPart, err)
	}
	target, ok := findRelationshipTarget(string(relsRaw), relTypeThemeSuffix)
	if !ok {
		return "", fmt.Errorf("docxpatch: document has no theme part (no theme relationship in %s)", docRelsPart)
	}
	return resolveWordRelTarget(target), nil
}

// resolveWordRelTarget resolves a Target from word/_rels/document.xml.rels
// (relative to word/, e.g. "theme/theme1.xml") to a full zip part path.
func resolveWordRelTarget(target string) string {
	if strings.HasPrefix(target, "/") {
		return strings.TrimPrefix(target, "/")
	}
	return "word/" + target
}

func findRelationshipTarget(relsXML, typeSuffix string) (string, bool) {
	dec := xml.NewDecoder(strings.NewReader(relsXML))
	for {
		tok, err := dec.Token()
		if err != nil {
			return "", false
		}
		se, isStart := tok.(xml.StartElement)
		if !isStart || se.Name.Local != "Relationship" {
			continue
		}
		var typ, target string
		for _, a := range se.Attr {
			switch a.Name.Local {
			case "Type":
				typ = a.Value
			case "Target":
				target = a.Value
			}
		}
		if strings.HasSuffix(typ, typeSuffix) {
			return target, true
		}
	}
}

// ---- extract ----

type themeColorVal struct {
	Srgb *struct {
		Val string `xml:"val,attr"`
	} `xml:"srgbClr"`
	Sys *struct {
		LastClr string `xml:"lastClr,attr"`
	} `xml:"sysClr"`
}

func (v themeColorVal) hex() string {
	if v.Srgb != nil {
		return strings.ToUpper(v.Srgb.Val)
	}
	if v.Sys != nil {
		return strings.ToUpper(v.Sys.LastClr)
	}
	return ""
}

type themeXML struct {
	Name     string `xml:"name,attr"`
	Elements struct {
		ClrScheme struct {
			Dk1      themeColorVal `xml:"dk1"`
			Lt1      themeColorVal `xml:"lt1"`
			Dk2      themeColorVal `xml:"dk2"`
			Lt2      themeColorVal `xml:"lt2"`
			Accent1  themeColorVal `xml:"accent1"`
			Accent2  themeColorVal `xml:"accent2"`
			Accent3  themeColorVal `xml:"accent3"`
			Accent4  themeColorVal `xml:"accent4"`
			Accent5  themeColorVal `xml:"accent5"`
			Accent6  themeColorVal `xml:"accent6"`
			Hlink    themeColorVal `xml:"hlink"`
			FolHlink themeColorVal `xml:"folHlink"`
		} `xml:"clrScheme"`
		FontScheme struct {
			MajorFont struct {
				Latin struct {
					Typeface string `xml:"typeface,attr"`
				} `xml:"latin"`
			} `xml:"majorFont"`
			MinorFont struct {
				Latin struct {
					Typeface string `xml:"typeface,attr"`
				} `xml:"latin"`
			} `xml:"minorFont"`
		} `xml:"fontScheme"`
	} `xml:"themeElements"`
}

// ExtractTheme reads the document's theme colors and major/minor fonts.
func ExtractTheme(docx []byte) (Theme, error) {
	zr, err := zip.NewReader(bytes.NewReader(docx), int64(len(docx)))
	if err != nil {
		return Theme{}, fmt.Errorf("docxpatch: not a readable .docx: %w", err)
	}
	part, err := themePartFor(zr)
	if err != nil {
		return Theme{}, err
	}
	raw, err := readPart(zr, part)
	if err != nil {
		return Theme{}, fmt.Errorf("docxpatch: theme part %q referenced but not found: %w", part, err)
	}
	var t themeXML
	if err := xml.Unmarshal(raw, &t); err != nil {
		return Theme{}, fmt.Errorf("docxpatch: parse theme: %w", err)
	}
	c := t.Elements.ClrScheme
	return Theme{
		Name: t.Name,
		Colors: map[ThemeColorSlot]string{
			ThemeDark1: c.Dk1.hex(), ThemeLight1: c.Lt1.hex(),
			ThemeDark2: c.Dk2.hex(), ThemeLight2: c.Lt2.hex(),
			ThemeAccent1: c.Accent1.hex(), ThemeAccent2: c.Accent2.hex(),
			ThemeAccent3: c.Accent3.hex(), ThemeAccent4: c.Accent4.hex(),
			ThemeAccent5: c.Accent5.hex(), ThemeAccent6: c.Accent6.hex(),
			ThemeHyperlink: c.Hlink.hex(), ThemeFollowedHyperlink: c.FolHlink.hex(),
		},
		MajorLatin: t.Elements.FontScheme.MajorFont.Latin.Typeface,
		MinorLatin: t.Elements.FontScheme.MinorFont.Latin.Typeface,
	}, nil
}

// ---- apply (surgical patch) ----

var hexColorRe = regexp.MustCompile(`^#?[0-9a-fA-F]{6}$`)

// ApplyTheme patches only the requested color slots and/or fonts within
// the EXISTING theme1.xml, leaving everything else in the document byte-
// identical. Errors (and changes nothing) if the document has no theme
// part, if a requested color isn't a valid 6-digit hex, or if a named slot
// isn't found in the theme XML (a malformed/non-standard theme part).
func ApplyTheme(docx []byte, patch ThemePatch) ([]byte, error) {
	if len(patch.Colors) == 0 && patch.MajorLatinFont == "" && patch.MinorLatinFont == "" {
		return nil, fmt.Errorf("docxpatch: theme patch is empty")
	}
	for slot, hex := range patch.Colors {
		if !hexColorRe.MatchString(hex) {
			return nil, fmt.Errorf("docxpatch: theme color %s: %q is not a 6-digit hex color", slot, hex)
		}
		if !isKnownThemeSlot(slot) {
			return nil, fmt.Errorf("docxpatch: unknown theme color slot %q", slot)
		}
	}

	zr, err := zip.NewReader(bytes.NewReader(docx), int64(len(docx)))
	if err != nil {
		return nil, fmt.Errorf("docxpatch: not a readable .docx: %w", err)
	}
	part, err := themePartFor(zr)
	if err != nil {
		return nil, err
	}
	raw, err := readPart(zr, part)
	if err != nil {
		return nil, fmt.Errorf("docxpatch: theme part %q referenced but not found: %w", part, err)
	}
	s := string(raw)

	// Colors: replace the ENTIRE inner content of the slot's element with
	// a plain <a:srgbClr val="HEX"/> — matches what Word itself does when
	// a user customizes a theme color that started as a sysClr default
	// (dk1/lt1 typically are; the rest are already srgbClr).
	for slot, hex := range patch.Colors {
		hex = strings.ToUpper(strings.TrimPrefix(hex, "#"))
		newS, err := replaceThemeElementInner(s, string(slot), fmt.Sprintf(`<a:srgbClr val="%s"/>`, hex))
		if err != nil {
			return nil, err
		}
		s = newS
	}

	if patch.MajorLatinFont != "" {
		newS, err := replaceFontLatinTypeface(s, "majorFont", patch.MajorLatinFont)
		if err != nil {
			return nil, err
		}
		s = newS
	}
	if patch.MinorLatinFont != "" {
		newS, err := replaceFontLatinTypeface(s, "minorFont", patch.MinorLatinFont)
		if err != nil {
			return nil, err
		}
		s = newS
	}

	return ApplyPatch(docx, Patch{Replace: map[string][]byte{part: []byte(s)}})
}

func isKnownThemeSlot(slot ThemeColorSlot) bool {
	for _, s := range themeColorSlotOrder {
		if s == slot {
			return true
		}
	}
	return false
}

// replaceThemeElementInner replaces the content of <a:NAME>...</a:NAME>
// (the FIRST occurrence — a:clrScheme's slots are each unique within the
// theme part) with newInner.
func replaceThemeElementInner(s, name, newInner string) (string, error) {
	re := regexp.MustCompile(`(?s)<a:` + regexp.QuoteMeta(name) + `>.*?</a:` + regexp.QuoteMeta(name) + `>`)
	loc := re.FindStringIndex(s)
	if loc == nil {
		return "", fmt.Errorf("docxpatch: theme has no <a:%s> element (non-standard theme part)", name)
	}
	replacement := "<a:" + name + ">" + newInner + "</a:" + name + ">"
	return s[:loc[0]] + replacement + s[loc[1]:], nil
}

// replaceFontLatinTypeface replaces the typeface attribute of the FIRST
// <a:latin .../> inside <a:majorFont>...</a:majorFont> or
// <a:minorFont>...</a:minorFont> — scoped to that block so it can never
// accidentally touch the ea/cs typefaces or the other font role.
func replaceFontLatinTypeface(s, fontRole, newTypeface string) (string, error) {
	blockRe := regexp.MustCompile(`(?s)<a:` + fontRole + `>.*?</a:` + fontRole + `>`)
	loc := blockRe.FindStringIndex(s)
	if loc == nil {
		return "", fmt.Errorf("docxpatch: theme has no <a:%s> element (non-standard theme part)", fontRole)
	}
	block := s[loc[0]:loc[1]]
	latinRe := regexp.MustCompile(`<a:latin typeface="[^"]*"`)
	if !latinRe.MatchString(block) {
		return "", fmt.Errorf("docxpatch: <a:%s> has no <a:latin> typeface (non-standard theme part)", fontRole)
	}
	newBlock := latinRe.ReplaceAllString(block, fmt.Sprintf(`<a:latin typeface=%q`, newTypeface))
	return s[:loc[0]] + newBlock + s[loc[1]:], nil
}

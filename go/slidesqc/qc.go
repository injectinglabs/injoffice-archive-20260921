package slidesqc

import (
	"fmt"
	"math"
	"regexp"
	"strings"
)

// ---- Stage geometry, all in "cqw" (% of slide width) — mirrors DeckView.tsx
// and MUST stay numerically identical to qc.ts. See that file for the full
// rationale: there is no real box geometry in the HTML-first renderer this
// audits, so overflow/overlap are estimated from a character-width text
// heuristic against each slide kind's known stacked-element layout, not
// measured from real boxes.

// NOTE: these are typed float64 constants, not untyped integer constants —
// stageH's division must be true (56.25), not truncating integer division,
// to match qc.ts' floating-point arithmetic exactly.
const (
	stageW   float64 = 100
	stageH   float64 = (stageW * 9) / 16 // 16:9 aspect, height in the same cqw unit
	margin   float64 = 6                 // SlideView's inset on all 4 sides
	contentW float64 = stageW - 2*margin
	contentH float64 = stageH - 2*margin // ≈ 44.25
)

// pageNumberReserveH is the reserved bottom strip for the "N / total"
// page-number chip every non-title kind renders — content stacking into
// this band risks visually colliding with it.
const pageNumberReserveH = 4

// accentBarReserveH is the reserved bottom strip for the accent bar under
// title/closing slides.
const accentBarReserveH = 3

// avgCharWidth is the average glyph width as a fraction of font-size, for a
// proportional display/body sans stack — a standard typographic rule of
// thumb (not measured per font; an estimate, as throughout this package).
const avgCharWidth = 0.52

var (
	boldRe   = regexp.MustCompile(`\*\*([^*]+)\*\*`)
	italicRe = regexp.MustCompile(`\*([^*]+)\*`)
)

// plainLength strips the bold/italic emphasis markers DeckView's renderInline
// consumes — they don't render as glyphs, so they shouldn't count toward
// width. Counted in runes (not bytes) to match JS string .length semantics
// closely enough for this heuristic's purposes.
func plainLength(text string) int {
	stripped := italicRe.ReplaceAllString(boldRe.ReplaceAllString(text, "$1"), "$1")
	return len([]rune(stripped))
}

// EstimateWrappedLines returns the estimated wrapped line count for one
// paragraph of text at fontSizeCqw, inside a box boxWidthCqw wide. Explicit
// '\n' segments (pre-wrap fields like body/notes) each wrap independently.
// Empty text = 0 lines.
func EstimateWrappedLines(text string, boxWidthCqw, fontSizeCqw float64) int {
	if text == "" {
		return 0
	}
	charsPerLine := int(math.Max(1, math.Floor(boxWidthCqw/(fontSizeCqw*avgCharWidth))))
	lines := 0
	for _, segment := range strings.Split(text, "\n") {
		l := plainLength(segment)
		if l == 0 {
			lines++
		} else {
			lines += int(math.Ceil(float64(l) / float64(charsPerLine)))
		}
	}
	return lines
}

// textBlockHeight is the estimated rendered height (cqw) of one paragraph:
// lines × fontSize × lineHeight.
func textBlockHeight(text string, boxWidthCqw, fontSizeCqw, lineHeight float64) float64 {
	return float64(EstimateWrappedLines(text, boxWidthCqw, fontSizeCqw)) * fontSizeCqw * lineHeight
}

// bulletsHeight is the bulleted-list height: each item wraps independently
// within boxWidthCqw (minus the list's left padding), items separated by the
// template's gap.
func bulletsHeight(items []string, boxWidthCqw float64) float64 {
	if len(items) == 0 {
		return 0
	}
	const fontSizeCqw = 2.6
	const gap = 1.5
	innerW := boxWidthCqw - 2.4 // ul padding-left
	linesTotal := 0.0
	for _, item := range items {
		linesTotal += math.Max(1, float64(EstimateWrappedLines(item, innerW, fontSizeCqw)))
	}
	return linesTotal*fontSizeCqw*1.55 + gap*float64(len(items)-1)
}

func eyebrowHeight(eyebrow string) float64 {
	if eyebrow == "" {
		return 0
	}
	return 1.4*1.2 + 1.6 // font-size*lineHeight-ish + its own margin-bottom
}

// kindBudget is the estimated layout budget for one slide.
type kindBudget struct {
	// contentH is the estimated stacked content height (cqw) for this slide.
	contentH float64
	// reserve is the bottom reserve this kind's fixed decoration needs
	// (cqw), if any.
	reserve      float64
	reserveLabel string
	// fieldNotes are per-field notes for width-driven issues worth calling
	// out individually (e.g. unbalanced two-col columns).
	fieldNotes []string
}

func budgetFor(slide SlideSpec) kindBudget {
	var notes []string
	switch slide.Kind {
	case KindTitle:
		titleW := contentW * 0.8
		subW := contentW * 0.62
		h := eyebrowHeight(slide.Eyebrow) + textBlockHeight(slide.Title, titleW, 7, 1.05)
		if slide.Subtitle != "" {
			h += textBlockHeight(slide.Subtitle, subW, 2.6, 1.55) + 2.4
		}
		if slide.Body != "" {
			h += textBlockHeight(slide.Body, subW, 2, 1.55) + 1.6
		}
		return kindBudget{contentH: h, reserve: accentBarReserveH, reserveLabel: "the title bar accent", fieldNotes: notes}

	case KindSection:
		titleW := contentW * 0.78
		subW := contentW * 0.64
		h := eyebrowHeight(slide.Eyebrow) + textBlockHeight(slide.Title, titleW, 5.4, 1.05)
		if slide.Subtitle != "" || slide.Body != "" {
			text := slide.Subtitle
			if text == "" {
				text = slide.Body
			}
			h += textBlockHeight(text, subW, 2.3, 1.55) + 2
		}
		return kindBudget{contentH: h, reserve: pageNumberReserveH, reserveLabel: "the page-number chip", fieldNotes: notes}

	case KindQuote:
		quoteW := contentW*0.8 - 5 // minus the callout box's own horizontal padding (5cqw)
		h := 8.0 + textBlockHeight(slide.Quote, quoteW, 3.4, 1.4)
		if slide.Subtitle != "" {
			h += textBlockHeight(slide.Subtitle, quoteW, 1.6, 1.4) + 2.2
		}
		return kindBudget{contentH: h, reserve: pageNumberReserveH, reserveLabel: "the page-number chip", fieldNotes: notes}

	case KindTwoCol:
		colW := (contentW - 5) / 2
		leftH := bulletsHeight(slide.Bullets, colW)
		rightH := bulletsHeight(slide.BulletsRight, colW)
		if math.Abs(leftH-rightH) > 12 {
			notes = append(notes, fmt.Sprintf(
				"two-col columns are visually unbalanced (left ≈%scqw tall vs right ≈%scqw) — consider redistributing items",
				fmtNum(leftH), fmtNum(rightH),
			))
		}
		colH := math.Max(leftH, rightH)
		h := eyebrowHeight(slide.Eyebrow) + textBlockHeight(slide.Title, contentW, 4, 1.05) + 3 + colH
		return kindBudget{contentH: h, reserve: pageNumberReserveH, reserveLabel: "the page-number chip", fieldNotes: notes}

	case KindClosing:
		title := slide.Title
		if title == "" {
			title = "Thank you"
		}
		h := textBlockHeight(title, contentW, 5.6, 1.05)
		if slide.Subtitle != "" {
			h += textBlockHeight(slide.Subtitle, contentW, 2.2, 1.55) + 2
		}
		return kindBudget{contentH: h, reserve: accentBarReserveH + 3.4, reserveLabel: "the closing accent bar", fieldNotes: notes}

	case KindBullets:
		fallthrough
	default:
		bodyW := contentW * 0.8
		h := eyebrowHeight(slide.Eyebrow)
		if slide.Title != "" {
			h += textBlockHeight(slide.Title, contentW, 4, 1.05) + 3
		}
		if slide.Body != "" {
			h += textBlockHeight(slide.Body, bodyW, 2.1, 1.55) + 2.4
		}
		h += bulletsHeight(slide.Bullets, contentW)
		return kindBudget{contentH: h, reserve: pageNumberReserveH, reserveLabel: "the page-number chip", fieldNotes: notes}
	}
}

// fmtNum matches JS's Number.prototype.toFixed(1) rounding behavior (round
// half AWAY from zero) rather than Go's fmt.Sprintf("%.1f", n), which rounds
// exact binary midpoints (e.g. 44.25) half TO EVEN (44.2, not 44.3) — a
// real mismatch this package hit in cross-checking against qc.ts, since
// several of these budgets land on exact quarter-cqw sums.
func fmtNum(n float64) string {
	rounded := math.Round(n*10) / 10
	return fmt.Sprintf("%.1f", rounded)
}

// AuditSlide audits one slide. Pure function — same input, same output, no
// rendering.
func AuditSlide(slide SlideSpec, slideIndex int) []QCIssue {
	var issues []QCIssue
	budget := budgetFor(slide)
	push := func(kind QCIssueKind, message string) {
		issues = append(issues, QCIssue{SlideID: slide.ID, SlideIndex: slideIndex, Kind: kind, Message: message})
	}

	for _, note := range budget.fieldNotes {
		push(IssueOverlap, note)
	}

	available := contentH
	switch {
	case budget.contentH > available:
		push(IssueOutOfBounds, fmt.Sprintf(
			"content is estimated at ≈%scqw tall, %scqw past the %scqw slide content area — it will be clipped (shorten text, split into another slide, or reduce bullet count)",
			fmtNum(budget.contentH), fmtNum(budget.contentH-available), fmtNum(available),
		))
	case budget.reserve > 0 && budget.contentH > available-budget.reserve:
		push(IssueOverlap, fmt.Sprintf(
			"content is estimated at ≈%scqw tall, within %scqw of %s — it may visually collide with it",
			fmtNum(budget.contentH), fmtNum(available-budget.contentH), budget.reserveLabel,
		))
	}

	return issues
}

// AuditDeck audits every slide in a deck. An empty (nil) slice = clean deck.
func AuditDeck(deck DeckSpec) []QCIssue {
	var issues []QCIssue
	for i, slide := range deck.Slides {
		issues = append(issues, AuditSlide(slide, i)...)
	}
	return issues
}

// FormatDeckAudit formats issues as text suitable for feeding back into an
// AI revise prompt (mirrors qc.ts' formatDeckAudit: a compact block a model
// can act on).
func FormatDeckAudit(issues []QCIssue) string {
	if len(issues) == 0 {
		return "<deck-audit>Passed: no estimated overflow/overlap issues.</deck-audit>"
	}
	lines := make([]string, len(issues))
	for i, iss := range issues {
		lines[i] = fmt.Sprintf("- Slide %d (%s) [%s]: %s", iss.SlideIndex+1, iss.SlideID, iss.Kind, iss.Message)
	}
	return fmt.Sprintf("<deck-audit>Found %d issue(s):\n%s\n</deck-audit>", len(issues), strings.Join(lines, "\n"))
}

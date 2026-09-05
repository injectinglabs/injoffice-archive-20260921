// Package slidesqc is a deterministic (no LLM) layout QC pass for
// @injoffice/slides' DeckSpec — a Go port of packages/slides/src/qc.ts, kept
// byte-for-byte equivalent in behavior (same constants, same per-kind
// budgets, same estimate) so a Go host and the TS renderer agree on what
// "probably overflows" means. See qc.ts for the full design rationale; this
// file only restates the parts a Go caller needs.
//
// DeckSpec here is a deliberately narrow mirror of the TS type: only the
// fields the audit actually reads (id/kind/eyebrow/title/subtitle/bullets/
// bulletsRight/quote/body). A host decodes its own JSON into this shape (or
// a superset embedding it) — this package has no opinion on transport.
package slidesqc

// DeckSpec is one deck: an id/title plus its ordered slides. Theme and other
// TS-side fields are irrelevant to layout QC and intentionally omitted.
type DeckSpec struct {
	ID     string      `json:"id"`
	Title  string      `json:"title"`
	Slides []SlideSpec `json:"slides"`
}

// SlideKind mirrors the TS union of the same name.
type SlideKind string

const (
	KindTitle   SlideKind = "title"
	KindSection SlideKind = "section"
	KindBullets SlideKind = "bullets"
	KindTwoCol  SlideKind = "two-col"
	KindQuote   SlideKind = "quote"
	KindClosing SlideKind = "closing"
)

// SlideSpec is one slide. Only the fields the layout-budget estimate reads;
// unknown/irrelevant TS fields (colTitles, notes, shapeOverrides, ...) have
// no Go equivalent here because auditSlide never looks at them.
type SlideSpec struct {
	ID           string    `json:"id"`
	Kind         SlideKind `json:"kind"`
	Eyebrow      string    `json:"eyebrow"`
	Title        string    `json:"title"`
	Subtitle     string    `json:"subtitle"`
	Bullets      []string  `json:"bullets"`
	BulletsRight []string  `json:"bulletsRight"`
	Quote        string    `json:"quote"`
	Body         string    `json:"body"`
}

// QCIssueKind mirrors the TS union of the same name.
type QCIssueKind string

const (
	IssueOverflow    QCIssueKind = "overflow"
	IssueOverlap     QCIssueKind = "overlap"
	IssueOutOfBounds QCIssueKind = "out-of-bounds"
)

// QCIssue is one flagged problem on one slide. Message is human-readable,
// meant to be fed straight back into a revise/fix prompt.
type QCIssue struct {
	SlideID    string      `json:"slideId"`
	SlideIndex int         `json:"slideIndex"`
	Kind       QCIssueKind `json:"kind"`
	Message    string      `json:"message"`
}

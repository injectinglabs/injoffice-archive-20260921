package slidesqc

import (
	"strings"
	"testing"
)

func TestEstimateWrappedLines(t *testing.T) {
	if got := EstimateWrappedLines("", 80, 3); got != 0 {
		t.Errorf("empty text: got %d, want 0", got)
	}
	if got := EstimateWrappedLines("Hi", 80, 3); got != 1 {
		t.Errorf("short text: got %d, want 1", got)
	}

	short := EstimateWrappedLines(strings.Repeat("a ", 5), 40, 3)
	long := EstimateWrappedLines(strings.Repeat("a ", 50), 40, 3)
	if !(long > short) {
		t.Errorf("expected long text to wrap to more lines than short: long=%d short=%d", long, short)
	}

	oneLine := EstimateWrappedLines("hello world", 80, 3)
	twoLines := EstimateWrappedLines("hello\nworld", 80, 3)
	if !(twoLines >= oneLine) {
		t.Errorf("explicit newline should not reduce line count: oneLine=%d twoLines=%d", oneLine, twoLines)
	}
	if got := EstimateWrappedLines("a\nb\nc", 80, 3); got != 3 {
		t.Errorf("three explicit segments: got %d, want 3", got)
	}

	plain := EstimateWrappedLines("hello world", 20, 3)
	bold := EstimateWrappedLines("**hello** world", 20, 3)
	if bold != plain {
		t.Errorf("markdown emphasis markers should not count toward width: plain=%d bold=%d", plain, bold)
	}
}

func hasKind(issues []QCIssue, kind QCIssueKind) bool {
	for _, i := range issues {
		if i.Kind == kind {
			return true
		}
	}
	return false
}

func TestAuditSlide_NormalTitleSlideHasNoIssues(t *testing.T) {
	s := SlideSpec{ID: "s1", Kind: KindTitle, Title: "Quarterly Update", Subtitle: "Q3 2026 results"}
	if issues := AuditSlide(s, 0); len(issues) != 0 {
		t.Errorf("expected no issues, got %+v", issues)
	}
}

func TestAuditSlide_EmptySlideHasNoIssues(t *testing.T) {
	s := SlideSpec{ID: "s1", Kind: KindBullets}
	if issues := AuditSlide(s, 0); len(issues) != 0 {
		t.Errorf("expected no issues, got %+v", issues)
	}
}

func TestAuditSlide_FlagsOutOfBoundsForManyLongBullets(t *testing.T) {
	bullets := make([]string, 30)
	for i := range bullets {
		bullets[i] = "This is a fairly long bullet point number with real substance to it"
	}
	s := SlideSpec{ID: "s1", Kind: KindBullets, Title: "Everything", Bullets: bullets}
	issues := AuditSlide(s, 2)
	if !hasKind(issues, IssueOutOfBounds) {
		t.Errorf("expected an out-of-bounds issue, got %+v", issues)
	}
	if len(issues) == 0 || issues[0].SlideIndex != 2 {
		t.Errorf("expected first issue slideIndex 2, got %+v", issues)
	}
	if len(issues) == 0 || issues[0].SlideID != "s1" {
		t.Errorf("expected first issue slideId s1, got %+v", issues)
	}
}

func TestAuditSlide_FlagsOutOfBoundsForExtremelyLongTitle(t *testing.T) {
	s := SlideSpec{ID: "s1", Kind: KindTitle, Title: strings.Repeat("A ", 200)}
	issues := AuditSlide(s, 0)
	if !hasKind(issues, IssueOutOfBounds) {
		t.Errorf("expected an out-of-bounds issue, got %+v", issues)
	}
}

func TestAuditSlide_FlagsNearMarginClosingSlide(t *testing.T) {
	s := SlideSpec{
		ID:       "s1",
		Kind:     KindClosing,
		Title:    "Thank You Everyone For Coming Together Today",
		Subtitle: strings.Repeat("We could not have done it without each and every one of you supporting this launch. ", 7),
	}
	issues := AuditSlide(s, 0)
	if len(issues) == 0 {
		t.Errorf("expected at least one issue (overlap or out-of-bounds), got none")
	}
}

func TestAuditSlide_FlagsUnbalancedTwoCol(t *testing.T) {
	rightBullets := make([]string, 12)
	for i := range rightBullets {
		rightBullets[i] = "A much longer right-column bullet point that goes on and on and on"
	}
	s := SlideSpec{
		ID: "s1", Kind: KindTwoCol, Title: "Comparison",
		Bullets:      []string{"short"},
		BulletsRight: rightBullets,
	}
	issues := AuditSlide(s, 0)
	found := false
	for _, i := range issues {
		if i.Kind == IssueOverlap && strings.Contains(i.Message, "unbalanced") {
			found = true
		}
	}
	if !found {
		t.Errorf("expected an unbalanced two-col overlap issue, got %+v", issues)
	}
}

func TestAuditSlide_ReasonableTwoColHasNoIssues(t *testing.T) {
	s := SlideSpec{
		ID: "s1", Kind: KindTwoCol, Title: "Before / After",
		Bullets:      []string{"Manual process", "Slow turnaround", "Error prone"},
		BulletsRight: []string{"Automated", "Fast", "Reliable"},
	}
	if issues := AuditSlide(s, 0); len(issues) != 0 {
		t.Errorf("expected no issues, got %+v", issues)
	}
}

func TestAuditSlide_ReasonableQuoteHasNoIssues(t *testing.T) {
	s := SlideSpec{ID: "s1", Kind: KindQuote, Quote: "Simplicity is the ultimate sophistication.", Subtitle: "Leonardo da Vinci"}
	if issues := AuditSlide(s, 0); len(issues) != 0 {
		t.Errorf("expected no issues, got %+v", issues)
	}
}

func TestAuditSlide_FlagsOverlongQuote(t *testing.T) {
	s := SlideSpec{ID: "s1", Kind: KindQuote, Quote: strings.Repeat("Lorem ipsum dolor sit amet. ", 40)}
	if issues := AuditSlide(s, 0); len(issues) == 0 {
		t.Errorf("expected at least one issue, got none")
	}
}

func TestAuditDeck_CleanAndDirty(t *testing.T) {
	clean := DeckSpec{
		ID: "d1", Title: "Deck",
		Slides: []SlideSpec{
			{ID: "a", Kind: KindTitle, Title: "Hello"},
			{ID: "b", Kind: KindBullets, Title: "Points", Bullets: []string{"One", "Two"}},
		},
	}
	if issues := AuditDeck(clean); len(issues) != 0 {
		t.Errorf("expected clean deck to have no issues, got %+v", issues)
	}

	longBullets := make([]string, 40)
	for i := range longBullets {
		longBullets[i] = strings.Repeat("bullet text ", 10)
	}
	dirty := DeckSpec{
		ID: "d2", Title: "Deck",
		Slides: []SlideSpec{
			{ID: "a", Kind: KindTitle, Title: "Hello"},
			{ID: "b", Kind: KindBullets, Title: "Points", Bullets: longBullets},
		},
	}
	issues := AuditDeck(dirty)
	if len(issues) == 0 {
		t.Fatalf("expected dirty deck to have issues")
	}
	for _, i := range issues {
		if i.SlideID != "b" {
			t.Errorf("expected all issues on slide b, got %+v", i)
		}
		if i.SlideIndex != 1 {
			t.Errorf("expected all issues at slideIndex 1, got %+v", i)
		}
	}
}

func TestFormatDeckAudit_Pass(t *testing.T) {
	if got := FormatDeckAudit(nil); !strings.Contains(got, "Passed") {
		t.Errorf("expected pass message, got %q", got)
	}
}

func TestFormatDeckAudit_ListsIssues(t *testing.T) {
	text := FormatDeckAudit([]QCIssue{{SlideID: "x", SlideIndex: 4, Kind: IssueOverflow, Message: "too tall"}})
	if !strings.Contains(text, "Slide 5") {
		t.Errorf("expected 'Slide 5' in %q", text)
	}
	if !strings.Contains(text, "[overflow]") {
		t.Errorf("expected '[overflow]' in %q", text)
	}
	if !strings.Contains(text, "too tall") {
		t.Errorf("expected message text in %q", text)
	}
}

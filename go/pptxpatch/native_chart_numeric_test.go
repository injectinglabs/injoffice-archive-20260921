package pptxpatch

import (
	"strings"
	"testing"
)

func TestNativeChartDecimalSourceLexemes(t *testing.T) {
	for _, text := range []string{"0", "-0", "+0", "1.25", ".5", "1.", "-12.34e-5", "0001.000", "1E+100", "1e-100", "12345678901234567890123456789012"} {
		d, err := parseNativeChartDecimal(text)
		if err != nil || d.lexeme != text {
			t.Fatalf("source lexeme changed/refused %q: %v", text, err)
		}
	}
	for _, text := range []string{"", " 1", "1 ", "1\n", "NaN", "INF", "Infinity", "0x10", "1_000", "1,5", ".", "1e", "1e101", "1e-101", strings.Repeat("1", 33), "1e99999999999999", strings.Repeat("0", 129)} {
		if _, err := parseNativeChartDecimal(text); err == nil {
			t.Fatalf("invalid/budgeted lexeme accepted %q", text)
		}
	}
}
func TestNativeChartDecimalExactComparison(t *testing.T) {
	for _, c := range []struct {
		a, b  string
		order int
	}{{"-0", "0", 0}, {"1.25e2", "125", 0}, {".001", "1e-3", 0}, {"1000000000000.000000000000000001", "1000000000000.000000000000000002", -1}, {"-1e100", "-1e99", -1}, {"1e-100", "0", 1}} {
		a, e1 := parseNativeChartDecimal(c.a)
		b, e2 := parseNativeChartDecimal(c.b)
		if e1 != nil || e2 != nil {
			t.Fatal(e1, e2)
		}
		if a.compare(b) != c.order {
			t.Fatalf("incorrect exact order %s vs %s", c.a, c.b)
		}
		before := a.coefficient.String()
		a.compare(b)
		if a.coefficient.String() != before {
			t.Fatal("comparison mutated number")
		}
	}
}

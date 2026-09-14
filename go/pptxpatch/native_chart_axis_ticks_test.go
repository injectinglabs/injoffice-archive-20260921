package pptxpatch

import "testing"

func TestNativeChartAxisExactTicks(t *testing.T) {
	for _, test := range []struct {
		low, high, step string
		count           int
	}{{"-1", "1", ".25", 9}, {"-1.00", ".9", "0.5", 4}, {"0", "255e-100", "1e-100", 256}, {"0", "1", "2", 1}, {"10000000000000000000000000000000", "10000000000000000000000000000002", "1", 3}} {
		count, ok := nativeChartAxisTickCount(test.low, test.high, test.step)
		if !ok || count != test.count {
			t.Fatalf("%+v ->%d,%v", test, count, ok)
		}
	}
	for _, args := range [][3]string{{"0", "256", "1"}, {"-.1", "1", ".3"}, {"0", "1", "0"}, {"0", "1", "-1"}, {"1", "1", "1"}, {"0", "1e101", "1"}, {"0", "1e100", "1e-100"}} {
		if _, ok := nativeChartAxisTickCount(args[0], args[1], args[2]); ok {
			t.Fatalf("invalid grid accepted: %v", args)
		}
	}
	for _, code := range []string{"0", "0.0", "0.000000"} {
		if !nativeChartFixedFormat(code) {
			t.Fatal(code)
		}
	}
	for _, code := range []string{"General", "0.", "0.0000000", "#,##0", "0%", "0;0", "0.0\n"} {
		if nativeChartFixedFormat(code) {
			t.Fatal(code)
		}
	}
}

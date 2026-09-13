package pptxpatch

import "math/big"

// This first explicit tick grid is anchored unambiguously: min is an exact
// integer multiple of majorUnit. No automatic interval or binary-float loop.
func nativeChartAxisTickCount(minimum, maximum, majorUnit string) (int, bool) {
	low, e := parseNativeChartDecimal(minimum)
	if e != nil {
		return 0, false
	}
	high, e := parseNativeChartDecimal(maximum)
	if e != nil {
		return 0, false
	}
	step, e := parseNativeChartDecimal(majorUnit)
	if e != nil || step.coefficient.Sign() <= 0 {
		return 0, false
	}
	exponent := min(low.exponent, high.exponent, step.exponent)
	a, b, s := low.scaled(exponent), high.scaled(exponent), step.scaled(exponent)
	if a.Cmp(b) >= 0 || new(big.Int).Rem(a, s).Sign() != 0 {
		return 0, false
	}
	count := new(big.Int).Quo(new(big.Int).Sub(b, a), s)
	count.Add(count, big.NewInt(1))
	if count.Sign() <= 0 || count.Cmp(big.NewInt(256)) > 0 {
		return 0, false
	}
	// Iterate only within the independently fixed small bound; no narrowing cast.
	for i := 1; i <= 256; i++ {
		if count.Cmp(big.NewInt(int64(i))) == 0 {
			return i, true
		}
	}
	return 0, false
}
func nativeChartFixedFormat(code string) bool {
	if code == "0" {
		return true
	}
	if len(code) < 3 || len(code) > 8 || code[:2] != "0." {
		return false
	}
	for _, c := range code[2:] {
		if c != '0' {
			return false
		}
	}
	return true
}

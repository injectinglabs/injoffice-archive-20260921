package pptxpatch

import (
	"fmt"
	"math/big"
	"regexp"
	"strconv"
)

// Keep source decimals as lexemes: the native JSON contract intentionally has
// integer-only numeric fields. Exact decimal arithmetic avoids cancellation in
// nearby explicit axis endpoints and never rescales source values into integers.
const nativeChartMaxNumericLexemeBytes = 128
const nativeChartMaxDecimalDigits = 32
const nativeChartMaxDecimalExponent = 100

var nativeChartDecimalPattern = regexp.MustCompile(`^([+-]?)(?:([0-9]+)(?:\.([0-9]*))?|\.([0-9]+))(?:[eE]([+-]?[0-9]+))?$`)

type nativeChartDecimal struct {
	lexeme      string
	coefficient *big.Int
	exponent    int
}

func parseNativeChartDecimal(lexeme string) (nativeChartDecimal, error) {
	if len(lexeme) == 0 || len(lexeme) > nativeChartMaxNumericLexemeBytes {
		return nativeChartDecimal{}, fmt.Errorf("chart numeric lexeme budget exceeded")
	}
	match := nativeChartDecimalPattern.FindStringSubmatch(lexeme)
	if match == nil {
		return nativeChartDecimal{}, fmt.Errorf("invalid chart decimal lexeme")
	}
	integer, fraction := match[2], match[3]
	if match[4] != "" {
		fraction = match[4]
	}
	digits := integer + fraction
	if len(digits) > nativeChartMaxDecimalDigits {
		return nativeChartDecimal{}, fmt.Errorf("chart decimal digit budget exceeded")
	}
	exponent := 0
	if match[5] != "" {
		var err error
		exponent, err = strconv.Atoi(match[5])
		if err != nil || exponent < -nativeChartMaxDecimalExponent || exponent > nativeChartMaxDecimalExponent {
			return nativeChartDecimal{}, fmt.Errorf("chart decimal exponent budget exceeded")
		}
	}
	coefficient, ok := new(big.Int).SetString(digits, 10)
	if !ok {
		return nativeChartDecimal{}, fmt.Errorf("invalid chart decimal digits")
	}
	if match[1] == "-" {
		coefficient.Neg(coefficient)
	}
	return nativeChartDecimal{lexeme: lexeme, coefficient: coefficient, exponent: exponent - len(fraction)}, nil
}
func (d nativeChartDecimal) scaled(exponent int) *big.Int {
	value := new(big.Int).Set(d.coefficient)
	if difference := d.exponent - exponent; difference > 0 {
		value.Mul(value, new(big.Int).Exp(big.NewInt(10), big.NewInt(int64(difference)), nil))
	}
	return value
}
func (d nativeChartDecimal) compare(other nativeChartDecimal) int {
	exponent := min(d.exponent, other.exponent)
	return d.scaled(exponent).Cmp(other.scaled(exponent))
}

// Used only for source text policy checks; negative zero retains its lexeme.
func (d nativeChartDecimal) isZero() bool { return d.coefficient.Sign() == 0 }

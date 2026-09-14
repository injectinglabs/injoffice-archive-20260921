package pptxpatch

import (
	"fmt"
	"math/big"
)

type nativeChartStackSeries struct {
	Index, Order int64
	Values       []string
}
type nativeChartStackBand struct {
	Index, Order int64
	Lower, Upper []*big.Rat
}

// Private exact boundaries; rational objects never cross native JSON. Values
// remain original lexemes in the source record. Percent boundaries use 1=100%.
// Negative stacked input is not yet qualified for the family-specific rules.
func nativeChartStackBands(series []nativeChartStackSeries, grouping string) ([]nativeChartStackBand, error) {
	if len(series) < 1 || len(series) > nativeChartMaxSeries || (grouping != "standard" && grouping != "stacked" && grouping != "percentStacked") {
		return nil, fmt.Errorf("invalid chart stack profile")
	}
	count := len(series[0].Values)
	if count < 1 || count > nativeChartMaxCategories {
		return nil, fmt.Errorf("invalid chart stack point count")
	}
	seen, orders := map[int64]bool{}, map[int64]bool{}
	values := make([][]*big.Rat, len(series))
	totals := make([]*big.Rat, count)
	for point := range totals {
		totals[point] = new(big.Rat)
	}
	for order, item := range series {
		if item.Index < 0 || item.Index > 4294967295 || seen[item.Index] || item.Order < 0 || item.Order >= int64(len(series)) || orders[item.Order] || len(item.Values) != count {
			return nil, fmt.Errorf("invalid chart stack series alignment")
		}
		seen[item.Index] = true
		orders[item.Order] = true
		values[order] = make([]*big.Rat, count)
		for point, raw := range item.Values {
			decimal, err := parseNativeChartDecimal(raw)
			if err != nil {
				return nil, err
			}
			if grouping != "standard" && decimal.coefficient.Sign() < 0 {
				return nil, fmt.Errorf("negative chart stacking is not qualified")
			}
			value := new(big.Rat).SetInt(decimal.coefficient)
			power := new(big.Int).Exp(big.NewInt(10), big.NewInt(int64(max(decimal.exponent, -decimal.exponent))), nil)
			if decimal.exponent < 0 {
				value.Quo(value, new(big.Rat).SetInt(power))
			} else {
				value.Mul(value, new(big.Rat).SetInt(power))
			}
			values[order][point] = value
			totals[point].Add(totals[point], value)
		}
	}
	accumulated := make([]*big.Rat, count)
	for point := range accumulated {
		accumulated[point] = new(big.Rat)
	}
	bands := make([]nativeChartStackBand, len(series))
	for order, item := range series {
		band := nativeChartStackBand{Index: item.Index, Order: item.Order, Lower: make([]*big.Rat, count), Upper: make([]*big.Rat, count)}
		for point, value := range values[order] {
			lower := new(big.Rat)
			if grouping != "standard" {
				lower.Set(accumulated[point])
			}
			upper := new(big.Rat).Add(lower, value)
			accumulated[point] = new(big.Rat).Set(upper)
			if grouping == "percentStacked" {
				if totals[point].Sign() == 0 {
					lower.SetInt64(0)
					upper.SetInt64(0)
				} else {
					lower.Quo(lower, totals[point])
					upper.Quo(upper, totals[point])
				}
			}
			band.Lower[point], band.Upper[point] = lower, upper
		}
		bands[order] = band
	}
	return bands, nil
}

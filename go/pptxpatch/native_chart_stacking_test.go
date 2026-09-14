package pptxpatch

import (
	"reflect"
	"testing"
)

func TestNativeChartStackBandsExact(t *testing.T) {
	input := []nativeChartStackSeries{{10, 0, []string{"+.10", "1e100", "-0"}}, {11, 1, []string{".20", "2e100", "0"}}}
	before := append([]string(nil), input[0].Values...)
	bands, err := nativeChartStackBands(input, "percentStacked")
	if err != nil {
		t.Fatal(err)
	}
	for _, point := range []int{0, 1} {
		if bands[0].Upper[point].RatString() != "1/3" || bands[1].Lower[point].RatString() != "1/3" || bands[1].Upper[point].RatString() != "1" {
			t.Fatal(bands)
		}
	}
	if bands[1].Upper[2].Sign() != 0 || !reflect.DeepEqual(before, input[0].Values) {
		t.Fatal("zero/source changed")
	}
	bands[0].Upper[0].SetInt64(99)
	if bands[1].Lower[0].RatString() != "1/3" {
		t.Fatal("boundaries share mutable arithmetic")
	}
}
func TestNativeChartStackBandsSignedStandard(t *testing.T) {
	input := []nativeChartStackSeries{{0, 0, []string{"-3", "2"}}, {1, 1, []string{"1", "-4"}}}
	bands, err := nativeChartStackBands(input, "standard")
	if err != nil || bands[1].Lower[0].Sign() != 0 || bands[1].Upper[1].RatString() != "-4" {
		t.Fatal(bands, err)
	}
	for _, grouping := range []string{"stacked", "percentStacked"} {
		if _, err := nativeChartStackBands(input, grouping); err == nil {
			t.Fatal("negative stacking admitted")
		}
	}
}
func TestNativeChartStackBandsRejectAlignment(t *testing.T) {
	for _, input := range [][]nativeChartStackSeries{nil, {{0, 0, nil}}, {{0, 1, []string{"1"}}}, {{0, 0, []string{"1"}}, {0, 1, []string{"2"}}}, {{0, 0, []string{"1"}}, {1, 1, []string{"1", "2"}}}, {{0, 0, []string{"NaN"}}}} {
		if _, err := nativeChartStackBands(input, "stacked"); err == nil {
			t.Fatal("malformed input admitted", input)
		}
	}
}

package xlsxpatch

import (
	"bytes"
	"fmt"
)

// A cell transaction may introduce the engine's existing exact calcPr rewrite.
// A separately classified bookViews record no longer masks that new generic
// record through deduplication. Admit only this byte-proven workbook transition;
// retain the original strict inventory comparison for every other record.
func nativeExpectedUnsupportedAfterCells(original, produced []byte, before *NativeWorkbookV1, transaction NativeWorkbookMutationTransactionV1) ([]NativeWorkbookUnsupportedV1, error) {
	expected := before.Unsupported
	if len(transaction.Cells) == 0 {
		return expected, nil
	}
	source, err := openNativeWorkbookPackage(original)
	if err != nil {
		return nil, err
	}
	route, err := locateWorkbookPartBytes(source.index, func(name string) ([]byte, bool) { b, ok := source.files[name]; return b, ok })
	if err != nil {
		return nil, err
	}
	for _, item := range expected {
		if item.Code == "UNMODELED_WORKBOOK_FEATURE" && item.Capability == "workbook-features" && item.ScopeID == "workbook" && item.PartName != nil && *item.PartName == route.part {
			return expected, nil
		}
	}
	candidate, err := openNativeWorkbookPackage(produced)
	if err != nil {
		return nil, err
	}
	next, err := locateWorkbookPartBytes(candidate.index, func(name string) ([]byte, bool) { b, ok := candidate.files[name]; return b, ok })
	if err != nil {
		return nil, err
	}
	if next.part != route.part {
		return nil, fmt.Errorf("xlsxpatch: native mutation: recalculation workbook route changed")
	}
	wanted, err := workbookWithFullRecalculation(source.files[route.part])
	if err != nil {
		return nil, err
	}
	if bytes.Equal(wanted, source.files[route.part]) || !bytes.Equal(wanted, candidate.files[next.part]) {
		return nil, fmt.Errorf("xlsxpatch: native mutation: new workbook inventory requires exact generated recalculation XML")
	}
	collector := nativeWorkbookExtractor{unsupportedKeys: map[string]bool{}}
	if err := collector.addUnsupported("UNMODELED_WORKBOOK_FEATURE", "workbook-features", "workbook", route.part, "", "workbook feature is preserved exactly outside the v1 sheet projection"); err != nil {
		return nil, err
	}
	return append(append([]NativeWorkbookUnsupportedV1(nil), expected...), collector.unsupported[0]), nil
}

package xlsxpatch

import (
	"fmt"
	"sort"
	"strings"
)

// styleParentApplyMismatch is one cellXf component that differs from its
// cellStyleXf parent while the matching apply flag is false or absent.
type styleParentApplyMismatch struct {
	cellXF            int
	component, flag   string
	direct, inherited int
}

// newStyleRegistryForExtraction reads the style table for read-only
// extraction. Reference ids outside their tables remain fatal. A cellXf whose
// font, fill, border, number format, or alignment differs from its cellStyleXf
// parent without an apply flag is recorded instead of refused: Excel renders a
// cell from the cellXf's own record (apply flags only govern cell-style
// inheritance and UI), so the in-memory registry displays that record. Source
// bytes are untouched, and newStyleRegistry stays the strict boundary for every
// mutation and write-back path, so inheritance-dependent edits remain refused.
func newStyleRegistryForExtraction(data []byte) (*styleRegistry, []styleParentApplyMismatch, error) {
	registry, err := readStyleRegistryRecords(data)
	if err != nil {
		return nil, nil, err
	}
	for position := range registry.styleXfs {
		if err := registry.validateXFReferences(registry.styleXfs[position], false); err != nil {
			return nil, nil, fmt.Errorf("cellStyleXf %d: %w", position, err)
		}
	}
	var mismatches []styleParentApplyMismatch
	for position := range registry.cellXfs {
		xf := &registry.cellXfs[position]
		if err := registry.validateXFReferenceIDs(*xf, true); err != nil {
			return nil, nil, fmt.Errorf("cellXf %d: %w", position, err)
		}
		base := effectiveCellStyleXF(registry.styleXfs[xf.xfID])
		for _, component := range []struct {
			name, flag        string
			direct, inherited int
			apply             **bool
		}{
			{"font", "applyFont", xf.fontID, base.fontID, &xf.applyFont},
			{"fill", "applyFill", xf.fillID, base.fillID, &xf.applyFill},
			{"border", "applyBorder", xf.borderID, base.borderID, &xf.applyBorder},
			{"number format", "applyNumberFormat", xf.numFmtID, base.numFmtID, &xf.applyNumberFormat},
		} {
			if (*component.apply == nil || !**component.apply) && component.direct != component.inherited {
				mismatches = append(mismatches, styleParentApplyMismatch{cellXF: position, component: component.name, flag: component.flag, direct: component.direct, inherited: component.inherited})
				*component.apply = boolPointer(true)
			}
		}
		if (xf.applyAlignment == nil || !*xf.applyAlignment) && xf.alignment.present && !styleAlignmentsEquivalent(xf.alignment, base.alignment) {
			mismatches = append(mismatches, styleParentApplyMismatch{cellXF: position, component: "alignment", flag: "applyAlignment", direct: -1, inherited: -1})
			xf.applyAlignment = boolPointer(true)
		}
		if err := registry.validateXFReferences(*xf, true); err != nil {
			return nil, nil, fmt.Errorf("cellXf %d: %w", position, err)
		}
	}
	return registry, mismatches, nil
}

// refuseNativeMutationForStyleParentApplyMismatch keeps the mutation boundary
// where it stood before extraction learned to preview these workbooks: every
// transaction is refused while the source styles carry apply-flag mismatches,
// because cellXf/cellStyleXf inheritance is not modeled for write-back.
func refuseNativeMutationForStyleParentApplyMismatch(before *NativeWorkbookV1) error {
	for _, item := range before.Unsupported {
		if item.Code == "STYLE_PARENT_APPLY_MISMATCH" {
			return fmt.Errorf("xlsxpatch: native mutation: STYLE_PARENT_APPLY_MISMATCH: source styles remain authority-bound (%s)", item.Message)
		}
	}
	return nil
}

const styleParentApplyMismatchMessageLimit = 8

// styleParentApplyMismatchMessage describes the recorded mismatches for one
// bounded workbook-scoped inventory entry, in cellXf order.
func styleParentApplyMismatchMessage(mismatches []styleParentApplyMismatch) string {
	if len(mismatches) == 0 {
		return ""
	}
	sorted := make([]styleParentApplyMismatch, len(mismatches))
	copy(sorted, mismatches)
	sort.SliceStable(sorted, func(i, j int) bool {
		if sorted[i].cellXF != sorted[j].cellXF {
			return sorted[i].cellXF < sorted[j].cellXF
		}
		return sorted[i].component < sorted[j].component
	})
	details := make([]string, 0, styleParentApplyMismatchMessageLimit)
	for _, item := range sorted {
		if len(details) == styleParentApplyMismatchMessageLimit {
			details = append(details, fmt.Sprintf("and %d more", len(sorted)-styleParentApplyMismatchMessageLimit))
			break
		}
		if item.direct >= 0 {
			details = append(details, fmt.Sprintf("cellXf %d %s id %d differs from cellStyleXf id %d without %s", item.cellXF, item.component, item.direct, item.inherited, item.flag))
		} else {
			details = append(details, fmt.Sprintf("cellXf %d %s differs from its cellStyleXf parent without %s", item.cellXF, item.component, item.flag))
		}
	}
	return strings.Join(details, "; ") + "; each cellXf's own record is displayed and style-inheritance-dependent mutation is refused"
}

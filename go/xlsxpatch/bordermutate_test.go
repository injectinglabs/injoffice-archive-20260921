package xlsxpatch

import "testing"

func TestNativeBorderEdgesRoundTrip(t *testing.T) {
	for _, strict := range []bool{false, true} {
		original := buildZip(t, nativeMutationFixture(strict))
		apply := func(delta StyleDelta) *NativeWorkbookV1 {
			t.Helper()
			before, err := ExtractNativeWorkbookV1(original)
			if err != nil {
				t.Fatal(err)
			}
			result, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Styles: []StylePatchMutation{{OperationID: "border", SheetID: "7", Kind: StylePatch, Range: StyleRange{Row: 0, Column: 1, EndRow: 0, EndColumn: 1}, Style: delta}}})
			if err != nil {
				t.Fatal(err)
			}
			original = result.Package
			reopened, err := ExtractNativeWorkbookV1(original)
			if err != nil {
				t.Fatal(err)
			}
			return reopened
		}
		edge := func(w *NativeWorkbookV1) *NativeWorkbookBorderV1 {
			for _, c := range w.Sheets[0].Cells {
				if c.Ref == "B1" {
					return w.Styles[c.StyleID].Effective.Border
				}
			}
			t.Fatal("missing B1")
			return nil
		}
		b := edge(apply(StyleDelta{BorderTop: SetStyleProperty(BorderEdge{"double", "#123456"}), BorderLeft: SetStyleProperty(BorderEdge{"thin", "#ABCDEF"})}))
		if b.Top == nil || b.Top.Style != "double" || b.Top.Color != "#123456" || b.Left == nil {
			t.Fatalf("border readback: %+v", b)
		}
		b = edge(apply(StyleDelta{BorderTop: SetStyleProperty(BorderEdge{"none", "#000000"})}))
		if b.Top != nil || b.Left == nil {
			t.Fatalf("clearing top lost left: %+v", b)
		}
		b = edge(apply(StyleDelta{BorderLeft: ClearStyleProperty[BorderEdge]()}))
		if b.Left != nil {
			t.Fatal("left did not inherit empty edge")
		}
	}
}

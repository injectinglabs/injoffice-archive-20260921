package pptxpatch

import "fmt"

type nativeInspectionBudgetError struct{ resource string }

func (e *nativeInspectionBudgetError) Error() string {
	return fmt.Sprintf("table inspection: %s budget exceeded", e.resource)
}

type nativeInspectionBudget struct{ nodes, cells, paragraphs, runs, text int }

// Count iteratively before grammar recursion and text concatenation, including
// empty nodes and refused frames. Limits are shared across the entire package.
func (b *nativeInspectionBudget) scan(root *nativeXMLNode, d nativeExtractDialect) error {
	stack := []*nativeXMLNode{root}
	for len(stack) > 0 {
		n := stack[len(stack)-1]
		stack = stack[:len(stack)-1]
		b.nodes++
		if b.nodes > 20000 {
			return &nativeInspectionBudgetError{"node"}
		}
		if n.Name.Space == d.drawing {
			switch n.Name.Local {
			case "tc":
				b.cells++
				if b.cells > 4096 {
					return &nativeInspectionBudgetError{"cell"}
				}
			case "txBody":
				count := 0
				for _, c := range n.Children {
					if c.Name.Space == d.drawing && c.Name.Local == "p" {
						count++
					}
				}
				if count > 256 {
					return &nativeInspectionBudgetError{"paragraphs per cell"}
				}
			case "p":
				b.paragraphs++
				if b.paragraphs > 4096 {
					return &nativeInspectionBudgetError{"paragraph"}
				}
				count := 0
				for _, c := range n.Children {
					if c.Name.Space == d.drawing && c.Name.Local == "r" {
						count++
					}
				}
				if count > 256 {
					return &nativeInspectionBudgetError{"runs per paragraph"}
				}
			case "r":
				b.runs++
				if b.runs > 16384 {
					return &nativeInspectionBudgetError{"run"}
				}
			case "t":
				b.text += int(utf16CodeUnitLengthBounded(n.Text, nativeTableInspectionMaxText+1))
				if b.text > nativeTableInspectionMaxText {
					return &nativeInspectionBudgetError{"text"}
				}
			}
		}
		if len(stack)+len(n.Children) > 20000-b.nodes {
			return &nativeInspectionBudgetError{"node"}
		}
		stack = append(stack, n.Children...)
	}
	return nil
}

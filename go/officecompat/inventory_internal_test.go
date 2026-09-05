package officecompat

import "testing"

func TestExceedsCompressionRatioUsesOverflowSafeArithmetic(t *testing.T) {
	maxUint64 := ^uint64(0)
	tests := []struct {
		name         string
		uncompressed uint64
		compressed   uint64
		want         bool
	}{
		{name: "slack permits empty compressed payload", uncompressed: compressionRatioSlack, compressed: 0},
		{name: "one byte beyond slack with empty compressed payload", uncompressed: compressionRatioSlack + 1, compressed: 0, want: true},
		{name: "exact ratio boundary", uncompressed: compressionRatioSlack + MaxCompressionRatio, compressed: 1},
		{name: "one byte beyond ratio boundary", uncompressed: compressionRatioSlack + MaxCompressionRatio + 1, compressed: 1, want: true},
		{name: "adversarial sizes do not overflow multiply", uncompressed: maxUint64, compressed: maxUint64},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			if got := exceedsCompressionRatio(test.uncompressed, test.compressed); got != test.want {
				t.Fatalf("exceedsCompressionRatio(%d, %d) = %v, want %v", test.uncompressed, test.compressed, got, test.want)
			}
		})
	}
}

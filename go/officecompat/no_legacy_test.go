package officecompat_test

import (
	"os"
	"strings"
	"testing"
)

func TestNativeMutationPolicyHasNoLegacyOrRendererAuthority(t *testing.T) {
	source, err := os.ReadFile("mutation.go")
	if err != nil {
		t.Fatal(err)
	}
	lower := strings.ToLower(string(source))
	for _, forbidden := range []string{"mammoth", "luckyexcel", "html", "document.objectmodel", "browser zip", "domparser"} {
		if strings.Contains(lower, forbidden) {
			t.Fatalf("renderer or reconstruction authority %q entered native mutation policy", forbidden)
		}
	}
}

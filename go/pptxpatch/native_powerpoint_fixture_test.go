package pptxpatch

import (
	"crypto/sha256"
	"encoding/hex"
	"os"
	"strings"
	"testing"
)

const powerpointAuthoredQRSHA256 = "b4a503d90634e117ca53fb62d6aaf0657b78bfe0269b7314a313077a4899e095"

func readPowerPointAuthoredFixture(t *testing.T, name, expectedSHA string) []byte {
	t.Helper()
	data, err := os.ReadFile("testdata/powerpoint-authored/" + name)
	if err != nil {
		t.Fatal(err)
	}
	digest := sha256.Sum256(data)
	if actual := hex.EncodeToString(digest[:]); actual != expectedSHA {
		t.Fatalf("%s bytes changed: sha256=%s, want %s", name, actual, expectedSHA)
	}
	return data
}

func TestPowerPointAuthoredNativeFixtureKeepsOfficeAuthority(t *testing.T) {
	qr := readPowerPointAuthoredFixture(t, "attendee-survey-qr.pptx", powerpointAuthoredQRSHA256)
	app := string(chartZipEntry(t, qr, "docProps/app.xml"))
	if !strings.Contains(app, "<Application>Microsoft Macintosh PowerPoint</Application>") || !strings.Contains(app, "<AppVersion>16.0000</AppVersion>") {
		t.Fatalf("positive fixture lost PowerPoint authoring provenance: %s", app)
	}
	deck, err := ExtractNativePPTX(qr, nativeTestExtractOptions())
	if err != nil {
		if !strings.Contains(err.Error(), "pptxpatch: native extract:") {
			t.Fatalf("PowerPoint-authored fixture failed outside native extract: %v", err)
		}
		t.Logf("native extract fail-closed: %v", err)
		return
	}
	if deck.ContractVersion != NativePPTXContractVersion || deck.Origin != NativeOriginParsed {
		t.Fatalf("unexpected contract identity: %#v", deck)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("extracted PowerPoint-authored deck is invalid: %#v", issues)
	}
}

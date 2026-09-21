// Command expectgen rewrites the checked-in native expectations of the
// accepted DOCX corpus fixtures from a fresh extraction. Run it after an
// intentional change to the native contract projection, then re-run the
// corpus tests.
package main

import (
	"bytes"
	"crypto/sha256"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"

	"github.com/injectinglabs/injoffice/go/docxpatch"
	"github.com/injectinglabs/injoffice/go/officecompat/corpus"
)

type expectation struct {
	Protocol  string          `json:"protocol"`
	FixtureID string          `json:"fixtureId"`
	Format    string          `json:"format"`
	Dialect   string          `json:"dialect"`
	Outcome   string          `json:"outcome"`
	Native    json.RawMessage `json:"native,omitempty"`
	Refusal   json.RawMessage `json:"refusal,omitempty"`
}

func main() {
	root := "corpus"
	manifestBytes, err := os.ReadFile(filepath.Join(root, "manifest.json"))
	if err != nil {
		panic(err)
	}
	var manifest corpus.Manifest
	if err := json.Unmarshal(manifestBytes, &manifest); err != nil {
		panic(err)
	}
	for _, fixture := range manifest.Fixtures {
		if fixture.Format != "docx" || fixture.Outcome != "accepted" {
			continue
		}
		packageBytes, err := os.ReadFile(filepath.Join(root, fixture.Package))
		if err != nil {
			panic(err)
		}
		doc, err := docxpatch.ExtractNativeDocumentV1(packageBytes)
		if err != nil {
			panic(err)
		}
		encoded, err := docxpatch.EncodeNativeDocumentV1(doc)
		if err != nil {
			panic(err)
		}
		path := filepath.Join(root, fixture.Expected)
		current, err := os.ReadFile(path)
		if err != nil {
			panic(err)
		}
		var value expectation
		if err := json.Unmarshal(current, &value); err != nil {
			panic(err)
		}
		value.Native = encoded
		output, err := json.MarshalIndent(value, "", "  ")
		if err != nil {
			panic(err)
		}
		output = append(output, '\n')
		if !bytes.Equal(output, current) {
			if err := os.WriteFile(path, output, 0o644); err != nil {
				panic(err)
			}
		}
		fmt.Printf("%s file=%x canonical=%x\n", fixture.ID, sha256.Sum256(output), sha256.Sum256(encoded))
	}
}

// digestReport is printed for every rewritten fixture so the corpus and
// production-e2e manifests can be repinned from the same run.

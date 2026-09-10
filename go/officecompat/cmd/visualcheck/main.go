// visualcheck compares externally captured pages against reviewed references.
// It never generates or updates reference images.
package main

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"flag"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	officecompat "github.com/injectinglabs/injoffice/go/officecompat"
)

const maxFileBytes = 64 << 20

type manifest struct {
	Version int          `json:"version"`
	Cases   []visualCase `json:"cases"`
}

type visualCase struct {
	ID                string                       `json:"id"`
	Format            string                       `json:"format"`
	Source            string                       `json:"source"`
	SourceSHA256      string                       `json:"sourceSha256"`
	ReferenceRenderer string                       `json:"referenceRenderer"`
	CandidateRenderer string                       `json:"candidateRenderer"`
	Limits            officecompat.VisualLimits    `json:"limits"`
	Tolerance         officecompat.VisualTolerance `json:"tolerance"`
	Pages             []pagePair                   `json:"pages"`
}

type pagePair struct {
	Reference       string `json:"reference"`
	Candidate       string `json:"candidate"`
	ReferenceSHA256 string `json:"referenceSha256"`
}

type pageResult struct {
	Page   int                        `json:"page"`
	Match  bool                       `json:"match"`
	Report *officecompat.VisualReport `json:"report,omitempty"`
	Error  string                     `json:"error,omitempty"`
}

type caseResult struct {
	ID                string       `json:"id"`
	Format            string       `json:"format"`
	SourceSHA256      string       `json:"sourceSha256"`
	ReferenceRenderer string       `json:"referenceRenderer"`
	CandidateRenderer string       `json:"candidateRenderer"`
	Pages             []pageResult `json:"pages"`
}

type result struct {
	Version int          `json:"version"`
	Passed  bool         `json:"passed"`
	Cases   []caseResult `json:"cases"`
}

func main() {
	path := flag.String("manifest", "", "reviewed visual comparison manifest (required)")
	flag.Parse()
	if *path == "" || flag.NArg() != 0 {
		fmt.Fprintln(os.Stderr, "usage: visualcheck -manifest path/to/manifest.json")
		os.Exit(2)
	}
	report, err := run(*path)
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(2)
	}
	encoder := json.NewEncoder(os.Stdout)
	encoder.SetIndent("", "  ")
	if err := encoder.Encode(report); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(2)
	}
	if !report.Passed {
		os.Exit(1)
	}
}

func run(path string) (result, error) {
	data, err := readBounded(path, 1<<20)
	if err != nil {
		return result{}, err
	}
	var spec manifest
	decoder := json.NewDecoder(strings.NewReader(string(data)))
	decoder.DisallowUnknownFields()
	if err := decoder.Decode(&spec); err != nil {
		return result{}, err
	}
	if err := decoder.Decode(new(any)); err != io.EOF {
		return result{}, fmt.Errorf("manifest must contain one JSON value")
	}
	if spec.Version != 1 || len(spec.Cases) == 0 || len(spec.Cases) > 256 {
		return result{}, fmt.Errorf("manifest requires version 1 and 1–256 cases")
	}
	base := filepath.Dir(path)
	report := result{Version: 1, Passed: true, Cases: []caseResult{}}
	seen := map[string]bool{}
	totalPages := 0
	for _, c := range spec.Cases {
		if strings.TrimSpace(c.ID) == "" || seen[c.ID] || strings.TrimSpace(c.ReferenceRenderer) == "" || strings.TrimSpace(c.CandidateRenderer) == "" {
			return result{}, fmt.Errorf("cases require unique IDs and explicit renderer versions")
		}
		seen[c.ID] = true
		switch c.Format {
		case "docx", "xlsx", "pptx", "pdf":
		default:
			return result{}, fmt.Errorf("%s: unsupported format", c.ID)
		}
		totalPages += len(c.Pages)
		if len(c.Pages) == 0 || totalPages > 2048 {
			return result{}, fmt.Errorf("requires pages, at most 2048 in total")
		}
		if c.Limits.MaxEncodedBytes == 0 || c.Limits.MaxEncodedBytes > maxFileBytes || c.Limits.MaxPixels == 0 || c.Limits.MaxPixels > 32_000_000 || c.Limits.MaxWidth == 0 || c.Limits.MaxWidth > 16384 || c.Limits.MaxHeight == 0 || c.Limits.MaxHeight > 16384 {
			return result{}, fmt.Errorf("%s: missing or excessive image limits", c.ID)
		}
		source, err := readRelative(base, c.Source, maxFileBytes)
		if err != nil {
			return result{}, fmt.Errorf("%s source: %w", c.ID, err)
		}
		if !matchesHash(source, c.SourceSHA256) {
			return result{}, fmt.Errorf("%s: source digest mismatch", c.ID)
		}
		entry := caseResult{ID: c.ID, Format: c.Format, SourceSHA256: c.SourceSHA256, ReferenceRenderer: c.ReferenceRenderer, CandidateRenderer: c.CandidateRenderer, Pages: []pageResult{}}
		for i, p := range c.Pages {
			page := pageResult{Page: i + 1}
			ref, refErr := readRelative(base, p.Reference, int64(c.Limits.MaxEncodedBytes))
			candidate, candidateErr := readRelative(base, p.Candidate, int64(c.Limits.MaxEncodedBytes))
			switch {
			case refErr != nil:
				page.Error = fmt.Sprintf("reference: %v", refErr)
			case !matchesHash(ref, p.ReferenceSHA256):
				page.Error = "reference digest mismatch"
			case candidateErr != nil:
				page.Error = fmt.Sprintf("candidate: %v", candidateErr)
			default:
				comparison, err := officecompat.ComparePNG(ref, candidate, c.Limits, c.Tolerance)
				if err != nil {
					page.Error = err.Error()
				} else {
					page.Report = &comparison
					page.Match = comparison.Matches()
				}
			}
			if !page.Match {
				report.Passed = false
			}
			entry.Pages = append(entry.Pages, page)
		}
		report.Cases = append(report.Cases, entry)
	}
	return report, nil
}

func matchesHash(data []byte, want string) bool {
	digest := sha256.Sum256(data)
	return len(want) == 64 && want == hex.EncodeToString(digest[:])
}

func readRelative(base, path string, limit int64) ([]byte, error) {
	if !filepath.IsLocal(path) {
		return nil, fmt.Errorf("path must remain relative to manifest")
	}
	// Resolve links too: reviewed manifests must not escape their corpus folder.
	root, err := filepath.EvalSymlinks(base)
	if err != nil {
		return nil, err
	}
	resolved, err := filepath.EvalSymlinks(filepath.Join(base, path))
	if err != nil {
		return nil, err
	}
	rel, err := filepath.Rel(root, resolved)
	if err != nil || !filepath.IsLocal(rel) {
		return nil, fmt.Errorf("path escapes corpus folder")
	}
	return readBounded(resolved, limit)
}

func readBounded(path string, limit int64) ([]byte, error) {
	file, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer file.Close()
	stat, err := file.Stat()
	if err != nil {
		return nil, err
	}
	if !stat.Mode().IsRegular() || stat.Size() > limit {
		return nil, fmt.Errorf("requires a regular file of at most %d bytes", limit)
	}
	data, err := io.ReadAll(io.LimitReader(file, limit+1))
	if err != nil {
		return nil, err
	}
	if int64(len(data)) > limit {
		return nil, fmt.Errorf("file exceeds byte budget")
	}
	return data, nil
}

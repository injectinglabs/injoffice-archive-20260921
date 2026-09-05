// Package xlsxpatch is InjOffice's fail-closed fidelity core for xlsx files.
//
// Philosophy (the roadmap's design rule #3): the ORIGINAL file is the source
// of truth. A save applies a small set of explicit part-level changes and
// copies every other zip entry raw — same compressed bytes, same metadata.
// After writing, the result is re-opened and every untouched entry is
// verified byte-identical against the original; any mismatch fails the save
// rather than shipping a file that silently lost content. This is what makes
// it safe for an editor that models only PART of the format (values, styles)
// to save workbooks containing parts it doesn't understand (charts, pivots,
// images, macros-free customXml, ...).
package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"sort"
)

// Patch is a set of part-level changes to apply to an xlsx (OPC zip) file.
// Paths are zip entry names, e.g. "xl/worksheets/sheet1.xml".
type Patch struct {
	// Replace swaps an existing entry's content. Replacing a missing entry is
	// an error (fail closed — a typo'd path must not silently become an add).
	Replace map[string][]byte
	// Add creates a new entry. Adding over an existing entry is an error.
	Add map[string][]byte
	// Delete removes an entry. Deleting a missing entry is an error.
	Delete map[string]bool
}

func (p Patch) empty() bool {
	return len(p.Replace) == 0 && len(p.Add) == 0 && len(p.Delete) == 0
}

// Apply produces a new xlsx where exactly the patched parts changed and every
// other entry is bit-for-bit the original (raw-copied, then verified). Returns
// the new file's bytes or an error; on error the original must keep being used.
func Apply(orig []byte, p Patch) ([]byte, error) {
	zr, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: open original: %w", err)
	}

	byName := make(map[string]*zip.File, len(zr.File))
	for _, f := range zr.File {
		if _, dup := byName[f.Name]; dup {
			// Duplicate entries are legal in zip but ambiguous to patch —
			// refuse rather than guess which one a reader will use.
			return nil, fmt.Errorf("xlsxpatch: original has duplicate entry %q", f.Name)
		}
		byName[f.Name] = f
	}

	for name := range p.Replace {
		if _, ok := byName[name]; !ok {
			return nil, fmt.Errorf("xlsxpatch: replace of missing entry %q", name)
		}
	}
	for name := range p.Add {
		if _, ok := byName[name]; ok {
			return nil, fmt.Errorf("xlsxpatch: add of already-existing entry %q", name)
		}
		if _, alsoReplaced := p.Replace[name]; alsoReplaced {
			return nil, fmt.Errorf("xlsxpatch: entry %q in both add and replace", name)
		}
	}
	for name := range p.Delete {
		if _, ok := byName[name]; !ok {
			return nil, fmt.Errorf("xlsxpatch: delete of missing entry %q", name)
		}
		if _, alsoReplaced := p.Replace[name]; alsoReplaced {
			return nil, fmt.Errorf("xlsxpatch: entry %q in both delete and replace", name)
		}
	}

	var out bytes.Buffer
	zw := zip.NewWriter(&out)

	// Preserve original entry order (readers shouldn't care, but byte-level
	// predictability makes diffs and debugging sane), appending adds at the end
	// in sorted order for determinism.
	for _, f := range zr.File {
		if p.Delete[f.Name] {
			continue
		}
		if replacement, ok := p.Replace[f.Name]; ok {
			w, err := zw.CreateHeader(&zip.FileHeader{Name: f.Name, Method: zip.Deflate})
			if err != nil {
				return nil, fmt.Errorf("xlsxpatch: write %q: %w", f.Name, err)
			}
			if _, err := w.Write(replacement); err != nil {
				return nil, fmt.Errorf("xlsxpatch: write %q: %w", f.Name, err)
			}
			continue
		}
		// Raw copy: compressed bytes and header move over untouched.
		if err := zw.Copy(f); err != nil {
			return nil, fmt.Errorf("xlsxpatch: raw-copy %q: %w", f.Name, err)
		}
	}
	added := make([]string, 0, len(p.Add))
	for name := range p.Add {
		added = append(added, name)
	}
	sort.Strings(added)
	for _, name := range added {
		w, err := zw.CreateHeader(&zip.FileHeader{Name: name, Method: zip.Deflate})
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: add %q: %w", name, err)
		}
		if _, err := w.Write(p.Add[name]); err != nil {
			return nil, fmt.Errorf("xlsxpatch: add %q: %w", name, err)
		}
	}
	if err := zw.Close(); err != nil {
		return nil, fmt.Errorf("xlsxpatch: finalize: %w", err)
	}

	if err := verify(orig, out.Bytes(), p); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

// verify re-opens the produced file and proves, entry by entry, that exactly
// the requested changes happened: untouched entries byte-identical, replaced/
// added entries present with the requested content, deleted entries gone.
// This is the fail-closed contract — Apply's writer bugs (or archive/zip
// surprises) surface here as errors, never as silently-mutated files.
func verify(orig, produced []byte, p Patch) error {
	or, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return fmt.Errorf("xlsxpatch: verify reopen original: %w", err)
	}
	nr, err := zip.NewReader(bytes.NewReader(produced), int64(len(produced)))
	if err != nil {
		return fmt.Errorf("xlsxpatch: verify reopen produced: %w", err)
	}
	newByName := make(map[string]*zip.File, len(nr.File))
	for _, f := range nr.File {
		newByName[f.Name] = f
	}

	readAll := func(f *zip.File) ([]byte, error) {
		rc, err := f.Open()
		if err != nil {
			return nil, err
		}
		defer rc.Close()
		return io.ReadAll(rc)
	}

	expected := 0
	for _, of := range or.File {
		if p.Delete[of.Name] {
			if _, still := newByName[of.Name]; still {
				return fmt.Errorf("xlsxpatch: verify: deleted entry %q still present", of.Name)
			}
			continue
		}
		expected++
		nf, ok := newByName[of.Name]
		if !ok {
			return fmt.Errorf("xlsxpatch: verify: entry %q lost", of.Name)
		}
		want, isReplaced := p.Replace[of.Name]
		got, err := readAll(nf)
		if err != nil {
			return fmt.Errorf("xlsxpatch: verify read %q: %w", of.Name, err)
		}
		if isReplaced {
			if !bytes.Equal(got, want) {
				return fmt.Errorf("xlsxpatch: verify: replaced entry %q content mismatch", of.Name)
			}
			continue
		}
		origContent, err := readAll(of)
		if err != nil {
			return fmt.Errorf("xlsxpatch: verify read original %q: %w", of.Name, err)
		}
		if !bytes.Equal(got, origContent) {
			return fmt.Errorf("xlsxpatch: verify: untouched entry %q changed", of.Name)
		}
	}
	for name, want := range p.Add {
		expected++
		nf, ok := newByName[name]
		if !ok {
			return fmt.Errorf("xlsxpatch: verify: added entry %q missing", name)
		}
		got, err := readAll(nf)
		if err != nil {
			return fmt.Errorf("xlsxpatch: verify read %q: %w", name, err)
		}
		if !bytes.Equal(got, want) {
			return fmt.Errorf("xlsxpatch: verify: added entry %q content mismatch", name)
		}
	}
	if len(nr.File) != expected {
		return fmt.Errorf("xlsxpatch: verify: produced file has %d entries, expected %d", len(nr.File), expected)
	}
	return nil
}

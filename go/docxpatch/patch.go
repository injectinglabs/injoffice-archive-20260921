// Generalizes docxpatch's zip rewriting from "replace exactly one part"
// (the original rewriteZip) to a full Patch{Replace/Add/Delete}, mirroring
// go/xlsxpatch/patch.go's Apply — same fail-closed contract: apply the
// patch, then reopen the produced file and verify every untouched entry is
// byte-identical to the original, every replaced/added entry matches what
// was requested, and every deleted entry is actually gone. D11 (inline
// images, footnotes) needs this: an image insert touches THREE parts
// (document.xml, word/_rels/document.xml.rels, [Content_Types].xml) plus
// ADDS a new one (word/media/imageN.ext) — the original single-file
// rewriteZip could not add parts at all.
package docxpatch

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"sort"
)

// Patch is a set of part-level changes to apply to a .docx (OPC zip) file.
// Paths are zip entry names, e.g. "word/document.xml".
type Patch struct {
	// Replace swaps an existing entry's content. Replacing a missing entry is
	// an error (fail closed — a typo'd path must not silently become an add).
	Replace map[string][]byte
	// Add creates a new entry. Adding over an existing entry is an error.
	Add map[string][]byte
	// Delete removes an entry. Deleting a missing entry is an error.
	Delete map[string]bool
}

// ApplyPatch produces a new .docx where exactly the patched parts changed
// and every other entry is bit-for-bit the original (raw-copied, then
// verified). Returns the new file's bytes or an error; on error the
// original must keep being used.
func ApplyPatch(orig []byte, p Patch) ([]byte, error) {
	zr, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return nil, fmt.Errorf("docxpatch: open original: %w", err)
	}

	byName := make(map[string]*zip.File, len(zr.File))
	for _, f := range zr.File {
		if _, dup := byName[f.Name]; dup {
			// Duplicate entries are legal in zip but ambiguous to patch —
			// refuse rather than guess which one a reader will use.
			return nil, fmt.Errorf("docxpatch: original has duplicate entry %q", f.Name)
		}
		byName[f.Name] = f
	}

	for name := range p.Replace {
		if _, ok := byName[name]; !ok {
			return nil, fmt.Errorf("docxpatch: replace of missing entry %q", name)
		}
	}
	for name := range p.Add {
		if _, ok := byName[name]; ok {
			return nil, fmt.Errorf("docxpatch: add of already-existing entry %q", name)
		}
		if _, alsoReplaced := p.Replace[name]; alsoReplaced {
			return nil, fmt.Errorf("docxpatch: entry %q in both add and replace", name)
		}
	}
	for name := range p.Delete {
		if _, ok := byName[name]; !ok {
			return nil, fmt.Errorf("docxpatch: delete of missing entry %q", name)
		}
		if _, alsoReplaced := p.Replace[name]; alsoReplaced {
			return nil, fmt.Errorf("docxpatch: entry %q in both delete and replace", name)
		}
	}

	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	if err := zw.SetComment(zr.Comment); err != nil {
		return nil, fmt.Errorf("docxpatch: preserve archive comment: %w", err)
	}

	// Preserve original entry order (readers shouldn't care, but byte-level
	// predictability makes diffs and debugging sane), appending adds at the
	// end in sorted order for determinism.
	for _, f := range zr.File {
		if p.Delete[f.Name] {
			continue
		}
		if replacement, ok := p.Replace[f.Name]; ok {
			w, err := zw.CreateHeader(&zip.FileHeader{Name: f.Name, Method: zip.Deflate})
			if err != nil {
				return nil, fmt.Errorf("docxpatch: write %q: %w", f.Name, err)
			}
			if _, err := w.Write(replacement); err != nil {
				return nil, fmt.Errorf("docxpatch: write %q: %w", f.Name, err)
			}
			continue
		}
		// Raw copy: compressed bytes and header move over untouched.
		if err := zw.Copy(f); err != nil {
			return nil, fmt.Errorf("docxpatch: raw-copy %q: %w", f.Name, err)
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
			return nil, fmt.Errorf("docxpatch: add %q: %w", name, err)
		}
		if _, err := w.Write(p.Add[name]); err != nil {
			return nil, fmt.Errorf("docxpatch: add %q: %w", name, err)
		}
	}
	if err := zw.Close(); err != nil {
		return nil, fmt.Errorf("docxpatch: finalize: %w", err)
	}

	if err := verifyPatch(orig, out.Bytes(), p); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

// verifyPatch re-opens the produced file and proves, entry by entry, that
// exactly the requested changes happened: untouched entries byte-identical,
// replaced/added entries present with the requested content, deleted
// entries gone. This is the fail-closed contract — ApplyPatch's writer bugs
// (or archive/zip surprises) surface here as errors, never as
// silently-mutated files.
func verifyPatch(orig, produced []byte, p Patch) error {
	or, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return fmt.Errorf("docxpatch: verify reopen original: %w", err)
	}
	nr, err := zip.NewReader(bytes.NewReader(produced), int64(len(produced)))
	if err != nil {
		return fmt.Errorf("docxpatch: verify reopen produced: %w", err)
	}
	if or.Comment != nr.Comment {
		return fmt.Errorf("docxpatch: verify: archive comment changed")
	}
	newByName := make(map[string]*zip.File, len(nr.File))
	for _, f := range nr.File {
		if _, duplicate := newByName[f.Name]; duplicate {
			return fmt.Errorf("docxpatch: verify: produced file has duplicate entry %q", f.Name)
		}
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
				return fmt.Errorf("docxpatch: verify: deleted entry %q still present", of.Name)
			}
			continue
		}
		expected++
		nf, ok := newByName[of.Name]
		if !ok {
			return fmt.Errorf("docxpatch: verify: entry %q lost", of.Name)
		}
		want, isReplaced := p.Replace[of.Name]
		got, err := readAll(nf)
		if err != nil {
			return fmt.Errorf("docxpatch: verify read %q: %w", of.Name, err)
		}
		if isReplaced {
			if !bytes.Equal(got, want) {
				return fmt.Errorf("docxpatch: verify: replaced entry %q content mismatch", of.Name)
			}
			continue
		}
		if !sameRawDOCXZipHeader(of, nf) {
			return fmt.Errorf("docxpatch: verify: untouched entry %q metadata changed", of.Name)
		}
		originalRaw, err := of.OpenRaw()
		if err != nil {
			return fmt.Errorf("docxpatch: verify raw original %q: %w", of.Name, err)
		}
		producedRaw, err := nf.OpenRaw()
		if err != nil {
			return fmt.Errorf("docxpatch: verify raw produced %q: %w", of.Name, err)
		}
		leftRaw, err := io.ReadAll(originalRaw)
		if err != nil {
			return fmt.Errorf("docxpatch: verify raw original %q: %w", of.Name, err)
		}
		rightRaw, err := io.ReadAll(producedRaw)
		if err != nil {
			return fmt.Errorf("docxpatch: verify raw produced %q: %w", of.Name, err)
		}
		if !bytes.Equal(leftRaw, rightRaw) {
			return fmt.Errorf("docxpatch: verify: untouched entry %q raw compressed bytes changed", of.Name)
		}
		origContent, err := readAll(of)
		if err != nil {
			return fmt.Errorf("docxpatch: verify read original %q: %w", of.Name, err)
		}
		if !bytes.Equal(got, origContent) {
			return fmt.Errorf("docxpatch: verify: untouched entry %q changed", of.Name)
		}
	}
	for name, want := range p.Add {
		expected++
		nf, ok := newByName[name]
		if !ok {
			return fmt.Errorf("docxpatch: verify: added entry %q missing", name)
		}
		got, err := readAll(nf)
		if err != nil {
			return fmt.Errorf("docxpatch: verify read %q: %w", name, err)
		}
		if !bytes.Equal(got, want) {
			return fmt.Errorf("docxpatch: verify: added entry %q content mismatch", name)
		}
	}
	if len(nr.File) != expected {
		return fmt.Errorf("docxpatch: verify: produced file has %d entries, expected %d", len(nr.File), expected)
	}
	return nil
}

func sameRawDOCXZipHeader(left, right *zip.File) bool {
	return left.Name == right.Name && left.Method == right.Method && left.Flags == right.Flags && left.CRC32 == right.CRC32 &&
		left.CompressedSize64 == right.CompressedSize64 && left.UncompressedSize64 == right.UncompressedSize64 &&
		left.Comment == right.Comment && bytes.Equal(left.Extra, right.Extra) && left.Modified.Equal(right.Modified) &&
		left.ModifiedTime == right.ModifiedTime && left.ModifiedDate == right.ModifiedDate && left.ExternalAttrs == right.ExternalAttrs &&
		left.CreatorVersion == right.CreatorVersion && left.ReaderVersion == right.ReaderVersion && left.NonUTF8 == right.NonUTF8
}

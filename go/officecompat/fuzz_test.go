package officecompat_test

import (
	"archive/zip"
	"bytes"
	"reflect"
	"testing"

	"github.com/injectinglabs/injoffice/go/officecompat"
)

func FuzzStructuralInspectDeterministic(f *testing.F) {
	f.Add(buildFuzzPackage(map[string][]byte{
		"[Content_Types].xml": []byte(contentTypes),
		"_rels/.rels":         []byte(emptyRootRels),
		"word/document.xml":   []byte(`<document><p>seed</p></document>`),
	}))
	f.Add([]byte("not a ZIP"))
	f.Add([]byte{})
	limits := officecompat.DefaultLimits()
	limits.MaxPackageBytes = 1 << 20
	limits.MaxPartBytes = 512 << 10
	limits.MaxExpandedBytes = 1 << 20
	limits.MaxXMLBytes = 512 << 10
	limits.MaxParts = 256
	f.Fuzz(func(t *testing.T, data []byte) {
		first, firstErr := officecompat.InspectWithLimits(data, limits)
		second, secondErr := officecompat.InspectWithLimits(data, limits)
		if !reflect.DeepEqual(first, second) || errorString(firstErr) != errorString(secondErr) {
			t.Fatalf("nondeterministic inspection:\nfirst:  %+v / %v\nsecond: %+v / %v", first, firstErr, second, secondErr)
		}
	})
}

func FuzzStructuralXMLDeterministic(f *testing.F) {
	f.Add([]byte(`<a xmlns="urn:x" x="1"><b>text</b></a>`), []byte(`<q:a xmlns:q="urn:x" x='1'><q:b>text</q:b></q:a>`))
	f.Add([]byte(`<!DOCTYPE a><a/>`), []byte(`<a/>`))
	f.Add([]byte(`<a><b></a>`), []byte(`<a/>`))
	limits := officecompat.DefaultLimits()
	limits.MaxXMLBytes = 512 << 10
	limits.MaxPartBytes = 512 << 10
	limits.MaxExpandedBytes = 1 << 20
	limits.MaxXMLTokens = 100_000
	limits.MaxXMLAttributes = 10_000
	f.Fuzz(func(t *testing.T, before, after []byte) {
		first, firstErr := officecompat.CompareXMLWithLimits(before, after, limits)
		second, secondErr := officecompat.CompareXMLWithLimits(before, after, limits)
		if !reflect.DeepEqual(first, second) || errorString(firstErr) != errorString(secondErr) {
			t.Fatalf("nondeterministic XML comparison:\nfirst:  %+v / %v\nsecond: %+v / %v", first, firstErr, second, secondErr)
		}
	})
}

func errorString(err error) string {
	if err == nil {
		return ""
	}
	return err.Error()
}

func buildFuzzPackage(entries map[string][]byte) []byte {
	var buffer bytes.Buffer
	writer := zip.NewWriter(&buffer)
	for _, name := range []string{"[Content_Types].xml", "_rels/.rels", "word/document.xml"} {
		part, err := writer.Create(name)
		if err != nil {
			panic(err)
		}
		if _, err := part.Write(entries[name]); err != nil {
			panic(err)
		}
	}
	if err := writer.Close(); err != nil {
		panic(err)
	}
	return buffer.Bytes()
}

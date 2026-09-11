package main

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"strings"
	"testing"
)

// The original JPEG bands become blue/red when half-turn + vertical flip are
// composed. This is a source-defined pixel oracle, not an Office screenshot.
func TestNativePreviewTransformedImageBrowserFixture(t *testing.T) {
	output := os.Getenv("INJOFFICE_TRANSFORM_FIXTURE_OUTPUT")
	if output == "" {
		t.Skip("optional native image transform browser fixture export")
	}
	font, err := os.ReadFile(os.Getenv("INJOFFICE_TRANSFORM_FIXTURE_FONT"))
	if err != nil {
		t.Fatal(err)
	}
	source, err := buildWithJPEG(font, true)
	if err != nil {
		t.Fatal(err)
	}
	reader, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
	if err != nil {
		t.Fatal(err)
	}
	var result bytes.Buffer
	writer := zip.NewWriter(&result)
	for _, entry := range reader.File {
		input, err := entry.Open()
		if err != nil {
			t.Fatal(err)
		}
		data, err := io.ReadAll(input)
		input.Close()
		if err != nil {
			t.Fatal(err)
		}
		if entry.Name == "word/document.xml" {
			if os.Getenv("INJOFFICE_TRANSFORM_CROP") == "left-half" {
				data = []byte(strings.Replace(string(data), `<pic:blipFill><a:blip`, `<pic:blipFill><a:srcRect l="50000"/><a:blip`, 1))
			}
			data = []byte(strings.Replace(string(data), `<a:xfrm/>`, `<a:xfrm rot="10800000" flipH="false" flipV="true"/>`, 1))
			angle := os.Getenv("INJOFFICE_TRANSFORM_ANGLE")
			if angle == "90" || angle == "270" {
				rotation := "5400000"
				if angle == "270" {
					rotation = "16200000"
				}
				data = []byte(strings.Replace(string(data), `<a:xfrm rot="10800000" flipH="false" flipV="true"/>`, `<a:xfrm rot="`+rotation+`"><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/></a:xfrm>`, 1))
				data = []byte(strings.Replace(string(data), `<wp:extent cx="1828800" cy="457200"/>`, `<wp:extent cx="457200" cy="1828800"/>`, 1))
			}
		}
		part, err := writer.Create(entry.Name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := part.Write(data); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(output, result.Bytes(), 0600); err != nil {
		t.Fatal(err)
	}
}

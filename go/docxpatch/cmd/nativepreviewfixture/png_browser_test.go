package main

import (
	"archive/zip"
	"bytes"
	"image"
	"image/color"
	"image/png"
	"io"
	"os"
	"strings"
	"testing"
)

// Export an original PNG variant for the native browser smoke without widening
// the public fixture generator. All embedded font/source bytes are preserved.
func TestNativePreviewPNGBrowserFixture(t *testing.T) {
	output := os.Getenv("INJOFFICE_PNG_FIXTURE_OUTPUT")
	if output == "" {
		t.Skip("optional native PNG browser fixture export")
	}
	font, err := os.ReadFile(os.Getenv("INJOFFICE_PNG_FIXTURE_FONT"))
	if err != nil {
		t.Fatal(err)
	}
	source, err := buildWithJPEG(font, true)
	if err != nil {
		t.Fatal(err)
	}
	pixels := image.NewRGBA(image.Rect(0, 0, 16, 8))
	for y := 0; y < 8; y++ {
		for x := 0; x < 16; x++ {
			shade := color.RGBA{220, 40, 40, 255}
			if x >= 8 {
				shade = color.RGBA{30, 80, 220, 255}
			}
			pixels.Set(x, y, shade)
		}
	}
	var media bytes.Buffer
	if err := png.Encode(&media, pixels); err != nil {
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
		name := entry.Name
		switch name {
		case "word/media/bands.jpg":
			name = "word/media/bands.png"
			data = media.Bytes()
		case "[Content_Types].xml":
			data = []byte(strings.Replace(string(data), `Extension="jpg" ContentType="image/jpeg"`, `Extension="png" ContentType="image/png"`, 1))
		case "word/_rels/document.xml.rels":
			data = []byte(strings.Replace(string(data), "media/bands.jpg", "media/bands.png", 1))
		}
		part, err := writer.Create(name)
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

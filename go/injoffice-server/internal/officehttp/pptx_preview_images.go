package officehttp

import (
	"archive/zip"
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/base64"
	"errors"
	"fmt"
	"io"

	"github.com/injectinglabs/injoffice/go/pptxpatch"
)

// Only the current, explicitly uploaded package supplies image bytes. Part
// anchors never grant access to a file, remote URL, or stored artifact.
func attachPPTXPreviewImages(ctx context.Context, data []byte, deck *pptxpatch.NativePPTXDeck, slide int) error {
	if slide < 0 || slide >= len(deck.Slides) {
		return errors.New("invalid image source slide")
	}
	wanted := map[string]bool{}
	var walk func([]pptxpatch.NativeElement, int) error
	walk = func(elements []pptxpatch.NativeElement, depth int) error {
		if depth > 32 {
			return errors.New("preview image nesting exceeds budget")
		}
		for _, element := range elements {
			if element.Kind == pptxpatch.NativeElementKindPicture && element.AssetID != nil {
				wanted[*element.AssetID] = true
			}
			if element.Kind == pptxpatch.NativeElementKindChart && element.Chart != nil && element.Chart.PreviewAssetID != nil {
				wanted[*element.Chart.PreviewAssetID] = true
			}
			if len(wanted) > 256 {
				return errors.New("preview image count exceeds budget")
			}
			if err := walk(element.Children, depth+1); err != nil {
				return err
			}
		}
		return nil
	}
	if err := walk(deck.Slides[slide].Elements, 0); err != nil {
		return err
	}
	if len(wanted) == 0 {
		return nil
	}
	archive, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return err
	}
	parts := map[string]*zip.File{}
	for _, part := range archive.File {
		if _, exists := parts[part.Name]; exists {
			return errors.New("ambiguous duplicate preview package part")
		}
		parts[part.Name] = part
	}
	total := 0
	for index := range deck.Assets {
		asset := &deck.Assets[index]
		if !wanted[asset.ID] {
			continue
		}
		if err := ctx.Err(); err != nil {
			return err
		}
		if asset.Source == nil || asset.Source.FingerprintSHA256 != asset.SHA256 || asset.ByteLength == nil || *asset.ByteLength < 1 || *asset.ByteLength > 8*1024*1024 {
			return errors.New("preview image source identity or byte budget is invalid")
		}
		if asset.ContentType != "image/png" && asset.ContentType != "image/jpeg" {
			return errors.New("native slide preview supports bounded PNG/JFIF images only")
		}
		part := parts[asset.Source.PartName]
		if part == nil || part.UncompressedSize64 != uint64(*asset.ByteLength) {
			return errors.New("preview image does not match its source part")
		}
		reader, err := part.Open()
		if err != nil {
			return err
		}
		content, readErr := io.ReadAll(io.LimitReader(reader, 8*1024*1024+1))
		reader.Close()
		if readErr != nil {
			return readErr
		}
		total += len(content)
		if total > 8*1024*1024 || int64(len(content)) != *asset.ByteLength || fmt.Sprintf("%x", sha256.Sum256(content)) != asset.SHA256 {
			return errors.New("preview image digest or cumulative byte budget failed")
		}
		encoded := base64.StdEncoding.EncodeToString(content)
		asset.DataBase64 = &encoded
		delete(wanted, asset.ID)
	}
	if len(wanted) != 0 {
		return errors.New("preview slide references an unavailable source image")
	}
	return ctx.Err()
}
